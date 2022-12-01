#include <iostream>
#include <stdexcept>
#include <pybind11/pybind11.h>

#include "Wrapper/ActProgType64_i.h"
#include "Wrapper/ActSupportMsg64_i.h"
#include "Wrapper/ActUtlType64_i.h"
#include "ActDefine.h"

namespace py = pybind11;

using i8 = int8_t;
using i16 = int16_t;
using i32 = int32_t;
using u8 = uint8_t;
using u16 = uint16_t;
using u32 = uint32_t;

class _BSTR {
public:
    BSTR bstr;
    _BSTR(): bstr(nullptr) {}
    _BSTR(const wchar_t* sz): bstr(::SysAllocString(sz)) {}
    ~_BSTR() {
        if (this->bstr) {
            ::SysFreeString(this->bstr);
            this->bstr = nullptr;
        }
    }
    operator BSTR() { return this->bstr; }
    operator BSTR*() { return &this->bstr; }
};

class _BUFFER {
public:
    Py_buffer py_buffer = {};
    _BUFFER(py::object obj, bool writable = false) {
        int flags = PyBUF_SIMPLE;
        if (writable) flags |= PyBUF_WRITABLE;
        if (PyObject_GetBuffer(obj.ptr(), &this->py_buffer, flags) != 0) {
            throw py::error_already_set();
        }
    }
    ~_BUFFER() {
        if (this->py_buffer.buf) {
            PyBuffer_Release(&this->py_buffer);
            this->py_buffer.buf = nullptr;
        }
    }
    u8* buf(size_t offset = 0) { return (u8*)(this->py_buffer.buf) + offset; }
};

template<typename T>
class ComObject {
public:
    void* _interface_pointer = nullptr;
    ComObject() {
        // std::cout << typeid(T).name() << " com object constructor" << std::endl;
        CoInitialize(NULL);
    }
    ~ComObject() {
        // std::cout << typeid(T).name() << " com object release" << std::endl;
        this->iptr().Release();
    }
    T& iptr() { return *(T*)(this->_interface_pointer); }
};

void _hresult(HRESULT hr);
void _return_code(LONG return_code);

#define _CALL0(f) {                     \
    HRESULT hr;                         \
    LONG return_code = 0;               \
    hr = f(&return_code);               \
    _hresult(hr);                       \
    _return_code(return_code);          \
}

#define _CALL(f, ...) {                 \
    HRESULT hr;                         \
    LONG return_code = 0;               \
    hr = f(__VA_ARGS__, &return_code);  \
    _hresult(hr);                       \
    _return_code(return_code);          \
}

class ActSupportMsg: public ComObject<IActSupportMsg64> {
public:
    ActSupportMsg() {
        HRESULT hr = CoCreateInstance(CLSID_ActSupportMsg64, NULL, CLSCTX_INPROC_SERVER, IID_IActSupportMsg64, &this->_interface_pointer);
        _hresult(hr);
    }

    std::wstring get_message(int error_code) {
        _BSTR error_message;
        _CALL(this->iptr().GetErrorMessage, error_code, error_message);
        return std::wstring(error_message);
    }
};

static ActSupportMsg _act_support_msg;

static std::wstring _get_act_error_message(int error_code) {
    return _act_support_msg.get_message(error_code);
}

class ActError : public std::exception {
private:
    int _error_code;

public:
    ActError(int error_code): _error_code(error_code) {}
    ActError(LONG error_code): ActError((int)error_code) {}
    int error_code() const { return this->_error_code; }
};

static std::wstring _get_window_error_message(int error_code) {
    LPWSTR buffer = nullptr;
    DWORD size = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, nullptr, error_code, 0, (LPWSTR)&buffer, 0,  nullptr);
    std::wstring error_message(buffer);
    ::LocalFree(buffer);
    return error_message;
}

class ActOleError : public ActError {
public:
    ActOleError(int error_code): ActError(error_code) {}
    ActOleError(HRESULT error_code): ActError((int)error_code) {}
};

static void _hresult(HRESULT hr) {
    if (!SUCCEEDED(hr)) {
        throw ActOleError(hr);
    }
}

static void _return_code(LONG return_code) {
    if (return_code != 0) {
        throw ActError(return_code);
    }
}

static void _dump_binary(u8* p, int size) {
    std::cout << std::hex;
    for (int i = 0; i < size; i++) {
        std::cout << (int)*p++ << " ";
    }
    std::cout << std::dec;
    std::cout << std::endl;
}

template<typename T>
class ActCommon: public ComObject<T> {
public:
    bool _check_writable = true;
    ActCommon(): _check_writable(true) {
        // std::cout << typeid(T).name() << " act common constructor" << std::endl;
    }

    void Open() {
        _CALL0(this->iptr().Open);
    }

    void Close() {
        _CALL0(this->iptr().Close);
    }

    py::object ReadDeviceBlock2(const wchar_t* device, int size, py::object data, int offset = 0) {
        if (data.is_none()) {
            data = py::bytearray("", size*sizeof(SHORT));
        }
        _BUFFER buffer(data, this->_check_writable);
        _CALL(this->iptr().ReadDeviceBlock2, _BSTR(device), size, (SHORT*)buffer.buf(offset));
        return data;
    }

    void WriteDeviceBlock2(const wchar_t* device, int size, py::object data, size_t offset = 0) {
        _BUFFER buffer(data);
        _CALL(this->iptr().WriteDeviceBlock2, _BSTR(device), size, (SHORT*)buffer.buf(offset));
    }

    py::object ReadDeviceRandom2(const wchar_t* device, int size, py::object data, int offset = 0) {
        if (data.is_none()) {
            data = py::bytearray("", size*sizeof(SHORT));
        }
        _BUFFER buffer(data, this->_check_writable);
        _CALL(this->iptr().ReadDeviceRandom2, _BSTR(device), size, (SHORT*)buffer.buf(offset));
        return data;
    }

    void WriteDeviceRandom2(const wchar_t* device, int size, py::object data, size_t offset = 0) {
        _BUFFER buffer(data);
        _CALL(this->iptr().WriteDeviceRandom2, _BSTR(device), size, (SHORT*)buffer.buf(offset));
    }

    void CheckControl() {
        _CALL0(this->iptr().CheckControl);
    }

    py::tuple GetCpuType() {
        _BSTR cpu_name;
        LONG cpu_code;
        _CALL(this->iptr().GetCpuType, cpu_name, &cpu_code);
        return py::make_tuple(std::wstring(cpu_name), cpu_code);
    }

    void SetCpuStatus(int operation) {
        _CALL(this->iptr().SetCpuStatus, operation);
    }

    py::object ReadDeviceBlock(const wchar_t* device, int size, py::object data, int offset = 0) {
        if (data.is_none()) {
            data = py::bytearray("", size*sizeof(LONG));
        }
        _BUFFER buffer(data, this->_check_writable);
        _CALL(this->iptr().ReadDeviceBlock, _BSTR(device), size, (LONG*)buffer.buf(offset));
        return data;
    }

    void WriteDeviceBlock(const wchar_t* device, int size, py::object data, size_t offset = 0) {
        _BUFFER buffer(data);
        _CALL(this->iptr().WriteDeviceBlock, _BSTR(device), size, (LONG*)buffer.buf(offset));
    }

    py::object ReadDeviceRandom(const wchar_t* device, int size, py::object data, int offset = 0) {
        if (data.is_none()) {
            data = py::bytearray("", size*sizeof(LONG));
        }
        _BUFFER buffer(data, this->_check_writable);
        _CALL(this->iptr().ReadDeviceRandom, _BSTR(device), size, (LONG*)buffer.buf(offset));
        return data;
    }

    void WriteDeviceRandom(const wchar_t* device, int size, py::object data, size_t offset = 0) {
        _BUFFER buffer(data);
        _CALL(this->iptr().WriteDeviceRandom, _BSTR(device), size, (LONG*)buffer.buf(offset));
    }

    py::object ReadBuffer(int startio, int address, int size, py::object data, int offset = 0) {
        if (data.is_none()) {
            data = py::bytearray("", size*sizeof(SHORT));
        }
        _BUFFER buffer(data, this->_check_writable);
        _CALL(this->iptr().ReadBuffer, startio, address, size, (SHORT*)buffer.buf(offset));
        return data;
    }

    void WriteBuffer(int startio, int address, int size, py::object data, size_t offset = 0) {
        _BUFFER buffer(data);
        _CALL(this->iptr().WriteBuffer, startio, address, size, (SHORT*)buffer.buf(offset));
    }

    py::tuple GetClockData() {
        SHORT year;
        SHORT month;
        SHORT day;
        SHORT day_of_week;
        SHORT hour;
        SHORT minute;
        SHORT second;
        _CALL(this->iptr().GetClockData, &year, &month, &day, &day_of_week, &hour, &minute, &second);
        return py::make_tuple(year, month, day, day_of_week, hour, minute, second);
    }

    void SetClockData(int year, int month, int day, int day_of_week, int hour, int minute, int second) {
        _CALL(this->iptr().SetClockData, year, month, day, day_of_week, hour, minute, second);
    }

    void SetDevice(const wchar_t* device, LONG data) {
        _CALL(this->iptr().SetDevice, _BSTR(device), data);
    }

    LONG GetDevice(const wchar_t* device) {
        LONG data = 0;
        _CALL(this->iptr().GetDevice, _BSTR(device), &data);
        return data;
    }

    SHORT GetDevice2(const wchar_t* device) {
        SHORT data = 0;
        _CALL(this->iptr().GetDevice2, _BSTR(device), &data);
        return data;
    }

    void SetDevice2(const wchar_t* device, SHORT data) {
        _CALL(this->iptr().SetDevice2, _BSTR(device), data);
    }
};

#define DEF_PROPERTY(property_type, property_name)                               \
    property_type get_Act##property_name() {                                     \
        property_type value;                                                     \
        HRESULT hr = this->iptr().get_Act##property_name(&value);                \
        _hresult(hr);                                                            \
        return value;                                                            \
    }                                                                            \
                                                                                 \
    void put_Act##property_name(property_type value) {                           \
        HRESULT hr = this->iptr().put_Act##property_name(value);                 \
        _hresult(hr);                                                            \
    }

#define DEF_PROPERTY_STR(property_type, property_name)                           \
    property_type get_Act##property_name() {                                     \
        _BSTR value;                                                             \
        HRESULT hr = this->iptr().get_Act##property_name(value);                 \
        _hresult(hr);                                                            \
        return property_type(value);                                             \
    }                                                                            \
                                                                                 \
    void put_Act##property_name(property_type value) {                           \
        HRESULT hr = this->iptr().put_Act##property_name(_BSTR(value.c_str()));  \
        _hresult(hr);                                                            \
    }

class ActUtlType: public ActCommon<IActUtlType64> {
public:
    ActUtlType() {
        HRESULT hr = CoCreateInstance(CLSID_ActUtlType64, NULL, CLSCTX_LOCAL_SERVER, IID_IActUtlType64, &this->_interface_pointer);
        _hresult(hr);
    }

    DEF_PROPERTY(LONG, LogicalStationNumber)
    DEF_PROPERTY_STR(std::wstring, Password)
};

class ActProgType: public ActCommon<IActProgType64> {
public:
    ActProgType() {
        HRESULT hr = CoCreateInstance(CLSID_ActProgType64, NULL, CLSCTX_LOCAL_SERVER, IID_IActProgType64, &this->_interface_pointer);
        _hresult(hr);
    }

    DEF_PROPERTY(LONG, NetworkNumber)
    DEF_PROPERTY(LONG, StationNumber)
    DEF_PROPERTY(LONG, UnitNumber)
    DEF_PROPERTY(LONG, ConnectUnitNumber)
    DEF_PROPERTY(LONG, IONumber)
    DEF_PROPERTY(LONG, CpuType)
    DEF_PROPERTY(LONG, PortNumber)
    DEF_PROPERTY(LONG, BaudRate)
    DEF_PROPERTY(LONG, DataBits)
    DEF_PROPERTY(LONG, Parity)
    DEF_PROPERTY(LONG, StopBits)
    DEF_PROPERTY(LONG, Control)
    DEF_PROPERTY_STR(std::wstring, HostAddress)
    DEF_PROPERTY(LONG, CpuTimeOut)
    DEF_PROPERTY(LONG, TimeOut)
    DEF_PROPERTY(LONG, SumCheck)
    DEF_PROPERTY(LONG, SourceNetworkNumber)
    DEF_PROPERTY(LONG, SourceStationNumber)
    DEF_PROPERTY(LONG, DestinationPortNumber)
    DEF_PROPERTY(LONG, DestinationIONumber)
    DEF_PROPERTY(LONG, MultiDropChannelNumber)
    DEF_PROPERTY(LONG, ThroughNetworkType)
    DEF_PROPERTY(LONG, IntelligentPreferenceBit)
    DEF_PROPERTY(LONG, DidPropertyBit)
    DEF_PROPERTY(LONG, DsidPropertyBit)
    DEF_PROPERTY(LONG, PacketType)
    DEF_PROPERTY_STR(std::wstring, Password)
    DEF_PROPERTY(LONG, TargetSimulator)
    DEF_PROPERTY(LONG, UnitType)
    DEF_PROPERTY(LONG, ProtocolType)
};

using namespace pybind11::literals;

PYBIND11_MODULE(_act5, m) {
    m.doc() = R"pbdoc(
        MX Component v5
        ---------------
    )pbdoc";

    static py::exception<ActError> act_error(m, "ActError");
    py::register_exception_translator([](std::exception_ptr p) {
        try {
            if (p) std::rethrow_exception(p);
        } catch (const ActError &e) {
            std::wstring msg = _get_act_error_message(e.error_code());
            // auto value = PyUnicode_FromWideChar(msg.c_str(), msg.length());
            auto value = py::make_tuple(msg, e.error_code());
            PyErr_SetObject(act_error.ptr(), value.ptr());
        }
    });

    static py::exception<ActOleError> act_ole_error(m, "ActOleError", act_error.ptr());
    py::register_exception_translator([](std::exception_ptr p) {
        try {
            if (p) std::rethrow_exception(p);
        } catch (const ActOleError &e) {
            PyErr_SetExcFromWindowsErr(act_ole_error.ptr(), e.error_code());
        }
    });

    py::class_<ActUtlType>(m, "ActUtlType")
        .def(py::init())
        .def_readwrite("CheckWritable", &ActUtlType::_check_writable)
        .def_property("ActLogicalStationNumber", &ActUtlType::get_ActLogicalStationNumber, &ActUtlType::put_ActLogicalStationNumber)
        .def_property("ActPassword", &ActUtlType::get_ActPassword, &ActUtlType::put_ActPassword)
        .def("Open", &ActUtlType::Open)
        .def("Close", &ActUtlType::Close)
        .def("ReadDeviceBlock2", &ActUtlType::ReadDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceBlock2", &ActUtlType::WriteDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("ReadDeviceRandom2", &ActUtlType::ReadDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceRandom2", &ActUtlType::WriteDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("CheckControl", &ActUtlType::CheckControl)
        .def("GetCpuType", &ActUtlType::GetCpuType)
        .def("SetCpuStatus", &ActUtlType::SetCpuStatus, py::arg("operation"))
        .def("ReadDeviceBlock", &ActUtlType::ReadDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceBlock", &ActUtlType::WriteDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("ReadDeviceRandom", &ActUtlType::ReadDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceRandom", &ActUtlType::WriteDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("ReadBuffer", &ActUtlType::ReadBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteBuffer", &ActUtlType::WriteBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("GetClockData", &ActUtlType::GetClockData)
        .def("SetClockData", &ActUtlType::SetClockData, py::arg("year"), py::arg("month"), py::arg("day"), py::arg("day_of_week"), py::arg("hour"), py::arg("minute"), py::arg("second"))
        .def("SetDevice", &ActUtlType::SetDevice, py::arg("device"), py::arg("value"))
        .def("GetDevice", &ActUtlType::GetDevice, py::arg("device"))
        .def("GetDevice2", &ActUtlType::GetDevice2, py::arg("device"))
        .def("SetDevice2", &ActUtlType::SetDevice2, py::arg("device"), py::arg("value"))
        .def_readwrite("check_writable", &ActUtlType::_check_writable)
        .def_property("logical_station_number", &ActUtlType::get_ActLogicalStationNumber, &ActUtlType::put_ActLogicalStationNumber)
        .def_property("password", &ActUtlType::get_ActPassword, &ActUtlType::put_ActPassword)
        .def("open", &ActUtlType::Open)
        .def("close", &ActUtlType::Close)
        .def("read_device_block2", &ActUtlType::ReadDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_block2", &ActUtlType::WriteDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("read_device_random2", &ActUtlType::ReadDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_random2", &ActUtlType::WriteDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("check_control", &ActUtlType::CheckControl)
        .def("get_cpu_type", &ActUtlType::GetCpuType)
        .def("set_cpu_status", &ActUtlType::SetCpuStatus, py::arg("operation"))
        .def("read_device_block", &ActUtlType::ReadDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_block", &ActUtlType::WriteDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("read_device_random", &ActUtlType::ReadDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_random", &ActUtlType::WriteDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("read_buffer", &ActUtlType::ReadBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_buffer", &ActUtlType::WriteBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("get_clock_data", &ActUtlType::GetClockData)
        .def("set_clock_data", &ActUtlType::SetClockData, py::arg("year"), py::arg("month"), py::arg("day"), py::arg("day_of_week"), py::arg("hour"), py::arg("minute"), py::arg("second"))
        .def("set_device", &ActUtlType::SetDevice, py::arg("device"), py::arg("value"))
        .def("get_device", &ActUtlType::GetDevice, py::arg("device"))
        .def("get_device2", &ActUtlType::GetDevice2, py::arg("device"))
        .def("set_device2", &ActUtlType::SetDevice2, py::arg("device"), py::arg("value"));

    py::class_<ActProgType>(m, "ActProgType")
        .def(py::init())
        .def_readwrite("CheckWritable", &ActProgType::_check_writable)
        .def_property("ActNetworkNumber", &ActProgType::get_ActNetworkNumber, &ActProgType::put_ActNetworkNumber)
        .def_property("ActStationNumber", &ActProgType::get_ActStationNumber, &ActProgType::put_ActStationNumber)
        .def_property("ActUnitNumber", &ActProgType::get_ActUnitNumber, &ActProgType::put_ActUnitNumber)
        .def_property("ActConnectUnitNumber", &ActProgType::get_ActConnectUnitNumber, &ActProgType::put_ActConnectUnitNumber)
        .def_property("ActIONumber", &ActProgType::get_ActIONumber, &ActProgType::put_ActIONumber)
        .def_property("ActCpuType", &ActProgType::get_ActCpuType, &ActProgType::put_ActCpuType)
        .def_property("ActPortNumber", &ActProgType::get_ActPortNumber, &ActProgType::put_ActPortNumber)
        .def_property("ActBaudRate", &ActProgType::get_ActBaudRate, &ActProgType::put_ActBaudRate)
        .def_property("ActDataBits", &ActProgType::get_ActDataBits, &ActProgType::put_ActDataBits)
        .def_property("ActParity", &ActProgType::get_ActParity, &ActProgType::put_ActParity)
        .def_property("ActStopBits", &ActProgType::get_ActStopBits, &ActProgType::put_ActStopBits)
        .def_property("ActControl", &ActProgType::get_ActControl, &ActProgType::put_ActControl)
        .def_property("ActHostAddress", &ActProgType::get_ActHostAddress, &ActProgType::put_ActHostAddress)
        .def_property("ActCpuTimeOut", &ActProgType::get_ActCpuTimeOut, &ActProgType::put_ActCpuTimeOut)
        .def_property("ActTimeOut", &ActProgType::get_ActTimeOut, &ActProgType::put_ActTimeOut)
        .def_property("ActSumCheck", &ActProgType::get_ActSumCheck, &ActProgType::put_ActSumCheck)
        .def_property("ActSourceNetworkNumber", &ActProgType::get_ActSourceNetworkNumber, &ActProgType::put_ActSourceNetworkNumber)
        .def_property("ActSourceStationNumber", &ActProgType::get_ActSourceStationNumber, &ActProgType::put_ActSourceStationNumber)
        .def_property("ActDestinationPortNumber", &ActProgType::get_ActDestinationPortNumber, &ActProgType::put_ActDestinationPortNumber)
        .def_property("ActDestinationIONumber", &ActProgType::get_ActDestinationIONumber, &ActProgType::put_ActDestinationIONumber)
        .def_property("ActMultiDropChannelNumber", &ActProgType::get_ActMultiDropChannelNumber, &ActProgType::put_ActMultiDropChannelNumber)
        .def_property("ActThroughNetworkType", &ActProgType::get_ActThroughNetworkType, &ActProgType::put_ActThroughNetworkType)
        .def_property("ActIntelligentPreferenceBit", &ActProgType::get_ActIntelligentPreferenceBit, &ActProgType::put_ActIntelligentPreferenceBit)
        .def_property("ActDidPropertyBit", &ActProgType::get_ActDidPropertyBit, &ActProgType::put_ActDidPropertyBit)
        .def_property("ActDsidPropertyBit", &ActProgType::get_ActDsidPropertyBit, &ActProgType::put_ActDsidPropertyBit)
        .def_property("ActPacketType", &ActProgType::get_ActPacketType, &ActProgType::put_ActPacketType)
        .def_property("ActPassword", &ActProgType::get_ActPassword, &ActProgType::put_ActPassword)
        .def_property("ActTargetSimulator", &ActProgType::get_ActTargetSimulator, &ActProgType::put_ActTargetSimulator)
        .def_property("ActUnitType", &ActProgType::get_ActUnitType, &ActProgType::put_ActUnitType)
        .def_property("ActProtocolType", &ActProgType::get_ActProtocolType, &ActProgType::put_ActProtocolType)
        .def("Open", &ActProgType::Open)
        .def("Close", &ActProgType::Close)
        .def("ReadDeviceBlock2", &ActProgType::ReadDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceBlock2", &ActProgType::WriteDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("ReadDeviceRandom2", &ActProgType::ReadDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceRandom2", &ActProgType::WriteDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("CheckControl", &ActProgType::CheckControl)
        .def("GetCpuType", &ActProgType::GetCpuType)
        .def("SetCpuStatus", &ActProgType::SetCpuStatus, py::arg("operation"))
        .def("ReadDeviceBlock", &ActProgType::ReadDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceBlock", &ActProgType::WriteDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("ReadDeviceRandom", &ActProgType::ReadDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteDeviceRandom", &ActProgType::WriteDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("ReadBuffer", &ActProgType::ReadBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("WriteBuffer", &ActProgType::WriteBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("GetClockData", &ActProgType::GetClockData)
        .def("SetClockData", &ActProgType::SetClockData, py::arg("year"), py::arg("month"), py::arg("day"), py::arg("day_of_week"), py::arg("hour"), py::arg("minute"), py::arg("second"))
        .def("SetDevice", &ActProgType::SetDevice, py::arg("device"), py::arg("value"))
        .def("GetDevice", &ActProgType::GetDevice, py::arg("device"))
        .def("GetDevice2", &ActProgType::GetDevice2, py::arg("device"))
        .def("SetDevice2", &ActProgType::SetDevice2, py::arg("device"), py::arg("value"))
        .def_readwrite("check_writable", &ActProgType::_check_writable)
        .def_property("network_number", &ActProgType::get_ActNetworkNumber, &ActProgType::put_ActNetworkNumber)
        .def_property("station_number", &ActProgType::get_ActStationNumber, &ActProgType::put_ActStationNumber)
        .def_property("unit_number", &ActProgType::get_ActUnitNumber, &ActProgType::put_ActUnitNumber)
        .def_property("connect_unit_number", &ActProgType::get_ActConnectUnitNumber, &ActProgType::put_ActConnectUnitNumber)
        .def_property("io_number", &ActProgType::get_ActIONumber, &ActProgType::put_ActIONumber)
        .def_property("cpu_type", &ActProgType::get_ActCpuType, &ActProgType::put_ActCpuType)
        .def_property("port_number", &ActProgType::get_ActPortNumber, &ActProgType::put_ActPortNumber)
        .def_property("baud_rate", &ActProgType::get_ActBaudRate, &ActProgType::put_ActBaudRate)
        .def_property("data_bits", &ActProgType::get_ActDataBits, &ActProgType::put_ActDataBits)
        .def_property("parity", &ActProgType::get_ActParity, &ActProgType::put_ActParity)
        .def_property("stop_bits", &ActProgType::get_ActStopBits, &ActProgType::put_ActStopBits)
        .def_property("control", &ActProgType::get_ActControl, &ActProgType::put_ActControl)
        .def_property("host_address", &ActProgType::get_ActHostAddress, &ActProgType::put_ActHostAddress)
        .def_property("cpu_time_out", &ActProgType::get_ActCpuTimeOut, &ActProgType::put_ActCpuTimeOut)
        .def_property("time_out", &ActProgType::get_ActTimeOut, &ActProgType::put_ActTimeOut)
        .def_property("sum_check", &ActProgType::get_ActSumCheck, &ActProgType::put_ActSumCheck)
        .def_property("source_network_number", &ActProgType::get_ActSourceNetworkNumber, &ActProgType::put_ActSourceNetworkNumber)
        .def_property("source_station_number", &ActProgType::get_ActSourceStationNumber, &ActProgType::put_ActSourceStationNumber)
        .def_property("destination_port_number", &ActProgType::get_ActDestinationPortNumber, &ActProgType::put_ActDestinationPortNumber)
        .def_property("destination_io_number", &ActProgType::get_ActDestinationIONumber, &ActProgType::put_ActDestinationIONumber)
        .def_property("multi_drop_channel_number", &ActProgType::get_ActMultiDropChannelNumber, &ActProgType::put_ActMultiDropChannelNumber)
        .def_property("through_network_type", &ActProgType::get_ActThroughNetworkType, &ActProgType::put_ActThroughNetworkType)
        .def_property("intelligent_preference_bit", &ActProgType::get_ActIntelligentPreferenceBit, &ActProgType::put_ActIntelligentPreferenceBit)
        .def_property("did_property_bit", &ActProgType::get_ActDidPropertyBit, &ActProgType::put_ActDidPropertyBit)
        .def_property("dsid_property_bit", &ActProgType::get_ActDsidPropertyBit, &ActProgType::put_ActDsidPropertyBit)
        .def_property("packet_type", &ActProgType::get_ActPacketType, &ActProgType::put_ActPacketType)
        .def_property("password", &ActProgType::get_ActPassword, &ActProgType::put_ActPassword)
        .def_property("target_simulator", &ActProgType::get_ActTargetSimulator, &ActProgType::put_ActTargetSimulator)
        .def_property("unit_type", &ActProgType::get_ActUnitType, &ActProgType::put_ActUnitType)
        .def_property("protocol_type", &ActProgType::get_ActProtocolType, &ActProgType::put_ActProtocolType)
        .def("open", &ActProgType::Open)
        .def("close", &ActProgType::Close)
        .def("read_device_block2", &ActProgType::ReadDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_block2", &ActProgType::WriteDeviceBlock2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("read_device_random2", &ActProgType::ReadDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_random2", &ActProgType::WriteDeviceRandom2, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("check_control", &ActProgType::CheckControl)
        .def("get_cpu_type", &ActProgType::GetCpuType)
        .def("set_cpu_status", &ActProgType::SetCpuStatus, py::arg("operation"))
        .def("read_device_block", &ActProgType::ReadDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_block", &ActProgType::WriteDeviceBlock, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("read_device_random", &ActProgType::ReadDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_device_random", &ActProgType::WriteDeviceRandom, py::arg("device"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("read_buffer", &ActProgType::ReadBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data") = py::none(), py::arg("offset") = 0)
        .def("write_buffer", &ActProgType::WriteBuffer, py::arg("startio"), py::arg("address"), py::arg("size"), py::arg("data"), py::arg("offset") = 0)
        .def("get_clock_data", &ActProgType::GetClockData)
        .def("set_clock_data", &ActProgType::SetClockData, py::arg("year"), py::arg("month"), py::arg("day"), py::arg("day_of_week"), py::arg("hour"), py::arg("minute"), py::arg("second"))
        .def("set_device", &ActProgType::SetDevice, py::arg("device"), py::arg("value"))
        .def("get_device", &ActProgType::GetDevice, py::arg("device"))
        .def("get_device2", &ActProgType::GetDevice2, py::arg("device"))
        .def("set_device2", &ActProgType::SetDevice2, py::arg("device"), py::arg("value"));
}

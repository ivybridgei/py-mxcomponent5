import logging
from pathlib import Path
from ctypes import c_short, c_int, create_string_buffer, cast, POINTER

import comtypes.client
try:
    from comtypes.gen import ActSupportMsg64Lib, ActUtlType64Lib, ActProgType64Lib
except ImportError:
    from winreg import HKEY_CLASSES_ROOT, OpenKeyEx, QueryValue, KEY_READ, KEY_WOW64_32KEY
    clsid = QueryValue(HKEY_CLASSES_ROOT, "ActUtlType.ActUtlType\\CLSID")
    clsid_key = OpenKeyEx(HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, KEY_READ | KEY_WOW64_32KEY)
    _dir = Path(QueryValue(clsid_key, "InprocServer32")).parent
    ActSupportMsg64Lib = comtypes.client.GetModule(str(_dir / "Wrapper" / "ActSupportMsg64.dll"))
    ActUtlType64Lib = comtypes.client.GetModule(str(_dir / "ActUtlType64.exe"))
    ActProgType64Lib = comtypes.client.GetModule(str(_dir / "ActProgType64.exe"))

from .ActDefine import *

logger = logging.getLogger(__package__)


def _patch(interface):
    methods = interface._methods_
    for method_i, item in enumerate(methods):
        restype, name, argtypes, paramflags, idlflags, doc = item
        for prefix in ("_get_", "_set_"):
            # 새로 정의될 수 있도록 기존 이름을 지운다.
            # 지우지 않으면 '_'이 덧붙여진다.
            _name = name if not name.startswith(prefix) else name.removeprefix(prefix)
            if hasattr(interface, _name):
                delattr(interface, _name)
        if name.startswith("Read"):
            paramflags = list(paramflags)
            for i, paramflag in enumerate(paramflags):
                pflags, argname, *defval = paramflag
                # 배열의 크기가 반영되지 않으므로
                # 버퍼를 직접 전달할 수 있도록 in으로 바꾼다.
                if argname in {"lplData", "lpsData"} and pflags == 2:
                    pflags = 1  # out -> in
                    paramflags[i] = (pflags, argname, *defval)
            paramflags = tuple(paramflags)
            methods[method_i] = (restype, name, argtypes, paramflags, idlflags, doc)
    interface._methods_ = methods  # 다시 만든다.
_patch(ActUtlType64Lib.IActUtlType64)
_patch(ActProgType64Lib.IActProgType64)
del _patch


__all__ = ["ActUtlType", "ActProgType"]


def act_get_error_message(error_code):
    obj = comtypes.client.CreateObject(ActSupportMsg64Lib.ActSupportMsg64)
    error_message, retval = obj.GetErrorMessage(error_code)
    if retval != 0:
        raise Exception(f"cannot get error message: {retval:08x}")
    return error_message


class ActError(Exception):
    def __init__(self, error_code: int, error_message: str = None):
        self.error_code = error_code
        if error_message is None:
            error_message = act_get_error_message(error_code)
        super().__init__(error_message)


class ActCommon:
    def __init__(self, *, error_check=True):
        self._object = None
        self.error_check = error_check

    def _call(self, method, *args, **kwds):
        retval = method(*args, **kwds)
        if self.error_check:
            error_code = retval[-1] if isinstance(retval, tuple) else retval
            if error_code != 0:
                raise ActError(error_code)
        return retval

    def close(self):
        obj = self._object
        self._call(obj.Close)

    def get_device(self, device):
        obj = self._object
        value, retval = self._call(obj.GetDevice2, device)
        return value

    def set_device(self, device, value):
        obj = self._object
        retval = self._call(obj.SetDevice2, device, value)

    def get_cpu_type(self):
        obj = self._object
        cpu_name, cpu_code, retval = self._call(obj.GetCpuType)
        return cpu_name, cpu_code

    def read_device_random2(self, device_list):
        device_count = len(device_list)
        data = create_string_buffer(device_count * 2)
        retval = self._call(self._object.ReadDeviceRandom2, "\n".join(device_list), device_count, cast(data, POINTER(c_short)))
        return data.raw

    def read_device_random(self, device_list):
        device_count = len(device_list)
        data = create_string_buffer(device_count * 4)
        retval = self._call(self._object.ReadDeviceRandom, "\n".join(device_list), device_count, cast(data, POINTER(c_int)))
        return data.raw

    def write_device_block2(self, device, data):
        device_count = len(data) // 2
        # TODO: 읽기만 하면 되는데...
        if isinstance(data, bytes):
            data = bytearray(data)
        data = (c_short*device_count).from_buffer(data)
        retval = self._call(self._object.WriteDeviceBlock2, device, device_count, data)

    def write_device_block(self, device, data):
        # 비트 디바이스인 시작 디바이스 번호는 16의 배수여야 한다.
        device_count = len(data) // 4
        data = create_string_buffer(data, len(data))
        data_p = cast(data, POINTER(c_int))
        retval = self._call(self._object.WriteDeviceBlock, device, device_count, data_p)

    get_error_message = staticmethod(act_get_error_message)


class ActUtlType(ActCommon):
    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._object = comtypes.client.CreateObject(ActUtlType64Lib.ActUtlType64)

    def open(self, stno: int):
        obj = self._object
        obj.ActLogicalStationNumber = stno
        self._call(obj.Open)
        return self


class ActProgType(ActCommon):
    def __init__(self, **kwds):
        super().__init__(**kwds)
        self._object = comtypes.client.CreateObject(ActProgType64Lib.ActProgType64)

    def open_simulator2(self, target = 0):
        """
        Specify the connection destination GX Simulator2 in start status.
        When connecting to FXCPU, specify "0" (0x00).
        ■Property value
            0 (0x00): None
            (When only one simulator is in start status, connects to the simulator in start status. When
            multiple simulators are in start status, search for the simulators in start status and connect them in
            alphabetical order.)
            1 (0x01): Simulator A
            2 (0x02): Simulator B
            3 (0x03): Simulator C
            4 (0x04): Simulator D
        ---
        Specify the PLC number of the connection destination GX Simulator3 in start status.
        ---
        Specify the connection destination MT Simulator2 in start status.
        ■Property value
            2 (0x02): Simulator No.2
            3 (0x03): Simulator No.3
            4 (0x04): Simulator No.4
        """
        obj = self._object
        obj.ActUnitType = UNIT_SIMULATOR2
        obj.ActTargetSimulator = target
        self._call(obj.Open)
        return self

import ctypes
from ctypes import byref, POINTER, WINFUNCTYPE
from ctypes.wintypes import LONG, PLONG, WCHAR, PSHORT

import comtypes

__all__ = ("ActUtlType", "ActProgType")

HRESULT = ctypes.c_int32
OLECHAR = WCHAR
BSTR = POINTER(OLECHAR)

GET_LONG = WINFUNCTYPE(HRESULT, PLONG)
PUT_LONG = WINFUNCTYPE(HRESULT, LONG)


class ActCommon:
    CLSID = None
    IID = None

    _QueryInterface = WINFUNCTYPE(HRESULT, POINTER(comtypes.GUID), POINTER(ctypes.c_void_p))(0, "QueryInterface")
    _AddRef = WINFUNCTYPE(ctypes.c_ulong)(1, "AddRef")
    _Release = WINFUNCTYPE(ctypes.c_ulong)(2, "Release")

    def __init__(self):
        if not self.CLSID:
            raise NotImplementedError()
        self.pv = ctypes.c_void_p()
        comtypes._ole32.CoCreateInstance(byref(self.CLSID), None, comtypes.CLSCTX_LOCAL_SERVER, byref(self.IID), byref(self.pv))

    def __del__(self):
        if self.pv:
            self._Release(self.pv)
            self.pv = None

    def Open(self):
        return_code = LONG()
        hr = self._Open(self.pv, byref(return_code))
        print(hr, return_code)

    def Close(self):
        return_code = LONG()
        hr = self._Close(self.pv, byref(return_code))
        print(hr, return_code)

    def ReadDeviceBlock2(self, device: str, size: int) -> bytes:
        return_code = LONG()
        hr = self._ReadDeviceBlock2(self.pv, byref(return_code))
        print(hr, return_code)

    def WriteDeviceBlock2(self, device: str, size: int, data: bytes):
        return_code = LONG()
        hr = self._WriteDeviceBlock2(self.pv, byref(return_code))
        print(hr, return_code)


        # virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ReadDeviceBlock2( 
        #     /* [string][in] */ BSTR szDevice,
        #     /* [in] */ LONG lSize,
        #     /* [size_is][out] */ SHORT *lplData,
        #     /* [retval][out] */ LONG *lplReturnCode) = 0;
        
        # virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE WriteDeviceBlock2( 
        #     /* [string][in] */ BSTR szDevice,
        #     /* [in] */ LONG lSize,
        #     /* [size_is][in] */ SHORT *lplData,
        #     /* [retval][out] */ LONG *lplReturnCode) = 0;


class ActUtlType(ActCommon):
    CLSID = comtypes.GUID.from_progid("ActUtlType64.ActUtlWrap")
    IID = comtypes.GUID("{174DD3F4-961E-4833-A4D2-6BFFE5DDFC6C}")

    get_ActLogicalStationNumber = GET_LONG(3, "get_ActLogicalStationNumber")
    put_ActLogicalStationNumber = PUT_LONG(4, "put_ActLogicalStationNumber")
    # (5, "get_ActPassword")
    # (6, "put_ActPassword")
    _Open = WINFUNCTYPE(HRESULT, PLONG)(7, "Open")
    _Close = WINFUNCTYPE(HRESULT, PLONG)(8, "Close")
    _ReadDeviceBlock2 = WINFUNCTYPE(HRESULT, BSTR, LONG, PSHORT, PLONG)(9, "ReadDeviceBlock2")
    _WriteDeviceBlock2 = WINFUNCTYPE(HRESULT, BSTR, LONG, PSHORT, PLONG)(10, "WriteDeviceBlock2")

    @property
    def ActLogicalStationNumber(self):
        value = LONG()
        self.get_ActLogicalStationNumber(self.pv, byref(value))
        return value.value

    @ActLogicalStationNumber.setter
    def ActLogicalStationNumber(self, value):
        self.put_ActLogicalStationNumber(self.pv, LONG(value))


class ActProgType(ActCommon):
    pass


if __name__ == "__main__":
    actutl = ActUtlType()
    actutl.ActLogicalStationNumber = 2
    actutl.Open()
    actutl.Close()

# MX Component Ver.5

```python
from mxcomponent5 import ActUtlType, connect_qcpu_ethernet


def test(control):
    print(control)
    control.Open()
    x = control.GetCpuType()
    print(x)
    x = control.GetDevice("ZR3000")
    print(x)
    x = control.GetDevice("L600")
    print(x)
    x = control.GetClockData()
    print(x)
    x = control.ReadDeviceBlock("ZR3000", 10)
    print(x)
    control.Close()


control = ActUtlType()
control.ActLogicalStationNumber = 10
test(control)

control = connect_qcpu_ethernet("192.168.0.100")
test(control)
```

MX Component V5 支持 64 位。然而，它没有实现支持 IDispatch 的 ...ML... 接口。
因此，必须使用 VTable 接口。
在 Read... 系列方法中，与数据相关的参数类型为 SHORT* 或 LONG*，并且是 out 参数。
实际上，需要传递与 size 相同大小的数组，但自动生成的代码总是只为一个数据分配空间。
_act5 是一个使用 pybind11 的扩展模块，基本上依赖于此模块。
_act5_comtypes.py 使用 comtypes 库生成包装代码，并将 Read... 系列方法中与数据相关的参数从 out 修改为 in，
以便外部可以直接传递数据。
_act5_ctypes.py 是直接编写的包装代码。
虽然有些繁琐，并且在 MX Component 升级时可能会失效，但直接处理 COM 接口本身具有一定的意义。
目前仍处于实验阶段。

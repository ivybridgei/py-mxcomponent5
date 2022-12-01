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

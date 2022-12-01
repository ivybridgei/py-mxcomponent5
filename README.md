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

MX Component V5는 64비트를 지원한다. 하지만 IDispatch를 구현하는 ...ML... 인터페이스가 없다.
그래서 VTable 인터페이스를 사용해야 하는데,
Read...류 메소드의 데이터 관련 파라메터 형이 SHORT* 혹은 LONG* 이고 out이라
실제로는 size 만큼의 배열이 전달되어야 하지만,
자동 생성되는 코드는 항상 데이터가 하나 뿐인 것처럼 공간을 확보한다.

_act5는 pybind11를 사용한 확장으로 기본적으로 이 모듈을 사용한다.

_act5_comtypes.py는 comtypes 라이브러리를 사용해서 일단 감싸는 코드를 생성하고,
Read...류 메소드의 데이터 관련 파라메터를 out이 아니라 in으로 고쳐서
외부에서 데이터를 직접 전달할 수 있도록 했다.

_act5_ctypes.py는 감싸는 코드를 직접 작성한다.
다소 번거롭고, MX Component가 업그레이드 되면 깨질 가능성도 있지만,
그럼에도 COM 인터페이스를 좀 더 직접적으로 다뤄본다는 점에는 나름 의미가 있다.
아직은 실험을 해 보는 단계다.

import logging
from typing import Any

from ._act5 import *
from .ActDefine import *

__all__ = (
    "connect_simulator2",
    "connect_qcpu_ethernet",
)

logger: logging.Logger = logging.getLogger(__package__)


def connect_simulator2(target=0, **kwds):
    """
    :param:target
        Specify the connection destination GX Simulator2 in start status.
        When connecting to FXCPU, specify "0" (0x00).
        ■Property value
        0 (0x00): None
        (When only one simulator is in start status, connects to the simulator in start status.
        When multiple simulators are in start status,
        search for the simulators in start status and connect them in alphabetical order.)
        1 (0x01): Simulator A
        2 (0x02): Simulator B
        3 (0x03): Simulator C
        4 (0x04): Simulator D
        - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        Specify the PLC number of the connection destination GX Simulator3 in start status.
        - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - 
        Specify the connection destination MT Simulator2 in start status.
        ■Property value
        2 (0x02): Simulator No.2
        3 (0x03): Simulator No.3
        4 (0x04): Simulator No.4
    """
    control = ActProgType()
    _set_cpu_type(control, kwds)
    control.ActTargetSimulator = target
    control.ActUnitType = UNIT_SIMULATOR2
    return control


def connect_qcpu_ethernet(host: str, protocol="udp", port: int = None, **kwds):
    control = ActProgType()
    _set_cpu_type(control, kwds)
    control.ActHostAddress = host
    control.ActDestinationPortNumber = port if port is not None else (5007 if protocol.lower() == "tcp" else 5006)
    control.ActProtocolType = PROTOCOL_TCPIP if protocol.lower() == "tcp" else PROTOCOL_UDPIP
    control.ActUnitType = UNIT_QNETHER
    return control


def _set_cpu_type(control: ActProgType, kwds: dict[str, Any]):
    cpu = kwds.pop("cpu", None)
    if cpu:
        if isinstance(cpu, str):
            if cpu.upper() not in cpu_name_to_code_table:
                logger.warning(f"unknown cpu: {cpu}")
                return
            cpu = cpu_name_to_code_table[cpu.upper()]
        control.ActCpuType = cpu

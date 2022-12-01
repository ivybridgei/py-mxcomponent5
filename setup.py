from pathlib import Path
from setuptools import setup
from pybind11.setup_helpers import Pybind11Extension


def _get_act_dir(*, cache=[None]) -> Path:
    if cache[0] is None:
        from winreg import HKEY_CLASSES_ROOT, OpenKeyEx, QueryValue, KEY_READ, KEY_WOW64_32KEY
        clsid = QueryValue(HKEY_CLASSES_ROOT, "ActUtlType.ActUtlType\\CLSID")
        clsid_key = OpenKeyEx(HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, KEY_READ | KEY_WOW64_32KEY)
        cache[0] = Path(QueryValue(clsid_key, "InprocServer32")).parents[1]
    return cache[0]


def _get_act_file_path(lst: list, dir=""):
    _act_dir = _get_act_dir()
    return [str(_act_dir.joinpath(f"{dir}{filename}")) for filename in lst]


act_sources = _get_act_file_path(["ActProgType64_i.c", "ActSupportMsg64_i.c", "ActUtlType64_i.c"], "Include/Wrapper/")
act_include_dir = str(_get_act_dir().joinpath("Include"))

ext_modules = [
    Pybind11Extension("mxcomponent5._act5",
        sources = ["_act5.cpp", *act_sources],
        include_dirs = [act_include_dir],
        libraries = ["Ole32", "OleAut32"],
        extra_compile_args = ["/utf-8", "/std:c++17"],
    ),
]

setup(
    ext_modules = ext_modules,
)

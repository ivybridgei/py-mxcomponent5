import io
from pathlib import Path


def _get_act_dir(*, cache=[None]) -> Path:
    if cache[0] is None:
        from winreg import HKEY_CLASSES_ROOT, OpenKeyEx, QueryValue, KEY_READ, KEY_WOW64_32KEY
        clsid = QueryValue(HKEY_CLASSES_ROOT, "ActUtlType.ActUtlType\\CLSID")
        clsid_key = OpenKeyEx(HKEY_CLASSES_ROOT, f"CLSID\\{clsid}", 0, KEY_READ | KEY_WOW64_32KEY)
        cache[0] = Path(QueryValue(clsid_key, "InprocServer32")).parents[1]
    return cache[0]


def gen_ActDefine_py():
    ActDefine_cs = _get_act_dir().joinpath("Include", "ActDefine.cs")
    lines = []
    with open(ActDefine_cs, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line.startswith("//"):
                if line.startswith("///"):
                    comment = line.removeprefix("///").strip()
                    lines.append(f"### {comment}\n")
                else:
                    comment = line.removeprefix("//").strip()
                    lines.append(f"# {comment}\n")
            if line.startswith("public const int"):
                name, value_comment = line.removeprefix("public const int").split("=")
                value, comment = value_comment.split(";")
                comment = comment.replace("//", "")
                lines.append(tuple(x.strip() for x in (name, value, comment)))
            if not line:
                lines.append("\n")

    max_name = 0
    max_value = 0
    max_cpu_name = 0
    cpu_list: list[str] = []
    for line in lines:
        if len(line) == 3:
            line: tuple[str, str, str]
            name, value, comment = line
            if len(name) > max_name:
                max_name = len(name)
            if len(value) > max_value:
                max_value = len(value)
            if name.startswith("CPU_"):
                if len(name) > max_cpu_name:
                    max_cpu_name = len(name)
                cpu_list.append(name)
    max_name += 3
    max_value += 3

    sb = io.StringIO()
    for line in lines:
        if len(line) == 3:
            name, value, comment = line
            define = f"{name:{max_name}} = {value:>{max_value}}  # {comment}"
            sb.write(f"{define}\n")
        else:
            sb.write(line)

    sb.write("cpu_name_to_code_table = {\n")
    for cpu in cpu_list:
        cpu_name = cpu.removeprefix("CPU_")
        key_name = f'"{cpu_name}"'
        sb.write(f"    {key_name:{max_cpu_name}}: {cpu},\n")
        if "CPU" in cpu_name:
            short_name = "_".join(x.removesuffix("CPU") for x in cpu_name.split("_"))
            if short_name != cpu_name:
                key_name = f'"{short_name}"'
                sb.write(f"    {key_name:{max_cpu_name}}: {cpu},\n")
    sb.write("}\n")

    with open("mxcomponent5/ActDefine.py", "w", encoding="utf-8") as f:
        f.write(sb.getvalue())


if __name__ == "__main__":
    gen_ActDefine_py()
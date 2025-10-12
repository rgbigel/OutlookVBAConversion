import re

builtin_map = {
    "len(": "len(",
    "ucase(": ".upper(",
    "lcase(": ".lower(",
    "trim(": ".strip(",
    "ltrim(": ".lstrip(",
    "rtrim(": ".rstrip(",
    "replace(": ".replace(",
    "instr(": ".find(",
    "instrrev(": ".rfind(",
    "strreverse(": "[::-1]",
    "space(": '" " * ',
    "string(": "*",  # handled separately
    "isnumeric(": ".isnumeric(",
    "isdate(": "isinstance(",
    "isempty(": "x is None or x == \"\"",
    "isnull(": "x is None",
    "msgbox": "print",
    "debug.print": "print",
    "vbcrlf": '"\\r\\n"',
}

def convert_vba_to_python(vba_code: str, filename: str) -> str:
    lines = vba_code.splitlines()
    py_lines = [f"# Converted from {filename}", ""]
    indent = "    "
    block_indent = ""
    in_try_block = False

    for line in lines:
        original = line.strip()
        lower = original.lower()

        if not original:
            py_lines.append("")
            continue

        # Sub / Function
        if lower.startswith("sub ") or lower.startswith("function "):
            name = re.findall(r"(sub|function)\s+(\w+)", lower)
            if name:
                py_lines.append(f"def {name[0][1]}():")
                block_indent = indent
            continue

        # End Sub / Function
        if lower.startswith("end sub") or lower.startswith("end function"):
            block_indent = ""
            continue

        # Class
        if lower.startswith("class "):
            name = re.findall(r"class\s+(\w+)", original)
            if name:
                py_lines.append(f"class {name[0]}:")
                py_lines.append(f"{indent}def __init__(self):")
                py_lines.append(f"{indent*2}pass")
            continue

        # Dim / Set
        if lower.startswith("dim ") or lower.startswith("set "):
            match = re.findall(r"(dim|set)\s+(\w+)(?:\s+as\s+(\w+))?", lower)
            if match:
                var = match[0][1]
                vtype = match[0][2].lower() if match[0][2] else ""
                default = '""' if vtype == "string" else "0" if vtype == "integer" else "None"
                py_lines.append(f"{block_indent}{var} = {default}")
            continue

        # If / ElseIf / Else
        if lower.startswith("if "):
            condition = original[3:].split("then")[0].strip()
            py_lines.append(f"{block_indent}if {condition}:")
            continue
        if lower.startswith("elseif "):
            condition = original[7:].split("then")[0].strip()
            py_lines.append(f"{block_indent}elif {condition}:")
            continue
        if lower.startswith("else"):
            py_lines.append(f"{block_indent}else:")
            continue
        if lower.startswith("end if"):
            continue

        # For / Next
        if lower.startswith("for each "):
            match = re.findall(r"for each (\w+) in (\w+)", lower)
            if match:
                py_lines.append(f"{block_indent}for {match[0][0]} in {match[0][1]}:")
            continue
        if lower.startswith("for "):
            match = re.findall(r"for (\w+) = (\d+) to (\d+)", lower)
            if match:
                py_lines.append(f"{block_indent}for {match[0][0]} in range({match[0][1]}, {int(match[0][2])+1}):")
            continue
        if lower.startswith("next"):
            continue

        # Select Case
        if lower.startswith("select case "):
            var = original.split(" ")[-1]
            py_lines.append(f"{block_indent}match {var}:")
            continue
        if lower.startswith("case "):
            value = original[5:].strip()
            if value.lower() == "else":
                py_lines.append(f"{block_indent}    case _:")
            else:
                py_lines.append(f"{block_indent}    case {value}:")
            continue
        if lower.startswith("end select"):
            continue

        # On Error
        if lower.startswith("on error"):
            py_lines.append(f"{block_indent}try:")
            block_indent += indent
            in_try_block = True
            continue
        if lower.startswith("resume next") and in_try_block:
            py_lines.append(f"{block_indent}pass  # resume next")
            continue

        # Built-in replacements
        replaced = original
        for vba, py in builtin_map.items():
            if vba in lower:
                replaced = re.sub(vba, py, replaced, flags=re.IGNORECASE)

        # Left / Mid / Right
        replaced = re.sub(r"Left\(([^,]+),\s*(\d+)\)", r"\1[:\2]", replaced, flags=re.IGNORECASE)
        replaced = re.sub(r"Right\(([^,]+),\s*(\d+)\)", r"\1[-\2:]", replaced, flags=re.IGNORECASE)
        replaced = re.sub(r"Mid\(([^,]+),\s*(\d+),\s*(\d+)\)", lambda m: f"{m.group(1)}[{int(m.group(2))-1}:{int(m.group(2))-1+int(m.group(3))}]", replaced, flags=re.IGNORECASE)

        # String(n, c)
        replaced = re.sub(r"String\((\d+),\s*\"(.)\"\)", r'"\2" * \1', replaced, flags=re.IGNORECASE)

        # COM references
        if "application" in lower or "namespace" in lower:
            py_lines.insert(1, "from com_wrappers import get_outlook_namespace, create_mail")
            replaced = replaced.replace("Application", "get_outlook_namespace()")
            replaced = replaced.replace("Namespace", "get_outlook_namespace()")

        py_lines.append(f"{block_indent}{replaced}")

    return "\r\n".join(py_lines)

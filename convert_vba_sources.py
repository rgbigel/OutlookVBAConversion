import os
import re
from datetime import datetime

# --- Parser: Enhanced VBA to Python ---
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

        # MsgBox / Debug.Print
        if "msgbox" in lower:
            msg = re.findall(r'msgbox\s+"([^"]+)"', original, re.IGNORECASE)
            if msg:
                py_lines.append(f"{block_indent}print('{msg[0]}')")
            continue
        if "debug.print" in lower:
            msg = original.split("debug.print")[-1].strip()
            py_lines.append(f"{block_indent}print({msg})")
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

        # Default: comment original line
        py_lines.append(f"{block_indent}# {original}")

    return "\r\n".join(py_lines)

# --- Conversion Runner ---
source_dir = r"D:\GitHubPrivate\OutlookVBAConversion\VBA_Sources"
output_dir = r"D:\GitHubPrivate\OutlookVBAConversion\scripts"
forms_dir = os.path.join(output_dir, "forms")
log_path = os.path.join(output_dir, "conversion_log.txt")

excluded_files = {"bughelp.bas", "ThisOutlookSession.cls"}

os.makedirs(output_dir, exist_ok=True)
os.makedirs(forms_dir, exist_ok=True)

converted_modules = {}

with open(log_path, "w", encoding="utf-8") as log:
    log.write(f"Conversion Log - {datetime.now().isoformat()}\n\n")
    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.lower().endswith((".bas", ".cls", ".frm")) and file not in excluded_files:
                full_path = os.path.join(root, file)
                with open(full_path, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()
                py_name = file.replace(".bas", ".py").replace(".cls", ".py").replace(".frm", ".py")
                py_code = convert_vba_to_python(content, py_name)
                target_dir = forms_dir if file.endswith(".frm") else output_dir
                target_path = os.path.join(target_dir, py_name)
                with open(target_path, "w", newline="\r\n", encoding="utf-8") as out:
                    out.write(py_code)
                log.write(f"[✓] {file} → {target_path}\n")
                converted_modules[py_name] = py_code

print("✅ Phase 1 complete: VBA converted to Python with CRLF endings and full logging.")

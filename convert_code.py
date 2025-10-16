import os
import re
import tkinter as tk
import argparse

excluded = {"BugHelp", "VbaProject.OTM", "Y_Errlf"}
def is_excluded(file):
    name = os.path.splitext(file)[0]
    return name in excluded or name.startswith("ErrEx")

def find_vba_files(source_dir):
    vba_files = []
    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.endswith((".bas", ".cls", ".frm")) and not is_excluded(file):
                vba_files.append(os.path.join(root, file))
    return vba_files

def flag_loop(line):
    match = re.match(r'For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+)', line.strip(), re.IGNORECASE)
    if match:
        var, start, end = match.groups()
        return [
            f"# originally: {line.strip()}",
            f"# consider replacing with: for {var} in range({start}, {end} + 1):"
        ]
    return None

def convert_bas_module(lines):
    py_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Attribute") or stripped == "":
            continue
        loop_flag = flag_loop(line)
        if loop_flag:
            py_lines.extend(loop_flag)
        elif stripped.startswith(("Public", "Private")):
            visibility = "Public" if "Public" in stripped else "Private"
            py_lines.append(f"# originally {visibility}")
            func = re.sub(r'\b(Public|Private)\b', 'def', stripped)
            func = re.sub(r'As\s+\w+', '', func)
            py_lines.append(func)
        else:
            py_lines.append(f"# {line}")
    return py_lines

def convert_cls_module(lines, module_name):
    py_lines = [f"class {module_name}:", "    def __init__(self):", "        pass\n"]
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Attribute") or stripped == "":
            continue
        loop_flag = flag_loop(line)
        if loop_flag:
            py_lines.extend([f"    {l}" for l in loop_flag])
        elif stripped.startswith(("Public", "Private")):
            visibility = "Public" if "Public" in stripped else "Private"
            py_lines.append(f"    # originally {visibility}")
            func = re.sub(r'\b(Public|Private)\b', 'def', stripped)
            func = re.sub(r'As\s+\w+', '', func)
            py_lines.append(f"    {func}")
        else:
            py_lines.append(f"    # {line}")
    return py_lines

def convert_frm_module(lines, module_name):
    py_lines = [f"class {module_name}(tk.Tk):", "    def __init__(self):", f"        super().__init__()", f"        self.title(\"{module_name}\")", "        # TODO: Add widgets\n"]
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Begin"):
            py_lines.append(f"        # {line} → widget start")
        elif stripped.startswith("End"):
            py_lines.append(f"        # {line} → widget end")
        elif stripped.startswith(("Caption", "Name", "Text", "Value")):
            py_lines.append(f"        # {line}")
        elif stripped.startswith(("Public", "Private", "Sub", "Function")):
            visibility = "Public" if "Public" in stripped else "Private"
            py_lines.append(f"    # originally {visibility}")
            func = re.sub(r'\b(Public|Private|Sub|Function)\b', 'def', stripped)
            func = re.sub(r'As\s+\w+', '', func)
            py_lines.append(f"    {func}")
        else:
            loop_flag = flag_loop(line)
            if loop_flag:
                py_lines.extend([f"    {l}" for l in loop_flag])
            else:
                py_lines.append(f"    # {line}")
    return py_lines

def convert_vba_to_python(vba_code, filename):
    lines = vba_code.splitlines()
    module_name = os.path.splitext(os.path.basename(filename))[0]
    ext = os.path.splitext(filename)[1].lower()

    if ext == ".bas":
        return "\n".join(convert_bas_module(lines))
    elif ext == ".cls":
        return "\n".join(convert_cls_module(lines, module_name))
    elif ext == ".frm":
        return "\n".join(convert_frm_module(lines, module_name))
    else:
        return "# Unsupported file type"

def convert_code(source_dir, output_dir):
    os.makedirs(output_dir, exist_ok=True)
    vba_files = find_vba_files(source_dir)

    for vba_path in vba_files:
        with open(vba_path, "r", encoding="utf-8") as f:
            vba_code = f.read()

        filename = os.path.basename(vba_path)
        py_code = convert_vba_to_python(vba_code, filename)

        py_filename = os.path.splitext(filename)[0] + ".py"
        py_path = os.path.join(output_dir, py_filename)

        with open(py_path, "w", encoding="utf-8") as f:
            f.write(py_code)

        print(f"Converted: {filename} → {py_filename}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert VBA .bas, .cls, and .frm files to Python stubs.")
    parser.add_argument("source_dir", nargs="?", default="VBA_Source", help="Directory containing VBA source files")
    parser.add_argument("output_dir", nargs="?", default="Converted_Python", help="Directory to write converted Python files")

    args = parser.parse_args()

    print(f"Using source_dir: {args.source_dir}")
    print(f"Using output_dir: {args.output_dir}")

    convert_code(args.source_dir, args.output_dir)

import os
import re

def find_vba_files(source_dir):
    """
    Recursively find all .bas, .cls, and .frm files in the given directory.
    """
    vba_files = []
    for root, _, files in os.walk(source_dir):
        for file in files:
            if file.endswith((".bas", ".cls", ".frm")):
                vba_files.append(os.path.join(root, file))
    return vba_files

def convert_bas_module(lines):
    """
    Convert .bas module lines to Python.
    """
    py_lines = []
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Attribute") or stripped == "":
            continue
        if stripped.startswith(("Public", "Private")):
            func = re.sub(r'\b(Public|Private)\b', 'def', stripped)
            func = re.sub(r'As\s+\w+', '', func)
            py_lines.append(func)
        else:
            py_lines.append(f"# {line}")
    return py_lines

def convert_cls_module(lines, module_name):
    """
    Convert .cls class module lines to Python class.
    """
    py_lines = [f"class {module_name}:", "    def __init__(self):", "        pass\n"]
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Attribute") or stripped == "":
            continue
        if stripped.startswith(("Public", "Private")):
            func = re.sub(r'\b(Public|Private)\b', 'def', stripped)
            func = re.sub(r'As\s+\w+', '', func)
            py_lines.append(f"    {func}")
        else:
            py_lines.append(f"    # {line}")
    return py_lines

def convert_frm_module(lines, module_name):
    """
    Convert .frm form module lines to Python tkinter stub.
    """
    py_lines = [f"class {module_name}(tk.Tk):", "    def __init__(self):", "        super().__init__()", "        self.title(\"{0}\")".format(module_name), "        # TODO: Add widgets\n"]
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("Begin"):
            py_lines.append(f"        # {line} → widget start")
        elif stripped.startswith("End"):
            py_lines.append(f"        # {line} → widget end")
        elif stripped.startswith(("Caption", "Name", "Text", "Value")):
            py_lines.append(f"        # {line}")
        elif stripped.startswith(("Public", "Private", "Sub", "Function")):
            func = re.sub(r'\b(Public|Private|Sub|Function)\b', 'def', stripped)
            func = re.sub(r'As\s+\w+', '', func)
            py_lines.append(f"    {func}")
        else:
            py_lines.append(f"    # {line}")
    return py_lines

def convert_vba_to_python(vba_code, filename):
    """
    Dispatch conversion based on file extension.
    """
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
    """
    Convert all eligible VBA files (.bas, .cls, .frm) from source_dir to Python scripts in output_dir.
    """
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
    import argparse

    parser = argparse.ArgumentParser(description="Convert VBA .bas, .cls, and .frm files to Python stubs.")
    parser.add_argument("source_dir", help="Directory containing VBA source files")
    parser.add_argument("output_dir", help="Directory to write converted Python files")

    args = parser.parse_args()
    convert_code(args.source_dir, args.output_dir)

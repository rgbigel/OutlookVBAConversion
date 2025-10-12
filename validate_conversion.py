import os
import ast

def find_python_files(input_dir):
    """
    Recursively find all .py files in the given directory.
    """
    py_files = []
    for root, _, files in os.walk(input_dir):
        for file in files:
            if file.endswith(".py"):
                py_files.append(os.path.join(root, file))
    return py_files

def check_syntax(py_path):
    """
    Check if the Python file is syntactically valid.
    """
    try:
        with open(py_path, "r", encoding="utf-8") as f:
            source = f.read()
        ast.parse(source)
        return True, []
    except SyntaxError as e:
        return False, [f"{e.msg} at line {e.lineno}"]

def check_structure(py_path):
    """
    Check for structural completeness based on filename and content.
    """
    issues = []
    filename = os.path.basename(py_path)
    module_name = os.path.splitext(filename)[0]
    ext_hint = module_name[:3].lower()

    with open(py_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    content = "".join(lines)

    if ext_hint == "frm":
        if f"class {module_name}(tk.Tk)" not in content:
            issues.append("Missing tkinter subclass")
        if "self.title(" not in content:
            issues.append("Missing window title")
    elif ext_hint == "cls":
        if f"class {module_name}" not in content:
            issues.append("Missing class definition")
    elif ext_hint == "mod" or ext_hint == "bas":
        if "def " not in content:
            issues.append("No function definitions found")

    raw_vba_keywords = ["Sub ", "Function ", "End ", "Dim ", "Set "]
    for keyword in raw_vba_keywords:
        if keyword in content:
            issues.append(f"Contains raw VBA keyword: {keyword.strip()}")

    return issues

def validate_conversion(input_dir):
    """
    Validate all converted Python files for syntax and structure.
    """
    py_files = find_python_files(input_dir)
    total = len(py_files)
    passed = 0
    flagged = {}

    for py_path in py_files:
        syntax_ok, syntax_issues = check_syntax(py_path)
        structure_issues = check_structure(py_path)

        if syntax_ok and not structure_issues:
            passed += 1
        else:
            flagged[py_path] = syntax_issues + structure_issues

    print(f"\n✅  Checked {total} files")
    print(f"✅  Passed: {passed}")
    print(f"⚠️  Flagged: {len(flagged)}")

    for path, issues in flagged.items():
        print(f"\n⚠️  {os.path.basename(path)}")
        for issue in issues:
            print(f"   - {issue}")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Validate converted Python files for syntax and structure.")
    parser.add_argument("input_dir", help="Directory containing converted Python files")

    args = parser.parse_args()
    validate_conversion(args.input_dir)

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

def extract_doc_info(py_path):
    """
    Extract class and function names from a Python file.
    """
    with open(py_path, "r", encoding="utf-8") as f:
        source = f.read()

    try:
        tree = ast.parse(source)
    except SyntaxError:
        return None  # Skip invalid files

    classes = []
    functions = []

    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            classes.append(node.name)
        elif isinstance(node, ast.FunctionDef):
            functions.append(node.name)

    return {
        "filename": os.path.basename(py_path),
        "classes": classes,
        "functions": functions
    }

def generate_markdown(doc_info):
    """
    Generate Markdown content from extracted doc info.
    """
    lines = [f"# {doc_info['filename']}\n"]

    if doc_info["classes"]:
        lines.append("## Classes")
        for cls in doc_info["classes"]:
            lines.append(f"- `{cls}`")

    if doc_info["functions"]:
        lines.append("\n## Functions")
        for func in doc_info["functions"]:
            lines.append(f"- `{func}()`")

    return "\n".join(lines)

def generate_docs(input_dir, output_dir):
    """
    Generate Markdown documentation for all Python files in input_dir.
    """
    os.makedirs(output_dir, exist_ok=True)
    py_files = find_python_files(input_dir)

    for py_path in py_files:
        doc_info = extract_doc_info(py_path)
        if not doc_info:
            continue

        md_content = generate_markdown(doc_info)
        md_filename = os.path.splitext(doc_info["filename"])[0] + ".md"
        md_path = os.path.join(output_dir, md_filename)

        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md_content)

        print(f"📄 Documented: {doc_info['filename']} → {md_filename}")

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Generate Markdown docs from converted Python files.")
    parser.add_argument("input_dir", help="Directory containing converted Python files")
    parser.add_argument("output_dir", help="Directory to write Markdown documentation")

    args = parser.parse_args()
    generate_docs(args.input_dir, args.output_dir)

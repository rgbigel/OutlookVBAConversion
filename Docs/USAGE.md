# OutlookVBAConversion — Usage Guide

This guide documents how to use the CLI tools in the OutlookVBAConversion pipeline. Each tool is modular and can be chained for batch processing, validation, and documentation.

---

## 🔁 Conversion Pipeline

### 1. Convert VBA to Python

```bash
python convert_code.py VBA_Sources scripts/

This command scans the VBA_Sources/ directory for .bas, .cls, and .frm files and converts them to Python .py files in the scripts/ directory.

.bas files → Python functions with type hints stripped

.cls files → Python classes with __init__ stubs

.frm files → tkinter subclasses with widget placeholders

Preserves original module names for traceability

### 2. Validate Conversion
bash
python validate_conversion.py scripts/
This step checks the converted .py files for syntax and structural integrity.

Validates Python syntax using ast

Flags raw VBA keywords (Sub, Function, Dim, etc.)

Ensures .cls files contain class definitions

Ensures .frm files subclass tk.Tk and include a window title

Reports missing wrappers or incomplete stubs

### 3. Generate Markdown Documentation
bash
python generate_docs.py scripts/ docs/
This command extracts class and function names from each .py file and generates a corresponding .md file in the docs/ directory.

One .md file per module

Lists all classes and functions

Skips files with syntax errors

🔗 Pipeline Chaining
You can chain all three steps into a single command:

bash
python convert_code.py VBA_Sources scripts/ \
  && python validate_conversion.py scripts/ \
  && python generate_docs.py scripts/ docs/
This ensures that only valid files are documented and errors are caught early.

### 🧪 Sample Output
bash
✅  Checked 12 files
✅  Passed: 10
⚠️  Flagged: 2

⚠️  modMath.py
   - Contains raw VBA keyword: Sub
   - No function definitions found

📄 Documented: cLogger.py → cLogger.md
📄 Documented: frmMain.py → frmMain.md
🛠️ Tips
Run validate_conversion.py before committing converted code

Use generate_docs.py to preview module structure

Add --fix flags in future versions to auto-wrap missing classes

Consider using a justfile or Makefile for repeatable builds

Extend the pipeline with GUI preview or docstring extraction

### 📁 Folder Conventions
VBA_Sources/ → Original .bas, .cls, .frm files

scripts/ → Converted Python .py files

docs/ → Generated Markdown documentation
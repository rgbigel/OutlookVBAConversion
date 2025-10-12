# OutlookVBAConversion

A modular pipeline for converting legacy Outlook VBA modules (`.bas`, `.cls`, `.frm`) into Pythonic equivalents — ready for automation, GUI scaffolding, and documentation. Built for transparency, batch processing, and iterative refinement.

---

## 🚀 Overview

OutlookVBAConversion is designed to:
- Convert `.bas`, `.cls`, and `.frm` files to Python stubs
- Preserve naming and structure for traceability
- Scaffold tkinter GUI classes from `.frm` files
- Validate syntax and structure post-conversion
- Generate Markdown documentation for each module

This project supports CLI chaining and modular extension — ideal for migrating large VBA codebases into maintainable Python workflows.

---
For full CLI usage and pipeline examples, see [USAGE.md](docs/USAGE.md).
---
## 🧱 Folder Structure

OutlookVBAConversion/ 
├── convert_code.py # Unified converter for .bas, .cls, .frm 
├── validate_conversion.py # Syntax and structure checker 
├── generate_docs.py # Markdown doc generator 
├── VBA_Sources/ # Original VBA files 
├── scripts/ # Converted Python files 
├── docs/ # Generated Markdown docs 
└── README.md


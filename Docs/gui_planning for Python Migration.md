# VBA to Python Mapping Guide

## 🎯 Purpose
This guide documents common VBA constructs and their Python equivalents to streamline the migration process and reduce ambiguity during refactoring.

---

## 🔄 Syntax Mapping

| VBA Syntax                  | Python Equivalent                    | Notes |
|----------------------------|--------------------------------------|-------|
| `Dim x As Integer`         | `x: int = 0`                         | Use type hints for clarity |
| `If x = 1 Then`            | `if x == 1:`                         | Python uses `==` and `:` |
| `ElseIf`                   | `elif`                               | No `Then` keyword |
| `For i = 1 To 10`          | `for i in range(1, 11):`             | `range` is exclusive of end |
| `Do While x < 5`           | `while x < 5:`                       | Python uses `while` directly |
| `Select Case x`            | `match x:` (Python 3.10+)            | Use `if/elif` for older versions |
| `MsgBox "Hello"`           | `print("Hello")` or `tkinter.messagebox` | GUI vs console |
| `Application.ScreenUpdating = False` | `# No direct equivalent` | Use GUI framework controls |
| `Set obj = CreateObject(...)` | `obj = win32com.client.Dispatch(...)` | Requires `pywin32` |

---

## 🧩 Object and Form Handling

| VBA Concept                | Python Equivalent                    | Notes |
|---------------------------|--------------------------------------|-------|
| `UserForm`                | `tkinter.Tk()` or `PyQt5.QDialog`    | Choose GUI framework |
| `Controls.Add`            | `tkinter.Button(...)` or similar     | GUI widgets |
| `ActiveSheet.Range("A1")`| `openpyxl` or `pandas`                | Excel automation |
| `Workbook.SaveAs`         | `wb.save("filename.xlsx")`           | `openpyxl` method |

---

## 🧪 Error Handling

| VBA                      | Python                                |
|--------------------------|----------------------------------------|
| `On Error Resume Next`   | `try: ... except: pass`                |
| `Err.Number`             | `except Exception as e:`              |
| `On Error GoTo`          | `try/except/finally`                  |

---

## 🔮 Migration Tips

- Use `type hints` and `docstrings` to clarify logic
- Modularize VBA macros into Python functions
- Replace global state with class-based design
- Use logging instead of `MsgBox` for diagnostics
- Consider `retry_logic.md` for robust error handling

---

## 🔗 Related Notes
- [[migration_overview]]
- [[retry_logic]]
- [[gui_planning]]

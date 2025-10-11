# Excel Automation in Python

## 🎯 Goal
Replicate and enhance VBA-based Excel macros using Python libraries for better control, logging, and modularity.

---

## 🧰 Recommended Libraries

| Library       | Use Case                      | Notes                         |
|---------------|-------------------------------|-------------------------------|
| `openpyxl`    | Read/write `.xlsx` files      | No VBA support, pure Python   |
| `xlwings`     | Control Excel via COM         | Requires Excel installed      |
| `pandas`      | Data manipulation             | Combine with `openpyxl`       |

---

## 🔄 VBA to Python Mapping

| VBA Macro                    | Python Equivalent                    |
|-----------------------------|--------------------------------------|
| `Range("A1").Value = 5`     | `ws["A1"].value = 5` (`openpyxl`)    |
| `ActiveSheet.Name = "Data"` | `ws.title = "Data"` (`openpyxl`)     |
| `Application.ScreenUpdating = False` | `xlwings.App.screen_updating = False` |

---

## 🧪 Sample Workflow

```python
import openpyxl

wb = openpyxl.load_workbook("report.xlsx")
ws = wb.active
ws["B2"].value = "Updated"
wb.save("report_updated.xlsx")

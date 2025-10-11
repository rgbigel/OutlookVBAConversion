# Logging Strategy for Python Automation

## 🎯 Purpose
Design a transparent, modular logging system to track script execution, errors, retries, and user interactions.

---

## 🧰 Logging Options

| Method        | Use Case                     | Notes                          |
|---------------|------------------------------|--------------------------------|
| `print()`     | Quick console feedback       | Not persistent                 |
| `logging`     | Structured, file-based logs  | Supports levels, rotation      |
| `custom GUI`  | Visual feedback in interface | Use with `tkinter` or `PyQt5`  |

---

## 🧩 Recommended Setup

```python
import logging

logging.basicConfig(
    filename="automation.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

logging.info("Script started")
logging.warning("Device not found")
logging.error("Recovery failed after 3 attempts")

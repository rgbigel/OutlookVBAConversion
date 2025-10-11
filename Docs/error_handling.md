
---

## 🛡️ `docs/error_handling.md`

```markdown
# Error Handling Patterns in Python Automation

## 🎯 Goal
Implement robust error handling to catch and respond to runtime issues without crashing scripts or losing state.

---

## 🧰 Patterns

| Pattern               | Use Case                     | Notes                          |
|-----------------------|------------------------------|--------------------------------|
| `try/except`          | Catch runtime errors         | Basic structure                |
| `try/except/finally`  | Cleanup after error          | Always runs `finally` block    |
| `raise`               | Propagate custom errors      | Use with custom exceptions     |

---

## 🧪 Example

```python
try:
    result = run_devcon_command()
    if not result.success:
        raise RuntimeError("DevCon failed")
except RuntimeError as e:
    logging.error(f"Recovery error: {e}")
finally:
    cleanup_temp_files()

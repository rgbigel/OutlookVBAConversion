# Retry Logic for Device Recovery Scripts

## 🧠 Purpose
Design a robust retry mechanism for hardware control scripts (e.g., GPU reset via DevCon) to handle transient failures, timing issues, and state inconsistencies.

## 🔁 Retry Strategy

### 1. **Trigger Conditions**
- DevCon returns non-zero exit code
- Device state remains unchanged after command
- Timeout exceeded without confirmation

### 2. **Retry Parameters**
| Parameter       | Description                          | Default |
|----------------|--------------------------------------|---------|
| `max_attempts` | Maximum number of retries            | `3`     |
| `delay`        | Wait time between attempts (seconds) | `2`     |
| `backoff`      | Optional exponential backoff         | `False` |

### 3. **Logging**
- Each attempt logs timestamp, command, result
- Final status logged with success/failure flag
- Optional: write to rotating log file (`.log`)

### 4. **Visual Feedback**
- Console output with color-coded status
- Optional GUI popup or tray notification
- Future: LED or screen indicator for physical feedback

## 🧪 Example Pseudocode

```python
for attempt in range(max_attempts):
    result = run_devcon_command()
    if result.success:
        break
    log_failure(attempt, result)
    time.sleep(delay)

import tkinter as tk
from tkinter import scrolledtext
from device_control import retry_devcon, reset_device
from config import MAX_ATTEMPTS, DELAY
import os

# === Device Options ===
DEVICE_OPTIONS = {
    "RX580": "PCI\\VEN_1002&DEV_67DF&SUBSYS_34171462&REV_E7",
    "GTX 1060": "PCI\\VEN_10DE&DEV_1C82&SUBSYS_85AE1043&REV_A1"
}
selected_device = tk.StringVar(value="RX580")  # Default selection

def get_device_id():
    return DEVICE_OPTIONS[selected_device.get()]

LOG_FILE = "device_control.log"

# === GUI Setup ===
root = tk.Tk()
root.title("GPU Control Panel")

# === Status Label ===
status_label = tk.Label(root, text="Ready", fg="black")
status_label.grid(row=6, column=0, columnspan=2, pady=(5, 0))

def update_status(message, success=True):
    status_label.config(text=message, fg="green" if success else "red")

# === Action Handlers ===
def handle_action(action):
    device_id = get_device_id()
    success = retry_devcon(action, device_id, MAX_ATTEMPTS, DELAY)
    update_status(f"{action.capitalize()} {'succeeded' if success else 'failed'}", success)
    refresh_log()

def handle_reset():
    device_id = get_device_id()
    success = reset_device(device_id)
    update_status(f"Reset {'succeeded' if success else 'failed'}", success)
    refresh_log()

# === Log Viewer ===
def refresh_log():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as f:
            log_content = f.read()
        log_viewer.delete(1.0, tk.END)
        log_viewer.insert(tk.END, log_content)
    else:
        log_viewer.delete(1.0, tk.END)
        log_viewer.insert(tk.END, "Log file not found.")

# === GUI Layout ===
tk.Label(root, text="Select Device:").grid(row=0, column=0, sticky="e")
tk.OptionMenu(root, selected_device, *DEVICE_OPTIONS.keys()).grid(row=0, column=1, sticky="w")

tk.Button(root, text="Enable", width=15, command=lambda: handle_action("enable")).grid(row=1, column=0, pady=5)
tk.Button(root, text="Disable", width=15, command=lambda: handle_action("disable")).grid(row=1, column=1, pady=5)
tk.Button(root, text="Reset", width=32, command=handle_reset).grid(row=2, column=0, columnspan=2, pady=10)

tk.Label(root, text="Log Viewer:").grid(row=3, column=0, columnspan=2, sticky="w", padx=5)
log_viewer = scrolledtext.ScrolledText(root, width=60, height=15)
log_viewer.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
refresh_log()

root.mainloop()

import tkinter as tk
from tkinter import messagebox, scrolledtext
from device_control import retry_devcon, reset_device
import os

DEVICE_ID = "PCI\\VEN_10DE&DEV_1C82"  # Replace with your actual device ID
LOG_FILE = "device_control.log"

def handle_action(action):
    success = retry_devcon(action, DEVICE_ID)
    if success:
        messagebox.showinfo("Success", f"{action.capitalize()} succeeded.")
    else:
        messagebox.showerror("Failure", f"{action.capitalize()} failed after retries.")
    refresh_log()

def handle_reset():
    success = reset_device(DEVICE_ID)
    if success:
        messagebox.showinfo("Success", "Device reset succeeded.")
    else:
        messagebox.showerror("Failure", "Device reset failed.")
    refresh_log()

def refresh_log():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, "r") as f:
            log_content = f.read()
        log_viewer.delete(1.0, tk.END)
        log_viewer.insert(tk.END, log_content)
    else:
        log_viewer.delete(1.0, tk.END)
        log_viewer.insert(tk.END, "Log file not found.")

# GUI setup
root = tk.Tk()
root.title("Device Control Panel")

tk.Label(root, text="Device ID:").grid(row=0, column=0, sticky="e")
tk.Label(root, text=DEVICE_ID).grid(row=0, column=1, sticky="w")

tk.Button(root, text="Enable", width=15, command=lambda: handle_action("enable")).grid(row=1, column=0, pady=5)
tk.Button(root, text="Disable", width=15, command=lambda: handle_action("disable")).grid(row=1, column=1, pady=5)
tk.Button(root, text="Reset", width=32, command=handle_reset).grid(row=2, column=0, columnspan=2, pady=10)

# Log viewer
tk.Label(root, text="Log Viewer:").grid(row=3, column=0, columnspan=2, sticky="w", padx=5)
log_viewer = scrolledtext.ScrolledText(root, width=60, height=15)
log_viewer.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
refresh_log()

root.mainloop()

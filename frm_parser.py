import re
import logging
from pathlib import Path
from datetime import datetime

# Setup logging
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

# Step 1: Parse .frm file
def parse_frm_file(path: Path) -> dict:
    try:
        with path.open("r", encoding="utf-8") as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with path.open("r", encoding="latin-1") as f:
            lines = f.readlines()


    form = {"name": None, "caption": None, "size": {}, "controls": []}
    current_control = None

    for line in lines:
        line = line.strip()
        if line.startswith("Begin VB.Form"):
            form["name"] = line.split()[-1]
        elif line.startswith("Caption"):
            form["caption"] = extract_value(line)
        elif line.startswith("ClientHeight"):
            form["size"]["height"] = extract_value(line)
        elif line.startswith("ClientWidth"):
            form["size"]["width"] = extract_value(line)
        elif line.startswith("Begin VB."):
            control_type, control_name = line.split()[1:3]
            current_control = {
                "type": control_type.replace("VB.", ""),
                "name": control_name,
                "properties": {}
            }
        elif line.startswith("End") and current_control:
            form["controls"].append(current_control)
            current_control = None
        elif current_control and "=" in line:
            key, value = map(str.strip, line.split("=", 1))
            current_control["properties"][key] = extract_value(value)

    return form

def extract_value(line: str) -> str:
    match = re.search(r'=\s*"?(.*?)"?$', line)
    return match.group(1) if match else line

# Step 2: Map VBA controls to tkinter widgets
def map_control_to_widget(control: dict) -> tuple[str, str]:
    widget_map = {
        "CommandButton": "Button",
        "TextBox": "Entry",
        "Label": "Label",
        "CheckBox": "Checkbutton",
        "OptionButton": "Radiobutton",
        "ComboBox": "ttk.Combobox",
        "ListBox": "Listbox"
    }

    widget_type = widget_map.get(control["type"], "Label")
    props = control["properties"]
    caption = props.get("Caption", control["name"])
    x = props.get("Left", "0")
    y = props.get("Top", "0")
    width = props.get("Width", "100")
    height = props.get("Height", "30")

    handler_name = f"on_{control['name'].lower()}_click"

    widget_code = (
        f'    # {control["type"]}: {caption}\n'
        f'    {control["name"]} = tk.{widget_type}(root, text="{caption}", command={handler_name})\n'
        f'    {control["name"]}.place(x={x}, y={y}, width={width}, height={height})\n'
    )

    handler_stub = (
        f'def {handler_name}():\n'
        f'    logging.info("🔘 {control["name"]} clicked")\n'
        f'    # Event handler for {control["name"]}\n'
        f'    pass\n'
    )

    return widget_code, handler_stub

# Step 3: Generate tkinter GUI code
def generate_tkinter_gui(parsed: dict, source_path: Path) -> str:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    metadata = [
        f'# Auto-generated from {source_path.name}',
        f'# Converted on {timestamp}',
        f'# Source: {source_path}',
        f'# Author: Rolf-Günther via frm_parser.py',
        ''
    ]

    lines = metadata + [
        "import tkinter as tk",
        "from tkinter import ttk",
        "import logging",
        "",
        "logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')",
        "",
        "def main():",
        "    root = tk.Tk()",
        f'    root.title("{parsed.get("caption", parsed["name"])}")'
    ]

    if "width" in parsed["size"] and "height" in parsed["size"]:
        lines.append(f'    root.geometry("{parsed["size"]["width"]}x{parsed["size"]["height"]}")')

    lines.append("")

    handlers = []

    for control in parsed["controls"]:
        widget_code, handler_stub = map_control_to_widget(control)
        lines.append(widget_code)
        handlers.append(handler_stub)

    lines.append("")
    lines.append("    root.mainloop()")
    lines.append("")
    lines.extend(handlers)
    lines.append("")
    lines.append("if __name__ == '__main__':")
    lines.append("    main()")

    return "\n".join(lines)

# Step 4: Convert one file
def convert_frm_to_tkinter(frm_path: Path, output_path: Path):
    parsed = parse_frm_file(frm_path)
    gui_code = generate_tkinter_gui(parsed, frm_path)
    with output_path.open("w", encoding="utf-8") as f:
        f.write(gui_code)
    logging.info(f"Converted {frm_path.name} → {output_path.name}")

# Step 5: Batch convert all .frm files
def batch_convert_frm_files(source_dir: Path, output_dir: Path):
    output_dir.mkdir(exist_ok=True)
    for frm_file in sorted(source_dir.glob("*.frm")):
        output_file = output_dir / f"{frm_file.stem}_gui.py"
        convert_frm_to_tkinter(frm_file, output_file)

if __name__ == "__main__":
    batch_convert_frm_files(Path("VBA_Sources"), Path("scripts"))

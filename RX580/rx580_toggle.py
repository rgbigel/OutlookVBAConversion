import subprocess
import time
import re

def get_rx580_instance_id():
    try:
        result = subprocess.run(
            ["pnputil", "/enum-devices", "/connected"],
            capture_output=True, text=True, check=True
        )
        lines = result.stdout.splitlines()
        instance_id = None
        current_block = []

        for line in lines:
            if line.strip() == "":
                current_block = []
                continue

            current_block.append(line.strip())

            if "Device Description" in line and "Radeon RX 580 Series" in line:
                for entry in current_block:
                    if entry.startswith("Instance ID"):
                        instance_id = entry.split(":", 1)[1].strip()
                        return instance_id
        return None
    except subprocess.CalledProcessError as e:
        print("Error running pnputil:", e)
        return None

def toggle_device(instance_id):
    print(f"Disabling device: {instance_id}")
    subprocess.run(["pnputil", "/disable-device", instance_id], shell=True)
    time.sleep(5)
    print(f"Enabling device: {instance_id}")
    subprocess.run(["pnputil", "/enable-device", instance_id], shell=True)

def main():
    instance_id = get_rx580_instance_id()
    if not instance_id:
        print("RX580 not found.")
        return
    toggle_device(instance_id)
    print("Toggle complete.")

if __name__ == "__main__":
    main()


import subprocess
import time
import logging

# Setup logging
logging.basicConfig(
    filename="device_control.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

def run_devcon(command: str, device_id: str) -> tuple[int, str]:
    """
    Executes a DevCon command and returns (exit_code, output).
    """
    cmd = ["devcon", command, device_id]
    logging.info(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=True, text=True)
    logging.info(f"Exit Code: {result.returncode}")
    logging.info(f"Output: {result.stdout.strip()}")
    return result.returncode, result.stdout.strip()

def retry_devcon(command: str, device_id: str, max_attempts: int = 3, delay: int = 2) -> bool:
    """
    Retries a DevCon command up to max_attempts with delay between tries.
    Returns True if successful, False otherwise.
    """
    for attempt in range(1, max_attempts + 1):
        logging.info(f"Attempt {attempt} for {command}")
        code, output = run_devcon(command, device_id)
        if code == 0:
            logging.info(f"Success on attempt {attempt}")
            return True
        logging.warning(f"Failed attempt {attempt}: {output}")
        time.sleep(delay)
    logging.error(f"All attempts failed for {command} on {device_id}")
    return False

def reset_device(device_id: str) -> bool:
    """
    Performs a device reset: disable → wait → enable.
    """
    logging.info(f"Starting reset for {device_id}")
    if not retry_devcon("disable", device_id):
        return False
    time.sleep(2)
    return retry_devcon("enable", device_id)

import time
import subprocess
import psutil
import pyautogui
from pywinauto import Desktop


# -------------------------------
# Configuration
# -------------------------------

FORTICLIENT_EXE = r"C:\Program Files\Fortinet\FortiClient\FortiClient.exe"
FORTICLIENT_TITLE_REGEX = ".*FortiClient.*"

VPN_ADAPTER_KEYWORD = "Forti"   # Keyword used to identify Forti VPN adapter
CONNECT_Y_RATIO = 0.60          # Vertical ratio of Connect button inside window
STARTUP_WAIT_SECONDS = 8
VPN_WAIT_TIMEOUT_SECONDS = 60


# -------------------------------
# VPN status detection
# -------------------------------

def is_vpn_connected() -> bool:
    """
    Check whether Forti VPN is connected by inspecting network interfaces.
    A connected VPN usually exposes a virtual adapter with an IPv4 address.
    """
    for iface_name, addresses in psutil.net_if_addrs().items():
        if VPN_ADAPTER_KEYWORD.lower() in iface_name.lower():
            for addr in addresses:
                if addr.family.name == "AF_INET":
                    return True
    return False


# -------------------------------
# FortiClient process control
# -------------------------------

def is_forticlient_running() -> bool:
    """
    Check whether FortiClient process is already running.
    """
    for proc in psutil.process_iter(attrs=["name"]):
        if proc.info["name"] and "FortiClient" in proc.info["name"]:
            return True
    return False


def start_forticlient():
    """
    Start FortiClient if it is not running.
    """
    if is_forticlient_running():
        return

    subprocess.Popen(FORTICLIENT_EXE)
    time.sleep(STARTUP_WAIT_SECONDS)


# -------------------------------
# UI automation (coordinate-based)
# -------------------------------

def get_forticlient_window():
    """
    Locate FortiClient top-level window using Desktop enumeration.
    This is more reliable than binding to a process.
    """
    windows = Desktop(backend="uia").windows(title_re=FORTICLIENT_TITLE_REGEX)

    if not windows:
        raise RuntimeError(
            "FortiClient window not found. "
            "Make sure the script is run as Administrator and FortiClient UI is visible."
        )

    return windows[0]


def click_connect_by_coordinates():
    """
    Click the FortiClient 'Connect' button using screen coordinates.
    This bypasses UI Automation limitations and works with custom UI controls.
    """
    window = get_forticlient_window()

    window.restore()
    window.set_focus()
    time.sleep(0.5)

    rect = window.rectangle()

    click_x = rect.left + rect.width() // 2
    click_y = rect.top + int(rect.height() * CONNECT_Y_RATIO)

    pyautogui.moveTo(click_x, click_y, duration=0.2)
    pyautogui.click()


# -------------------------------
# Connection wait logic
# -------------------------------

def wait_for_vpn_connection(timeout_seconds: int) -> bool:
    """
    Wait until VPN becomes connected or timeout expires.
    """
    start_time = time.time()

    while time.time() - start_time < timeout_seconds:
        if is_vpn_connected():
            return True
        time.sleep(2)

    return False


# -------------------------------
# Main orchestration
# -------------------------------

def main():
    """
    Main workflow:
    1. Check VPN status
    2. Start FortiClient if needed
    3. Trigger Connect via coordinate click
    4. Let user complete browser-based SSO
    5. Wait for VPN connection
    """
    print("Checking VPN connection status...")

    if is_vpn_connected():
        print("VPN is already connected.")
        return

    print("VPN not connected. Starting FortiClient...")
    start_forticlient()

    print("Triggering Connect button (browser SSO should appear)...")
    click_connect_by_coordinates()

    print("Waiting for VPN connection (user completes SSO)...")
    if wait_for_vpn_connection(VPN_WAIT_TIMEOUT_SECONDS):
        print("VPN connected successfully.")
    else:
        print("VPN connection timed out. Please complete SSO manually.")


if __name__ == "__main__":
    main()

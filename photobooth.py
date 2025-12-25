import tkinter as tk
from tkinter import Canvas, Button
from PIL import Image, ImageTk, ImageWin
import os
import glob
import time
import threading
import subprocess
import win32print
import win32ui
import random
import shutil
from typing import Optional
import json
import ctypes
import sys
from abc import ABC, abstractmethod

# ============================================================================
# CONFIGURATION
# ============================================================================

# Directories - relative to script location
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PHOTO_DIR = os.path.join(SCRIPT_DIR, 'Photos')
TEMP_DIR = os.path.join(SCRIPT_DIR, 'Temp')
DRIVE_DIR = r"G:\My Drive\New_Year_Photo_Booth"  # Google Drive location
LOG_PATH = os.path.join(TEMP_DIR, "booth.log")
LOGO_CANDIDATES = [
    os.environ.get("LOGO_PATH", ""),
    os.path.join(SCRIPT_DIR, "logo.png"),
    os.path.join(SCRIPT_DIR, "logo.PNG"),
    os.path.join(SCRIPT_DIR, "logo.jpg"),
    os.path.join(SCRIPT_DIR, "logo.jpeg"),
    "logo.png",
    "logo.PNG",
]

# Promo logo for right side
PROMO_LOGO_CANDIDATES = [
    os.environ.get("PROMO_LOGO_PATH", ""),
    os.path.join(SCRIPT_DIR, "promo.png"),
    os.path.join(SCRIPT_DIR, "promo.PNG"),
    os.path.join(SCRIPT_DIR, "promo.jpg"),
    os.path.join(SCRIPT_DIR, "promo.jpeg"),
    os.path.join(SCRIPT_DIR, "sponsor.png"),
    os.path.join(SCRIPT_DIR, "sponsor.jpg"),
    "promo.png",
    "promo.PNG",
]

# Create all required directories
os.makedirs(PHOTO_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(os.path.join(TEMP_DIR, "captures"), exist_ok=True)
os.makedirs(os.path.join(TEMP_DIR, "cache"), exist_ok=True)
try:
    os.makedirs(DRIVE_DIR, exist_ok=True)
except Exception:
    DRIVE_DIR = os.path.join(SCRIPT_DIR, 'Saved')
    os.makedirs(DRIVE_DIR, exist_ok=True)

# WSL / gphoto2 settings
WSL_DISTRO = "Ubuntu-22.04"
GPHOTO_CMD = "gphoto2"

# USB camera settings - Nikon Z6_3
CAMERA_VID_PID = "04b0:0454"
CAMERA_DEVICE_NAME = "Z6_3"
USBIPD_EXE = r"C:\Program Files\usbipd-win\usbipd.exe"

# Theme Colors - Premium New Year 2026
THEME_BG = "#0d0d1a"
THEME_ACCENT = "#FFD700"
THEME_ACCENT_2 = "#FF6B9D"
THEME_TEXT = "#FFFFFF"
THEME_SUCCESS = "#00D9A5"
THEME_DANGER = "#FF4757"
THEME_PRINT = "#FF6B35"
THEME_GLOW = "#B388FF"

# UI Settings
FONT_FAMILY = "Segoe UI"
COUNTDOWN_SIZE = 200
TITLE_SIZE = 64
SUBTITLE_SIZE = 32
STATUS_SIZE = 26
BUTTON_SIZE = 20

# Animation Settings
SNOW_COUNT = 100
SNOW_SPEED = (1, 4)
ANIMATION_FPS = 33

# Print Settings - Customize for your paper size
# Set to None to use printer defaults, or specify in inches
PRINT_PAPER_WIDTH_INCHES = 6       # Paper width in inches (e.g., 4, 5, 6)
PRINT_PAPER_HEIGHT_INCHES = 4      # Paper height in inches (e.g., 4, 6)
PRINT_MARGIN_INCHES = 0.0          # Margin around the image in inches
PRINT_FIT_TO_PAGE = True           # True = fit image to page, False = use original size
PRINT_CENTER_ON_PAGE = True        # Center the image on the page

# Global print job counter
_print_job_counter = 0

# ============================================================================
# HELPERS
# ============================================================================

def windows_path_to_wsl(path: str) -> str:
    """Convert Windows path to WSL path."""
    drive, rest = os.path.splitdrive(path)
    drive_letter = drive.rstrip(':').lower()
    rest = rest.replace('\\', '/').lstrip('/')
    return f"/mnt/{drive_letter}/{rest}"


def write_log(msg: str):
    """Log to console and file."""
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"{ts} {msg}"
    try:
        print(line, flush=True)
    except Exception:
        pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def is_admin() -> bool:
    """Check if current process is running with administrator privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def ensure_elevated():
    """If not running elevated, relaunch with UAC and exit. Otherwise, continue."""
    if is_admin():
        write_log("✓ Running with administrator privileges")
        return
    
    write_log("Not elevated; requesting UAC for admin mode...")
    try:
        # Relaunch current script with admin privileges
        python_exe = sys.executable
        script_path = os.path.abspath(__file__)
        ctypes.windll.shell32.ShellExecuteW(None, "runas", python_exe, f'"{script_path}"', None, 1)
        write_log("Relaunched with elevation; exiting non-admin instance...")
        sys.exit(0)
    except Exception as e:
        write_log(f"ERROR: Failed to elevate: {e}")
        write_log("Please run the app from an elevated PowerShell or Command Prompt.")
        sys.exit(1)


def resolve_logo_path() -> Optional[str]:
    """Find first existing logo file."""
    for path in LOGO_CANDIDATES:
        if path and os.path.exists(path):
            return path
    return None


def resolve_promo_logo_path() -> Optional[str]:
    """Find first existing promotional logo file."""
    for path in PROMO_LOGO_CANDIDATES:
        if path and os.path.exists(path):
            return path
    return None


def cleanup_old_temp_files(max_age_days: int = 7):
    """Clean up temporary files older than max_age_days."""
    try:
        cutoff_time = time.time() - (max_age_days * 86400)
        temp_captures = os.path.join(TEMP_DIR, "captures")
        
        if os.path.exists(temp_captures):
            for filename in os.listdir(temp_captures):
                filepath = os.path.join(temp_captures, filename)
                try:
                    if os.path.isfile(filepath) and os.path.getmtime(filepath) < cutoff_time:
                        os.remove(filepath)
                        write_log(f"Deleted old temp file: {filename}")
                except Exception as e:
                    write_log(f"Could not delete {filename}: {e}")
    except Exception as e:
        write_log(f"Cleanup error: {e}")


def disable_usb_selective_suspend():
    """Disable USB selective suspend to prevent Windows from disconnecting USB devices.
    This is a common cause of camera disconnections on laptops.
    """
    try:
        # Disable USB selective suspend on current power scheme
        # GUID 2a737441-1930-4402-8d77-b2bebba308a3 is USB settings
        # Sub-GUID 48e6b7a6-50f5-4782-a5d4-53bb8f07e226 is USB selective suspend
        result = subprocess.run(
            ["powercfg", "/SETACVALUEINDEX", "SCHEME_CURRENT", 
             "2a737441-1930-4402-8d77-b2bebba308a3", 
             "48e6b7a6-50f5-4782-a5d4-53bb8f07e226", "0"],
            capture_output=True, timeout=10
        )
        if result.returncode == 0:
            # Apply changes
            subprocess.run(["powercfg", "/SETACTIVE", "SCHEME_CURRENT"], 
                          capture_output=True, timeout=10)
            write_log("✓ USB selective suspend disabled (prevents camera disconnects)")
            return True
    except Exception as e:
        write_log(f"Could not disable USB selective suspend: {e}")
    return False


def run_command(cmd, timeout=10, shell=False):
    """Run a command and return (returncode, stdout, stderr)."""
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, shell=shell)
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return -1, "", f"Timeout after {timeout}s"
    except Exception as e:
        return -2, "", str(e)


def _ps_escape(arg: str) -> str:
    """Escape a string for PowerShell single-quoted context."""
    return str(arg).replace("'", "''")


def run_powershell_elevated(ps_command: str) -> bool:
    """Prompt for UAC and run a PowerShell command elevated (no output captured)."""
    # This launches: powershell.exe -NoProfile -WindowStyle Hidden -Command <ps_command>
    # with elevation via Start-Process -Verb RunAs from a non-elevated shell.
    # We cannot capture output; caller should verify by polling state.
    bootstrap = (
        "Start-Process -Verb RunAs -FilePath 'powershell' "
        "-ArgumentList '-NoProfile','-WindowStyle','Hidden','-Command'," +
        f"'" + _ps_escape(ps_command) + "' "
        "-WindowStyle Hidden"
    )
    write_log("Requesting elevation (UAC) for PowerShell command...")
    rc, out, err = run_command(["powershell", "-NoProfile", "-Command", bootstrap], timeout=8)
    if rc != 0:
        write_log(f"Elevation bootstrap failed rc={rc}: {err.strip()}")
        return False
    return True


def ensure_wsl_running():
    """Start WSL distro if not running."""
    try:
        run_command(["wsl", "-d", WSL_DISTRO, "--", "bash", "-lc", "exit"], timeout=5)
        return True
    except Exception:
        return False


def load_wsl_usb_modules() -> bool:
    """Load required USB kernel modules in WSL for usbip to work."""
    write_log("Loading USB kernel modules in WSL...")
    # Load vhci-hcd module (required for USB/IP client)
    rc, stdout, stderr = run_command(
        ["wsl", "-d", WSL_DISTRO, "--", "bash", "-lc", "sudo modprobe vhci-hcd 2>/dev/null || true"],
        timeout=10
    )
    if rc == 0:
        write_log("USB modules loaded (or already loaded)")
        return True
    else:
        write_log(f"Warning: Could not load USB modules: {stderr.strip()}")
        return False


def check_wsl_usbip_ready() -> bool:
    """Check if WSL is ready to receive USB devices."""
    write_log("Checking if WSL USB/IP is ready...")
    rc, stdout, stderr = run_command(
        ["wsl", "-d", WSL_DISTRO, "--", "bash", "-lc", "ls /sys/devices/platform/vhci_hcd.0 2>/dev/null && echo READY"],
        timeout=8
    )
    if "READY" in stdout:
        write_log("WSL USB/IP subsystem is ready")
        return True
    else:
        write_log("WSL USB/IP subsystem not ready, attempting to load modules...")
        return load_wsl_usb_modules()


def get_camera_busid():
    """Query usbipd and return camera BusID, or None if not found."""
    if not os.path.exists(USBIPD_EXE):
        write_log("ERROR: usbipd-win not found at " + USBIPD_EXE)
        return None

    rc, stdout, stderr = run_command([USBIPD_EXE, "list"], timeout=8)
    if rc != 0:
        write_log(f"usbipd list failed: {stderr}")
        return None

    # Look for camera in "Connected:" section only
    in_connected_section = False
    for line in stdout.splitlines():
        if line.strip().lower().startswith("connected:"):
            in_connected_section = True
            continue
        if line.strip().lower().startswith("persisted:"):
            in_connected_section = False
            continue
        
        if not in_connected_section:
            continue
            
        line_lower = line.lower()
        if CAMERA_DEVICE_NAME.lower() in line_lower or CAMERA_VID_PID.lower() in line_lower:
            write_log(f"Found camera: {line.strip()}")
            # Extract BusID
            tokens = line.split()
            for tok in tokens:
                if '-' in tok:
                    parts = tok.split('-')
                    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                        return tok
    
    write_log("Camera not found in connected devices. Is the camera plugged in and turned on?")
    return None


def get_camera_state(busid: str) -> str:
    """Return camera state: 'not-shared', 'shared', 'attached', or 'unknown'."""
    if not os.path.exists(USBIPD_EXE):
        return "unknown"

    rc, stdout, stderr = run_command([USBIPD_EXE, "list"], timeout=8)
    if rc != 0:
        return "unknown"

    for line in stdout.splitlines():
        if busid in line:
            if "Attached" in line:
                return "attached"
            elif "Shared" in line:
                return "shared"
            elif "Not shared" in line:
                return "not-shared"
    return "unknown"


def bind_camera(busid: str) -> bool:
    """Bind camera (changes 'Not shared' → 'Shared')."""
    if not os.path.exists(USBIPD_EXE):
        return False

    write_log(f"usbipd bind --busid {busid}")
    rc, stdout, stderr = run_command([USBIPD_EXE, "bind", "--busid", busid], timeout=15)
    
    if rc == 0:
        write_log(f"bind successful")
        return True
    else:
        errl = (stderr or "").lower()
        write_log(f"bind failed: {stderr.strip()}")
        # Try elevated bind
        ps_cmd = f"Start-Process -FilePath '{_ps_escape(USBIPD_EXE)}' -ArgumentList 'bind --busid {_ps_escape(busid)}' -Wait"
        if run_powershell_elevated(ps_cmd):
            # Poll state for up to 10 seconds
            for _ in range(10):
                time.sleep(1)
                st = get_camera_state(busid)
                if st in ("shared", "attached"):
                    write_log("bind successful (elevated)")
                    return True
        return False


def attach_camera(busid: str) -> bool:
    """Attach camera to WSL (requires --wsl flag in usbipd 5.x)."""
    if not os.path.exists(USBIPD_EXE):
        return False

    cmd = [USBIPD_EXE, "attach", "--wsl", WSL_DISTRO, "--busid", busid]
    
    cmd_str = " ".join(cmd[1:])  # Skip exe path for logging
    write_log(f"usbipd {cmd_str}")
    
    # Use longer timeout (90s) since attach can be slow
    rc, stdout, stderr = run_command(cmd, timeout=90)
    
    if rc == 0:
        write_log(f"attach successful")
        return True
    else:
        stderr_lower = stderr.lower()
        if "already attached" in stderr_lower:
            write_log(f"camera already attached (ok)")
            return True
        elif "timeout" in stderr_lower:
            write_log(f"attach timed out")
            return False
        else:
            write_log(f"attach failed: {stderr.strip()}")
            return False


def attach_camera_auto(busid: str) -> bool:
    """Start auto-attach in background and poll for success."""
    if not os.path.exists(USBIPD_EXE):
        return False
    
    write_log(f"Starting auto-attach for busid {busid}...")
    
    # Start auto-attach in background
    cmd = [USBIPD_EXE, "attach", "--wsl", WSL_DISTRO, "--busid", busid, "--auto-attach"]
    write_log(f"usbipd attach --wsl {WSL_DISTRO} --busid {busid} --auto-attach")
    
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        creationflags=subprocess.CREATE_NO_WINDOW if hasattr(subprocess, 'CREATE_NO_WINDOW') else 0
    )
    
    # Poll for camera to become attached (up to 60 seconds)
    for i in range(30):
        time.sleep(2)
        state = get_camera_state(busid)
        write_log(f"auto-attach poll {i+1}/30: state={state}")
        if state == "attached":
            write_log("auto-attach: camera attached!")
            return True
    
    # Kill background process if still running
    if proc.poll() is None:
        write_log("auto-attach: terminating background process...")
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except:
            proc.kill()
    
    return False


def detach_camera(busid: str) -> bool:
    """Detach camera from any client (safe to call if not attached)."""
    if not os.path.exists(USBIPD_EXE):
        return False

    write_log(f"usbipd detach --busid {busid}")
    rc, stdout, stderr = run_command([USBIPD_EXE, "detach", "--busid", busid], timeout=12)
    if rc == 0:
        write_log("detach successful or not attached")
        return True
    else:
        write_log(f"detach returned rc={rc}: {stderr.strip()}")
        # Try elevated detach
        ps_cmd = f"Start-Process -FilePath '{_ps_escape(USBIPD_EXE)}' -ArgumentList 'detach --busid {_ps_escape(busid)}' -Wait"
        if run_powershell_elevated(ps_cmd):
            # Poll until not attached
            for _ in range(10):
                time.sleep(1)
                st = get_camera_state(busid)
                if st in ("shared", "not-shared", "unknown"):
                    write_log("detach successful (elevated)")
                    return True
        return False


def restart_usbipd_service() -> None:
    """Best-effort restart of usbipd service. Requires elevation; logs outcome."""
    # Try PowerShell Restart-Service
    rc, out, err = run_command(["powershell", "-NoProfile", "-Command", "Restart-Service -Name usbipd -ErrorAction SilentlyContinue"], timeout=8)
    if rc == 0:
        write_log("usbipd service restart attempted")
        return
    write_log(f"Restart-Service usbipd failed (rc={rc}): {err.strip()}")
    # Try elevated PowerShell to restart the service
    if not run_powershell_elevated("Restart-Service -Name usbipd -Force"):
        # Fallback to sc stop/start (non-elevated may fail)
        run_command(["sc.exe", "stop", "usbipd"], timeout=6)
        time.sleep(1)
        run_command(["sc.exe", "start", "usbipd"], timeout=6)


def shutdown_wsl() -> None:
    """Shut down all WSL instances (best-effort)."""
    rc, out, err = run_command(["wsl", "--shutdown"], timeout=10)
    if rc != 0:
        write_log(f"wsl --shutdown rc={rc}: {err.strip()}")
    else:
        write_log("WSL shut down")


def usbipd_state_json() -> Optional[dict]:
    """Return parsed json from 'usbipd state', or None."""
    if not os.path.exists(USBIPD_EXE):
        return None
    rc, stdout, stderr = run_command([USBIPD_EXE, "state"], timeout=8)
    if rc != 0:
        return None
    try:
        return json.loads(stdout)
    except Exception:
        return None


def robust_attach_camera(busid: str, max_attempts: int = 2) -> bool:
    """Attach camera to WSL with retries and fallbacks.
    Strategy: if attached → ok; else try attach; on failure: detach → restart usbipd → wsl shutdown/start → bind (if needed) → attach.
    """
    # Quick success if already attached
    state = get_camera_state(busid)
    write_log(f"robust attach; initial state: {state}")
    if state == "attached":
        return True

    # Ensure shared state if needed
    if state == "not-shared":
        if not bind_camera(busid):
            return False
        time.sleep(1)

    # Ensure WSL USB/IP subsystem is ready before attempting attach
    check_wsl_usbip_ready()

    # Attempts
    for attempt in range(1, max_attempts + 1):
        write_log(f"attach attempt {attempt}/{max_attempts}")
        write_log(f"Calling attach_camera for busid {busid}...")
        attach_result = attach_camera(busid)
        write_log(f"attach_camera returned: {attach_result}")
        if attach_result:
            # Confirm
            write_log("Confirming camera state...")
            final_state = get_camera_state(busid)
            write_log(f"Camera state after attach: {final_state}")
            if final_state == "attached":
                write_log("✓ Camera successfully attached!")
                return True
            else:
                write_log(f"WARNING: attach_camera succeeded but state is '{final_state}', not 'attached'")
        # Fallback sequence
        write_log("attach failed; running fallback: detach → restart usbipd → wsl --shutdown/start → re-bind")
        detach_camera(busid)
        write_log("Waiting 2s after detach...")
        time.sleep(2)
        restart_usbipd_service()
        write_log("Waiting 3s after service restart...")
        time.sleep(3)
        shutdown_wsl()
        write_log("Waiting 5s after WSL shutdown...")
        time.sleep(5)
        write_log("Starting WSL...")
        ensure_wsl_running()
        write_log("Waiting 3s for WSL to fully start...")
        time.sleep(3)
        # Load USB modules again after WSL restart
        check_wsl_usbip_ready()
        # Re-evaluate share state
        s2 = get_camera_state(busid)
        write_log(f"Camera state after fallback: {s2}")
        if s2 == "not-shared":
            write_log("Binding camera...")
            bind_camera(busid)
            time.sleep(2)
        write_log(f"Ready for next attach attempt...")

    # Try auto-attach method
    write_log("Trying auto-attach method...")
    if attach_camera_auto(busid):
        final_state = get_camera_state(busid)
        if final_state == "attached":
            write_log("✓ Auto-attach succeeded!")
            return True

    # Final attempt: try non-blocking attach via PowerShell
    write_log("Trying non-blocking attach via PowerShell...")
    if attach_camera_nonblocking(busid):
        final_state = get_camera_state(busid)
        if final_state == "attached":
            write_log("✓ Non-blocking attach succeeded!")
            return True

    # Final verification
    final_state = get_camera_state(busid)
    write_log(f"robust attach; final state: {final_state}")
    return final_state == "attached"


def attach_camera_nonblocking(busid: str) -> bool:
    """Try to attach camera using PowerShell Start-Process (non-blocking with polling)."""
    if not os.path.exists(USBIPD_EXE):
        return False
    
    # Start the attach process in background via PowerShell
    ps_cmd = f"Start-Process -FilePath '{_ps_escape(USBIPD_EXE)}' -ArgumentList 'attach','--wsl','{WSL_DISTRO}','--busid','{busid}' -NoNewWindow -Wait"
    write_log(f"Starting attach via PowerShell: {ps_cmd[:80]}...")
    
    # Run PowerShell in background and poll for state change
    proc = subprocess.Popen(
        ["powershell", "-NoProfile", "-Command", ps_cmd],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE
    )
    
    # Poll for up to 45 seconds
    start_time = time.time()
    while time.time() - start_time < 45:
        # Check if process finished
        poll_result = proc.poll()
        if poll_result is not None:
            write_log(f"PowerShell attach process finished with code {poll_result}")
            break
        
        # Check camera state
        state = get_camera_state(busid)
        if state == "attached":
            write_log("Camera attached while polling!")
            proc.terminate()
            return True
        
        time.sleep(2)
    
    # Kill process if still running
    if proc.poll() is None:
        write_log("Terminating stuck attach process...")
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except:
            proc.kill()
    
    # Final check
    state = get_camera_state(busid)
    return state == "attached"


def gphoto2_detect() -> bool:
    """Run gphoto2 --auto-detect inside WSL and check for camera."""
    rc, stdout, stderr = run_command(
        ["wsl", "-d", WSL_DISTRO, "--", "bash", "-lc", "gphoto2 --auto-detect"],
        timeout=12
    )
    
    write_log(f"gphoto2 --auto-detect rc={rc}")
    write_log(f"gphoto2 output: {stdout.strip()}")
    
    if rc != 0:
        return False

    for ln in stdout.splitlines():
        if ln.strip() and not ln.lower().startswith('model') and not ln.strip().startswith('-'):
            write_log("gphoto2 detected camera!")
            return True
    return False


def fully_automated_camera_setup() -> bool:
    """Complete automation: Bind → Attach → Detect. Returns True if camera ready."""
    write_log("\n" + "="*60)
    write_log("STARTING FULLY AUTOMATED CAMERA SETUP")
    write_log("="*60)

    # Step 1: Ensure WSL is running
    write_log("\nStep 1: Ensuring WSL is running...")
    if not ensure_wsl_running():
        write_log("WARNING: WSL may not have started, continuing anyway")

    # Step 2: Get camera BusID
    write_log("\nStep 2: Finding camera BusID via usbipd...")
    busid = get_camera_busid()
    if not busid:
        write_log("ERROR: Camera not found in usbipd list")
        return False
    write_log(f"Found camera at BusID: {busid}")

    # Step 3: Check and bind if needed
    write_log("\nStep 3: Checking camera share state...")
    state = get_camera_state(busid)
    write_log(f"Current state: {state}")

    if state == "not-shared":
        write_log("Camera is 'Not shared', binding now...")
        if not bind_camera(busid):
            write_log("ERROR: Failed to bind camera")
            return False
        time.sleep(1)  # Brief wait after bind
        state = get_camera_state(busid)
        write_log(f"New state after bind: {state}")

    # Step 4: Attach if needed (robust with retries)
    write_log("\nStep 4: Attaching camera to WSL...")
    if state != "attached":
        if not robust_attach_camera(busid, max_attempts=2):
            write_log("ERROR: Failed to attach camera after retries")
            return False
        time.sleep(2)  # Brief wait after attach

    # Step 5: Verify with gphoto2
    write_log("\nStep 5: Verifying camera with gphoto2 --auto-detect...")
    if not gphoto2_detect():
        write_log("ERROR: gphoto2 failed to detect camera")
        return False

    write_log("\n" + "="*60)
    write_log("✓ CAMERA SETUP COMPLETE AND VERIFIED")
    write_log("="*60 + "\n")
    return True

# ============================================================================
# CAMERA INTERFACE
# ============================================================================

class CameraInterface(ABC):
    @abstractmethod
    def connect(self) -> bool: ...
    @abstractmethod
    def disconnect(self): ...
    @abstractmethod
    def is_connected(self) -> bool: ...
    @abstractmethod
    def capture(self, save_path: str) -> bool: ...
    @abstractmethod
    def get_name(self) -> str: ...


class WSLGPhotoCamera(CameraInterface):
    """Camera control via gphoto2 inside WSL"""

    def __init__(self, distro: Optional[str] = WSL_DISTRO):
        self.distro = distro
        self.camera_name = "USB PTP Class Camera"
        self.camera_port = None  # Store the USB port for explicit targeting
        self.connected = False

    def _wsl_prefix(self):
        return ["wsl", "-d", self.distro] if self.distro else ["wsl"]

    def _run_wsl(self, cmd: str, timeout: int = 20) -> subprocess.CompletedProcess:
        full_cmd = self._wsl_prefix() + ["bash", "-lc", cmd]
        write_log(f"[wsl] {cmd}")
        try:
            result = subprocess.run(full_cmd, capture_output=True, text=True, timeout=timeout)
            if result.returncode != 0:
                write_log(f"[wsl] rc={result.returncode} {result.stderr.strip()}")
            return result
        except subprocess.TimeoutExpired:
            write_log(f"[wsl] TIMEOUT after {timeout}s")
            return subprocess.CompletedProcess(full_cmd, 1, "", f"Timeout")

    def connect(self) -> bool:
        try:
            write_log("WSLGPhotoCamera: connecting...")
            result = self._run_wsl(f"{GPHOTO_CMD} --auto-detect", timeout=10)
            
            if result.returncode == 0:
                # Look for USB PTP camera first (preferred over Mass Storage)
                for ln in result.stdout.splitlines():
                    if ln.strip() and not ln.lower().startswith('model') and not ln.strip().startswith('-'):
                        # Parse camera name and port
                        parts = ln.strip().split()
                        if len(parts) >= 2:
                            port = parts[-1]  # Last column is the port
                            # Prefer USB PTP over disk/mass storage
                            if port.startswith('usb:'):
                                self.camera_port = port
                                self.camera_name = ' '.join(parts[:-1])
                                self.connected = True
                                write_log(f"✓ Camera connected: {self.camera_name} on {self.camera_port}")
                                return True
                
                # Fallback: accept any camera if no USB PTP found
                for ln in result.stdout.splitlines():
                    if ln.strip() and not ln.lower().startswith('model') and not ln.strip().startswith('-'):
                        self.camera_name = ln.strip()
                        self.connected = True
                        write_log(f"✓ Camera connected (fallback): {self.camera_name}")
                        return True
            
            write_log("✗ Camera detection failed")
            return False
        except Exception as ex:
            write_log(f"WSLGPhotoCamera exception: {ex}")
            return False

    def disconnect(self):
        self.connected = False
        self.camera_port = None

    def is_connected(self) -> bool:
        return self.connected

    def capture(self, save_path: str) -> bool:
        try:
            if not self.connected:
                write_log("capture: camera not connected")
                return False

            capture_start = time.time()
            write_log(f"capture: starting capture to {save_path}")
            
            # Skip autofocus - let camera use its current focus setting (MF or AF)
            # This avoids 3-4 second timeout when camera is in manual focus mode

            wsl_path = windows_path_to_wsl(save_path)
            write_log(f"capture: WSL path = {wsl_path}")
            
            # Build gphoto2 command with explicit port if available (fixes multi-interface cameras)
            port_arg = f"--port '{self.camera_port}'" if self.camera_port else ""
            gphoto_cmd = f"{GPHOTO_CMD} {port_arg} --capture-image-and-download --filename '{wsl_path}' 2>&1"
            write_log(f"capture: running {gphoto_cmd}")
            
            # Start gphoto2 in background and poll for file
            full_cmd = self._wsl_prefix() + ["bash", "-lc", gphoto_cmd]
            proc = subprocess.Popen(full_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            # Poll for file existence (up to 30 seconds)
            start_time = time.time()
            last_size = 0
            stable_count = 0
            file_first_seen = None
            
            while time.time() - start_time < 30:
                if os.path.exists(save_path):
                    if file_first_seen is None:
                        file_first_seen = time.time()
                        write_log(f"capture: file appeared after {file_first_seen - capture_start:.1f}s")
                    
                    current_size = os.path.getsize(save_path)
                    write_log(f"capture: file detected, size = {current_size} bytes")
                    
                    # Wait for file to be completely written (size stable for 0.5s)
                    if current_size > 100000 and current_size == last_size:
                        stable_count += 1
                        if stable_count >= 2:  # Stable for ~0.5 seconds
                            total_time = time.time() - capture_start
                            write_log(f"capture: file complete, size = {current_size} bytes, total time = {total_time:.1f}s")
                            # Kill gphoto2 process if still running
                            if proc.poll() is None:
                                proc.terminate()
                                try:
                                    proc.wait(timeout=2)
                                except:
                                    proc.kill()
                            write_log("capture: SUCCESS (fast path)")
                            return True
                    else:
                        stable_count = 0
                    last_size = current_size
                
                time.sleep(0.25)
            
            # If we get here, file wasn't detected in time - wait for process
            write_log("capture: fast path failed, waiting for gphoto2 to finish...")
            try:
                stdout, stderr = proc.communicate(timeout=30)
                write_log(f"capture: gphoto2 finished, rc={proc.returncode}")
            except subprocess.TimeoutExpired:
                proc.kill()
                write_log("capture: gphoto2 killed after timeout")
            
            # Final check
            if os.path.exists(save_path):
                file_size = os.path.getsize(save_path)
                if file_size > 1000:
                    write_log(f"capture: SUCCESS (slow path), size = {file_size} bytes")
                    return True
            
            write_log("capture: FAILED - file not created")
            return False
            
        except Exception as e:
            write_log(f"capture: EXCEPTION - {e}")
            return False

    def get_name(self) -> str:
        return self.camera_name


class WIACamera(CameraInterface):
    """Windows Image Acquisition fallback"""

    def __init__(self):
        self.device = None
        self.device_name = "WIA Camera"

    def connect(self) -> bool:
        try:
            import win32com.client
            device_manager = win32com.client.Dispatch("WIA.DeviceManager")
            for device_info in device_manager.DeviceInfos:
                if device_info.Type == 2:
                    try:
                        self.device = device_info.Connect()
                        self.device_name = device_info.Properties("Name").Value
                        return True
                    except Exception:
                        continue
            return False
        except Exception:
            return False

    def disconnect(self):
        self.device = None

    def is_connected(self) -> bool:
        return self.device is not None

    def capture(self, save_path: str) -> bool:
        try:
            if not self.device:
                return False
            import win32com.client
            for cmd in self.device.Commands:
                if "capture" in cmd.Name.lower() or cmd.CommandID == "{AF933CAC-ACAD-11D2-A093-00C04F72DC3C}":
                    self.device.ExecuteCommand(cmd.CommandID)
                    break
            items = self.device.Items
            if items.Count > 0:
                item = items[items.Count]
                image = item.Transfer()
                image.SaveFile(save_path)
                return True
            return False
        except Exception:
            return False

    def get_name(self) -> str:
        return self.device_name


class CameraManager:
    """Manage camera detection and selection"""

    def __init__(self):
        self.camera: Optional[CameraInterface] = None

    def detect_cameras(self) -> list:
        methods = []
        try:
            cam = WSLGPhotoCamera()
            if cam.connect():
                methods.append(("wsl-gphoto2", cam.get_name()))
                cam.disconnect()
        except Exception:
            pass
        try:
            cam = WIACamera()
            if cam.connect():
                methods.append(("wia", cam.get_name()))
                cam.disconnect()
        except Exception:
            pass
        return methods

    def connect_best(self) -> bool:
        methods = self.detect_cameras()
        for key, name in methods:
            if key == "wsl-gphoto2":
                cam = WSLGPhotoCamera()
                if cam.connect():
                    self.camera = cam
                    return True
            elif key == "wia":
                cam = WIACamera()
                if cam.connect():
                    self.camera = cam
                    return True
        return False

    def disconnect(self):
        if self.camera:
            try:
                self.camera.disconnect()
            except Exception:
                pass
            self.camera = None

    def is_connected(self) -> bool:
        return self.camera is not None and self.camera.is_connected()

    def get_name(self) -> str:
        return self.camera.get_name() if self.camera else "No Camera"

    def capture(self, save_path: str) -> bool:
        return self.camera.capture(save_path) if self.camera else False


# ============================================================================
# SNOWFLAKE ANIMATION
# ============================================================================

class Snowflake:
    """Festive particle - snow or golden sparkle"""
    COLORS = ["white", "white", "white", "#FFD700", "#FFD700", "#FF6B9D", "#B388FF"]  # Mix of snow and sparkles
    
    def __init__(self, canvas: Canvas, width: int, height: int):
        self.canvas = canvas
        self.width = width
        self.height = height
        self.reset()

    def reset(self):
        self.x = random.randint(0, self.width)
        self.y = random.randint(-self.height, 0)
        self.size = random.randint(2, 8)
        self.speed = random.uniform(*SNOW_SPEED)
        self.drift = random.uniform(-0.8, 0.8)
        self.color = random.choice(self.COLORS)
        # Golden particles are slightly larger
        if self.color != "white":
            self.size = random.randint(3, 6)
        self.id = self.canvas.create_oval(
            self.x, self.y, self.x + self.size, self.y + self.size,
            fill=self.color, outline="", tags="snow"
        )

    def update(self):
        self.y += self.speed
        self.x += self.drift
        if self.y > self.height:
            self.y = -10
            self.x = random.randint(0, self.width)
        if self.x < 0 or self.x > self.width:
            self.x = random.randint(0, self.width)
        self.canvas.coords(self.id, self.x, self.y, self.x + self.size, self.y + self.size)

# ============================================================================
# THEMED DIALOG (Canvas overlay - stays in fullscreen)
# ============================================================================

class ThemedDialog:
    """Custom dialog overlaid on the main canvas - no fullscreen exit needed"""
    
    @staticmethod
    def ask_yes_no(parent, title: str, message: str, icon: str = "⚠️", canvas=None) -> bool:
        """Show a themed yes/no dialog as canvas overlay. Returns True if yes."""
        
        # Try to find the canvas if not provided
        if canvas is None:
            for child in parent.winfo_children():
                if isinstance(child, Canvas):
                    canvas = child
                    break
        
        if canvas is None:
            # Fallback to simple dialog if no canvas found
            from tkinter import messagebox
            return messagebox.askyesno(title, message)
        
        result = [None]  # Store result
        
        # Get canvas dimensions
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        
        # Create dark overlay background
        overlay = canvas.create_rectangle(
            0, 0, cw, ch,
            fill="#000000", stipple="gray50",
            tags="dialog_overlay"
        )
        
        # Dialog box dimensions (larger for better readability)
        dw, dh = 600, 340
        dx = (cw - dw) // 2
        dy = (ch - dh) // 2
        
        # Dialog background (dark with gold border effect)
        canvas.create_rectangle(
            dx-4, dy-4, dx+dw+4, dy+dh+4,
            fill=THEME_ACCENT, outline="",
            tags="dialog_overlay"
        )
        canvas.create_rectangle(
            dx, dy, dx+dw, dy+dh,
            fill=THEME_BG, outline="",
            tags="dialog_overlay"
        )
        
        # Icon
        canvas.create_text(
            dx + dw//2, dy + 65,
            text=icon, font=("Segoe UI Emoji", 52),
            fill=THEME_ACCENT,
            tags="dialog_overlay"
        )
        
        # Title
        canvas.create_text(
            dx + dw//2, dy + 135,
            text=title, font=("Segoe UI", 28, "bold"),
            fill=THEME_TEXT,
            tags="dialog_overlay"
        )
        
        # Message
        canvas.create_text(
            dx + dw//2, dy + 190,
            text=message, font=("Segoe UI", 20),
            fill="#CCCCCC", width=550,
            tags="dialog_overlay"
        )
        
        # YES button
        yes_x, yes_y = dx + dw//2 - 100, dy + 275
        yes_btn = canvas.create_rectangle(
            yes_x-65, yes_y-28, yes_x+65, yes_y+28,
            fill=THEME_SUCCESS, outline="white", width=2,
            tags="dialog_overlay"
        )
        yes_txt = canvas.create_text(
            yes_x, yes_y, text="✓  YES",
            font=("Segoe UI", 20, "bold"), fill="white",
            tags="dialog_overlay"
        )
        
        # NO button
        no_x, no_y = dx + dw//2 + 100, dy + 275
        no_btn = canvas.create_rectangle(
            no_x-65, no_y-28, no_x+65, no_y+28,
            fill=THEME_DANGER, outline="white", width=2,
            tags="dialog_overlay"
        )
        no_txt = canvas.create_text(
            no_x, no_y, text="✗  NO",
            font=("Segoe UI", 20, "bold"), fill="white",
            tags="dialog_overlay"
        )
        
        def on_yes(event=None):
            result[0] = True
            canvas.delete("dialog_overlay")
        
        def on_no(event=None):
            result[0] = False
            canvas.delete("dialog_overlay")
        
        # Bind click events
        canvas.tag_bind(yes_btn, "<Button-1>", on_yes)
        canvas.tag_bind(yes_txt, "<Button-1>", on_yes)
        canvas.tag_bind(no_btn, "<Button-1>", on_no)
        canvas.tag_bind(no_txt, "<Button-1>", on_no)
        
        # Raise dialog to top
        canvas.tag_raise("dialog_overlay")
        
        # Wait for user response
        while result[0] is None:
            parent.update()
            parent.after(50)
        
        return result[0]
    
    @staticmethod  
    def show_message(parent, title: str, message: str, icon: str = "✓", canvas=None):
        """Show a themed message dialog as canvas overlay."""
        
        if canvas is None:
            for child in parent.winfo_children():
                if isinstance(child, Canvas):
                    canvas = child
                    break
        
        if canvas is None:
            from tkinter import messagebox
            messagebox.showinfo(title, message)
            return
        
        cw = canvas.winfo_width()
        ch = canvas.winfo_height()
        
        # Overlay
        canvas.create_rectangle(
            0, 0, cw, ch,
            fill="#000000", stipple="gray50",
            tags="dialog_overlay"
        )
        
        # Dialog box
        dw, dh = 450, 220
        dx = (cw - dw) // 2
        dy = (ch - dh) // 2
        
        canvas.create_rectangle(
            dx-4, dy-4, dx+dw+4, dy+dh+4,
            fill=THEME_ACCENT, outline="",
            tags="dialog_overlay"
        )
        canvas.create_rectangle(
            dx, dy, dx+dw, dy+dh,
            fill=THEME_BG, outline="",
            tags="dialog_overlay"
        )
        
        # Icon
        canvas.create_text(
            dx + dw//2, dy + 50,
            text=icon, font=("Segoe UI Emoji", 40),
            fill=THEME_SUCCESS,
            tags="dialog_overlay"
        )
        
        # Title
        canvas.create_text(
            dx + dw//2, dy + 100,
            text=title, font=("Segoe UI", 20, "bold"),
            fill=THEME_TEXT,
            tags="dialog_overlay"
        )
        
        # Message
        canvas.create_text(
            dx + dw//2, dy + 140,
            text=message, font=("Segoe UI", 14),
            fill="#CCCCCC", width=400,
            tags="dialog_overlay"
        )
        
        # OK button
        ok_x, ok_y = dx + dw//2, dy + 185
        ok_btn = canvas.create_rectangle(
            ok_x-50, ok_y-18, ok_x+50, ok_y+18,
            fill=THEME_ACCENT, outline="white", width=2,
            tags="dialog_overlay"
        )
        ok_txt = canvas.create_text(
            ok_x, ok_y, text="OK",
            font=("Segoe UI", 14, "bold"), fill="black",
            tags="dialog_overlay"
        )
        
        done = [False]
        
        def on_ok(event=None):
            done[0] = True
            canvas.delete("dialog_overlay")
        
        canvas.tag_bind(ok_btn, "<Button-1>", on_ok)
        canvas.tag_bind(ok_txt, "<Button-1>", on_ok)
        canvas.tag_raise("dialog_overlay")
        
        while not done[0]:
            parent.update()
            parent.after(50)

# ============================================================================
# PHOTO ZOOM & GALLERY VIEWER (Lightweight)
# ============================================================================

class PhotoZoomViewer:
    """Simple, lightweight photo viewer with basic zoom/pan"""

    def __init__(self, canvas: Canvas, ui_photo_id, width: int, height: int):
        self.canvas = canvas
        self.ui_photo_id = ui_photo_id
        self.width = width
        self.height = height

        # Photo management
        self.photos = []
        self.current_idx = 0
        self.original_image = None
        self.photo_display = None
        self.cached_image = None  # Pre-scaled for display

        # Zoom/Pan state
        self.zoom_level = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0
        self.min_zoom = 0.5
        self.max_zoom = 6.0  # Allow more zoom for detail
        
        # Base display dimensions
        self.base_display_w = 0
        self.base_display_h = 0
        
        # Drag state
        self.drag_start = None

        # UI elements
        self.zoom_text_id = None
        self.photo_counter_id = None

        # Bind mouse events
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind("<Button-1>", self._on_mouse_press)
        self.canvas.bind("<B1-Motion>", self._on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_mouse_release)
        self.canvas.bind("<Double-Button-1>", self._on_double_click)

    def load_existing_photos(self, photo_dir: str):
        """Load existing photos from directory"""
        try:
            if os.path.exists(photo_dir):
                photos = sorted([
                    os.path.join(photo_dir, f) 
                    for f in os.listdir(photo_dir) 
                    if f.lower().endswith(('.jpg', '.jpeg', '.png'))
                ], key=os.path.getmtime)
                self.photos = photos
                if photos:
                    write_log(f"Loaded {len(photos)} existing photos from gallery")
        except Exception as e:
            write_log(f"Error loading existing photos: {e}")

    def add_photo(self, photo_path: str):
        """Add photo to gallery"""
        if photo_path not in self.photos:
            self.photos.append(photo_path)
            self.current_idx = len(self.photos) - 1
        else:
            self.current_idx = self.photos.index(photo_path)

    def load_photo(self, photo_path: str):
        """Load and display a photo"""
        try:
            with Image.open(photo_path) as img:
                self.original_image = img.copy()
            
            self.zoom_level = 1.0
            self.pan_x = 0.0
            self.pan_y = 0.0
            
            self.add_photo(photo_path)
            self._cache_image()
            self._render_photo()
            self._update_photo_counter()
        except Exception as e:
            write_log(f"Error loading photo: {e}")

    def _cache_image(self):
        """Cache scaled version for display"""
        if not self.original_image:
            return
        try:
            display_width = int(self.width * 0.80)
            display_height = int(self.height * 0.65)
            ratio = min(display_width / self.original_image.width, display_height / self.original_image.height)
            
            self.base_display_w = int(self.original_image.width * ratio)
            self.base_display_h = int(self.original_image.height * ratio)
            
            # Cache at 2x for zoom quality
            cache_w = self.base_display_w * 2
            cache_h = self.base_display_h * 2
            self.cached_image = self.original_image.resize((cache_w, cache_h), Image.Resampling.LANCZOS)
        except Exception as e:
            write_log(f"Error caching image: {e}")

    def _render_photo(self):
        """Render photo at current zoom/pan"""
        if not self.cached_image:
            return
        try:
            display_w = max(1, int(self.base_display_w * self.zoom_level))
            display_h = max(1, int(self.base_display_h * self.zoom_level))
            
            img_resized = self.cached_image.resize((display_w, display_h), Image.Resampling.BILINEAR)
            self.photo_display = ImageTk.PhotoImage(img_resized)
            self.canvas.itemconfig(self.ui_photo_id, image=self.photo_display)
            
            center_x = self.width // 2 + int(self.pan_x)
            center_y = self.height // 2 + int(self.pan_y)
            self.canvas.coords(self.ui_photo_id, center_x, center_y)
            
            self._update_zoom_indicator()
        except Exception as e:
            write_log(f"Error rendering photo: {e}")

    def _clamp_pan(self):
        """Keep image on screen but allow full panning when zoomed"""
        if self.base_display_w == 0:
            return
        display_w = self.base_display_w * self.zoom_level
        display_h = self.base_display_h * self.zoom_level
        
        # Allow panning up to image edge, with some margin for buttons
        # When zoomed in (image larger than viewport), allow panning to see edges
        visible_w = self.width * 0.75  # Account for side buttons
        visible_h = self.height * 0.70  # Account for header/footer
        
        if display_w > visible_w:
            max_pan_x = (display_w - visible_w) / 2
        else:
            max_pan_x = 0
            
        if display_h > visible_h:
            max_pan_y = (display_h - visible_h) / 2
        else:
            max_pan_y = 0
            
        self.pan_x = max(-max_pan_x, min(max_pan_x, self.pan_x))
        self.pan_y = max(-max_pan_y, min(max_pan_y, self.pan_y))

    def _update_zoom_indicator(self):
        """Update zoom percentage display"""
        zoom_pct = int(self.zoom_level * 100)
        if self.zoom_text_id:
            self.canvas.itemconfig(self.zoom_text_id, text=f"🔍 {zoom_pct}%")
        else:
            self.zoom_text_id = self.canvas.create_text(
                self.width - 80, self.height - 40,
                text=f"🔍 {zoom_pct}%",
                font=("Segoe UI", 18, "bold"),
                fill="#FFD700",
                tags="zoom_ui"
            )

    def _update_photo_counter(self):
        """Show photo X of Y counter"""
        if not self.photos:
            return
        try:
            counter_text = f"📷 {self.current_idx + 1} / {len(self.photos)}"
            if self.photo_counter_id:
                self.canvas.itemconfig(self.photo_counter_id, text=counter_text)
            else:
                # Position below header bar (header is 80px tall)
                self.photo_counter_id = self.canvas.create_text(
                    self.width // 2, 100,
                    text=counter_text,
                    font=("Segoe UI", 20, "bold"),
                    fill="#FFFFFF",
                    tags="zoom_ui"
                )
        except Exception:
            pass

    def _switch_photo(self, photo_idx: int):
        """Switch to a different photo in gallery"""
        if 0 <= photo_idx < len(self.photos):
            self.current_idx = photo_idx
            with Image.open(self.photos[photo_idx]) as img:
                self.original_image = img.copy()
            self.zoom_level = 1.0
            self.pan_x = 0.0
            self.pan_y = 0.0
            self._cache_image()
            self._render_photo()
            self._update_photo_counter()
            
            # Notify callback if set (for updating button states)
            if hasattr(self, 'on_photo_changed') and self.on_photo_changed:
                self.on_photo_changed(self.photos[photo_idx])

    def _on_mousewheel(self, event):
        """Simple zoom on mouse wheel"""
        try:
            old_zoom = self.zoom_level
            if event.delta > 0:
                self.zoom_level *= 1.2
            else:
                self.zoom_level *= 0.83
            
            self.zoom_level = max(self.min_zoom, min(self.max_zoom, self.zoom_level))
            
            # Zoom toward mouse position
            if old_zoom != self.zoom_level:
                img_center_x = self.width // 2 + self.pan_x
                img_center_y = self.height // 2 + self.pan_y
                mouse_offset_x = event.x - img_center_x
                mouse_offset_y = event.y - img_center_y
                zoom_ratio = self.zoom_level / old_zoom
                self.pan_x = self.pan_x - mouse_offset_x * (zoom_ratio - 1)
                self.pan_y = self.pan_y - mouse_offset_y * (zoom_ratio - 1)
            
            self._clamp_pan()
            self._render_photo()
        except Exception as e:
            write_log(f"Zoom error: {e}")

    def _on_mouse_press(self, event):
        """Start drag"""
        self.drag_start = (event.x, event.y)

    def _on_mouse_drag(self, event):
        """Pan while dragging"""
        if self.drag_start:
            dx = event.x - self.drag_start[0]
            dy = event.y - self.drag_start[1]
            self.pan_x += dx
            self.pan_y += dy
            self._clamp_pan()
            self.drag_start = (event.x, event.y)
            self._render_photo()

    def _on_mouse_release(self, event):
        """End drag"""
        self.drag_start = None

    def _on_double_click(self, event):
        """Double-click to toggle zoom"""
        button_margin = 250
        if event.x < button_margin or event.x > self.width - button_margin:
            return
        
        if self.zoom_level > 1.2:
            self.zoom_level = 1.0
            self.pan_x = 0.0
            self.pan_y = 0.0
        else:
            old_zoom = self.zoom_level
            self.zoom_level = 2.0
            img_center_x = self.width // 2 + self.pan_x
            img_center_y = self.height // 2 + self.pan_y
            click_offset_x = event.x - img_center_x
            click_offset_y = event.y - img_center_y
            zoom_ratio = self.zoom_level / old_zoom
            self.pan_x = self.pan_x - click_offset_x * (zoom_ratio - 1)
            self.pan_y = self.pan_y - click_offset_y * (zoom_ratio - 1)
        
        self._clamp_pan()
        self._render_photo()

    def create_control_buttons(self):
        """Create modern control buttons with cool styling"""
        try:
            self.hide_control_buttons()
            
            # Gallery header bar at top
            header_y = 50
            self.canvas.create_rectangle(
                0, 0, self.width, 80,
                fill="#1a1a2e", outline="", tags="zoom_button"
            )
            self.canvas.create_text(
                self.width // 2, header_y,
                text="📸 PHOTO GALLERY",
                font=("Segoe UI", 28, "bold"),
                fill=THEME_ACCENT,
                tags="zoom_button"
            )
            
            # Control buttons - modern circular design
            button_y = int(self.height * 0.50)  # Centered vertically
            button_size = 70
            
            # LEFT SIDE - Zoom controls (vertical stack)
            left_x = 70
            zoom_spacing = 85
            
            # Zoom In button
            self._create_modern_btn(left_x, button_y - zoom_spacing, button_size, "➕", THEME_ACCENT, "#000", self._touch_zoom_in)
            
            # Zoom Out button
            self._create_modern_btn(left_x, button_y, button_size, "➖", THEME_ACCENT, "#000", self._touch_zoom_out)
            
            # Reset button
            self._create_modern_btn(left_x, button_y + zoom_spacing, button_size, "⟳", THEME_SUCCESS, "#FFF", self._touch_zoom_reset)

            # RIGHT SIDE - Navigation (vertical stack)
            right_x = self.width - 70
            
            # Previous Photo button
            self._create_modern_btn(right_x, button_y - zoom_spacing // 2, button_size, "◀", THEME_PRINT, "#FFF", self._touch_prev_photo)
            
            # Next Photo button
            self._create_modern_btn(right_x, button_y + zoom_spacing // 2, button_size, "▶", THEME_PRINT, "#FFF", self._touch_next_photo)
            
            # Photo counter and zoom indicator
            self._update_photo_counter()
            self._update_zoom_indicator()

        except Exception as e:
            write_log(f"Error creating control buttons: {e}")

    def _create_modern_btn(self, x, y, size, text, bg_color, text_color, callback):
        """Create modern circular button with shadow"""
        radius = size // 2
        
        # Shadow (offset circle)
        shadow = self.canvas.create_oval(
            x - radius + 4, y - radius + 4,
            x + radius + 4, y + radius + 4,
            fill="#0a0a14", outline="", tags="zoom_button"
        )
        
        # Main button circle
        btn = self.canvas.create_oval(
            x - radius, y - radius,
            x + radius, y + radius,
            fill=bg_color, outline="#FFFFFF", width=3, tags="zoom_button"
        )
        
        # Button text/icon
        txt = self.canvas.create_text(
            x, y, text=text,
            font=("Segoe UI Emoji", 24, "bold"),
            fill=text_color, tags="zoom_button"
        )
        
        self.canvas.tag_bind(btn, "<Button-1>", lambda e: callback())
        self.canvas.tag_bind(txt, "<Button-1>", lambda e: callback())

    def _create_control_btn(self, x, y, size, text, bg_color, text_color, callback):
        """Legacy helper - redirects to modern button"""
        self._create_modern_btn(x, y, size, text, bg_color, text_color, callback)

    def hide_control_buttons(self):
        """Hide control buttons and zoom UI"""
        try:
            self.canvas.delete("zoom_button")
            self.canvas.delete("zoom_ui")
            self.control_button_ids = {}
            self.zoom_text_id = None
            self.photo_counter_id = None
        except Exception:
            pass

    def _touch_zoom_in(self):
        """Zoom in button"""
        self.zoom_level = min(self.zoom_level * 1.3, self.max_zoom)
        self._clamp_pan()
        self._render_photo()

    def _touch_zoom_out(self):
        """Zoom out button"""
        self.zoom_level = max(self.zoom_level * 0.77, self.min_zoom)
        self._clamp_pan()
        self._render_photo()

    def _touch_zoom_reset(self):
        """Reset zoom and pan"""
        self.zoom_level = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0
        self._render_photo()

    def _touch_prev_photo(self):
        """Touch button: previous photo"""
        if self.current_idx > 0:
            self._switch_photo(self.current_idx - 1)

    def _touch_next_photo(self):
        """Touch button: next photo"""
        if self.current_idx < len(self.photos) - 1:
            self._switch_photo(self.current_idx + 1)


# ============================================================================
# MAIN APP
# ============================================================================

class PhotoBooth:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("New Year Sparkle Booth 2026")
        self.root.attributes("-fullscreen", True)
        self.root.configure(bg=THEME_BG)

        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()

        self.mode = "INIT"
        self.running = True
        self.camera_ready = False
        self.current_photo_path: Optional[str] = None
        self.capture_in_progress = False
        self.monitoring_camera = False
        self.last_retry_time = 0
        self.zoom_viewer = None  # Will be initialized after UI setup
        self.snow_paused = False  # Pause snow during photo view for performance
        
        # Threading lock to prevent race conditions on camera reconnect
        self.camera_reconnect_lock = threading.Lock()
        self.pause_keepalive = False  # Pause keepalive during active capture

        self.camera_manager = CameraManager()

        self.setup_canvas()
        self.setup_snowflakes()
        self.setup_ui_elements()
        self.setup_buttons()

        self.root.bind("<Escape>", lambda e: self.quit_app())
        self.root.bind("<F5>", lambda e: self.retry_camera())
        self.root.bind("<plus>", lambda e: self._zoom_in())
        self.root.bind("<equal>", lambda e: self._zoom_in())
        self.root.bind("<minus>", lambda e: self._zoom_out())
        self.root.bind("<0>", lambda e: self._zoom_reset())
        self.root.bind("<Left>", lambda e: self._prev_photo())
        self.root.bind("<Right>", lambda e: self._next_photo())
        self.root.protocol("WM_DELETE_WINDOW", self.quit_app)
        self.canvas.bind("<Button-1>", self.on_canvas_click)

        self.init_camera()
        self.start_camera_monitoring()
        self.animate()

    def setup_canvas(self):
        self.canvas = Canvas(self.root, bg=THEME_BG, highlightthickness=0)
        self.canvas.place(x=0, y=0, width=self.width, height=self.height)

    def setup_snowflakes(self):
        self.snowflakes = [Snowflake(self.canvas, self.width, self.height) for _ in range(SNOW_COUNT)]

    def setup_ui_elements(self):
        center_x = self.width / 2
        
        # Animation state for pulsing effects
        self.animation_frame = 0
        self.twinkling_stars = []
        
        # Create twinkling stars around the screen
        self._create_twinkling_stars()
        
        # Title with clean shadow (not messy glow)
        title_y = self.height * 0.08
        
        # Shadow layer (single, offset)
        self.canvas.create_text(
            center_x + 3, title_y + 3,
            text="HAPPY NEW YEAR",
            font=(FONT_FAMILY, TITLE_SIZE - 10, "bold"),
            fill="#1a1a2e",
            tags="ui_shadow"
        )
        
        # Main title - clean and readable
        self.ui_title = self.canvas.create_text(
            center_x, title_y,
            text="HAPPY NEW YEAR",
            font=(FONT_FAMILY, TITLE_SIZE - 10, "bold"),
            fill=THEME_TEXT,
            tags="ui"
        )
        
        # Large year "2026" with pulsing effect
        year_y = self.height * 0.20
        
        # Year shadow
        self.canvas.create_text(
            center_x + 4, year_y + 4,
            text="2026",
            font=(FONT_FAMILY, 140, "bold"),
            fill="#1a1a2e",
            tags="ui_shadow"
        )
        
        # Year main text (will be animated)
        self.ui_year = self.canvas.create_text(
            center_x, year_y,
            text="2026",
            font=(FONT_FAMILY, 140, "bold"),
            fill=THEME_ACCENT,
            tags="ui_year"
        )
        
        # Decorative sparkle line
        line_y = self.height * 0.32
        line_length = 250
        
        # Left line with gradient effect (multiple lines)
        for i, alpha in enumerate([0.3, 0.5, 0.7, 1.0]):
            offset = i * 60
            color = THEME_ACCENT if alpha == 1.0 else "#997a00"
            self.canvas.create_line(
                center_x - line_length - offset, line_y,
                center_x - 80 - offset, line_y,
                fill=color, width=2, tags="ui_deco"
            )
        
        # Center diamond
        diamond_size = 12
        self.canvas.create_polygon(
            center_x, line_y - diamond_size,
            center_x + diamond_size, line_y,
            center_x, line_y + diamond_size,
            center_x - diamond_size, line_y,
            fill=THEME_ACCENT, outline="", tags="ui_deco"
        )
        
        # Right line
        for i, alpha in enumerate([0.3, 0.5, 0.7, 1.0]):
            offset = i * 60
            color = THEME_ACCENT if alpha == 1.0 else "#997a00"
            self.canvas.create_line(
                center_x + 80 + offset, line_y,
                center_x + line_length + offset, line_y,
                fill=color, width=2, tags="ui_deco"
            )
        
        # Instruction text
        self.ui_instruction = self.canvas.create_text(
            center_x, self.height * 0.38,
            text="Initializing...",
            font=(FONT_FAMILY, SUBTITLE_SIZE + 2, "bold"),
            fill=THEME_TEXT,
            tags="ui"
        )
        
        # Status text
        self.ui_status = self.canvas.create_text(
            center_x, self.height * 0.44,
            text="",
            font=(FONT_FAMILY, STATUS_SIZE),
            fill=THEME_SUCCESS,
            tags="ui"
        )
        
        # Countdown display
        self.ui_countdown = self.canvas.create_text(
            center_x, self.height * 0.55,
            text="",
            font=(FONT_FAMILY, COUNTDOWN_SIZE + 60, "bold"),
            fill=THEME_ACCENT,
            tags="ui"
        )
        
        # Countdown message
        self.ui_countdown_msg = self.canvas.create_text(
            center_x, self.height * 0.55,
            text="",
            font=(FONT_FAMILY, 72, "bold"),
            fill=THEME_ACCENT,
            state="hidden",
            tags="ui"
        )
        
        # Photo display
        self.ui_photo = self.canvas.create_image(
            center_x, self.height * 0.52,
            image=None,
            state="hidden",
            tags="photo"
        )
        
        # Setup logos and viewer
        self._setup_logos_and_viewer()
    
    def _create_twinkling_stars(self):
        """Create decorative twinkling stars around the edges"""
        star_positions = [
            (0.05, 0.10), (0.95, 0.10),
            (0.08, 0.25), (0.92, 0.25),
            (0.03, 0.40), (0.97, 0.40),
            (0.06, 0.55), (0.94, 0.55),
            (0.04, 0.70), (0.96, 0.70),
            (0.07, 0.85), (0.93, 0.85),
            (0.15, 0.05), (0.85, 0.05),
            (0.25, 0.08), (0.75, 0.08),
        ]
        
        for rel_x, rel_y in star_positions:
            x = self.width * rel_x
            y = self.height * rel_y
            size = random.randint(6, 12)
            color = random.choice([THEME_ACCENT, "#FFFFFF", THEME_ACCENT_2, THEME_GLOW])
            
            star_id = self.canvas.create_text(
                x, y, text="✦",
                font=("Arial", size, "bold"),
                fill=color,
                tags="twinkle_star"
            )
            self.twinkling_stars.append({
                'id': star_id,
                'base_size': size,
                'phase': random.uniform(0, 6.28),
                'speed': random.uniform(0.05, 0.15)
            })

    def _setup_logos_and_viewer(self):
        """Setup logos and photo viewer - called from setup_ui_elements"""
        # Initialize zoom viewer and load existing photos
        self.zoom_viewer = PhotoZoomViewer(self.canvas, self.ui_photo, self.width, self.height)
        self.zoom_viewer.load_existing_photos(PHOTO_DIR)
        
        # Set callback to update button states when photo changes
        self.zoom_viewer.on_photo_changed = self._on_gallery_photo_changed

        # Logo
        self.logo_image = None
        logo_path = resolve_logo_path()
        if logo_path:
            try:
                img = Image.open(logo_path)
                max_size = min(self.width * 0.35, self.height * 0.28)
                ratio = min(max_size / img.width, max_size / img.height)
                new_w = int(img.width * ratio)
                new_h = int(img.height * ratio)
                img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                self.logo_image = ImageTk.PhotoImage(img)
                logo_x = self.width * 0.12
                logo_y = self.height * 0.16
                self.ui_logo = self.canvas.create_image(logo_x, logo_y, image=self.logo_image, tags="logo")
                self.canvas.tag_lower("logo", "snow")
            except Exception as e:
                write_log(f"logo failed: {e}")
        
        # Promotional Logo (right side)
        self.promo_logo_image = None
        promo_path = resolve_promo_logo_path()
        if promo_path:
            try:
                img = Image.open(promo_path)
                max_size = min(self.width * 0.35, self.height * 0.28)  # Same as main logo
                ratio = min(max_size / img.width, max_size / img.height)
                new_w = int(img.width * ratio)
                new_h = int(img.height * ratio)
                img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                self.promo_logo_image = ImageTk.PhotoImage(img)
                promo_x = self.width * 0.88  # Mirror of 0.12
                promo_y = self.height * 0.16  # Same as main logo
                self.ui_promo = self.canvas.create_image(promo_x, promo_y, image=self.promo_logo_image, tags="promo")
                self.canvas.tag_lower("promo", "snow")
                write_log(f"Loaded promotional logo: {promo_path}")
            except Exception as e:
                write_log(f"promo logo failed: {e}")

    def setup_buttons(self):
        btn_font = (FONT_FAMILY, BUTTON_SIZE + 2, "bold")
        
        # Common button style settings
        btn_style = {
            "relief": "flat",
            "cursor": "hand2",
            "borderwidth": 0,
            "highlightthickness": 0,
        }
        
        # Action buttons for RESULT mode - Premium styling
        self.btn_delete = Button(self.root, text="❌  DELETE", font=btn_font, 
                                 bg=THEME_DANGER, fg="white",
                                 activebackground="#E8384F", activeforeground="white",
                                 padx=25, pady=12, **btn_style,
                                 command=self.action_delete)
        
        self.btn_save = Button(self.root, text="☁️  UPLOAD", font=btn_font, 
                               bg=THEME_SUCCESS, fg="white",
                               activebackground="#00F5B8", activeforeground="white",
                               padx=25, pady=12, **btn_style,
                               command=self.action_save)
        
        self.btn_print = Button(self.root, text="🖨️  PRINT", font=btn_font, 
                                bg=THEME_PRINT, fg="white",
                                activebackground="#FF8357", activeforeground="white",
                                padx=25, pady=12, **btn_style,
                                command=self.action_print)
        
        self.btn_new = Button(self.root, text="🏠  HOME", font=btn_font, 
                              bg=THEME_ACCENT, fg="#0d0d1a",
                              activebackground="#FFE55C", activeforeground="#0d0d1a",
                              padx=25, pady=12, **btn_style,
                              command=self.action_go_home)
        
        # VIEW PHOTOS button for READY mode (home screen)
        self.btn_view_photos = Button(self.root, text="📸  VIEW PHOTOS", font=btn_font, 
                                      bg=THEME_GLOW, fg="white",
                                      activebackground="#D4A5FF", activeforeground="white",
                                      padx=25, pady=12, **btn_style,
                                      command=self.action_view_photos)
        
        # START button for taking photos (home screen) - Extra large and prominent
        start_font = (FONT_FAMILY, BUTTON_SIZE + 14, "bold")
        self.btn_start = Button(self.root, text="📷  TAP TO START", font=start_font, 
                                bg=THEME_SUCCESS, fg="white",
                                activebackground="#00F5B8", activeforeground="white",
                                padx=50, pady=25, **btn_style,
                                command=self.start_photo_workflow)
        self.hide_buttons()

    def init_camera(self):
        # Clean up old temp files on startup
        cleanup_old_temp_files(max_age_days=7)
        self.update_ui_state("INIT")
        threading.Thread(target=self._init_camera_thread, daemon=True).start()

    def _init_camera_thread(self):
        try:
            # Run the complete automated setup
            if fully_automated_camera_setup():
                # Now connect via camera manager
                if self.camera_manager.connect_best():
                    name = self.camera_manager.get_name()
                    self.camera_ready = True
                    self.root.after(0, lambda: self.update_ui_state("READY", name))
                else:
                    self.camera_ready = False
                    self.root.after(0, lambda: self.update_ui_state("ERROR", "Camera not detected by manager"))
            else:
                self.camera_ready = False
                self.root.after(0, lambda: self.update_ui_state("ERROR", "Camera setup failed"))
        except Exception as ex:
            write_log(f"init exception: {ex}")
            self.camera_ready = False
            self.root.after(0, lambda: self.update_ui_state("ERROR", "Connection error"))

    def start_camera_monitoring(self):
        """Background health check and keepalive."""
        self.monitoring_camera = True
        self.last_keepalive = time.time()
        threading.Thread(target=self._monitor_camera_health, daemon=True).start()

    def _camera_keepalive(self):
        """Send a lightweight command to keep the camera connection alive."""
        try:
            # Use --auto-detect which is proven to work for Nikon Z6_3
            # The --get-config command was timing out for this camera
            result = subprocess.run(
                ["wsl", "-d", WSL_DISTRO, "gphoto2", "--auto-detect"],
                capture_output=True,
                timeout=10  # Increased timeout for more reliability
            )
            # Check if camera is detected in output
            if result.returncode == 0 and b"usb:" in result.stdout:
                return True
            return False
        except Exception as e:
            write_log(f"Keepalive failed: {e}")
            return False

    def _monitor_camera_health(self):
        while self.running and getattr(self, 'monitoring_camera', False):
            try:
                # Skip keepalive if paused (during active capture workflow)
                if getattr(self, 'pause_keepalive', False):
                    time.sleep(2)
                    continue
                
                # Keepalive in READY or RESULT mode (not during active capture)
                if self.mode in ("READY", "RESULT") and self.camera_ready and not getattr(self, 'capture_in_progress', False):
                    # Keepalive every 60 seconds (was 15, but that was too aggressive)
                    current_time = time.time()
                    if current_time - self.last_keepalive > 60:
                        if self._camera_keepalive():
                            self.last_keepalive = current_time
                            write_log("Keepalive: camera still connected")
                        else:
                            # Try to acquire lock before reconnecting (non-blocking)
                            if self.camera_reconnect_lock.acquire(blocking=False):
                                try:
                                    write_log("Keepalive failed, auto-reconnecting...")
                                    # Immediate reconnect attempt
                                    if fully_automated_camera_setup():
                                        self.camera_manager.connect_best()
                                        write_log("Camera auto-reconnected successfully")
                                    else:
                                        write_log("Auto-reconnect failed, will retry on next check")
                                finally:
                                    self.camera_reconnect_lock.release()
                            else:
                                write_log("Keepalive: reconnect skipped (capture in progress)")
                            self.last_keepalive = current_time
                # Poll every 10 seconds (was 5, reduced frequency)
                time.sleep(10)
            except Exception:
                time.sleep(10)

    def update_ui_state(self, state: str, message: str = ""):
        self.mode = state
        self.canvas.itemconfig(self.ui_title, state="normal")
        self.canvas.itemconfig(self.ui_instruction, state="hidden")
        self.canvas.itemconfig(self.ui_status, state="hidden")
        self.canvas.itemconfig(self.ui_countdown, state="hidden")
        self.canvas.itemconfig(self.ui_countdown_msg, state="hidden")
        self.canvas.itemconfig(self.ui_photo, state="hidden")
        # Show snow for non-RESULT states
        self.canvas.itemconfig("snow", state="normal")
        self.hide_buttons()

        if state == "INIT":
            self.canvas.itemconfig(self.ui_instruction, text="Setting up the countdown...", state="normal")
            self.canvas.itemconfig(self.ui_status, text="⏳ Binding & attaching camera...", state="normal")
        elif state == "READY":
            # Restore ALL decorative elements when returning to home
            self.canvas.itemconfig("ui_year", state="normal")
            self.canvas.itemconfig("ui_shadow", state="normal")
            self.canvas.itemconfig("ui_deco", state="normal")
            self.canvas.itemconfig("twinkle_star", state="normal")
            self.canvas.itemconfig("logo", state="normal")
            self.canvas.itemconfig("promo", state="normal")
            if hasattr(self, 'ui_year'):
                self.canvas.itemconfig(self.ui_year, state="normal")
            # Restore logos to original top position
            if hasattr(self, 'ui_logo'):
                self.canvas.coords(self.ui_logo, self.width * 0.12, self.height * 0.16)
            if hasattr(self, 'ui_promo'):
                self.canvas.coords(self.ui_promo, self.width * 0.88, self.height * 0.16)
            # Clean home screen - no status text, just buttons
            self.canvas.itemconfig(self.ui_instruction, text="", state="hidden")
            self.canvas.itemconfig(self.ui_status, text="", state="hidden")
            # Show START and VIEW PHOTOS buttons on home screen
            self.btn_start.place(relx=0.5, rely=0.5, anchor="center")
            self.btn_view_photos.place(relx=0.5, rely=0.92, anchor="center")
        elif state == "COUNTDOWN":
            # Hide year and title during countdown for cleaner look
            self.canvas.itemconfig("ui_year", state="hidden")
            self.canvas.itemconfig("ui_shadow", state="hidden")
            self.canvas.itemconfig("ui_deco", state="hidden")
            self.canvas.itemconfig("twinkle_star", state="hidden")
            self.canvas.itemconfig(self.ui_title, state="hidden")
            if hasattr(self, 'ui_year'):
                self.canvas.itemconfig(self.ui_year, state="hidden")
            self.canvas.itemconfig(self.ui_countdown, state="normal")
            # Move status text to instruction position (higher up)
            self.canvas.itemconfig(self.ui_instruction, text="✨ Get Ready! ✨", fill=THEME_ACCENT, state="normal")
        elif state == "RESULT":
            # Hide decorative elements but keep logos (repositioned)
            self.canvas.itemconfig(self.ui_title, state="hidden")
            self.canvas.itemconfig("ui_year", state="hidden")
            self.canvas.itemconfig("ui_shadow", state="hidden")
            self.canvas.itemconfig("ui_deco", state="hidden")
            self.canvas.itemconfig("twinkle_star", state="hidden")
            if hasattr(self, 'ui_year'):
                self.canvas.itemconfig(self.ui_year, state="hidden")
            
            # Move logos to just below header for gallery view
            if hasattr(self, 'ui_logo'):
                self.canvas.coords(self.ui_logo, self.width * 0.10, self.height * 0.22)
                self.canvas.itemconfig("logo", state="normal")
            if hasattr(self, 'ui_promo'):
                self.canvas.coords(self.ui_promo, self.width * 0.90, self.height * 0.22)
                self.canvas.itemconfig("promo", state="normal")
            
            self.canvas.itemconfig(self.ui_photo, state="normal")
            # Hide snow for cleaner photo view
            self.canvas.itemconfig("snow", state="hidden")
            self.show_buttons()
            # Show touch-friendly control buttons for zoom/navigation
            if hasattr(self, 'zoom_viewer') and self.zoom_viewer:
                self.zoom_viewer.create_control_buttons()
        elif state == "ERROR":
            # Show snow again
            self.canvas.itemconfig("snow", state="normal")
            self.canvas.itemconfig(self.ui_instruction, text="Tap to retry", fill=THEME_ACCENT, state="normal")
            self.canvas.itemconfig(self.ui_status, text=f"✗ {message}", fill=THEME_DANGER, state="normal")

    def show_buttons(self):
        btn_y = 0.92
        
        # Check if current photo is already uploaded
        is_uploaded = False
        if self.current_photo_path and os.path.exists(self.current_photo_path):
            filename = os.path.basename(self.current_photo_path)
            dest = os.path.join(DRIVE_DIR, filename)
            is_uploaded = os.path.exists(dest)
        
        # Update upload button based on status
        if is_uploaded:
            self.btn_save.config(text="☁️ ✓ DONE", bg="#2D6A4F", fg="white", state="disabled")
        else:
            self.btn_save.config(text="☁️  UPLOAD", bg=THEME_SUCCESS, fg="white", state="normal")
        
        # 4 buttons evenly spaced
        self.btn_delete.place(relx=0.12, rely=btn_y, anchor="center")
        self.btn_save.place(relx=0.37, rely=btn_y, anchor="center")
        self.btn_print.place(relx=0.62, rely=btn_y, anchor="center")
        self.btn_new.place(relx=0.87, rely=btn_y, anchor="center")

    def _on_gallery_photo_changed(self, new_photo_path: str):
        """Called when user navigates to a different photo in gallery"""
        self.current_photo_path = new_photo_path
        # Refresh button states for the new photo
        if self.mode == "RESULT":
            self.show_buttons()

    def hide_buttons(self):
        for btn in [self.btn_delete, self.btn_save, self.btn_print, self.btn_new, self.btn_view_photos, self.btn_start]:
            btn.place_forget()
        # Hide touch control buttons too
        if hasattr(self, 'zoom_viewer') and self.zoom_viewer:
            self.zoom_viewer.hide_control_buttons()

    def on_canvas_click(self, event):
        """Handle canvas clicks - only for error retry, NOT for photo capture"""
        # Photo capture is ONLY triggered by the TAP TO START button
        # Canvas click is only used for error retry
        if self.mode == "ERROR":
            self.retry_camera()

    def start_photo_workflow(self):
        if not self.camera_ready:
            return
        
        self.update_ui_state("COUNTDOWN")  # Show countdown UI
        self.root.update()
        
        # Check if we need to reconnect or can skip it
        # If keepalive succeeded recently (within 90 seconds), skip reconnect
        time_since_keepalive = time.time() - getattr(self, 'last_keepalive', 0)
        if time_since_keepalive < 90 and self.camera_manager.is_connected():
            write_log(f"Skipping reconnect (keepalive {time_since_keepalive:.0f}s ago)")
            # Start countdown directly
            threading.Thread(target=self._quick_capture, daemon=True).start()
        else:
            write_log("Force reconnecting camera before capture...")
            # Run reconnect in background then start countdown
            threading.Thread(target=self._reconnect_and_capture, daemon=True).start()
    
    def _quick_capture(self):
        """Start capture without reconnecting (camera already verified)"""
        try:
            self.pause_keepalive = True
            self.capture_in_progress = True
            self._countdown_and_capture()
        except Exception as e:
            write_log(f"_quick_capture error: {e}")
            self.root.after(0, self.capture_failed)
        finally:
            self.pause_keepalive = False
    
    def _reconnect_and_capture(self):
        """Reconnect camera and then start capture"""
        try:
            # Pause keepalive to prevent race conditions
            self.pause_keepalive = True
            self.capture_in_progress = True
            
            # Show "Connecting" message instead of countdown
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_instruction, text="📷 Connecting Camera...", fill=THEME_TEXT))
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_countdown, text="", state="hidden"))
            
            # Acquire lock to prevent any other reconnection attempts
            with self.camera_reconnect_lock:
                # Force reattach to ensure fresh USB connection
                if fully_automated_camera_setup():
                    self.camera_manager.connect_best()
                    write_log("Camera reconnected for capture")
                else:
                    write_log("Camera reconnect warning - proceeding anyway")
            
            # Now show "Get Ready" and start countdown
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_instruction, text="✨ Get Ready! ✨", fill=THEME_ACCENT))
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_countdown, state="normal"))
            
            self._countdown_and_capture()
        except Exception as e:
            write_log(f"_reconnect_and_capture error: {e}")
            self.root.after(0, self.capture_failed)
        finally:
            # Resume keepalive after capture workflow completes
            self.pause_keepalive = False

    def _countdown_and_capture(self):
        try:
            for i in [5, 4, 3, 2, 1]:
                self.root.after(0, lambda v=i: self.canvas.itemconfig(self.ui_countdown, text=str(v)))
                time.sleep(0.8)
            # Hide big countdown, show smaller message
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_countdown, state="hidden"))
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_countdown_msg, text="LOOK AT THE CAMERA!", state="normal"))
            write_log("Countdown complete, capturing photo...")
            
            # Quick camera check before capture - reconnect if needed
            if not self.camera_manager.is_connected():
                write_log("Camera not responding before capture, attempting reconnect...")
                if fully_automated_camera_setup():
                    self.camera_manager.connect_best()
                    write_log("Camera reconnected successfully")
                else:
                    write_log("Camera reconnect failed")

            filename = f"IMG_{int(time.time())}.jpg"
            save_path = os.path.join(os.path.join(TEMP_DIR, "captures"), filename)
            write_log(f"Saving photo to: {save_path}")
            
            # Try capture with automatic retry on failure
            ok = False
            for attempt in range(3):  # Up to 3 attempts
                if attempt > 0:
                    write_log(f"Retry attempt {attempt + 1}/3 - reconnecting camera...")
                    self.root.after(0, lambda: self.canvas.itemconfig(self.ui_countdown_msg, text="Reconnecting camera..."))
                    # Force reconnect
                    if fully_automated_camera_setup():
                        self.camera_manager.connect_best()
                        write_log("Camera reconnected for retry")
                        time.sleep(1)  # Brief pause after reconnect
                    else:
                        write_log("Reconnect failed, trying capture anyway")
                
                ok = self.camera_manager.capture(save_path)
                write_log(f"Capture attempt {attempt + 1} result: {ok}")
                if ok:
                    break
                elif attempt < 2:
                    write_log("Capture failed, will retry...")
                    time.sleep(2)  # Wait before retry
            
            if ok:
                write_log(f"Photo saved successfully, moving to gallery...")
                # Move photo from temp to permanent gallery location
                backup_path = os.path.join(PHOTO_DIR, os.path.basename(save_path))
                try:
                    shutil.move(save_path, backup_path)
                    write_log(f"Photo moved to gallery: {backup_path}")
                except Exception as e:
                    write_log(f"Move failed, trying copy: {e}")
                    try:
                        shutil.copy2(save_path, backup_path)
                        os.remove(save_path)  # Clean up temp file
                        write_log(f"Photo copied to gallery: {backup_path}")
                    except Exception as e2:
                        write_log(f"Copy also failed: {e2}")
                        backup_path = save_path  # Fall back to temp path
                
                # Set current photo to the gallery location (not temp)
                self.current_photo_path = backup_path
                self.root.after(0, self._show_captured_photo)
            else:
                write_log("All capture attempts failed, showing error...")
                self.root.after(0, self.capture_failed)
        except Exception as e:
            write_log(f"_countdown_and_capture exception: {e}")
            self.root.after(0, self.capture_failed)

    def _show_captured_photo(self):
        write_log("_show_captured_photo called")
        try:
            write_log(f"Loading photo from: {self.current_photo_path}")
            # Use zoom viewer to display photo
            if self.zoom_viewer:
                write_log("Using zoom_viewer to load photo...")
                # Add photo to gallery and set as current
                self.zoom_viewer.add_photo(self.current_photo_path)
                self.zoom_viewer.load_photo(self.current_photo_path)
                write_log("zoom_viewer.load_photo completed")
            else:
                write_log("zoom_viewer not available, using fallback...")
                # Fallback if zoom_viewer not initialized
                img = Image.open(self.current_photo_path)
                display_width = int(self.width * 0.8)
                display_height = int(self.height * 0.8)
                ratio = min(display_width / img.width, display_height / img.height)
                new_w = int(img.width * ratio)
                new_h = int(img.height * ratio)
                img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
                self.photo_display = ImageTk.PhotoImage(img_resized)
                self.canvas.itemconfig(self.ui_photo, image=self.photo_display)
                write_log("Fallback photo display completed")
            write_log("Updating UI state to RESULT...")
            self.update_ui_state("RESULT")
            write_log("Photo display complete!")
        except Exception as e:
            write_log(f"Error showing photo: {e}")
            import traceback
            write_log(f"Traceback: {traceback.format_exc()}")
            self.capture_failed()

    def capture_failed(self):
        self.capture_in_progress = False
        self.update_ui_state("ERROR", "Capture failed")

    def action_delete(self):
        # Get the currently viewed photo from zoom_viewer
        photo_to_delete = None
        if hasattr(self, 'zoom_viewer') and self.zoom_viewer and self.zoom_viewer.photos:
            if 0 <= self.zoom_viewer.current_idx < len(self.zoom_viewer.photos):
                photo_to_delete = self.zoom_viewer.photos[self.zoom_viewer.current_idx]
        
        # Fallback to current_photo_path if zoom_viewer doesn't have it
        if not photo_to_delete:
            photo_to_delete = self.current_photo_path
        
        if photo_to_delete and os.path.exists(photo_to_delete):
            # Show themed confirmation dialog
            confirm = ThemedDialog.ask_yes_no(
                self.root,
                "Delete Photo?",
                "Are you sure you want to delete this photo?",
                "🗑️"
            )
            if confirm:
                # Show deleting feedback
                self.btn_delete.config(text="🗑️ DELETING...", state="disabled")
                self.root.update()
                
                try:
                    # Delete local file
                    os.remove(photo_to_delete)
                    write_log(f"Deleted photo: {photo_to_delete}")
                    
                    # Also delete from cloud if it was uploaded
                    filename = os.path.basename(photo_to_delete)
                    cloud_path = os.path.join(DRIVE_DIR, filename)
                    if os.path.exists(cloud_path):
                        os.remove(cloud_path)
                        write_log(f"Deleted from cloud: {cloud_path}")
                except Exception as e:
                    write_log(f"Delete error: {e}")
                
                # Reset button
                self.btn_delete.config(text="🗑️ DELETE", state="normal")
                
                # Update gallery and show next photo if available
                if hasattr(self, 'zoom_viewer') and self.zoom_viewer:
                    old_idx = self.zoom_viewer.current_idx
                    self.zoom_viewer.load_existing_photos(PHOTO_DIR)
                    if self.zoom_viewer.photos:
                        # Show next photo in gallery (or previous if we were at end)
                        new_idx = min(old_idx, len(self.zoom_viewer.photos) - 1)
                        self.zoom_viewer.current_idx = new_idx
                        self.current_photo_path = self.zoom_viewer.photos[new_idx]
                        self.zoom_viewer.load_photo(self.current_photo_path)
                        # Refresh button states for the new photo
                        self.show_buttons()
                        return  # Stay in RESULT mode showing next photo
                
                # No more photos, go home
                self.current_photo_path = None
                self.capture_in_progress = False
                self.update_ui_state("READY", self.camera_manager.get_name())
            # If not confirmed, stay on current photo
        else:
            self.current_photo_path = None
            self.capture_in_progress = False
            self.update_ui_state("READY", self.camera_manager.get_name())

    def action_view_photos(self):
        """Open gallery to view and print existing photos"""
        # Get all photos in the photo directory
        photos = sorted(glob.glob(os.path.join(PHOTO_DIR, "*.jpg")), key=os.path.getmtime, reverse=True)
        if not photos:
            from tkinter import messagebox
            messagebox.showinfo("No Photos", "No photos in gallery yet.\nTap anywhere to take your first photo!")
            return
        
        # Load the most recent photo into viewer
        self.current_photo_path = photos[0]
        
        # Set up the zoom viewer with the gallery (use existing or create with correct params)
        if not self.zoom_viewer:
            self.zoom_viewer = PhotoZoomViewer(self.canvas, self.ui_photo, self.width, self.height)
        
        # Load gallery and display
        self.zoom_viewer.load_existing_photos(PHOTO_DIR)
        self.zoom_viewer.current_idx = 0
        self.zoom_viewer.load_photo(self.current_photo_path)
        
        self.update_ui_state("RESULT")

    def action_go_home(self):
        """Go back to home screen without starting new capture"""
        self.capture_in_progress = False
        self.current_photo_path = None
        self.update_ui_state("READY", self.camera_manager.get_name())

    def action_save(self):
        # Check if photo already uploaded (exists in DRIVE_DIR)
        if self.current_photo_path and os.path.exists(self.current_photo_path):
            filename = os.path.basename(self.current_photo_path)
            dest = os.path.join(DRIVE_DIR, filename)
            
            if os.path.exists(dest):
                # Already uploaded - show message
                ThemedDialog.show_message(
                    self.root,
                    "Already Uploaded",
                    "This photo has already been uploaded to cloud!",
                    "☁️"
                )
                return
        
        # Show confirmation dialog
        confirm = ThemedDialog.ask_yes_no(
            self.root,
            "Upload to Cloud?",
            "Do you want to upload this photo?",
            "☁️"
        )
        if not confirm:
            return
        
        self.btn_save.config(text="☁️ UPLOADING...", state="disabled")
        threading.Thread(target=self._save_thread, daemon=True).start()

    def _save_thread(self):
        success = False
        try:
            if self.current_photo_path and os.path.exists(self.current_photo_path):
                filename = os.path.basename(self.current_photo_path)
                dest = os.path.join(DRIVE_DIR, filename)
                shutil.copy2(self.current_photo_path, dest)
                write_log(f"saved to drive: {dest}")
                success = True
        except Exception as e:
            write_log(f"save error: {e}")
        self.root.after(0, lambda: self._save_complete(success))

    def _save_complete(self, success: bool):
        # Update button to show done status
        if success:
            self.btn_save.config(text="☁️ ✓ DONE", bg="#2D6A4F", fg="white", state="disabled")
            ThemedDialog.show_message(
                self.root,
                "Upload Complete",
                "Photo uploaded to Google Drive!",
                "✓"
            )
        else:
            self.btn_save.config(text="☁️  UPLOAD", bg=THEME_SUCCESS, fg="white", state="normal")
            ThemedDialog.show_message(
                self.root,
                "Upload Failed",
                "Could not upload photo. Check connection.",
                "✗"
            )
        # Stay on current photo

    def action_print(self):
        # Confirm before printing
        confirm = ThemedDialog.ask_yes_no(
            self.root,
            "Print Photo?",
            "Do you want to print this photo?",
            "🖨️"
        )
        if not confirm:
            return
        
        self.btn_print.config(text="🖨️ PRINTING...", state="disabled")
        threading.Thread(target=self._print_thread, daemon=True).start()

    def _print_thread(self):
        global _print_job_counter
        _print_job_counter += 1
        job_number = _print_job_counter
        success = False
        
        try:
            if self.current_photo_path and os.path.exists(self.current_photo_path):
                filename = os.path.basename(self.current_photo_path)
                dest = os.path.join(DRIVE_DIR, filename)
                shutil.copy2(self.current_photo_path, dest)
                write_log(f"[Print #{job_number}] Saved to drive: {dest}")
                
                try:
                    printer_name = win32print.GetDefaultPrinter()
                    write_log(f"[Print #{job_number}] Printer: {printer_name}")
                    
                    img = Image.open(self.current_photo_path)
                    img_width, img_height = img.size
                    write_log(f"[Print #{job_number}] Image size: {img_width}x{img_height} pixels")
                    
                    hdc = win32ui.CreateDC()
                    hdc.CreatePrinterDC(printer_name)
                    
                    # Get printer capabilities
                    # HORZRES (8) = width in pixels, VERTRES (10) = height in pixels
                    # LOGPIXELSX (88) = DPI horizontal, LOGPIXELSY (90) = DPI vertical
                    printable_width = hdc.GetDeviceCaps(8)   # HORZRES
                    printable_height = hdc.GetDeviceCaps(10)  # VERTRES
                    printer_dpi_x = hdc.GetDeviceCaps(88)     # LOGPIXELSX
                    printer_dpi_y = hdc.GetDeviceCaps(90)     # LOGPIXELSY
                    
                    write_log(f"[Print #{job_number}] Printable area: {printable_width}x{printable_height} pixels at {printer_dpi_x}x{printer_dpi_y} DPI")
                    
                    # Calculate margins in printer pixels
                    margin_x = int(PRINT_MARGIN_INCHES * printer_dpi_x)
                    margin_y = int(PRINT_MARGIN_INCHES * printer_dpi_y)
                    
                    # Available print area after margins
                    available_width = printable_width - (2 * margin_x)
                    available_height = printable_height - (2 * margin_y)
                    
                    if PRINT_FIT_TO_PAGE:
                        # Scale image to fit within available area while maintaining aspect ratio
                        img_aspect = img_width / img_height
                        area_aspect = available_width / available_height
                        
                        if img_aspect > area_aspect:
                            # Image is wider than available area - fit to width
                            dest_width = available_width
                            dest_height = int(available_width / img_aspect)
                        else:
                            # Image is taller than available area - fit to height
                            dest_height = available_height
                            dest_width = int(available_height * img_aspect)
                    else:
                        # Use original size scaled to printer DPI
                        # Assume image is 300 DPI if not specified
                        dest_width = int(img_width * printer_dpi_x / 300)
                        dest_height = int(img_height * printer_dpi_y / 300)
                    
                    # Calculate position (centered or top-left)
                    if PRINT_CENTER_ON_PAGE:
                        x_offset = margin_x + (available_width - dest_width) // 2
                        y_offset = margin_y + (available_height - dest_height) // 2
                    else:
                        x_offset = margin_x
                        y_offset = margin_y
                    
                    write_log(f"[Print #{job_number}] Output: {dest_width}x{dest_height} at ({x_offset},{y_offset})")
                    
                    # Start print job
                    hdc.StartDoc(f"PhotoBooth Print #{job_number}")
                    hdc.StartPage()
                    
                    dib = ImageWin.Dib(img)
                    hdc_rect = (x_offset, y_offset, x_offset + dest_width, y_offset + dest_height)
                    dib.draw(hdc.GetHandleOutput(), hdc_rect)
                    
                    hdc.EndPage()
                    hdc.EndDoc()
                    hdc.DeleteDC()
                    
                    write_log(f"[Print #{job_number}] ✓ Print job sent successfully!")
                    success = True
                    
                except Exception as e:
                    write_log(f"[Print #{job_number}] ERROR: {e}")
        except Exception as e:
            write_log(f"[Print #{job_number}] Save/print error: {e}")
        
        self.root.after(0, lambda: self._print_complete(success))

    def _print_complete(self, success: bool):
        if success:
            # Show brief success on button, then reset
            self.btn_print.config(text="🖨️ ✓ SENT!", bg="#2D6A4F", state="disabled")
            self.root.after(2000, lambda: self.btn_print.config(text="🖨️  PRINT", bg=THEME_PRINT, state="normal"))
        else:
            # Show brief error on button, then reset
            self.btn_print.config(text="🖨️ ✗ FAILED", bg=THEME_DANGER, state="disabled")
            self.root.after(2000, lambda: self.btn_print.config(text="🖨️  PRINT", bg=THEME_PRINT, state="normal"))
        # Stay on current photo - no dialog interruption

    def retry_camera(self):
        now = time.time()
        if now - self.last_retry_time < 2:
            return
        self.last_retry_time = now

        def retry_thread():
            try:
                self.camera_manager.disconnect()
                self.root.after(0, self.init_camera)
            except Exception:
                self.root.after(0, self.init_camera)

        threading.Thread(target=retry_thread, daemon=True).start()

    def animate(self):
        if not self.running:
            return
        
        # Increment animation frame
        self.animation_frame += 1
        
        # Pause animations during RESULT mode for better performance
        if self.mode != "RESULT":
            # Update snowflakes
            for flake in self.snowflakes:
                flake.update()
            self.canvas.tag_lower("snow")
            
            # Pulsing year color animation (every 15 frames)
            if self.animation_frame % 15 == 0 and hasattr(self, 'ui_year'):
                import math
                # Cycle between gold colors
                colors = [THEME_ACCENT, "#FFA500", "#FF8C00", "#FFB347", THEME_ACCENT_2, THEME_ACCENT]
                color_idx = (self.animation_frame // 15) % len(colors)
                try:
                    self.canvas.itemconfig(self.ui_year, fill=colors[color_idx])
                except Exception:
                    pass
            
            # Twinkling stars animation (every 3 frames for smooth effect)
            if self.animation_frame % 3 == 0 and hasattr(self, 'twinkling_stars'):
                import math
                for star in self.twinkling_stars:
                    try:
                        # Calculate pulsing size
                        phase = star['phase'] + self.animation_frame * star['speed']
                        scale = 0.6 + 0.4 * abs(math.sin(phase))
                        new_size = int(star['base_size'] * scale)
                        new_size = max(4, min(16, new_size))
                        
                        # Update font size for pulsing effect
                        self.canvas.itemconfig(
                            star['id'],
                            font=("Arial", new_size, "bold")
                        )
                    except Exception:
                        pass
        
        self.root.after(ANIMATION_FPS, self.animate)

    def _zoom_in(self):
        """Keyboard shortcut: zoom in"""
        if self.zoom_viewer and self.mode == "RESULT":
            self.zoom_viewer.zoom_level *= 1.2
            self.zoom_viewer.zoom_level = min(self.zoom_viewer.zoom_level, self.zoom_viewer.max_zoom)
            self.zoom_viewer._display_current_photo_fast()

    def _zoom_out(self):
        """Keyboard shortcut: zoom out"""
        if self.zoom_viewer and self.mode == "RESULT":
            self.zoom_viewer.zoom_level *= 0.8
            self.zoom_viewer.zoom_level = max(self.zoom_viewer.zoom_level, self.zoom_viewer.min_zoom)
            self.zoom_viewer._display_current_photo_fast()

    def _zoom_reset(self):
        """Keyboard shortcut: reset zoom (press 0)"""
        if self.zoom_viewer and self.mode == "RESULT":
            self.zoom_viewer.zoom_level = 1.0
            self.zoom_viewer.pan_x = 0
            self.zoom_viewer.pan_y = 0
            self.zoom_viewer._display_current_photo_fast()

    def _prev_photo(self):
        """Keyboard shortcut: previous photo (Left arrow)"""
        if self.zoom_viewer and self.mode == "RESULT" and self.zoom_viewer.current_idx > 0:
            self.zoom_viewer._switch_photo(self.zoom_viewer.current_idx - 1)

    def _next_photo(self):
        """Keyboard shortcut: next photo (Right arrow)"""
        if self.zoom_viewer and self.mode == "RESULT" and self.zoom_viewer.current_idx < len(self.zoom_viewer.photos) - 1:
            self.zoom_viewer._switch_photo(self.zoom_viewer.current_idx + 1)

    def quit_app(self):
        self.running = False
        self.monitoring_camera = False
        self.camera_manager.disconnect()
        self.root.quit()
        self.root.destroy()


def main():
    print("=" * 60)
    print("New Year Sparkle Booth 2026 - Fully Automated")
    print("=" * 60)
    print("Ensuring administrator privileges...")
    print("=" * 60)
    
    # Elevate to admin once at startup
    ensure_elevated()
    
    # Disable USB power management that can disconnect the camera
    disable_usb_selective_suspend()
    
    print("Camera will auto-bind, auto-attach, and auto-detect.")
    print("(Press ESC to exit, F5 to retry camera)")
    print("=" * 60)

    root = tk.Tk()
    app = PhotoBooth(root)
    root.mainloop()


if __name__ == "__main__":
    main()

"""
New Year Photo Booth 2026 - Fully Automated
Primary: gphoto2 inside WSL (Windows Subsystem for Linux)
Auto-handles: usbipd bind, attach, WSL startup, camera detection
Fallback: Windows Image Acquisition (WIA)
Author: AI Assistant
Date: December 2025
"""

import tkinter as tk
from tkinter import Canvas, Button
from PIL import Image, ImageTk, ImageWin
import os
import time
import threading
import subprocess
import win32print
import win32ui
import random
import shutil
from typing import Optional
from abc import ABC, abstractmethod

# ============================================================================
# CONFIGURATION
# ============================================================================

# Directories
PHOTO_DIR = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'Photobooth', 'Photos')
TEMP_DIR = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'Photobooth_Temp')
DRIVE_DIR = r"G:\My Drive\New_Year_Photo_Booth"
LOG_PATH = os.path.join(TEMP_DIR, "booth.log")
LOGO_CANDIDATES = [
    os.environ.get("LOGO_PATH", ""),
    "logo.png",
    "logo.PNG",
    "logo.jpg",
    "logo.jpeg",
]

os.makedirs(PHOTO_DIR, exist_ok=True)
os.makedirs(TEMP_DIR, exist_ok=True)
try:
    os.makedirs(DRIVE_DIR, exist_ok=True)
except Exception:
    DRIVE_DIR = os.path.join(os.environ['USERPROFILE'], 'Desktop', 'Photobooth_Saved')
    os.makedirs(DRIVE_DIR, exist_ok=True)

# WSL / gphoto2 settings
WSL_DISTRO = "Ubuntu"
GPHOTO_CMD = "gphoto2"

# USB camera settings - Nikon Z6_3
CAMERA_VID_PID = "04b0:0454"
CAMERA_DEVICE_NAME = "Z6_3"
USBIPD_EXE = r"C:\Program Files\usbipd-win\usbipd.exe"

# Theme Colors
THEME_BG = "#0a0e1a"
THEME_ACCENT = "#FFD700"
THEME_TEXT = "#FFFFFF"
THEME_SUCCESS = "#28a745"
THEME_DANGER = "#dc3545"
THEME_PRINT = "#FF6B35"

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
        print(line)
    except Exception:
        pass
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def resolve_logo_path() -> Optional[str]:
    """Find first existing logo file."""
    for path in LOGO_CANDIDATES:
        if path and os.path.exists(path):
            return path
    return None


def run_command(cmd, timeout=10, shell=False):
    """Run a command and return (returncode, stdout, stderr)."""
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, shell=shell)
        return result.returncode, result.stdout, result.stderr
    except subprocess.TimeoutExpired:
        return -1, "", f"Timeout after {timeout}s"
    except Exception as e:
        return -2, "", str(e)


def ensure_wsl_running():
    """Start WSL distro if not running."""
    try:
        run_command(["wsl", "-d", WSL_DISTRO, "--", "bash", "-lc", "exit"], timeout=5)
        return True
    except Exception:
        return False


def get_camera_busid():
    """Query usbipd and return camera BusID, or None if not found."""
    if not os.path.exists(USBIPD_EXE):
        write_log("ERROR: usbipd-win not found at " + USBIPD_EXE)
        return None

    rc, stdout, stderr = run_command([USBIPD_EXE, "list"], timeout=8)
    if rc != 0:
        write_log(f"usbipd list failed: {stderr}")
        return None

    for line in stdout.splitlines():
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
        write_log(f"bind failed: {stderr.strip()}")
        return False


def attach_camera(busid: str) -> bool:
    """Attach camera to WSL (requires --wsl flag in usbipd 5.x)."""
    if not os.path.exists(USBIPD_EXE):
        return False

    write_log(f"usbipd attach --wsl {WSL_DISTRO} --busid {busid}")
    rc, stdout, stderr = run_command(
        [USBIPD_EXE, "attach", "--wsl", WSL_DISTRO, "--busid", busid],
        timeout=20
    )
    
    if rc == 0:
        write_log(f"attach successful")
        return True
    else:
        stderr_lower = stderr.lower()
        if "already attached" in stderr_lower:
            write_log(f"camera already attached (ok)")
            return True
        else:
            write_log(f"attach failed: {stderr.strip()}")
            return False


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

    # Step 4: Attach if needed
    write_log("\nStep 4: Attaching camera to WSL...")
    if state != "attached":
        if not attach_camera(busid):
            write_log("ERROR: Failed to attach camera")
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
                for ln in result.stdout.splitlines():
                    if ln.strip() and not ln.lower().startswith('model') and not ln.strip().startswith('-'):
                        self.camera_name = ln.strip()
                        self.connected = True
                        write_log(f"✓ Camera connected: {self.camera_name}")
                        return True
            
            write_log("✗ Camera detection failed")
            return False
        except Exception as ex:
            write_log(f"WSLGPhotoCamera exception: {ex}")
            return False

    def disconnect(self):
        self.connected = False

    def is_connected(self) -> bool:
        return self.connected

    def capture(self, save_path: str) -> bool:
        try:
            if not self.connected:
                return False

            # Autofocus
            self._run_wsl(f"{GPHOTO_CMD} --set-config autofocusdrive=1 2>&1", timeout=5)

            wsl_path = windows_path_to_wsl(save_path)
            result = self._run_wsl(f"{GPHOTO_CMD} --capture-image-and-download --filename '{wsl_path}' 2>&1", timeout=45)

            return result.returncode == 0 and os.path.exists(save_path)
        except Exception:
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
    def __init__(self, canvas: Canvas, width: int, height: int):
        self.canvas = canvas
        self.width = width
        self.height = height
        self.reset()

    def reset(self):
        self.x = random.randint(0, self.width)
        self.y = random.randint(-self.height, 0)
        self.size = random.randint(2, 7)
        self.speed = random.uniform(*SNOW_SPEED)
        self.drift = random.uniform(-0.5, 0.5)
        self.id = self.canvas.create_oval(
            self.x, self.y, self.x + self.size, self.y + self.size,
            fill="white", outline="", tags="snow"
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

        self.camera_manager = CameraManager()

        self.setup_canvas()
        self.setup_snowflakes()
        self.setup_ui_elements()
        self.setup_buttons()

        self.root.bind("<Escape>", lambda e: self.quit_app())
        self.root.bind("<F5>", lambda e: self.retry_camera())
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
        self.ui_title = self.canvas.create_text(
            center_x, self.height * 0.12,
            text="HAPPY NEW YEAR 2026",
            font=(FONT_FAMILY, TITLE_SIZE + 6, "bold"),
            fill=THEME_ACCENT,
            tags="ui"
        )
        self.ui_instruction = self.canvas.create_text(
            center_x, self.height * 0.32,
            text="Initializing...",
            font=(FONT_FAMILY, SUBTITLE_SIZE + 6, "bold"),
            fill=THEME_TEXT,
            tags="ui"
        )
        self.ui_status = self.canvas.create_text(
            center_x, self.height * 0.40,
            text="",
            font=(FONT_FAMILY, STATUS_SIZE),
            fill=THEME_SUCCESS,
            tags="ui"
        )
        self.ui_countdown = self.canvas.create_text(
            center_x, self.height * 0.52,
            text="",
            font=(FONT_FAMILY, COUNTDOWN_SIZE + 20, "bold"),
            fill=THEME_ACCENT,
            tags="ui"
        )
        self.ui_photo = self.canvas.create_image(
            center_x, self.height * 0.52,
            image=None,
            state="hidden",
            tags="photo"
        )

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
                logo_x = self.width * 0.08
                logo_y = self.height * 0.16
                self.ui_logo = self.canvas.create_image(logo_x, logo_y, image=self.logo_image, tags="logo")
                self.canvas.tag_lower("logo", "snow")
            except Exception as e:
                write_log(f"logo failed: {e}")

    def setup_buttons(self):
        btn_font = (FONT_FAMILY, BUTTON_SIZE, "bold")
        self.btn_delete = Button(self.root, text="🗑️ DELETE", font=btn_font, bg=THEME_DANGER, fg="white",
                                 activebackground="#c82333", activeforeground="white", relief="flat",
                                 cursor="hand2", command=self.action_delete)
        self.btn_save = Button(self.root, text="💾 SAVE TO DRIVE", font=btn_font, bg=THEME_SUCCESS, fg="white",
                               activebackground="#218838", activeforeground="white", relief="flat",
                               cursor="hand2", command=self.action_save)
        self.btn_print = Button(self.root, text="🖨️ PRINT & DRIVE", font=btn_font, bg=THEME_PRINT, fg="white",
                                activebackground="#e55a2b", activeforeground="white", relief="flat",
                                cursor="hand2", command=self.action_print)
        self.hide_buttons()

    def init_camera(self):
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
        """Background health check."""
        self.monitoring_camera = True
        threading.Thread(target=self._monitor_camera_health, daemon=True).start()

    def _monitor_camera_health(self):
        while self.running and getattr(self, 'monitoring_camera', False):
            try:
                if self.mode in ["READY", "COUNTDOWN"] and self.camera_ready:
                    if not gphoto2_detect():
                        write_log("monitor: camera missing, reattaching...")
                        if fully_automated_camera_setup():
                            if self.mode == "READY":
                                self.root.after(0, lambda: self.update_ui_state("READY", "USB reattached"))
                        else:
                            self.root.after(0, lambda: self.update_ui_state("ERROR", "Camera lost"))
                time.sleep(3)
            except Exception:
                time.sleep(3)

    def update_ui_state(self, state: str, message: str = ""):
        self.mode = state
        self.canvas.itemconfig(self.ui_title, state="normal")
        self.canvas.itemconfig(self.ui_instruction, state="hidden")
        self.canvas.itemconfig(self.ui_status, state="hidden")
        self.canvas.itemconfig(self.ui_countdown, state="hidden")
        self.canvas.itemconfig(self.ui_photo, state="hidden")
        self.hide_buttons()

        if state == "INIT":
            self.canvas.itemconfig(self.ui_instruction, text="Setting up the countdown...", state="normal")
            self.canvas.itemconfig(self.ui_status, text="⏳ Binding & attaching camera...", state="normal")
        elif state == "READY":
            self.canvas.itemconfig(self.ui_instruction, text="TAP TO START YOUR 2026 SHOT", fill=THEME_TEXT, state="normal")
            self.canvas.itemconfig(self.ui_status, text=(f"✓ {message}" if message else "✓ Camera ready for midnight"), fill=THEME_SUCCESS, state="normal")
        elif state == "COUNTDOWN":
            self.canvas.itemconfig(self.ui_countdown, state="normal")
            self.canvas.itemconfig(self.ui_status, text="Get ready!", fill=THEME_ACCENT, state="normal")
        elif state == "RESULT":
            self.canvas.itemconfig(self.ui_title, state="hidden")
            self.canvas.itemconfig(self.ui_photo, state="normal")
            self.show_buttons()
        elif state == "ERROR":
            self.canvas.itemconfig(self.ui_instruction, text="Tap to retry", fill=THEME_ACCENT, state="normal")
            self.canvas.itemconfig(self.ui_status, text=f"✗ {message}", fill=THEME_DANGER, state="normal")

    def show_buttons(self):
        btn_y = 0.92
        self.btn_delete.place(relx=0.15, rely=btn_y, anchor="center")
        self.btn_save.place(relx=0.5, rely=btn_y, anchor="center")
        self.btn_print.place(relx=0.85, rely=btn_y, anchor="center")

    def hide_buttons(self):
        for btn in [self.btn_delete, self.btn_save, self.btn_print]:
            btn.place_forget()

    def on_canvas_click(self, event):
        if self.mode == "READY" and self.camera_ready and not self.capture_in_progress:
            self.start_photo_workflow()
        elif self.mode == "ERROR":
            self.retry_camera()

    def start_photo_workflow(self):
        if not self.camera_ready:
            return
        self.capture_in_progress = True
        self.update_ui_state("COUNTDOWN")
        threading.Thread(target=self._countdown_and_capture, daemon=True).start()

    def _countdown_and_capture(self):
        try:
            for i in [3, 2, 1]:
                self.root.after(0, lambda v=i: self.canvas.itemconfig(self.ui_countdown, text=str(v)))
                time.sleep(0.8)
            self.root.after(0, lambda: self.canvas.itemconfig(self.ui_countdown, text="CHEESE!"))

            filename = f"IMG_{int(time.time())}.jpg"
            save_path = os.path.join(TEMP_DIR, filename)
            ok = self.camera_manager.capture(save_path)
            if ok:
                self.current_photo_path = save_path
                try:
                    backup_path = os.path.join(PHOTO_DIR, os.path.basename(save_path))
                    shutil.copy2(save_path, backup_path)
                except Exception:
                    pass
                self.root.after(0, self._show_captured_photo)
            else:
                self.root.after(0, self.capture_failed)
        except Exception:
            self.root.after(0, self.capture_failed)

    def _show_captured_photo(self):
        try:
            img = Image.open(self.current_photo_path)
            display_width = int(self.width * 0.8)
            display_height = int(self.height * 0.8)
            ratio = min(display_width / img.width, display_height / img.height)
            new_w = int(img.width * ratio)
            new_h = int(img.height * ratio)
            img_resized = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            self.photo_display = ImageTk.PhotoImage(img_resized)
            self.canvas.itemconfig(self.ui_photo, image=self.photo_display)
            self.update_ui_state("RESULT")
        except Exception:
            self.capture_failed()

    def capture_failed(self):
        self.capture_in_progress = False
        self.update_ui_state("ERROR", "Capture failed")

    def action_delete(self):
        if self.current_photo_path and os.path.exists(self.current_photo_path):
            try:
                os.remove(self.current_photo_path)
            except Exception:
                pass
        self.current_photo_path = None
        self.capture_in_progress = False
        self.update_ui_state("READY", self.camera_manager.get_name())

    def action_save(self):
        self.btn_save.config(text="💾 SAVING TO DRIVE...", state="disabled")
        threading.Thread(target=self._save_thread, daemon=True).start()

    def _save_thread(self):
        try:
            if self.current_photo_path and os.path.exists(self.current_photo_path):
                filename = os.path.basename(self.current_photo_path)
                dest = os.path.join(DRIVE_DIR, filename)
                shutil.copy2(self.current_photo_path, dest)
                write_log(f"saved to drive: {dest}")
        except Exception as e:
            write_log(f"save error: {e}")
        self.root.after(0, self._save_complete)

    def _save_complete(self):
        self.btn_save.config(text="💾 SAVE TO DRIVE", state="normal")
        self.current_photo_path = None
        self.capture_in_progress = False
        self.update_ui_state("READY", self.camera_manager.get_name())

    def action_print(self):
        self.btn_print.config(text="🖨️ PRINT & DRIVE...", state="disabled")
        threading.Thread(target=self._print_thread, daemon=True).start()

    def _print_thread(self):
        try:
            if self.current_photo_path and os.path.exists(self.current_photo_path):
                filename = os.path.basename(self.current_photo_path)
                dest = os.path.join(DRIVE_DIR, filename)
                shutil.copy2(self.current_photo_path, dest)
                write_log(f"saved to drive (print path): {dest}")
                try:
                    printer_name = win32print.GetDefaultPrinter()
                    img = Image.open(self.current_photo_path)
                    hdc = win32ui.CreateDC()
                    hdc.CreatePrinterDC(printer_name)
                    hdc.StartDoc("PhotoBooth Print")
                    hdc.StartPage()
                    dib = ImageWin.Dib(img)
                    (px, py) = img.size
                    hdc_rect = (0, 0, px, py)
                    dib.draw(hdc.GetHandleOutput(), hdc_rect)
                    hdc.EndPage()
                    hdc.EndDoc()
                    hdc.DeleteDC()
                    write_log("printed")
                except Exception as e:
                    write_log(f"print error: {e}")
        except Exception as e:
            write_log(f"save/print error: {e}")
        self.root.after(0, self._print_complete)

    def _print_complete(self):
        self.btn_print.config(text="🖨️ PRINT & DRIVE", state="normal")
        self.current_photo_path = None
        self.capture_in_progress = False
        self.update_ui_state("READY", self.camera_manager.get_name())

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
        for flake in self.snowflakes:
            flake.update()
        self.canvas.tag_lower("snow")
        self.root.after(ANIMATION_FPS, self.animate)

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
    print("Camera will auto-bind, auto-attach, and auto-detect.")
    print("(Press ESC to exit, fn + F5 to retry camera)")
    print("=" * 60)

    root = tk.Tk()
    app = PhotoBooth(root)
    root.mainloop()


if __name__ == "__main__":
    main()

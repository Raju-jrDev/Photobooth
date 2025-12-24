# PhotoBooth - Manual Commands for WSL/Camera Setup

**Use this if automation fails or you need to troubleshoot manually.**

---

## Quick Start

```bash
cd C:\Users\saikr\OneDrive\Desktop\Photobooth
python photobooth.py
```

---

## Manual WSL Setup (If Automation Fails)

### 1. Check WSL Ubuntu Installation

```powershell
# List installed WSL distributions
wsl --list --verbose

# Expected output should show: ubuntu (Running, version 2)
```

### 2. Start WSL Ubuntu

```powershell
# Open WSL Ubuntu shell
wsl

# Inside WSL, verify you're in home directory
pwd
```

### 3. Update System Packages

```bash
# Inside WSL Ubuntu
sudo apt update
sudo apt upgrade -y
```

### 4. Install gphoto2 (Camera Control)

```bash
# Inside WSL Ubuntu
sudo apt install -y gphoto2 libgphoto2-dev
sudo apt install -y libmtp-dev mtp-tools

# Verify installation
gphoto2 --version
```

### 5. Configure USB Permissions

```bash
# Inside WSL Ubuntu
sudo usermod -a -G plugdev $(whoami)

# Restart to apply group changes (exit WSL, restart terminal)
exit
```

---

## Manual Camera Binding (Windows PowerShell)

### 1. Check USB Device ID

```powershell
# In Windows PowerShell (NOT WSL)
usbipd list
```

Expected output (for Nikon Z6_3):
```
BUSID  VID:PID      DEVICE
1-15   04b0:0454    Nikon Z6_3
```

**Note the BUSID (1-15 in this example)**

### 2. Bind Camera to WSL (First Time Only)

```powershell
# Windows PowerShell - Run as Administrator
usbipd bind --busid 1-15

# Output should show: Successfully bound
```

### 3. Attach Camera to WSL

```powershell
# Windows PowerShell (does NOT need admin, but can be run as admin)
usbipd attach --wsl --busid 1-15

# Output should show: Successfully attached
```

**Note**: The `--wsl` flag is required for WSL 5.x

### 4. Verify Camera is Attached

```powershell
usbipd list
```

Look for BUSID 1-15 with ATTACHED status next to it.

---

## Manual Camera Detection (Inside WSL)

### 1. List Connected Cameras

```bash
# Inside WSL Ubuntu
gphoto2 --list-cameras
```

Expected output:
```
Model                          Port                                            
----------------------------------------------------------
Nikon Z6_3                     usb:001,015
```

### 2. Get Camera Info

```bash
# Inside WSL Ubuntu
gphoto2 --camera "Nikon Z6_3" --summary
```

### 3. Capture Test Photo

```bash
# Inside WSL Ubuntu
gphoto2 --camera "Nikon Z6_3" --capture-image-and-download --filename test_photo.jpg
```

---

## Troubleshooting

### Camera Not Detected?

**Step 1**: Verify USB binding
```powershell
# Windows: Check if bound
usbipd list

# Should show: Nikon Z6_3 with state "Bound"
```

**Step 2**: Attach to WSL
```powershell
# Windows: Attach to WSL
usbipd attach --wsl --busid 1-15
```

**Step 3**: Check inside WSL
```bash
# Inside WSL: List USB devices
lsusb

# Should show Nikon device
# Example: Bus 001 Device 015: ID 04b0:0454 Nikon Corp.
```

**Step 4**: Force gphoto2 reset
```bash
# Inside WSL
pkill -9 gphoto2
sleep 2
gphoto2 --list-cameras
```

### Permission Denied on USB?

```bash
# Inside WSL
sudo gphoto2 --list-cameras

# If this works, user needs group permissions:
sudo usermod -a -G plugdev $USER

# Then exit and restart WSL to apply changes
exit
```

### WSL Not Running?

```powershell
# Windows PowerShell
wsl --status

# Start WSL
wsl

# If issues persist, restart WSL
wsl --shutdown
wsl
```

---

## Testing Camera Manually

### Minimal Test Script

Create `test_camera.py`:

```python
#!/usr/bin/env python3
import subprocess
import os

# Check if gphoto2 works
result = subprocess.run(["gphoto2", "--list-cameras"], capture_output=True, text=True)
print("=== Camera Detection ===")
print(result.stdout)
print(result.stderr)

# Try to take a photo
print("\n=== Attempting Capture ===")
result = subprocess.run(
    ["gphoto2", "--capture-image-and-download", "--filename", "test.jpg"],
    capture_output=True, text=True
)
print(result.stdout)
print(result.stderr)

# Check if file exists
if os.path.exists("test.jpg"):
    size = os.path.getsize("test.jpg")
    print(f"\n✓ Photo captured successfully! Size: {size} bytes")
else:
    print("\n✗ Photo capture failed")
```

Run it:
```bash
cd /path/to/photobooth
python3 test_camera.py
```

---

## Full Manual Automation Sequence

Use this if you need to set everything up from scratch:

### Windows PowerShell

```powershell
# 1. Check USB device
usbipd list

# 2. Bind USB (first time only)
usbipd bind --busid 1-15

# 3. Attach USB to WSL
usbipd attach --wsl --busid 1-15

# 4. Verify attachment
usbipd list
```

### WSL Ubuntu (wsl command opens this)

```bash
# 5. Update packages
sudo apt update
sudo apt upgrade -y

# 6. Install gphoto2
sudo apt install -y gphoto2 libgphoto2-dev libmtp-dev mtp-tools

# 7. Configure USB access
sudo usermod -a -G plugdev $(whoami)

# 8. Exit and restart for group changes
exit
```

### Back to Windows PowerShell

```powershell
# 9. Re-open WSL to apply group changes
wsl
```

### Back to WSL Ubuntu

```bash
# 10. Verify camera detection
gphoto2 --list-cameras

# 11. Test photo capture
gphoto2 --capture-image-and-download --filename test.jpg

# 12. If working, exit WSL
exit
```

### Back to Windows PowerShell

```powershell
# 13. Navigate to PhotoBooth
cd C:\Users\saikr\OneDrive\Desktop\Photobooth

# 14. Run the app
python photobooth.py
```

---

## Environment Setup (Python)

### Create Virtual Environment (First Time)

```powershell
cd C:\Users\saikr\OneDrive\Desktop\Photobooth
python -m venv .venv
.venv\Scripts\Activate.ps1
```

### Install Dependencies

```powershell
# Inside venv
pip install pillow
pip install python-pptx
```

### Run PhotoBooth

```powershell
# Inside venv
python photobooth.py
```

---

## Quick Commands Reference

| Task | Command |
|------|---------|
| List USB devices | `usbipd list` |
| Bind USB | `usbipd bind --busid 1-15` |
| Attach to WSL | `usbipd attach --wsl --busid 1-15` |
| List cameras (WSL) | `gphoto2 --list-cameras` |
| Capture photo (WSL) | `gphoto2 --capture-image-and-download --filename photo.jpg` |
| Open WSL | `wsl` |
| Exit WSL | `exit` |
| Start PhotoBooth | `python photobooth.py` |
| Activate venv | `.venv\Scripts\Activate.ps1` |

---

## Camera Specs (Reference)

- **Model**: Nikon Z6_3
- **USB VID**: 04b0 (Nikon)
- **USB PID**: 0454
- **Default BUSID**: 1-15
- **Connection**: USB 3.1
- **Driver**: gphoto2

---

## File Locations

| Item | Path |
|------|------|
| App | `C:\Users\saikr\OneDrive\Desktop\Photobooth\photobooth.py` |
| Photos Output | `C:\Users\saikr\OneDrive\Desktop\Photobooth\Photos\` |
| Logs | Console output (or pipe to file) |
| Config | Hardcoded in photobooth.py |

---

## Error Messages & Solutions

### "gphoto2: command not found"
```bash
# Install gphoto2
sudo apt install -y gphoto2
```

### "usb permission denied"
```bash
# Add user to plugdev group
sudo usermod -a -G plugdev $(whoami)

# Then restart WSL:
exit
wsl
```

### "gphoto2: camera not found"
```powershell
# Reattach USB
usbipd detach --busid 1-15
usbipd attach --wsl --busid 1-15
```

### "Python module not found"
```powershell
# Inside venv
pip install pillow
# Then restart app
python photobooth.py
```

---

## Status Checks

### Check Python Version
```powershell
python --version
# Expected: Python 3.13.2 or higher
```

### Check WSL Ubuntu
```powershell
wsl --status
# Should show: Running
```

### Check gphoto2
```bash
gphoto2 --version
# Should show version number
```

### Check USB Device
```powershell
usbipd list
# Should show Nikon Z6_3
```

---

## Daily Use

### Start PhotoBooth
```powershell
cd C:\Users\saikr\OneDrive\Desktop\Photobooth
python photobooth.py
```

### If Camera Not Found at Startup
```powershell
# Reattach USB
usbipd attach --wsl --busid 1-15

# Then restart PhotoBooth
python photobooth.py
```

### Check Saved Photos
```powershell
dir C:\Users\saikr\OneDrive\Desktop\Photobooth\Photos
```

---

**Last Updated**: Dec 22, 2025  
**Status**: Ready for manual use

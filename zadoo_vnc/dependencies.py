"""Optional dependency probes and runtime feature flags."""
from __future__ import annotations

import logging
import subprocess
import sys


CORE_PACKAGES = {
    "websockets": "websockets",
    "mss": "mss",
    "pyautogui": "PyAutoGUI",
    "PIL": "Pillow",
    "numpy": "numpy",
    "pyperclip": "pyperclip",
    "keyboard": "keyboard",
    "win32api": "pywin32",
    "dotenv": "python-dotenv",
}

OPTIONAL_PACKAGES = {
    "dxcam": "dxcam",
    "fast_ctypes_screenshots": "fast-ctypes-screenshots",
    "bettercam": "bettercam",
    "winrt": "winrt",
    "imagecodecs": "imagecodecs",
    "aiortc": "aiortc",
    "sounddevice": "sounddevice",
    "av": "av",
    "resend": "resend",
    "soundcard": "soundcard",
    "paramiko": "paramiko",
    "winpty": "pywinpty",
    "paddleocr": "paddleocr",
}


def install_package(package_name):
    """Install a Python package using pip."""
    try:
        print(f"Installing {package_name}...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", package_name, "--quiet"],
            capture_output=True,
            text=True,
            check=True,
        )
        print(f"OK: {package_name} installed")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {package_name}: {e.stderr}")
        return False


def check_and_install_dependencies(include_optional=False):
    """Check and install core dependencies. Optional extras are documented only."""
    packages = dict(CORE_PACKAGES)
    if include_optional:
        packages.update(OPTIONAL_PACKAGES)

    missing_packages = []
    for module_name, package_name in packages.items():
        try:
            __import__(module_name)
            print(f"OK: {package_name}")
        except ImportError:
            print(f"Missing: {package_name}")
            missing_packages.append(package_name)

    if missing_packages:
        print("\n[!] Required packages are missing.")
        print("Required:", ", ".join(missing_packages))
        try:
            print("\n[*] Attempting to install missing packages...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", *missing_packages])
            print("\n[+] Packages installed successfully. Please restart the application.")
        except subprocess.CalledProcessError:
            print("\n[-] Failed to install packages automatically.")
            print("Please install them manually using: pip install " + " ".join(missing_packages))
        sys.exit(1)

    if not include_optional:
        missing_optional = []
        for module_name, package_name in OPTIONAL_PACKAGES.items():
            try:
                __import__(module_name)
            except ImportError:
                missing_optional.append(package_name)
        if missing_optional:
            logging.info("Optional packages not installed: %s", ", ".join(missing_optional))

    print("[+] Core dependencies are satisfied.")


try:
    import mss
    HAS_MSS = True
except ImportError:
    mss = None
    HAS_MSS = False

try:
    import numpy as np
    HAS_NUMPY = True
except ImportError:
    np = None
    HAS_NUMPY = False

try:
    imagecodecs = None
    if HAS_NUMPY:
        import imagecodecs
        HAS_IMAGECODECS = True
    else:
        HAS_IMAGECODECS = False
except ImportError:
    imagecodecs = None
    HAS_IMAGECODECS = False

try:
    import pyautogui
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0
    HAS_PYAUTOGUI = True
except ImportError:
    pyautogui = None
    HAS_PYAUTOGUI = False

try:
    import keyboard
    HAS_KEYBOARD = True
except ImportError:
    keyboard = None
    HAS_KEYBOARD = False

try:
    from PIL import Image, ImageGrab
    HAS_PIL = True
except ImportError:
    Image = None
    ImageGrab = None
    HAS_PIL = False

try:
    import win32api
    import win32con
    import win32gui
    import win32ui
    WIN32_AVAILABLE = True
except ImportError:
    win32api = win32con = win32gui = win32ui = None
    WIN32_AVAILABLE = False

try:
    import dxcam
    HAS_DXCAM = True
except ImportError:
    dxcam = None
    HAS_DXCAM = False

try:
    import bettercam
    HAS_BETTERCAM = True
except ImportError:
    bettercam = None
    HAS_BETTERCAM = False

if HAS_BETTERCAM:
    try:
        logging.info("BetterCam detected. Version: %s", getattr(bettercam, "__version__", "unknown"))
        CamCls = getattr(bettercam, "BetterCam", None)
        if CamCls and not hasattr(CamCls, "_zadoo_patched"):
            orig_stop = getattr(CamCls, "stop", None)

            def _zadoo_safe_stop(self, *args, **kwargs):
                if not hasattr(self, "is_capturing"):
                    try:
                        setattr(self, "is_capturing", False)
                    except Exception:
                        pass
                if orig_stop:
                    try:
                        return orig_stop(self, *args, **kwargs)
                    except Exception:
                        return None

            if orig_stop:
                try:
                    CamCls.stop = _zadoo_safe_stop
                except Exception:
                    pass
            orig_del = getattr(CamCls, "__del__", None)
            if orig_del:

                def _zadoo_safe_del(self):
                    try:
                        return orig_del(self)
                    except Exception:
                        return None

                try:
                    CamCls.__del__ = _zadoo_safe_del
                except Exception:
                    pass
            try:
                CamCls._zadoo_patched = True
            except Exception:
                pass
    except Exception:
        logging.debug("BetterCam patching failed", exc_info=True)

HAS_D3DSHOT = False

try:
    import winrt  # type: ignore
    HAS_WINRT = True
except ImportError:
    winrt = None
    HAS_WINRT = False

try:
    import fast_ctypes_screenshots
    HAS_FAST_CTYPES = True
except ImportError:
    fast_ctypes_screenshots = None
    HAS_FAST_CTYPES = False

try:
    import pyperclip
    HAS_PYPERCLIP = True
except ImportError:
    pyperclip = None
    HAS_PYPERCLIP = False

try:
    import winpty
    HAS_WINPTY = True
except Exception:
    winpty = None
    HAS_WINPTY = False

try:
    from aiortc import MediaStreamTrack, RTCPeerConnection, RTCSessionDescription
    from aiortc.rtcrtpsender import RTCRtpSender
    HAS_AIORTC = True
except Exception:
    MediaStreamTrack = RTCPeerConnection = RTCSessionDescription = RTCRtpSender = None
    HAS_AIORTC = False

try:
    import sounddevice as sd
    HAS_SOUNDDEVICE = True
except Exception:
    sd = None
    HAS_SOUNDDEVICE = False

try:
    import av
    from fractions import Fraction
    HAS_AV = True
except Exception:
    av = None
    Fraction = None
    HAS_AV = False

try:
    import soundcard as sc
    HAS_SOUNDCARD = True
except Exception:
    sc = None
    HAS_SOUNDCARD = False

try:
    import paramiko
    HAS_PARAMIKO = True
except Exception:
    paramiko = None
    HAS_PARAMIKO = False

try:
    import resend
    HAS_RESEND = True
except Exception:
    resend = None
    HAS_RESEND = False

"""Optional dependency probes and runtime feature flags."""
from __future__ import annotations

import logging


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

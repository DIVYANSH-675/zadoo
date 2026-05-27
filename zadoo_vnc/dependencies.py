"""Optional dependency probes and runtime feature flags."""
from __future__ import annotations

import logging
import threading


_BETTERCAM_PATCH_LOCK = threading.RLock()


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
    import dxcam
    HAS_DXCAM = True
except Exception:
    logging.debug("DXCam import/probe failed", exc_info=True)
    dxcam = None
    HAS_DXCAM = False

try:
    import bettercam
    HAS_BETTERCAM = True
except Exception:
    logging.debug("BetterCam import/probe failed", exc_info=True)
    bettercam = None
    HAS_BETTERCAM = False

if HAS_BETTERCAM:
    try:
        logging.info("BetterCam detected. Version: %s", getattr(bettercam, "__version__", "unknown"))
        CamCls = getattr(bettercam, "BetterCam", None)
        with _BETTERCAM_PATCH_LOCK:
            should_patch = bool(CamCls and not getattr(CamCls, "_zadoo_patched", False))
            if should_patch:
                orig_stop = getattr(CamCls, "_zadoo_original_stop", None) or getattr(CamCls, "stop", None)
                orig_del = getattr(CamCls, "_zadoo_original_del", None) or getattr(CamCls, "__del__", None)
                try:
                    CamCls._zadoo_original_stop = orig_stop
                    CamCls._zadoo_original_del = orig_del
                except Exception:
                    pass

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
    import sounddevice as sd
    HAS_SOUNDDEVICE = True
except Exception:
    sd = None
    HAS_SOUNDDEVICE = False

try:
    import av
    HAS_AV = True
except Exception:
    av = None
    HAS_AV = False

try:
    import soundcard as sc
except Exception:
    sc = None

try:
    import resend
    HAS_RESEND = True
except Exception:
    resend = None
    HAS_RESEND = False

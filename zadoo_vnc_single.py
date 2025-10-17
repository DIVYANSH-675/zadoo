# --- Custom Alert (type-to-alert) helpers ---
def _capture_char(self, ch: str):
    # Append one character if capture is active
    try:
        if getattr(self, "custom_alert_active", False):
            self.custom_alert_buf.append(ch)
            _log_try_ok("_capture_char.append", ch)
    except Exception:
        _log_except("_capture_char", sys.exc_info()[1])

def _capture_backspace(self):
    try:
        if getattr(self, "custom_alert_active", False) and self.custom_alert_buf:
            self.custom_alert_buf.pop()
            _log_try_ok("_capture_backspace.pop")
    except Exception:
        _log_except("_capture_backspace", sys.exc_info()[1])
import subprocess as _sp
import ctypes as _ct

def _run_hidden(cmd, *, timeout=8, text=True):
    """Run a child process without flashing a window in a frozen GUI app."""
    si = _sp.STARTUPINFO()
    si.dwFlags |= _sp.STARTF_USESHOWWINDOW
    cf = getattr(_sp, 'CREATE_NO_WINDOW', 0) | getattr(_sp, 'DETACHED_PROCESS', 0x00000008)
    return _sp.run(cmd, capture_output=True, text=text, timeout=timeout,
                   startupinfo=si, creationflags=cf)
#!/usr/bin/env python3
"""
ZADDOO
"""
import sys
try:
    import asyncio
    if sys.platform.startswith("win"):
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
except Exception:
    pass

try:
    import winpty._winpty  # noqa: F401
except Exception:
    pass
import subprocess
import os
import ctypes
import threading
import time
import signal
import atexit
import urllib.request
import urllib.parse
import platform
import asyncio
import json
import socket
import threading
import io
import time
import logging
import queue
import http.server
import functools
from http.server import HTTPServer, BaseHTTPRequestHandler
from contextlib import contextmanager
from typing import Set, Optional
from collections import deque
import websockets
from websockets.http11 import Response as WSResponse
from websockets.datastructures import Headers
from datetime import datetime

from ctypes import wintypes, windll

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

# Reduce noisy websockets handshake logs
try:
    logging.getLogger("websockets.server").setLevel(logging.WARNING)
except Exception:
    pass

# Hide console window only in packaged (frozen) mode to keep dev runs visible
try:
    if getattr(sys, 'frozen', False) or os.environ.get('HIDE_CONSOLE') == '1':
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)  
except Exception:
    pass

# --- Log to both terminal and file (daily file in ./logs) ---
_LOG_FILE_HANDLE = None
def _setup_logging_to_file():
    global _LOG_FILE_HANDLE
    try:
        base_dir = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
        log_dir = os.path.join(base_dir, 'logs')
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, f"zadoo_{datetime.now():%Y%m%d}.log")

        class _Tee(io.TextIOBase):
            def __init__(self, stream, file):
                self._stream = stream
                self._file = file
                self._lock = threading.Lock()
            def write(self, s):
                with self._lock:
                    try:
                        self._stream.write(s)
                    except Exception:
                        pass
                    try:
                        self._file.write(s)
                        self._file.flush()
                    except Exception:
                        pass
                return len(s)
            def flush(self):
                with self._lock:
                    try:
                        self._stream.flush()
                    except Exception:
                        pass
                    try:
                        self._file.flush()
                    except Exception:
                        pass

        # Open file in line-buffered append mode
        fh = open(log_path, 'a', encoding='utf-8', buffering=1)
        sys.stdout = _Tee(sys.stdout, fh)
        sys.stderr = _Tee(sys.stderr, fh)
        _LOG_FILE_HANDLE = fh
        atexit.register(lambda: (_LOG_FILE_HANDLE and _LOG_FILE_HANDLE.close()))
        print(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] 📜 Logging to {log_path}")
    except Exception as e:
        try:
            print(f"⚠️ Logging setup failed: {e}")
        except Exception:
            pass

_setup_logging_to_file()

# Configure Python logging to include timestamps (goes to stderr → captured by tee)
try:
    logging.basicConfig(level=logging.INFO,
                        format='[%(asctime)s] %(levelname)s %(message)s',
                        datefmt='%Y-%m-%d %H:%M:%S')
except Exception:
    pass

# --- Simple logging helpers for consistent TRY/EXCEPT markers --------------------
def _log_try_ok(label: str, extra: str = ""):
    try:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{ts}] TRY_OK {label}{(' ' + extra) if extra else ''}")
    except Exception:
        pass

def _log_except(label: str, e: Exception):
    try:
        ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{ts}] EXCEPT {label}: {e}")
    except Exception:
        pass

# --- Network helpers -----------------------------------------------------------
def get_local_ip() -> str:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        try:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
        finally:
            s.close()
    except Exception:
        try:
            return socket.gethostbyname(socket.gethostname())
        except Exception:
            return "127.0.0.1"

# --- Function call/exception logging decorators --------------------------------
def log_calls(name: str = None, level: int = logging.INFO):
    """Decorator to log function entry/exit, duration, and exceptions (sync/async)."""
    def _decorator(func):
        label = name or getattr(func, '__qualname__', getattr(func, '__name__', 'func'))
        try:
            import asyncio as _asyncio
        except Exception:
            _asyncio = None

        def _safe(v, limit=200):
            try:
                s = str(v)
            except Exception:
                s = object.__repr__(v)
            return (s if len(s) <= limit else (s[:limit] + '…'))

        if _asyncio and _asyncio.iscoroutinefunction(func):
            async def _aw(*args, **kwargs):
                t0 = time.perf_counter()
                try:
                    logging.log(level, f"[CALL] {label} args={_safe(args[:1])} kwargs_keys={list(kwargs.keys())}")
                except Exception:
                    pass
                try:
                    result = await func(*args, **kwargs)
                    try:
                        dt = (time.perf_counter() - t0) * 1000.0
                        logging.log(level, f"[RETURN] {label} dt_ms={dt:.2f}")
                    except Exception:
                        pass
                    return result
                except Exception as e:
                    logging.error(f"[EXC] {label}", exc_info=True)
                    raise
            return functools.wraps(func)(_aw)
        else:
            def _w(*args, **kwargs):
                t0 = time.perf_counter()
                try:
                    logging.log(level, f"[CALL] {label} args={_safe(args[:1])} kwargs_keys={list(kwargs.keys())}")
                except Exception:
                    pass
                try:
                    result = func(*args, **kwargs)
                    try:
                        dt = (time.perf_counter() - t0) * 1000.0
                        logging.log(level, f"[RETURN] {label} dt_ms={dt:.2f}")
                    except Exception:
                        pass
                    return result
                except Exception as e:
                    logging.error(f"[EXC] {label}", exc_info=True)
                    raise
            return functools.wraps(func)(_w)
    return _decorator

# --- Auto-instrumentation for class methods ------------------------------------
def _instrument_class_methods(cls):
    """Wrap all methods of a class with log_calls, preserving static/class methods."""
    try:
        import inspect as _inspect
    except Exception:
        _inspect = None
    for _n, _attr in list(getattr(cls, '__dict__', {}).items()):
        if _n.startswith('__'):
            continue
        try:
            is_static = isinstance(_attr, staticmethod)
            is_class = isinstance(_attr, classmethod)
            func = _attr.__func__ if (is_static or is_class) else _attr
            if not callable(func):
                continue
            # Avoid wrapping properties or already-wrapped functions
            if _inspect and (isinstance(func, property) or getattr(func, '__wrapped__', None)):
                continue
            wrapped = log_calls(f"{cls.__name__}.{_n}")(func)
            if is_static:
                setattr(cls, _n, staticmethod(wrapped))
            elif is_class:
                setattr(cls, _n, classmethod(wrapped))
            else:
                setattr(cls, _n, wrapped)
        except Exception:
            # Best-effort; skip problematic attributes
            continue

# --- Mouse State Tracking Functions ---
user32 = windll.user32

# Help coordinates match physical pixels on HiDPI
try:
    user32.SetProcessDPIAware()
except Exception:
    pass

# --- Line-by-line tracer ---------------------------------------------------------
def _make_line_tracer(include_files: tuple):
    """Return a sys.settrace function that logs each executed line for given files."""
    include_files = tuple(os.path.abspath(p) for p in include_files)

    def _tracer(frame, event, arg):
        try:
            fpath = os.path.abspath(frame.f_code.co_filename)
            if fpath not in include_files:
                return _tracer

            fname = frame.f_code.co_name
            lineno = frame.f_lineno

            if event == 'call':
                # Log function call with arguments
                try:
                    arg_names = frame.f_code.co_varnames[: frame.f_code.co_argcount]
                    local_map = frame.f_locals or {}
                    parts = []
                    for name in arg_names:
                        try:
                            val = local_map.get(name, '<unset>')
                            sval = repr(val)
                            if len(sval) > 200:
                                sval = sval[:200] + '…'
                            parts.append(f"{name}={sval}")
                        except Exception:
                            parts.append(f"{name}=<error>")
                    arg_str = ", ".join(parts)
                    logging.info("[trace.call] %s:%d in %s(%s)", fpath, lineno, fname, arg_str)
                    print(f"[trace.call] {fpath}:{lineno} in {fname}({arg_str})", flush=True)
                except Exception:
                    pass

            elif event == 'line':
                try:
                    # Attempt to fetch the current line text
                    with open(fpath, 'r', encoding='utf-8', errors='ignore') as f:
                        for i, line in enumerate(f, start=1):
                            if i == lineno:
                                code_line = line.rstrip('\n')
                                break
                        else:
                            code_line = ''
                except Exception:
                    code_line = ''
                try:
                    logging.info("[trace] %s:%d in %s | %s", fpath, lineno, fname, code_line)
                except Exception:
                    pass
                try:
                    # Tag control-flow lines explicitly
                    _sl = code_line.lstrip()
                    _tag = None
                    if _sl.startswith('try:'):
                        _tag = 'TRY'
                    elif _sl.startswith('except'):
                        _tag = 'EXCEPT'
                    elif _sl.startswith('finally:'):
                        _tag = 'FINALLY'
                    elif _sl.startswith('if '):
                        _tag = 'IF'
                    elif _sl.startswith('elif '):
                        _tag = 'ELIF'
                    elif _sl == 'else:':
                        _tag = 'ELSE'
                    if _tag:
                        print(f"{_tag} line {lineno} - {code_line}", flush=True)
                    # Minimal per-line progress format as requested
                    print(f"line {lineno} done - {code_line}", flush=True)
                except Exception:
                    pass

            elif event == 'return':
                try:
                    r = repr(arg)
                    if len(r) > 200:
                        r = r[:200] + '…'
                    logging.info("[trace.return] %s:%d in %s → %s", fpath, lineno, fname, r)
                    print(f"[trace.return] {fpath}:{lineno} in {fname} → {r}", flush=True)
                except Exception:
                    pass

            elif event == 'exception':
                try:
                    etype, evalue, _ = arg or (None, None, None)
                    ename = getattr(etype, '__name__', str(etype))
                    evals = repr(evalue)
                    if len(evals) > 200:
                        evals = evals[:200] + '…'
                    logging.info("[trace.exc] %s:%d in %s !! %s: %s", fpath, lineno, fname, ename, evals)
                    print(f"[trace.exc] {fpath}:{lineno} in {fname} !! {ename}: {evals}", flush=True)
                    # Plain exception tag line
                    print(f"EXCEPT raised at line {lineno} - {ename}: {evals}", flush=True)
                except Exception:
                    pass
        except Exception:
            pass
        return _tracer

    return _tracer

# --- Mouse control using Windows API directly ---
# We use Windows API instead of PyAutoGUI to avoid ctypes POINT errors

# --- Win32 structs ---
# Use canonical wintypes.POINT to avoid cross-module type mismatches
POINT = wintypes.POINT
LPPOINT = ctypes.POINTER(POINT)

# --- Win32 APIs we use ---
# Bind a local prototype to avoid global argtypes clashes from other modules
GetCursorPosProto = ctypes.WINFUNCTYPE(wintypes.BOOL, LPPOINT)
_GetCursorPos = GetCursorPosProto(("GetCursorPos", user32))

_GetAsyncKeyState = user32.GetAsyncKeyState
_GetAsyncKeyState.argtypes = [wintypes.INT]
_GetAsyncKeyState.restype = wintypes.SHORT

# Virtual-screen metrics (multi-monitor)
SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79

VK_LBUTTON = 0x01
VK_RBUTTON = 0x02

# Mouse injection constants and SendInput structures
MOUSEEVENTF_MOVE = 0x0001
MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040
MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_ABSOLUTE = 0x8000
MOUSEEVENTF_VIRTUALDESK = 0x4000

# ---- Cursor shape detection (Windows) ----
class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hCursor", wintypes.HANDLE),
        ("ptScreenPos", POINT),
    ]

CURSOR_SHOWING = 0x00000001

try:
    user32.GetCursorInfo.argtypes = [ctypes.POINTER(CURSORINFO)]
    user32.GetCursorInfo.restype = wintypes.BOOL
    user32.LoadCursorW.argtypes = [wintypes.HINSTANCE, ctypes.c_void_p]
    user32.LoadCursorW.restype = wintypes.HANDLE
except Exception:
    pass

# Standard cursor IDs (IDC_*)
IDC_ARROW       = 32512
IDC_IBEAM       = 32513
IDC_WAIT        = 32514
IDC_CROSS       = 32515
IDC_UPARROW     = 32516
IDC_SIZENWSE    = 32642
IDC_SIZENESW    = 32643
IDC_SIZEWE      = 32644
IDC_SIZENS      = 32645
IDC_SIZEALL     = 32646
IDC_NO          = 32648
IDC_HAND        = 32649
IDC_APPSTARTING = 32650
IDC_HELP        = 32651

__CURSOR_HANDLE_TO_CSS = {}

def _init_cursor_map():
    global __CURSOR_HANDLE_TO_CSS
    if __CURSOR_HANDLE_TO_CSS:
        return
    css_to_idc = {
        "default":      IDC_ARROW,
        "text":         IDC_IBEAM,
        "wait":         IDC_WAIT,
        "crosshair":    IDC_CROSS,
        "nwse-resize":  IDC_SIZENWSE,
        "nesw-resize":  IDC_SIZENESW,
        "ew-resize":    IDC_SIZEWE,
        "ns-resize":    IDC_SIZENS,
        "move":         IDC_SIZEALL,
        "not-allowed":  IDC_NO,
        "pointer":      IDC_HAND,
        "progress":     IDC_APPSTARTING,
        "help":         IDC_HELP,
    }
    for css, cid in css_to_idc.items():
        try:
            h = user32.LoadCursorW(None, ctypes.c_void_p(cid))
            if h:
                __CURSOR_HANDLE_TO_CSS[int(h)] = css
        except Exception:
            continue

def _get_css_cursor_from_system() -> str:
    try:
        _init_cursor_map()
        ci = CURSORINFO()
        ci.cbSize = ctypes.sizeof(CURSORINFO)
        if not user32.GetCursorInfo(ctypes.byref(ci)):
            return "default"
        if not (ci.flags & CURSOR_SHOWING):
            return "default"
        css = __CURSOR_HANDLE_TO_CSS.get(int(ci.hCursor))
        return css or "default"
    except Exception:
        return "default"

try:
    ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong

    class MOUSEINPUT(ctypes.Structure):
        _fields_ = (
            ("dx", wintypes.LONG),
            ("dy", wintypes.LONG),
            ("mouseData", wintypes.DWORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ULONG_PTR),
        )

    class _INPUT_UNION(ctypes.Union):
        _fields_ = (("mi", MOUSEINPUT),)

    class INPUT(ctypes.Structure):
        _fields_ = (("type", wintypes.DWORD), ("union", _INPUT_UNION))

    def _sendinput_mouse_move_abs(ax, ay):
        try:
            inp = INPUT()
            inp.type = 0  # INPUT_MOUSE
            inp.union.mi = MOUSEINPUT(ax, ay, 0, MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK, 0, 0)
            sent = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            return bool(sent)
        except Exception:
            return False

    def _sendinput_mouse_button(flag):
        try:
            inp = INPUT()
            inp.type = 0  # INPUT_MOUSE
            inp.union.mi = MOUSEINPUT(0, 0, 0, flag, 0, 0)
            sent = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            return bool(sent)
        except Exception:
            return False
except Exception:
    # If SendInput path fails to initialize, we'll fall back to mouse_event below
    def _sendinput_mouse_move_abs(ax, ay):
        return False
    def _sendinput_mouse_button(flag):
        return False

def get_cursor_pos():
    """Return (x, y) screen coordinates of the cursor."""
    # Prefer GetCursorInfo to avoid LP_POINT argtype conflicts across modules
    ci = CURSORINFO()
    ci.cbSize = ctypes.sizeof(CURSORINFO)
    if user32.GetCursorInfo(ctypes.byref(ci)):
        return ci.ptScreenPos.x, ci.ptScreenPos.y
    # Fallback to GetCursorPos if available
    try:
        pt = POINT()
        if _GetCursorPos(ctypes.byref(pt)):
            return pt.x, pt.y
    except Exception:
        pass
    return 0, 0

def left_pressed():
    """True while left mouse button is down."""
    return bool(_GetAsyncKeyState(VK_LBUTTON) & 0x8000)

def right_pressed():
    """True while right mouse button is down."""
    return bool(_GetAsyncKeyState(VK_RBUTTON) & 0x8000)

def get_virtual_screen_bounds():
    """Return (vx, vy, vw, vh) for the entire virtual desktop (all monitors)."""
    vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
    vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
    vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
    vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    if vw > 0 and vh > 0:
        return vx, vy, vw, vh
    return 0, 0, 1920, 1080  # fallback default

class ProcessProtector:
    def __init__(self):
        self.protected = True
        self.start_protection()
    
    def start_protection(self):
        """Start protection mechanisms"""
        # Register cleanup handlers
        atexit.register(self.cleanup)
        signal.signal(signal.SIGTERM, self.signal_handler)
        signal.signal(signal.SIGINT, self.signal_handler)
        
        # Start monitoring thread
        self.monitor_thread = threading.Thread(target=self.monitor_processes, daemon=True)
        self.monitor_thread.start()
    
    def monitor_processes(self):
        """Monitor for termination attempts"""
        while self.protected:
            try:
                time.sleep(30)  # Check every 30 seconds
            except:
                pass
    
    def restart_protection(self):
        """Restart protection mechanisms"""
        try:
            subprocess.Popen([sys.executable, __file__], 
                           creationflags=subprocess.CREATE_NO_WINDOW)
        except:
            pass
    
    def signal_handler(self, signum, frame):
        """Handle termination signals"""
        if self.protected:
            self.restart_protection()
    
    def cleanup(self):
        """Cleanup on exit"""
        self.protected = False

# Initialize protection
protector = ProcessProtector()




def install_package(package_name):
    """Install a Python package using pip"""
    try:
        print(f"Installing {package_name}...")
        result = subprocess.run([
            sys.executable, '-m', 'pip', 'install', package_name, '--quiet'
        ], capture_output=True, text=True, check=True)
        print(f"OK: {package_name} installed")
        return True
    except subprocess.CalledProcessError as e:
        print(f"Failed to install {package_name}: {e.stderr}")
        return False

def check_and_install_dependencies():
    """Check and install all required dependencies"""
    print("Setting up dependencies (one-time setup)...")
    
    required_packages = {
        'websockets': 'websockets',
        'mss': 'mss', 
        'pyautogui': 'pyautogui',
        'PIL': 'Pillow',
        'numpy': 'numpy',
        'dxcam': 'dxcam',
        'fast_ctypes_screenshots': 'fast-ctypes-screenshots',
        'bettercam': 'bettercam',
        'winrt': 'winrt',
        'imagecodecs': 'imagecodecs',
        'pyperclip': 'pyperclip',
        'psutil': 'psutil',
        'pygame': 'pygame',
        'aiortc': 'aiortc',
        'sounddevice': 'sounddevice',
        'av': 'av',
        'resend': 'resend',
        'soundcard': 'soundcard',
        'paramiko': 'paramiko'
    }
    
    # Check if running in a PyInstaller bundle
    if getattr(sys, 'frozen', False):
        bundle_dir = sys._MEIPASS
    else:
        bundle_dir = os.path.dirname(os.path.abspath(__file__))

    # Check for cloudflared executable (non-fatal; will attempt download later if missing)
    CLOUDFLARED_PATH = os.path.join(bundle_dir, 'cloudflared.exe')
    if not os.path.exists(CLOUDFLARED_PATH):
        try:
            print(f"⚠️ cloudflared.exe not found at {CLOUDFLARED_PATH}; will try to download when starting tunnel.")
        except Exception:
            # Best-effort logging only
            pass

    missing_packages = []
    
    for module_name, package_name in required_packages.items():
        try:
            __import__(module_name)
            print(f"OK: {package_name}")
        except ImportError:
            print(f"Missing: {package_name}")
            missing_packages.append(package_name)
    
    if missing_packages:
        print("\n[!] Some required packages are missing!")
        print("Required:", ', '.join(missing_packages))
        
        try:
            # Attempt to install missing packages via pip
            print("\n[*] Attempting to install missing packages...")
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', *missing_packages])
            print("\n[+] Packages installed successfully. Please restart the application.")
        except subprocess.CalledProcessError:
            print("\n[-] Failed to install packages automatically.")
            print("Please install them manually using: pip install " + ' '.join(missing_packages))
        
        # Exit after attempting install
        time.sleep(10)
        sys.exit(1)
        
    print("[+] All required dependencies are satisfied.")

# Now i
try:
    import mss
    HAS_MSS = True
except ImportError:
    HAS_MSS = False
try:
    import numpy as np
    import imagecodecs
    HAS_IMAGECODECS = True
except ImportError:
    HAS_IMAGECODECS = False
try:
    import pyautogui
    pyautogui.FAILSAFE = False
    pyautogui.PAUSE = 0
    HAS_PYAUTOGUI = True
except ImportError:
    HAS_PYAUTOGUI = False

try:
    import keyboard
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False
try:
    from PIL import Image, ImageGrab
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
try:
    import win32api, win32con, win32gui, win32ui
    WIN32_AVAILABLE = True
except ImportError:
    WIN32_AVAILABLE = False
try:
    import dxcam
    HAS_DXCAM = True
except ImportError:
    HAS_DXCAM = False
# Additional capture libs
try:
    import bettercam
    HAS_BETTERCAM = True
except ImportError:
    HAS_BETTERCAM = False
if HAS_BETTERCAM:
    try:
        logging.info(f"BetterCam detected. Version: {getattr(bettercam, '__version__', 'unknown')}")
    except Exception:
        pass
if 'HAS_BETTERCAM' in globals() and HAS_BETTERCAM:
    try:
        CamCls = getattr(bettercam, 'BetterCam', None)
        if CamCls and not hasattr(CamCls, '_zadoo_patched'):
            orig_stop = getattr(CamCls, 'stop', None)
            def _zadoo_safe_stop(self, *args, **kwargs):
                if not hasattr(self, 'is_capturing'):
                    try:
                        setattr(self, 'is_capturing', False)
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
            orig_del = getattr(CamCls, '__del__', None)
            if orig_del:
                def _zadoo_safe_del(self):
                    try:
                        return orig_del(self)
                    except Exception:
                        pass
                try:
                    CamCls.__del__ = _zadoo_safe_del
                except Exception:
                    pass
            try:
                CamCls._zadoo_patched = True
            except Exception:
                pass
    except Exception:
        pass
HAS_D3DSHOT = False
try:
    import winrt  # type: ignore
    HAS_WINRT = True
except ImportError:
    HAS_WINRT = False
# Fast capture using fast_ctypes_screenshots library
try:
    import fast_ctypes_screenshots
    HAS_FAST_CTYPES = True
    print("✅ fast_ctypes_screenshots imported successfully")
except ImportError:
    HAS_FAST_CTYPES = False
    print("❌ fast_ctypes_screenshots not available")
try:
    import pyperclip
    HAS_PYPERCLIP = True
except ImportError:
    HAS_PYPERCLIP = False
try:
    import winpty  # Windows PTY for proper interactive console behavior
    HAS_WINPTY = True
except Exception:
    HAS_WINPTY = False

# Optional audio/WebRTC deps
try:
    from aiortc import RTCPeerConnection, RTCSessionDescription, MediaStreamTrack
    from aiortc.rtcrtpsender import RTCRtpSender
    HAS_AIORTC = True
except Exception:
    HAS_AIORTC = False
try:
    import sounddevice as sd
    HAS_SOUNDDEVICE = True
except Exception:
    HAS_SOUNDDEVICE = False
try:
    import av
    from fractions import Fraction
    HAS_AV = True
except Exception:
    HAS_AV = False
try:
    import soundcard as sc
    HAS_SOUNDCARD = True
except Exception:
    HAS_SOUNDCARD = False
try:
    import paramiko
    HAS_PARAMIKO = True
except Exception:
    HAS_PARAMIKO = False
try:
    import resend
    HAS_RESEND = True
except Exception:
    HAS_RESEND = False

# --- Optional PaddleOCR support (English + code-friendly)
try:
    from paddleocr import PaddleOCR
    HAS_PADDLEOCR = True
except Exception:
    HAS_PADDLEOCR = False

_PADDLE_OCR = None

def _get_paddle_ocr():
    """
    Lazy-initialize PaddleOCR (English).
    Set use_angle_cls for rotated text; keep CPU by default.
    """
    global _PADDLE_OCR
    if _PADDLE_OCR is None:
        if not HAS_PADDLEOCR:
            raise ImportError(
                "paddleocr not installed. Install with:\n"
                "  pip install paddlepaddle paddleocr"
            )
        # Prefer modern args; fall back loosely for compatibility
        try:
            _PADDLE_OCR = PaddleOCR(
                use_textline_orientation=True,
                lang='en'
            )
        except TypeError:
            _PADDLE_OCR = PaddleOCR(lang='en')
    return _PADDLE_OCR

# Warm up OCR models once at startup to avoid first-call latency
def _warmup_ocr_once():
    try:
        ocr = _get_paddle_ocr()
        import numpy as _np
        # Tiny dummy image triggers model weights load and JIT warmups
        _ = ocr.predict(_np.zeros((32, 128, 3), dtype=_np.uint8))
        try:
            print("🔥 OCR warmup complete", flush=True)
        except Exception:
            pass
    except Exception:
        pass

# --- Setup Logging --- (append a rotating file handler without hijacking console)
try:
    from logging.handlers import RotatingFileHandler
    _file_handler = RotatingFileHandler('vnc_debug.log', maxBytes=5_000_000, backupCount=2, encoding='utf-8')
    _file_handler.setFormatter(logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s'))
    _root = logging.getLogger()
    if not any(isinstance(h, RotatingFileHandler) for h in _root.handlers):
        _root.addHandler(_file_handler)
except Exception:
    pass
# --- Load .env (optional) ---
def _load_dotenv(path: str = '.env'):
    """Load simple KEY=VALUE pairs from .env in CWD and script directory."""
    def _apply(p: str):
        try:
            if not os.path.exists(p):
                return
            with open(p, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    if '=' not in line:
                        continue
                    key, val = line.split('=', 1)
                    key = key.strip()
                    val = val.strip().strip('"').strip("'")
                    os.environ.setdefault(key, val)
        except Exception:
            pass
    # Try current working directory
    _apply(path)
    # Try alongside this file
    try:
        base = os.path.dirname(os.path.abspath(__file__))
        _apply(os.path.join(base, '.env'))
    except Exception:
        pass

_load_dotenv()

# --- Configuration ---
QUALITY = 85
STOP_FILE = "stop_vnc.flag"
BRAND_HEADER_IMAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "brand-header.png")
TRIGGER_ICON_IMAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "trigger-icon.png")
SPLASH_IMAGE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "splash.png")

# --- Embedded HTML Client --- (same as original)
HTML_CONTENT = r"""
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Zadoo - Remote Desktop</title>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/codemirror@5/lib/codemirror.min.css">
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/lib/codemirror.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/addon/edit/closebrackets.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/addon/edit/matchbrackets.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/javascript/javascript.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/python/python.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/xml/xml.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/css/css.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/htmlmixed/htmlmixed.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/clike/clike.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/codemirror@5/mode/json/json.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; 
            background: #000;
            overflow: hidden; 
            height: 100vh;
        }
        
        /* Main container */
        .container { 
            display: flex; 
            height: 100vh; 
            position: relative;
        }
        @supports (height: 100dvh) {
            .container { height: 100dvh; }
        }
        
        /* Video area */
        .video-area {
            flex: 1;
            display: flex;
            flex-direction: column;
            background: #000;
            position: relative;
        }
        
        #screen { 
            width: 100%; 
            height: 100vh; 
            object-fit: contain; 
            background: #000;
            cursor: default;
        }
        @supports (height: 100dvh) {
            #screen { height: 100dvh; }
        }
        
        /* Dynamic cursor styles */
        .cursor-default { cursor: default !important; }
        .cursor-pointer { cursor: pointer !important; }
        .cursor-text { cursor: text !important; }
        .cursor-wait { cursor: wait !important; }
        .cursor-move { cursor: move !important; }
        .cursor-grab { cursor: grab !important; }
        .cursor-grabbing { cursor: grabbing !important; }
        .cursor-crosshair { cursor: crosshair !important; }
        .cursor-help { cursor: help !important; }
        .cursor-not-allowed { cursor: not-allowed !important; }
        
        /* Status bar */
        .status-bar {
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            background: rgba(0,0,0,0.8);
            color: white;
            padding: 8px 15px;
            font-size: 12px;
            z-index: 100;
            display: flex;
            justify-content: space-between;
            align-items: center;
            backdrop-filter: blur(10px);
        }
        
        .status-left { display: flex; align-items: center; gap: 15px; }
        .status-right { display: flex; align-items: center; gap: 10px; }
        
        .status-dot {
            width: 8px; height: 8px; 
            border-radius: 50%; 
            background: #00ff00;
            margin-right: 5px;
            animation: pulse 2s infinite;
        }
        
        @keyframes pulse {
            0% { opacity: 1; }
            50% { opacity: 0.5; }
            100% { opacity: 1; }
        }
        
        /* Sidebar Styles */
        .sidebar {
            position: fixed;
            right: 0;
            top: 0;
            height: 100vh;
            width: 0px;
            background: #1c1c1e;
            border-left: 1px solid #333;
            transition: all 0.4s cubic-bezier(0.25, 0.8, 0.25, 1);
            z-index: 1000;
            overflow-x: hidden;
            overflow-y: auto;
            -webkit-overflow-scrolling: touch;
            /* iOS safe-area padding */
            padding-top: constant(safe-area-inset-top);
            padding-top: env(safe-area-inset-top);
            padding-bottom: constant(safe-area-inset-bottom);
            padding-bottom: env(safe-area-inset-bottom);
        }
        @supports (height: 100dvh) { .sidebar { height: 100dvh; } }
        
        .sidebar.open {
            width: 280px;
        }
        /* Advanced SSH Terminal Panel */
        .terminal-panel {
            position: fixed;
            top: 0;
            right: -50vw;   /* hidden off-screen when closed */
            width: 50vw;     /* covers right half of screen */
            height: 100vh;   /* full length */
            background: #0d0f13;
            border-left: 1px solid #333;
            box-shadow: -8px 0 24px rgba(0,0,0,0.4);
            transition: right 0.35s ease;
            z-index: 3000;
            display: flex;
            flex-direction: column;
            overflow: hidden; /* hide internal scrollbars */
        }
        .terminal-panel.open { right: 0; }
        .terminal-body { flex: 1; overflow: hidden; }
        .terminal-host { width: 100%; height: 100%; display: flex; }
        .terminal-iframe { width: 100%; height: 100%; flex: 1 1 auto; border: none; background: #0d0f13; display:block; }
        .terminal-header {
            display: flex; align-items: center; justify-content: space-between;
            padding: 8px 12px; color: #cfd8dc; background: #151922; border-bottom: 1px solid #222;
        }
        
        
        /* Notepad Panel */
        .notepad-panel {
            position: fixed;
            top: 0;
            right: -35vw;
            width: 35vw;
            height: 100vh;
            background: #0d0f13;
            border-left: 1px solid #1f2937;
            box-shadow: -6px 0 16px rgba(0,0,0,0.4);
            transition: right 0.35s ease;
            z-index: 2800;
            display: flex;
            flex-direction: column;
            overflow: hidden;
        }
        .notepad-panel.open { right: 0; }
        .notepad-header { display:flex; align-items:center; justify-content:space-between; padding:8px 12px; color:#cfd8dc; background:#151922; border-bottom:1px solid #222; }
        .notepad-body { flex:1; display:flex; flex-direction:column; height:100%; }
        .notepad-main { flex: 1 1 auto; display:flex; flex-direction:column; min-height:0; }
        #code-editor { flex:1 1 auto; height:auto; }
        .notepad-ctrl-tab { display:flex; align-items:center; justify-content:space-between; padding:6px 12px; background:#0f1420; border-top:1px solid #1f2937; color:#cfd8dc; cursor:pointer; }
        .ctrl-arrow { transition: transform 0.2s ease; }
        .ctrl-arrow.open { transform: rotate(180deg); }
        .notepad-control { flex: 0 0 0; height:0; overflow:hidden; background:#0b0f17; border-top:1px solid #1f2937; transition: flex-basis 0.25s ease, height 0.25s ease; }
        .notepad-panel.control-open .notepad-control { flex: 0 0 20%; height:auto; }
        .notepad-panel.control-open .notepad-main { flex: 0 0 80%; }
        .control-buttons { display:flex; gap:8px; padding:10px 12px; align-items:center; flex-wrap:wrap; }
        .notepad-textarea { flex:1; width:100%; resize:none; background:#0b0f17; color:#e5e7eb; border:none; padding:12px; font-family:'Fira Code', Consolas, 'Segoe UI', monospace; font-size:14px; line-height:1.45; outline:none; }
        .notepad-actions { display:flex; gap:8px; padding:8px 12px; border-top:1px solid #1f2937; }
        .np-btn { background:#1f2937;color:#e5e7eb;border:1px solid #334155;border-radius:10px;padding:8px;cursor:pointer;display:inline-flex;align-items:center;justify-content:center;gap:8px;box-shadow: 0 6px 20px rgba(0,0,0,0.25), inset 0 1px 0 rgba(255,255,255,0.05); transition: all 0.2s ease; }
        .np-btn:hover { transform: translateY(-1px); box-shadow: 0 10px 24px rgba(0,0,0,0.35), inset 0 1px 0 rgba(255,255,255,0.08); background: linear-gradient(145deg, #243244, #1b2635); }
        .np-btn:active { transform: translateY(1px); box-shadow: 0 4px 12px rgba(0,0,0,0.3), inset 0 2px 8px rgba(0,0,0,0.25); }
        .np-btn:disabled { opacity:0.6; cursor:default; }
        /* Play / Pause Neon Button */
        .play { width:56px; height:56px; min-width:56px; border-radius:14px; display:flex; align-items:center; justify-content:center; background: linear-gradient(180deg, rgba(255,255,255,0.01), rgba(255,255,255,0.015)); border:3px solid #00d4ff; box-shadow: 0 4px 16px rgba(0,212,255,0.05), 0 0 14px rgba(0,212,255,0.06); cursor:pointer; transition: transform .18s ease, box-shadow .18s ease; }
        .play:active{ transform: translateY(1px) scale(.995) }
        .play:hover{ box-shadow: 0 12px 36px rgba(0,212,255,0.09), 0 0 30px rgba(0,212,255,0.08) }
        .play svg { width:24px; height:24px; display:block; fill:#00d4ff; filter: drop-shadow(0 2px 6px rgba(0,212,255,0.12)); }
        /* WPM Arrow Buttons */
        .wpm-col { display:flex; flex-direction:column; gap:8px; align-items:flex-start; margin-left:6px; }
        .wpm-pill { display:flex; align-items:center; justify-content:center; height:36px; padding:0 16px; min-width:180px; border-radius:20px; border:2px solid #00d4ff; background: linear-gradient(180deg, rgba(255,255,255,0.01), rgba(255,255,255,0.015)); box-shadow: 0 4px 14px rgba(0,212,255,0.03); cursor:pointer; font-weight:700; letter-spacing:.2px; transition: transform .12s ease, background .12s; color:#90ee90; }
        .wpm-pill .arrow { margin-right:10px; color:#00d4ff; font-size:16px; }
        .wpm-pill:active { transform:translateY(1px) }
        /* Neon Refresh Button */
        .refresh { width:48px; height:48px; border-radius:50%; display:flex; align-items:center; justify-content:center; border:3px solid #00d4ff; background: linear-gradient(180deg, rgba(255,255,255,0.01), rgba(255,255,255,0.015)); box-shadow: 0 6px 22px rgba(0,212,255,0.05); cursor:pointer; transition: transform .2s ease, box-shadow .2s ease; }
        .refresh:hover { transform: translateY(-3px) rotate(8deg); box-shadow: 0 12px 30px rgba(0,212,255,0.09); }
        .refresh svg { width:24px; height:24px; stroke:#00d4ff; stroke-width:2.2; fill:none; }
        /* Echo & Robo Toggles */
        .toggles-col { display:flex; flex-direction:column; gap:8px; align-items:flex-start; }
        .toggle-pill { height:36px; min-width:180px; padding:0 16px; border-radius:20px; border:2px solid #00d4ff; display:flex; align-items:center; justify-content:space-between; gap:10px; background: linear-gradient(180deg, rgba(255,255,255,0.01), rgba(255,255,255,0.006)); box-shadow: 0 6px 16px rgba(0,212,255,0.02); cursor:pointer; transition: transform .12s, box-shadow .12s; user-select:none; color:#fff; }
        .toggle-pill:active { transform:translateY(1px) }
        .toggle-label { font-weight:600; color:#fff; padding-left:6px; }
        .switch { width:38px; height:20px; background:#22272b; border-radius:999px; position:relative; display:inline-flex; align-items:center; padding:2px; transition: background .18s ease; }
        .knob { width:14px; height:14px; border-radius:50%; background:#fff; transform:translateX(0); transition: transform .18s cubic-bezier(.2,.9,.3,1), background .18s; box-shadow:0 2px 6px rgba(2,6,10,0.6); }
        .switch.on { background: linear-gradient(90deg, rgba(0,212,255,0.18), rgba(0,212,255,0.12)); box-shadow: 0 6px 18px rgba(0,212,255,0.08); }
        .switch.on .knob { transform:translateX(18px); background:#06131a; }

        /* Neon controls removed - reverted to original simple controls */
        .np-icon { width:18px; height:18px; display:inline-block; }
        .notepad-resizer { position:absolute; left:-6px; top:0; width:6px; height:100%; cursor: ew-resize; }
        .CodeMirror { height: 100% !important; background: #0b0f17; color: #e5e7eb; }
        .CodeMirror pre { font-size: 18px; }
        .CodeMirror-cursor { border-left: 2px solid #e5e7eb !important; }
        .cm-s-default .CodeMirror-selected { background: #1f2937 !important; }
        .CodeMirror-gutters { background: #000 !important; border-right: 1px solid #1f2937; }
        .CodeMirror-linenumber { color: #9ca3af; background: #000; }
        .terminal-body { flex: 1; overflow: hidden; }
        .terminal-iframe { width: 100%; height: 100%; border: none; background: #0d0f13; }
        .advanced-btn {
            position: fixed; bottom: 16px; right: 16px; z-index: 2500;
            background: #1e293b; color: #fff; border: 1px solid #334155; border-radius: 6px;
            padding: 8px 12px; cursor: pointer; font-size: 12px;
        }
        .drag-resize {
            position: absolute; left: -6px; top: 0; width: 6px; height: 100%; cursor: ew-resize; background: transparent;
        }
        
        /* Minimal Toggle Trigger */
        .sidebar-trigger {
            position: fixed;
            right: 20px;
            top: 50%;
            transform: translateY(-50%);
            width: 50px;
            height: 50px;
            background: transparent;
            border: none;
            border-radius: 0;
            cursor: move;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: transform 0.3s ease;
            z-index: 3500;
            box-shadow: none;
            user-select: none;
        }
        
        .sidebar-trigger:hover {
            transform: translateY(-50%) scale(1.05);
        }
        
        .sidebar-trigger.dragging {
            cursor: grabbing;
            transform: scale(1.1);
        }
        
        .trigger-icon {
            width: 100%;
            height: 100%;
            pointer-events: none;
            display: flex; align-items: center; justify-content: center;
            background-size: contain;
            background-repeat: no-repeat;
            background-position: center;
            background-image: url('/trigger-icon.png');
        }
        
        .trigger-button {
            position: absolute;
            top: -8px;
            right: -8px;
            width: 20px;
            height: 20px;
            background: #007acc;
            border: none;
            border-radius: 10px;
            color: white;
            font-size: 12px;
            cursor: pointer;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: all 0.2s ease;
            z-index: 1;
        }
        
        .trigger-button:hover {
            background: #005a9e;
            transform: scale(1.1);
        }
        
        .sidebar.open .sidebar-trigger {
            right: 340px;
        }
        
        
        
        /* Sidebar Content */
        .sidebar-content {
            padding: 0;
            opacity: 0;
            transform: translateX(20px);
            transition: all 0.4s cubic-bezier(0.25, 0.8, 0.25, 1);
            height: 100%;
            overflow-x: hidden;
            overflow-y: auto;
        }
        
        .sidebar.open .sidebar-content {
            opacity: 1;
            transform: translateX(0);
            transition-delay: 0.1s;
        }
        
        /* Header */
        .sidebar-header {
            padding: 16px 12px 12px;
            border-bottom: 1px solid #333;
            background: #2a2a2e;
        }
        
        .sidebar-logo {
            width: 100%;
            height: auto;
            display: block;
            border-radius: 8px;
            /* no border */
            -webkit-user-select: none;
            -moz-user-select: none;
            -ms-user-select: none;
            user-select: none;
            -webkit-touch-callout: none;
            -webkit-tap-highlight-color: transparent;
            touch-action: manipulation;
        }
        
        .sidebar-title {
            color: #fff;
            font-size: 14px;
            font-weight: 600;
            margin: 0;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Inter', sans-serif;
        }
        
        .sidebar-subtitle {
            color: rgba(255, 255, 255, 0.6);
            font-size: 10px;
            margin: 2px 0 0 0;
            font-weight: 400;
        }
        
        /* Sections */
        .sidebar-section {
            padding: 0 12px;
            margin: 8px 0;
        }
        
        .section-title {
            color: #e5f2ff;
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin: 6px 0 6px 0;
            font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            display: flex; align-items: center; gap: 8px;
        }
        .section-title::after { content: ''; flex:1; height:1px; background: linear-gradient(90deg, rgba(0,212,255,0.28), rgba(0,212,255,0)); border-radius: 1px; }
        
        /* Buttons */
        .sidebar-btn {
            width: 100%;
            padding: 8px 10px;
            margin-bottom: 6px;
            background: linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01));
            border: 1px solid #2a3344;
            border-radius: 10px;
            color: #e5f2ff;
            font-size: 11px;
            font-weight: 600;
            letter-spacing: .2px;
            cursor: pointer;
            transition: transform .18s ease, box-shadow .18s ease, border-color .18s ease;
            display: flex;
            align-items: center;
            gap: 8px;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Inter', sans-serif;
            box-shadow: 0 2px 10px rgba(0, 212, 255, 0.06);
        }
        
        .btn-group {
            display: flex;
            gap: 4px;
        }
        
        .sidebar-btn:hover {
            border-color: #00d4ff;
            box-shadow: 0 8px 24px rgba(0, 212, 255, 0.12);
            transform: translateY(-1px);
        }
        
        .sidebar-btn:active {
            transform: translateY(0) scale(.995);
        }
        
        .sidebar-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
            transform: none;
        }
        .sidebar-btn.icon-only {
            justify-content: center;
            gap: 0;
            padding: 8px;
        }
        

        .btn-icon { width:16px; height:16px; display:inline-block; }
        .btn-label { font-size: 11px; }
        
        /* Toggle Input Animation */
        .toggle-input-container {
            margin-bottom: 12px;
        }
        
        .toggle-btn {
            width: 100%;
            padding: 8px 10px;
            margin-bottom: 6px;
            background: #1e293b;
            color: #00d4ff;
            border: 1px solid #334155;
            border-radius: 6px;
            cursor: pointer;
            font-size: 11px;
            display: flex;
            align-items: center;
            justify-content: space-between;
            transition: all 0.2s ease;
        }
        
        .toggle-btn:hover {
            border-color: #00d4ff;
            box-shadow: 0 8px 24px rgba(0, 212, 255, 0.12);
            transform: translateY(-1px);
        }
        
        .toggle-btn.active {
            background: #0f172a;
            border-color: #00d4ff;
            color: #00d4ff;
        }
        
        .toggle-switch {
            width: 32px;
            height: 16px;
            background: #334155;
            border-radius: 8px;
            position: relative;
            transition: all 0.3s ease;
        }
        
        .toggle-switch.active {
            background: #00d4ff;
        }
        
        .toggle-switch::after {
            content: '';
            position: absolute;
            width: 12px;
            height: 12px;
            background: #fff;
            border-radius: 50%;
            top: 2px;
            left: 2px;
            transition: all 0.3s ease;
        }
        
        .toggle-switch.active::after {
            transform: translateX(16px);
        }
        
        .animated-input {
            max-height: 0;
            overflow: hidden;
            transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
            transform: scaleY(0);
            transform-origin: top;
            opacity: 0;
        }
        
        .animated-input.show {
            max-height: 160px;
            transform: scaleY(1);
            opacity: 1;
        }
        .keystroke-display {
            width: 100%;
            min-height: 40px;
            padding: 8px 10px;
            background: #1e293b;
            color: #00d4ff;
            border: 1px solid #334155;
            border-radius: 6px;
            font-size: 11px;
            margin-top: 6px;
            transition: all 0.2s ease;
            display: flex;
            flex-wrap: wrap;
            gap: 4px;
            align-items: center;
            font-family: 'Courier New', monospace;
            height: 140px;
            overflow-y: auto;
            -webkit-overflow-scrolling: touch;
            align-content: flex-start;
        }
        
        .keystroke-display.active {
            border-color: #00d4ff;
            box-shadow: 0 0 0 2px rgba(0, 212, 255, 0.2);
        }
        
        .keystroke-item {
            background: #0f172a;
            color: #00d4ff;
            padding: 2px 6px;
            border-radius: 4px;
            border: 1px solid #334155;
            font-size: 10px;
            font-weight: 600;
            display: inline-block;
            margin: 1px;
        }
        
        .keystroke-item.modifier {
            background: #1e40af;
            color: #fff;
            border-color: #3b82f6;
        }
        
        .keystroke-item.regular {
            background: #374151;
            color: #e5e7eb;
        }
        
        .keystroke-placeholder {
            color: #6b7280;
            font-style: italic;
        }
        
        /* Input Controls */
        .input-group {
            margin-bottom: 12px;
        }
        
        .input-label {
            color: rgba(255, 255, 255, 0.7);
            font-size: 9px;
            font-weight: 500;
            margin-bottom: 4px;
            display: block;
        }
        
        .input-control {
            width: 100%;
            padding: 6px 0;
            background: transparent;
            border: none;
            border-bottom: 1px solid rgba(255, 255, 255, 0.2);
            color: white;
            font-size: 10px;
            transition: border-color 0.3s ease;
        }
        
        .input-control:focus {
            outline: none;
            border-bottom-color: #667eea;
        }
        .preset-active { border-color:#00d4ff !important; box-shadow:0 8px 24px rgba(0,212,255,0.12); }
        .audioq-wrap { position: relative; padding: 10px 6px 0; }
        .audioq-range { -webkit-appearance:none; width:100%; height: 10px; background: linear-gradient(90deg, rgba(0,212,255,0.12) 0%, rgba(0,212,255,0.08) 100%); border:1px solid #2a3344; border-radius:8px; outline:none; }
        .audioq-range::-webkit-slider-thumb { -webkit-appearance:none; width: 38px; height: 18px; background:#0b1220; border:2px solid #00d4ff; border-radius:4px; box-shadow: 0 6px 18px rgba(0,212,255,0.16); cursor:pointer; margin-top: -4px; }
        .audioq-range::-moz-range-thumb { width:38px; height:18px; background:#0b1220; border:2px solid #00d4ff; border-radius:4px; box-shadow: 0 6px 18px rgba(0,212,255,0.16); cursor:pointer; }
        .audioq-ticks { display:flex; justify-content:space-between; color:#8aa3b8; font-size:9px; margin-top:6px; padding:0 2px; }
        .futuristic-range { 
            -webkit-appearance:none; width:100%; height: 10px; 
            --val: 50%;
            background: linear-gradient(90deg, #00d4ff 0%, #00d4ff var(--val), rgba(0,212,255,0.08) var(--val), rgba(0,212,255,0.06) 100%);
            border:1px solid #2a3344; border-radius:8px; outline:none; box-shadow: inset 0 0 14px rgba(0,212,255,0.09);
        }
        .futuristic-range::-webkit-slider-runnable-track {
            height:10px; border-radius:8px; border:1px solid #2a3344;
            background: linear-gradient(90deg, #00d4ff 0%, #00d4ff var(--val), rgba(0,212,255,0.08) var(--val), rgba(0,212,255,0.06) 100%);
        }
        .futuristic-range::-moz-range-track {
            height:10px; border-radius:8px; border:1px solid #2a3344;
            background: linear-gradient(90deg, #00d4ff 0%, #00d4ff var(--val), rgba(0,212,255,0.08) var(--val), rgba(0,212,255,0.06) 100%);
        }
        .futuristic-range::-webkit-slider-thumb { -webkit-appearance:none; width: 18px; height: 18px; background:#0b1220; border:2px solid #00d4ff; border-radius:50%; box-shadow: 0 8px 22px rgba(0,212,255,0.18); cursor:pointer; margin-top: -5px; transition: transform .15s ease; }
        .futuristic-range::-moz-range-thumb { width:18px; height:18px; background:#0b1220; border:2px solid #00d4ff; border-radius:50%; box-shadow: 0 8px 22px rgba(0,212,255,0.18); cursor:pointer; transition: transform .15s ease; }
        .futuristic-range:hover::-webkit-slider-thumb, .futuristic-range:hover::-moz-range-thumb { transform: scale(1.08); }
        .futuristic-range:focus { outline: none; box-shadow: 0 0 0 2px rgba(0,212,255,0.25); }
        
        .futuristic-select {
            width: 100%;
            padding: 8px 12px;
            background: linear-gradient(180deg, rgba(0,212,255,0.05), rgba(0,212,255,0.02));
            border: 1px solid #2a3344;
            border-radius: 8px;
            color: #fff;
            font-size: 10px;
            font-weight: 500;
            outline: none;
            cursor: pointer;
            transition: all 0.3s ease;
            box-shadow: inset 0 0 14px rgba(0,212,255,0.09);
        }
        
        .futuristic-select:hover {
            border-color: #00d4ff;
            box-shadow: 0 8px 24px rgba(0, 212, 255, 0.12), inset 0 0 14px rgba(0,212,255,0.15);
        }
        
        .futuristic-select:focus {
            outline: none;
            box-shadow: 0 0 0 2px rgba(0,212,255,0.25), inset 0 0 14px rgba(0,212,255,0.15);
        }
        
        .futuristic-select option {
            background: #0b1220;
            color: #fff;
            padding: 8px;
        }
        
        .capture-status {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 6px 8px;
            background: rgba(0,212,255,0.05);
            border: 1px solid #2a3344;
            border-radius: 6px;
            font-size: 9px;
        }
        
        .status-indicator {
            font-size: 12px;
            color: #ff6b6b;
            animation: pulse 2s infinite;
        }
        
        .status-indicator.working {
            color: #51cf66;
            animation: none;
        }
        
        .status-indicator.warning {
            color: #ffd43b;
            animation: pulse 1s infinite;
        }
        
        .status-text {
            flex: 1;
            color: rgba(255, 255, 255, 0.8);
            font-weight: 500;
        }
        
        .verify-btn {
            background: rgba(0,212,255,0.1);
            border: 1px solid #00d4ff;
            color: #00d4ff;
            border-radius: 4px;
            padding: 2px 6px;
            font-size: 8px;
            cursor: pointer;
            transition: all 0.2s ease;
        }
        
        .verify-btn:hover {
            background: rgba(0,212,255,0.2);
            transform: scale(1.05);
        }
        
        @keyframes pulse {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.5; }
        }
        
        /* URL Display */
        .url-display {
            background: #111;
            border: 1px solid #333;
            border-radius: 5px;
            padding: 6px; /* internal padding */
            margin-bottom: 6px;
            position: relative;
            display: flex;
            align-items: center;
            gap: 8px;
        }
        
        .url-info {
            flex: 1 1 auto;
            display: flex;
            flex-direction: column;
            min-width: 0;
        }
        .url-text {
            color: #00ff88;
            font-size: 10px;
            word-break: break-all;
            line-height: 1.3;
            min-height: 12px;
            font-family: 'JetBrains Mono', 'Fira Code', monospace;
            max-width: 100%;
        }
        .url-port {
            color: #888;
            font-size: 10px;
            margin-top: 3px;
        }
        .url-actions {
            margin-left: auto;
            display: flex;
            flex-direction: column;
            align-items: center;
            gap: 6px;
            border-left: 1px solid #333;
            padding-left: 6px;
            flex: 0 0 10%;
            max-width: 10%;
            justify-content: center;
        }
        .url-action-btn {
            width: 18px;
            height: 18px;
            padding: 0;
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 11px;
            line-height: 1;
            background: linear-gradient(180deg, #1f2937, #0f172a);
            color: #fff;
            border: 1px solid #475569;
            border-radius: 4px;
            cursor: pointer;
            box-shadow: 0 1px 2px rgba(0,0,0,0.25), 0 0 6px rgba(0,212,255,0.08);
            transition: transform .12s ease, box-shadow .12s ease, background .2s ease;
        }
        .url-action-btn:hover {
            transform: translateY(-1px);
            box-shadow: 0 3px 8px rgba(0,0,0,0.28), 0 0 10px rgba(0,212,255,0.12);
        }
        .url-action-btn:active {
            transform: translateY(0);
        }
        /* Alert Top Box */
        #alertSidebar{ position: fixed; top: -30vh; left: 40%; width: 20vw; height: 20vh; background: rgba(17,19,26,.98); color:#fff; box-shadow: 0 20px 40px rgba(0,0,0,.45); backdrop-filter: blur(8px); border:1px solid rgba(255,255,255,.08); transition: top .28s ease; z-index: 10000; display: flex; flex-direction: column; border-radius: 12px; }
        #alertSidebar.open{ top: 0; }
        #alertSidebar .as-header{ display:flex; align-items:center; justify-content:space-between; padding: 10px 12px; border-bottom: 1px solid rgba(255,255,255,.08); }
        #alertSidebar .as-title{ font-weight: 700; letter-spacing: .2px; font-size: 14px; opacity: .9; }
        #alertSidebar .as-close{ background: transparent; color:#bfc7d5; border:0; font-size: 18px; cursor:pointer; }
        #alertSidebar .as-close:hover{ color:#fff; }
        #alertSidebar .as-body{ padding: 12px; font-size: 13px; line-height: 1.5; color:#dfe6f3; display:flex; flex-direction:column; gap:8px; }
        #alertSidebar .as-time{ font-size: 11px; color:#9fb3c8; opacity:.9; }
        
        /* Remote mouse overlay */
        .video-area { position: relative; }
        #mouse-overlay { position:absolute; width:14px; height:14px; border:2px solid #00d4ff; border-radius:50%; box-shadow: 0 0 10px rgba(0,212,255,0.4), inset 0 0 2px rgba(0,212,255,0.9); pointer-events:none; display:none; transform: translate(-50%, -50%); z-index: 5; }
        #mouse-overlay::after { content:''; position:absolute; left:50%; top:50%; width:4px; height:4px; background:#00d4ff; border-radius:50%; transform: translate(-50%, -50%); }
        
        /* Info Items */
        .info-item { display:flex; justify-content:space-between; align-items:center; padding:3px 0; border:0; border-bottom:1px solid rgba(255,255,255,0.06); background:transparent; border-radius:0; margin:0; box-shadow:none; }
        
        .info-item:last-child {
            border-bottom: none;
        }
        
        .info-label { color:#9fb6c8; font-size: 9px; font-weight:600; letter-spacing:.2px; }
        
        .info-value { color:#e5f2ff; font-size: 9px; font-weight:700; }
        
        /* Loading Animation */
        .loading {
            display: inline-block;
            width: 12px;
            height: 12px;
            border: 2px solid rgba(255, 255, 255, 0.3);
            border-radius: 50%;
            border-top-color: #fff;
            animation: spin 1s ease-in-out infinite;
        }
        
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        
        /* Responsive */
        /* 13"–14" laptops (e.g., 1366x768, 1440x900) */
        @media (min-width: 1024px) and (max-width: 1440px) {
            .sidebar.open { width: 300px; }
            .terminal-panel { width: 55vw; right: -55vw; }
            .notepad-panel { width: 42vw; right: -42vw; }
            .play { width:52px; height:52px; }
            .refresh { width:46px; height:46px; }
            .wpm-pill, .toggle-pill { min-width:160px; height:34px; padding:0 14px; }
            .section-title { font-size: 11px; }
            .url-text { font-size: 11px; }
            .url-port { font-size: 10px; }
            .CodeMirror pre { font-size: 16px; }
            .sidebar.open .sidebar-trigger { right: calc(300px + 20px); }
        }
        /* 14"–15" laptops inc. MacBook 14 (1280–1535 logical widths) */
        @media (min-width: 1280px) and (max-width: 1535px) {
            /* Sidebar width and trigger offset */
            .sidebar.open { width: 320px; }
            .sidebar.open .sidebar-trigger { right: calc(320px + 20px); }
            /* Panels */
            .terminal-panel { width: 54vw; right: -54vw; }
            .notepad-panel { width: 41vw; right: -41vw; }
            /* Codex control sizing */
            .notepad-control .control-buttons { gap: 6px; }
            .play { width:50px; height:50px; }
            .refresh { width:44px; height:44px; }
            .wpm-pill, .toggle-pill { min-width:150px; height:32px; padding:0 12px; }
            .wpm-pill .wpm-text { font-size: 12px; }
            .toggle-label { font-size: 12px; }
            .CodeMirror pre { font-size: 15px; }
        }
        @media (max-width: 1024px) {
            .sidebar.open { width: 260px; }
            .terminal-panel { width: 60vw; right: -60vw; }
            .notepad-panel { width: 45vw; right: -45vw; }
        }
        @media (max-width: 768px) {
            .sidebar.open { width: 240px; }
            .terminal-panel { width: 90vw; right: -90vw; }
            .terminal-panel.open { right: 0; }
            .notepad-panel { width: 85vw; right: -85vw; }
            .notepad-panel.open { right: 0; }
            .sidebar-trigger { right: 12px; width: 44px; height: 44px; }
            .wpm-pill, .toggle-pill { min-width: 140px; height: 32px; padding: 0 12px; }
            .play { width:48px; height:48px; }
            .refresh { width:42px; height:42px; }
            .url-actions { gap: 4px; }

            .url-action-btn { width: 16px; height: 16px; }
            .section-title { font-size: 10px; }
            .info-label, .info-value { font-size: 8px; }
        }
        @media (max-width: 480px) {
            .sidebar.open { width: 100vw; }
            .trigger-icon { transform: scale(0.9); }
            .wpm-pill, .toggle-pill { min-width: 120px; height: 30px; padding: 0 10px; }
            .url-text { font-size: 9px; }
            .url-port { font-size: 9px; }
            #auth-form { top: 76%; left: 20%; width: 60%; height: 18%; }
            #auth-code { font-size: 16px; }
            .terminal-panel { width: 100vw; right: -100vw; }
            .notepad-panel { width: 100vw; right: -100vw; }
        }
        /* Phones (<= 576px) */
        @media (max-width: 576px) {
            .sidebar.open { width: 100vw; }
            .sidebar.open .sidebar-trigger { right: 12px; }
            .terminal-panel, .notepad-panel { width: 100vw; right: -100vw; }
        }
        /* iPad portrait */
        @media (min-width: 768px) and (max-width: 1024px) and (orientation: portrait) {
            .sidebar.open { width: 280px; }
            .terminal-panel { width: 85vw; right: -85vw; }
            .notepad-panel { width: 65vw; right: -65vw; }
        }
        /* Large laptop / desktop (16-inch, >= 1536px width) */
        @media (min-width: 1536px) {
            .sidebar.open { width: 340px; }
            .sidebar.open .sidebar-trigger { right: calc(340px + 20px); }
            .terminal-panel { width: 50vw; right: -50vw; }
            .notepad-panel { width: 40vw; right: -40vw; }
        }
        /* 1280–1535px (common 15" and many 14") */
        @media (min-width: 1280px) and (max-width: 1535px) {
            .sidebar.open { width: 320px; }
            .sidebar.open .sidebar-trigger { right: calc(320px + 20px); }
            .terminal-panel { width: 55vw; right: -55vw; }
            .notepad-panel { width: 42vw; right: -42vw; }
        }
        /* Retina tweaks (Macs) */
        @media (-webkit-min-device-pixel-ratio: 2), (min-resolution: 192dpi) {
            .play svg, .refresh svg { filter: drop-shadow(0 1px 3px rgba(0,212,255,0.12)); }
        }
        
    </style>
    <style id="kb-style">
        /* Keyboard Popup Overlay */
        .kb-overlay{position:fixed;inset:0;background:rgba(0,0,0,0.8);display:none;align-items:center;justify-content:center;z-index:10000;overflow:auto;padding:10px}
        .keyboard-popup{background:linear-gradient(145deg,#2c3e50,#34495e);padding:20px;border-radius:15px;box-shadow:0 20px 60px rgba(0,0,0,0.5);border:3px solid #1a252f;position:relative;max-width:95vw;max-height:95vh}
        .keyboard-popup::before{content:'';position:absolute;top:-2px;left:-2px;right:-2px;bottom:-2px;background:linear-gradient(45deg,#667eea,#764ba2,#667eea);border-radius:17px;z-index:-1}
        .keyboard{display:grid;gap:3px;grid-template-columns:repeat(20,1fr);grid-template-rows:repeat(6,1fr);width:100%;max-width:1000px;margin:0 auto}
        .keyboard-popup .key{background:linear-gradient(145deg,#27ae60,#2ecc71);border:2px solid #229954;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:600;color:#fff;cursor:pointer;transition:all .2s cubic-bezier(.175,.885,.32,1.275);box-shadow:0 4px 8px rgba(0,0,0,.2),inset 0 2px 4px rgba(255,255,255,.1);min-height:40px;text-align:center;line-height:1.2;user-select:none;position:relative;text-shadow:1px 1px 2px rgba(0,0,0,.3)}
        .keyboard-popup .key::before{content:'';position:absolute;top:2px;left:2px;right:2px;height:50%;background:linear-gradient(180deg,rgba(255,255,255,.2),transparent);border-radius:4px 4px 0 0;pointer-events:none}
        .keyboard-popup .key:hover{transform:translateY(-1px);box-shadow:0 6px 12px rgba(0,0,0,.3),inset 0 2px 4px rgba(255,255,255,.2);background:linear-gradient(145deg,#2ecc71,#27ae60)}
        .keyboard-popup .key:active{transform:translateY(1px);box-shadow:0 2px 4px rgba(0,0,0,.3),inset 0 2px 8px rgba(0,0,0,.2)}
        .keyboard-popup .key.pressed{background:linear-gradient(145deg,#e74c3c,#c0392b);border-color:#a93226;transform:translateY(2px) scale(.98);box-shadow:0 2px 4px rgba(0,0,0,.3),inset 0 2px 8px rgba(0,0,0,.3)}
        .keyboard-popup .key.pressed::before{background:linear-gradient(180deg,rgba(255,255,255,.1),transparent)}
        @keyframes kbPress{0%{transform:translateY(-1px) scale(1);box-shadow:0 6px 12px rgba(0,0,0,.3)}50%{transform:translateY(3px) scale(.95);box-shadow:0 1px 2px rgba(0,0,0,.4)}100%{transform:translateY(2px) scale(.98);box-shadow:0 2px 4px rgba(0,0,0,.3)}}
        .keyboard-popup .key.small-text{font-size:8px}
        .two-line{display:flex;flex-direction:column;align-items:center;justify-content:center;line-height:1}
        .two-line span{display:block;margin:1px 0}
        .kb-close-btn{position:absolute;top:-10px;right:-10px;width:30px;height:30px;background:linear-gradient(145deg,#e74c3c,#c0392b);border:2px solid #a93226;border-radius:50%;color:#fff;font-weight:bold;font-size:16px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all .2s ease;z-index:10}
        .kb-close-btn:hover{transform:scale(1.1);background:linear-gradient(145deg,#c0392b,#e74c3c)}
        @media (max-width:1200px){.keyboard{grid-template-columns:repeat(20,minmax(25px,1fr))}.keyboard-popup .key{min-height:35px;font-size:10px}.keyboard-popup{padding:15px;margin:10px}}
        @media (max-width:768px){.keyboard{grid-template-columns:repeat(20,minmax(20px,1fr));gap:2px}.keyboard-popup .key{min-height:30px;font-size:8px}.keyboard-popup{padding:10px;margin:5px}}
        /* Grid positions */
        .keyboard .key-esc{grid-column:1;grid-row:1}
        .keyboard .key-f1{grid-column:3;grid-row:1}
        .keyboard .key-f2{grid-column:4;grid-row:1}
        .keyboard .key-f3{grid-column:5;grid-row:1}
        .keyboard .key-f4{grid-column:6;grid-row:1}
        .keyboard .key-f5{grid-column:7;grid-row:1}
        .keyboard .key-f6{grid-column:8;grid-row:1}
        .keyboard .key-f7{grid-column:9;grid-row:1}
        .keyboard .key-f8{grid-column:10;grid-row:1}
        .keyboard .key-f9{grid-column:11;grid-row:1}
        .keyboard .key-f10{grid-column:12;grid-row:1}
        .keyboard .key-f11{grid-column:13;grid-row:1}
        .keyboard .key-f12{grid-column:14;grid-row:1}
        .keyboard .key-fn{grid-column:16;grid-row:1}
        .keyboard .key-home{grid-column:17;grid-row:1}
        .keyboard .key-pgup{grid-column:18;grid-row:1}
        .keyboard .key-eq{grid-column:19;grid-row:1}
        .keyboard .key-div{grid-column:20;grid-row:1}
        .keyboard .key-backtick{grid-column:1;grid-row:2}
        .keyboard .key-1{grid-column:2;grid-row:2}
        .keyboard .key-2{grid-column:3;grid-row:2}
        .keyboard .key-3{grid-column:4;grid-row:2}
        .keyboard .key-4{grid-column:5;grid-row:2}
        .keyboard .key-5{grid-column:6;grid-row:2}
        .keyboard .key-6{grid-column:7;grid-row:2}
        .keyboard .key-7{grid-column:8;grid-row:2}
        .keyboard .key-8{grid-column:9;grid-row:2}
        .keyboard .key-9{grid-column:10;grid-row:2}
        .keyboard .key-0{grid-column:11;grid-row:2}
        .keyboard .key-minus{grid-column:12;grid-row:2}
        .keyboard .key-equals{grid-column:13;grid-row:2}
        .keyboard .key-delete{grid-column:14 / 16;grid-row:2}
        .keyboard .key-fn2{grid-column:16;grid-row:2}
        .keyboard .key-home2{grid-column:17;grid-row:2}
        .keyboard .key-pgup2{grid-column:18;grid-row:2}
        .keyboard .key-num7{grid-column:19;grid-row:2}
        .keyboard .key-num8{grid-column:20;grid-row:2}
        .keyboard .key-tab{grid-column:1 / 3;grid-row:3}
        .keyboard .key-q{grid-column:3;grid-row:3}
        .keyboard .key-w{grid-column:4;grid-row:3}
        .keyboard .key-e{grid-column:5;grid-row:3}
        .keyboard .key-r{grid-column:6;grid-row:3}
        .keyboard .key-t{grid-column:7;grid-row:3}
        .keyboard .key-y{grid-column:8;grid-row:3}
        .keyboard .key-u{grid-column:9;grid-row:3}
        .keyboard .key-i{grid-column:10;grid-row:3}
        .keyboard .key-o{grid-column:11;grid-row:3}
        .keyboard .key-p{grid-column:12;grid-row:3}
        .keyboard .key-bracket-left{grid-column:13;grid-row:3}
        .keyboard .key-bracket-right{grid-column:14;grid-row:3}
        .keyboard .key-backslash{grid-column:15;grid-row:3}
        .keyboard .key-delete2{grid-column:16;grid-row:3}
        .keyboard .key-end{grid-column:17;grid-row:3}
        .keyboard .key-pgdn{grid-column:18;grid-row:3}
        .keyboard .key-num9{grid-column:19;grid-row:3}
        .keyboard .key-numminus{grid-column:20;grid-row:3}
        .keyboard .key-caps{grid-column:1 / 3;grid-row:4}
        .keyboard .key-a{grid-column:3;grid-row:4}
        .keyboard .key-s{grid-column:4;grid-row:4}
        .keyboard .key-d{grid-column:5;grid-row:4}
        .keyboard .key-f{grid-column:6;grid-row:4}
        .keyboard .key-g{grid-column:7;grid-row:4}
        .keyboard .key-h{grid-column:8;grid-row:4}
        .keyboard .key-j{grid-column:9;grid-row:4}
        .keyboard .key-k{grid-column:10;grid-row:4}
        .keyboard .key-l{grid-column:11;grid-row:4}
        .keyboard .key-semicolon{grid-column:12;grid-row:4}
        .keyboard .key-quote{grid-column:13;grid-row:4}
        .keyboard .key-return{grid-column:14 / 16;grid-row:4}
        .keyboard .key-num4{grid-column:17;grid-row:4}
        .keyboard .key-num5{grid-column:18;grid-row:4}
        .keyboard .key-num6{grid-column:19;grid-row:4}
        .keyboard .key-plus{grid-column:20;grid-row:4}
        .keyboard .key-shift-left{grid-column:1 / 3;grid-row:5}
        .keyboard .key-z{grid-column:3;grid-row:5}
        .keyboard .key-x{grid-column:4;grid-row:5}
        .keyboard .key-c{grid-column:5;grid-row:5}
        .keyboard .key-v{grid-column:6;grid-row:5}
        .keyboard .key-b{grid-column:7;grid-row:5}
        .keyboard .key-n{grid-column:8;grid-row:5}
        .keyboard .key-m{grid-column:9;grid-row:5}
        .keyboard .key-comma{grid-column:10;grid-row:5}
        .keyboard .key-period{grid-column:11;grid-row:5}
        .keyboard .key-slash{grid-column:12;grid-row:5}
        .keyboard .key-shift-right{grid-column:13 / 15;grid-row:5}
        .keyboard .key-up{grid-column:16;grid-row:5}
        .keyboard .key-num1{grid-column:17;grid-row:5}
        .keyboard .key-num2{grid-column:18;grid-row:5}
        .keyboard .key-num3{grid-column:19;grid-row:5}
        .keyboard .key-enter{grid-column:20;grid-row:4 / 6}
        .keyboard .key-ctrl-left{grid-column:1;grid-row:6}
        .keyboard .key-alt-left{grid-column:2;grid-row:6}
        .keyboard .key-space{grid-column:3 / 13;grid-row:6}
        .keyboard .key-alt-right{grid-column:13;grid-row:6}
        .keyboard .key-ctrl-right{grid-column:14;grid-row:6}
        .keyboard .key-left{grid-column:15;grid-row:6}
        .keyboard .key-down{grid-column:16;grid-row:6}
        .keyboard .key-right{grid-column:17;grid-row:6}
        .keyboard .key-num0{grid-column:18 / 20;grid-row:6}
        .keyboard .key-dot{grid-column:20;grid-row:6}
    </style>
    <style id="splash-style">
        /* Splash overlay */
        body { overflow: hidden; }
        #splash-screen { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: transparent; display: flex; justify-content: center; align-items: center; z-index: 9999; transition: opacity 0.5s ease-out; }
        .splash-container { position: relative; max-width: 1000px; width: 90%; }
        .splash-container img { width: 100%; height: auto; display: block; }
        #auth-form { position: absolute; top: 76%; left: 23%; width: 54%; height: 17%; }
        #auth-code { width: 100%; height: 100%; box-sizing: border-box; border: none; background-color: transparent; text-align: center; font-size: 2.5vw; font-family: 'Courier New', Courier, monospace; font-weight: bold; color: #4a2c0f; outline: none; text-transform: uppercase; }
        @media (min-width: 1000px) { #auth-code { font-size: 40px; } }
        .hidden { display: none; }
        .fade-out { opacity: 0; }
        .shake { animation: shake 0.5s; }
        @keyframes shake { 10%, 90% { transform: translateX(-5px); } 20%, 80% { transform: translateX(5px); } 30%, 50%, 70% { transform: translateX(-5px); } 40%, 60% { transform: translateX(5px); } }


        /* YOU WILL SEE THIS CODE ABOVE */
        .shake { animation: shake 0.5s; }
        @keyframes shake { 10%, 90% { transform: translateX(-5px); } 20%, 80% { transform: translateX(5px); } 30%, 50%, 70% { transform: translateX(-5px); } 40%, 60% { transform: translateX(5px); } }

        /* --- VVV PASTE THE NEW CSS CODE HERE VVV --- */

        /* OCR Selection Box */
        #ocr-selection-box {
            position: fixed;
            border: 2px dashed #00d4ff;
            background-color: rgba(0, 212, 255, 0.2);
            box-shadow: 0 0 10px rgba(0, 212, 255, 0.5);
            z-index: 9998;
            pointer-events: none;
            display: none;
        }


        /* Add this rule for the blurred background layer */
        .blurred-background {
            position: fixed;
            top: 0;
            left: 0;
            width: 100vw;
            height: 100vh;
            z-index: 9998; /* below #splash-screen (9999), above page content */
            background-image: url('https://uploads.onecompiler.io/43vbtkw3d/43vbtm4hc/zadoo.png');
            background-size: cover;
            background-position: center;
            filter: blur(15px);
            transform: scale(1.1);
            pointer-events: none;
            opacity: 0; /* hidden unless splash-active */
            transition: opacity 0.4s ease;
        }
        body.splash-active .blurred-background { opacity: 1; }
        /* Blur the remote display while splash is active */
        body.splash-active #screen {
            filter: blur(18px);
            transform: scale(1.02);
            transition: filter 0.25s ease, transform 0.25s ease;
        }
    </style>
</head>
<body>
    <div class="blurred-background"></div>
    <div id="splash-screen">
        <div class="splash-container">
            <img src="/splash.png" alt="Authentication">
            <form id="auth-form">
                <input type="text" id="auth-code" placeholder="Enter Code" autocomplete="off" autofocus maxlength="10">
            </form>
        </div>
    </div>
    
    <!-- Keyboard Configure Popup Overlay -->
    <div id="kb-overlay" class="kb-overlay" style="display:none;">
      <div class="keyboard-popup">
        <div class="kb-close-btn" onclick="closeKeyboardPopup()">&times;</div>
        <div class="keyboard">
            <div class="key key-esc">esc</div>
            <div class="key key-f1">F1</div>
            <div class="key key-f2">F2</div>
            <div class="key key-f3">F3</div>
            <div class="key key-f4">F4</div>
            <div class="key key-f5">F5</div>
            <div class="key key-f6">F6</div>
            <div class="key key-f7">F7</div>
            <div class="key key-f8">F8</div>
            <div class="key key-f9">F9</div>
            <div class="key key-f10">F10</div>
            <div class="key key-f11">F11</div>
            <div class="key key-f12">F12</div>
            <div class="key key-fn small-text">fn</div>
            <div class="key key-home small-text">home</div>
            <div class="key key-pgup small-text">page<br>up</div>
            <div class="key key-eq">=</div>
            <div class="key key-div">/</div>

            <div class="key key-backtick two-line"><span>~</span><span>`</span></div>
            <div class="key key-1 two-line"><span>!</span><span>1</span></div>
            <div class="key key-2 two-line"><span>@</span><span>2</span></div>
            <div class="key key-3 two-line"><span>#</span><span>3</span></div>
            <div class="key key-4 two-line"><span>$</span><span>4</span></div>
            <div class="key key-5 two-line"><span>%</span><span>5</span></div>
            <div class="key key-6 two-line"><span>^</span><span>6</span></div>
            <div class="key key-7 two-line"><span>&</span><span>7</span></div>
            <div class="key key-8 two-line"><span>*</span><span>8</span></div>
            <div class="key key-9 two-line"><span>(</span><span>9</span></div>
            <div class="key key-0 two-line"><span>)</span><span>0</span></div>
            <div class="key key-minus two-line"><span>_</span><span>-</span></div>
            <div class="key key-equals two-line"><span>+</span><span>=</span></div>
            <div class="key key-delete">delete</div>
            <div class="key key-fn2 small-text">fn</div>
            <div class="key key-home2 small-text">home</div>
            <div class="key key-pgup2 small-text">page<br>up</div>
            <div class="key key-num7">7</div>
            <div class="key key-num8">8</div>

            <div class="key key-tab">tab</div>
            <div class="key key-q">Q</div>
            <div class="key key-w">W</div>
            <div class="key key-e">E</div>
            <div class="key key-r">R</div>
            <div class="key key-t">T</div>
            <div class="key key-y">Y</div>
            <div class="key key-u">U</div>
            <div class="key key-i">I</div>
            <div class="key key-o">O</div>
            <div class="key key-p">P</div>
            <div class="key key-bracket-left two-line"><span>{</span><span>[</span></div>
            <div class="key key-bracket-right two-line"><span>}</span><span>]</span></div>
            <div class="key key-backslash two-line"><span>|</span><span>\\</span></div>
            <div class="key key-delete2 small-text">delete</div>
            <div class="key key-end small-text">end</div>
            <div class="key key-pgdn small-text">page<br>down</div>
            <div class="key key-num9">9</div>
            <div class="key key-numminus">-</div>

            <div class="key key-caps">caps lock</div>
            <div class="key key-a">A</div>
            <div class="key key-s">S</div>
            <div class="key key-d">D</div>
            <div class="key key-f">F</div>
            <div class="key key-g">G</div>
            <div class="key key-h">H</div>
            <div class="key key-j">J</div>
            <div class="key key-k">K</div>
            <div class="key key-l">L</div>
            <div class="key key-semicolon two-line"><span>:</span><span>;</span></div>
            <div class="key key-quote two-line"><span>"</span><span>'</span></div>
            <div class="key key-return">return</div>
            <div class="key key-num4">4</div>
            <div class="key key-num5">5</div>
            <div class="key key-num6">6</div>
            <div class="key key-plus">+</div>

            <div class="key key-shift-left">shift</div>
            <div class="key key-z">Z</div>
            <div class="key key-x">X</div>
            <div class="key key-c">C</div>
            <div class="key key-v">V</div>
            <div class="key key-b">B</div>
            <div class="key key-n">N</div>
            <div class="key key-m">M</div>
            <div class="key key-comma two-line"><span>&lt;</span><span>,</span></div>
            <div class="key key-period two-line"><span>&gt;</span><span>.</span></div>
            <div class="key key-slash two-line"><span>?</span><span>/</span></div>
            <div class="key key-shift-right">shift</div>
            <div class="key key-up">▲</div>
            <div class="key key-num1">1</div>
            <div class="key key-num2">2</div>
            <div class="key key-num3">3</div>
            <div class="key key-enter">enter</div>

            <div class="key key-ctrl-left">ctrl</div>
            <div class="key key-alt-left">alt</div>
            <div class="key key-space"></div>
            <div class="key key-alt-right">alt</div>
            <div class="key key-ctrl-right">ctrl</div>
            <div class="key key-left">◄</div>
            <div class="key key-down">▼</div>
            <div class="key key-right">►</div>
            <div class="key key-num0">0</div>
            <div class="key key-dot">.</div>
        </div>
      </div>
    </div>
    <div class="container">
        <!-- Video Area -->
        <div class="video-area">
            <div id="mouse-overlay"></div>
            <canvas id="screen"></canvas>
            <div id="snap-overlay"></div>
            <div id="snap-selection-box"></div>
            <!-- Hidden status element for JavaScript -->
            <div id="status" style="display: none;">Connecting...</div>
            <div id="current-quality" style="display: none;">75</div>
            <div id="current-fps" style="display: none;">30</div>
            <div id="uptime-display" style="display: none;">00:00:00</div>
        </div>
        
        <!-- Floating Sidebar Trigger -->
        <div class="sidebar-trigger" id="sidebar-trigger" title="Open Controls">
            <div class="trigger-icon"></div>
        </div>
        
        <!-- Collapsible Sidebar -->
        <div class="sidebar" id="sidebar">
            <!-- Sidebar Content -->
            <div class="sidebar-content">
                <!-- Header -->
                <div class="sidebar-header">
                    <img class="sidebar-logo" src="/brand-header.png" alt="Zadoo Remote Control Panel"/>
                </div>
                
                <div id="sidebar-main" class="sidebar-main">
                <!-- Public URL Row -->
                <div class="sidebar-section">
                    <div class="url-display">
                        <div class="url-info">
                        <div class="url-text" id="link-text">Getting public URL...</div>
                            <div class="url-port" id="port-info"></div>
                    </div>
                        <div class="url-actions">
                            <button class="url-action-btn" title="Copy" onclick="copyPublicLink(event)">⧉</button>
                            <button class="url-action-btn" id="btn-refresh-url" title="Switch & Get New Link" onclick="refreshPublicLink()">⟳</button>
                        </div>
                    </div>
                </div>
                
                <!-- Graphics Quality Section -->
                <div class="sidebar-section">
                    <div class="section-title">🎨 Graphics</div>
                    <div class="input-group">
                        <label class="input-label">Visual Quality <span id="quality-value" style="color:#00d4ff; font-weight:700;">75</span>%</label>
                        <input type="range" class="futuristic-range" id="quality-slider" 
                               min="10" max="100" value="75" oninput="updateQuality(this.value)">
                    </div>
                    <div class="input-group">
                        <label class="input-label">Frame Rate <span id="fps-value" style="color:#00d4ff; font-weight:700;">30</span></label>
                        <input type="range" class="futuristic-range" id="fps-slider" 
                               min="5" max="120" value="30" oninput="updateFPS(this.value)">
                    </div>
                    <div class="input-group">
                        <label class="input-label">Capture Method</label>
                        <select class="futuristic-select" id="capture-method-select" onchange="changeCaptureMethod(this.value)">
                            <option value="auto">Auto (Best Performance)</option>
                            <option value="dxcam">DXCam (100+ FPS)</option>
                            <option value="fast_ctypes">Fast CTypes (70-125 FPS)</option>
                            <option value="mss">MSS (30-50 FPS)</option>
                            <option value="bettercam">BetterCam (High FPS)</option>
                            <option value="winrt">WinRT (GraphicsCapture)</option>
                        </select>
                    </div>
                    <div class="input-group">
                        <label class="input-label">Performance Mode</label>
                        <div style="display:flex;gap:6px;align-items:center;">
                            <button class="sidebar-btn" id="perf-toggle" onclick="togglePerformanceMode()" title="Enable performance optimizations">Enable</button>
                            <select class="futuristic-select" id="perf-region" title="Region">
                                <option value="full">Full</option>
                                <option value="center_0.75">Center 75%</option>
                                <option value="center_0.5">Center 50%</option>
                                <option value="custom">Custom…</option>
                            </select>
                            <select class="futuristic-select" id="perf-scale" title="Downscale">
                                <option value="1">1x</option>
                                <option value="2">0.5x</option>
                                <option value="3">~0.33x</option>
                            </select>
                            <button class="sidebar-btn" id="perf-gray" onclick="togglePerfGray()" title="Grayscale (B/W)" style="margin-left:auto;">B/W</button>
                        </div>
                    </div>
                    <div class="input-group">
                        <div class="capture-status" id="capture-status">
                            <div class="status-indicator" id="status-indicator">●</div>
                            <span class="status-text" id="status-text">Checking...</span>
                            <button class="verify-btn" onclick="verifyCaptureMethod()" title="Verify Capture Method">✓</button>
                        </div>
                    </div>
                </div>
                
                <!-- Server Info Section -->
                <div class="sidebar-section">
                    <div class="section-title">📊 Live Status</div>
                    <div class="info-item">
                        <span class="info-label">Connection</span>
                        <span class="info-value" id="server-status">Connecting</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">Uptime</span>
                        <span class="info-value" id="server-uptime">00:00:00</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">Live FPS</span>
                        <span class="info-value"><span id="sidebar-fps">0</span></span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">Audio</span>
                        <span class="info-value" id="audio-status">Idle</span>
                    </div>
                    <div class="info-item">
                        <span class="info-label">Email</span>
                        <span class="info-value" id="email-status">Pending</span>
                    </div>
                    
                </div>
                
                <!-- Clipboard Section -->
                <div class="sidebar-section">
                    <div class="section-title">📋 Clipboard</div>
                    <div class="btn-group">
                        <button class="sidebar-btn" onclick="getRemoteClipboard()" title="Pull from remote">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                            <span>Pull</span>
                        </button>
                        <button class="sidebar-btn" onclick="setRemoteClipboard()" title="Push to remote">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 14 12 9 7 14"/><line x1="12" y1="9" x2="12" y2="21"/></svg>
                            <span>Push</span>
                        </button>
                    </div>
                </div>
                
                <!-- Quick Actions Section -->
                <div class="sidebar-section">
                    <div class="section-title">⚡ Quick Actions</div>
                    <div class="btn-group">
                        <button class="sidebar-btn" id="btn-snap" onclick="startSnapSelection(event)" oncontextmenu="takeSnapshotFull(event)" title="Snapshot (Left: select area • Right: full screen)">
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h3l2-3h8l2 3h3a2 2 0 0 1 2 2z"/><circle cx="12" cy="13" r="4"/></svg>
                            <span>Snap</span>
                        </button>
                        <button class="sidebar-btn" id="btn-ocr" onclick="performOCR(event)" oncontextmenu="startOCRSelection(event); return false;" title="OCR (Right-click to select area)">
                        
                            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 4h5v5H4zM15 4h5v5h-5zM4 15h5v5H4zM15 15h5v5h-5z"/></svg>
                            <span>OCR</span>
                        </button>
                    </div>
                    <button class="sidebar-btn" id="mouse-toggle-btn" onclick="window.toggleMouseOverlay()" title="Show Mouse">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="6" y="2" width="12" height="20" rx="6"/><line x1="12" y1="6" x2="12" y2="10"/></svg>
                        <span id="mouse-toggle-text">Show Mouse</span>
                        </button>
                    <button class="sidebar-btn" onclick="toggleFullscreen()" title="Fullscreen">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M8 3H5a2 2 0 0 0-2 2v3m0 8v3a2 2 0 0 0 2 2h3m8-18h3a2 2 0 0 1 2 2v3m0 8v3a2 2 0 0 1-2 2h-3"/></svg>
                        <span>Fullscreen</span>
                        </button>
                    

                    
                </div>
            </div>
                
                <div id="sidebar-advanced" class="sidebar-advanced" style="display: none;">
                    <div class="sidebar-section">
                        <div class="section-title">🛠️ Advanced Tab</div>
                        
                        <!-- Toggle Input Container -->
                        <div class="toggle-input-container">
                            <button class="toggle-btn" id="toggle-input-btn" onclick="toggleAnimatedInput()">
                                <span style="display:inline-flex;align-items:center;gap:6px;">
                                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 12h2"></path><path d="M8 12h2"></path><path d="M12 12h2"></path><path d="M16 12h2"></path><circle cx="12" cy="12" r="9"></circle></svg>
                                    <span>Sniff</span>
                                </span>
                                <div class="toggle-switch" id="toggle-switch"></div>
                            </button>
                            <div class="animated-input" id="animated-input">
                                <div class="keystroke-display" id="keystroke-display">
                                    <span class="keystroke-placeholder">Press keys on host to capture keystrokes...</span>
                                </div>
                            </div>
                        </div>
                        
                        <div class="btn-group">
                            <button class="sidebar-btn" id="btn-terminal" onclick="toggleTerminal()" title="Terminal">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="4 17 10 11 4 5"/><line x1="12" y1="19" x2="20" y2="19"/></svg>
                                <span class="btn-label">Terminal</span>
                            </button>
                        </div>
                        <div class="btn-group">
                            <button class="sidebar-btn" id="btn-webcam" onclick="toggleWebcam()" title="Webcam">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="3" y="6" width="13" height="12" rx="2"/><polygon points="16 8 22 12 16 16 16 8"/></svg>
                                <span class="btn-label">Webcam</span>
                            </button>
                        </div>
                        <div id="webcam-container" style="position: relative; margin-top: 8px; border: 1px solid #334155; border-radius: 6px; overflow: hidden; background: #0b0f17; display:none;">
                            <img id="webcam-img" alt="webcam" style="width: 100%; height: 180px; object-fit: cover; display: block;"/>
                            <button title="Fullscreen" onclick="toggleWebcamFullscreen()" style="position: absolute; right: 6px; top: 6px; z-index: 5; background: #334155; color: #fff; border: 1px solid #475569; border-radius: 4px; padding: 2px 6px; cursor: pointer;">⛶</button>
                        </div>
                        <div class="input-group" style="margin-top:6px;">
                            <div style="display:flex; gap:6px;">
                                <select id="camera-select" class="input-control" style="flex:1;">
                                    <option>Loading...</option>
                                </select>
                                <button class="sidebar-btn" style="padding:6px 8px;" onclick="refreshCameras()" title="Refresh Cameras">⟳</button>
                            </div>
                        </div>
                        <div class="btn-group">
                            <button class="sidebar-btn" id="audio-toggle" onclick="toggleAudio(event)" title="Audio">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polygon points="11 5 6 9 2 9 2 15 6 15 11 19 11 5"/><path d="M19 12a4 4 0 0 0-4-4"/><path d="M19 12a4 4 0 0 1-4 4"/></svg>
                                <span class="btn-label">Audio</span>
                            </button>
                            <button class="sidebar-btn" id="mic-toggle" onclick="toggleMic(event)" title="Microphone" style="margin-left:6px;">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                                  <path d="M12 1a3 3 0 0 0-3 3v6a3 3 0 0 0 6 0V4a3 3 0 0 0-3-3z"/>
                                  <path d="M19 10a7 7 0 0 1-14 0"/><line x1="12" y1="17" x2="12" y2="23"/><line x1="8" y1="23" x2="16" y2="23"/>
                                </svg>
                                <span class="btn-label">Mic</span>
                            </button>
                        </div>
                        <div class="input-group">
                            <label class="input-label" style="display:flex; align-items:center; gap:8px;">
                                <span style="font-weight:700; letter-spacing:.4px; color:#9fb6c8;">Audio Quality</span>
                                <span id="audioq-value" style="color:#00d4ff; font-weight:700;">Balanced</span>
                            </label>
                            <div class="audioq-wrap">
                                <input type="range" class="audioq-range" id="audioq-steps" min="0" max="3" step="1" value="1" oninput="handleAudioQSlider(this.value)">
                                <div class="audioq-ticks"><span>Low</span><span>Balanced</span><span>High</span><span>Ultra</span></div>
                            </div>
                        </div>
                        <div class="btn-group" style="display:flex; gap:6px;">
                            <button class="sidebar-btn icon-only" id="btn-disable-mouse" onclick="toggleMouseBlock()" title="Disable Mouse">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="6" y="2" width="12" height="20" rx="6"/><line x1="12" y1="6" x2="12" y2="10"/><line x1="4" y1="4" x2="20" y2="20"/></svg>
                            </button>
                            <button class="sidebar-btn icon-only" id="btn-disable-keyboard" onclick="toggleKeyboardBlock()" title="Disable Keyboard">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="6" width="20" height="12" rx="2"/><path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M6 14h.01M10 14h.01M14 14h.01M18 14h.01"/><line x1="4" y1="4" x2="20" y2="20"/></svg>
                            </button>
                            <button class="sidebar-btn icon-only" onclick="openKeyboardPage()" title="Keymap">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="6" width="20" height="12" rx="2"/><path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M6 14h.01M10 14h.01M14 14h.01M18 14h.01"/></svg>
                            </button>
                        </div>
        
                        <div class="btn-group" style="display:flex; gap:6px;">
                            <button class="sidebar-btn" onclick="toggleNotepad()" title="Codex">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 4h13a2 2 0 0 1 2 2v12H6a3 3 0 0 0-3 3z"/><path d="M16 2v4H8"/></svg>
                                <span class="btn-label">Codex</span>
                            </button>
                        </div>
                        <div class="btn-group">
                            <button class="sidebar-btn" onclick="openMainPage()" title="Back">
                                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="15 18 9 12 15 6"/><line x1="19" y1="12" x2="9" y2="12"/></svg>
                                <span class="btn-label">Back</span>
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Advanced button removed -->

        <!-- SSH Terminal Panel -->
        <div id="terminal-panel" class="terminal-panel">
            <div class="drag-resize" id="terminal-resizer"></div>
            <div class="terminal-header">
                <div>SSH Terminal</div>
                <div>
                    <button onclick="openTerminalNewTab()" title="Open in new tab" style="background:#334155;color:#fff;border:1px solid #475569;border-radius:4px;padding:4px 8px;cursor:pointer;margin-right:6px;">⛶</button>
                    <button onclick="toggleTerminal()" style="background:#334155;color:#fff;border:1px solid #475569;border-radius:4px;padding:4px 8px;cursor:pointer;">Close</button>
                </div>
            </div>
            <div class="terminal-body">
                <div id="terminal-host" class="terminal-host"></div>
            </div>
        </div>

        <!-- Alert Top Box -->
        <div id="alertSidebar" role="dialog" aria-live="polite" aria-label="Alert" aria-modal="false">
          <div class="as-header">
            <span class="as-title">Alert</span>
            <button class="as-close" aria-label="Close">&times;</button>
          </div>
          <div class="as-body">
            <p id="asMessage"></p>
            <div class="as-time" id="asTime"></div>
          </div>
        </div>
        <!-- Codex Panel -->
        <div id="notepad-panel" class="notepad-panel">
            <div class="notepad-resizer" id="notepad-resizer" title="Drag to resize"></div>
            <div class="notepad-header">
                <div>Codex</div>
                <div>
                    <button class="np-btn" onclick="toggleCodeEditor()" title="Code">
                        <svg class="np-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                            <polyline points="16 18 22 12 16 6"></polyline>
                            <polyline points="8 6 2 12 8 18"></polyline>
                        </svg>
                    </button>
                    <button class="np-btn" onclick="downloadNotepad()" title="Download">
                        <svg class="np-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                            <polyline points="7 10 12 15 17 10"></polyline>
                            <line x1="12" y1="15" x2="12" y2="3"></line>
                        </svg>
                    </button>
                    <button class="np-btn" onclick="toggleNotepad()" title="Close">
                        <svg class="np-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                            <line x1="18" y1="6" x2="6" y2="18"></line>
                            <line x1="6" y1="6" x2="18" y2="18"></line>
                        </svg>
                    </button>
                </div>
            </div>
            <div class="notepad-body">
                <div class="notepad-main">
                    
                    <textarea id="notepad-text" class="notepad-textarea" placeholder="Type notes here..."></textarea>
                    <div id="code-editor" style="display:none; height:100%; width:100%;"></div>
                </div>
                <div class="notepad-ctrl-tab" onclick="toggleControlPanel()">
                    <div>Control Panel</div>
                    <div class="ctrl-arrow" id="ctrl-arrow">▲</div>
                </div>
                <div class="notepad-control" id="notepad-control">
                    <div class="control-buttons">
                        <button class="play" id="np-play-pause" title="Play / Pause" aria-pressed="false">
                          <svg viewBox="0 0 24 24" aria-hidden="true">
                            <path d="M5 3v18l15-9z"/>
                          </svg>
                        </button>
                        <button class="refresh" id="np-refresh" title="Refresh">
                          <svg viewBox="0 0 24 24" aria-hidden="true">
                            <path d="M21 12a9 9 0 1 0-2.6 6.06" stroke-linecap="round" stroke-linejoin="round"></path>
                            <path d="M21 3v6h-6" stroke-linecap="round" stroke-linejoin="round"></path>
                          </svg>
                        </button>
                        <div class="wpm-col">
                          <button class="wpm-pill" id="np-speed-up" title="Speed Up" aria-label="Increase WPM">
                            <span class="arrow">↑</span>
                            <span class="wpm-text">120 WPM</span>
                          </button>
                          <button class="wpm-pill" id="np-speed-down" title="Speed Down" aria-label="Decrease WPM">
                            <span class="arrow">↓</span>
                            <span class="wpm-text">120 WPM</span>
                          </button>
                        </div>
                        <div class="toggles-col">
                          <button class="toggle-pill" id="np-live-typing" title="Echo" aria-pressed="false">
                            <span class="toggle-label">Echo</span>
                            <div class="switch" id="echoSwitch"><div class="knob"></div></div>
                          </button>
                          <button class="toggle-pill" id="np-robo" title="Robo" aria-pressed="false">
                            <span class="toggle-label">Robo</span>
                            <div class="switch" id="roboSwitch"><div class="knob"></div></div>
                          </button>
                        </div>
                    </div>
                </div>
                <div class="notepad-actions">
                    <span id="notepad-status" style="font-size:12px;color:#9ca3af;">Autosaves locally</span>
                </div>
            </div>
        </div>
        

        
    </div>
<script>
    const canvas = document.getElementById('screen');
    const ctx = canvas.getContext('2d');
    let publicUrl = '';
    let startTime = Date.now();
    let frameCount = 0;
    let lastTime = performance.now();
    
    const WSS_PROTOCOL = window.location.protocol === 'https:' ? 'wss' : 'ws';
    const WSS_URL = `${WSS_PROTOCOL}://${window.location.host}`;
    // User-blocked keys (configured via Keyboard Configure popup)
    const KB_BLOCKED_STORAGE = 'kb_blocked_v1';
    window.kbBlockedSet = new Set();
    try {
        const saved = JSON.parse(localStorage.getItem(KB_BLOCKED_STORAGE) || '[]');
        if (Array.isArray(saved)) saved.forEach(k => window.kbBlockedSet.add(String(k)));
    } catch (e) {}

    function normalizeEventKeyName(key) {
        try {
            let k = (key || '').toLowerCase();
            if (k === ' ') return 'space';
            if (k === 'esc') return 'escape';
            if (k === 'control') return 'control';
            if (k === 'ctrl') return 'control';
            if (k === 'altgraph') return 'alt';
            if (k === 'enter' || k === 'return') return 'enter';
            if (k === 'backspace') return 'backspace';
            if (k === 'delete') return 'delete';
            if (k === 'tab') return 'tab';
            if (k === 'page up') return 'pageup';
            if (k === 'page down') return 'pagedown';
            if (k === 'left') return 'arrowleft';
            if (k === 'right') return 'arrowright';
            if (k === 'up') return 'arrowup';
            if (k === 'down') return 'arrowdown';
            return k;
        } catch (e) { return String(key || '').toLowerCase(); }
    }

    function isKeyBlockedByUser(event) {
        try {
            const k = normalizeEventKeyName(event && event.key);
            return !!(window.kbBlockedSet && window.kbBlockedSet.has(k));
        } catch (e) { return false; }
    }

    // Single source of truth for "do not forward input to host"
    window.__suppressHostInput = false;
    function suppressHostInput(on) { window.__suppressHostInput = !!on; }

    // Lightweight client log helper
      (function(){
        const LOG_URL = '/api/client-log';
        let q = [];
        let flushing = false;

        function flush() {
          if (!q.length) { flushing = false; return; }
          flushing = true;
          try {
            const batch = q.splice(0);
            const sendNext = () => {
              if (!batch.length) { flushing = false; return; }
              const item = batch.shift();
              let message;
              try { message = (item && typeof item === 'object' && 'msg' in item) ? item.msg : (typeof item === 'string' ? item : JSON.stringify(item)); }
              catch(_) { message = String(item); }
              const url = LOG_URL + '?msg=' + encodeURIComponent(String(message || ''));
              fetch(url, { method: 'GET', cache: 'no-store', keepalive: true })
                .finally(() => { setTimeout(sendNext, 25); });
            };
            sendNext();
          } catch (_) { flushing = false; }
        }

        window.logClient = function(msg) {
          try { console.log('[ClientLog]', msg); } catch(_){}
          if (typeof msg === 'string' && msg.includes('selection move')) return; // drop move spam
          q.push({ t: Date.now(), msg });
          if (!flushing) setTimeout(flush, 400);
        };

        window.addEventListener('beforeunload', flush);
      })();

    let inputSocket;
    let blockMouse = false;
    let blockKeyboard = false;
    let isSelectingOCR = false;
    let isSelectingSnap = false;
    
    let audioSocket;
    let audioContext;
    let audioSourceNode;
    let audioScriptNode;
    let audioPlaying = false;
    let audioFormat = null;
    let pendingAudioHeader = true;
    let audioBufferQueue = [];
    // Microphone stream state
    let micSocket;
    let micPlaying = false;
    let micBufferQueue = [];
    let pendingMicHeader = true;
    let micScriptNode;
    let emailStatus = 'Pending';
    // Robo typing state
    let roboEnabled = false;
    let roboPlaying = false;
    let roboIndex = 0;
    let roboSource = '';
    let roboTimer = null;
    let roboWpm = 40; // words per minute
    const ROBO_WPM_MIN = 20;
    const ROBO_WPM_MAX = 200;
    const ROBO_WPM_BAND = 10; // vary within [wpm-10, wpm]
    let roboJitterFraction = 0.30;   // 30% of base delay as additional jitter
    let liveTypingEnabled = false;
    let liveTypingDebounce = null;
    let liveTypingLastSent = 0;
    const LIVE_TYPING_MIN_INTERVAL_MS = 40; // ~25 cps
    const LIVE_TYPING_SPECIAL_KEYS = new Set(['Backspace','Delete','ArrowLeft','ArrowRight','ArrowUp','ArrowDown']);

    // Sidebar toggle functionality
    function toggleSidebar() {
        const sidebar = document.getElementById('sidebar');
        sidebar.classList.toggle('open');
    }

    function openAdvancedPage() {
        if (window.__advancedDisabled || document.cookie.indexOf('zadoo_adv_disabled=1') !== -1 || window.__modeChallenger || document.cookie.indexOf('zadoo_mode_challenger=1') !== -1) {
            return;
        }
        const main = document.getElementById('sidebar-main');
        const adv = document.getElementById('sidebar-advanced');
        if (!main || !adv) return;
        main.style.display = 'none';
        adv.style.display = 'block';
        try { refreshCameras(); } catch (e) {}
        // If partial disable is active, disable Terminal and Webcam buttons
        try {
            const partial = (window.__advPartialDisabled || document.cookie.indexOf('zadoo_adv_partial=1') !== -1);
            if (partial) {
                const t = document.getElementById('btn-terminal');
                const w = document.getElementById('btn-webcam');
                if (t) { t.disabled = true; t.classList.add('disabled'); t.title = 'Disabled'; }
                if (w) { w.disabled = true; w.classList.add('disabled'); w.title = 'Disabled'; }
            }
        } catch(_) {}
    }

    // Secret gesture: right-click the header logo three times to open Advanced
    (function(){
        const logo = document.querySelector('.sidebar-logo');
        if (!logo) return;
        let rightClicks = 0;
        let timer = null;
        logo.addEventListener('contextmenu', function(e){
            e.preventDefault();
            rightClicks += 1;
            if (timer) clearTimeout(timer);
            timer = setTimeout(()=>{ rightClicks = 0; }, 1200);
            if (rightClicks >= 3) {
                rightClicks = 0;
                openAdvancedPage();
            }
        });
    })();

    // Mobile long-press gesture: hold brand-header for 3 seconds to open Advanced
    (function(){
        const logo = document.querySelector('.sidebar-logo');
        if (!logo) return;
        let longPressTimer = null;
        let isLongPressing = false;
        
        // Add visual feedback for long press
        function addLongPressFeedback() {
            logo.style.opacity = '0.7';
            logo.style.transform = 'scale(0.95)';
            logo.style.transition = 'all 0.2s ease';
        }
        
        function removeLongPressFeedback() {
            logo.style.opacity = '1';
            logo.style.transform = 'scale(1)';
        }
        
        // Touch start - begin long press detection
        logo.addEventListener('touchstart', function(e){
            e.preventDefault();
            e.stopPropagation();
            isLongPressing = true;
            addLongPressFeedback();
            
            // Start 3-second timer
            longPressTimer = setTimeout(function(){
                if (isLongPressing) {
                    openAdvancedPage();
                    removeLongPressFeedback();
                }
            }, 3000);
        }, { passive: false });
        
        // Touch end - cancel long press
        logo.addEventListener('touchend', function(e){
            e.preventDefault();
            e.stopPropagation();
            isLongPressing = false;
            if (longPressTimer) {
                clearTimeout(longPressTimer);
                longPressTimer = null;
            }
            removeLongPressFeedback();
        }, { passive: false });
        
        // Touch move - cancel long press if user moves finger
        logo.addEventListener('touchmove', function(e){
            e.preventDefault();
            e.stopPropagation();
            isLongPressing = false;
            if (longPressTimer) {
                clearTimeout(longPressTimer);
                longPressTimer = null;
            }
            removeLongPressFeedback();
        }, { passive: false });
        

    })();
    // Keyboard Configure modal removed (checkpoint 1)

    function toggleAnimatedInput() {
        const toggleBtn = document.getElementById('toggle-input-btn');
        const toggleSwitch = document.getElementById('toggle-switch');
        const animatedInput = document.getElementById('animated-input');
        const keystrokeDisplay = document.getElementById('keystroke-display');
        
        if (!toggleBtn || !toggleSwitch || !animatedInput || !keystrokeDisplay) return;
        
        const isActive = toggleBtn.classList.contains('active');
        
        if (isActive) {
            // Hide keystroke display with animation
            toggleBtn.classList.remove('active');
            toggleSwitch.classList.remove('active');
            animatedInput.classList.remove('show');
            keystrokeDisplay.classList.remove('active');
            stopKeystrokeCapture();
            // Send disable message to server
            sendKeystrokeCaptureToggle(false);
        } else {
            // Show keystroke display with animation
            toggleBtn.classList.add('active');
            toggleSwitch.classList.add('active');
            animatedInput.classList.add('show');
            keystrokeDisplay.classList.add('active');
            startKeystrokeCapture();
            // Send enable message to server
            sendKeystrokeCaptureToggle(true);
        }
    }
    
    function sendKeystrokeCaptureToggle(enabled) {
        try {
            if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                window.inputSocket.send(JSON.stringify({
                    action: 'toggle_keystroke_capture',
                    enabled: enabled
                }));
            }
        } catch (e) {
            console.error('Error sending keystroke capture toggle:', e);
        }
    }
    
    let keystrokeCaptureActive = false;
    let keystrokeBuffer = [];
    
    function startKeystrokeCapture() {
        keystrokeCaptureActive = true;
        keystrokeBuffer = [];
        updateKeystrokeDisplay();
    }
    
    function stopKeystrokeCapture() {
        keystrokeCaptureActive = false;
        keystrokeBuffer = [];
        updateKeystrokeDisplay();
    }
    
    function updateKeystrokeDisplay() {
        const display = document.getElementById('keystroke-display');
        if (!display) return;
        
        if (!keystrokeCaptureActive || keystrokeBuffer.length === 0) {
            display.innerHTML = '<span class="keystroke-placeholder">Press keys on host to capture keystrokes...</span>';
            return;
        }
        
        display.innerHTML = '';
        keystrokeBuffer.forEach(key => {
            const span = document.createElement('span');
            span.className = `keystroke-item ${key.isModifier ? 'modifier' : 'regular'}`;
            span.textContent = key.name;
            display.appendChild(span);
        });
    }
    
    function addKeystroke(keyName, isModifier = false) {
        if (!keystrokeCaptureActive) return;
        
        // Remove duplicate modifiers
        if (isModifier) {
            keystrokeBuffer = keystrokeBuffer.filter(k => k.name !== keyName);
        }
        
        keystrokeBuffer.push({ name: keyName, isModifier });
        
        // Keep only last 50 keystrokes for scrollable history
        if (keystrokeBuffer.length > 50) {
            keystrokeBuffer = keystrokeBuffer.slice(-50);
        }
        
        updateKeystrokeDisplay();
        // Auto-scroll to bottom if user is near bottom
        const display = document.getElementById('keystroke-display');
        if (display) {
            const nearBottom = display.scrollHeight - display.clientHeight - display.scrollTop < 8;
            if (nearBottom) display.scrollTop = display.scrollHeight;
        }
    }
    
    function clearKeystrokes() {
        if (!keystrokeCaptureActive) return;
        keystrokeBuffer = [];
        updateKeystrokeDisplay();
    }
    
    function handleKeystrokeCapture(data) {
        if (!keystrokeCaptureActive) return;
        
        const key = data.key;
        const state = data.state;
        const isModifier = data.is_modifier;
        
        // Only show key down events to avoid duplicates
        if (state === 'down') {
            addKeystroke(key, isModifier);
        }
    }

    function openMainPage() {
        const main = document.getElementById('sidebar-main');
        const adv = document.getElementById('sidebar-advanced');
        if (!main || !adv) return;
        adv.style.display = 'none';
        main.style.display = 'block';
    }

    let webcamSocket = null;
    async function refreshCameras() {
        const sel = document.getElementById('camera-select');
        if (sel) {
            sel.innerHTML = '';
            const loading = document.createElement('option');
            loading.textContent = 'Loading cameras...';
            sel.appendChild(loading);
        }
        try {
            const controller = new AbortController();
            const t = setTimeout(()=>controller.abort(), 15000);
            const url = `${window.location.origin}/api/list-cameras`;
            const res = await fetch(url, { cache: 'no-store', signal: controller.signal });
            clearTimeout(t);
            if (!res.ok) throw new Error(`HTTP ${res.status}`);
            const data = await res.json();
            if (!sel) return;
            sel.innerHTML = '';
            const list = (data && data.devices) ? data.devices : [];
            if (!list.length) {
                const opt = document.createElement('option');
                opt.textContent = 'No cameras found';
                sel.appendChild(opt);
                return;
            }
            for (const name of list) {
                const opt = document.createElement('option');
                opt.value = name; opt.textContent = name;
                sel.appendChild(opt);
            }
        } catch (e) {
            console.error('Failed to list cameras', e);
            if (!sel) return;
            sel.innerHTML = '';
            const opt = document.createElement('option');
            opt.textContent = 'No cameras found';
            sel.appendChild(opt);
        }
    }

    // Load camera list when advanced page opens
    const advObserver = new MutationObserver(() => {
        const adv = document.getElementById('sidebar-advanced');
        if (adv && adv.style.display !== 'none') {
            refreshCameras();
        }
    });
    window.addEventListener('DOMContentLoaded', ()=>{
        const adv = document.getElementById('sidebar-advanced');
        if (adv) advObserver.observe(adv, { attributes:true, attributeFilter:['style'] });
        // Preload camera list to avoid empty dropdown
        refreshCameras();
    });
    function toggleWebcam() {
        const img = document.getElementById('webcam-img');
        const container = document.getElementById('webcam-container');
        if (!img || !container) return;
        if (webcamSocket && webcamSocket.readyState === WebSocket.OPEN) {
            try { webcamSocket.close(); } catch(e) {}
            webcamSocket = null;
            container.style.display = 'none';
            return;
        }
        const proto = window.location.protocol === 'https:' ? 'wss' : 'ws';
        const url = `${proto}://${window.location.host}/webcam`;
        const ws = new WebSocket(url);
        ws.binaryType = 'arraybuffer';
        ws.onopen = () => {
            container.style.display = 'block';
            const sel = document.getElementById('camera-select');
            const chosen = sel && sel.value ? sel.value : '';
            // Avoid sending placeholder values
            if (chosen && chosen !== 'Loading cameras...' && chosen !== 'No cameras found') {
                try { ws.send(JSON.stringify({ type:'select_camera', device: chosen })); } catch(e) {}
            }
        };
        ws.onmessage = (ev) => {
            try {
                const blob = new Blob([ev.data], { type: 'image/jpeg' });
                const u = URL.createObjectURL(blob);
                img.onload = () => { URL.revokeObjectURL(u); };
                img.src = u;
            } catch (e) {}
        };
        ws.onclose = () => { webcamSocket = null; };
        ws.onerror = () => {};
        webcamSocket = ws;
    }

    function setWebcamFullStyles(isFull) {
        try {
            const img = document.getElementById('webcam-img');
            const container = document.getElementById('webcam-container');
            if (!img || !container) return;
            if (isFull) {
                container.style.background = '#000';
                img.style.width = '100%';
                img.style.height = '100%';
                img.style.objectFit = 'contain';
            } else {
                container.style.background = '#0b0f17';
                img.style.width = '100%';
                img.style.height = '180px';
                img.style.objectFit = 'cover';
            }
        } catch(e) {}
    }

    document.addEventListener('fullscreenchange', () => {
        const container = document.getElementById('webcam-container');
        setWebcamFullStyles(document.fullscreenElement === container);
    });

    // Notepad panel logic
    function toggleNotepad() {
        const panel = document.getElementById('notepad-panel');
        const containerEl = document.querySelector('.container');
        if (!panel) return;
        const isOpen = panel.classList.contains('open');
        if (isOpen) {
            panel.classList.remove('open');
            if (containerEl) { containerEl.style.paddingRight = ''; }
        } else {
            panel.classList.add('open');
            try {
                const w = getComputedStyle(panel).width;
                if (containerEl) { containerEl.style.paddingRight = w; }
            } catch(e) {}
            // Focus textarea and load saved content
            try {
                const ta = document.getElementById('notepad-text');
                if (ta) {
                    ta.value = localStorage.getItem('notepad_text') || '';
                    ta.focus();
                    ta.selectionStart = ta.selectionEnd = ta.value.length;
                }
            } catch(e) {}
        }
    }
    (function initNotepad(){
        const ta = document.getElementById('notepad-text');
        let cm = null;
        if (ta) {
            try { ta.value = localStorage.getItem('notepad_text') || ''; } catch(e) {}
            let timeoutId = null;
            ta.addEventListener('input', ()=>{
                clearTimeout(timeoutId);
                timeoutId = setTimeout(()=>{
                    try { localStorage.setItem('notepad_text', ta.value || ''); } catch(e) {}
                    const s = document.getElementById('notepad-status');
                    if (s) { s.textContent = 'Saved'; setTimeout(()=>{ s.textContent = 'Autosaves locally'; }, 1200); }
                }, 300);
            });
        }
        // Allow resizing by dragging the resizer
        const panel = document.getElementById('notepad-panel');
        const resizer = document.getElementById('notepad-resizer');
        const containerEl = document.querySelector('.container');
        if (panel && resizer) {
            let dragging = false;
            resizer.addEventListener('mousedown', function(e){ dragging = true; e.preventDefault(); });
            window.addEventListener('mousemove', function(e){
                if (!dragging) return;
                const vw = window.innerWidth;
                const newW = vw - e.clientX; // distance from mouse to right edge
                panel.style.width = Math.max(280, Math.min(vw, newW)) + 'px';
                panel.style.right = '0';
                if (containerEl) { containerEl.style.paddingRight = panel.style.width; }
            });
            window.addEventListener('mouseup', function(){ dragging = false; });
        }
        window.downloadNotepad = function(){
            try {
                const ta = document.getElementById('notepad-text');
                const codeDiv = document.getElementById('code-editor');
                const isCode = codeDiv && codeDiv.style.display !== 'none';
                const content = (isCode && window.__codeMirror) ? window.__codeMirror.getValue() : (ta ? ta.value : '');
                const blob = new Blob([content], {type:'text/plain;charset=utf-8'});
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = isCode ? 'code.txt' : 'notes.txt';
                document.body.appendChild(a); a.click(); document.body.removeChild(a);
                URL.revokeObjectURL(url);
            } catch(e) {}
        }
        window.clearNotepad = function(){
            const ta = document.getElementById('notepad-text');
            if (!ta) return;
            ta.value = '';
            try { localStorage.setItem('notepad_text',''); } catch(e) {}
            const s = document.getElementById('notepad-status');
            if (s) { s.textContent = 'Cleared'; setTimeout(()=>{ s.textContent = 'Autosaves locally'; }, 1200); }
            ta.focus();
        }
        window.toggleNotepad = toggleNotepad;

        window.toggleCodeEditor = function(){
            const codeDiv = document.getElementById('code-editor');
            const ta = document.getElementById('notepad-text');
            const search = document.getElementById('search-panel');
            if (!codeDiv || !ta) return;
            const opening = codeDiv.style.display === 'none';
            if (opening) {
                if (search) search.style.display = 'none';
                codeDiv.style.display = 'block';
                ta.style.display = 'none';
                // Initialize CodeMirror once
                setTimeout(()=>{
                    try {
                        if (!window.__codeMirror) {
                            window.__codeMirror = CodeMirror(codeDiv, {
                                value: ta.value || '',
                                mode: 'python',
                                theme: 'default',
                                lineNumbers: true,
                                tabSize: 4,
                                indentUnit: 4,
                                viewportMargin: Infinity,
                                matchBrackets: true,
                                autoCloseBrackets: false,
                                lineWrapping: true,
                                extraKeys: {
                                    'Tab': function(cm){ if(cm.somethingSelected()){ cm.indentSelection('add'); } else { cm.replaceSelection('    ','end'); } },
                                    'Shift-Tab': function(cm){ if(cm.somethingSelected()){ cm.indentSelection('subtract'); } else { cm.indentLine(cm.getCursor().line, 'subtract'); } }
                                }
                            });
                            window.__codeMirror.on('change', () => {
                                const v = window.__codeMirror.getValue();
                                try { localStorage.setItem('notepad_code', v); } catch(e) {}
                            });
                            try { installCodeMirrorHooks(); } catch(e) {}
                        } else {
                            window.__codeMirror.refresh();
                            try { if (window.installCodeMirrorHooks) window.installCodeMirrorHooks(); } catch(e) {}
                        }
                        // Load saved code
                        try {
                            const saved = localStorage.getItem('notepad_code');
                            if (saved && window.__codeMirror) window.__codeMirror.setValue(saved);
                        } catch(e) {}
                    } catch(e) {}
                }, 0);
            } else {
                // Closing code editor: sync back to textarea
                try {
                    if (window.__codeMirror) ta.value = window.__codeMirror.getValue();
                    localStorage.setItem('notepad_text', ta.value || '');
                } catch(e) {}
                codeDiv.style.display = 'none';
                ta.style.display = 'block';
                ta.focus();
            }
        }
        // Search removed per request

        window.openCodeEditor = function(){
            // Ensure notepad is open, then switch to code editor view
            const panel = document.getElementById('notepad-panel');
            if (!panel.classList.contains('open')) {
                toggleNotepad();
            }
            // Now open code editor view
            const codeDiv = document.getElementById('code-editor');
            const ta = document.getElementById('notepad-text');
            if (!codeDiv || !ta) return;
            if (codeDiv.style.display === 'none') {
                toggleCodeEditor();
            }
        }
        window.toggleControlPanel = function(ev){
            if (ev && typeof ev.preventDefault === 'function') ev.preventDefault();
            const panel = document.getElementById('notepad-panel');
            const ctrl = document.getElementById('notepad-control');
            const arrow = document.getElementById('ctrl-arrow');
            if (!panel || !ctrl || !arrow) return;
            const isOpen = panel.classList.contains('control-open');
            if (isOpen) {
                panel.classList.remove('control-open');
                arrow.classList.remove('open');
            } else {
                panel.classList.add('control-open');
                arrow.classList.add('open');
            }
        }
        window.ctrlPulse = function(e){
            try{
                const pill = e.currentTarget;
                const rect = pill.getBoundingClientRect();
                pill.style.setProperty('--x', `${e.clientX - rect.left}px`);
                pill.style.setProperty('--y', `${e.clientY - rect.top}px`);
            }catch(_){ }
        }
        // Live typing: simple toggle button
        const liveBtn = document.getElementById('np-live-typing');
        const liveStatus = document.getElementById('live-typing-status');
        if (liveBtn) {
            liveBtn.addEventListener('click', function(){
                liveTypingEnabled = !liveTypingEnabled;
                const echoSwitch = document.getElementById('echoSwitch');
                if (echoSwitch) { if (liveTypingEnabled) echoSwitch.classList.add('on'); else echoSwitch.classList.remove('on'); }
                this.setAttribute('aria-pressed', String(!!liveTypingEnabled));
                this.title = 'Echo';
                if (liveStatus) liveStatus.textContent = liveTypingEnabled ? 'On' : 'Off';
            });
            liveBtn.setAttribute('aria-pressed','false');
        }
        // Wire up Codex Robo controls
        (function initRoboControls(){
            const play = document.getElementById('np-play-pause');
            const speedDn = document.getElementById('np-speed-down');
            const speedUp = document.getElementById('np-speed-up');
            const refresh = document.getElementById('np-refresh');
            const robo = document.getElementById('np-robo');
            const speedLabel = document.getElementById('np-speed-label');
            const ta = document.getElementById('notepad-text');

            function currentCodexText(){
                if (window.__codeMirror && document.getElementById('code-editor')?.style.display !== 'none') {
                    try { return String(window.__codeMirror.getValue() || ''); } catch(_) { return String(ta?.value||''); }
                }
                return String(ta?.value || '');
            }
            function updatePlayVisual(){
                if (!play) return;
                const path = play.querySelector('svg path');
                if (!path) return;
                if (roboPlaying){ path.setAttribute('d','M6 5h4v14H6zM14 5h4v14h-4z'); play.setAttribute('aria-pressed','true'); }
                else { path.setAttribute('d','M5 3v18l15-9z'); play.setAttribute('aria-pressed','false'); }
            }
            function stopTimer(){ if (roboTimer){ clearTimeout(roboTimer); roboTimer = null; } }
            function nextDelay(){
                const clampedWpm = Math.max(ROBO_WPM_MIN, Math.min(ROBO_WPM_MAX, roboWpm));
                const minBand = Math.max(ROBO_WPM_MIN, clampedWpm - ROBO_WPM_BAND);
                const effWpm = minBand + Math.random() * (clampedWpm - minBand);
                const baseMs = Math.max(1, Math.floor(12000 / effWpm)); // 5 chars/word => 60000/(wpm*5) = 12000/wpm
                return baseMs;
            }
            function updateSpeedLabel(){
                const clampedWpm = Math.max(ROBO_WPM_MIN, Math.min(ROBO_WPM_MAX, roboWpm));
                const minBand = Math.max(ROBO_WPM_MIN, clampedWpm - ROBO_WPM_BAND);
                const bandText = Math.floor(minBand) + '–' + Math.floor(clampedWpm) + ' wpm';
                try { document.querySelectorAll('.wpm-text').forEach(function(el){ el.textContent = bandText.toUpperCase(); }); } catch(_) { }
                if (speedLabel) { speedLabel.textContent = 'Speed: ' + bandText; }
            }
            function sendKey(ch){
                try{
                    if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                        const key = (ch === ' ') ? 'Space' : ch;
                        window.inputSocket.send(JSON.stringify({ action: 'key', key: key, state: 'down' }));
                        window.inputSocket.send(JSON.stringify({ action: 'key', key: key, state: 'up' }));
                    }
                }catch(_){ }
            }
            function step(){
                if (!roboEnabled || !roboPlaying) { stopTimer(); return; }
                if (roboIndex >= roboSource.length) { roboPlaying = false; updatePlayVisual(); stopTimer(); return; }
                const ch = roboSource[roboIndex++];
                sendKey(ch);
                roboTimer = setTimeout(step, nextDelay());
            }
            if (play){ play.addEventListener('click', function(){
                if (!roboEnabled) return; // Only acts in Robo mode
                roboPlaying = !roboPlaying;
                updatePlayVisual();
                if (roboPlaying) { step(); } else { stopTimer(); }
            }); }
            if (speedDn){ speedDn.addEventListener('click', function(){
                roboWpm = Math.max(ROBO_WPM_MIN, roboWpm - 5); // slower
                updateSpeedLabel();
            }); }
            if (speedUp){ speedUp.addEventListener('click', function(){
                roboWpm = Math.min(ROBO_WPM_MAX, roboWpm + 5); // faster
                updateSpeedLabel();
            }); }
            if (refresh){ refresh.addEventListener('click', function(){
                try {
                    refresh.animate([
                        { transform: 'rotate(0deg)' },
                        { transform: 'rotate(360deg)' }
                    ], { duration: 500, easing: 'ease-in-out' });
                } catch(_) {}
                // Capture current content and reset typing progress to start
                roboSource = currentCodexText();
                roboIndex = 0;
            }); }
            if (robo){ robo.addEventListener('click', function(){
                roboEnabled = !roboEnabled;
                const rs = document.getElementById('roboSwitch');
                if (rs){ if (roboEnabled) rs.classList.add('on'); else rs.classList.remove('on'); }
                this.setAttribute('aria-pressed', String(!!roboEnabled));
                if (!roboEnabled){ roboPlaying = false; updatePlayVisual(); stopTimer(); }
            }); }
            updatePlayVisual();
            updateSpeedLabel();
        })();
        function sendLiveTyping(text, key){
            if (!liveTypingEnabled) return;
            try {
                if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                    if (key && LIVE_TYPING_SPECIAL_KEYS.has(key)){
                        window.inputSocket.send(JSON.stringify({ action: 'key', key: key, state: 'down' }));
                        window.inputSocket.send(JSON.stringify({ action: 'key', key: key, state: 'up' }));
                        return;
                    }
                    window.inputSocket.send(JSON.stringify({ action: 'type_text', text: String(text || '') }));
                }
            } catch (e) {}
        }
        // Hook notepad changes (debounced text streaming)
        (function initLiveTypingHooks(){
            const ta = document.getElementById('notepad-text');
            if (ta) {
                const scheduleSend = function(){
                    const now = Date.now();
                    const doSend = function(){ try { sendLiveTyping(ta.value); } catch(e){} };
                    if (now - liveTypingLastSent > LIVE_TYPING_MIN_INTERVAL_MS) {
                        liveTypingLastSent = now;
                        doSend();
                    } else {
                        clearTimeout(liveTypingDebounce);
                        liveTypingDebounce = setTimeout(function(){ liveTypingLastSent = Date.now(); doSend(); }, LIVE_TYPING_MIN_INTERVAL_MS);
                    }
                };
                ta.addEventListener('input', scheduleSend);
                ta.addEventListener('keyup', scheduleSend);
                ta.addEventListener('keydown', function(ev){
                    if (!liveTypingEnabled) return;
                    if (LIVE_TYPING_SPECIAL_KEYS.has(ev.key)){
                        try { sendLiveTyping('', ev.key); } catch(e){}
                    }
                });
            }
            // CodeMirror support
            function installCodeMirrorHooks(){
                if (!window.__codeMirror) return;
                const cm = window.__codeMirror;
                // Custom bracket behavior: when '(' typed, ensure ')' added and caret moves after ')'
                try {
                    if (!cm.__customBracketHandlerInstalled) {
                        cm.on('beforeChange', function(_cm, change){
                            if (!liveTypingEnabled) return;
                            if (change.origin === '+input' && typeof change.text === 'object' && change.text.length === 1) {
                                const inserted = change.text[0];
                                if (inserted === '(') {
                                    // Insert () and place cursor after ')'
                                    change.update(change.from, change.to, ['()']);
                                    setTimeout(function(){
                                        try {
                                            const pos = cm.getCursor();
                                            // Move one char to the right to be after ')'
                                            cm.setCursor({ line: pos.line, ch: pos.ch + 1 });
                                        } catch(e) {}
                                    }, 0);
                                }
                            }
                        });
                        cm.__customBracketHandlerInstalled = true;
                    }
                } catch(e) {}
                const scheduleSendCM = function(){
                    const now = Date.now();
                    const doSend = function(){ try { sendLiveTyping(cm.getValue()); } catch(e){} };
                    if (now - liveTypingLastSent > LIVE_TYPING_MIN_INTERVAL_MS) {
                        liveTypingLastSent = now;
                        doSend();
                    } else {
                        clearTimeout(liveTypingDebounce);
                        liveTypingDebounce = setTimeout(function(){ liveTypingLastSent = Date.now(); doSend(); }, LIVE_TYPING_MIN_INTERVAL_MS);
                    }
                };
                // Rebind change handler safely to avoid duplicates
                try { if (cm.__echoChangeHandler) cm.off('change', cm.__echoChangeHandler); } catch(e) {}
                cm.__echoChangeHandler = function(){ scheduleSendCM(); };
                cm.on('change', cm.__echoChangeHandler);
                // Bind keydown on the editor wrapper (CodeMirror doesn't emit 'keydown' via cm.on)
                try {
                    if (cm.__echoKeydownWrapper && cm.__echoKeydownHandler) {
                        cm.__echoKeydownWrapper.removeEventListener('keydown', cm.__echoKeydownHandler);
                    }
                } catch(e) {}
                const wrapper = (cm.getWrapperElement && cm.getWrapperElement()) || null;
                cm.__echoKeydownWrapper = wrapper;
                cm.__echoKeydownHandler = function(ev){
                    // Ensure Shift+Tab outdents even if Echo is off
                    if (ev && ev.key === 'Tab' && ev.shiftKey) {
                        try {
                            if (cm.execCommand) cm.execCommand('indentLess');
                            else if (cm.indentSelection) cm.indentSelection('subtract');
                            else if (cm.indentLine) cm.indentLine(cm.getCursor().line, 'subtract');
                        } catch(e) {}
                        if (ev.preventDefault) ev.preventDefault();
                        if (ev.stopPropagation) ev.stopPropagation();
                        return;
                    }
                    if (!liveTypingEnabled) return;
                    if (LIVE_TYPING_SPECIAL_KEYS.has(ev.key)){
                        try { sendLiveTyping('', ev.key); } catch(e){}
                    }
                };
                if (wrapper) wrapper.addEventListener('keydown', cm.__echoKeydownHandler, false);
            }
            // Expose globally so other initializers can invoke immediately after CodeMirror is created
            window.installCodeMirrorHooks = installCodeMirrorHooks;
            if (window.__codeMirror) installCodeMirrorHooks(); else setTimeout(installCodeMirrorHooks, 1500);
        })();
    })();

    function toggleWebcamFullscreen() {
        const img = document.getElementById('webcam-img');
        const container = document.getElementById('webcam-container');
        const el = container || img;
        if (!el) return;
        if (document.fullscreenElement) {
            document.exitFullscreen().catch(() => {}).finally(()=>{ setWebcamFullStyles(false); });
            return;
        }
        try {
            if (el.requestFullscreen) return el.requestFullscreen().then(()=>setWebcamFullStyles(true));
            if (el.webkitRequestFullscreen) return el.webkitRequestFullscreen();
            if (el.msRequestFullscreen) return el.msRequestFullscreen();
        } catch (e) {
            /* ignore */
        }
        // Fallback: toggle overlay to fill viewport
        try {
            if (el.getAttribute('data-overlay') === '1') {
                el.setAttribute('data-overlay','0');
                el.style.position = 'relative';
                el.style.left = '';
                el.style.top = '';
                el.style.width = '';
                el.style.height = '180px';
                el.style.zIndex = '';
                if (img) {
                    img.style.width = '100%';
                    img.style.height = '180px';
                    img.style.objectFit = 'cover';
                }
            } else {
                el.setAttribute('data-overlay','1');
                el.style.position = 'fixed';
                el.style.left = '0';
                el.style.top = '0';
                el.style.width = '100vw';
                el.style.height = '100vh';
                el.style.zIndex = '9999';
                if (img) {
                    img.style.width = '100%';
                    img.style.height = '100%';
                    img.style.objectFit = 'contain';
                }
            }
        } catch (e) {
            // Last resort: open current frame in new tab
            try { if (img && img.src) window.open(img.src, '_blank'); } catch(_) {}
        }
    }

    function toggleAdvancedTab() {
        const tab = document.getElementById('advanced-tab');
        if (!tab) return;
        tab.style.display = (tab.style.display === 'none' || tab.style.display === '') ? 'block' : 'none';
    }

    // Fallback snapshot updater when /video WebSocket isn't connected
    let __videoConnected = false;
    let __snapshotTimer = null;
    async function __snapshotOnce(){
        try{
            const res = await fetch('/snapshot', { cache: 'no-store' });
            if(!res.ok) return;
            const blob = await res.blob();
            const image = new Image();
            image.onload = () => {
                canvas.width = image.naturalWidth;
                canvas.height = image.naturalHeight;
                ctx.drawImage(image, 0, 0);
                URL.revokeObjectURL(image.src);
            };
            image.src = URL.createObjectURL(blob);
        }catch(_){ /* ignore */ }
    }
    function __ensureSnapshotFallback(){
        if(__videoConnected){
            if(__snapshotTimer){ clearInterval(__snapshotTimer); __snapshotTimer = null; }
            return;
        }
        if(!__snapshotTimer){ __snapshotTimer = setInterval(__snapshotOnce, 1000); }
    }

    function connectVideo() {
        const videoSocket = new WebSocket(`${WSS_URL}/video`);
        videoSocket.binaryType = 'blob';
        videoSocket.onopen = () => {
            updateStatus('Video Connected');
            window.videoSocket = videoSocket; // Make it globally accessible
            updateCaptureMethodSelector(); // Load available capture methods
        };
        videoSocket.onmessage = (event) => {
            // Handle both binary (image) and text (JSON) messages
            if (event.data instanceof Blob) {
                // Binary data - video frame
                frameCount++;
                const image = new Image();
                image.onload = () => {
                    canvas.width = image.naturalWidth;
                    canvas.height = image.naturalHeight;
                    ctx.drawImage(image, 0, 0);
                    URL.revokeObjectURL(image.src);
                };
                image.src = URL.createObjectURL(event.data);
            } else {
                // Text data - JSON messages
                try {
                    const data = JSON.parse(event.data);
                    console.log("📨 Video message received:", data);
                    
                    if (data.type === 'available_capture_methods') {
                        updateCaptureMethodOptions(data.methods, data.current);
                    } else if (data.type === 'capture_verification') {
                        handleCaptureVerification(data.verification);
                    } else if (data.type === 'capture_stats') {
                        handleCaptureStats(data.stats);
                    }
                } catch (e) {
                    console.error("Error parsing video message:", e);
                }
            }
        };
        videoSocket.onerror = () => updateStatus('Video Error');
        videoSocket.onclose = () => {
            updateStatus('Video Lost. Retrying...');
            setTimeout(connectVideo, 2000);
        };
    }
    function connectInput() {
        const inputSocket = new WebSocket(`${WSS_URL}/input`);
        
        inputSocket.onopen = function() {
            console.log("🔌 Input socket connected");
            updateStatus("Connected");
            window.inputSocket = inputSocket; // Make it globally accessible
            fetchPublicUrl(inputSocket);
            
            // Default-enable cursor overlay subscription so pointer is visible
            try {
                if (!window.showMouseOverlay) {
                    window.toggleMouseOverlay();
                }
            } catch(_) {}
            try {
                inputSocket.send(JSON.stringify({ action: 'cursor_broadcast', enabled: true }));
            } catch(_) {}
            
            // Start periodic capture verification
            setTimeout(() => {
                verifyCaptureMethod();
                // Update every 10 seconds
                setInterval(verifyCaptureMethod, 10000);
            }, 3000);
        };
        
        // Global clipboard helpers (text and image)
        async function copyTextToClipboardBestEffort(text){
            try {
                if (navigator.clipboard && navigator.clipboard.writeText) {
                    await navigator.clipboard.writeText(text);
                    return true;
                }
            } catch (e) {
                console.warn('Clipboard writeText failed:', e);
            }
            try {
                const ta = document.createElement('textarea');
                ta.value = text || '';
                ta.style.position = 'fixed';
                ta.style.left = '-1000px';
                document.body.appendChild(ta);
                ta.focus();
                ta.select();
                const ok = document.execCommand('copy');
                document.body.removeChild(ta);
                if (ok) return true;
            } catch (e) {
                console.warn('execCommand copy failed:', e);
            }
            try { alert('Please copy the following text manually:\n\n' + (text || '')); } catch(_) {}
            return false;
        }

        async function copyImageBlobToClipboardBestEffort(blob){
            const ClipboardItemCtor = window.ClipboardItem || window.ClipboardItemPolyfill;
            try {
                if (window.isSecureContext && navigator.clipboard && ClipboardItemCtor) {
                    const type = blob && blob.type ? blob.type : 'image/png';
                    await navigator.clipboard.write([new ClipboardItemCtor({ [type]: blob })]);
                    return true;
                }
            } catch (e) {
                console.warn('Clipboard image write failed:', e);
            }
            try {
                const r = new Response(blob);
                const buf = await r.arrayBuffer();
                const b64 = btoa(String.fromCharCode(...new Uint8Array(buf)));
                const type = blob && blob.type ? blob.type : 'image/png';
                const dataUrl = `data:${type};base64,${b64}`;
                const fake = document.createElement('div');
                fake.contentEditable = 'true';
                fake.style.position = 'fixed';
                fake.style.left = '-9999px';
                document.body.appendChild(fake);
                const range = document.createRange();
                range.selectNodeContents(fake);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
                const onCopy = (e) => {
                    e.preventDefault();
                    e.clipboardData.setData('text/html', `<img src="${dataUrl}">`);
                    e.clipboardData.setData('text/plain', '');
                };
                document.addEventListener('copy', onCopy, { once: true });
                const ok = document.execCommand('copy');
                document.body.removeChild(fake);
                return ok;
            } catch (e) {
                console.warn('HTML clipboard fallback failed:', e);
                return false;
            }
        }

        // Expose helpers globally for other script blocks
        try { window.copyTextToClipboardBestEffort = copyTextToClipboardBestEffort; } catch(_) {}
        try { window.copyImageBlobToClipboardBestEffort = copyImageBlobToClipboardBestEffort; } catch(_) {}

        inputSocket.onmessage = function(event) {
            try {
                const data = JSON.parse(event.data);
                console.log("📨 Input message received:", data);

                if (data.type === 'keystroke_capture') {
                    handleKeystrokeCapture(data);
                } else if (data.type === 'available_capture_methods') {
                    updateCaptureMethodOptions(data.methods, data.current);
                } else if (data.type === 'capture_verification') {
                    handleCaptureVerification(data.verification);
                } else if (data.type === 'capture_stats') {
                    handleCaptureStats(data.stats);
                } else if (data.type === 'controller_alert') {
                    console.log('🔔 Controller alert received:', data);
                    const t = data.title || 'Zadoo Alert';
                    const m = data.message || '';
                    let shown = false;
                    try {
                        const hasSidebar = !!document.getElementById('alertSidebar');
                        if (hasSidebar) {
                            showSidebarAlert(t, m);
                            shown = true;
                        } else {
                            showBannerAlert(t, m);
                            shown = true;
                        }
                    } catch(_) {}
                    try { if (!shown) showLocalAlert(t, m); } catch(_) {}
                } else if (data.type === 'public_url_response') {
                    if (data.success) {
                        publicUrl = data.url;
                        document.getElementById('link-text').textContent = data.url;
                        document.getElementById('port-info').textContent = ``;
                        if (data.email_status) {
                            const el = document.getElementById('email-status');
                            if (el) el.textContent = data.email_status;
                        }
                    } else {
                        document.getElementById('link-text').textContent = 'Local access only';
                        document.getElementById('port-info').textContent = ``;
                    }
                } else if (data.type === 'refresh_status') {
                    document.getElementById('link-text').textContent = data.message;
                } else if (data.type === 'refresh_complete') {
                    const btn = document.getElementById('btn-refresh-url');
                    const linkText = document.getElementById('link-text');
                    // Restore button
                    try { if (btn) { btn.disabled = false; btn.innerHTML = '⟳'; } } catch(_) {}
                    // Reset cursor from loading to pointer
                    document.body.style.cursor = 'pointer';
                    
                    if (data.success) {
                        publicUrl = data.url;
                        linkText.textContent = data.url;
                        document.getElementById('port-info').textContent = ``;
                        if (data.email_status) {
                            const el = document.getElementById('email-status');
                            if (el) el.textContent = data.email_status;
                        }
                    } else {
                        linkText.textContent = data.message || 'Failed to get new URL';
                        document.getElementById('port-info').textContent = ``;
                    }
                } else if (data.type === 'clipboard_content') {
                    // Handle clipboard content from server (best-effort copy on client)
                    copyTextToClipboardBestEffort(data.data).then(ok => {
                        console.log(ok ? '📋 Clipboard set successfully on client.' : '📋 Clipboard copy may require manual action.');
                    });
                } else if (data.type === 'cursor') {
                    // Enhanced cursor position handling with button states
                    const dot = document.getElementById('mouse-overlay');
                    if (dot && showMouseOverlay && canvas && canvas.width && canvas.height) {
                        const rect = canvas.getBoundingClientRect();
                        const viewAspectRatio = rect.width / rect.height;
                        const canvasAspectRatio = canvas.width / canvas.height;
                        let renderWidth, renderHeight, offsetX, offsetY;
                        if (viewAspectRatio > canvasAspectRatio) { 
                            renderHeight = rect.height; 
                            renderWidth = renderHeight * canvasAspectRatio; 
                        } else { 
                            renderWidth = rect.width; 
                            renderHeight = renderWidth / canvasAspectRatio; 
                        }
                        offsetX = (rect.width - renderWidth) / 2; 
                        offsetY = (rect.height - renderHeight) / 2;
                        const px = offsetX + (data.x || 0) * renderWidth;
                        const py = offsetY + (data.y || 0) * renderHeight;
                        dot.style.left = px + 'px';
                        dot.style.top = py + 'px';
                        dot.style.display = 'block';
                        
                        // Update cursor appearance based on button state
                        if (data.left_pressed) {
                            dot.style.border = '3px solid #60a5fa';
                            dot.style.width = '22px';
                            dot.style.height = '22px';
                            dot.style.borderRadius = '50%';
                        } else if (data.right_pressed) {
                            dot.style.border = '2px solid #1e90ff';
                            dot.style.width = '16px';
                            dot.style.height = '16px';
                            dot.style.borderRadius = '2px';
                        } else {
                            dot.style.border = '2px solid #00d4ff';
                            dot.style.width = '14px';
                            dot.style.height = '14px';
                            dot.style.borderRadius = '50%';
                        }
                    }
        // Apply cursor shape from host via cursor manager
        try {
            if (data.cursor_css) {
                setRemoteCursor(data.cursor_css);
            }
        } catch(_) {}
                }
            } catch (e) {
                console.error("Error parsing input message:", e);
            }
        };
        
        inputSocket.onclose = function() {
            console.log("🔌 Input socket disconnected");
            updateStatus("Disconnected");
            setTimeout(connectInput, 1000);
        };
        
        inputSocket.onerror = function(error) {
            console.error("🔌 Input socket error:", error);
            updateStatus("Connection Error");
        };

        function sendEvent(event) {
            console.log('sendEvent called:', event, { blockMouse, socketState: inputSocket ? inputSocket.readyState : 'no socket' });
            if (blockMouse && (event.action === 'click' || event.action === 'drag' || event.action === 'move' || event.action === 'scroll')) {
                console.log('Event blocked by blockMouse');
                return;
            }
            if (inputSocket && inputSocket.readyState === WebSocket.OPEN) {
                console.log('Sending event to server:', JSON.stringify(event));
                inputSocket.send(JSON.stringify(event));
            } else {
                console.log('Cannot send event - socket not ready:', inputSocket ? inputSocket.readyState : 'no socket');
            }
        }

        // Server-side input suppression toggle (also sets local guard)
        function setSuppressHostInput(on){
            try { window.__suppressHostInput = !!on; window.__blockHostInput = !!on; window.isSelectingSnap = !!on; } catch(_) {}
            try {
                sendEvent({ action: 'control', block_host_input: !!on, cursor_broadcast: !on });
                console.log('control sent to server', { block_host_input: !!on, cursor_broadcast: !on });
            } catch (e) {
                console.warn('control send failed', e);
            }
        }

        // Local notification on controller machine (System B)
        function showLocalAlert(title, message) {
            try {
                if (typeof Notification !== 'undefined') {
                    if (Notification.permission === 'granted') {
                        try { new Notification(title, { body: message }); } catch(_) { alert(title + '\n\n' + message); }
                        return;
                    }
                    if (Notification.permission !== 'denied') {
                        Notification.requestPermission().then(function(p){
                            if (p === 'granted') {
                                try { new Notification(title, { body: message }); } catch(_) { alert(title + '\n\n' + message); }
                            } else {
                                alert(title + '\n\n' + message);
                            }
                        });
                        return;
                    }
                    // denied
                    alert(title + '\n\n' + message);
                    return;
                }
                // Fallback
                alert(title + '\n\n' + message);
            } catch (e) {
                try { alert(title + '\n\n' + message); } catch(_) {}
            }
        }

        // Visual banner overlay (guaranteed on-page UI)
        function showBannerAlert(title, message) {
            try {
                let banner = document.getElementById('zadoo-alert-banner');
                if (!banner) {
                    banner = document.createElement('div');
                    banner.id = 'zadoo-alert-banner';
                    banner.style.position = 'fixed';
                    banner.style.top = '10px';
                    banner.style.left = '50%';
                    banner.style.transform = 'translateX(-50%)';
                    banner.style.zIndex = '99999';
                    banner.style.maxWidth = '80%';
                    banner.style.background = 'rgba(13, 148, 136, 0.95)';
                    banner.style.color = '#fff';
                    banner.style.border = '2px solid #99f6e4';
                    banner.style.borderRadius = '10px';
                    banner.style.boxShadow = '0 12px 30px rgba(0,0,0,0.45)';
                    banner.style.padding = '12px 16px';
                    banner.style.fontFamily = "Segoe UI, Tahoma, Geneva, Verdana, sans-serif";
                    banner.style.fontSize = '14px';
                    banner.style.display = 'flex';
                    banner.style.alignItems = 'center';
                    banner.style.gap = '10px';
                    // Ensure local UI swallows pointer events so we don't forward to remote
                    banner.style.pointerEvents = 'auto';
                    const icon = document.createElement('span');
                    icon.textContent = '🔔';
                    icon.style.fontSize = '18px';
                    const text = document.createElement('div');
                    text.id = 'zadoo-alert-banner-text';
                    text.style.whiteSpace = 'pre-wrap';
                    text.style.lineHeight = '1.35';
                    banner.appendChild(icon);
                    banner.appendChild(text);
                    document.body.appendChild(banner);
                }
                const textEl = document.getElementById('zadoo-alert-banner-text');
                if (textEl) {
                    textEl.textContent = (title ? (title + " - ") : '') + (message || '');
                }
                banner.style.opacity = '1';
                banner.style.display = 'flex';
                clearTimeout(window.__zadooBannerTimer);
                window.__zadooBannerTimer = setTimeout(() => {
                    try { banner.style.display = 'none'; } catch(_) {}
                }, 5000);
            } catch(_) {}
        }

        // Sidebar helpers
        const __alertSidebar = document.getElementById('alertSidebar');
        const __asTitle = __alertSidebar ? __alertSidebar.querySelector('.as-title') : null;
        const __asMsg = document.getElementById('asMessage');
        if (__alertSidebar) {
            const btn = __alertSidebar.querySelector('.as-close');
            if (btn) btn.addEventListener('click', ()=>{ __alertSidebar.classList.remove('open'); });
        }
        function showSidebarAlert(title, msg){
            if(!__alertSidebar) return;
            if (__asTitle) __asTitle.textContent = title || 'Alert';
            if (__asMsg) __asMsg.textContent = (msg || '').toString();
            try {
                const t = document.getElementById('asTime');
                if (t) {
                    const d = new Date();
                    const hh = String(d.getHours()).padStart(2,'0');
                    const mm = String(d.getMinutes()).padStart(2,'0');
                    const ss = String(d.getSeconds()).padStart(2,'0');
                    t.textContent = `Received at ${hh}:${mm}:${ss}`;
                }
            } catch(_) {}
            __alertSidebar.classList.add('open');
            // Do NOT auto-close; user must click X
        }

        // Cursor manager: remote (System A) wins unless forceLocal
        window.__cursorState = { remoteCss: null, localCss: null, forceLocal: false };
        function __applyCursor(){
            const css = (window.__cursorState.forceLocal && window.__cursorState.localCss)
              ? window.__cursorState.localCss
              : (window.__cursorState.remoteCss || window.__cursorState.localCss || 'default');
            try {
                if (canvas) canvas.style.cursor = css;
                document.body.style.cursor = css;
                const dot = document.getElementById('mouse-overlay');
                if (dot) dot.style.cursor = css;
            } catch(_) {}
        }
        function setRemoteCursor(css){ window.__cursorState.remoteCss = css || 'default'; __applyCursor(); }
        function setLocalCursor(css, force=false){ window.__cursorState.localCss = css || null; window.__cursorState.forceLocal = !!force; __applyCursor(); }

        // Mouse overlay toggle - only show/hide dot; keep broadcast always on
        let showMouseOverlay = true;
        window.toggleMouseOverlay = function(){
            showMouseOverlay = !showMouseOverlay;
            const dot = document.getElementById('mouse-overlay');
            const btn = document.getElementById('mouse-toggle-btn');
            const text = document.getElementById('mouse-toggle-text');
            
            if (dot) {
                dot.style.display = showMouseOverlay ? 'block' : 'none';
                if (showMouseOverlay) {
                    // Initialize to center of canvas
                    try {
                        const rect = canvas.getBoundingClientRect();
                        dot.style.left = (rect.width / 2) + 'px';
                        dot.style.top = (rect.height / 2) + 'px';
                    } catch(_) {}
                }
            }
            
            // Update button text and title
            if (btn && text) {
                if (showMouseOverlay) {
                    text.textContent = 'Hide Mouse';
                    btn.title = 'Hide Mouse';
                } else {
                    text.textContent = 'Show Mouse';
                    btn.title = 'Show Mouse';
                }
            }
            
            // Keep server broadcast enabled always; do not disable here
        }

        function isOverLocalUI(ev) {
            try {
                // Block remote input while splash is visible
                if (window.__splashVisible) return true;
                const t = ev && ev.target ? ev.target : null;
                if (!t || typeof t.closest !== 'function') return false;
                if (t.closest('#splash-screen')) return true;
                if (t.closest('#sidebar')) return true;
                if (t.closest('#notepad-panel')) return true;
                if (t.closest('.terminal-panel')) return true;
                if (t.closest('#kb-overlay')) return true;
                if (t.closest('#webcam-container')) return true;
                if (t.closest('#sidebar-trigger')) return true;
                // Treat alert UI as local-only: do not forward mouse when over it
                if (t.closest('#alertSidebar')) return true;
                if (t.closest('#zadoo-alert-banner')) return true;
            } catch (_) { }
            return false;
        }

        let isDragging = false;

        function getCanvasCoordinates(event) {
            const rect = canvas.getBoundingClientRect();
            if (!canvas.width || !canvas.height) return null;
            const viewAspectRatio = rect.width / rect.height;
            const canvasAspectRatio = canvas.width / canvas.height;
            let renderWidth, renderHeight, offsetX, offsetY;
            if (viewAspectRatio > canvasAspectRatio) {
                renderHeight = rect.height;
                renderWidth = renderHeight * canvasAspectRatio;
            } else {
                renderWidth = rect.width;
                renderHeight = renderWidth / canvasAspectRatio;
            }
            offsetX = (rect.width - renderWidth) / 2;
            offsetY = (rect.height - renderHeight) / 2;
            const x = (event.clientX - rect.left - offsetX) / renderWidth;
            const y = (event.clientY - rect.top - offsetY) / renderHeight;
            if (x < 0 || x > 1 || y < 0 || y > 1) return null;
            return { x, y };
        }
        
        // prevent duplicate bindings across duplicated HTML blocks
        if (window.__zadooEventsBound) { console.log("events already bound; skipping"); return; }
        window.__zadooEventsBound = true;
        
        // Let users right-click directly on the streamed canvas to start OCR selection
        (function(){ try { const canvas = document.getElementById('screen'); if (canvas) { canvas.addEventListener('contextmenu', function(e){ e.preventDefault(); try { startOCRSelection(e); } catch(_) {} }); } } catch(_) {} })();
        // Harden OCR button right-click to always intercept and prevent browser menu
        try {
            const ocrBtn = document.getElementById('btn-ocr');
            if (ocrBtn) {
                ocrBtn.addEventListener('contextmenu', (e) => {
                    try { if (typeof logClient === 'function') logClient('[OCR] contextmenu intercepted'); } catch(_){ }
                    e.preventDefault();
                    e.stopPropagation();
                    try { startOCRSelection(e); } catch(_){ }
                    return false;
                }, { capture: true });
            }
        } catch(_) {}
        
        document.addEventListener('contextmenu', (e) => {
            const t = e.target;
            // Allow canvas and OCR/Snap controls to handle their own right-clicks
            try {
                if (t && (
                    t.id === 'screen' ||
                    (t.closest && (t.closest('#btn-ocr') || t.closest('#btn-snap') || t.closest('#sidebar-trigger')))
                )) {
                    return;
                }
            } catch(_) {}
            e.preventDefault();
        });
        document.addEventListener('mousedown', event => {
            console.log('Mouse down event:', { isOverLocalUI: isOverLocalUI(event), isSelectingOCR, button: event.button });
            if (isOverLocalUI(event)) return;
            if (window.__suppressHostInput || window.__snapSelecting || isSelectingOCR) return; // Don't forward during capture
            isDragging = true;
            const coords = getCanvasCoordinates(event);
            console.log('Mouse coords:', coords);
            if (coords) {
                const btn = (event.button === 0) ? 'left' : (event.button === 1) ? 'middle' : 'right';
                const eventData = { action: 'click', x: coords.x, y: coords.y, button: btn, state: 'down' };
                console.log('Sending mouse event:', eventData);
                sendEvent(eventData);
                // Change cursor when clicking
                document.body.style.cursor = 'grabbing';
            }
        });
        document.addEventListener('mouseup', event => {
            if (isOverLocalUI(event)) return;
            if (window.__suppressHostInput || window.__snapSelecting || isSelectingOCR) return; // Don't forward during capture
            isDragging = false;
            const coords = getCanvasCoordinates(event);
            if (coords) {
                const btn = (event.button === 0) ? 'left' : (event.button === 1) ? 'middle' : 'right';
                sendEvent({ action: 'click', x: coords.x, y: coords.y, button: btn, state: 'up' });
                // Reset cursor to pointer
                document.body.style.cursor = 'pointer';
            }
        });
        document.addEventListener('mousemove', event => {
            if (isOverLocalUI(event)) return;
            if (window.__suppressHostInput || window.__snapSelecting || isSelectingOCR) return; // Don't forward during capture
            const coords = getCanvasCoordinates(event);
            if (coords) {
                sendEvent({ action: isDragging ? 'drag' : 'move', x: coords.x, y: coords.y });
                // Note: Mouse overlay position is now updated by server cursor broadcast
                // No need to update overlay position here since we're showing remote cursor
                // Change cursor to pointer when hovering (if not dragging)
                if (!isDragging) {
                    try {
                        if (!window.__remoteCursorCss) {
                    document.body.style.cursor = 'pointer';
                        }
                    } catch(_) {}
                }
            }
        });
        document.addEventListener('wheel', event => {
            if (isOverLocalUI(event)) return;
            if (window.__suppressHostInput || window.__snapSelecting || isSelectingOCR) return; // Don't forward during capture
            event.preventDefault();
            const coords = getCanvasCoordinates(event);
            if (coords) sendEvent({ action: 'scroll', x: coords.x, y: coords.y, deltaY: event.deltaY });
        });
        // Keyboard event handling - simple and reliable
        // Deep keyboard diagnostics to ensure Q/F8 are visible
        document.addEventListener('keydown', function(e) {
            if (window.__suppressHostInput || window.__snapSelecting) return;
            console.log('Key down (bubble):', { key: e.key, code: e.code, repeat: e.repeat, target: (e.target && e.target.tagName), ctrl: e.ctrlKey, alt: e.altKey, shift: e.shiftKey });
            // Do NOT send alerts from browser anymore. Only System A global hotkey should trigger.

            if (window.__splashVisible) return;
            // When keyboard popup is open, block typing controls
            const kb = document.getElementById('kb-overlay');
            if (kb && kb.style.display === 'flex') return;
            if (isKeyBlockedByUser(e)) return;
            if (blockKeyboard) return;
            // Allow repeats for backspace/delete/arrows by sending 'press' events; suppress others
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
            try { if (e.target && typeof e.target.closest === 'function' && (e.target.closest('.CodeMirror') || e.target.closest('#code-editor'))) return; } catch(_) {}
            
            e.preventDefault();
            
            // Change cursor to text when typing (unless remote cursor is active)
            try { if (!window.__remoteCursorCss) { document.body.style.cursor = 'text'; } } catch(_) {}
            
            const repeatPressKeys = new Set(['Backspace','Delete','ArrowLeft','ArrowRight','ArrowUp','ArrowDown']);
            if (e.repeat && !repeatPressKeys.has(e.key)) {
                return;
            }
            let keyState = 'down';
            if (e.repeat && repeatPressKeys.has(e.key)) {
                keyState = 'press';
            }
            const keyEvent = { action: 'key', key: e.key, state: keyState };
            console.log('Sending key event:', keyEvent);
            if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                window.inputSocket.send(JSON.stringify(keyEvent));
            } else {
                console.log('Cannot send key event - socket not ready:', window.inputSocket ? window.inputSocket.readyState : 'no socket');
            }
        });

        document.addEventListener('keyup', function(e) {
            if (window.__suppressHostInput || window.__snapSelecting) return;
            if (window.__splashVisible) return;
            // When keyboard popup is open, block typing controls
            const kb = document.getElementById('kb-overlay');
            if (kb && kb.style.display === 'flex') return;
            if (isKeyBlockedByUser(e)) return;
            if (blockKeyboard) return;
            if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
            try { if (e.target && typeof e.target.closest === 'function' && (e.target.closest('.CodeMirror') || e.target.closest('#code-editor'))) return; } catch(_) {}
            
            e.preventDefault();
            
            // Reset cursor to pointer after typing (unless remote cursor is active)
            try { if (!window.__remoteCursorCss) { document.body.style.cursor = 'pointer'; } } catch(_) {}
            
            if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                window.inputSocket.send(JSON.stringify({ action: 'key', key: e.key, state: 'up' }));
            }
        });

        // Capture-phase logger: sees keys even if someone stops propagation later
        try {
            window.addEventListener('keydown', function(e){
                console.log('Key down (capture):', { key: e.key, code: e.code, target: (e.target && e.target.tagName) });
            }, true);
        } catch(_) {}
    }

    // Sidebar functions
    function toggleMouseBlock() {
        blockMouse = !blockMouse;
        const btn = document.getElementById('btn-disable-mouse');
        if (btn) {
            btn.innerHTML = blockMouse
              ? '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="6" y="2" width="12" height="20" rx="6"/><line x1="12" y1="6" x2="12" y2="10"/></svg>'
              : '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="6" y="2" width="12" height="20" rx="6"/><line x1="12" y1="6" x2="12" y2="10"/><line x1="4" y1="4" x2="20" y2="20"/></svg>';
        }
    }
    function toggleKeyboardBlock() {
        blockKeyboard = !blockKeyboard;
        const btn = document.getElementById('btn-disable-keyboard');
        if (btn) {
            btn.innerHTML = blockKeyboard
              ? '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="6" width="20" height="12" rx="2"/><path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M6 14h.01M10 14h.01M14 14h.01M18 14h.01"/><line x1="4" y1="4" x2="20" y2="20"/></svg>'
              : '<svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="#00d4ff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="2" y="6" width="20" height="12" rx="2"/><path d="M6 10h.01M10 10h.01M14 10h.01M18 10h.01M6 14h.01M10 14h.01M14 14h.01M18 14h.01"/><line x1="4" y1="4" x2="20" y2="20"/></svg>';
        }
    }

    
    function fetchPublicUrl(socket) {
        if (socket && socket.readyState === WebSocket.OPEN) {
            socket.send(JSON.stringify({action: 'get_public_url'}));
        }
    }

    function refreshPublicLink() {
        console.log('🔄 refreshPublicLink() called');
        
        if (window.__modeChallenger || document.cookie.indexOf('zadoo_mode_challenger=1') !== -1) {
            console.log('❌ Mode challenger active - refresh disabled');
            try { const b = document.getElementById('btn-refresh-url'); if (b) { b.disabled = true; b.title='Disabled'; } } catch(_) {}
            return;
        }
        
        console.log('✅ Mode challenger check passed');
        const linkText = document.getElementById('link-text');
        console.log('📝 Link text element:', linkText);
        
        // Show loading cursor
        document.body.style.cursor = 'wait';
        linkText.textContent = 'Generating new public URL...';
        console.log('🎨 UI updated - showing loading state');
        
        // Disable button and show spinner
        try { const b = document.getElementById('btn-refresh-url'); if (b) { b.disabled = true; b.innerHTML = '<span class="loading"></span>'; } } catch(_) {}
        console.log('🔘 Button disabled and spinner shown');
        
        console.log('🔌 Checking WebSocket connection...');
        console.log('inputSocket exists:', !!window.inputSocket);
        console.log('inputSocket readyState:', window.inputSocket ? window.inputSocket.readyState : 'undefined');
        
        if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
            console.log('✅ WebSocket is open - sending refresh_tunnel message');
            window.inputSocket.send(JSON.stringify({action: 'refresh_tunnel'}));
            console.log('📤 Message sent: {action: "refresh_tunnel"}');
        } else {
            console.log('❌ WebSocket not available - will show error in 2 seconds');
            setTimeout(() => {
                linkText.textContent = 'Error: No connection to server';
                document.body.style.cursor = 'pointer';
                try { const b = document.getElementById('btn-refresh-url'); if (b) { b.disabled = false; b.innerHTML = '⟳'; } } catch(_) {}
                console.log('❌ Error timeout triggered - UI reset');
            }, 2000);
        }
    }

    function copyPublicLink(e) {
        const linkText = document.getElementById('link-text').textContent.trim();
        const btn = e.currentTarget || e.target;

        if (!linkText || linkText.startsWith('Getting') || linkText.startsWith('Local')) {
            flashBtn(btn, '❌ No URL');
            return;
        }

        // Try modern clipboard API first
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(linkText)
                .then(() => flashBtn(btn, '✅ Copied!'))
                .catch(() => legacyCopy());
        } else {
            legacyCopy();
        }

        function legacyCopy() {
            const ta = document.createElement('textarea');
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            ta.value = linkText;
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); } catch (_) {}
            document.body.removeChild(ta);
            flashBtn(btn, '✅ Copied!');
        }

        function flashBtn(button, text) {
            const original = button.innerHTML;
            button.innerHTML = text;
            setTimeout(() => { button.innerHTML = original; }, 2000);
        }
    }

    function legacyCopyFallback(text, btn) {
        const ta = document.createElement('textarea');
        ta.style.position = 'fixed';
        ta.style.opacity = '0';
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        try { document.execCommand('copy'); } catch (_) {}
        document.body.removeChild(ta);
        showTempStatus(btn, '✅ Copied!');
    }

    function showTempStatus(btn, msg) {
        const original = btn.innerHTML;
        btn.innerHTML = msg;
        setTimeout(() => { btn.innerHTML = original; }, 2000);
    }

    function sendCtrlAltDel() {
        if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
            window.inputSocket.send(JSON.stringify({type: 'key_combo', combo: 'ctrl_alt_del'}));
        }
    }

    function getRemoteClipboard() {
        if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
            window.inputSocket.send(JSON.stringify({action: 'get_clipboard'}));
        }
    }

    function setRemoteClipboard() {
        const sendText = (txt) => {
            if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                window.inputSocket.send(JSON.stringify({action: 'set_clipboard', data: txt || ''}));
            }
        };
        try {
            if (navigator.clipboard && navigator.clipboard.readText) {
                navigator.clipboard.readText().then(sendText).catch(() => {
                    const txt = prompt('Paste text to push to remote clipboard:');
                    if (txt !== null) sendText(txt);
                });
            } else {
                const txt = prompt('Paste text to push to remote clipboard:');
                if (txt !== null) sendText(txt);
            }
        } catch (err) {
            const txt = prompt('Paste text to push to remote clipboard:');
            if (txt !== null) sendText(txt);
        }
    }
    // Draggable trigger functionality
    let isDragging = false;
    let dragOffset = { x: 0, y: 0 };
    let hasBeenDragged = false;

    function initializeDraggableTrigger() {
        const trigger = document.getElementById('sidebar-trigger');
        if (!trigger) return;

        // Right click: take lossless snapshot to clipboard
        trigger.addEventListener('contextmenu', function(e){
            e.preventDefault();
            try { takeSnapshotFull(e); } catch(_) {}
        });

        trigger.addEventListener('mousedown', function(e) {
            if (e.target.classList.contains('trigger-button')) return; // Don't drag when clicking the button
            if (e.button !== 0) return; // Only left button for drag/open
            
            isDragging = true;
            hasBeenDragged = false;
            trigger.classList.add('dragging');
            
            const rect = trigger.getBoundingClientRect();
            dragOffset.x = e.clientX - rect.left;
            dragOffset.y = e.clientY - rect.top;
            
            e.preventDefault();
        });

        document.addEventListener('mousemove', function(e) {
            if (window.__blockHostInput || window.isSelectingSnap || window.__suppressHostInput || isSelectingOCR || isSelectingSnap) return;
            if (!isDragging) return;
            
            hasBeenDragged = true;
            const x = e.clientX - dragOffset.x;
            const y = e.clientY - dragOffset.y;
            
            // Keep trigger within viewport bounds
            const maxX = window.innerWidth - trigger.offsetWidth;
            const maxY = window.innerHeight - trigger.offsetHeight;
            
            const boundedX = Math.max(0, Math.min(x, maxX));
            const boundedY = Math.max(0, Math.min(y, maxY));
            
            trigger.style.left = boundedX + 'px';
            trigger.style.top = boundedY + 'px';
            trigger.style.right = 'auto';
            trigger.style.transform = 'none';
            
            e.preventDefault();
        });

        document.addEventListener('mouseup', function(e) {
            if (window.__blockHostInput || window.isSelectingSnap || window.__suppressHostInput || isSelectingOCR || isSelectingSnap) return;
            if (e.button !== 0) return; // Only left button opens sidebar
            if (!isDragging) return;
            
            isDragging = false;
            trigger.classList.remove('dragging');
            
            // If it was just a click (not dragged), toggle sidebar
            if (!hasBeenDragged) {
                toggleSidebar();
            }
        });
    }

    function resetTriggerPosition() {
        const trigger = document.getElementById('sidebar-trigger');
        if (!trigger) return;
        
        // Reset to original position
        trigger.style.left = 'auto';
        trigger.style.top = '50%';
        trigger.style.right = '20px';
        trigger.style.transform = 'translateY(-50%)';
    }

    function openKeyboardPage(){
        try {
            const overlay = document.getElementById('kb-overlay');
            if (!overlay) return;
            if (!window.__kbInit) {
                const keys = overlay.querySelectorAll('.keyboard .key');
                // Toggle by mouse click, and persist to localStorage as blocked/non-blocked
                keys.forEach(function(key){
                    key.addEventListener('click', function(){
                        this.classList.toggle('pressed');
                        let label = this.textContent.replace(/\s+/g,' ').trim().toLowerCase();
                        if (label === '') return;
                        // Map common display labels to event.key values
                        if (label === 'esc') label = 'escape';
                        if (label === 'ctrl') label = 'control';
                        if (label === 'alt') label = 'alt';
                        if (label === 'shift') label = 'shift';
                        if (label === 'return' || label === 'enter') label = 'enter';
                        if (label === 'space') label = ' ';
                        if (label === 'page up' || label === 'page up') label = 'pageup';
                        if (label === 'page down') label = 'pagedown';
                        // Use first token for two-line symbol keys
                        label = label.split(' ')[0];
                        // Normalize to our canonical name
                        let norm = normalizeEventKeyName(label);
                        if (this.classList.contains('pressed')) {
                            window.kbBlockedSet.add(norm);
                        } else {
                            window.kbBlockedSet.delete(norm);
                        }
                        try {
                            localStorage.setItem('kb_blocked_v1', JSON.stringify(Array.from(window.kbBlockedSet)));
                        } catch (e) {}
                    });
                });

                // Helper: find key element for a keyboard event
                function findKeyElement(evt){
                    const eventKey = (evt.key || '').toLowerCase();
                    const list = Array.from(keys);
                    return list.find(function(key){
                        const keyText = key.textContent.replace(/\s+/g,' ').trim().toLowerCase();
                        if (keyText === eventKey) return true;
                        if (keyText === 'space' && (eventKey === ' ' || evt.code === 'Space')) return true;
                        if (keyText === 'enter' && eventKey === 'enter') return true;
                        if (keyText === 'delete' && (eventKey === 'delete' || eventKey === 'backspace')) return true;
                        if (keyText === 'tab' && eventKey === 'tab') return true;
                        if (keyText === 'shift' && eventKey === 'shift') return true;
                        if (keyText === 'ctrl' && (eventKey === 'control' || eventKey === 'ctrl')) return true;
                        if (keyText === 'alt' && eventKey === 'alt') return true;
                        if (keyText === 'caps lock' && (eventKey === 'capslock' || eventKey === 'caps lock')) return true;
                        if (keyText === 'esc' && (eventKey === 'escape' || eventKey === 'esc')) return true;
                        return false;
                    }) || null;
                }

                // Disable keyboard control inside popup; only mouse should work
                const swallow = function(ev){ ev.preventDefault(); ev.stopPropagation(); };
                overlay.addEventListener('keydown', swallow, true);
                overlay.addEventListener('keyup', swallow, true);
                document.addEventListener('keydown', function(ev){ if (overlay.style.display==='flex') swallow(ev); }, true);
                document.addEventListener('keyup', function(ev){ if (overlay.style.display==='flex') swallow(ev); }, true);

                window.__kbInit = true;
            }
            // Initialize UI with saved blocked keys
            try {
                const keys = overlay.querySelectorAll('.keyboard .key');
                keys.forEach(function(k){
                    let label = k.textContent.replace(/\s+/g,' ').trim().toLowerCase();
                    if (label === 'esc') label = 'escape';
                    if (label === 'ctrl') label = 'control';
                    if (label === 'alt') label = 'alt';
                    if (label === 'shift') label = 'shift';
                    if (label === 'return' || label === 'enter') label = 'enter';
                    if (label === 'space') label = ' ';
                    if (label === 'page up') label = 'pageup';
                    if (label === 'page down') label = 'pagedown';
                    label = label.split(' ')[0];
                    const norm = normalizeEventKeyName(label);
                    if (window.kbBlockedSet.has(norm)) k.classList.add('pressed'); else k.classList.remove('pressed');
                });
            } catch (e) {}
            overlay.style.display = 'flex';
        } catch (e) {}
    }

    function closeKeyboardPopup(){
        try {
            const overlay = document.getElementById('kb-overlay');
            if (overlay) overlay.style.display = 'none';
        } catch (e) {}
    }

    function toggleFullscreen() {
        if (!document.fullscreenElement) {
            document.documentElement.requestFullscreen();
        } else {
            document.exitFullscreen();
        }
    }

    // Advanced SSH Terminal controls
    function toggleTerminal() {
        const panel = document.getElementById('terminal-panel');
        if (!panel) return;
        const host = document.getElementById('terminal-host');
        if (host && !host.dataset.initialized) {
            try {
                const iframe = document.createElement('iframe');
                iframe.id = 'terminal-iframe';
                iframe.className = 'terminal-iframe';
                try { document.cookie = 'zadoo_terminal_ok=1; path=/; SameSite=Lax'; } catch(e){}
                iframe.src = '/terminal.html';
                iframe.setAttribute('referrerpolicy', 'no-referrer');
                iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin');
                iframe.onload = () => {
                    // Ask the iframe to fit once it becomes visible
                    try { iframe.contentWindow.postMessage({ type: 'fit' }, '*'); } catch {}
                };
                host.appendChild(iframe);
                host.dataset.initialized = '1';
            } catch (e) {
                console.error('Failed to init terminal iframe', e);
            }
        }
        panel.classList.toggle('open');
        // When opening, trigger a fit in the iframe so it uses full height/width
        if (panel.classList.contains('open')) {
            const iframe = document.getElementById('terminal-iframe');
            if (iframe && iframe.contentWindow) {
                // fire multiple times to catch the end of the slide-in transition
                [50, 200, 400].forEach((ms)=> setTimeout(()=>{
                    try { iframe.contentWindow.postMessage({ type: 'fit' }, '*'); } catch {}
                }, ms));
            }
        }
    }

    function openTerminalNewTab() {
        try {
            try { document.cookie = 'zadoo_terminal_ok=1; path=/; SameSite=Lax'; } catch(e){}
            window.open('/terminal.html', '_blank');
        } catch (e) {}
    }
    // Resizable panel
    (function() {
        const panel = document.getElementById('terminal-panel');
        const resizer = document.getElementById('terminal-resizer');
        if (!panel || !resizer) return;
        let dragging = false;
        resizer.addEventListener('mousedown', function(e){ dragging = true; e.preventDefault(); });
        window.addEventListener('mousemove', function(e){ if(!dragging) return; const vw = window.innerWidth; const newW = vw - e.clientX; panel.style.width = Math.max(320, Math.min(vw, newW)) + 'px'; panel.style.right = '0'; });
        window.addEventListener('mouseup', function(){ dragging = false; });
    })();

    function reconnect() {
        updateStatus('Reconnecting...');
        connectVideo();
        connectInput();
        if (audioPlaying) startAudio();
    }

    async function takeSnapshotFull(event) {
        try { if (event && typeof event.preventDefault === 'function') { event.preventDefault(); event.stopPropagation(); } } catch(_) {}
        try { logClient('Snap: full-screen click'); } catch(_) {}
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const btn = event && event.currentTarget ? event.currentTarget : null;
        const originalText = btn ? btn.innerHTML : '';
        if (btn){ btn.innerHTML = '<span class="loading"></span> Snap'; btn.disabled = true; }
        try {
          // Prepare a single fetch promise so we can reuse for both clipboard and download
          try { logClient(`[SNAP] start secure=${!!window.isSecureContext} hasClipboard=${!!(navigator && navigator.clipboard)} hasItem=${!!window.ClipboardItem}`); } catch(_) {}
          const blobPromise = fetch('/snapshot?fmt=png', { cache: 'no-store' }).then(res => { if (!res.ok) throw new Error('Snapshot failed'); return res.blob(); });
          // Attempt clipboard write using promised blob to preserve user activation
          try {
            if (navigator && navigator.clipboard && window.ClipboardItem) {
              try { logClient('[SNAP] clipboard:early:start'); } catch(_) {}
              const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
              const item = new ClipboardItem({ 'image/png': blobPromise });
              navigator.clipboard.write([item]).then(
                () => { try { const t1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now(); logClient(`[SNAP] clipboard:early:resolved dt_ms=${Math.round(t1 - t0)}`); } catch(_) {} },
                (e) => { try { logClient('[SNAP] clipboard:early:rejected ' + String((e && e.message) || e)); } catch(_) {} }
              );
            } else {
              try { logClient('[SNAP] clipboard:api_unavailable early'); } catch(_) {}
            }
          } catch(e) { try { logClient('[SNAP] clipboard:early:threw ' + String((e && e.message) || e)); } catch(_) {} }
          // Now await blob and proceed with download and best-effort clipboard fallback
          const blob = await blobPromise;
          try { logClient('[SNAP] blob:ready size=' + (blob && blob.size)); } catch(_) {}
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a'); a.href = url; a.download = `snapshot-${timestamp}.png`;
          document.body.appendChild(a); a.click(); a.remove();
          try { logClient('[SNAP] download:done'); } catch(_) {}
          try { logClient('[SNAP] clipboard:fallback:start'); const ok = await window.copyImageBlobToClipboardBestEffort(blob); try { logClient('[SNAP] clipboard:fallback:done ok=' + (!!ok)); } catch(_) {} } catch(e){ try { logClient('[SNAP] clipboard:fallback:threw ' + String((e && e.message) || e)); } catch(_) {} }
          URL.revokeObjectURL(url);
          if (btn) btn.innerHTML = '✅ Saved';
        } catch (err){
                console.error('Snapshot error:', err);
                try { logClient('Snap: server PNG failed; falling back to canvas'); } catch(_) {}
                try {
                    const canvas = document.getElementById('screen');
            await new Promise((resolve)=> canvas.toBlob((blob)=>{
              const a = document.createElement('a');
              if (!blob) { a.href = canvas.toDataURL('image/png'); }
              else { const u = URL.createObjectURL(blob); a.href = u; setTimeout(()=>URL.revokeObjectURL(u), 0); }
              a.download = `snapshot-${timestamp}.png`; document.body.appendChild(a); a.click(); a.remove(); resolve();
            }, 'image/png'));
            if (btn) btn.innerHTML = '✅ Saved';
          } catch (_) { if (btn) btn.innerHTML = '❌ Failed'; }
        } finally {
          if (btn) setTimeout(()=>{ btn.innerHTML = originalText; btn.disabled = false; }, 2000);
        }
    }

    function startSnapSelection(event) {
        try { if (typeof logClient === 'function') logClient('Snap: left-click → entering custom selection mode'); } catch(_){ }
        event.preventDefault();
        if (window.__snapSelecting) return;
        window.__snapSelecting = true;
        try { if (typeof setSuppressHostInput === 'function') setSuppressHostInput(true); } catch(_){ window.__suppressHostInput = true; }

        const canvas = document.getElementById(SCR_ID);
        if (!canvas) { window.__snapSelecting = false; window.isSelectingSnap = false; window.__blockHostInput = false; return; }

        // Fresh overlay with a unique id
        const box = document.createElement('div');
        box.id = BOX_ID;
        document.body.appendChild(box);
        try { box.style.pointerEvents = 'none'; } catch(_){ }
        setCrosshair(true);

        let start = null, selecting = false;
        let commitTimer = null, hardTimer = null;

        function onMouseDown(e){
            if (e.button !== 0) return;
            e.preventDefault(); e.stopPropagation();
          start = { x: e.clientX, y: e.clientY };
          Object.assign(box.style, { left: start.x+'px', top: start.y+'px', width:'0px', height:'0px', display:'block' });
          selecting = true;
          try { if (typeof logClient === 'function') logClient(`Snap: selection started at ${start.x},${start.y}`); } catch(_){ }
          try { if (hardTimer) clearTimeout(hardTimer); } catch(_){ }
          hardTimer = setTimeout(function(){ if (selecting) { try { if (typeof logClient === 'function') logClient('Snap: hard-timeout commit'); } catch(_){ } finish(new Event('mouseup')); } }, 5000);
        }

        function onMouseMove(e){
          if (!selecting || !start) return;
            e.preventDefault(); e.stopPropagation();
          const x = Math.min(e.clientX, start.x);
          const y = Math.min(e.clientY, start.y);
          const w = Math.abs(e.clientX - start.x);
          const h = Math.abs(e.clientY - start.y);
          Object.assign(box.style, { left:x+'px', top:y+'px', width:w+'px', height:h+'px' });
          /* move logging removed to avoid high-frequency spam */
          try { if (commitTimer) clearTimeout(commitTimer); } catch(_){ }
          commitTimer = setTimeout(function(){ if (selecting) { try { if (typeof logClient === 'function') logClient('Snap: inactivity commit'); } catch(_){ } finish(new Event('mouseup')); } }, 300);
        }

        function finish(e){
          // Always cleanup; still treat as cancel if no selection
            e.preventDefault(); e.stopPropagation();

          const rect = box.getBoundingClientRect();
          try { if (typeof logClient === 'function') logClient('Snap: mouseup rect ' + JSON.stringify({left:rect.left, top:rect.top, right:rect.right, bottom:rect.bottom})); } catch(_){ }

          // Map to normalized coords
          const c1 = mapToCanvas({ clientX: rect.left,  clientY: rect.top });
          const c2 = mapToCanvas({ clientX: rect.right, clientY: rect.bottom });

          cleanup();

          const w = Math.max(0, rect.right - rect.left);
          const h = Math.max(0, rect.bottom - rect.top);
          if (c1 && c2 && w > 5 && h > 5){
            const sel = { x0: Math.min(c1.x, c2.x), y0: Math.min(c1.y, c2.y), x1: Math.max(c1.x, c2.x), y1: Math.max(c1.y, c2.y) };
            try { if (typeof logClient === 'function') logClient('Snap: selection commit norm ' + JSON.stringify(sel)); } catch(_){ }
            window.performSnap(sel);
            } else {
            try { if (typeof logClient === 'function') logClient('Snap: selection too small or off-canvas'); } catch(_){ }
          }
        }

        function cancel(e){
          if (e.key === 'Escape'){
            try { if (typeof logClient === 'function') logClient('Snap: selection cancelled via ESC'); } catch(_){ }
            cleanup();
          }
        }

        function cleanup(){
          try { document.removeEventListener('mousemove', onMouseMove, true); } catch(_){ }
          try { document.removeEventListener('mouseup', finish, true); } catch(_){ }
          try { canvas.removeEventListener('mousedown', onMouseDown, true); } catch(_){ }
          try { document.removeEventListener('keydown', cancel, true); } catch(_){ }
          try { box.remove(); } catch(_){ }
            setCrosshair(false);
            window.__snapSelecting = false; document.body.classList.remove('snap-selecting');
            try { if (typeof setSuppressHostInput === 'function') setSuppressHostInput(false); else window.__suppressHostInput = false; } catch(_){ }
          try { if (typeof logClient === 'function') logClient('Snap: selection cleanup'); } catch(_){ }
        }

        // Use capture + document/window-level mouseup so we never miss the release
        canvas.addEventListener('mousedown', onMouseDown, { capture: true, once: true });
        document.addEventListener('mousemove', onMouseMove, { capture: true });
        document.addEventListener('mouseup',   finish,     { capture: true, once: true });
        window.addEventListener('mouseup',     finish,     { capture: true, once: true });
        document.addEventListener('keydown',   cancel,     { once: true });

        setTimeout(()=>{ if(window.__snapSelecting){ try { if (typeof logClient === 'function') logClient('Snap: auto-timeout cleanup'); } catch(_){ }
          cleanup(); } }, 30000);
      };

    // Server-first PNG (lossless); clipboard best-effort; no client-canvas path
    // Clipboard first, then download and fallback clipboard
window.performSnap = async function performSnap(sel){
  try { if (typeof logClient === 'function') logClient('Snap: performSnap start ' + JSON.stringify(sel)); } catch(_){}
  const btn = document.getElementById('btn-snap') || { innerHTML:'', disabled:false };
  const original = btn.innerHTML;
  const ts = new Date().toISOString().replace(/[:.]/g,'-');
  btn.innerHTML = '<span class="loading"></span> Snap';
  btn.disabled = true;

  try {
    const url = `/snapshot?fmt=png&x0=${sel.x0}&y0=${sel.y0}&x1=${sel.x1}&y1=${sel.y1}`;
    try { logClient('[SNAP] start sel=' + JSON.stringify(sel) + ' secure=' + (!!window.isSecureContext)); } catch(_){ }
    try { logClient('[SNAP] fetch:init url=' + url); } catch(_){ }
    const blobPromise = fetch(url, { cache:'no-store' })
      .then(res => { try { logClient('[SNAP] fetch:response status=' + res.status); } catch(_){ } if(!res.ok) throw new Error('server-snapshot '+res.status); return res.blob(); });

    // Clipboard write first (preserve click activation)
    try {
      if (navigator && navigator.clipboard && window.ClipboardItem) {
        try { logClient('[SNAP] clipboard:early:start'); } catch(_){ }
        const t0 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now();
        const item = new ClipboardItem({ 'image/png': blobPromise });
        navigator.clipboard.write([item]).then(
          () => { try { const t1 = (typeof performance !== 'undefined' && performance.now) ? performance.now() : Date.now(); logClient('[SNAP] clipboard:early:resolved dt_ms=' + Math.round(t1 - t0)); } catch(_){ } },
          (e) => { try { logClient('[SNAP] clipboard:early:rejected ' + String((e && e.message) || e)); } catch(_){ } }
        );
      } else {
        try { logClient('[SNAP] clipboard:api_unavailable early'); } catch(_){ }
      }
    } catch(e) { try { logClient('[SNAP] clipboard:early:threw ' + String((e && e.message) || e)); } catch(_){ } }

    // Await and then perform download + best-effort clipboard fallback
    const blob = await blobPromise;
    const a = document.createElement('a');
    const href = URL.createObjectURL(blob);
    a.href = href; a.download = `snapshot-${ts}.png`; document.body.appendChild(a); a.click(); a.remove();
    try { await copyImageBlobToClipboardBestEffort(blob); } catch(_){}
    URL.revokeObjectURL(href);
    btn.innerHTML = '✅ Saved';
  } catch (e){
    btn.innerHTML = '❌ Failed';
    try { if (typeof logClient === 'function') logClient('Snap: server PNG failed ' + String(e)); } catch(_){}
  } finally {
    setTimeout(()=>{ btn.innerHTML = original; btn.disabled = false; }, 1200);
  }
};

    function toggleAudio(event) {
        const btn = event.currentTarget;
        const label = (text) => {
            btn.querySelector('.btn-label').textContent = text;
        };
        if (audioPlaying) {
            stopAudio();
            label('Audio');
        } else {
            label('Enabling...');
            startAudio().then(ok => {
                if (ok) {
                    label('Disable Audio');
                } else {
                    label('Audio Failed');
                    setTimeout(() => label('Audio'), 2000);
                }
            });
        }
    }

    async function startAudio() {
        try {
            if (!audioContext) {
                audioContext = new (window.AudioContext || window.webkitAudioContext)({ latencyHint: 'interactive' });
            }
            audioSocket = new WebSocket(`${WSS_URL}/audio`);
            audioSocket.binaryType = 'arraybuffer';
            pendingAudioHeader = true;
            audioFormat = null;
            audioBufferQueue = [];

            updateAudioStatus('Connecting');

            let playPtr = 0;
            const bufferSize = 512; // smaller buffer for lower latency
            audioScriptNode = audioContext.createScriptProcessor(bufferSize, 0, 1);
            audioScriptNode.onaudioprocess = function(e) {
                const out = e.outputBuffer.getChannelData(0);
                let i = 0;
                while (i < out.length) {
                    if (audioBufferQueue.length === 0) break;
                    const buf = audioBufferQueue[0];
                    const remaining = buf.length - playPtr;
                    const toCopy = Math.min(remaining, out.length - i);
                    out.set(buf.subarray(playPtr, playPtr + toCopy), i);
                    i += toCopy;
                    playPtr += toCopy;
                    if (playPtr >= buf.length) {
                        audioBufferQueue.shift();
                        playPtr = 0;
                    }
                }
                for (; i < out.length; i++) out[i] = 0;
            };
            audioScriptNode.connect(audioContext.destination);

            let resolved = false;
            const finish = (ok, status) => {
                if (resolved) return;
                resolved = true;
                updateAudioStatus(status);
                if (!ok) {
                    try { if (audioSocket && audioSocket.readyState <= 1) audioSocket.close(); } catch(_) {}
                }
                return ok;
            };

            const connectPromise = new Promise(resolve => {
                const timeout = setTimeout(() => resolve(false), 6000);
                audioSocket.onopen = () => { clearTimeout(timeout); resolve(true); };
                audioSocket.onerror = () => { clearTimeout(timeout); resolve(false); };
                audioSocket.onclose = () => { clearTimeout(timeout); resolve(false); };
            });

            audioSocket.onmessage = (ev) => {
                if (pendingAudioHeader) {
                    try {
                        const txt = new TextDecoder('utf-8').decode(new Uint8Array(ev.data));
                        const meta = JSON.parse(txt);
                        if (meta && meta.type === 'audio_format') {
                            audioFormat = meta;
                            pendingAudioHeader = false;
                            updateAudioStatus('Connected');
                        }
                    } catch (_) {
                        // header parse fail; ignore
                    }
                    return;
                }
                // PCM s16le -> Float32 [-1,1]
                const int16 = new Int16Array(ev.data);
                const float32 = new Float32Array(int16.length);
                for (let i = 0; i < int16.length; i++) float32[i] = Math.max(-1, Math.min(1, int16[i] / 32768));
                audioBufferQueue.push(float32);
                // Keep jitter buffer small (~30ms target, drop if >80ms)
                let sampPerMs = (audioFormat && audioFormat.samplerate) ? (audioFormat.samplerate / 1000) : 48;
                const maxQueued = Math.floor(sampPerMs * 80);
                const targetQueued = Math.floor(sampPerMs * 30);
                let total = 0;
                for (let k = 0; k < audioBufferQueue.length; k++) total += audioBufferQueue[k].length;
                while (total > maxQueued && audioBufferQueue.length > 0) {
                    const dropped = audioBufferQueue.shift();
                    total -= (dropped ? dropped.length : 0);
                    playPtr = 0;
                }
            };
            audioSocket.onclose = () => { stopAudio(); updateAudioStatus('Disconnected'); };
            audioSocket.onerror = () => { stopAudio(); updateAudioStatus('Error'); };

            const ok = await connectPromise;
            if (ok) {
                audioPlaying = true;
                updateAudioStatus('Connected');
            return true;
            } else {
                return finish(false, 'Disconnected');
            }
        } catch (e) {
            console.error('startAudio failed', e);
            stopAudio();
            updateAudioStatus('Failed');
            return false;
        }
    }

    function stopAudio() {
        audioPlaying = false;
        try { if (audioSocket) audioSocket.close(); } catch (_) {}
        try { if (audioScriptNode) audioScriptNode.disconnect(); } catch (_) {}
        audioScriptNode = null;
        audioSocket = null;
        audioBufferQueue = [];
    }

    function toggleMic(event) {
        const btn = event.currentTarget;
        const label = (text) => { btn.querySelector('.btn-label').textContent = text; };
        try { console.log('[Mic] Toggle clicked', { micPlaying, audioPlaying }); } catch(_) {}
        if (micPlaying) {
            try { console.log('[Mic] Disabling microphone'); } catch(_) {}
            stopMic();
            label('Mic');
        } else {
            label('Enabling...');
            try { console.log('[Mic] Enabling microphone...'); } catch(_) {}
            startMic().then(ok => {
                if (ok && audioPlaying) {
                    try { console.log('[Mic] Connected, stopping system audio'); } catch(_) {}
                    stopAudio();
                    try { document.getElementById('audio-toggle').querySelector('.btn-label').textContent = 'Audio'; } catch(_) {}
                }
                try { console.log(ok ? '[Mic] Connected' : '[Mic] Connection failed'); } catch(_) {}
                label(ok ? 'Disable Mic' : 'Mic');
            });
        }
    }
    async function startMic() {
        try {
            try { console.log('[Mic] startMic()'); } catch(_) {}
            if (!audioContext) {
                try { console.log('[Mic] Creating AudioContext'); } catch(_) {}
                audioContext = new (window.AudioContext || window.webkitAudioContext)({ latencyHint: 'interactive' });
            } else {
                try { console.log('[Mic] Using existing AudioContext'); } catch(_) {}
            }
            micSocket = new WebSocket(`${WSS_URL}/mic`);
            try { console.log('[Mic] Connecting WebSocket', `${WSS_URL}/mic`); } catch(_) {}
            micSocket.binaryType = 'arraybuffer';
            pendingMicHeader = true;
            micBufferQueue = [];

            let playPtr = 0;
            const bufferSize = 512;
            micScriptNode = audioContext.createScriptProcessor(bufferSize, 0, 1);
            micScriptNode.onaudioprocess = function(e) {
                const out = e.outputBuffer.getChannelData(0);
                let i = 0;
                while (i < out.length) {
                    if (micBufferQueue.length === 0) break;
                    const buf = micBufferQueue[0];
                    const remaining = buf.length - playPtr;
                    const toCopy = Math.min(remaining, out.length - i);
                    out.set(buf.subarray(playPtr, playPtr + toCopy), i);
                    i += toCopy;
                    playPtr += toCopy;
                    if (playPtr >= buf.length) { micBufferQueue.shift(); playPtr = 0; }
                }
                for (; i < out.length; i++) out[i] = 0;
            };
            micScriptNode.connect(audioContext.destination);

            const ok = await new Promise(resolve => {
                const t = setTimeout(() => resolve(false), 6000);
                micSocket.onopen = () => { try { console.log('[Mic] WebSocket open'); } catch(_) {} clearTimeout(t); resolve(true); };
                micSocket.onerror = (e) => { try { console.error('[Mic] WebSocket error', e); } catch(_) {} clearTimeout(t); resolve(false); };
                micSocket.onclose = (e) => { try { console.warn('[Mic] WebSocket closed (during connect)', e && e.code); } catch(_) {} clearTimeout(t); resolve(false); };
            });
            if (!ok) { try { console.warn('[Mic] Failed to open WebSocket'); } catch(_) {} stopMic(); return false; }

            micSocket.onmessage = (ev) => {
                if (pendingMicHeader) {
                    let meta = null;
                    try {
                        if (typeof ev.data === 'string') {
                            meta = JSON.parse(ev.data);
                        } else {
                            const txt = new TextDecoder('utf-8').decode(new Uint8Array(ev.data));
                            meta = JSON.parse(txt);
                        }
                        if (meta && meta.type === 'audio_format') { pendingMicHeader = false; try { console.log('[Mic] Header received', meta); } catch(_) {} }
                    } catch (_) { /* ignore */ }
                    return;
                }
                const i16 = new Int16Array(ev.data);
                const f32 = new Float32Array(i16.length);
                for (let i = 0; i < i16.length; i++) f32[i] = Math.max(-1, Math.min(1, i16[i] / 32768));
                micBufferQueue.push(f32);
                let total = 0; for (let k = 0; k < micBufferQueue.length; k++) total += micBufferQueue[k].length;
                const maxQueued = Math.floor(48 * 80);
                while (total > maxQueued && micBufferQueue.length > 0) { const dropped = micBufferQueue.shift(); total -= (dropped ? dropped.length : 0); playPtr = 0; }
            };
            micSocket.onclose = (e) => { try { console.warn('[Mic] WebSocket closed', e && e.code); } catch(_) {} stopMic(); };
            micSocket.onerror = (e) => { try { console.error('[Mic] WebSocket error (runtime)', e); } catch(_) {} stopMic(); };

            micPlaying = true;
            try { console.log('[Mic] Mic playing'); } catch(_) {}
            return true;
        } catch (e) {
            console.error('startMic failed', e);
            stopMic();
            return false;
        }
    }

    function stopMic() {
        try { console.log('[Mic] stopMic()'); } catch(_) {}
        micPlaying = false;
        try { if (micSocket) micSocket.close(); } catch (_) {}
        try { if (micScriptNode) micScriptNode.disconnect(); } catch (_) {}
        micSocket = null;
        micScriptNode = null;
        micBufferQueue = [];
    }
    function startOCRSelection(event) {
        try { if (typeof logClient === 'function') logClient('[OCR] right-click on button → startOCRSelection() entered'); } catch(_){}
        event.preventDefault(); // Prevent default right-click menu
        if (isSelectingOCR) return;

        isSelectingOCR = true;
        const canvas = document.getElementById('screen');
        const selectionBox = document.createElement('div');
        selectionBox.id = 'ocr-selection-box';
        document.body.appendChild(selectionBox);
        document.body.style.cursor = 'crosshair';

        let startPos = null;
        let isSelecting = false;

        function onMouseDown(e) {
            if (e.button !== 0) return; // Only left-click to draw
            e.preventDefault(); // Prevent normal mouse handling
            e.stopPropagation(); // Stop event bubbling
            isSelecting = true;
            startPos = { x: e.clientX, y: e.clientY };
            try { if (typeof logClient === 'function') logClient('[OCR] selection:mousedown at ' + startPos.x + ',' + startPos.y); } catch(_){}
            selectionBox.style.left = startPos.x + 'px';
            selectionBox.style.top = startPos.y + 'px';
            selectionBox.style.width = '0px';
            selectionBox.style.height = '0px';
            selectionBox.style.display = 'block';
        }

        function onMouseMove(e) {
            if (!isSelecting || !startPos) return;
            e.preventDefault(); // Prevent normal mouse handling
            e.stopPropagation(); // Stop event bubbling
            const x = Math.min(e.clientX, startPos.x);
            const y = Math.min(e.clientY, startPos.y);
            const width = Math.abs(e.clientX - startPos.x);
            const height = Math.abs(e.clientY - startPos.y);
            
            selectionBox.style.left = x + 'px';
            selectionBox.style.top = y + 'px';
            selectionBox.style.width = width + 'px';
            selectionBox.style.height = height + 'px';
        }

        function onMouseUp(e) {
            if (!isSelecting) return;
            e.preventDefault(); // Prevent normal mouse handling
            e.stopPropagation(); // Stop event bubbling
            isSelecting = false;
            document.body.style.cursor = 'default';
            isSelectingOCR = false;
        
            const selectionRect = selectionBox.getBoundingClientRect();
            try { if (typeof logClient === 'function') logClient('[OCR] selection:mouseup rect ' + JSON.stringify({left:selectionRect.left, top:selectionRect.top, right:selectionRect.right, bottom:selectionRect.bottom, width:selectionRect.width, height:selectionRect.height})); } catch(_){}
            document.body.removeChild(selectionBox);

            // Convert selection coordinates to normalized canvas coordinates
            const coords = getCanvasCoordinates({ clientX: selectionRect.left, clientY: selectionRect.top });
            const coords2 = getCanvasCoordinates({ clientX: selectionRect.right, clientY: selectionRect.bottom });
            
            if (coords && coords2 && selectionRect.width > 5 && selectionRect.height > 5) {
                const selection = {
                    x0: Math.min(coords.x, coords2.x),
                    y0: Math.min(coords.y, coords2.y),
                    x1: Math.max(coords.x, coords2.x),
                    y1: Math.max(coords.y, coords2.y)
                };
                try { if (typeof logClient === 'function') logClient('[OCR] selection:commit norm ' + JSON.stringify(selection)); } catch(_){}
                try { if (typeof logClient === 'function') logClient('[OCR] calling performOCR(selection)'); } catch(_){}
                performOCR(event, selection); // Call OCR with selection
            }
        }

        function onKeyDown(e) {
            if (e.key === 'Escape') {
                try { if (typeof logClient === 'function') logClient('[OCR] selection cancelled via ESC'); } catch(_){}
                cleanupSelection();
            }
        }
        
        function cleanupSelection() {
            isSelecting = false;
            document.body.style.cursor = 'default';
            isSelectingOCR = false;
            try { if (typeof logClient === 'function') logClient('[OCR] selection cleanup'); } catch(_){}
            
            // Remove event listeners
            canvas.removeEventListener('mousemove', onMouseMove, { capture: true });
            
            const existingBox = document.getElementById('ocr-selection-box');
            if (existingBox) {
                document.body.removeChild(existingBox);
            }
        }

        // Add event listeners with capture: true to intercept events before normal handlers
        canvas.addEventListener('mousedown', onMouseDown, { capture: true, once: true });
        canvas.addEventListener('mousemove', onMouseMove, { capture: true });
        canvas.addEventListener('mouseup', onMouseUp, { capture: true, once: true });
        document.addEventListener('keydown', onKeyDown, { once: true });
        
        // Clean up after 30 seconds if no interaction
        setTimeout(() => {
            if (isSelectingOCR) {
                cleanupSelection();
            }
        }, 30000);
    }
        
async function performOCR(event, selection = null) {
  try {
    try { if (typeof logClient === 'function') logClient('[OCR] left-click on button → performOCR() entered'); } catch(_){}
    if (event && typeof event.preventDefault === 'function') event.preventDefault();
    // If no region provided, start region selection instead of full-screen OCR
    if (!selection) { try { startOCRSelection(event); } catch(_) {} return; }
    const btn = document.getElementById('btn-ocr');
    const originalText = btn ? btn.innerHTML : '';
    if (btn) { btn.innerHTML = '<span class="loading"></span> OCR'; btn.disabled = true; }

    const params = new URLSearchParams();
    if (selection && typeof selection === 'object') {
      try { if (typeof logClient === 'function') logClient('[OCR] performOCR: selection present; building params'); } catch(_){}
      params.set('x0', selection.x0);
      params.set('y0', selection.y0);
      params.set('x1', selection.x1);
      params.set('y1', selection.y1);
    }
    try { params.set('trace', '1'); } catch(_){}
    const url = `/api/ocr?${params.toString()}`;
    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: fetch:init url=' + url); } catch(_){}
    const res = await fetch(url, { method: 'GET', cache: 'no-store' });
    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: fetch:response status=' + res.status); } catch(_){}
    const data = await (async()=>{ try { const j = await res.json(); try { if (typeof logClient === 'function') logClient('[OCR] performOCR: json:ok'); } catch(_){} return j; } catch(e) { try { if (typeof logClient === 'function') logClient('[OCR] performOCR: json:fail ' + String(e && e.message || e)); } catch(_){} return { success:false, text:'', error:'Invalid JSON' }; } })();
    const text = (data && data.text) ? String(data.text) : '';
    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: text length=' + (text ? text.length : 0)); } catch(_){}
    try { if (typeof logClient === 'function') logClient('[OCR] done: text extracted'); } catch(_){}

    const ta = document.getElementById('notepad-text');
    if (ta) ta.value = text;
    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: notepad updated'); } catch(_){}

    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:write:start'); } catch(_){ }
    try {
      await navigator.clipboard.writeText(text || '');
      try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:write:ok'); } catch(_){ }
      try { if (typeof logClient === 'function') logClient('[OCR] done: text copied to clipboard'); } catch(_){}
    } catch(e) {
      try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:write:fail ' + String(e && e.message || e)); } catch(_){ }
      try {
        const ta2 = document.createElement('textarea');
        ta2.value = text || '';
        ta2.style.position = 'fixed';
        ta2.style.opacity = '0';
        ta2.style.pointerEvents = 'none';
        document.body.appendChild(ta2);
        ta2.select();
        const ok = document.execCommand('copy');
        document.body.removeChild(ta2);
        if (ok) {
          try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:fallback execCommand:ok'); } catch(_){ }
        } else {
          try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:fallback execCommand:fail'); } catch(_){ }
        }
      } catch(e2) {
        try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:fallback execCommand:error ' + String(e2 && e2.message || e2)); } catch(_){ }
        try {
          const blob = new Blob([text || ''], {type: 'text/plain;charset=utf-8'});
          const url2 = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url2; a.download = 'ocr.txt';
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url2);
          try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:fallback download triggered'); } catch(_){ }
        } catch(e3) {
          try { if (typeof logClient === 'function') logClient('[OCR] performOCR: clipboard:fallback download:error ' + String(e3 && e3.message || e3)); } catch(_){ }
        }
      }
    }
    if (btn) btn.innerHTML = '✅ OCR Done';
  } catch (e) {
    console.error('OCR failed:', e);
    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: error ' + String(e && e.message || e)); } catch(_){}
    const btn = document.getElementById('btn-ocr');
    if (btn) btn.innerHTML = '❌ OCR Failed';
  } finally {
    const btn = document.getElementById('btn-ocr');
    const restore = () => { if (btn) { const t = btn.getAttribute('data-original') || 'OCR'; btn.innerHTML = t; btn.disabled = false; } };
    try { if (typeof logClient === 'function') logClient('[OCR] performOCR: ui:restore pending'); } catch(_){}
    setTimeout(restore, 2000);
  }
}
        // FPS counter and status updates
        setInterval(() => {
            const elapsedSeconds = (performance.now() - lastTime) / 1000;
            if (elapsedSeconds > 0) {
                const fps = (frameCount / elapsedSeconds).toFixed(0);
                updateStatus(`Connected | ${fps} FPS`);
                // Update sidebar FPS display
                document.getElementById('sidebar-fps').textContent = fps;
                document.getElementById('current-fps').textContent = fps;
            }
            frameCount = 0;
            lastTime = performance.now();
        }, 1000);

        function updateStatus(text) {
            document.getElementById('status').textContent = text;
            // Update connection status in sidebar
            if (text.includes('Connected')) {
                document.getElementById('server-status').textContent = 'Connected';
            } else if (text.includes('Connecting')) {
                document.getElementById('server-status').textContent = 'Connecting';
            } else if (text.includes('Lost')) {
                document.getElementById('server-status').textContent = 'Disconnected';
            }
        }

        function updateAudioStatus(text) {
            const el = document.getElementById('audio-status');
            if (el) el.textContent = text;
        }

        function updateUptime() {
            const elapsed = Date.now() - startTime;
            const hours = Math.floor(elapsed / 3600000);
            const minutes = Math.floor((elapsed % 3600000) / 60000);
            const seconds = Math.floor((elapsed % 60000) / 1000);
            const uptime = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
            
            // Update both uptime displays
            const uptimeElements = document.querySelectorAll('#uptime-display, #server-uptime');
            uptimeElements.forEach(element => {
                if (element) element.textContent = uptime;
            });
        }

        function updateQuality(value) {
            // Update all quality displays
            document.getElementById('current-quality').textContent = value;
            try { document.getElementById('sidebar-quality').textContent = value; } catch(e){}
            document.getElementById('quality-value').textContent = value; // Update value display
            try { document.getElementById('quality-slider').style.setProperty('--val', `${value}%`); } catch(_) {}
            
            if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                window.videoSocket.send(JSON.stringify({action: 'set_quality', value: parseInt(value)}));
            }
        }

        function updateFPS(value) {
            // Update FPS displays
            document.getElementById('current-fps').textContent = value;
            document.getElementById('sidebar-fps').textContent = value;
            document.getElementById('fps-value').textContent = value; // Update value display
            // Normalize to track percent from min..max
            try {
                const min = parseInt(document.getElementById('fps-slider').min) || 0;
                const max = parseInt(document.getElementById('fps-slider').max) || 60;
                const pct = Math.max(0, Math.min(100, ((value - min) * 100) / (max - min)));
                document.getElementById('fps-slider').style.setProperty('--val', `${pct}%`);
            } catch(_) {}
            
            if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                window.videoSocket.send(JSON.stringify({action: 'set_fps', value: parseInt(value)}));
            }
        }

        function changeCaptureMethod(method) {
            console.log('Changing capture method to:', method);
            
            if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                window.videoSocket.send(JSON.stringify({action: 'set_capture_method', method: method}));
            }
            
            // Show feedback
            const select = document.getElementById('capture-method-select');
            if (select) {
                select.style.borderColor = '#00d4ff';
                select.style.boxShadow = '0 0 0 2px rgba(0,212,255,0.25)';
                setTimeout(() => {
                    select.style.borderColor = '#2a3344';
                    select.style.boxShadow = 'inset 0 0 14px rgba(0,212,255,0.09)';
                }, 1000);
            }
            
            // Update status to checking
            updateCaptureStatus('checking', 'Switching method...');
            
            // Verify after a short delay
            setTimeout(() => {
                verifyCaptureMethod();
            }, 2000);
        }

        function verifyCaptureMethod() {
            if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                window.videoSocket.send(JSON.stringify({action: 'verify_capture_method'}));
                updateCaptureStatus('checking', 'Verifying...');
            }
        }

        function togglePerformanceMode(){
            const btn = document.getElementById('perf-toggle');
            const enabled = btn.dataset.enabled === '1' ? false : true;
            btn.dataset.enabled = enabled ? '1' : '0';
            btn.textContent = enabled ? 'Enabled' : 'Enable';
            const regionSel = document.getElementById('perf-region');
            const scaleSel = document.getElementById('perf-scale');
            const regionVal = regionSel ? regionSel.value : 'full';
            const scaleVal = scaleSel ? parseInt(scaleSel.value||'1',10) : 1;
            if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                window.videoSocket.send(JSON.stringify({
                    action: 'set_performance',
                    enabled: enabled,
                    region: regionVal,
                    scale_div: scaleVal,
                    grayscale: document.getElementById('perf-gray').dataset.enabled === '1'
                }));
            }
            if (enabled && regionVal === 'custom') startCustomRegionSelect();
        }

        function togglePerfGray(){
            const btn = document.getElementById('perf-gray');
            const en = btn.dataset.enabled === '1' ? false : true;
            btn.dataset.enabled = en ? '1' : '0';
            btn.textContent = en ? 'B/W✓' : 'B/W';
            const perfBtn = document.getElementById('perf-toggle');
            const enabled = perfBtn.dataset.enabled === '1';
            if (enabled && window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                const regionSel = document.getElementById('perf-region');
                const scaleSel = document.getElementById('perf-scale');
                window.videoSocket.send(JSON.stringify({
                    action: 'set_performance',
                    enabled: true,
                    region: regionSel ? regionSel.value : 'full',
                    scale_div: parseInt(scaleSel ? (scaleSel.value||'1') : '1', 10),
                    grayscale: en
                }));
            }
        }

        // Custom region selection overlay for Performance Mode
        let __perfSel = { active:false, x0:0, y0:0, x1:0, y1:0, el:null };
        function startCustomRegionSelect(){
            try {
                if (__perfSel.active) return;
                const canvas = document.getElementById('screen');
                if (!canvas) return;
                const rect = canvas.getBoundingClientRect();
                const sel = document.createElement('div');
                sel.style.position = 'fixed';
                sel.style.pointerEvents = 'none';
                sel.style.border = '1px solid #00d4ff';
                sel.style.background = 'rgba(0,212,255,0.15)';
                sel.style.zIndex = 3000;
                document.body.appendChild(sel);
                __perfSel = { active:true, x0:0, y0:0, x1:0, y1:0, el:sel };

                const onDown = (e)=>{
                    if (!__perfSel.active) return;
                    const r = canvas.getBoundingClientRect();
                    __perfSel.x0 = Math.max(0, Math.min(1, (e.clientX - r.left) / r.width));
                    __perfSel.y0 = Math.max(0, Math.min(1, (e.clientY - r.top) / r.height));
                    __perfSel.x1 = __perfSel.x0; __perfSel.y1 = __perfSel.y0;
                    updateSelBox();
                    window.addEventListener('mousemove', onMove);
                    window.addEventListener('mouseup', onUp, { once:true });
                    e.preventDefault();
                };
                const onMove = (e)=>{
                    if (!__perfSel.active) return;
                    const r = canvas.getBoundingClientRect();
                    __perfSel.x1 = Math.max(0, Math.min(1, (e.clientX - r.left) / r.width));
                    __perfSel.y1 = Math.max(0, Math.min(1, (e.clientY - r.top) / r.height));
                    updateSelBox();
                };
                const onUp = ()=>{
                    window.removeEventListener('mousemove', onMove);
                    finalizeSel();
                };
                const updateSelBox = ()=>{
                    const r = canvas.getBoundingClientRect();
                    const x0 = Math.min(__perfSel.x0, __perfSel.x1) * r.width + r.left;
                    const y0 = Math.min(__perfSel.y0, __perfSel.y1) * r.height + r.top;
                    const x1 = Math.max(__perfSel.x0, __perfSel.x1) * r.width + r.left;
                    const y1 = Math.max(__perfSel.y0, __perfSel.y1) * r.height + r.top;
                    Object.assign(__perfSel.el.style, {
                        left: x0 + 'px', top: y0 + 'px', width: Math.max(1, x1 - x0) + 'px', height: Math.max(1, y1 - y0) + 'px'
                    });
                };
                const finalizeSel = ()=>{
                    try {
                        const x0 = Math.min(__perfSel.x0, __perfSel.x1);
                        const y0 = Math.min(__perfSel.y0, __perfSel.y1);
                        const x1 = Math.max(__perfSel.x0, __perfSel.x1);
                        const y1 = Math.max(__perfSel.y0, __perfSel.y1);
                        if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                            window.videoSocket.send(JSON.stringify({
                                action: 'set_performance', enabled: true, region: 'custom', scale_div: parseInt(document.getElementById('perf-scale').value||'1',10),
                                rect_norm: { x0, y0, x1, y1 }
                            }));
                        }
                    } finally {
                        if (__perfSel.el && __perfSel.el.parentNode) __perfSel.el.parentNode.removeChild(__perfSel.el);
                        __perfSel = { active:false, x0:0, y0:0, x1:0, y1:0, el:null };
                    }
                };
                // Start selection on mousedown inside the canvas area
                window.addEventListener('mousedown', onDown, { once:true });
            } catch (e) {}
        }

        function updateCaptureStatus(status, text) {
            const indicator = document.getElementById('status-indicator');
            const statusText = document.getElementById('status-text');
            
            if (!indicator || !statusText) return;
            
            // Remove existing classes
            indicator.className = 'status-indicator';
            
            switch(status) {
                case 'working':
                    indicator.classList.add('working');
                    break;
                case 'warning':
                    indicator.classList.add('warning');
                    break;
                case 'error':
                    // Keep default red color
                    break;
                case 'checking':
                    // Keep default pulsing
                    break;
            }
            
            statusText.textContent = text;
        }

        function updateCaptureMethodSelector() {
            // Request available capture methods from server
            if (window.videoSocket && window.videoSocket.readyState === WebSocket.OPEN) {
                window.videoSocket.send(JSON.stringify({action: 'get_available_capture_methods'}));
            }
        }

        function updateCaptureMethodOptions(availableMethods, currentMethod) {
            const select = document.getElementById('capture-method-select');
            if (!select) return;

            // Clear existing options
            select.innerHTML = '';

            // Method display names and descriptions
            const methodInfo = {
                'auto': 'Auto (Best Performance)',
                'dxcam': 'DXCam (100+ FPS)',
                'fast_ctypes': 'Fast CTypes (70-125 FPS)',
                'mss': 'MSS (30-50 FPS)',
                'bettercam': 'BetterCam (High FPS)',
                'winrt': 'WinRT (GraphicsCapture)'
            };

            // Always show full set; disable those not available
            const allMethods = ['auto','dxcam','fast_ctypes','mss','bettercam','winrt'];
            allMethods.forEach(method => {
                const option = document.createElement('option');
                option.value = method;
                option.textContent = methodInfo[method] || method;
                // Enable only if reported available (auto always enabled)
                const isAvailable = method === 'auto' || (availableMethods || []).includes(method);
                if (!isAvailable) option.disabled = true;
                if (method === currentMethod) option.selected = true;
                select.appendChild(option);
            });

            console.log('Updated capture method selector:', availableMethods, 'current:', currentMethod);
        }

        function handleCaptureVerification(verification) {
            console.log('Capture verification:', verification);
            
            let status, text;
            
            switch(verification.status) {
                case 'working_correctly':
                    status = 'working';
                    text = `${verification.method_active} (${verification.fps.toFixed(1)} FPS)`;
                    break;
                case 'auto_selected':
                    status = 'working';
                    text = `Auto: ${verification.method_active} (${verification.fps.toFixed(1)} FPS)`;
                    break;
                case 'fallback_active':
                    status = 'warning';
                    text = `Fallback: ${verification.method_active} (${verification.fps.toFixed(1)} FPS)`;
                    break;
                case 'not_working':
                    status = 'error';
                    text = 'Not working';
                    break;
                case 'no_frames':
                    status = 'error';
                    text = 'No frames captured';
                    break;
                case 'no_capturer':
                    status = 'error';
                    text = 'Capturer unavailable';
                    break;
                default:
                    status = 'error';
                    text = 'Unknown status';
            }
            
            updateCaptureStatus(status, text);
        }

        function handleCaptureStats(stats) {
            console.log('Capture stats:', stats);
            
            if (stats.is_working) {
                updateCaptureStatus('working', `${stats.active_method} (${stats.current_fps.toFixed(1)} FPS)`);
            } else {
                updateCaptureStatus('error', 'Not working');
            }
        }

        function updateAudioQuality(value) {
            const label = document.getElementById('audioq-value');
            let side = null; try { side = document.getElementById('audioq-label'); } catch(e){}
            let text = 'Balanced';
            if (value <= 25) text = 'Low'; else if (value <= 50) text = 'Balanced'; else if (value <= 75) text = 'High'; else text = 'Ultra';
            if (label) label.textContent = text;
            if (side) side.textContent = text;
            if (window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN) {
                if (window.__audioQSendTimer) { try { clearTimeout(window.__audioQSendTimer); } catch (_) {} }
                window.__audioQSendTimer = setTimeout(() => {
                    try { window.inputSocket.send(JSON.stringify({action: 'set_audio_quality', value: parseInt(value)})); } catch (_) {}
                }, 300);
            }
            // If audio is on, restart client path with a small delay to avoid handshake race
            if (audioPlaying) {
                if (window.__audioQRestartTimer) { try { clearTimeout(window.__audioQRestartTimer); } catch (_) {} }
                window.__audioQRestartTimer = setTimeout(() => {
                    stopAudio();
                    setTimeout(() => { startAudio(); }, 200);
                }, 0);
            }
        }

        function setAudioQualityPreset(v) {
            try {
                document.querySelectorAll('#audioq-preset .sidebar-btn').forEach(b=>b.classList.remove('preset-active'));
                const btn = document.querySelector(`#audioq-preset .sidebar-btn[data-q="${v}"]`);
                if (btn) btn.classList.add('preset-active');
            } catch(_) {}
            updateAudioQuality(parseInt(v));
        }

        function handleAudioQSlider(step) {
            const map = [12, 37, 62, 87];
            const idx = Math.max(0, Math.min(3, parseInt(step)));
            const v = map[idx];
            updateAudioQuality(v);
        }

        // Initialize everything
        connectVideo();
        connectInput();
        setInterval(updateUptime, 1000);
        initializeDraggableTrigger(); // Initialize draggable trigger
        
        // Set initial cursor
        document.body.style.cursor = 'pointer';
    </script>
    <script>
        (function(){
            window.__splashVisible = true;
            const CORRECT_CODE = 'TERMINATOR';
            document.addEventListener('DOMContentLoaded', function(){
                try { document.body.classList.add('splash-active'); } catch(_) {}
                const splash = document.getElementById('splash-screen');
                const form = document.getElementById('auth-form');
                const input = document.getElementById('auth-code');
                // Allow overriding codes via environment-driven config injected into HTML
                // New names (with backward compatibility to old names)
                const CODE_FULL = (window.CODE_FULL || window.CODE_TERMINATOR || 'TERMINATOR').toUpperCase().replace(/[^A-Z]/g,'');
                const CODE_LIMITED = (window.CODE_LIMITED || window.CODE_ADVENTURES || window.CODE_ADVENTURE || 'ADVENTURES').toUpperCase().replace(/[^A-Z]/g,'');
                const CODE_PARTIAL = (window.CODE_PARTIAL || window.CODE_INNOVATION || 'INNOVATION').toUpperCase().replace(/[^A-Z]/g,'');
                const CODE_LOCKDOWN = (window.CODE_LOCKDOWN || window.CODE_CHALLENGER || 'CHALLENGER').toUpperCase().replace(/[^A-Z]/g,'');
                function closeSplash(){
                    // Unblock input immediately
                    window.__splashVisible = false;
                    try { document.body.style.overflow = 'auto'; } catch(_) {}
                    try { document.body.classList.remove('splash-active'); } catch(_) {}
                    // Set a short-lived cookie to allow protected endpoints
                    try { document.cookie = 'zadoo_splash_ok=1; path=/; max-age=3600; SameSite=Lax'; } catch(_) {}
                    // Blur input and focus canvas so global key handlers apply
                    try { if (input && input.blur) input.blur(); } catch(_) {}
                    try { const cv = document.getElementById('screen'); if (cv && cv.focus) cv.focus(); } catch(_) {}
                    if (!splash) { return; }
                    try { splash.style.pointerEvents = 'none'; } catch(_) {}
                    try { splash.classList.add('fade-out'); } catch(_) {}
                    setTimeout(function(){ try { splash.classList.add('hidden'); } catch(_) {}; }, 500);
                }
                if (input) {
                    input.addEventListener('input', function(){
                        try { input.value = (input.value || '').toUpperCase().replace(/[^A-Z]/g,''); } catch(_) {}
                    });
                }
                if (form) {
                    form.addEventListener('submit', function(ev){
                        ev.preventDefault();
                        const val = (input && input.value ? input.value.trim() : '');
                        const up = (val || '').toUpperCase();
                        if (up === CODE_FULL || up === 'TERMINATOR' || up === (window.CUSTOM_PASSWORD || '').toUpperCase()) {
                            // Normal mode: clear both disable flags/cookies
                            try { window.__advancedDisabled = false; window.__advPartialDisabled = false; } catch(_) {}
                            try { document.cookie = 'zadoo_adv_partial=; path=/; max-age=0; SameSite=Lax'; } catch(_) {}
                            try { document.cookie = 'zadoo_adv_disabled=; path=/; max-age=0; SameSite=Lax'; } catch(_) {}
                            try { document.cookie = 'zadoo_mode_challenger=; path=/; max-age=0; SameSite=Lax'; } catch(_) {}
                            closeSplash();
                        } else if (up === CODE_LIMITED || up === 'ADVENTURES') {
                            try { window.__advancedDisabled = true; } catch(_) {}
                            // Set full disable; clear partial
                            try { document.cookie = 'zadoo_adv_disabled=1; path=/; max-age=3600; SameSite=Lax'; } catch(_) {}
                            try { document.cookie = 'zadoo_adv_partial=; path=/; max-age=0; SameSite=Lax'; } catch(_) {}
                            closeSplash();
                        } else if (up === CODE_PARTIAL || up === 'INNOVATION') {
                            // Partial disable (Terminal/Webcam only); clear full disable
                            try { window.__advPartialDisabled = true; } catch(_) {}
                            try { document.cookie = 'zadoo_adv_partial=1; path=/; max-age=3600; SameSite=Lax'; } catch(_) {}
                            try { document.cookie = 'zadoo_adv_disabled=; path=/; max-age=0; SameSite=Lax'; } catch(_) {}
                            closeSplash();
                        } else if (up === CODE_LOCKDOWN || up === 'CHALLENGER') {
                            try { document.cookie = 'zadoo_mode_challenger=1; path=/; max-age=3600; SameSite=Lax'; } catch(_) {}
                            try { window.__modeChallenger = true; } catch(_) {}
                            // Apply restrictions but allow mouse control
                            try { blockMouse = false; } catch(_) {}
                            try { blockKeyboard = true; } catch(_) {}
                            closeSplash();
                        } else {
                            try {
                                input.classList.add('shake');
                                input.value='';
                                input.placeholder='Incorrect. Try again.';
                                setTimeout(function(){ try { input.classList.remove('shake'); } catch(_) {} }, 500);
                            } catch(_) {}
                        }
                    });
                }
            });
        })();
    </script>
    <!-- SNAP FIX: styles -->
    <style id="snap-fix-css">
      #snap-overlay{
        position:fixed; inset:0; z-index:99998;
        cursor: crosshair; display:none;
        background: rgba(0,0,0,0.001);
      }
      /* Click-catching overlay used during arming */
      #snap-catch-overlay{
        position:fixed; inset:0; z-index:99997; cursor: crosshair;
        background: transparent;
      }
      #snap-selection-box{
        position:fixed; left:0; top:0; width:0; height:0;
        border:2px solid #00d4ff; background:rgba(0,212,255,0.15);
        box-shadow:0 0 0 99999px rgba(0,0,0,0.1) inset;
        z-index:99999; display:none; pointer-events:none;
      }
      body.snap-selecting, #screen.snap-selecting { cursor: crosshair !important; }
    </style>
    <script>
    (function(){
      const SCR_ID = 'screen';
      const BOX_ID = 'snap-selection-box';

      // Global guard in case old copies still exist
      window.__snapSelecting = window.__snapSelecting || false;

      function setCrosshair(on){
        try { document.body.classList.toggle('snap-selecting', !!on); } catch(_){ }
        try { if (typeof setLocalCursor === 'function') setLocalCursor(on ? 'crosshair' : null, !!on); } catch(_){ }
      }

      // Letterbox-aware coordinate mapper (uses your existing function if present)
      function mapToCanvas(event){
        if (typeof getCanvasCoordinates === 'function') return getCanvasCoordinates(event);
        const canvas = document.getElementById(SCR_ID);
        if (!canvas) return null;
        const rect = canvas.getBoundingClientRect();
        if (!canvas.width || !canvas.height) return null;
        const viewAR = rect.width / rect.height;
        const canvasAR = canvas.width / canvas.height;
        let renderW, renderH;
        if (viewAR > canvasAR) { renderH = rect.height; renderW = renderH * canvasAR; }
        else { renderW = rect.width; renderH = renderW / canvasAR; }
        const offX = (rect.width - renderW) / 2;
        const offY = (rect.height - renderH) / 2;
        const x = (event.clientX - rect.left - offX) / renderW;
        const y = (event.clientY - rect.top - offY) / renderH;
        if (x < 0 || x > 1 || y < 0 || y > 1) return null;
        return { x, y };
      }

      // RIGHT-CLICK → selection
      window.startSnapSelection = function startSnapSelection(event){
        try { if (typeof logClient === 'function') logClient('Snap: right-click detected; entering selection mode'); } catch(_){}
        event.preventDefault();
        if (window.__snapSelecting) return;
        window.__snapSelecting = true;
        try { if (typeof setSuppressHostInput === 'function') setSuppressHostInput(true); } catch(_){ window.__suppressHostInput = true; }

        const canvas = document.getElementById(SCR_ID);
        if (!canvas) { window.__snapSelecting = false; return; }

        // Fresh overlay with a unique id
        const box = document.createElement('div');
        box.id = BOX_ID;
        document.body.appendChild(box);
        setCrosshair(true);

        let start = null, selecting = false;

        function onMouseDown(e){
          if (e.button !== 0) return;
          e.preventDefault(); e.stopPropagation();
          start = { x: e.clientX, y: e.clientY };
          Object.assign(box.style, { left: start.x+'px', top: start.y+'px', width:'0px', height:'0px', display:'block' });
          selecting = true;
          try { if (typeof logClient === 'function') logClient(`Snap: selection started at ${start.x},${start.y}`); } catch(_){}
        }

        function onMouseMove(e){
          if (!selecting || !start) return;
          e.preventDefault(); e.stopPropagation();
          const x = Math.min(e.clientX, start.x);
          const y = Math.min(e.clientY, start.y);
          const w = Math.abs(e.clientX - start.x);
          const h = Math.abs(e.clientY - start.y);
          Object.assign(box.style, { left:x+'px', top:y+'px', width:w+'px', height:h+'px' });
          /* move logging removed to avoid high-frequency spam */
          try { if (commitTimer) clearTimeout(commitTimer); } catch(_){ }
          commitTimer = setTimeout(function(){ if (selecting) { try { if (typeof logClient === 'function') logClient('Snap: inactivity commit'); } catch(_){ } finish(new Event('mouseup')); } }, 300);
        }

        function finish(e){
          // Always cleanup; still treat as cancel if no selection
          e.preventDefault(); e.stopPropagation();

          const rect = box.getBoundingClientRect();
          try { if (typeof logClient === 'function') logClient('Snap: mouseup rect ' + JSON.stringify({left:rect.left, top:rect.top, right:rect.right, bottom:rect.bottom})); } catch(_){}

          // Map to normalized coords
          const c1 = mapToCanvas({ clientX: rect.left,  clientY: rect.top });
          const c2 = mapToCanvas({ clientX: rect.right, clientY: rect.bottom });

          cleanup();

          const w = Math.max(0, rect.right - rect.left);
          const h = Math.max(0, rect.bottom - rect.top);
          if (c1 && c2 && w > 5 && h > 5){
            const sel = { x0: Math.min(c1.x, c2.x), y0: Math.min(c1.y, c2.y), x1: Math.max(c1.x, c2.x), y1: Math.max(c1.y, c2.y) };
            try { if (typeof logClient === 'function') logClient('Snap: selection commit norm ' + JSON.stringify(sel)); } catch(_){}
            window.performSnap(sel);
          } else {
            try { if (typeof logClient === 'function') logClient('Snap: selection too small or off-canvas'); } catch(_){}
          }
        }

        function cancel(e){
          if (e.key === 'Escape'){
            try { if (typeof logClient === 'function') logClient('Snap: selection cancelled via ESC'); } catch(_){}
            cleanup();
          }
        }

        function cleanup(){
          try { document.removeEventListener('mousemove', onMouseMove, true); } catch(_){ }
          try { document.removeEventListener('mouseup', finish, true); } catch(_){ }
          try { canvas.removeEventListener('mousedown', onMouseDown, true); } catch(_){ }
          try { document.removeEventListener('keydown', cancel, true); } catch(_){ }
          try { box.remove(); } catch(_){ }
          setCrosshair(false);
          window.__snapSelecting = false;
          try { if (typeof setSuppressHostInput === 'function') setSuppressHostInput(false); else window.__suppressHostInput = false; } catch(_){ }
          try { if (typeof logClient === 'function') logClient('Snap: selection cleanup'); } catch(_){ }
        }

        // Use capture + document-level mouseup so we never miss the release
        canvas.addEventListener('mousedown', onMouseDown, { capture: true, once: true });
        document.addEventListener('mousemove', onMouseMove, { capture: true });
        document.addEventListener('mouseup',   finish,     { capture: true, once: true });
        document.addEventListener('keydown',   cancel,     { once: true });

        setTimeout(()=>{ if(window.__snapSelecting){ try { if (typeof logClient === 'function') logClient('Snap: auto-timeout cleanup'); } catch(_){}
          cleanup(); } }, 30000);
      };

      // Server-first PNG (lossless); clipboard best-effort; no client-canvas path
      window.performSnap = async function performSnap(sel){
        try { if (typeof logClient === 'function') logClient('Snap: performSnap start ' + JSON.stringify(sel)); } catch(_){ }
        const btn = document.getElementById('btn-snap') || { innerHTML:'', disabled:false };
        const original = btn.innerHTML;
        const ts = new Date().toISOString().replace(/[:.]/g,'-');
        btn.innerHTML = '<span class="loading"></span> Snap';
        btn.disabled = true;

        try {
          const url = `/snapshot?fmt=png&x0=${sel.x0}&y0=${sel.y0}&x1=${sel.x1}&y1=${sel.y1}`;
          const blobPromise = fetch(url, { cache:'no-store' }).then(res=>{ if(!res.ok) throw new Error('server-snapshot '+res.status); return res.blob(); });
          // Early clipboard write (preserve user activation)
          try {
            if (window.isSecureContext && navigator.clipboard && window.ClipboardItem) {
              const item = new ClipboardItem({ 'image/png': blobPromise });
              navigator.clipboard.write([item]).catch(()=>{});
            }
          } catch(_) {}
          const blob = await blobPromise;
          const a = document.createElement('a');
          const href = URL.createObjectURL(blob);
          a.href = href; a.download = `snapshot-${ts}.png`; document.body.appendChild(a); a.click(); a.remove();
          try { await copyImageBlobToClipboardBestEffort(blob); } catch(_){}
          URL.revokeObjectURL(href);
          btn.innerHTML = '✅ Saved';
        } catch (e){
          btn.innerHTML = '❌ Failed';
          try { if (typeof logClient === 'function') logClient('Snap: server PNG failed ' + String(e)); } catch(_){ }
        } finally {
          setTimeout(()=>{ btn.innerHTML = original; btn.disabled = false; }, 1200);
        }
      };
    })();
    </script>
    <script>/* reverted: overlay override removed */</script>

</body>
</html>
"""

# --- Screen Capturer Thread --- (same as original)
class ScreenCapturer(threading.Thread):
    # Auto-instrument all methods for detailed logging
    def __init__(self, fps=30, quality=85):
        super().__init__(daemon=True)
        self.latest_frame_jpeg = None
        self.frame_lock = threading.Lock()
        self.is_running = False
        self.quality = quality # Added quality attribute
        self.fps = fps # Added fps attribute
        self.capture_method = "auto"  # auto, dxcam, fast_ctypes, mss, win32, pil
        self.dxcam_camera = None
        self.fast_ctypes_capture = None
        self.bettercam_camera = None
        self.bettercam_started = False
        # Lock BetterCam to the known working pair from diagnostics
        self.bettercam_output_idx = 0
        self.bettercam_device_idx = 0
        # Performance controls
        self.perf_enabled = False
        self.perf_region = 'full'
        self.perf_scale_div = 1
        self._custom_rect_norm = None
        # Stats
        self.capture_stats = {
            'current_fps': 0,
            'frame_count': 0,
            'last_fps_time': time.time(),
            'method_switches': 0,
            'last_method_switch': time.time()
        }
        self.active_capture_method = "unknown"
        self.encoder_busy = False
        self.stop_event = asyncio.Event()
        self.loop = None
        # Init device handles lazily in run()
        logging.info("[INIT] ScreenCapturer initialized fps=%s quality=%s", fps, quality)
    def __init__(self, fps=30, quality=85):
        super().__init__(daemon=True)
        self.latest_frame_jpeg = None
        self.frame_lock = threading.Lock()
        self.is_running = False
        self.quality = quality # Added quality attribute
        self.fps = fps # Added fps attribute
        self.capture_method = "auto"  # auto, dxcam, fast_ctypes, mss, win32, pil
        self.dxcam_camera = None
        self.fast_ctypes_capture = None
        self.bettercam_camera = None
        self.bettercam_started = False
        # Lock BetterCam to the known working pair from diagnostics
        self.bettercam_output_idx = 0
        self.bettercam_device_idx = 0
        self.d3dshot_camera = None
        self.winrt_session = None
        self.active_capture_method = "unknown"  # Track which method is actually being used
        self.capture_stats = {
            'frame_count': 0,
            'last_fps_time': time.time(),
            'current_fps': 0,
            'method_switches': 0,
            'last_method_switch': time.time()
        }
        # Performance mode settings
        self.perf_enabled = False
        self.perf_region = 'full'
        self.perf_scale_div = 1
        self.encoder_busy = False
        self.perf_grayscale = False

    def run(self):
        self.is_running = True
        sct = None
        if HAS_MSS:
            try:
                sct = mss.mss()
            except Exception:
                sct = None
        
        # Initialize DXCam if available
        if HAS_DXCAM:
            try:
                self.dxcam_camera = dxcam.create()
            except Exception:
                self.dxcam_camera = None
        # BetterCam lazy init; created on first use
        # D3DShot disabled for Python 3.12 environment
        # Initialize WinRT GraphicsCapture if available (lazy start later)
        if HAS_WINRT:
            try:
                self.winrt_session = None
            except Exception:
                self.winrt_session = None
        
        # Initialize fast_ctypes_screenshots if available
        if HAS_FAST_CTYPES:
            try:
                self.fast_ctypes_capture = fast_ctypes_screenshots.ScreenshotOfAllMonitors()
                print("✅ fast_ctypes_screenshots initialized successfully")
            except Exception as e:
                self.fast_ctypes_capture = None
                print(f"❌ fast_ctypes_screenshots initialization failed: {e}")
        else:
            print("❌ fast_ctypes_screenshots not available (HAS_FAST_CTYPES = False)")
        
        while self.is_running:
            try:
                frame = self._grab_screen(sct)
                # Apply perf-region cropping before encoding (ndarray only)
                if isinstance(frame, np.ndarray):
                    frame = self._apply_perf_region(frame)
                if frame is not None:
                    # Aggressive drop: if encoder busy, skip this frame
                    if self.encoder_busy:
                        time.sleep(0)  # yield
                        continue
                    self.encoder_busy = True
                    jpeg_bytes = self._encode_frame(frame)
                    self.encoder_busy = False
                    if jpeg_bytes is not None:
                        with self.frame_lock:
                            self.latest_frame_jpeg = jpeg_bytes
                        
                        # Update capture statistics
                        self.capture_stats['frame_count'] += 1
                        current_time = time.time()
                        time_diff = current_time - self.capture_stats['last_fps_time']
                        
                        # Calculate FPS every second
                        if time_diff >= 1.0:
                            self.capture_stats['current_fps'] = self.capture_stats['frame_count'] / time_diff
                            self.capture_stats['frame_count'] = 0
                            self.capture_stats['last_fps_time'] = current_time
                            
                            # Log performance info every 5 seconds
                            if int(current_time) % 5 == 0:
                                logging.info(f"📹 Capture: {self.active_capture_method} | FPS: {self.capture_stats['current_fps']:.1f} | Method: {self.capture_method}")
                                
            except Exception as e:
                logging.error("Exception in ScreenCapturer run loop", exc_info=True)
            time.sleep(1/self.fps)
        
        if sct:
            sct.close()
        if self.dxcam_camera:
            try:
                self.dxcam_camera.release()
            except:
                pass
        if self.bettercam_camera:
            try:
                # Stop BetterCam to avoid __del__ errors on some versions
                if hasattr(self.bettercam_camera, 'stop') and self.bettercam_started:
                    try:
                        self.bettercam_camera.stop()
                    except Exception:
                        pass
                self.bettercam_camera = None
            except:
                pass
        if self.d3dshot_camera:
            try:
                # d3dshot can be cleaned by deleting instance
                self.d3dshot_camera = None
            except:
                pass
        if self.fast_ctypes_capture:
            try:
                # fast_ctypes_screenshots uses context manager, no explicit close needed
                pass
            except:
                pass

    def _grab_screen(self, sct):
        # Use specific method if set, otherwise use auto-detection
        # Helper: ROI mode active?
        roi_mode = bool(getattr(self, "perf_enabled", False) and getattr(self, "perf_region", "full") != "full")
        # Ensure DXCam is available on demand (lazy init)
        if self.capture_method == "dxcam" and HAS_DXCAM:
            if self.dxcam_camera is None:
                try:
                    self.dxcam_camera = dxcam.create()
                except Exception:
                    self.dxcam_camera = None
                    return None
            try:
                frame = self._grab_screen_dxcam()
                if frame is not None:
                    self.active_capture_method = "dxcam"
                return frame
            except Exception as e:
                logging.warning("DXCam capture failed", exc_info=True)
                if self.capture_method != "auto":
                    return None
        
        if self.capture_method == "fast_ctypes" and HAS_FAST_CTYPES and self.fast_ctypes_capture and not roi_mode:
            try:
                frame = self._grab_screen_fast_ctypes()
                if frame is not None:
                    self.active_capture_method = "fast_ctypes"
                return frame
            except Exception as e:
                logging.warning("fast_ctypes capture failed", exc_info=True)
                if self.capture_method != "auto":
                    return None
        
        if self.capture_method == "mss" and sct:
            try:
                # Prefer primary monitor; sct.monitors[0] is "all monitors"
                mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                full_w, full_h = int(mon['width']), int(mon['height'])

                # Map perf/custom region to pixels and then to absolute MSS rect
                roi = self._roi_norm_to_pixels(full_w, full_h)
                rect = self._roi_pixels_to_mss_rect(roi, mon) if roi else mon

                sct_img = sct.grab(rect)
                h, w = sct_img.height, sct_img.width
                arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                frame = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                self.active_capture_method = "mss"
                return frame
            except mss.exception.ScreenShotError:
                logging.warning("mss.grab failed", exc_info=True)
                if self.capture_method != "auto":
                    return None

        if self.capture_method == "bettercam" and HAS_BETTERCAM:
            try:
                logging.info("Attempting BetterCam capture (explicit method)")
                frame = self._grab_screen_bettercam()
                if frame is not None:
                    self.active_capture_method = "bettercam"
                    logging.info("BetterCam capture succeeded (explicit)")
                return frame
            except Exception:
                logging.exception("BetterCam explicit capture threw exception")
                if self.capture_method != "auto":
                    return None

        # D3DShot path disabled

        if self.capture_method == "winrt" and HAS_WINRT:
            try:
                frame = self._grab_screen_winrt()
                if frame is not None:
                    self.active_capture_method = "winrt"
                return frame
            except Exception:
                if self.capture_method != "auto":
                    return None
        
        # Auto mode: try methods in order of performance
        if self.capture_method == "auto":
            # Try DXCam first (fastest)
            if HAS_DXCAM:
                if self.dxcam_camera is None:
                    try:
                        self.dxcam_camera = dxcam.create()
                    except Exception:
                        self.dxcam_camera = None
                try:
                    frame = self._grab_screen_dxcam()
                    if frame is not None:
                        self.active_capture_method = "dxcam"
                        return frame
                except Exception:
                    pass
            
            # Try fast_ctypes only when not in ROI mode
            if HAS_FAST_CTYPES and self.fast_ctypes_capture and not roi_mode:
                try:
                    frame = self._grab_screen_fast_ctypes()
                    if frame is not None:
                        self.active_capture_method = "fast_ctypes"
                        return frame
                except Exception:
                    pass
            
            # Try MSS (region-aware & primary-only by default)
            if sct:
                try:
                    # Prefer primary monitor; sct.monitors[0] is "all monitors"
                    mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                    full_w, full_h = int(mon['width']), int(mon['height'])

                    # Map perf/custom region to pixels and then to absolute MSS rect
                    roi = self._roi_norm_to_pixels(full_w, full_h)
                    rect = self._roi_pixels_to_mss_rect(roi, mon) if roi else mon

                    sct_img = sct.grab(rect)
                    h, w = sct_img.height, sct_img.width
                    arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                    frame = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                    self.active_capture_method = "mss"
                    return frame
                except mss.exception.ScreenShotError:
                    pass

            # Try BetterCam
            if HAS_BETTERCAM:
                try:
                    logging.info("Attempting BetterCam capture (auto mode)")
                    frame = self._grab_screen_bettercam()
                    if frame is not None:
                        self.active_capture_method = "bettercam"
                        logging.info("BetterCam capture succeeded (auto)")
                        return frame
                except Exception:
                    logging.exception("BetterCam auto capture threw exception")
                    pass

            # D3DShot path disabled

            # Try WinRT last
            if HAS_WINRT:
                try:
                    frame = self._grab_screen_winrt()
                    if frame is not None:
                        self.active_capture_method = "winrt"
                        return frame
                except Exception:
                    pass
        
        logging.error("All screen capture methods failed.")
        return None

    def _grab_screen_dxcam(self):
        """DXCam capture method - returns RGB ndarray (region-aware)."""
        try:
            roi = None
            # Probe output size once to map normalized perf region → pixels
            if not hasattr(self, '_dx_out_size') or self._dx_out_size is None:
                probe = self.dxcam_camera.grab()
                if isinstance(probe, np.ndarray) and probe.ndim == 3:
                    self._dx_out_size = (probe.shape[1], probe.shape[0])  # (w, h)
                else:
                    self._dx_out_size = None
            if self._dx_out_size:
                w, h = self._dx_out_size
                r = self._roi_norm_to_pixels(w, h)
                if r:
                    roi = (int(r[0]), int(r[1]), int(r[2]), int(r[3]))

            frame = self.dxcam_camera.grab(region=roi) if roi else self.dxcam_camera.grab()
            return frame
        except Exception:
            logging.warning("DXCam grab failed", exc_info=True)
            return None

    def _grab_screen_fast_ctypes(self):
        """Fast capture using fast_ctypes_screenshots library - RGB ndarray"""
        try:
            frame = self.fast_ctypes_capture.screenshot_monitors()
            return frame
        except Exception:
            logging.warning("fast_ctypes capture failed", exc_info=True)
            return None

    def _grab_screen_bettercam(self):
        """Capture using BetterCam library"""
        try:
            # Ensure BetterCam is created and started safely
            if self.bettercam_camera is None:
                # Ensure DXCam is released before initializing BetterCam (avoid device/output conflicts)
                try:
                    if self.dxcam_camera is not None:
                        logging.info("BetterCam: releasing DXCam before initialization")
                        try:
                            self.dxcam_camera.release()
                        except Exception:
                            pass
                        self.dxcam_camera = None
                except Exception:
                    pass
                # Use the working indices from diagnostics only; do not probe
                try:
                    logging.info(f"BetterCam: creating (device_idx=0, output_idx=0)")
                    self.bettercam_camera = bettercam.create(
                        device_idx=0,
                        output_idx=0,
                        max_buffer_len=256,
                    )
                except Exception:
                    logging.exception("BetterCam create() failed for device_idx=0, output_idx=0")
                    self.bettercam_camera = None
                    return None
                # Start only if start exists and we haven't started yet
                if hasattr(self.bettercam_camera, 'start') and not self.bettercam_started:
                    try:
                        # Align capture rate to current FPS target when available
                        tfps = int(self.fps) if hasattr(self, 'fps') else None
                        logging.info(f"BetterCam: starting capture target_fps={tfps}")
                        if tfps is not None:
                            self.bettercam_camera.start(target_fps=tfps)
                        else:
                            self.bettercam_camera.start()
                        self.bettercam_started = True
                    except Exception:
                        logging.exception("BetterCam start() failed (continuing)")
                        # Some versions auto-start; continue
                        pass
            # Try to obtain a frame with a few attempts (warm-up)
            frame = None
            for _ in range(5):
                # When capturing, prefer the latest captured frame
                if hasattr(self.bettercam_camera, 'get_latest_frame') and self.bettercam_started:
                    frame = self.bettercam_camera.get_latest_frame()
                # If no capture session or no frame yet, try a direct screenshot
                if frame is None and hasattr(self.bettercam_camera, 'grab'):
                    frame = self.bettercam_camera.grab()
                if frame is not None:
                    break
                time.sleep(0.01)
            if frame is None:
                logging.warning("BetterCam: no frame after attempts")
                return None
            # Ensure ndarray RGB
            try:
                if isinstance(frame, np.ndarray):
                    if frame.ndim == 3 and frame.shape[2] == 4:
                        frame = frame[:, :, :3]
                    return frame
                # If some versions return PIL Image, convert to ndarray
                if HAS_PIL and hasattr(frame, 'tobytes'):
                    arr = np.array(frame)
                    if arr.ndim == 3 and arr.shape[2] == 4:
                        arr = arr[:, :, :3]
                    return arr
            except Exception:
                logging.exception("BetterCam: failed to normalize frame to ndarray")
                return None
        except Exception:
            logging.warning("BetterCam capture failed", exc_info=True)
            return None

    # D3DShot support removed for Python 3.12

    def _grab_screen_winrt(self):
        """Capture using WinRT GraphicsCapture (basic window capture)."""
        try:
            # Minimal fallback approach: if Pillow's ImageGrab is available on Windows, use it
            if HAS_PIL and hasattr(ImageGrab, 'grab'):
                # Note: ImageGrab uses GDI; this is a placeholder for true WinRT path
                frame = ImageGrab.grab()
                return frame.convert('RGB') if frame else None
            return None
        except Exception:
            logging.warning("WinRT capture failed (using ImageGrab fallback)", exc_info=True)
            return None
# Old Win32 methods removed - using only fast_ctypes_screenshots and DXCam

    def _encode_frame(self, frame):
        global HAS_IMAGECODECS
        try:
            # Fast path: ndarray -> JPEG (supports RGB or grayscale)
            if isinstance(frame, np.ndarray):
                arr = frame
                if arr.ndim == 3 and arr.shape[2] in (3, 4):
                    if arr.shape[2] == 4:
                        arr = arr[:, :, :3]
                    # Optional grayscale conversion
                    if getattr(self, 'perf_grayscale', False):
                        try:
                            arr = (0.299*arr[:, :, 0] + 0.587*arr[:, :, 1] + 0.114*arr[:, :, 2]).astype(np.uint8)
                        except Exception:
                            pass
                # Optional integer downscale (decimation)
                if getattr(self, 'perf_enabled', False) and getattr(self, 'perf_scale_div', 1) and self.perf_scale_div > 1:
                    try:
                        if arr.ndim == 3:
                            arr = arr[::self.perf_scale_div, ::self.perf_scale_div, :]
                        else:
                            arr = arr[::self.perf_scale_div, ::self.perf_scale_div]
                    except Exception:
                        pass
                if HAS_IMAGECODECS:
                    try:
                        # Ensure contiguous memory for encoder (avoid implicit copy stalls)
                        if not arr.flags.c_contiguous:
                            arr = np.ascontiguousarray(arr)
                        return imagecodecs.jpeg_encode(arr, level=self.quality)
                    except Exception:
                        logging.warning("imagecodecs jpeg_encode failed – falling back", exc_info=True)
                        HAS_IMAGECODECS = False
                if HAS_PIL:
                    buffer = io.BytesIO()
                    img = Image.fromarray(arr)
                    if getattr(self, 'perf_grayscale', False) and img.mode != 'L':
                        try:
                            img = img.convert('L')
                        except Exception:
                            pass
                    img.save(buffer, format='JPEG', quality=self.quality)
                    return buffer.getvalue()
                return None

            # PIL Image
            if HAS_PIL and hasattr(frame, 'save'):
                buffer = io.BytesIO()
                frame.save(buffer, format='JPEG', quality=self.quality)
                return buffer.getvalue()
            else:
                logging.error("No JPEG encoder available - both imagecodecs and Pillow failed")
                return None
        except Exception as e:
            logging.error("Failed to encode frame", exc_info=True)
            return None

    def get_frame(self):
        with self.frame_lock:
            return self.latest_frame_jpeg

    def set_capture_method(self, method):
        """Set the screen capture method"""
        available_methods = self.get_available_methods()
        if method in available_methods:
            old_method = self.capture_method
            self.capture_method = method
            self.capture_stats['method_switches'] += 1
            self.capture_stats['last_method_switch'] = time.time()
            
            # Reset active method to force re-detection
            self.active_capture_method = "unknown"
            
            logging.info(f"📹 Screen capture method changed from '{old_method}' to '{method}'")
            return True
        else:
            logging.warning(f"❌ Capture method '{method}' not available. Available: {available_methods}")
            return False

    def get_available_methods(self):
        """Get list of available capture methods"""
        methods = ["auto"]
        
        print(f"🔍 Checking available methods:")
        print(f"  HAS_DXCAM: {HAS_DXCAM}, dxcam_camera: {self.dxcam_camera is not None}")
        print(f"  HAS_FAST_CTYPES: {HAS_FAST_CTYPES}, fast_ctypes_capture: {self.fast_ctypes_capture is not None}")
        print(f"  HAS_MSS: {HAS_MSS}")
        print(f"  HAS_BETTERCAM: {HAS_BETTERCAM}, bettercam_camera: {self.bettercam_camera is not None}")
        print(f"  HAS_D3DSHOT: {HAS_D3DSHOT}, d3dshot_camera: {self.d3dshot_camera is not None}")
        print(f"  HAS_BETTERCAM: {HAS_BETTERCAM}, bettercam_camera: {self.bettercam_camera is not None}")
        print(f"  HAS_D3DSHOT: {HAS_D3DSHOT}, d3dshot_camera: {self.d3dshot_camera is not None}")
        
        if HAS_DXCAM:
            methods.append("dxcam")
            print("  ✅ Added dxcam")
        if HAS_FAST_CTYPES:
            methods.append("fast_ctypes")
            print("  ✅ Added fast_ctypes")
        if HAS_MSS:
            methods.append("mss")
            print("  ✅ Added mss")
        if HAS_BETTERCAM:
            methods.append("bettercam")
            print("  ✅ Added bettercam")
        if HAS_D3DSHOT:
            methods.append("d3dshot")
            print("  ✅ Added d3dshot")
        # Expose WinRT option if either winrt is available or PIL ImageGrab fallback can be used
        if HAS_WINRT or HAS_PIL:
            if "winrt" not in methods:
                methods.append("winrt")
                print("  ✅ Added winrt")
        if HAS_D3DSHOT and self.d3dshot_camera:
            methods.append("d3dshot")
            print("  ✅ Added d3dshot")
            
        # de-dup and keep a stable order preference
        pref = ["dxcam", "fast_ctypes", "mss", "bettercam", "d3dshot", "winrt"]
        methods = [m for m in pref if m in dict.fromkeys(methods)]
        print(f"📋 Available methods: {methods}")
        return methods

    def get_current_method(self):
        """Get current capture method"""
        return self.capture_method

    def get_active_method(self):
        """Get the method that's actually being used (may differ from set method)"""
        return self.active_capture_method

    def get_capture_stats(self):
        """Get capture performance statistics"""
        return {
            'current_fps': self.capture_stats['current_fps'],
            'active_method': self.active_capture_method,
            'set_method': self.capture_method,
            'method_switches': self.capture_stats['method_switches'],
            'is_working': self.active_capture_method != "unknown" and self.capture_stats['current_fps'] > 0
        }

    def verify_capture_method(self):
        """Verify that the selected capture method is working"""
        verification = {
            'method_requested': self.capture_method,
            'method_active': self.active_capture_method,
            'is_working': False,
            'fps': self.capture_stats['current_fps'],
            'status': 'unknown'
        }
        
        if self.active_capture_method == "unknown":
            verification['status'] = 'not_working'
        elif self.capture_stats['current_fps'] == 0:
            verification['status'] = 'no_frames'
        elif self.capture_method == "auto":
            verification['status'] = 'auto_selected'
            verification['is_working'] = True
        elif self.active_capture_method == self.capture_method:
            verification['status'] = 'working_correctly'
            verification['is_working'] = True
        else:
            verification['status'] = 'not_working'
            verification['is_working'] = False
            
        return verification
    def set_performance_mode(self, enabled: bool, region: str, scale_div: int):
        self.perf_enabled = bool(enabled)
        self.perf_region = str(region or 'full')
        try:
            v = int(scale_div)
            self.perf_scale_div = v if v >= 1 else 1
        except Exception:
            self.perf_scale_div = 1
        if self.perf_region != 'custom':
            self._custom_rect_norm = None

    def _roi_norm_to_pixels(self, width: int, height: int):
        """Return (l, t, r, b) in pixels from current perf settings; None if not active."""
        if not getattr(self, 'perf_enabled', False):
            return None

        def center_box(scale: float):
            w = int(width * scale); h = int(height * scale)
            x = max(0, (width - w) // 2); y = max(0, (height - h) // 2)
            return (x, y, x + w, y + h)

        try:
            if self.perf_region == 'center_0.75':
                return center_box(0.75)
            if self.perf_region == 'center_0.5':
                return center_box(0.5)
            if self.perf_region == 'custom' and getattr(self, '_custom_rect_norm', None):
                x0 = float(self._custom_rect_norm.get('x0', 0.0))
                y0 = float(self._custom_rect_norm.get('y0', 0.0))
                x1 = float(self._custom_rect_norm.get('x1', 1.0))
                y1 = float(self._custom_rect_norm.get('y1', 1.0))
                l = max(0, min(width,  int(x0 * width)))
                t = max(0, min(height, int(y0 * height)))
                r = max(0, min(width,  int(x1 * width)))
                b = max(0, min(height, int(y1 * height)))
                if r > l and b > t:
                    return (l, t, r, b)
        except Exception:
            pass
        return None

    def _roi_pixels_to_mss_rect(self, roi, mon):
        """Convert (l,t,r,b) relative to a monitor to MSS dict with absolute desktop coords."""
        if not roi: return None
        l, t, r, b = roi
        return {
            'left': int(mon['left'] + l),
            'top':  int(mon['top']  + t),
            'width':  int(max(1, r - l)),
            'height': int(max(1, b - t)),
        }

    def _apply_perf_region(self, frame: np.ndarray) -> np.ndarray:
        if not (self.perf_enabled and isinstance(frame, np.ndarray) and frame.ndim == 3 and frame.shape[2] in (3,4)):
            return frame
        try:
            h, w = frame.shape[:2]
            if self.perf_region == 'center_0.75':
                rh, rw = int(h*0.75), int(w*0.75)
            elif self.perf_region == 'center_0.5':
                rh, rw = int(h*0.5), int(w*0.5)
            elif self.perf_region == 'custom' and getattr(self, '_custom_rect_norm', None):
                x0 = max(0.0, min(1.0, float(self._custom_rect_norm.get('x0', 0))))
                y0 = max(0.0, min(1.0, float(self._custom_rect_norm.get('y0', 0))))
                x1 = max(0.0, min(1.0, float(self._custom_rect_norm.get('x1', 1))))
                y1 = max(0.0, min(1.0, float(self._custom_rect_norm.get('y1', 1))))
                ix0, iy0 = int(x0 * w), int(y0 * h)
                ix1, iy1 = int(x1 * w), int(y1 * h)
                ix0, iy0 = max(0, ix0), max(0, iy0)
                ix1, iy1 = min(w, ix1), min(h, iy1)
                if ix1 > ix0 and iy1 > iy0:
                    return frame[iy0:iy1, ix0:ix1, :]
                return frame
            else:
                return frame
            y0 = max(0, (h - rh)//2)
            x0 = max(0, (w - rw)//2)
            return frame[y0:y0+rh, x0:x0+rw, :]
        except Exception:
            return frame

    def set_custom_region(self, rect_norm: dict):
        try:
            self._custom_rect_norm = {
                'x0': float(rect_norm.get('x0', 0)),
                'y0': float(rect_norm.get('y0', 0)),
                'x1': float(rect_norm.get('x1', 1)),
                'y1': float(rect_norm.get('y1', 1)),
            }
            self.perf_region = 'custom'
        except Exception:
            self._custom_rect_norm = None

    def set_grayscale(self, enabled: bool):
        self.perf_grayscale = bool(enabled)

    def stop(self):
        self.is_running = False
# --- WebSocket Server --- (same as original)
class VNCServer:
    def __init__(self, port, secondary_port):
        self.port = port
        self.secondary_port = secondary_port
        self.tunnel_manager = None
        self.screen_capturer = None
        self.current_quality = 75
        self.current_fps = 60
        self.video_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.audio_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.input_clients: Set[websockets.WebSocketServerProtocol] = set()
        self.cursor_subscribers: Set[websockets.WebSocketServerProtocol] = set()
        self.cursor_broadcast_enabled = False
        self.audio_thread = None
        self.audio_queue = queue.Queue(maxsize=10)
        self.audio_running = False
        self.stop_event = asyncio.Event()
        self.loop = None
        # Audio quality targets
        self.audio_samplerate_target = 24000
        self.audio_channels_target = 1
        self.audio_blocksize_target = 960
        # Audio format state
        self.audio_samplerate = 24000
        self.audio_channels = 1
        self.audio_blocksize = 960
        self.selected_camera = None
        self.keyboard_hook_active = False
        self.keystroke_capture_enabled = False
        self.keyboard_hook = None
        # Global keyboard hook suppression flag (true only when explicitly enabled)
        self._global_hook_suppress = False           # default: don't swallow anything
        self._synth_injecting = False                # guard to bypass suppression when we inject
        # mic streaming state
        self.mic_running = False
        self.mic_stream = None
        self.mic_queue = queue.Queue(maxsize=64)
        self.mic_samplerate = 48000
        self.mic_blocksize = 960
        self.mic_channels = 1
        # Typematic repeat state for non-modifier keys
        self._repeat_keys = {}
        # Feature flags
        self.enable_legacy_alert_hotkeys = True  # (legacy Q/F8 disabled in code below; flag kept for compatibility)
        # Alert presets for host controls
        self.alert_presets = {
            'A': ("Heads up", "Please check this now."),
            'B': ("Break", "Take a short break."),
            'C': ("Call", "Join the call."),
            'D': ("Stop", "Stop and review immediately."),
        }
        # Allow overriding alert presets from environment / .env
        try:
            self._load_alert_presets_from_env()
        except Exception:
            pass
        # Custom type-to-alert capture state
        self.custom_alert_active = False
        self.custom_alert_buf = []
        self._blocked_keys_in_capture = set()
        self._custom_alert_cooldown_until = 0.0
    def _load_alert_presets_from_env(self):
        """Override alert presets (A/B/C/D) from environment variables.
        Supported keys per preset (e.g. for A):
          - ALERT_A_TITLE, ALERT_A_MESSAGE
          - ALERT_A="Title|Message" (or "Title::Message")
        """
        # Optionally load .env if python-dotenv is available
        try:
            from dotenv import load_dotenv  # type: ignore
            try:
                load_dotenv()
            except Exception:
                pass
        except Exception:
            pass
        try:
            import os as _os
        except Exception:
            _os = None
        if _os is None:
            return
        def _pair_for(key: str):
            base = f"ALERT_{key}"
            title = _os.getenv(f"{base}_TITLE")
            message = _os.getenv(f"{base}_MESSAGE")
            combined = _os.getenv(base)
            if (not title and not message) and combined:
                if '|' in combined:
                    parts = combined.split('|', 1)
                elif '::' in combined:
                    parts = combined.split('::', 1)
                else:
                    parts = [combined, '']
                title = (parts[0] or '').strip()
                message = (parts[1] or '').strip() if len(parts) > 1 else ''
            return (title, message)
        changed = []
        for k in ('A','B','C','D'):
            t, m = _pair_for(k)
            if (t and t.strip()) or (m and m.strip()):
                cur = self.alert_presets.get(k, ("Alert", k))
                new_title = (t or '').strip() or cur[0]
                new_message = (m or '').strip() or cur[1]
                self.alert_presets[k] = (new_title, new_message)
                changed.append(k)
        try:
            if changed:
                print(f"[alerts] Presets overridden from env for: {', '.join(changed)}")
            else:
                print("[alerts] Using default alert presets (no env overrides)")
        except Exception:
            pass
        self._custom_capture_hotkeys = []
        self._repeat_lock = threading.Lock()
        

    def _capture_char(self, ch: str):
        try:
            if getattr(self, "custom_alert_active", False):
                self.custom_alert_buf.append(ch)
                _log_try_ok("_capture_char.append", ch)
        except Exception as e:
            _log_except("_capture_char", e)

    def _capture_backspace(self):
        try:
            if getattr(self, "custom_alert_active", False) and self.custom_alert_buf:
                self.custom_alert_buf.pop()
                _log_try_ok("_capture_backspace.pop")
        except Exception as e:
            _log_except("_capture_backspace", e)

    async def start_server(self):
        """Start WebSocket servers on both ports"""
        # Capture the running event loop for cross-thread broadcasts
        try:
            self.loop = asyncio.get_running_loop()
        except Exception:
            self.loop = None
        print("=" * 60)
        print("🎮 VNC SERVER STARTING...")
        print(f"🌐 Primary: http://localhost:{self.port}")
        print(f"🌐 Secondary: http://localhost:{self.secondary_port}")
        print(f"🏠 Network: http://{get_local_ip()}:{self.secondary_port}")
        print("🌍 Internet: Check above for public URL")
        print("=" * 60)
        # Start host hotkey capture on server start (A/B/C/D)
        try:
            self.start_global_keyboard_hook()
        except Exception as e:
            print(f"⚠️ Keyboard hook failed to initialize: {e}")
        # If hook isn't active, start fallback poller
        try:
            if not getattr(self, 'keyboard_hook_active', False):
                self.start_host_hotkey_poller()
                print("✅ Fallback hotkey poller started (A/B/C/D)")
            else:
                print("✅ Global keyboard hook armed for host alerts (A/B/C/D)")
        except Exception:
            pass
        
        # Check if screen capturer is working
        if self.screen_capturer:
            print("✅ Screen capturer initialized")
            # Wait a moment for it to capture first frame
            await asyncio.sleep(1)
            if self.screen_capturer.latest_frame_jpeg:
                print("✅ Screen capture working - first frame captured")
            else:
                print("⚠️  Screen capture not producing frames yet")
        else:
            print("❌ Screen capturer not initialized")
        
        # Initialize tunnel manager if available
        try:
            from cursor import CloudflareTunnelManager  # type: ignore
        except Exception:
            CloudflareTunnelManager = None
        if CloudflareTunnelManager is not None:
            try:
                self.tunnel_manager = CloudflareTunnelManager(self.port)
                threading.Thread(target=self.tunnel_manager.start_primary_tunnel, daemon=True).start()
            except Exception:
                self.tunnel_manager = None
                
        # Start broadcast task for video streaming
        broadcast_task = asyncio.create_task(self.broadcast_frames())
        
        # Start cursor broadcasting task
        cursor_broadcast_task = asyncio.create_task(self.broadcast_cursor_position())
        
        # Start both servers
        primary_server = await websockets.serve(
            self.main_handler,
            "0.0.0.0",
            self.port,
            process_request=self.process_request,
        )
        
        secondary_server = await websockets.serve(
            self.main_handler,
            "0.0.0.0",
            self.secondary_port,
            process_request=self.process_request,
        )
        
        print(f"✅ VNC servers running on ports {self.port} and {self.secondary_port}")
        print("🎬 Video streaming started")
        
        # Keep servers running
        await asyncio.gather(
            primary_server.wait_closed(),
            secondary_server.wait_closed(),
            broadcast_task,
            cursor_broadcast_task
        )

    def set_tunnel_manager(self, tunnel_manager):
        """Set the tunnel manager for API access"""
        self.tunnel_manager = tunnel_manager

    def enumerate_cameras(self):
        devices = []
        try:
            # Windows: use multiple strategies to list cameras
            if sys.platform.startswith('win'):
                import subprocess, json as _json, os as _os
                ps = _os.path.join(_os.environ.get('SystemRoot','C:\\Windows'), 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
                # 1) Windows.Devices.Enumeration (covers most devices)
                try:
                    cmd = [
                        ps, '-NoProfile', '-Command',
                        "$ErrorActionPreference='SilentlyContinue'; try { Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction SilentlyContinue; $t=[Windows.Devices.Enumeration.DeviceInformation]::FindAllAsync([Windows.Devices.Enumeration.DeviceClass]::VideoCapture); $r=$t.AsTask().GetAwaiter().GetResult(); ($r | Select-Object -ExpandProperty Name) | ConvertTo-Json -Compress } catch { '[]' }"
                    ]
                    result = _run_hidden(cmd, timeout=8)
                    out = (result.stdout or '').strip()
                    if out:
                        try:
                            parsed = _json.loads(out)
                            if isinstance(parsed, list):
                                devices.extend([str(x) for x in parsed if x])
                            elif isinstance(parsed, str) and parsed not in ('[]',''):
                                devices.append(parsed)
                        except Exception:
                            for line in out.splitlines():
                                line = line.strip()
                                if line:
                                    devices.append(line)
                except Exception:
                    pass
                # 2) PnP / CIM
                candidates_cmds = [
                    [ps, '-NoProfile', '-Command', "$ErrorActionPreference='SilentlyContinue'; Get-PnpDevice -Class Camera,Image | Select-Object -ExpandProperty FriendlyName | ConvertTo-Json -Compress"],
                    [ps, '-NoProfile', '-Command', "$ErrorActionPreference='SilentlyContinue'; Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPClass -eq 'Image' -or $_.Name -match 'camera|webcam|video' } | Select-Object -ExpandProperty Name | ConvertTo-Json -Compress"],
                ]
                for cmd in candidates_cmds:
                    try:
                        result = _run_hidden(cmd, timeout=6)
                        out = (result.stdout or '').strip()
                        if result.returncode == 0 and out:
                            try:
                                parsed = _json.loads(out)
                                if isinstance(parsed, list):
                                    devices.extend([str(x) for x in parsed if x])
                                elif isinstance(parsed, str):
                                    devices.append(parsed)
                            except Exception:
                                for line in out.splitlines():
                                    line = line.strip()
                                    if line:
                                        devices.append(line)
                    except Exception:
                        continue
                # 3) Registry DirectShow devices
                if not devices:
                    try:
                        reg_cmd = [ps, '-NoProfile', '-Command', "$ErrorActionPreference='SilentlyContinue'; Get-ChildItem 'HKLM:SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\MMDevices\\VideoCapture' | ForEach-Object { Get-ItemProperty $_.PsPath } | Select-Object -ExpandProperty FriendlyName | ConvertTo-Json -Compress"]
                        result = _run_hidden(reg_cmd, timeout=6)
                        out = (result.stdout or '').strip()
                        if out:
                            try:
                                parsed = _json.loads(out)
                                if isinstance(parsed, list):
                                    devices.extend([str(x) for x in parsed if x])
                                elif isinstance(parsed, str):
                                    devices.append(parsed)
                            except Exception:
                                for line in out.splitlines():
                                    line = line.strip()
                                    if line:
                                        devices.append(line)
                    except Exception:
                        pass
                # 4) WMIC last resort
                if not devices:
                    try:
                        result = _run_hidden(['wmic', 'path', 'Win32_PnPEntity', 'where', "PNPClass='Image' or Name like '%Camera%' or Name like '%Webcam%' or Name like '%Video%'", 'get', 'Name'], timeout=6)
                        out = (result.stdout or '').strip()
                        for line in out.splitlines():
                            line = line.strip()
                            if line and line.lower()!='name':
                                devices.append(line)
                    except Exception:
                        pass
        except Exception:
            pass
        # De-duplicate while preserving order
        unique = []
        seen = set()
        for d in devices:
            if d not in seen:
                seen.add(d)
                unique.append(d)
        # Normalize names and remove audio-only entries (e.g., "Microphone (Iriun Webcam)" -> "Iriun Webcam")
        try:
            import re as _re
            normalized = []
            seen2 = set()
            for name in unique:
                if not name:
                    continue
                s = str(name).strip()
                # If it looks like a microphone wrapper, extract the inner device name
                mic_match = _re.match(r"\s*Microphone\s*\((.+?)\)\s*$", s, _re.IGNORECASE)
                if mic_match:
                    s = mic_match.group(1).strip()
                # Skip obvious audio-only devices
                if _re.search(r"\b(microphone|mic|audio|speaker|line[- ]?in|line[- ]?out|headset)\b", s, _re.IGNORECASE):
                    # If it's still audio after normalization, skip it
                    continue
                if s and s not in seen2:
                    seen2.add(s)
                    normalized.append(s)
            unique = normalized
        except Exception:
            pass
        # Do not fabricate devices; return exactly what the OS exposes
        return unique
    async def process_request(self, *args, **kwargs):
        """Process HTTP requests - compatible with websockets v10-v15.

        Accepts either (path, request_headers) or a single ServerConnection object.
        """
        path = None
        request_headers = None

        # Unpack arguments depending on websockets version
        try:
            # websockets classic style: (path, request_headers)
            if len(args) >= 2 and isinstance(args[0], str):
                path = args[0]
                request_headers = args[1]
            elif len(args) >= 2:
                # websockets v15 style: (ServerConnection, Request)
                connection, request = args[0], args[1]
                path = getattr(request, "path", None)
                request_headers = getattr(request, "headers", None)
            elif len(args) == 1:
                # Fallback: older style may pass a single connection-like object
                connection = args[0]
                # Try to resolve request then path
                request = getattr(connection, "request", None)
                if request is not None:
                    path = getattr(request, "path", None)
                    request_headers = getattr(request, "headers", None)
                else:
                    path = getattr(connection, "path", None)
                    request_headers = getattr(connection, "request_headers", None)
        except Exception:
            pass

        if not isinstance(path, str):
            path = "/"

        logging.debug("process_request: path=%s", path)

        # IMPORTANT: Do not intercept WebSocket upgrade requests. Let handshake proceed.
        try:
            hdrs = request_headers
            get = None
            try:
                get = hdrs.get  # websockets Headers
            except Exception:
                get = None
            upgrade_val = None
            connection_val = None
            if get:
                try:
                    upgrade_val = (get("Upgrade") or get("upgrade") or "").lower()
                    connection_val = (get("Connection") or get("connection") or "").lower()
                except Exception:
                    pass
            if (isinstance(upgrade_val, str) and "websocket" in upgrade_val) or (
                isinstance(connection_val, str) and "upgrade" in connection_val
            ):
                return None
        except Exception:
            # On any issue determining headers, default to allowing handshake
            return None

        # Process routes
        if path == "/":
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            try:
                headers["Permissions-Policy"] = "clipboard-read=(self), clipboard-write=(self)"
            except Exception:
                pass
            try:
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
            except Exception:
                pass
            try:
                inj = (
                    "<script>"
                    f"window.CODE_FULL={json.dumps(os.getenv('CODE_FULL','TERMINATOR'))};"
                    f"window.CODE_LIMITED={json.dumps(os.getenv('CODE_LIMITED','ADVENTURES'))};"
                    f"window.CODE_PARTIAL={json.dumps(os.getenv('CODE_PARTIAL','INNOVATION'))};"
                    f"window.CODE_LOCKDOWN={json.dumps(os.getenv('CODE_LOCKDOWN','CHALLENGER'))};"
                    f"window.CUSTOM_PASSWORD={json.dumps(os.getenv('CUSTOM_PASSWORD',''))};"
                    "(function(){\n"
                    "  function safeUnblock(){ try{ if(typeof setSuppressHostInput==='function'){ setSuppressHostInput(false); } }catch(e){} }\n"
                    "  function enableCursor(){ try{ if(typeof sendEvent==='function'){ sendEvent({ action: 'cursor_broadcast', enabled: true }); } }catch(e){} }\n"
                    "  window.addEventListener('load', function(){ safeUnblock(); enableCursor(); });\n"
                    "  document.addEventListener('keydown', function(e){ if(e.key==='Escape' && window.__snapSelecting){ safeUnblock(); } });\n"
                    "  document.addEventListener('mouseup', function(){ if(window.__snapSelecting){ safeUnblock(); } });\n"
                    "  window.addEventListener('beforeunload', function(){ safeUnblock(); });\n"
                    "})();"
                    "</script>"
                )
                html = HTML_CONTENT.replace("</head>", inj + "\n</head>")
            except Exception:
                html = HTML_CONTENT
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=html.encode("utf-8"),
            )
        elif isinstance(path, str) and path.startswith("/api/client-log"):
            try:
                # Accept simple GET with ?msg=... or POST with text body
                if request_headers is None:
                    body_bytes = b""
                else:
                    body_bytes = getattr(args[1], "body", b"") if len(args) >= 2 else b""
                from urllib.parse import urlparse, parse_qs, unquote
                parsed = urlparse(path)
                qs = parse_qs(parsed.query or "")
                msg = (qs.get("msg") or [""])[0]
                if not msg and isinstance(body_bytes, (bytes, bytearray)) and body_bytes:
                    try:
                        msg = body_bytes.decode("utf-8", "ignore")
                    except Exception:
                        msg = str(body_bytes)
                msg = unquote(msg)
                _log_try_ok("client.log", msg[:500])
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=b"{\"ok\":true}",
                )
            except Exception as e:
                _log_except("client.log", e)
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                    reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                    headers=headers,
                    body=b"{\"ok\":false}",
                )
        elif path == "/terminal.html":
            # Require splash cookie and advanced-gated cookie to access terminal UI
            try:
                cookie = request_headers.get("Cookie") if request_headers else None
            except Exception:
                cookie = None
            if not (isinstance(cookie, str) and "zadoo_splash_ok=1" in cookie and "zadoo_terminal_ok=1" in cookie):
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Not Found",
                )
            # Terminal with xterm.js and a no-CDN fallback (triple-quoted to avoid quoting issues)
            term_html = ("""
<!DOCTYPE html><html><head><meta charset='utf-8'><title>Terminal</title>
<meta name='viewport' content='width=device-width,initial-scale=1'/>
<link rel='stylesheet' href='https://cdn.jsdelivr.net/npm/xterm@5.3.0/css/xterm.css'/>
<link rel='preconnect' href='https://fonts.googleapis.com'/>
<link href='https://fonts.googleapis.com/css2?family=Fira+Code:wght@500&display=swap' rel='stylesheet'/>
<style>
  @font-face{font-family:'FiraCodeMedium';src:url('https://cdn.jsdelivr.net/gh/tonsky/FiraCode@6.2/distr/woff2/FiraCode-Medium.woff2') format('woff2');font-weight:500;font-style:normal;font-display:swap;}
  html,body{margin:0;height:100%;background:#0b0e12;color:#cfd8dc;font-family:monospace;}
  #term{display:none;height:100%;width:100%;}
  .xterm{padding:8px;}
  .xterm .xterm-viewport{background:#0b0e12;}
  .xterm-rows{color:#cfd8dc;}
  #fallback{display:block;height:100vh;box-sizing:border-box;padding:8px;}
  #out{height:calc(100% - 36px);overflow:auto;white-space:pre-wrap;background:#0b0e12;color:#cfd8dc;border:1px solid #222;padding:8px;}
  #inp{width:100%;margin-top:6px;background:#0b0e12;color:#cfd8dc;border:1px solid #222;padding:8px;}
</style></head><body>
<div id='term'></div>
<div id='fallback'><div id='out'>Loading terminal...</div><input id='inp' placeholder='Type your commands and press Enter'/></div>
<script src='https://cdn.jsdelivr.net/npm/xterm@5.3.0/lib/xterm.min.js'></script>
<script src='https://cdn.jsdelivr.net/npm/xterm-addon-fit@0.8.0/lib/xterm-addon-fit.min.js'></script>
<script>
const proto=location.protocol==='https:'?'wss':'ws'; const ws=new WebSocket(`${proto}://${location.host}/ssh`); ws.binaryType='arraybuffer';
function startFallback(){ document.getElementById('term').style.display='none'; const fb=document.getElementById('fallback'); fb.style.display='block'; const out=document.getElementById('out'); const inp=document.getElementById('inp');
  out.textContent = 'Connecting to local shell...\\n';
  ws.onopen=()=>{ out.textContent += 'Connected. Type your commands and press Enter.\\n'; };
  ws.onerror=()=>{ out.textContent += '\\n[Error] Terminal connection error.\\n'; };
  ws.onclose=()=>{ out.textContent += '\\n[Closed] Terminal disconnected.\\n'; };
  ws.onmessage=e=>{ let d=e.data; if(d instanceof ArrayBuffer){ d=new TextDecoder().decode(new Uint8Array(d)); } out.textContent+=d; if(out.scrollHeight - out.clientHeight - out.scrollTop < 8){ out.scrollTop=out.scrollHeight; } };
  inp.addEventListener('keydown',ev=>{ if(ev.key==='Enter'){ ws.send(new TextEncoder().encode(inp.value+"\\r\\n")); inp.value=''; } });
}
function startXterm(){ if(!window.Terminal||!window.FitAddon){ return startFallback(); } const term=new window.Terminal({cols:120,rows:34,convertEol:false,disableStdin:false,windowsMode:true,fontFamily:"'Fira Code', monospace",fontWeight:500,fontSize:16,theme:{
  background:'#0b0e12',
  foreground:'#f8f8f2',
  cursor:'#f8f8f2',
  selection:'#44475a',
  black:'#21222c',
  red:'#ff5555',
  green:'#50fa7b',
  yellow:'#f1fa8c',
  blue:'#bd93f9',
  magenta:'#ff79c6',
  cyan:'#8be9fd',
  white:'#f8f8f2',
  brightBlack:'#6272a4',
  brightRed:'#ff6e6e',
  brightGreen:'#69ff94',
  brightYellow:'#ffffa5',
  brightBlue:'#d6acff',
  brightMagenta:'#ff92df',
  brightCyan:'#a4ffff',
  brightWhite:'#ffffff'}});
  const fit=new window.FitAddon.FitAddon(); term.loadAddon(fit); term.open(document.getElementById('term')); try{ term.setOption('fontFamily', "'FiraCodeMedium','Fira Code', monospace"); term.setOption('fontWeight','500'); }catch(e){}; fit.fit();
  // Apply Fira Code to xterm only (no layout/color changes)
  try { term.setOption('fontFamily', "'Fira Code', monospace"); term.setOption('fontWeight', '500'); } catch(e){}
  if (document.fonts && document.fonts.ready) { document.fonts.ready.then(()=>{ try{ fit.fit(); sendSize(); }catch(e){} }); }
  function sendSize(){ try{ if(ws.readyState===1){ ws.send(JSON.stringify({type:'resize', cols:term.cols, rows:term.rows})); } }catch(e){} }
  window.addEventListener('resize',()=>{ fit.fit(); sendSize(); }); term.focus();
  // Respond to parent messages to refit when panel opens
  window.addEventListener('message', (ev)=>{ try{ if(ev.data && ev.data.type==='fit'){ fit.fit(); sendSize(); } }catch(e){} });
  // Toggle views
  document.getElementById('fallback').style.display='none';
  document.getElementById('term').style.display='block';
  term.write('\u001b[36mConnecting to local shell...\u001b[0m\\r\\n');
  ws.onopen=()=>{ term.write('\u001b[32mConnected. Type your commands...\u001b[0m\\r\\n'); sendSize(); };
  ws.onerror=()=>{ term.write('\\r\\n\u001b[31m[Error] Terminal connection error.\u001b[0m\\r\\n'); };
  ws.onmessage=e=>{ let d=e.data; if(d instanceof ArrayBuffer){ d=new TextDecoder('utf-8').decode(new Uint8Array(d)); } term.write(d); };
  ws.onclose=()=>term.write('\\r\\n\u001b[31mDisconnected\u001b[0m\\r\\n');
  term.onData(data=>{ if(ws.readyState===1){ if(data==='\x7f'){ data='\b'; } ws.send(new TextEncoder().encode(data)); } });
}
setTimeout(startXterm,50);
</script>
</body></html>
""").encode()
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=term_html,
            )
        elif path == "/video":
            return None  # Let WebSocket handler take over
        elif path == "/input":
            return None  # Let WebSocket handler take over
        elif path == "/audio":
            return None  # Let WebSocket handler take over
        elif path == "/ssh":
            # Require splash cookie to allow SSH websocket upgrades
            try:
                cookie = request_headers.get("Cookie") if request_headers else None
            except Exception:
                cookie = None
            if not (isinstance(cookie, str) and "zadoo_splash_ok=1" in cookie):
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.FORBIDDEN),
                    reason_phrase=http.HTTPStatus.FORBIDDEN.phrase,
                    headers=headers,
                    body=b"Forbidden",
                )
            return None  # WebSocket bridge for SSH
        elif path == "/webcam":
            return None  # Let WebSocket handler take over
        elif path == "/api/list-cameras":
            try:
                devices = self.enumerate_cameras()
                payload = json.dumps({ 'success': True, 'devices': devices }).encode('utf-8')
            except Exception:
                payload = json.dumps({ 'success': False, 'devices': [] }).encode('utf-8')
            headers = Headers()
            headers["Content-Type"] = "application/json; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=payload,
            )
        elif path == "/host-controls":
            page = (
"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Host Controls — Alerts</title>
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial; background:#0f1115; color:#eaeef2; margin:0; }
    .wrap { max-width: 720px; margin: 32px auto; padding: 16px; }
    h1 { font-size: 20px; font-weight:600; margin: 0 0 16px; }
    .grid { display:grid; grid-template-columns: repeat(4, minmax(0,1fr)); gap:12px; }
    button { appearance: none; border:0; border-radius:12px; padding:16px 0; font-size:18px; font-weight:700; background:#1b1f2a; color:#eaeef2; cursor:pointer; transition:transform .05s ease, background .2s ease; box-shadow: 0 1px 0 rgba(255,255,255,.05) inset, 0 6px 24px rgba(0,0,0,.4); }
    button:hover{ background:#222835; }
    button:active{ transform:scale(.98); }
    .ok { margin-top:14px; color:#9fb3c8; font-size:13px; min-height: 18px; }
  </style>
  </head>
  <body>
    <div class="wrap">
      <h1>Send Sidebar Alert to Viewers</h1>
      <div class="grid">
        <button data-code="A">A</button>
        <button data-code="B">B</button>
        <button data-code="C">C</button>
        <button data-code="D">D</button>
      </div>
      <div class="ok" id="ok"></div>
    </div>
    <script>
      const ok = document.getElementById('ok');
      document.querySelectorAll('button[data-code]').forEach(btn=>{
        btn.addEventListener('click', async ()=>{
          const code = btn.getAttribute('data-code');
          try{
            const res = await fetch('/api/alert?code='+encodeURIComponent(code));
            ok.textContent = res.ok ? '✅ Sent '+code : '❌ Failed ('+res.status+')';
          }catch(e){ ok.textContent = '❌ Network error'; }
          setTimeout(()=> ok.textContent='', 2000);
        });
      });
    </script>
  </body>
  </html>
"""
            ).strip().encode("utf-8")
            headers = Headers()
            headers["Content-Type"] = "text/html; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.OK),
                reason_phrase=http.HTTPStatus.OK.phrase,
                headers=headers,
                body=page,
            )
        elif path.startswith("/api/alert"):
            # Accept GET or POST (websockets.process_request exposes only path/headers)
            try:
                from urllib.parse import urlparse, parse_qs
                qs = parse_qs(urlparse(path).query or "")
                code = (qs.get("code") or [""])[0].upper().strip()
                _log_try_ok("api.alert.parse", code)
            except Exception:
                code = ""
                _log_except("api.alert.parse", sys.exc_info()[1])
            try:
                title, message = self.alert_presets.get(code, ("Alert", f"Code: {code}"))
                _log_try_ok("api.alert.lookup", f"{title}|{message}")
            except Exception:
                title, message = ("Alert", f"Code: {code}")
                _log_except("api.alert.lookup", sys.exc_info()[1])
            try:
                self._broadcast_controller_alert(title, message)
                resp = b'{"ok":true}'
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                _log_try_ok("api.alert.broadcast")
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=resp,
                )
            except Exception as e:
                err = ("{\"ok\":false,\"error\":\"" + str(e) + "\"}").encode("utf-8", "ignore")
                headers = Headers()
                headers["Content-Type"] = "application/json; charset=utf-8"
                _log_except("api.alert.broadcast", e)
                return WSResponse(
                    status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                    reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                    headers=headers,
                    body=err,
                )
        elif path == "/brand-header.png":
            try:
                with open(BRAND_HEADER_IMAGE_PATH, 'rb') as f:
                    img_bytes = f.read()
                headers = Headers()
                headers["Content-Type"] = "image/png"
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=img_bytes,
                )
            except Exception:
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Header image not found",
                )
        elif path == "/trigger-icon.png":
            try:
                with open(TRIGGER_ICON_IMAGE_PATH, 'rb') as f:
                    img_bytes = f.read()
                headers = Headers()
                headers["Content-Type"] = "image/png"
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=img_bytes,
                )
            except Exception:
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Trigger icon not found",
                )
        elif path == "/splash.png":
            try:
                with open(SPLASH_IMAGE_PATH, 'rb') as f:
                    img_bytes = f.read()
                headers = Headers()
                headers["Content-Type"] = "image/png"
                headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                return WSResponse(
                    status_code=int(http.HTTPStatus.OK),
                    reason_phrase=http.HTTPStatus.OK.phrase,
                    headers=headers,
                    body=img_bytes,
                )
            except Exception:
                headers = Headers()
                headers["Content-Type"] = "text/plain; charset=utf-8"
                return WSResponse(
                    status_code=int(http.HTTPStatus.NOT_FOUND),
                    reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                    headers=headers,
                    body=b"Splash image not found",
            )
        elif path.startswith("/api/ocr"):
                return await self.handle_ocr(path)
        elif path.startswith("/snapshot"):
                try:
                    import time as _t
                    t_req0 = _t.perf_counter()
                    # NEW: Parse URL for crop coordinates
                    parsed_url = urllib.parse.urlparse(path)
                    query_params = urllib.parse.parse_qs(parsed_url.query)
                    
                    rect_norm = None
                    if all(k in query_params for k in ['x0', 'y0', 'x1', 'y1']):
                        try:
                            rect_norm = {
                                'x0': float(query_params['x0'][0]),
                                'y0': float(query_params['y0'][0]),
                                'x1': float(query_params['x1'][0]),
                                'y1': float(query_params['y1'][0])
                            }
                        except (ValueError, IndexError):
                            pass # Ignore invalid coordinates
                    # Optional encode params (default to PNG to match clipboard expectations)
                    fmt = query_params.get('fmt', ['png'])[0].lower().strip()
                    if fmt not in ('jpeg', 'jpg', 'png'):
                        fmt = 'png'
                    try:
                        quality = int(query_params.get('q', [85])[0])
                    except Exception:
                        quality = 85
                    try:
                        max_w = int(query_params.get('max_w', [0])[0]) or None
                    except Exception:
                        max_w = None
                    try:
                        max_h = int(query_params.get('max_h', [0])[0]) or None
                    except Exception:
                        max_h = None
                    # Log incoming snapshot request
                    try:
                        logging.info(
                            "[snapshot.request] fmt=%s q=%s max_w=%s max_h=%s rect=%s",
                            fmt,
                            quality,
                            str(max_w),
                            str(max_h),
                            'yes' if rect_norm else 'no'
                        )
                    except Exception:
                        pass

                    img_bytes = self._generate_snapshot_png(rect_norm=rect_norm, fmt=fmt, quality=quality, max_w=max_w, max_h=max_h)
                    t_req1 = _t.perf_counter()

                    if img_bytes:
                        headers = Headers()
                        headers["Content-Type"] = ("image/png" if fmt == 'png' else "image/jpeg")
                        headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
                        try:
                            import time as _t
                            ext = ("png" if fmt == 'png' else "jpg")
                            fname = f"snap-{_t.strftime('%Y%m%d-%H%M%S')}.{ext}"
                            headers["Content-Disposition"] = f"attachment; filename=\"{fname}\""
                        except Exception:
                            pass
                        # Log end-to-end server timing before returning
                        try:
                            logging.info(
                                "[snapshot.server] total=%.2fms generate=%.2fms bytes=%.1fKB fmt=%s rect=%s",
                                (t_req1 - t_req0) * 1000.0,
                                (t_req1 - t_req0) * 1000.0,
                                len(img_bytes) / 1024.0,
                                fmt,
                                'yes' if rect_norm else 'no'
                            )
                        except Exception:
                            pass
                        return WSResponse(
                            status_code=int(http.HTTPStatus.OK),
                            reason_phrase=http.HTTPStatus.OK.phrase,
                            headers=headers,
                            body=img_bytes,
                        )
                    else:
                        headers = Headers()
                        headers["Content-Type"] = "text/plain; charset=utf-8"
                        return WSResponse(
                            status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                            reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                            headers=headers,
                            body=b"Failed to capture snapshot",
                        )
                except Exception as e:
                    logging.error("Error generating snapshot", exc_info=True)
                    headers = Headers()
                    headers["Content-Type"] = "text/plain; charset=utf-8"
                    return WSResponse(
                        status_code=int(http.HTTPStatus.INTERNAL_SERVER_ERROR),
                        reason_phrase=http.HTTPStatus.INTERNAL_SERVER_ERROR.phrase,
                        headers=headers,
                        body=b"Error generating snapshot",
                    )
        else:
            headers = Headers()
            headers["Content-Type"] = "text/plain; charset=utf-8"
            return WSResponse(
                status_code=int(http.HTTPStatus.NOT_FOUND),
                reason_phrase=http.HTTPStatus.NOT_FOUND.phrase,
                headers=headers,
                body=b"Not Found",
            )

    @log_calls("main_handler")
    async def main_handler(self, websocket):
        """Handles incoming connections and routes them (websockets v15 ServerConnection)."""
        # Support both legacy and v15 APIs
        path = getattr(websocket, 'path', None)
        if not isinstance(path, str):
            try:
                path = websocket.request.path
            except Exception:
                path = '/video'
        
        if path == "/video":
            await self.video_stream_handler(websocket)
        elif path == "/input":
            await self.input_event_handler(websocket)
        elif path == "/audio":
            await self.audio_stream_handler(websocket)
        elif path == "/mic":
            await self.mic_stream_handler(websocket)
        elif path == "/ssh":
            await self.ssh_ws_handler(websocket)
        elif path == "/webcam":
            await self.webcam_stream_handler(websocket)
        else:
            # Default to video stream handler for compatibility
            await self.video_stream_handler(websocket)

    @log_calls("serve")
    async def serve(self):
        self.loop = asyncio.get_running_loop()
        asyncio.create_task(self.broadcast_frames())
        # Ensure global keyboard hook is active so System A can trigger alerts
        try:
            self.start_global_keyboard_hook()
        except Exception:
            pass
        # Start fallback hotkey poller in case keyboard hook is unavailable (A/B/C/D only)
        try:
            self.start_host_hotkey_poller()
        except Exception:
            pass

        async def process_request(path, request_headers):
            if request_headers.get("Upgrade") != "websocket":
                # Handle API endpoints
                if path == '/api/public-url':
                    return await self.handle_get_public_url()
                elif path == '/api/refresh-tunnel':
                    return await self.handle_refresh_tunnel()
                elif path.startswith('/api/client-log'):
                    try:
                        from urllib.parse import urlparse, parse_qs, unquote
                        parsed = urlparse(path)
                        qs = parse_qs(parsed.query or "")
                        msg = (qs.get("msg") or [""])[0]
                        msg = unquote(msg)
                        _log_try_ok("client.log", msg[:500])
                        return http.HTTPStatus.OK, {"Content-Type": "application/json; charset=utf-8"}, b'{"ok":true}'
                    except Exception as e:
                        _log_except("client.log", e)
                        return http.HTTPStatus.INTERNAL_SERVER_ERROR, {"Content-Type": "application/json; charset=utf-8"}, b'{"ok":false}'
                elif path.startswith('/api/ocr'):
                    return await self.handle_ocr(path)
                elif path == '/api/set-quality':
                    return await self.handle_set_quality(request_headers)
                elif path == '/api/set-fps':
                    return await self.handle_set_fps(request_headers)
                elif path == '/':
                    return http.HTTPStatus.OK, {"Content-Type": "text/html"}, HTML_CONTENT.encode()
                else:
                    return http.HTTPStatus.NOT_FOUND, {}, b"Not Found"
            return None

    async def handle_get_public_url(self):
        """API endpoint to get current public URL"""
        try:
            if self.tunnel_manager and self.tunnel_manager.get_current_url():
                current_url = self.tunnel_manager.get_current_url()
                current_port = self.tunnel_manager.current_port
                return json.dumps({
                    'success': True, 
                    'url': current_url,
                    'port': current_port,
                    'message': f'Current public URL for port {current_port}'
                })
            else:
                return json.dumps({
                    'success': False, 
                    'error': 'No public URL available'
                })
        except Exception as e:
            return json.dumps({
                'success': False, 
                'error': f'Error getting URL: {str(e)}'
            })

    async def handle_refresh_tunnel(self):
        """API endpoint to refresh tunnel and get new URL"""
        try:
            if self.tunnel_manager:
                print("🔄 Refreshing tunnel...")
                
                # Run the tunnel restart in a separate thread to avoid blocking
                def restart_in_thread():
                    return self.tunnel_manager.refresh_tunnel()
                
                # Use run_in_executor to run the blocking operation
                loop = asyncio.get_event_loop()
                success = await loop.run_in_executor(None, restart_in_thread)
                
                if success:
                    # Wait for URL to be captured
                    for i in range(15):  # Wait up to 15 seconds
                        await asyncio.sleep(1)
                        if self.tunnel_manager.tunnel_url:
                            break
                    
                    url = self.tunnel_manager.tunnel_url or ""
                    print(f"🌍 New tunnel URL: {url}")
                else:
                    url = ""
                    print("❌ Failed to restart tunnel")
            else:
                url = ""
                print("❌ No tunnel manager available")
            return json.dumps({ 'success': bool(url), 'url': url })
        except Exception as e:
            return json.dumps({ 'success': False, 'error': str(e) })

    @log_calls("handle_ocr")
    async def handle_ocr(self, path: str):
        """API: GET /api/ocr?x0=&y0=&x1=&y1=  (normalized coords) → PaddleOCR text"""
        try:
            q = urllib.parse.parse_qs(urllib.parse.urlparse(path).query or "")
            # Optional trace toggle
            do_trace = (q.get('trace', ['0'])[0] in ('1','true','yes'))
            tracer = None
            sel = None
            if all(k in q for k in ("x0", "y0", "x1", "y1")):
                sel = {
                    "x0": float(q["x0"][0]),
                    "y0": float(q["y0"][0]),
                    "x1": float(q["x1"][0]),
                    "y1": float(q["y1"][0]),
                }

            # STEP 1: tracing setup
            print("1 done - parsed query & selection", flush=True)
            # Enable line-by-line tracing unconditionally for this request and child threads
            try:
                tracer = _make_line_tracer(include_files=(os.path.abspath(__file__),))
                sys.settrace(tracer)
                try:
                    threading.settrace(tracer)
                except Exception:
                    pass
            except Exception:
                tracer = None
            t0 = time.time()
            logging.info("[ocr.server] handle_ocr begin trace=%s sel=%s", do_trace, ('set' if sel else 'none'))
            print(f"[ocr.server] begin trace={do_trace} sel={'set' if sel else 'none'}", flush=True)

            # STEP 2: capture
            img = self._capture_screen_image(rect_norm=sel)
            t1 = time.time()
            logging.info("[ocr.server] capture done in %.2f ms", (t1 - t0) * 1000.0)
            print(f"[ocr.server] capture done in {(t1 - t0) * 1000.0:.2f} ms", flush=True)
            print("2 done - capture", flush=True)
            # STEP 3: defensive downscale (full-screen and ROI)
            # Avoid huge full-screen OCR on CPU: downscale if no ROI; also trim ROI if large
            if img is not None:
                try:
                    max_side = 1024 if sel is not None else 1280
                    w, h = img.size
                    print(f"[ocr.server] pre-resize size={w}x{h} sel={'yes' if sel else 'no'} max_side={max_side}", flush=True)
                    if max(w, h) > max_side:
                        img.thumbnail((max_side, max_side), Image.LANCZOS if hasattr(Image, 'LANCZOS') else Image.BICUBIC)
                        w2, h2 = img.size
                        print(f"[ocr.server] downscaled to {w2}x{h2}", flush=True)
                except Exception:
                    pass
            if img is None:
                payload = {"success": False, "error": "No image captured"}
                return http.HTTPStatus.OK, {
                    "Content-Type": "application/json",
                    "Access-Control-Allow-Origin": "*",
                }, json.dumps(payload).encode("utf-8")

            print("3 done - resize guards", flush=True)
            # STEP 4: pre-OCR prep
            ocr = _get_paddle_ocr()
            np_img = np.array(img.convert("RGB"))
            t2 = time.time()
            logging.info("[ocr.server] pre-OCR prep in %.2f ms", (t2 - t1) * 1000.0)
            print(f"[ocr.server] pre-OCR prep in {(t2 - t1) * 1000.0:.2f} ms", flush=True)
            print("4 done - pre-OCR prep", flush=True)

            # STEP 5: predict
            result_list = ocr.predict(
                np_img,
                use_textline_orientation=True,
                return_word_box=False
            )
            t3 = time.time()
            logging.info("[ocr.server] OCR predict took %.2f ms", (t3 - t2) * 1000.0)
            print(f"[ocr.server] OCR predict took {(t3 - t2) * 1000.0:.2f} ms", flush=True)
            print("5 done - predict", flush=True)

            # STEP 6: parse results
            lines = []
            try:
                first = result_list[0] if isinstance(result_list, list) and result_list else None
                if isinstance(first, dict):
                    texts = first.get('rec_texts') or []
                    lines.extend([t if isinstance(t, str) else (t[0] if isinstance(t, (list, tuple)) and t else '') for t in texts])
                else:
                    for block in (first or []):
                        try:
                            lines.append(block[1][0])
                        except Exception:
                            pass
            except Exception:
                pass

            text = "\n".join([s for s in lines if s]).strip()
            t4 = time.time()
            logging.info("[ocr.server] post-process took %.2f ms", (t4 - t3) * 1000.0)
            print(f"[ocr.server] post-process took {(t4 - t3) * 1000.0:.2f} ms", flush=True)
            print("6 done - parse & post-process", flush=True)
            payload = {"success": True, "text": text}
            print("7 done - response ready", flush=True)
            return http.HTTPStatus.OK, {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*",
            }, json.dumps(payload).encode("utf-8")
        except Exception as e:
            try:
                logging.error("[ocr.server] handle_ocr error", exc_info=True)
            except Exception:
                pass
            payload = {"success": False, "error": str(e)}
            return http.HTTPStatus.OK, {
                "Content-Type": "application/json",
                "Access-Control-Allow-Origin": "*",
            }, json.dumps(payload).encode("utf-8")
        finally:
            try:
                if tracer is not None:
                    sys.settrace(None)
                    try:
                        threading.settrace(None)
                    except Exception:
                        pass
            except Exception:
                pass

    async def handle_set_quality(self, request_headers):
        """API endpoint to set graphics quality"""
        # This would need request body parsing in a real implementation
        # For now, just return success
        response_data = json.dumps({"success": True}).encode()
        return http.HTTPStatus.OK, {
            "Content-Type": "application/json",
            "Access-Control-Allow-Origin": "*"
        }, response_data

    async def handle_set_fps(self, request_headers):
        """API endpoint to set FPS"""
        # This would need request body parsing in a real implementation
        # For now, just return success
        response_data = json.dumps({"success": True}).encode()
        return http.HTTPStatus.OK, {
            "Content-Type": "application/json",
            "Access-Control-Allow-Origin": "*"
        }, response_data

    @log_calls("_capture_screen_image")
    def _capture_screen_image(self, rect_norm=None):
        """One-shot snapshot; reuse instances and capture at source (dxcam region → MSS region)."""
        import time as _t
        t0 = _t.perf_counter()
        # Track which backend we actually used for logging
        try:
            self._last_snapshot_backend = 'none'
        except Exception:
            pass
        # Helper: convert normalized rect to absolute desktop rect
        def _norm_to_abs_rect():
            if not rect_norm:
                return None
            try:
                try:
                    import ctypes as _ct
                    u32 = _ct.windll.user32
                    u32.SetProcessDPIAware()
                    sw = u32.GetSystemMetrics(0)
                    sh = u32.GetSystemMetrics(1)
                except Exception:
                    import pyautogui as _pg
                    sw, sh = _pg.size()
                x0 = max(0, min(sw, int(float(rect_norm['x0']) * sw)))
                y0 = max(0, min(sh, int(float(rect_norm['y0']) * sh)))
                x1 = max(0, min(sw, int(float(rect_norm['x1']) * sw)))
                y1 = max(0, min(sh, int(float(rect_norm['y1']) * sh)))
                if x1 <= x0 or y1 <= y0:
                    return None
                return (x0, y0, x1, y1)
            except Exception:
                return None

        abs_rect = _norm_to_abs_rect()

        # Heuristic: for small ROIs, MSS region can be faster than dxcam
        prefer_mss_first = False
        if abs_rect:
            try:
                # Get desktop size for ratio
                try:
                    import ctypes as _ct2
                    u322 = _ct2.windll.user32
                    u322.SetProcessDPIAware()
                    sw2 = u322.GetSystemMetrics(0)
                    sh2 = u322.GetSystemMetrics(1)
                except Exception:
                    try:
                        import pyautogui as _pg2
                        sw2, sh2 = _pg2.size()
                    except Exception:
                        sw2, sh2 = 0, 0
                if sw2 > 0 and sh2 > 0:
                    l, t, r, b = abs_rect
                    area_ratio = ((r - l) * (b - t)) / float(sw2 * sh2)
                    prefer_mss_first = area_ratio <= 0.25
            except Exception:
                prefer_mss_first = False

        def _try_dxcam_then_none():
            try:
                cam = getattr(self, 'dxcam_camera', None)
                if cam is not None:
                    arr = cam.grab(region=abs_rect)
                    if arr is not None:
                        rgb = arr[:, :, :3][:, :, ::-1].copy(order='C')
                        try:
                            self._last_snapshot_backend = 'dxcam'
                        except Exception:
                            pass
                        # Pillow deprecation: mode parameter on fromarray will be removed in Pillow 13
                        try:
                            return Image.fromarray(rgb)
                        except Exception:
                            return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
            except Exception:
                logging.warning('Snapshot dxcam failed; falling back', exc_info=True)
            return None

        def _try_mss_then_none():
            if HAS_MSS:
                try:
                    if not hasattr(self, '_snapshot_sct') or self._snapshot_sct is None:
                        self._snapshot_sct = mss.mss()
                    sct = self._snapshot_sct
                    mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                    rect = None
                    if abs_rect:
                        l, t, r, b = abs_rect
                        rect = {'left': l, 'top': t, 'width': max(1, r - l), 'height': max(1, b - t)}
                    sct_img = sct.grab(rect or mon)
                    h, w = sct_img.height, sct_img.width
                    arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                    rgb = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                    try:
                        self._last_snapshot_backend = 'mss'
                    except Exception:
                        pass
                    try:
                        return Image.fromarray(rgb)
                    except Exception:
                        return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
                except Exception:
                    logging.warning('Snapshot MSS failed', exc_info=True)
            return None

        # Order backends based on ROI size
        if prefer_mss_first:
            img = _try_mss_then_none()
            if img is not None:
                return img
            img = _try_dxcam_then_none()
            if img is not None:
                return img
        else:
            img = _try_dxcam_then_none()
            if img is not None:
                return img
            img = _try_mss_then_none()
            if img is not None:
                return img

        # 2) fast_ctypes (full, then crop)
        if HAS_FAST_CTYPES:
            try:
                sct = getattr(self, 'fast_ctypes_capture', None)
                if sct is None:
                    self.fast_ctypes_capture = fast_ctypes_screenshots.ScreenshotOfAllMonitors()
                    sct = self.fast_ctypes_capture
                frame = sct.screenshot_monitors()
                if frame is not None:
                    img = Image.fromarray(frame)
                    if abs_rect:
                        img = img.crop(abs_rect)
                    return img
            except Exception:
                logging.warning('Snapshot fast_ctypes failed; falling back', exc_info=True)

        # 3) MSS region (reuse instance)
        if HAS_MSS:
            try:
                if not hasattr(self, '_snapshot_sct') or self._snapshot_sct is None:
                    self._snapshot_sct = mss.mss()
                sct = self._snapshot_sct
                mon = sct.monitors[1] if len(sct.monitors) > 1 else sct.monitors[0]
                rect = None
                if abs_rect:
                    l, t, r, b = abs_rect
                    rect = {'left': l, 'top': t, 'width': max(1, r - l), 'height': max(1, b - t)}
                sct_img = sct.grab(rect or mon)
                h, w = sct_img.height, sct_img.width
                arr = np.frombuffer(sct_img.bgra, dtype=np.uint8).reshape((h, w, 4))
                rgb = np.ascontiguousarray(arr[..., :3][:, :, ::-1])
                try:
                    return Image.fromarray(rgb)
                except Exception:
                    return Image.frombuffer('RGB', (rgb.shape[1], rgb.shape[0]), rgb.tobytes())
            except Exception:
                logging.warning('Snapshot MSS failed', exc_info=True)

        return None

    @log_calls("_generate_snapshot_png")
    def _generate_snapshot_png(self, rect_norm=None, fmt='png', quality=85, max_w=None, max_h=None):
        """
        Returns image bytes (PNG or JPEG) for the current (or region) snapshot.
        Supports format/quality/downscale and logs capture→convert→encode timings.
        """
        try:
            import time as _t, io as _io
            t0 = _t.perf_counter()
            img = self._capture_screen_image(rect_norm=rect_norm)
            t1 = _t.perf_counter()
            if img is None:
                return None
            # Optional downscale to limit size
            try:
                if (max_w and img.width > max_w) or (max_h and img.height > max_h):
                    # Keep aspect ratio
                    tw = img.width; th = img.height
                    rw = (max_w / tw) if max_w else 1.0
                    rh = (max_h / th) if max_h else 1.0
                    r = min(rw, rh)
                    if r < 1.0:
                        new_size = (max(1, int(tw * r)), max(1, int(th * r)))
                        img = img.resize(new_size, Image.BILINEAR)
            except Exception:
                pass
            arr = np.array(img, copy=False)
            if arr.ndim == 3 and arr.shape[2] == 4:
                arr = arr[:, :, :3]
            t2 = _t.perf_counter()
            out = None
            if fmt == 'png' and HAS_IMAGECODECS:
                try:
                    out = imagecodecs.png_encode(arr, level=0)
                except Exception:
                    out = None
            if out is None:
                buf = _io.BytesIO()
                try:
                    if fmt == 'png':
                        img.save(buf, format='PNG', optimize=False, compress_level=0)
                    else:
                        # JPEG fast path via imagecodecs not used here to keep code simple; Pillow is fine
                        img.save(buf, format='JPEG', quality=max(1, min(95, int(quality))))
                    out = buf.getvalue()
                except Exception:
                    out = None
            t3 = _t.perf_counter()
            logging.info(
                "[snapshot.timing] total=%.2fms | capture=%.2fms convert=%.2fms encode=%.2fms | size=%dx%d bytes=%.1fKB backend=%s fmt=%s rect=%s",
                (t3 - t0) * 1000.0,
                (t1 - t0) * 1000.0,
                (t2 - t1) * 1000.0,
                (t3 - t2) * 1000.0,
                getattr(img, 'width', 0), getattr(img, 'height', 0),
                0.0 if not out else (len(out) / 1024.0),
                getattr(self, '_last_snapshot_backend', 'unknown'),
                fmt,
                'yes' if rect_norm else 'no'
            )
            return out
        except Exception:
            logging.error("Failed to generate snapshot PNG", exc_info=True)
            return None

    def stop(self):
        if self.loop:
            self.loop.call_soon_threadsafe(self.stop_event.set)
    def _start_audio_capture(self, want_samplerate=48000, packet_frames=960, want_channels=2, mono_method="left"):
        """
        Robust SYSTEM-audio capture (loopback) for Windows 11 using python-soundcard.
        - No dependency on sounddevice.WasapiSettings(loopback=...)
        - Jitter buffer to emit exact 20ms frames (960 @ 48k)
        - Auto-rebind when default output changes (speakers/headphones)
        """
        if getattr(self, "_audio_running", False):
            return
        self._audio_running = True
        self._audio_backend = "soundcard_loopback"
        self._audio_thread = None
        self._audio_stream = None
        self._audio_dev_name = None
        self._audio_sr = want_samplerate

        log = logging.getLogger("sysaudio")
        ring = deque()
        ring_len = 0

        def _to_safe_mono(x: np.ndarray, method="left"):
            if x.ndim == 1 or x.shape[1] == 1:
                return x.reshape(-1, 1)
            if method == "left":
                return x[:, :1]
            if method == "right":
                return x[:, 1:2]
            return ((x[:, :1] + x[:, 1:2]) * 0.5).astype(np.float32)

        def _pick_loopback():
            spk = sc.default_speaker()
            loopback = None
            try:
                for mic in sc.all_microphones(include_loopback=True):
                    if getattr(mic, "isloopback", False) and (spk.name.split(" (")[0] in mic.name or getattr(spk, 'id', None) in getattr(mic, 'id', '')):
                        loopback = mic
                        break
                if loopback is None:
                    for mic in sc.all_microphones(include_loopback=True):
                        if getattr(mic, "isloopback", False):
                            loopback = mic
                            break
            except Exception:
                loopback = None
            return spk, loopback

        stop_evt = threading.Event()

        def _audio_worker():
            nonlocal ring_len
            current_loop = None
            last_pick = 0.0
            try:
                while self._audio_running and not stop_evt.is_set():
                    now = time.time()
                    if current_loop is None or (now - last_pick) > 1.5:
                        last_pick = now
                        spk, loop = _pick_loopback()
                        if loop is not None and loop != current_loop:
                            try:
                                if self._audio_stream is not None:
                                    self._audio_stream.__exit__(None, None, None)
                            except Exception:
                                pass
                            try:
                                self._audio_stream = loop.recorder(samplerate=want_samplerate, channels=2)
                                self._audio_stream.__enter__()
                                current_loop = loop
                                log.info(f"🎵 Loopback device: {loop.name} @ {want_samplerate} Hz")
                            except Exception as e:
                                log.warning(f"Loopback open failed ({e}); retrying...")
                                current_loop = None
                                time.sleep(0.2)
                                continue

                    if self._audio_stream is None:
                        time.sleep(0.01)
                        continue

                    try:
                        data = self._audio_stream.record(numframes=packet_frames)
                        # ensure 2-D float32
                        if data.ndim == 1:
                            data = data.reshape(-1, 1)
                        if want_channels == 1:
                            data = _to_safe_mono(data, mono_method)
                        else:
                            if data.shape[1] == 1 and want_channels == 2:
                                data = np.repeat(data, 2, axis=1)
                        ring.append(data)
                        ring_len += data.shape[0]

                        while ring_len >= packet_frames:
                            need = packet_frames
                            chunks = []
                            while need > 0:
                                block = ring[0]
                                if block.shape[0] <= need:
                                    chunks.append(block)
                                    ring.popleft()
                                    ring_len -= block.shape[0]
                                    need -= block.shape[0]
                                else:
                                    chunks.append(block[:need])
                                    ring[0] = block[need:]
                                    ring_len -= need
                                    need = 0
                            out = np.vstack(chunks).astype(np.float32, copy=False)
                            # Convert to s16le mono (if needed) and enqueue for /audio
                            try:
                                mono = out[:, 0] if out.ndim == 2 else out.reshape(-1)
                                pcm = (np.clip(mono, -1.0, 1.0) * 32767.0).astype(np.int16).tobytes()
                                if not self.audio_queue.full():
                                    self.audio_queue.put(pcm)
                            except Exception as e:
                                log.debug(f"sender err: {e}")
                    except Exception as e:
                        try:
                            if self._audio_stream is not None:
                                self._audio_stream.__exit__(None, None, None)
                        except Exception:
                            pass
                        self._audio_stream = None
                        current_loop = None
                        time.sleep(0.05)
                        continue
            finally:
                try:
                    if self._audio_stream is not None:
                        self._audio_stream.__exit__(None, None, None)
                except Exception:
                    pass
                self._audio_stream = None
                self._audio_running = False
                log.info("🎵 System-audio worker stopped")

        t = threading.Thread(target=_audio_worker, daemon=True)
        t.start()
        self._audio_thread = t

    def _stop_audio_capture(self):
        self._audio_running = False
        try:
            if getattr(self, '_audio_thread', None) and self._audio_thread.is_alive():
                self._audio_thread.join(timeout=0.5)
        except Exception:
                pass
        try:
            if getattr(self, '_audio_stream', None):
                if hasattr(self._audio_stream, 'stop'):
                    self._audio_stream.stop()
                if hasattr(self._audio_stream, 'close'):
                    self._audio_stream.close()
                if hasattr(self._audio_stream, '__exit__'):
                    try:
                        self._audio_stream.__exit__(None, None, None)
                    except Exception:
                        pass
        except Exception:
            pass
        self._audio_stream = None
        self._audio_thread = None
        logging.info("🎵 Audio capture stopped")

    def _start_mic_capture(self, samplerate=48000, blocksize=960, channels=1):
        """
        Microphone of System A using sounddevice InputStream (no WASAPI special args).
        Always sends mono s16le frames of 'blocksize' samples at 'samplerate'.
        """
        if self.mic_running:
            return True
        self.mic_running = True
        self.mic_samplerate = int(samplerate)
        self.mic_blocksize = int(blocksize)
        self.mic_channels = 1

        def mic_callback(indata, frames, time_info, status):
            try:
                if status:
                    logging.debug(f"Mic status: {status}")
                x = indata
                if x.ndim == 2 and x.shape[1] > 1:
                    x = x[:, 0]
                x = (np.clip(x, -1.0, 1.0) * 32767.0).astype(np.int16)
                if not self.mic_queue.full():
                    self.mic_queue.put(x.tobytes())
            except Exception:
                logging.error("Mic callback error", exc_info=True)

        try:
            self.mic_stream = sd.InputStream(
                samplerate=self.mic_samplerate,
                channels=2,
                dtype='float32',
                blocksize=self.mic_blocksize,
                callback=mic_callback,
                device=None,
                latency='low'
            )
            self.mic_stream.start()
            print(f"🎙️ Mic capture started @ {self.mic_samplerate} Hz (block={self.mic_blocksize})")
            return True
        except Exception:
            logging.error("Failed to start microphone capture", exc_info=True)
            self.mic_running = False
            self.mic_stream = None
            return False

    def _stop_mic_capture(self):
        try:
            self.mic_running = False
            if self.mic_stream:
                try:
                    self.mic_stream.stop()
                except Exception:
                    pass
                try:
                    self.mic_stream.close()
                except Exception:
                    pass
            self.mic_stream = None
            with self.mic_queue.mutex:
                self.mic_queue.queue.clear()
        except Exception:
            logging.error("Failed to stop mic capture", exc_info=True)
    async def audio_stream_handler(self, websocket: websockets.WebSocketServerProtocol):
        print(f"🔊 New audio client connected from {websocket.remote_address}")
        self.audio_clients.add(websocket)

        # Start capture on first client
        if not self.audio_running:
            self._start_audio_capture()

        try:
            # Simple header to let client know PCM format
            header = json.dumps({
                'type': 'audio_format',
                'codec': 'pcm_s16le',
                'samplerate': 48000,
                'channels': 1,
                'blocksize': 960
            }).encode()
            await websocket.send(header)

            while True:
                try:
                    chunk = await asyncio.get_event_loop().run_in_executor(None, self.audio_queue.get)
                except Exception:
                    chunk = None
                if chunk is None:
                    await asyncio.sleep(0.005)
                    continue
                try:
                    await websocket.send(chunk)
                except websockets.exceptions.ConnectionClosed:
                    break
        except Exception as e:
            print(f"Error in audio_stream_handler: {e}")
        finally:
            self.audio_clients.discard(websocket)
            print(f"🔊 Removed audio client, {len(self.audio_clients)} clients remaining")
            if not self.audio_clients and self.audio_running:
                try:
                    if hasattr(self, '_audio_stream') and self._audio_stream:
                        try:
                            self._audio_stream.stop(); self._audio_stream.close()
                        except Exception:
                            pass
                except Exception:
                    pass
                self.audio_running = False
                with self.audio_queue.mutex:
                    self.audio_queue.queue.clear()
                print("🎵 Audio capture stopped (no clients)")

    async def mic_stream_handler(self, websocket: websockets.WebSocketServerProtocol):
        print(f"🎤 New mic client connected from {getattr(websocket, 'remote_address', None)}")
        # Send format header mirroring /audio
        hdr = {
            "type": "audio_format",
            "samplerate": int(self.mic_samplerate),
            "channels": 1,
            "samplefmt": "s16le",
            "blocksize": int(self.mic_blocksize)
        }
        try:
            await websocket.send(json.dumps(hdr).encode('utf-8'))
            print(f"🎤 Mic header sent: sr={hdr['samplerate']} ch={hdr['channels']} fmt={hdr['samplefmt']} block={hdr['blocksize']}")

            if not self.mic_running:
                ok = self._start_mic_capture(samplerate=48000, blocksize=960, channels=1)
                if not ok:
                    print("❌ Mic open failed")
                    try:
                        await websocket.close(code=1011, reason="Mic open failed")
                    except Exception:
                        pass
                    return

            while True:
                try:
                    chunk = await asyncio.get_event_loop().run_in_executor(None, self.mic_queue.get)
                except Exception:
                    chunk = None
                if not chunk:
                    await asyncio.sleep(0.005)
                    continue
                try:
                    await websocket.send(chunk)
                except websockets.exceptions.ConnectionClosed:
                    break
        except Exception as e:
            print(f"Error in mic_stream_handler: {e}")
        finally:
            # Stop mic when client disconnects
            self._stop_mic_capture()
            print("🎤 Mic client disconnected; mic capture stopped")

    async def video_stream_handler(self, websocket):
        print(f"📺 New video client connected from {websocket.remote_address}")
        self.video_clients.add(websocket)
        
        try:
            # Keep connection alive and handle any incoming messages
            async for message in websocket:
                try:
                    # Handle any control messages for video stream
                    event = json.loads(message)
                    action = event.get('action')
                    
                    if action == 'refresh_tunnel':
                        await self.handle_refresh_via_websocket(websocket)
                    elif action == 'get_public_url':
                        await self.handle_get_url_via_websocket(websocket)
                    elif action == 'set_quality':
                        value = event.get('value', 75)
                        self.current_quality = value
                        if self.screen_capturer:
                            self.screen_capturer.quality = value
                        print(f"🎨 Quality set to: {value}%")
                    elif action == 'set_fps':
                        value = event.get('value', 30)
                        self.current_fps = value
                        if self.screen_capturer:
                            self.screen_capturer.fps = value
                        print(f"🎬 FPS set to: {value}")
                    elif action == 'set_capture_method':
                        method = event.get('method', 'auto')
                        if self.screen_capturer:
                            success = self.screen_capturer.set_capture_method(method)
                            if success:
                                print(f"📹 Capture method changed to: {method}")
                            else:
                                print(f"❌ Failed to set capture method to: {method}")
                        else:
                            print("❌ Screen capturer not available")
                    elif action == 'get_available_capture_methods':
                        if self.screen_capturer:
                            methods = self.screen_capturer.get_available_methods()
                            current = self.screen_capturer.get_current_method()
                            await websocket.send(json.dumps({
                                'type': 'available_capture_methods',
                                'methods': methods,
                                'current': current
                            }))
                        else:
                            await websocket.send(json.dumps({
                                'type': 'available_capture_methods',
                                'methods': ['auto'],
                                'current': 'auto'
                            }))
                    elif action == 'set_performance':
                        enabled = bool(event.get('enabled'))
                        region = str(event.get('region') or 'full')
                        scale_div = int(event.get('scale_div') or 1)
                        rect_norm = event.get('rect_norm')
                        grayscale = bool(event.get('grayscale'))
                        if self.screen_capturer:
                            self.screen_capturer.set_performance_mode(enabled, region, scale_div)
                            if rect_norm and isinstance(rect_norm, dict):
                                self.screen_capturer.set_custom_region(rect_norm)
                            self.screen_capturer.set_grayscale(grayscale)
                            print(f"⚙️ Performance mode: enabled={enabled} region={region} scale_div={scale_div} gray={grayscale} custom={bool(rect_norm)}")
                    elif action == 'get_capture_stats':
                        if self.screen_capturer:
                            stats = self.screen_capturer.get_capture_stats()
                            await websocket.send(json.dumps({
                                'type': 'capture_stats',
                                'stats': stats
                            }))
                        else:
                            await websocket.send(json.dumps({
                                'type': 'capture_stats',
                                'stats': {'is_working': False, 'current_fps': 0}
                            }))
                    elif action == 'verify_capture_method':
                        if self.screen_capturer:
                            verification = self.screen_capturer.verify_capture_method()
                            await websocket.send(json.dumps({
                                'type': 'capture_verification',
                                'verification': verification
                            }))
                        else:
                            await websocket.send(json.dumps({
                                'type': 'capture_verification',
                                'verification': {'is_working': False, 'status': 'no_capturer'}
                            }))
                    elif action == 'set_audio_quality':
                        value = int(event.get('value', 50))
                        # Map 0..100 to samplerate
                        if value <= 25:
                            self.audio_samplerate_target = 16000
                        elif value <= 50:
                            self.audio_samplerate_target = 24000
                        elif value <= 75:
                            self.audio_samplerate_target = 32000
                        else:
                            self.audio_samplerate_target = 48000
                        # Restart capture with new rate on next client connect
                        print(f"🎧 Audio quality set: {self.audio_samplerate_target} Hz")
                    elif action == 'toggle_keystroke_capture':
                        enabled = event.get('enabled', False)
                        if enabled:
                            self.enable_keystroke_capture()
                        else:
                            self.disable_keystroke_capture()
                        await websocket.send(json.dumps({
                            'type': 'keystroke_capture_status',
                            'enabled': self.keystroke_capture_enabled
                        }))
                    else:
                        # Handle other input events
                        self.process_event(event, websocket)
                except Exception as e:
                    print(f"Error handling video client message: {e}")

        except websockets.exceptions.ConnectionClosed:
            print(f"📺 Video client {websocket.remote_address} disconnected")
        except Exception as e:
            print(f"Error in video_stream_handler: {e}")
        finally:
            self.video_clients.discard(websocket)
            print(f"📺 Removed video client, {len(self.video_clients)} clients remaining")
    async def webcam_stream_handler(self, websocket):
        """Stream JPEG frames from the server's webcam (DirectShow on Windows) to the client."""
        print(f"🎥 Webcam client connected from {websocket.remote_address}")
        # Prefer PyAV (FFmpeg) on Windows via DirectShow
        if not HAS_AV:
            print("❌ PyAV not available; webcam streaming disabled")
            try:
                await websocket.send(b"")
            except Exception:
                pass
            return

        container = None
        player = None
        try:
            # First message may include selected device from client
            try:
                first = await asyncio.wait_for(websocket.recv(), timeout=0.2)
                try:
                    data = json.loads(first) if isinstance(first, str) else json.loads(first.decode('utf-8','ignore'))
                    if isinstance(data, dict) and data.get('type') == 'select_camera' and data.get('device'):
                        self.selected_camera = str(data.get('device'))
                        print(f"🎥 Client selected camera: {self.selected_camera}")
                except Exception:
                    pass
            except asyncio.TimeoutError:
                pass
            open_ok = False
            is_windows = sys.platform.startswith('win')
            # Try common device names for Windows DirectShow
            if is_windows:
                # Try multiple option combinations for broader compatibility
                option_sets = [
                    None,
                    { 'video_size': '640x480' },
                    { 'framerate': '15', 'video_size': '640x480' },
                    { 'framerate': '30', 'video_size': '640x480' },
                    { 'framerate': '30', 'video_size': '1280x720' },
                    { 'framerate': '30', 'video_size': '1920x1080' },
                ]
                candidates = []
                env_name = os.environ.get('CAMERA_NAME')
                if env_name:
                    candidates.append(env_name)
                # Include selected camera and a normalized variant if it was an audio label like "Microphone (Name)"
                if self.selected_camera:
                    candidates.insert(0, self.selected_camera)
                    try:
                        import re as _re
                        m = _re.match(r"\s*Microphone\s*\((.+?)\)\s*$", str(self.selected_camera), _re.IGNORECASE)
                        if m:
                            inside = m.group(1).strip()
                            if inside and inside not in candidates:
                                candidates.insert(0, inside)
                    except Exception:
                        pass
                candidates.extend([
                    'Iriun Webcam',
                    'Integrated Camera',
                    'USB Video Device',
                    'HD Webcam',
                    'HP Wide Vision FHD Camera',
                    'Logitech',
                    'OBS Virtual Camera'
                ])
                # Try PyAV DirectShow device strings
                for name in candidates:
                    for opts in option_sets:
                        try:
                            device = f"video={name}"
                            if opts is None:
                                print(f"🎯 Trying dshow open: {device} opts=None")
                                container = av.open(device, format='dshow')
                            else:
                                print(f"🎯 Trying dshow open: {device} opts={opts}")
                                container = av.open(device, format='dshow', options=opts)
                            open_ok = True
                            print(f"✅ Opened webcam via dshow device: {device} opts={opts}")
                            break
                        except Exception as e:
                            print(f"⚠️  Open failed: {device} opts={opts} err={e}")
                            container = None
                            continue
                    if open_ok:
                        break
                # Fallback attempts with generic device strings
                if not open_ok:
                    for generic in ("video=0", "0", "video=1", "1"):
                        for opts in option_sets:
                            try:
                                if opts is None:
                                    print(f"🎯 Trying dshow open: {generic} opts=None")
                                    container = av.open(generic, format='dshow')
                                else:
                                    print(f"🎯 Trying dshow open: {generic} opts={opts}")
                                    container = av.open(generic, format='dshow', options=opts)
                                open_ok = True
                                print(f"✅ Opened webcam via dshow generic: {generic} opts={opts}")
                                break
                            except Exception as e:
                                print(f"⚠️  Open failed: {generic} opts={opts} err={e}")
                                container = None
                                continue
                        if open_ok:
                            break
                # Non-Windows simple attempt (may not be used in this environment)
                

            if not open_ok and HAS_AIORTC:
                # Fallback to aiortc MediaPlayer with dshow
                try:
                    from aiortc.contrib.media import MediaPlayer
                    for name in candidates:
                        for opts in option_sets:
                            try:
                                if opts is None:
                                    print(f"🎯 Trying MediaPlayer open: video={name} opts=None")
                                    player = MediaPlayer(f"video={name}", format='dshow')
                                else:
                                    print(f"🎯 Trying MediaPlayer open: video={name} opts={opts}")
                                    player = MediaPlayer(f"video={name}", format='dshow', options=opts)
                                print(f"✅ Opened webcam via aiortc MediaPlayer: video={name} opts={opts}")
                                break
                            except Exception as e:
                                print(f"⚠️  MediaPlayer open failed: video={name} opts={opts} err={e}")
                                player = None
                                continue
                        if player is not None:
                            break
                    if player is None:
                        for generic in ("video=0", "video=1"):
                            for opts in option_sets:
                                try:
                                    if opts is None:
                                        print(f"🎯 Trying MediaPlayer open: {generic} opts=None")
                                        player = MediaPlayer(generic, format='dshow')
                                    else:
                                        print(f"🎯 Trying MediaPlayer open: {generic} opts={opts}")
                                        player = MediaPlayer(generic, format='dshow', options=opts)
                                    print(f"✅ Opened webcam via aiortc MediaPlayer: {generic} opts={opts}")
                                    break
                                except Exception as e:
                                    print(f"⚠️  MediaPlayer open failed: {generic} opts={opts} err={e}")
                                    player = None
                                    continue
                            if player is not None:
                                break
                except Exception as e:
                    player = None

            if container is None and player is None:
                print("❌ No webcam device could be opened")
                try:
                    await websocket.send(b"")
                except Exception:
                    pass
                return

            # Send frames
            if container is not None:
                stream = None
                try:
                    stream = next((s for s in container.streams if s.type == 'video'), None)
                except Exception:
                    stream = None
                if stream is not None:
                    try:
                        stream.thread_type = 'AUTO'
                    except Exception:
                        pass
                    # Decode using stream index; passing a stream object causes errors
                    try:
                        video_stream_index = 0 if stream is None else int(getattr(stream, 'index', 0))
                    except Exception:
                        video_stream_index = 0
                    for frame in container.decode(video=video_stream_index):
                        try:
                            img = frame.to_image()
                            buf = io.BytesIO()
                            img.save(buf, format='JPEG', quality=70)
                            await websocket.send(buf.getvalue())
                        except websockets.exceptions.ConnectionClosed:
                            break
                        except Exception:
                            pass
                        await asyncio.sleep(0.05)
            else:
                # aiortc MediaPlayer path
                video_track = getattr(player, 'video', None)
                if video_track is None:
                    print("❌ MediaPlayer has no video track")
                    return
                while True:
                    try:
                        frame = await video_track.recv()
                        img = frame.to_image()
                        buf = io.BytesIO()
                        img.save(buf, format='JPEG', quality=70)
                        await websocket.send(buf.getvalue())
                    except websockets.exceptions.ConnectionClosed:
                        break
                    except Exception:
                        await asyncio.sleep(0.05)
                        continue
        except Exception as e:
            print(f"Error in webcam_stream_handler: {e}")
        finally:
            try:
                if container is not None:
                    container.close()
            except Exception:
                pass
            try:
                if player is not None:
                    player.audio and player.audio.stop()
                    player.video and player.video.stop()
            except Exception:
                pass

    async def ssh_ws_handler(self, websocket):
        """Route immediately to a local interactive shell for responsiveness on Windows."""
        await self.local_shell_ws_handler(websocket)
        return

    async def local_shell_ws_handler(self, websocket):
        """Spawn a local shell (PowerShell/cmd) and bridge it over the websocket."""
        import shutil
        use_pty = False
        # Prefer Windows PowerShell for a full shell experience; fallback to cmd.exe
        ps_path = os.path.join(os.environ.get('SystemRoot', 'C\\Windows'), 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe')
        cmd_path = os.path.join(os.environ.get('SystemRoot', 'C\\Windows'), 'System32', 'cmd.exe')
        if not os.path.exists(cmd_path):
            cmd_path = os.environ.get('ComSpec', shutil.which('cmd')) or os.path.join(os.environ.get('SystemRoot', 'C\\Windows'), 'System32', 'cmd.exe')
        try:
            # Use a Windows PTY when available for correct interactive behavior (Backspace, arrow keys)
            if 'HAS_WINPTY' in globals() and HAS_WINPTY:
                try:
                    from winpty import PtyProcess
                    shell_cmd = ps_path if os.path.exists(ps_path) else cmd_path
                    # Start with a sane default size matching the client
                    proc = PtyProcess.spawn(shell_cmd, dimensions=(34, 120))
                    proc_writer = proc
                    proc_reader = proc
                    use_pty = True
                except Exception:
                    proc = None
                    use_pty = False

            if not use_pty:
                shell_cmd = [ps_path, '-NoLogo', '-NoExit'] if os.path.exists(ps_path) else [cmd_path, '/K', 'chcp 65001']
                proc = subprocess.Popen(
                        shell_cmd,
                    stdin=subprocess.PIPE,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    bufsize=0,
                    creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0)
                )
            proc_writer = proc
            proc_reader = proc
            # Proactively show something
            try:
                await websocket.send("Connected to local shell. Type commands and press Enter.\\r\\n")
            except Exception:
                pass
            # Trigger prompt output
            try:
                if use_pty:
                    proc_writer.write("\r\n")
                else:
                    proc_writer.stdin.write(b"\r\n")
                    proc_writer.stdin.flush()
            except Exception:
                pass
        except Exception as e:
            try:
                await websocket.send(json.dumps({'type': 'error', 'message': f'Local shell failed: {e}'}))
            except Exception:
                pass
            return

        stop_flag = False

        async def ws_to_proc():
            nonlocal stop_flag
            try:
                async for msg in websocket:
                    try:
                        if isinstance(msg, str):
                            data = msg
                        else:
                            data = msg.decode('utf-8', errors='ignore')
                        # Handle terminal resize from client
                        if data and data.startswith('{'):
                            try:
                                evt = json.loads(data)
                                if evt.get('type')=='resize':
                                    cols = int(evt.get('cols', 120))
                                    rows = int(evt.get('rows', 34))
                                    if use_pty:
                                        try:
                                            proc_writer.set_size(rows, cols)
                                        except Exception:
                                            pass
                                    continue
                            except Exception:
                                pass
                        
                        # Normalize input: map DEL->BS always; only expand CR to CRLF for non-PTY
                        if use_pty:
                            data = data.replace('\x7f', '\b')
                            proc_writer.write(data)
                        else:
                            data = data.replace('\r', '\r\n').replace('\x7f', '\b')
                            proc_writer.stdin.write(data.encode('utf-8', errors='ignore'))
                            proc_writer.stdin.flush()
                    except Exception:
                        break
            finally:
                stop_flag = True

        async def proc_to_ws():
            nonlocal stop_flag
            loop = asyncio.get_event_loop()
            import locale
            import re
            enc = 'utf-8'
            poll_alive = (lambda: (proc.isalive() if use_pty else proc.poll() is None))
            while not stop_flag and poll_alive():
                try:
                    # Read in small chunks for responsiveness
                    if use_pty:
                        chunk = await loop.run_in_executor(None, proc_reader.read, 256)
                    else:
                        chunk = await loop.run_in_executor(None, proc_reader.stdout.read, 256)
                    if not chunk:
                        await asyncio.sleep(0.02)
                        continue
                    if isinstance(chunk, bytes):
                        try:
                            text = chunk.decode(enc, errors='ignore')
                        except Exception:
                            text = chunk.decode(locale.getpreferredencoding(False) or 'utf-8', errors='ignore')
                    else:
                        # winpty returns str
                        text = chunk
                    # Normalize bare CR from PTY to CRLF to avoid cursor overlays
                    if use_pty and text:
                        text = re.sub(r"\r(?!\n)", "\r\n", text)
                    await websocket.send(text)
                except Exception:
                    break

        sender = asyncio.create_task(ws_to_proc())
        receiver = asyncio.create_task(proc_to_ws())
        try:
            # Keep the session open as long as either side is active
            await asyncio.gather(sender, receiver)
        finally:
            for t in (sender, receiver):
                if not t.done():
                    t.cancel()
        try:
            if proc:
                if use_pty:
                    try:
                        proc.close()
                    except Exception:
                        pass
                else:
                    if proc.poll() is None:
                        try:
                            proc.terminate()
                            proc.wait(timeout=2)
                        except Exception:
                            try:
                                proc.kill()
                            except Exception:
                                pass
        except Exception:
            pass

    async def broadcast_frames(self):
        """Broadcast video frames to all connected video clients"""
        frame_count = 0
        last_debug = 0
        
        while not self.stop_event.is_set():
            if self.screen_capturer and self.screen_capturer.latest_frame_jpeg:
                frame = self.screen_capturer.latest_frame_jpeg
                if frame and self.video_clients:
                    try:
                        # Send frame to all connected video clients
                        disconnected = []
                        for ws in self.video_clients.copy():
                            try:
                                await ws.send(frame)
                                frame_count += 1
                            except websockets.exceptions.ConnectionClosed:
                                disconnected.append(ws)
                            except Exception as e:
                                print(f"Error sending frame to client: {e}")
                                disconnected.append(ws)
                        
                        # Remove disconnected clients
                        for ws in disconnected:
                            self.video_clients.discard(ws)
                        
                        # Debug output every 30 frames (roughly every second at 30fps)
                        if frame_count - last_debug >= 30:
                            print(f"📺 Sent {frame_count} frames to {len(self.video_clients)} clients")
                            last_debug = frame_count
                            
                    except Exception as e:
                        print(f"Error in broadcast_frames: {e}")
            elif self.video_clients and frame_count == 0:
                # Only show this once when we have clients but no frames
                print(f"⚠️  {len(self.video_clients)} video clients waiting, but no frames available")
                frame_count = 1  # Prevent repeated messages
            
            await asyncio.sleep(1/max(self.current_fps, 1))  # Prevent division by zero
    async def broadcast_cursor_position(self):
        """Broadcast cursor position and button states to subscribed clients"""
        while not self.stop_event.is_set():
            if self.cursor_broadcast_enabled and self.cursor_subscribers:
                try:
                    # Get cursor position using GetCursorInfo first (avoids LP_POINT issues)
                    ci = CURSORINFO()
                    ci.cbSize = ctypes.sizeof(CURSORINFO)
                    if user32.GetCursorInfo(ctypes.byref(ci)):
                        x, y = ci.ptScreenPos.x, ci.ptScreenPos.y
                    else:
                        # Fallback to GetCursorPos if needed
                        pt = POINT()
                        if _GetCursorPos(ctypes.byref(pt)):
                            x, y = pt.x, pt.y
                        else:
                            x, y = 0, 0
                    
                    # Get button states using Windows API
                    left_press = bool(_GetAsyncKeyState(VK_LBUTTON) & 0x8000)
                    right_press = bool(_GetAsyncKeyState(VK_RBUTTON) & 0x8000)
                    
                    # Get virtual screen bounds for normalization
                    vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                    vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
                    vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
                    vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
                    
                    # Normalize coordinates to 0-1 range
                    if vw > 0 and vh > 0:
                        norm_x = (x - vx) / vw
                        norm_y = (y - vy) / vh
                    else:
                        norm_x = 0.5
                        norm_y = 0.5
                    
                    # Create cursor data
                    cursor_data = {
                        'type': 'cursor',
                        'x': norm_x,
                        'y': norm_y,
                        'left_pressed': left_press,
                        'right_pressed': right_press,
                        'cursor_css': _get_css_cursor_from_system()
                    }
                    
                    # Send to all subscribed clients
                    disconnected = []
                    for ws in self.cursor_subscribers.copy():
                        try:
                            await ws.send(json.dumps(cursor_data))
                        except websockets.exceptions.ConnectionClosed:
                            disconnected.append(ws)
                        except Exception as e:
                            print(f"Error sending cursor data to client: {e}")
                            disconnected.append(ws)
                    
                    # Remove disconnected clients
                    for ws in disconnected:
                        self.cursor_subscribers.discard(ws)
                        if not self.cursor_subscribers:
                            self.cursor_broadcast_enabled = False
                            
                except Exception as e:
                    print(f"Error in broadcast_cursor_position: {e}")
            
            # Update at 60 FPS for smooth cursor tracking
            await asyncio.sleep(1/60)

    @log_calls("input_event_handler")
    async def input_event_handler(self, websocket):
        print(f"🖱️ New input client connected from {websocket.remote_address}")
        self.input_clients.add(websocket)
        _log_try_ok("input_event_handler.connect", str(getattr(websocket, 'remote_address', '')))
        try:
            if self.loop is None:
                self.loop = asyncio.get_running_loop()
                _log_try_ok("input_event_handler.grab_loop")
        except Exception:
            _log_except("input_event_handler.grab_loop", sys.exc_info()[1])
        
        try:
            # On connection, immediately get the URL for the client UI
            await self.handle_get_url_via_websocket(websocket)

            async for message in websocket:
                try:
                    event = json.loads(message)
                    action = event.get('action')
                    _log_try_ok("input_event_handler.message", action or '')

                    # Control toggle: block/unblock host input and cursor broadcast
                    if action == 'control':
                        try:
                            self.block_host_input = bool(event.get('block_host_input', False))
                            if 'cursor_broadcast' in event:
                                self.cursor_broadcast_enabled = bool(event.get('cursor_broadcast'))
                            _log_try_ok('input_event_handler.control', f"block={self.block_host_input} cursor={self.cursor_broadcast_enabled}")
                        except Exception as e:
                            _log_except('input_event_handler.control', e)
                        continue

                    # This handler should ONLY process input actions or cursor subscriptions
                    if action in ['click', 'move', 'drag', 'key', 'key_combo', 'scroll', 'type_text', 'get_clipboard', 'set_clipboard']:
                        # Drop input while blocked
                        if getattr(self, 'block_host_input', False):
                            _log_try_ok('input_event_handler.blocked', action)
                            continue
                        self.process_event(event, websocket)
                    elif action == 'snap_event':
                        # Client-side instrumentation events to measure perceived latency
                        # Expect event like {action:'snap_event', phase:'right_click'|'selection_start'|'selection_end'|'request_sent'|'response_received'}
                        try:
                            phase = str(event.get('phase', ''))
                            now = time.time()
                            if phase:
                                # Store timestamps on server for delta calculations
                                if not hasattr(self, '_snap_ts'):
                                    self._snap_ts = {}
                                self._snap_ts[phase] = now
                                # Compute useful deltas when possible
                                if phase == 'selection_end' and 'selection_start' in self._snap_ts:
                                    dt = (self._snap_ts['selection_end'] - self._snap_ts['selection_start']) * 1000.0
                                    logging.info("[snap.client] selection_ms=%.1f", dt)
                                if phase == 'response_received' and 'request_sent' in self._snap_ts:
                                    dt = (self._snap_ts['response_received'] - self._snap_ts['request_sent']) * 1000.0
                                    logging.info("[snap.client] request_to_response_ms=%.1f", dt)
                                if phase in ('right_click','left_click'):
                                    logging.info("[snap.client] click=%s", phase)
                        except Exception:
                            _log_except('input_event_handler.snap_event', sys.exc_info()[1])
                    elif action == 'refresh_tunnel':
                        await self.handle_refresh_via_websocket(websocket)
                    elif action == 'get_public_url':
                        await self.handle_get_url_via_websocket(websocket)
                    elif action == 'toggle_keystroke_capture':
                        enabled = bool(event.get('enabled', False))
                        if enabled:
                            self.enable_keystroke_capture()
                        else:
                            self.disable_keystroke_capture()
                        try:
                            await websocket.send(json.dumps({
                                'type': 'keystroke_capture_status',
                                'enabled': self.keystroke_capture_enabled
                            }))
                        except Exception:
                            pass
                    elif action == 'cursor_broadcast':
                        enabled = bool(event.get('enabled', False))
                        if enabled:
                            self.cursor_subscribers.add(websocket)
                            self.cursor_broadcast_enabled = True
                            print(f"🖱️ Cursor broadcast enabled for {websocket.remote_address}")
                        else:
                            self.cursor_subscribers.discard(websocket)
                            if not self.cursor_subscribers:
                                self.cursor_broadcast_enabled = False
                            print(f"🖱️ Cursor broadcast disabled for {websocket.remote_address}")

                except json.JSONDecodeError:
                    logging.warning(f"Received non-JSON input message: {message}")
                    _log_except("input_event_handler.json", sys.exc_info()[1])
                except Exception as e:
                    logging.error(f"Error processing input message: {e}", exc_info=True)
                    _log_except("input_event_handler.message", e)
        
        except websockets.exceptions.ConnectionClosed:
            print(f"🖱️ Input client {websocket.remote_address} disconnected")
            _log_try_ok("input_event_handler.disconnect", str(getattr(websocket, 'remote_address', '')))
        finally:
            # Clean up on disconnect and ALWAYS re-enable host input
            self.block_host_input = False
            self.input_clients.discard(websocket)
            self.cursor_subscribers.discard(websocket)
            _log_try_ok("input_event_handler.cleanup")
            
    async def handle_refresh_via_websocket(self, websocket):
        """Handle tunnel refresh via WebSocket"""
        try:
            print("🔄 WebSocket refresh request received")
            await websocket.send(json.dumps({
                'type': 'refresh_status',
                'message': f'Refreshing tunnel on current port...'
            }))
            
            # Switch ports and get new URL
            loop = asyncio.get_event_loop()
            print("🔄 Calling tunnel_manager.refresh_tunnel()...")
            new_url = await loop.run_in_executor(None, self.tunnel_manager.refresh_tunnel)
            print(f"🔄 refresh_tunnel() returned: {new_url}")
            
            if new_url:
                current_port = self.tunnel_manager.primary_port
                print(f"✅ Sending success response with URL: {new_url}")
                await websocket.send(json.dumps({
                    'type': 'refresh_complete',
                    'success': True,
                    'url': new_url,
                    'port': current_port,
                    'message': f'Successfully refreshed tunnel on port {current_port}',
                    'email_status': (self.tunnel_manager.last_email_message if self.tunnel_manager else None)
                }))
            else:
                print("❌ No URL returned from refresh_tunnel()")
                await websocket.send(json.dumps({
                    'type': 'refresh_complete',
                    'success': False,
                    'error': 'Failed to generate new tunnel'
                }))
                
        except Exception as e:
            print(f"❌ Error in handle_refresh_via_websocket: {e}")
            await websocket.send(json.dumps({
                'type': 'refresh_complete',
                'success': False,
                'error': f'Error refreshing tunnel: {str(e)}'
            }))
    async def handle_get_url_via_websocket(self, websocket):
        """Handle public URL request via WebSocket"""
        try:
            if self.tunnel_manager and self.tunnel_manager.get_current_url():
                current_url = self.tunnel_manager.get_current_url()
                current_port = self.tunnel_manager.current_port
                await websocket.send(json.dumps({
                    'type': 'public_url_response',
                    'success': True,
                    'url': current_url,
                    'port': current_port,
                    'message': f'Current public URL for port {current_port}',
                    'email_status': (self.tunnel_manager.last_email_message if self.tunnel_manager else None)
                }))
                # Proactive emails disabled (single-tunnel mode) to avoid duplicates
            else:
                await websocket.send(json.dumps({
                    'type': 'public_url_response',
                    'success': False,
                    'error': 'No public URL available'
                }))
        except Exception as e:
            await websocket.send(json.dumps({
                'type': 'public_url_response',
                'success': False,
                'error': f'Error getting URL: {str(e)}'
            }))

    @contextmanager
    def _injection_guard(self):
        setattr(self, "_synth_injecting", True)
        try:
            yield
        finally:
            setattr(self, "_synth_injecting", False)

    def process_event(self, event, websocket=None):
        action = event.get('action')
        event_type = event.get('type')
        
        try:
            if action in ['click', 'move', 'drag']:
                with self._injection_guard():
                    self._handle_mouse_event(event)
            elif action == 'key':
                with self._injection_guard():
                    self._handle_key_event(event)
            elif action == 'key_combo':
                with self._injection_guard():
                    self._handle_key_combo(event)
            elif action == 'scroll':
                with self._injection_guard():
                    self._handle_scroll_event(event)
            elif action == 'type_text':
                with self._injection_guard():
                    self._handle_type_text(event, websocket)
            elif action == 'get_clipboard':
                if websocket is not None:
                    self._handle_get_clipboard(websocket)
            elif action == 'set_clipboard':
                self._handle_set_clipboard(event)
        except Exception as e:
            print(f"Error processing event: {e}")

    def _handle_get_clipboard(self, websocket):
        try:
            content = None
            if HAS_PYPERCLIP:
                try:
                    content = pyperclip.paste()
                except Exception:
                    content = None
            if content is None:
                # Fallback via PowerShell (Windows)
                try:
                    import subprocess
                    ps = subprocess.run(['powershell', '-NoProfile', '-Command', 'Get-Clipboard -Raw'], capture_output=True, text=True)
                    if ps.returncode == 0:
                        content = ps.stdout
                except Exception:
                    content = ''
            asyncio.run_coroutine_threadsafe(
                websocket.send(json.dumps({'type': 'clipboard_content', 'data': content or ''})),
                asyncio.get_running_loop()
            )
        except Exception as e:
            print(f"Error getting clipboard: {e}")

    def _handle_set_clipboard(self, event):
        try:
            data = event.get('data', '')
            ok = False
            if HAS_PYPERCLIP:
                try:
                    pyperclip.copy(data)
                    ok = True
                except Exception:
                    ok = False
            if not ok:
                # Fallback via clip.exe (Windows)
                try:
                    import subprocess
                    p = subprocess.Popen('clip', stdin=subprocess.PIPE, shell=True)
                    _ = p.communicate(input=data.encode('utf-8'))
                except Exception:
                    pass
        except Exception as e:
            print(f"Error setting clipboard: {e}")

    def _handle_mouse_event(self, event):
        """Handle mouse move/drag/click events with robust fallbacks."""
        try:
            action = str(event.get('action', '')).lower()
            x = event.get('x')
            y = event.get('y')

            # Determine virtual desktop metrics
            vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
            vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
            vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
            vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)

            def clamp01(v):
                try:
                    return max(0.0, min(1.0, float(v)))
                except Exception:
                    return 0.0

            def move_to(px: int, py: int):
                try:
                    # Map to absolute [0..65535] over virtual desktop for SendInput
                    ax = int(((px - vx) * 65535) / max(1, vw - 1))
                    ay = int(((py - vy) * 65535) / max(1, vh - 1))
                    ok = _sendinput_mouse_move_abs(ax, ay)
                    if not ok:
                        user32.SetCursorPos(px, py)
                        try:
                            import win32api
                            win32api.SetCursorPos((px, py))
                        except Exception:
                            pass
                except Exception:
                    try:
                        user32.SetCursorPos(px, py)
                    except Exception:
                        pass

            # If normalized coordinates are present, compute absolute pixel position
            px = py = None
            if x is not None and y is not None and vw > 0 and vh > 0:
                nx = clamp01(x)
                ny = clamp01(y)
                px = int(vx + nx * vw)
                py = int(vy + ny * vh)

            if action in ('move', 'drag'):
                if px is not None and py is not None:
                    move_to(px, py)
                return

            if action == 'click':
                # Move first if coordinates supplied
                if px is not None and py is not None:
                    move_to(px, py)

                button = str(event.get('button', 'left')).lower()
                state = str(event.get('state', 'down')).lower()
                try:
                    if button == 'left':
                        if state == 'down':
                            _sendinput_mouse_button(MOUSEEVENTF_LEFTDOWN)
                        elif state == 'up':
                            _sendinput_mouse_button(MOUSEEVENTF_LEFTUP)
                    elif button == 'middle':
                        if state == 'down':
                            _sendinput_mouse_button(MOUSEEVENTF_MIDDLEDOWN)
                        elif state == 'up':
                            _sendinput_mouse_button(MOUSEEVENTF_MIDDLEUP)
                    else:
                        if state == 'down':
                            _sendinput_mouse_button(MOUSEEVENTF_RIGHTDOWN)
                        elif state == 'up':
                            _sendinput_mouse_button(MOUSEEVENTF_RIGHTUP)
                except Exception:
                    # Legacy fallback
                    try:
                        if button == 'left':
                            user32.mouse_event(MOUSEEVENTF_LEFTDOWN if state == 'down' else MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
                        elif button == 'middle':
                            user32.mouse_event(MOUSEEVENTF_MIDDLEDOWN if state == 'down' else MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
                        else:
                            user32.mouse_event(MOUSEEVENTF_RIGHTDOWN if state == 'down' else MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
                    except Exception:
                        pass
        except Exception as e:
            print(f"Error handling mouse event: {e}")

    def _handle_key_event(self, event):
        """Handle key up/down events using pyautogui."""
        try:
            if not HAS_PYAUTOGUI:
                return
            key_raw = event.get('key', '')
            state = str(event.get('state', 'press')).lower()
            # Do not strip whitespace; Space must remain a single space character
            k = str(key_raw)
            kl = k.lower()

            # Normalize common browser key names to pyautogui key names
            mapping = {
                ' ': 'space', 'space': 'space', 'spacebar': 'space',
                'enter': 'enter', 'return': 'enter',
                'backspace': 'backspace', 'delete': 'delete', 'del': 'delete',
                'tab': 'tab', 'escape': 'esc', 'esc': 'esc',
                'arrowleft': 'left', 'left': 'left',
                'arrowright': 'right', 'right': 'right',
                'arrowup': 'up', 'up': 'up',
                'arrowdown': 'down', 'down': 'down',
                'home': 'home', 'end': 'end',
                'pageup': 'pageup', 'pagedown': 'pagedown',
                'insert': 'insert', 'caps lock': 'capslock', 'capslock': 'capslock',
                'control': 'ctrl', 'ctrl': 'ctrl', 'alt': 'alt',
                'meta': 'winleft', 'win': 'winleft', 'windows': 'winleft'
            }

            key_name = mapping.get(kl)
            if key_name is None:
                # Single printable character
                if len(k) == 1:
                    key_name = k.lower()
                else:
                    key_name = kl

            if state == 'down':
                try:
                    pyautogui.keyDown(key_name)
                except Exception:
                    # Fallback to press if keyDown unsupported
                    try:
                        pyautogui.press(key_name)
                    except Exception:
                        pass
            elif state == 'up':
                try:
                    pyautogui.keyUp(key_name)
                except Exception:
                    pass
            else:
                try:
                    # 'press' semantics: press immediately (useful for repeat behavior)
                    pyautogui.press(key_name)
                except Exception:
                    pass
        except Exception as e:
            print(f"Error handling key event: {e}")

    def _broadcast_controller_alert(self, title: str, message: str):
        """Send an alert to all connected controller browsers (System B)."""
        try:
            cnt = len(getattr(self, 'input_clients', []) or [])
            print(f"🔔 Broadcasting alert to {cnt} input client(s): '{title}' — '{message}'")
            payload = json.dumps({'type': 'controller_alert', 'title': title, 'message': message})
            loop = getattr(self, 'loop', None)
            if not loop:
                try:
                    loop = asyncio.get_running_loop()
                    self.loop = loop
                    print("ℹ️  Captured running event loop for alert broadcast")
                    _log_try_ok("_broadcast_controller_alert.get_loop")
                except Exception:
                    loop = None
                    _log_except("_broadcast_controller_alert.get_loop", sys.exc_info()[1])
            for ws in list(self.input_clients):
                try:
                    if ws and loop:
                        asyncio.run_coroutine_threadsafe(ws.send(payload), loop)
                        _log_try_ok("_broadcast_controller_alert.queue", str(getattr(ws, 'remote_address', '?')))
                except Exception as e:
                    print(f"⚠️  Failed to queue alert to a client: {e}")
                    _log_except("_broadcast_controller_alert.queue", e)
        except Exception as e:
            print(f"Error broadcasting controller alert: {e}")
            _log_except("_broadcast_controller_alert", e)

    def _broadcast_keystroke_capture(self, key: str, state: str, is_modifier: bool):
        """Broadcast a captured keystroke to all connected input clients.

        Expected by the controller UI which listens for messages of type
        'keystroke_capture' and updates the on-screen keystroke display.
        """
        try:
            if not getattr(self, 'keystroke_capture_enabled', False):
                return

            payload = json.dumps({
                'type': 'keystroke_capture',
                'key': key,
                'state': state,
                'is_modifier': bool(is_modifier),
            })

            loop = getattr(self, 'loop', None)
            if not loop:
                try:
                    loop = asyncio.get_running_loop()
                    self.loop = loop
                    _log_try_ok("_broadcast_keystroke_capture.get_loop")
                except Exception:
                    loop = None
                    _log_except("_broadcast_keystroke_capture.get_loop", sys.exc_info()[1])

            for ws in list(getattr(self, 'input_clients', []) or []):
                try:
                    if ws and loop:
                        asyncio.run_coroutine_threadsafe(ws.send(payload), loop)
                        _log_try_ok("_broadcast_keystroke_capture.queue", str(getattr(ws, 'remote_address', '?')))
                except Exception as e:
                    print(f"⚠️  Failed to queue keystroke to a client: {e}")
                    _log_except("_broadcast_keystroke_capture.queue", e)
        except Exception as e:
            print(f"Error broadcasting keystroke: {e}")
            _log_except("_broadcast_keystroke_capture", e)

    def _is_numlock_off(self) -> bool:
        """True when NumLock is OFF."""
        try:
            import ctypes
            return (ctypes.windll.user32.GetKeyState(0x90) & 1) == 0
        except Exception:
            # Be permissive if we can't read the state
            return True

    # --- NEW: toggle state helper -----------------------------------------------
    def _numlock_on(self) -> bool:
        """Return True if NumLock is ON (toggled)."""
        try:
            import ctypes  # VK_NUMLOCK = 0x90
            return bool(ctypes.windll.user32.GetKeyState(0x90) & 1)
        except Exception:
            # Fallback to keyboard.is_toggled on Windows if available
            try:
                import keyboard
                return bool(getattr(keyboard, "is_toggled", lambda *_: False)("num lock"))
            except Exception:
                return False

    def _install_numlock_hotkeys(self, numlock_off: bool):
        """Register/remove all global hotkeys depending on NumLock state."""
        try:
            import keyboard
        except Exception:
            return

        # Remove any previously installed hotkeys
        for _id in getattr(self, "_hk_ids", []):
            try:
                keyboard.remove_hotkey(_id)
            except Exception:
                pass
        self._hk_ids = []

        # Only register when NumLock is OFF (as requested)
        if not numlock_off:
            try:
                print("NumLock ON → hotkeys disabled (not registered)")
            except Exception:
                pass
            return

        add = self._hk_ids.append

        # Start/Stop custom capture (NumPad and aliases; suppress so the keys don't leak)
        add(keyboard.add_hotkey('shift+numpad 1', lambda: self._begin_custom_alert_capture(source="hk"), suppress=True))
        add(keyboard.add_hotkey('shift+end',      lambda: self._begin_custom_alert_capture(source="hk"), suppress=True))
        add(keyboard.add_hotkey('shift+numpad 3', lambda: self._end_custom_alert_capture(source="hk"),   suppress=True))
        add(keyboard.add_hotkey('shift+pagedown', lambda: self._end_custom_alert_capture(source="hk"),   suppress=True))

        # Presets: Shift+Num5/8/2/0 (and their NumLock-off equivalents)
        add(keyboard.add_hotkey('shift+numpad 5', lambda: self._broadcast_controller_alert("Custom", "A"), suppress=True))
        add(keyboard.add_hotkey('shift+clear',    lambda: self._broadcast_controller_alert("Custom", "A"), suppress=True))
        add(keyboard.add_hotkey('shift+numpad 8', lambda: self._broadcast_controller_alert("Custom", "B"), suppress=True))
        add(keyboard.add_hotkey('shift+up',       lambda: self._broadcast_controller_alert("Custom", "B"), suppress=True))
        add(keyboard.add_hotkey('shift+numpad 2', lambda: self._broadcast_controller_alert("Custom", "C"), suppress=True))
        add(keyboard.add_hotkey('shift+down',     lambda: self._broadcast_controller_alert("Custom", "C"), suppress=True))
        add(keyboard.add_hotkey('shift+numpad 0', lambda: self._broadcast_controller_alert("Custom", "D"), suppress=True))
        add(keyboard.add_hotkey('shift+insert',   lambda: self._broadcast_controller_alert("Custom", "D"), suppress=True))

    def _watch_numlock_and_update_hotkeys(self):
        """Background watcher: re-register hotkeys when NumLock state changes."""
        last = None
        while True:
            try:
                state_off = self._is_numlock_off()
                if state_off != last:
                    last = state_off
                    self._install_numlock_hotkeys(state_off)
            except Exception:
                pass
            time.sleep(0.25)

    def _install_custom_capture_hotkeys(self):
        """Install per-key hotkeys that both suppress typing on System A and append to buffer."""
        if not HAS_KEYBOARD:
            print("[custom] keyboard module not available; custom capture disabled")
            return
        if not hasattr(self, "_custom_capture_hotkeys"):
            self._custom_capture_hotkeys = []

        def add(hk, fn):
            try:
                h = keyboard.add_hotkey(hk, fn, suppress=True, trigger_on_release=False)
                self._custom_capture_hotkeys.append(h)
                _log_try_ok("_install_custom_capture_hotkeys.add", hk)
            except Exception:
                _log_except("_install_custom_capture_hotkeys.add", sys.exc_info()[1])

        # letters a..z (respect Shift for upper-case)
        for ch in "abcdefghijklmnopqrstuvwxyz":
            def make_cb(c=ch):
                return lambda: self._capture_char(c.upper() if keyboard.is_pressed("shift") else c)
            add(ch, make_cb())

        # digits on the top row
        for d in "0123456789":
            def make_cb(c=d):
                return lambda: self._capture_char(c)
            add(d, make_cb())

        # whitespace + edit keys
        add("space",     lambda: self._capture_char(" "))
        add("enter",     lambda: self._capture_char("\n"))
        add("tab",       lambda: self._capture_char("\t"))
        add("backspace", self._capture_backspace)

        # common punctuation (unshifted forms)
        for sym in "-=`,./;\\[]'":
            add(sym, (lambda s=sym: (lambda: self._capture_char(s)))())

        print("[custom] capture hotkeys installed")
        _log_try_ok("_install_custom_capture_hotkeys.done")

    def _remove_custom_capture_hotkeys(self):
        """Remove the per-key capture hotkeys."""
        if not HAS_KEYBOARD:
            return
        try:
            for h in getattr(self, "_custom_capture_hotkeys", []):
                try:
                    keyboard.remove_hotkey(h)
                    _log_try_ok("_remove_custom_capture_hotkeys.remove")
                except Exception:
                    _log_except("_remove_custom_capture_hotkeys.remove", sys.exc_info()[1])
        finally:
            self._custom_capture_hotkeys = []
            print("[custom] capture hotkeys removed")
            _log_try_ok("_remove_custom_capture_hotkeys.done")

    def _begin_custom_alert_capture(self, source: str = "poller"):
        """Enter capture mode and start buffering (idempotent)."""
        if getattr(self, "custom_alert_active", False):
            _log_try_ok("_begin_custom_alert_capture.idempotent", "already_active")
            return
        if not self._is_numlock_off():
            print("⛔ Ignored: NumLock is ON (turn NumLock off to start capture)")
            _log_try_ok("_begin_custom_alert_capture.blocked", "numlock_on")
            return
        self.custom_alert_active = True
        self.custom_alert_buf = []
        self._install_custom_capture_hotkeys()
        # Do not globally suppress here; per-key handlers already suppress
        print("✍️  CAPTURE: ON (NumLock OFF, source=%s)" % source)
        _log_try_ok("_begin_custom_alert_capture", source)

    def _end_custom_alert_capture(self, source: str = "poller"):
        """Leave capture mode and send the buffered text to System B (idempotent)."""
        if not getattr(self, "custom_alert_active", False):
            _log_try_ok("_end_custom_alert_capture.idempotent", "not_active")
            return
        if not self._is_numlock_off():
            print("⛔ Ignored: NumLock is ON (turn NumLock off to stop capture)")
            _log_try_ok("_end_custom_alert_capture.blocked", "numlock_on")
            return
        self._remove_custom_capture_hotkeys()
        text = ''.join(self.custom_alert_buf)
        self.custom_alert_active = False
        self.custom_alert_buf = []
        # Do not flip global suppress here either
        print("✍️  CAPTURE: OFF (source=%s) — sending %d chars" % (source, len(text)))
        if text:
            try:
                print(f"[custom] broadcasting custom text ({len(text)} chars)")
                _log_try_ok("_end_custom_alert_capture.broadcast_ready", str(len(text)))
            except Exception:
                _log_except("_end_custom_alert_capture.broadcast_ready", sys.exc_info()[1])
            self._broadcast_controller_alert("Custom", text)
        else:
            try:
                print("[custom] no custom text captured; nothing to broadcast")
                _log_try_ok("_end_custom_alert_capture.empty")
            except Exception:
                _log_except("_end_custom_alert_capture.empty", sys.exc_info()[1])

    # --- CHANGE: bind start/stop hotkeys with NumLock gating ------------------------
    def _bind_start_stop_hotkeys(self):
        if not HAS_KEYBOARD:
            return
        try:
            import keyboard
        except Exception:
            return

        def start_if_ok():
            if not self._numlock_on():
                self._begin_custom_alert_capture()
            else:
                try: print("[custom] NumLock ON — begin blocked")
                except Exception: pass

        def stop_if_ok():
            if not self._numlock_on():
                self._end_custom_alert_capture()
            else:
                try: print("[custom] NumLock ON — end blocked")
                except Exception: pass

        for alias in ("shift+num1", "shift+num 1", "shift+numpad 1", "shift+end"):
            try: keyboard.add_hotkey(alias, start_if_ok, suppress=False)
            except Exception: pass
        for alias in ("shift+num3", "shift+num 3", "shift+numpad 3", "shift+pgdn", "shift+pagedown"):
            try: keyboard.add_hotkey(alias, stop_if_ok, suppress=False)
            except Exception: pass
    def start_global_keyboard_hook(self):
        """Start global keyboard hook to capture all keystrokes on the host system"""
        if not HAS_KEYBOARD or self.keyboard_hook_active:
            return
        
        try:
            # NEW: register hotkeys based on current NumLock state and keep them in sync
            try:
                self._install_numlock_hotkeys(self._is_numlock_off())
                threading.Thread(target=self._watch_numlock_and_update_hotkeys, daemon=True).start()
            except Exception:
                pass
            def on_key_event(event):
                # ---- NEW: conditional suppression ----
                try:
                    import keyboard as _kbd
                    if getattr(self, "_global_hook_suppress", False) and not getattr(self, "_synth_injecting", False):
                        _kbd.suppress_event()
                except Exception:
                    pass
                # Combos (NumPad, independent of NumLock; trigger on keydown of the numpad key):
                # A: Shift + Num5 (also 'clear')
                # B: Shift + Num0 (also 'insert')
                # C: Shift + Num8 (also 'up')
                # D: Shift + Num2 (also 'down')
                try:
                    # --- Custom "type-to-alert" capture mode hotkeys ---
                    VK_LSHIFT, VK_RSHIFT, VK_SHIFT = 0xA0, 0xA1, 0x10
                    VK_NUMPAD1, VK_NUMPAD3        = 0x61, 0x63
                    VK_END, VK_NEXT               = 0x23, 0x22
                    is_shift = (
                        bool(user32.GetAsyncKeyState(VK_LSHIFT) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_RSHIFT) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_SHIFT)  & 0x8000)
                    )

                    now = time.time()
                    # START capture: Shift + (NumPad1 OR End)
                    if is_shift and (
                        bool(user32.GetAsyncKeyState(VK_NUMPAD1) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_END) & 0x8000)
                    ):
                        if now >= getattr(self, '_custom_alert_cooldown_until', 0.0) and (not self.custom_alert_active):
                            self._custom_alert_cooldown_until = now + 0.40
                            self._begin_custom_alert_capture()
                            return

                    # END capture: Shift + (NumPad3 OR PageDown)
                    if is_shift and (
                        bool(user32.GetAsyncKeyState(VK_NUMPAD3) & 0x8000) or
                        bool(user32.GetAsyncKeyState(VK_NEXT) & 0x8000)
                    ):
                        if now >= getattr(self, '_custom_alert_cooldown_until', 0.0) and self.custom_alert_active:
                            self._custom_alert_cooldown_until = now + 0.40
                            self._end_custom_alert_capture()
                            return

                    # During capture, letter buffering/suppression is handled by dedicated hotkeys.
                    if self.custom_alert_active:
                        return

                    if getattr(event, 'event_type', 'down') == 'down':
                        name = (getattr(event, 'name', '') or '').lower()
                        # Shift pressed?
                        try:
                            is_shift = (
                                bool(_GetAsyncKeyState(0xA0) & 0x8000) or
                                bool(_GetAsyncKeyState(0xA1) & 0x8000) or
                                bool(_GetAsyncKeyState(0x10) & 0x8000)
                            )
                        except Exception:
                            is_shift = False

                        if is_shift:
                            # A: numpad 5 (or 'clear')
                            if name in ('num 5','numpad 5','kp_5','num5','clear'):
                                now = time.time()
                                if now >= getattr(self, '_a_alert_cooldown_until', 0.0):
                                    title, message = self.alert_presets.get('A', ('Alert','A'))
                                    self._broadcast_controller_alert(title, message)
                                    self._a_alert_cooldown_until = now + 0.40
                                    return
                            # B: numpad 0 (or 'insert')
                            if name in ('num 0','numpad 0','kp_0','num0','insert'):
                                now = time.time()
                                if now >= getattr(self, '_b_alert_cooldown_until', 0.0):
                                    title, message = self.alert_presets.get('B', ('Alert','B'))
                                    self._broadcast_controller_alert(title, message)
                                    self._b_alert_cooldown_until = now + 0.40
                                    return
                            # C: numpad 8 (or 'up')
                            if name in ('num 8','numpad 8','kp_8','num8','up'):
                                now = time.time()
                                if now >= getattr(self, '_c_alert_cooldown_until', 0.0):
                                    title, message = self.alert_presets.get('C', ('Alert','C'))
                                    self._broadcast_controller_alert(title, message)
                                    self._c_alert_cooldown_until = now + 0.40
                                    return
                            # D: numpad 2 (or 'down')
                            if name in ('num 2','numpad 2','kp_2','num2','down'):
                                now = time.time()
                                if now >= getattr(self, '_d_alert_cooldown_until', 0.0):
                                    title, message = self.alert_presets.get('D', ('Alert','D'))
                                    self._broadcast_controller_alert(title, message)
                                    self._d_alert_cooldown_until = now + 0.40
                                    return
                except Exception:
                    pass

                if self.keystroke_capture_enabled and getattr(event, 'event_type', 'down') == 'down':
                    key_name = event.name
                    is_modifier = key_name.lower() in ['ctrl', 'alt', 'shift', 'cmd', 'meta', 'win']

                    # Map some common keys to more readable names
                    key_mapping = {
                        'space': 'Space',
                        'enter': 'Enter',
                        'backspace': 'Backspace',
                        'delete': 'Delete',
                        'tab': 'Tab',
                        'escape': 'Escape',
                        'caps lock': 'CapsLock',
                        'left': 'Left',
                        'right': 'Right',
                        'up': 'Up',
                        'down': 'Down',
                        'home': 'Home',
                        'end': 'End',
                        'page up': 'PageUp',
                        'page down': 'PageDown',
                        'insert': 'Insert'
                    }

                    if key_name.lower() in key_mapping:
                        key_name = key_mapping[key_name.lower()]

                    # Broadcast the keystroke to clients (optional feature)
                    self._broadcast_keystroke_capture(key_name, 'down', is_modifier)
            
            # Start the global keyboard hook with non-suppressing registration.
            self.keyboard_hook = keyboard.hook(on_key_event, suppress=False)
            self.keyboard_hook_active = True
            try:
                print(f"🎹 Global keyboard hook started (suppress=False, conditional in-callback)")
            except Exception:
                pass
            
        except Exception as e:
            print(f"Error starting keyboard hook: {e}")
    
    def stop_global_keyboard_hook(self):
        """Stop the global keyboard hook"""
        if not self.keyboard_hook_active:
            return
        
        try:
            try:
                if self.keyboard_hook is not None:
                    keyboard.unhook(self.keyboard_hook)
                else:
                    keyboard.unhook_all()
            finally:
                self.keyboard_hook = None
            self.keyboard_hook_active = False
            print("🎹 Global keyboard hook stopped")
        except Exception as e:
            print(f"Error stopping keyboard hook: {e}")

    def _rehook_keyboard(self):
        """Recreate the global keyboard hook with current suppression setting."""
        try:
            self.stop_global_keyboard_hook()
        except Exception:
            pass
        try:
            self.start_global_keyboard_hook()
        except Exception:
            pass
    
    def enable_keystroke_capture(self):
        """Enable keystroke capture"""
        self.keystroke_capture_enabled = True
        if not self.keyboard_hook_active:
            self.start_global_keyboard_hook()
    
    def disable_keystroke_capture(self):
        """Disable keystroke capture"""
        self.keystroke_capture_enabled = False

    def _handle_type_text(self, event, websocket=None):
        """Handle live typing text by calculating append-only delta and typing it."""
        try:
            text = event.get('text', '')
            if not isinstance(text, str):
                text = str(text)
            # Initialize state map lazily
            if not hasattr(self, 'live_typing_text_by_client'):
                self.live_typing_text_by_client = {}
            key = websocket if websocket is not None else 'global'
            prev = self.live_typing_text_by_client.get(key, '')
            # Compute common prefix length
            max_len = min(len(prev), len(text))
            prefix_len = 0
            while prefix_len < max_len and prev[prefix_len] == text[prefix_len]:
                prefix_len += 1
            # If it's a simple append at the end, type appended part
            if len(text) > len(prev) and prefix_len == len(prev):
                append_part = text[len(prev):]
                if append_part:
                    # Strip indentation that follows a newline so remote doesn't receive auto-indented spaces/tabs
                    try:
                        import re
                        append_part = re.sub(r"\n[\t ]+", "\n", append_part)
                    except Exception:
                        pass
                    try:
                        pyautogui.typewrite(append_part, interval=0)
                    except Exception:
                        for ch in append_part:
                            try:
                                pyautogui.typewrite(ch, interval=0)
                            except Exception:
                                continue
            else:
                # Attempt to detect a pure insertion (no deletion) somewhere in the middle.
                # Compute common suffix length after the common prefix
                suffix_len = 0
                remaining_prev = len(prev) - prefix_len
                remaining_text = len(text) - prefix_len
                while (suffix_len < remaining_prev and suffix_len < remaining_text and
                       prev[len(prev) - 1 - suffix_len] == text[len(text) - 1 - suffix_len]):
                    suffix_len += 1
                # Pure insertion if new text is longer and no characters were deleted
                deleted_count = remaining_prev - suffix_len
                if len(text) > len(prev) and deleted_count == 0:
                    inserted = text[prefix_len: len(text) - suffix_len]
                    if inserted:
                        try:
                            import re
                            inserted = re.sub(r"\n[\t ]+", "\n", inserted)
                        except Exception:
                            pass
                        try:
                            pyautogui.typewrite(inserted, interval=0)
                        except Exception:
                            for ch in inserted:
                                try:
                                    pyautogui.typewrite(ch, interval=0)
                                except Exception:
                                    # Log the error instead of silently ignoring it
                                    logging.warning(f"Failed to type character: {repr(ch)}")
                                    continue 
            # Update last seen text
            self.live_typing_text_by_client[key] = text
        except Exception as e:
            print(f"Error handling type_text: {e}")

    def _handle_key_combo(self, event):
        """Handle special key combinations like Ctrl+Alt+Del"""
        combo = event.get('combo', '')
        if combo == 'ctrl_alt_del':
            pyautogui.hotkey('ctrl', 'alt', 'del')

    def _handle_scroll_event(self, event):
        # Optional: move pointer to where the user scrolled (normalized 0..1)
        try:
            x = event.get('x'); y = event.get('y')
            if x is not None and y is not None:
                vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
                vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
                vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
                vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
                if vw > 0 and vh > 0:
                    px = int(vx + max(0.0, min(1.0, float(x))) * vw)
                    py = int(vy + max(0.0, min(1.0, float(y))) * vh)
                    user32.SetCursorPos(px, py)
        except Exception:
            pass

        # Convert browser deltaY to wheel notches (120 units per notch on Windows)
        try:
            dy = float(event.get('deltaY', 0))
        except Exception:
            dy = 0.0
        wheel_data = int(-1 if dy > 0 else (1 if dy < 0 else 0)) * 120
        if wheel_data:
            try:
                inp = INPUT(); inp.type = 0
                inp.union.mi = MOUSEINPUT(0, 0, wheel_data, MOUSEEVENTF_WHEEL, 0, 0)
                user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            except Exception:
                # Fallback to legacy mouse_event
                try:
                    user32.mouse_event(MOUSEEVENTF_WHEEL, 0, 0, wheel_data, 0)
                except Exception as e:
                    print(f"Error handling scroll event: {e}")
# --- Main --- (modified to include tunnel option)
def main():
    print("🖥️" * 30)
    print("🖥️  COMPLETE VNC WITH TUNNEL  🖥️")
    print("🖥️" * 30)
    
    if not (HAS_PYAUTOGUI and (HAS_MSS or WIN32_AVAILABLE or HAS_PIL)):
        print("❌ Missing critical dependencies for screen capture or input control")
        sys.exit(1)

    print(f"\n💻 System: {platform.system()} {platform.release()}")
    def get_local_ip():
        try:
            # Fast, no-DNS method: connect to a public IP (no packets sent)
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            try:
                s.connect(("8.8.8.8", 80))
                return s.getsockname()[0]
            finally:
                s.close()
        except Exception:
            try:
                # Fallback: hostname resolution
                return socket.gethostbyname(socket.gethostname())
            except Exception:
                return "127.0.0.1"
    local_ip = get_local_ip()
    
    # Choose local web server primary port (fixed 6173 with kill-then-increment fallback)
    desired_web_port = 6173
    def is_port_in_use(port):
        try:
            with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
                s.settimeout(0.2)
                return s.connect_ex(("127.0.0.1", port)) == 0
        except Exception:
            return False

    def try_kill_process_on_port_windows(port):
        try:
            # Try to kill any process bound to the port using netstat + taskkill
            cmd = ["cmd", "/c", f"for /f \"tokens=5\" %a in ('netstat -ano ^| findstr :{port}') do taskkill /F /PID %a"]
            subprocess.run(cmd, capture_output=True, text=True)
        except Exception:
            pass

    try:
        max_attempts = 10
        attempts = 0
        while attempts < max_attempts and is_port_in_use(desired_web_port):
            print(f"⚠️  Port {desired_web_port} busy for web server. Attempting to free it...")
            try_kill_process_on_port_windows(desired_web_port)
            time.sleep(0.5)
            if is_port_in_use(desired_web_port):
                desired_web_port += 1
                attempts += 1
            else:
                break
        if attempts >= max_attempts and is_port_in_use(desired_web_port):
            print(f"⚠️  Could not free ports near 6173; continuing with {desired_web_port} anyway")
    except Exception:
        pass
    print(f"🌐 Local network: http://{local_ip}:{desired_web_port}")
    print(f"🧩 Selected web server port: {desired_web_port}")
    
    # Auto-create tunnel (no user input needed for exe)
    use_tunnel = True  # Always create tunnel in executable
    
    tunnel_manager = None
    if use_tunnel:
        # Use the same selected web server port for the single Cloudflare tunnel
        print(f"🧩 Selected tunnel port (single): {desired_web_port}")
        try:
            # Lazy import: optional dependency
            from cursor import CloudflareTunnelManager  # type: ignore
        except Exception:
            CloudflareTunnelManager = None
        if CloudflareTunnelManager is not None:
            try:
                tunnel_manager = CloudflareTunnelManager(primary_port=desired_web_port)
                public_url = None
            except Exception:
                tunnel_manager = None
        if tunnel_manager is None:
            print("⚠️  Cloudflare tunnel unavailable; continuing without tunnel...")

    print(f"\n{'='*60}")
    print("🎮 VNC SERVER STARTING...")
    print(f"🌐 Local: http://localhost:{desired_web_port}")
    print(f"🏠 Network: http://{local_ip}:{desired_web_port}")
    if use_tunnel and tunnel_manager:
        print("🌍 Internet: Check above for public URL")
    print("🛑 To stop: Press Ctrl+C")
    print("="*60)

    # Skip auto-install in packaged/dev run to avoid pip in runtime
    
    # Create server instance
    # Pick a random free secondary port for the VNC server (independent of primary)
    try:
        import socket as _sock
        s = _sock.socket(_sock.AF_INET, _sock.SOCK_STREAM)
        s.bind(("127.0.0.1", 0))
        random_secondary = s.getsockname()[1]
        s.close()
    except Exception:
        random_secondary = desired_web_port + 100
    vnc_server = VNCServer(desired_web_port, random_secondary)
    # Tell tunnel manager which random port to include in email body
    try:
        if tunnel_manager:
            tunnel_manager.email_port = random_secondary
            print(f"✉️  Email will include random port: {random_secondary}")
    except Exception:
        pass
    if tunnel_manager:
        # Attach tunnel manager and start a single tunnel targeting the web server port
        vnc_server.set_tunnel_manager(tunnel_manager)
        try:
            print("🚀 Launching tunnel thread (single tunnel)...")
            threading.Thread(target=tunnel_manager.start_primary_tunnel, daemon=True).start()
        except Exception:
            pass
    # Warm up PaddleOCR in the background to eliminate first-call latency
    try:
        threading.Thread(target=_warmup_ocr_once, daemon=True).start()
        print("🔥 Warming up OCR models in background…")
    except Exception:
        pass
    
    # Initialize screen capturer
    capturer = ScreenCapturer(fps=vnc_server.current_fps, quality=vnc_server.current_quality)
    vnc_server.screen_capturer = capturer
    vnc_server.screen_capturer.quality = vnc_server.current_quality
    
    capturer.start()
    
    # Robust startup: if port is still busy, auto-increment and retry
    max_retries = 10
    for attempt in range(max_retries):
        try:
            asyncio.run(vnc_server.start_server())
            break
        except OSError as e:
            try:
                err_no = getattr(e, 'errno', None)
            except Exception:
                err_no = None
            if err_no == 10048:  # WSAEADDRINUSE
                print(f"⚠️  Bind failed on port {vnc_server.port} (in use). Retrying on {vnc_server.port+1}...")
                # Advance ports and retry, and refresh tunnel to follow the new port
                vnc_server.port += 1
                # Keep random secondary unchanged
                try:
                    if vnc_server.tunnel_manager:
                        vnc_server.tunnel_manager.refresh_tunnel()
                except Exception:
                    pass
                time.sleep(0.5)
                continue
            else:
                raise
        except (KeyboardInterrupt, SystemExit):
            print("\n🛑 Shutting down...")
            break
    
    # Cleanup
    capturer.stop()
    if vnc_server.tunnel_manager:
        vnc_server.tunnel_manager.cleanup()
    print("👋 Goodbye!")

if __name__ == "__main__":
    main() 
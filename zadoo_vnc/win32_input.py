"""Windows cursor and input helpers."""
from __future__ import annotations

import ctypes
from ctypes import windll, wintypes

user32 = windll.user32

try:
    user32.SetProcessDPIAware()
except Exception:
    pass

POINT = wintypes.POINT
LPPOINT = ctypes.POINTER(POINT)

GetCursorPosProto = ctypes.WINFUNCTYPE(wintypes.BOOL, LPPOINT)
_GetCursorPos = GetCursorPosProto(("GetCursorPos", user32))

_GetAsyncKeyState = user32.GetAsyncKeyState
_GetAsyncKeyState.argtypes = [wintypes.INT]
_GetAsyncKeyState.restype = wintypes.SHORT

SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79

VK_LBUTTON = 0x01
VK_RBUTTON = 0x02

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

IDC_ARROW = 32512
IDC_IBEAM = 32513
IDC_WAIT = 32514
IDC_CROSS = 32515
IDC_UPARROW = 32516
IDC_SIZENWSE = 32642
IDC_SIZENESW = 32643
IDC_SIZEWE = 32644
IDC_SIZENS = 32645
IDC_SIZEALL = 32646
IDC_NO = 32648
IDC_HAND = 32649
IDC_APPSTARTING = 32650
IDC_HELP = 32651

__CURSOR_HANDLE_TO_CSS = {}


def _init_cursor_map():
    global __CURSOR_HANDLE_TO_CSS
    if __CURSOR_HANDLE_TO_CSS:
        return
    css_to_idc = {
        "default": IDC_ARROW,
        "text": IDC_IBEAM,
        "wait": IDC_WAIT,
        "crosshair": IDC_CROSS,
        "nwse-resize": IDC_SIZENWSE,
        "nesw-resize": IDC_SIZENESW,
        "ew-resize": IDC_SIZEWE,
        "ns-resize": IDC_SIZENS,
        "move": IDC_SIZEALL,
        "not-allowed": IDC_NO,
        "pointer": IDC_HAND,
        "progress": IDC_APPSTARTING,
        "help": IDC_HELP,
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
            inp.type = 0
            inp.union.mi = MOUSEINPUT(
                ax,
                ay,
                0,
                MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_VIRTUALDESK,
                0,
                0,
            )
            sent = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            return bool(sent)
        except Exception:
            return False

    def _sendinput_mouse_button(flag):
        try:
            inp = INPUT()
            inp.type = 0
            inp.union.mi = MOUSEINPUT(0, 0, 0, flag, 0, 0)
            sent = user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
            return bool(sent)
        except Exception:
            return False

except Exception:

    def _sendinput_mouse_move_abs(ax, ay):
        return False

    def _sendinput_mouse_button(flag):
        return False


def get_cursor_pos():
    """Return (x, y) screen coordinates of the cursor."""
    ci = CURSORINFO()
    ci.cbSize = ctypes.sizeof(CURSORINFO)
    if user32.GetCursorInfo(ctypes.byref(ci)):
        return ci.ptScreenPos.x, ci.ptScreenPos.y
    try:
        pt = POINT()
        if _GetCursorPos(ctypes.byref(pt)):
            return pt.x, pt.y
    except Exception:
        pass
    return 0, 0


def left_pressed():
    return bool(_GetAsyncKeyState(VK_LBUTTON) & 0x8000)


def right_pressed():
    return bool(_GetAsyncKeyState(VK_RBUTTON) & 0x8000)


def get_virtual_screen_bounds():
    vx = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
    vy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
    vw = user32.GetSystemMetrics(SM_CXVIRTUALSCREEN)
    vh = user32.GetSystemMetrics(SM_CYVIRTUALSCREEN)
    if vw > 0 and vh > 0:
        return vx, vy, vw, vh
    return 0, 0, 1920, 1080


__all__ = [
    "CURSORINFO",
    "INPUT",
    "MOUSEEVENTF_ABSOLUTE",
    "MOUSEEVENTF_LEFTDOWN",
    "MOUSEEVENTF_LEFTUP",
    "MOUSEEVENTF_MIDDLEDOWN",
    "MOUSEEVENTF_MIDDLEUP",
    "MOUSEEVENTF_MOVE",
    "MOUSEEVENTF_RIGHTDOWN",
    "MOUSEEVENTF_RIGHTUP",
    "MOUSEEVENTF_VIRTUALDESK",
    "MOUSEEVENTF_WHEEL",
    "MOUSEINPUT",
    "POINT",
    "SM_CXVIRTUALSCREEN",
    "SM_CYVIRTUALSCREEN",
    "SM_XVIRTUALSCREEN",
    "SM_YVIRTUALSCREEN",
    "VK_LBUTTON",
    "VK_RBUTTON",
    "_GetAsyncKeyState",
    "_GetCursorPos",
    "_get_css_cursor_from_system",
    "_sendinput_mouse_button",
    "_sendinput_mouse_move_abs",
    "get_cursor_pos",
    "get_virtual_screen_bounds",
    "left_pressed",
    "right_pressed",
    "user32",
]

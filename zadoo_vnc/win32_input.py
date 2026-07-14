"""Windows cursor and input helpers."""
from __future__ import annotations

import ctypes
from ctypes import windll, wintypes

from .dpi import ensure_process_dpi_aware_once

user32 = windll.user32

ensure_process_dpi_aware_once()

POINT = wintypes.POINT

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
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002
KEYEVENTF_UNICODE = 0x0004


class CURSORINFO(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("flags", wintypes.DWORD),
        ("hCursor", wintypes.HANDLE),
        ("ptScreenPos", POINT),
    ]


CURSOR_SHOWING = 0x00000001

user32.GetCursorInfo.argtypes = [ctypes.POINTER(CURSORINFO)]
user32.GetCursorInfo.restype = wintypes.BOOL
user32.LoadCursorW.argtypes = [wintypes.HINSTANCE, ctypes.c_void_p]
user32.LoadCursorW.restype = wintypes.HANDLE

IDC_ARROW = 32512
IDC_IBEAM = 32513
IDC_WAIT = 32514
IDC_CROSS = 32515
IDC_SIZENWSE = 32642
IDC_SIZENESW = 32643
IDC_SIZEWE = 32644
IDC_SIZENS = 32645
IDC_SIZEALL = 32646
IDC_NO = 32648
IDC_HAND = 32649
IDC_APPSTARTING = 32650
IDC_HELP = 32651

_CURSOR_HANDLE_TO_CSS = {
    int(handle): css
    for css, cid in {
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
    }.items()
    if (handle := user32.LoadCursorW(None, ctypes.c_void_p(cid)))
}


def _get_css_cursor(cursor_handle) -> str:
    return _CURSOR_HANDLE_TO_CSS.get(int(cursor_handle), "default")


ULONG_PTR = ctypes.c_ulonglong

class MOUSEINPUT(ctypes.Structure):
    _fields_ = (
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    )

class KEYBDINPUT(ctypes.Structure):
    _fields_ = (
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    )

class _INPUT_UNION(ctypes.Union):
    _fields_ = (("mi", MOUSEINPUT), ("ki", KEYBDINPUT))

class INPUT(ctypes.Structure):
    _fields_ = (("type", wintypes.DWORD), ("union", _INPUT_UNION))

def _sendinput_mouse_move_abs(ax, ay):
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
    if user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) != 1:
        raise OSError(f"SendInput mouse move failed (GetLastError={ctypes.get_last_error()})")

def _sendinput_mouse_button(flag):
    inp = INPUT()
    inp.type = 0
    inp.union.mi = MOUSEINPUT(0, 0, 0, flag, 0, 0)
    if user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) != 1:
        raise OSError(f"SendInput mouse button failed (GetLastError={ctypes.get_last_error()})")


_VK_CODES = {
    "backspace": 0x08,
    "tab": 0x09,
    "enter": 0x0D,
    "shift": 0x10,
    "ctrl": 0x11,
    "alt": 0x12,
    "pause": 0x13,
    "capslock": 0x14,
    "esc": 0x1B,
    "space": 0x20,
    "pageup": 0x21,
    "pagedown": 0x22,
    "end": 0x23,
    "home": 0x24,
    "left": 0x25,
    "up": 0x26,
    "right": 0x27,
    "down": 0x28,
    "insert": 0x2D,
    "delete": 0x2E,
    "winleft": 0x5B,
}
_EXTENDED_KEYS = {"pageup", "pagedown", "end", "home", "left", "up", "right", "down", "insert", "delete", "winleft"}


def _virtual_key(name):
    name = str(name).lower()
    if name in _VK_CODES:
        return _VK_CODES[name], name in _EXTENDED_KEYS
    if len(name) == 1:
        code = user32.VkKeyScanW(ord(name))
        if code != -1:
            return code & 0xFF, False
    if name.startswith("f") and name[1:].isdigit() and 1 <= int(name[1:]) <= 24:
        return 0x6F + int(name[1:]), False
    raise ValueError(f"Unsupported keyboard key: {name}")


def _sendinput_key(name, state):
    if state not in {"down", "up", "press"}:
        raise ValueError(f"Invalid keyboard state: {state}")
    vk, extended = _virtual_key(name)
    states = (False, True) if state == "press" else (state == "up",)
    inputs = (INPUT * len(states))()
    for index, key_up in enumerate(states):
        flags = KEYEVENTF_EXTENDEDKEY if extended else 0
        if key_up:
            flags |= KEYEVENTF_KEYUP
        inputs[index].type = 1
        inputs[index].union.ki = KEYBDINPUT(vk, 0, flags, 0, 0)
    sent = user32.SendInput(len(inputs), inputs, ctypes.sizeof(INPUT))
    if sent != len(inputs):
        raise OSError(
            f"SendInput key {name} {state} sent {sent} of {len(inputs)} events "
            f"(GetLastError={ctypes.get_last_error()})"
        )


def _sendinput_unicode(text):
    units = memoryview(str(text).encode("utf-16-le")).cast("H")
    for offset in range(0, len(units), 256):
        chunk = units[offset:offset + 256]
        inputs = (INPUT * (len(chunk) * 2))()
        for index, code_unit in enumerate(chunk):
            inputs[index * 2].type = 1
            inputs[index * 2].union.ki = KEYBDINPUT(0, code_unit, KEYEVENTF_UNICODE, 0, 0)
            inputs[index * 2 + 1].type = 1
            inputs[index * 2 + 1].union.ki = KEYBDINPUT(
                0,
                code_unit,
                KEYEVENTF_UNICODE | KEYEVENTF_KEYUP,
                0,
                0,
            )
        count = len(inputs)
        sent = user32.SendInput(count, inputs, ctypes.sizeof(INPUT))
        if sent != count:
            raise OSError(
                f"SendInput Unicode typing sent {sent} of {count} events "
                f"(GetLastError={ctypes.get_last_error()})"
            )

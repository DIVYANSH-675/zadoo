"""Windows DPI awareness helpers."""
from __future__ import annotations

import sys

_DPI_AWARENESS_ATTEMPTED = False


def ensure_process_dpi_aware_once():
    """Set Win32 process DPI awareness at most once per process."""
    global _DPI_AWARENESS_ATTEMPTED
    if _DPI_AWARENESS_ATTEMPTED or sys.platform != "win32":
        return
    _DPI_AWARENESS_ATTEMPTED = True
    try:
        import ctypes as _ct
        _ct.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


def get_primary_screen_size():
    if sys.platform == "win32":
        try:
            import ctypes as _ct
            ensure_process_dpi_aware_once()
            u32 = _ct.windll.user32
            return int(u32.GetSystemMetrics(0)), int(u32.GetSystemMetrics(1))
        except Exception:
            pass
    import pyautogui as _pg
    return _pg.size()

"""Windows DPI awareness helpers."""
from __future__ import annotations

import ctypes

_DPI_AWARENESS_ATTEMPTED = False


def ensure_process_dpi_aware_once():
    """Set Win32 process DPI awareness at most once per process."""
    global _DPI_AWARENESS_ATTEMPTED
    if _DPI_AWARENESS_ATTEMPTED:
        return
    try:
        if not ctypes.windll.user32.SetProcessDPIAware():
            raise ctypes.WinError()
    except OSError as exc:
        raise RuntimeError(f"Failed to enable Windows DPI awareness: {exc}") from exc
    _DPI_AWARENESS_ATTEMPTED = True


def get_primary_screen_size():
    ensure_process_dpi_aware_once()
    size = int(ctypes.windll.user32.GetSystemMetrics(0)), int(ctypes.windll.user32.GetSystemMetrics(1))
    if min(size) <= 0:
        raise RuntimeError(f"Windows reported an invalid primary screen size: {size[0]}x{size[1]}")
    return size

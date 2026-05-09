"""Subprocess helpers."""
from __future__ import annotations

import subprocess as _sp


def _run_hidden(cmd, *, timeout=8, text=True):
    """Run a child process without flashing a window in a frozen GUI app."""
    si = _sp.STARTUPINFO()
    si.dwFlags |= _sp.STARTF_USESHOWWINDOW
    cf = getattr(_sp, "CREATE_NO_WINDOW", 0)
    return _sp.run(
        cmd,
        capture_output=True,
        text=text,
        timeout=timeout,
        startupinfo=si,
        creationflags=cf,
    )

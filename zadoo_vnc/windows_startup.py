"""Windows startup task helpers for the Zadoo Settings app."""
from __future__ import annotations

import subprocess
import sys

from .config import windows_system_executable

TASK_NAME = "Zadoo"


def startup_task_exists() -> bool:
    result = subprocess.run(
        [windows_system_executable("schtasks.exe"), "/query", "/tn", TASK_NAME],
        capture_output=True,
        text=True,
        timeout=10,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    if result.returncode == 0:
        return True
    message = (result.stderr or result.stdout or "").strip()
    if "cannot find" in message.lower() or "does not exist" in message.lower():
        return False
    raise RuntimeError(message or f"schtasks query exited {result.returncode}")


def set_startup_task(enabled: bool) -> tuple[bool, str]:
    if enabled:
        command = [
            windows_system_executable("schtasks.exe"),
            "/create",
            "/tn",
            TASK_NAME,
            "/tr",
            f'"{sys.executable}"' + ("" if getattr(sys, "frozen", False) else " -m zadoo_vnc"),
            "/sc",
            "onlogon",
            "/rl",
            "highest",
            "/f",
        ]
    else:
        command = [windows_system_executable("schtasks.exe"), "/delete", "/tn", TASK_NAME, "/f"]
    try:
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    except subprocess.TimeoutExpired:
        return False, "schtasks did not finish within 10 seconds"
    if result.returncode == 0:
        return True, "Startup enabled" if enabled else "Startup disabled"
    message = (result.stderr or result.stdout or "").strip() or f"schtasks exited {result.returncode}"
    return False, message

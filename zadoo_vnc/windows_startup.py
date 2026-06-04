"""Windows startup task helpers for the Zadoo Settings app."""
from __future__ import annotations

import os
import subprocess
import sys

TASK_NAME = "Zadoo"


def _hidden_flag() -> int:
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _runtime_command_text() -> str:
    if getattr(sys, "frozen", False):
        return f'"{sys.executable}" --open'
    return f'"{sys.executable}" -m zadoo_vnc.app --open'


def startup_task_exists() -> bool:
    if os.name != "nt":
        return False
    result = subprocess.run(
        ["schtasks", "/query", "/tn", TASK_NAME],
        capture_output=True,
        creationflags=_hidden_flag(),
    )
    return result.returncode == 0


def set_startup_task(enabled: bool) -> tuple[bool, str]:
    if os.name != "nt":
        return False, "Startup tasks are only available on Windows"
    if enabled:
        command = [
            "schtasks",
            "/create",
            "/tn",
            TASK_NAME,
            "/tr",
            _runtime_command_text(),
            "/sc",
            "onlogon",
            "/rl",
            "highest",
            "/f",
        ]
    else:
        command = ["schtasks", "/delete", "/tn", TASK_NAME, "/f"]
    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        creationflags=_hidden_flag(),
    )
    if result.returncode == 0:
        return True, "Startup enabled" if enabled else "Startup disabled"
    message = (result.stderr or result.stdout or "").strip() or f"schtasks exited {result.returncode}"
    return False, message

"""Configuration constants and filesystem paths."""
from __future__ import annotations

import os
import sys
from pathlib import Path

from dotenv import load_dotenv

PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PACKAGE_DIR.parent
APP_PORT = 6173
TRUE_VALUES = {"1", "true", "yes", "on", "enabled"}
FALSE_VALUES = {"0", "false", "no", "off", "disabled"}


def resource_path(name: str) -> Path:
    """Resolve a resource in the source tree or pinned PyInstaller layout."""
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / name  # type: ignore[attr-defined]
    return PROJECT_DIR / name


def windows_system_executable(*parts: str) -> str:
    if sys.platform != "win32":
        raise RuntimeError(f"Windows system executables require Windows; current platform is {sys.platform}")
    system_root = os.environ.get("SYSTEMROOT", "").strip()
    if not system_root:
        raise RuntimeError("SYSTEMROOT environment variable is not set")
    path = Path(system_root, "System32", *parts)
    if not path.is_file():
        raise FileNotFoundError(f"Windows system executable not found: {path}")
    return str(path)


def env_int(name, default, minimum=None, maximum=None):
    raw_value = str(os.getenv(name, default)).strip()
    try:
        value = int(raw_value)
    except ValueError as exc:
        raise ValueError(f"{name} must be an integer; got {raw_value!r}") from exc
    if minimum is not None and value < int(minimum):
        raise ValueError(f"{name} must be at least {minimum}; got {value}")
    if maximum is not None and value > int(maximum):
        raise ValueError(f"{name} must be at most {maximum}; got {value}")
    return value


def env_bool(name, default=False):
    raw_value = os.getenv(name)
    if raw_value is None:
        return bool(default)
    normalized = str(raw_value).strip().lower()
    if normalized in TRUE_VALUES:
        return True
    if normalized in FALSE_VALUES:
        return False
    raise ValueError(f"{name} must be a boolean; got {raw_value!r}")


def _load_dotenv():
    """Load source-development overrides from the project .env file."""
    load_dotenv(PROJECT_DIR / ".env", override=False)

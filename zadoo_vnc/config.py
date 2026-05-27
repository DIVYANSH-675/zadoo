"""Configuration constants and filesystem paths."""
from __future__ import annotations

import os
import sys
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PACKAGE_DIR.parent
RUNTIME_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else PROJECT_DIR

QUALITY = 85
STOP_FILE = "stop_vnc.flag"
BRAND_HEADER_IMAGE_PATH = str(RUNTIME_DIR / "brand-header.png")
TRIGGER_ICON_IMAGE_PATH = str(RUNTIME_DIR / "trigger-icon.png")
SPLASH_IMAGE_PATH = str(RUNTIME_DIR / "splash.png")


def env_int(name, default, minimum=None, maximum=None):
    try:
        value = int(str(os.getenv(name, default)).strip())
    except Exception:
        value = int(default)
    if minimum is not None:
        value = max(int(minimum), value)
    if maximum is not None:
        value = min(int(maximum), value)
    return value


def _load_dotenv(path: str = ".env"):
    """Load simple KEY=VALUE pairs from .env in CWD and project directory."""
    loaded_paths = set()

    def _apply(p: str):
        try:
            resolved = os.path.realpath(p)
            if resolved in loaded_paths or not os.path.exists(resolved):
                return
            loaded_paths.add(resolved)
            with open(resolved, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#"):
                        continue
                    if "=" not in line:
                        continue
                    key, val = line.split("=", 1)
                    key = key.strip()
                    val = val.strip().strip('"').strip("'")
                    os.environ[key] = val
        except Exception:
            pass

    _apply(path)
    try:
        _apply(str(PROJECT_DIR / ".env"))
    except Exception:
        pass

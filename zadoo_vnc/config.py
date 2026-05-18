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


def _load_dotenv(path: str = ".env"):
    """Load simple KEY=VALUE pairs from .env in CWD and project directory."""
    def _apply(p: str):
        try:
            if not os.path.exists(p):
                return
            with open(p, "r", encoding="utf-8", errors="ignore") as f:
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

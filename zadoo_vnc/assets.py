"""Template and static asset helpers."""
from __future__ import annotations

import sys
from pathlib import Path

PACKAGE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = PACKAGE_DIR.parent


def _template_candidates(name: str):
    yield PACKAGE_DIR / "templates" / name
    if getattr(sys, "frozen", False):
        try:
            yield Path(sys._MEIPASS) / "zadoo_vnc" / "templates" / name
            yield Path(sys._MEIPASS) / "templates" / name
        except Exception:
            pass
    yield PROJECT_DIR / "zadoo_vnc" / "templates" / name


def load_template(name: str) -> str:
    for path in _template_candidates(name):
        try:
            if path.exists():
                return path.read_text(encoding="utf-8")
        except Exception:
            continue
    raise FileNotFoundError(f"Unable to locate template: {name}")


def load_index_html() -> str:
    return load_template("index.html")


def load_terminal_html() -> str:
    return load_template("terminal.html")


def load_host_controls_html() -> str:
    return load_template("host_controls.html")

"""Template and static asset helpers."""
from __future__ import annotations

from functools import cache
from pathlib import Path

from .config import resource_path

PACKAGE_DIR = Path(__file__).resolve().parent
STATIC_DIR = PACKAGE_DIR / "static"


@cache
def load_template(name: str) -> str:
    path = PACKAGE_DIR / "templates" / name
    if not path.is_file():
        raise FileNotFoundError(f"Template not found: {path}")
    return path.read_text(encoding="utf-8")


@cache
def load_binary(name: str) -> bytes:
    return resource_path(name).read_bytes()


@cache
def load_static(name: str) -> bytes:
    path = STATIC_DIR / name
    if not path.is_file():
        raise FileNotFoundError(f"Static asset not found: {path}")
    return path.read_bytes()

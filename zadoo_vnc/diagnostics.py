"""Bounded, redacted local diagnostic bundle generation."""
from __future__ import annotations

import json
import os
import platform
import re
import sys
import time
import zipfile
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from . import __version__
from .config import PROJECT_DIR, env_int
from .settings import SettingsStore, get_settings_store, settings_dir

_EMAIL_RE = re.compile(r"(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b")
_BEARER_RE = re.compile(r"(?i)(authorization\s*[:=]\s*bearer\s+)[^\s,;]+")
_COOKIE_RE = re.compile(r"(?i)\b(zadoo_(?:auth|csrf)=)[^;\s]+")
_DPAPI_RE = re.compile(r"\bdpapi(?:-user)?:[A-Za-z0-9+/=]+")
_TUNNEL_RE = re.compile(r"https://[a-z0-9-]+\.trycloudflare\.com", re.IGNORECASE)
_USER_PROFILE_RE = re.compile(r"(?i)\b[A-Z]:\\Users\\[^\\\s\"']+")
_SECRET_FIELD_RE = re.compile(
    r'(?i)(["\']?(?:access_code|device_token|poll_secret|pfx_password|api_key|'
    r'workspace_id|device_id|activation_id|email_to|user_email)["\']?\s*[:=]\s*["\']?)'
    r'([^"\'\s,}]+)'
)


def _log_dir() -> Path:
    return settings_dir() / "logs" if getattr(sys, "frozen", False) else PROJECT_DIR / "logs"


def _secret_values(store: SettingsStore) -> list[str]:
    data = store.load(reload=True)
    values = [
        store.get_access_code(),
        store.get_device_token(),
        data.get("email_to", ""),
        data.get("user_email", ""),
        data.get("workspace_id", ""),
        data.get("device_id", ""),
    ]
    activation = data.get("activation", {})
    if isinstance(activation, dict):
        values.extend(str(activation.get(key, "")) for key in ("poll_secret", "activation_id"))
    return sorted({value for value in values if isinstance(value, str) and value}, key=len, reverse=True)


def redact_diagnostic_text(text: str, secrets: list[str] | tuple[str, ...] = ()) -> str:
    if not isinstance(text, str):
        raise TypeError("Diagnostic text must be a string")
    redacted = text
    for value in secrets:
        redacted = redacted.replace(value, "[REDACTED]")
    redacted = _BEARER_RE.sub(r"\1[REDACTED]", redacted)
    redacted = _COOKIE_RE.sub(r"\1[REDACTED]", redacted)
    redacted = _DPAPI_RE.sub("[REDACTED_DPAPI]", redacted)
    redacted = _SECRET_FIELD_RE.sub(r"\1[REDACTED]", redacted)
    redacted = _EMAIL_RE.sub("[REDACTED_EMAIL]", redacted)
    redacted = _USER_PROFILE_RE.sub("[REDACTED_USER_PROFILE]", redacted)
    return _TUNNEL_RE.sub("https://[REDACTED].trycloudflare.com", redacted)


def _diagnostic_metadata(store: SettingsStore, runtime_status: dict[str, Any] | None) -> dict[str, Any]:
    data = store.load(reload=True)
    status = runtime_status if isinstance(runtime_status, dict) else {"available": False}
    return {
        "schema": 1,
        "created_utc": datetime.now(UTC).isoformat(),
        "zadoo_version": __version__,
        "platform": platform.platform(),
        "python_version": platform.python_version(),
        "frozen_build": bool(getattr(sys, "frozen", False)),
        "settings": {
            "location": (
                "override"
                if os.environ.get("ZADOO_SETTINGS_DIR") is not None
                or os.environ.get("ZADOO_SETTINGS_PATH") is not None
                else "current-user-local-app-data"
            ),
            "configured": bool(data["setup_complete"]),
            "signed_in": bool(store.get_device_token()),
            "storage_scope": "current-windows-user",
        },
        "runtime": status,
        "log_policy": {
            "retention_days": env_int("ZADOO_LOG_RETENTION_DAYS", 14, 1, 365),
            "max_bytes": env_int(
                "ZADOO_LOG_MAX_BYTES", 5 * 1024 * 1024, 65_536, 1024 * 1024 * 1024
            ),
            "backup_count": env_int("ZADOO_LOG_BACKUP_COUNT", 2, 1, 20),
        },
    }


def build_diagnostic_bundle(
    destination: Path,
    *,
    store: SettingsStore | None = None,
    runtime_status: dict[str, Any] | None = None,
) -> Path:
    """Write an atomic ZIP containing metadata and bounded, redacted log tails."""
    target = Path(destination).expanduser().resolve()
    if target.suffix.lower() != ".zip":
        raise ValueError("Diagnostic bundle path must end with .zip")
    target.parent.mkdir(parents=True, exist_ok=True)
    store = store if store is not None else get_settings_store()
    secrets = _secret_values(store)
    max_log_bytes = env_int(
        "ZADOO_DIAGNOSTIC_MAX_LOG_BYTES", 2 * 1024 * 1024, 65_536, 5 * 1024 * 1024
    )
    max_log_files = env_int("ZADOO_DIAGNOSTIC_MAX_LOG_FILES", 6, 1, 20)
    logs = sorted(
        (path for path in _log_dir().glob("zadoo_*.log*") if path.is_file()),
        key=lambda path: path.stat().st_mtime_ns,
        reverse=True,
    )[:max_log_files]
    temp = target.with_name(f".{target.name}.{os.getpid()}.{time.time_ns()}.tmp")
    try:
        with zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
            metadata = _diagnostic_metadata(store, runtime_status)
            metadata_text = redact_diagnostic_text(
                json.dumps(metadata, indent=2, sort_keys=True), secrets
            )
            archive.writestr("diagnostics.json", metadata_text.encode("utf-8"))
            for path in logs:
                with path.open("rb") as handle:
                    size = path.stat().st_size
                    if size > max_log_bytes:
                        handle.seek(size - max_log_bytes)
                    raw = handle.read(max_log_bytes)
                text = raw.decode("utf-8", errors="replace")
                archive.writestr(
                    f"logs/{path.name}",
                    redact_diagnostic_text(text, secrets).encode("utf-8"),
                )
        temp.replace(target)
    except Exception:
        temp.unlink(missing_ok=True)
        raise
    return target

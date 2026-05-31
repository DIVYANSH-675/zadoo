"""Persistent installed-app settings for Zadoo."""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import secrets
import sys
import time
from pathlib import Path
from typing import Any

from .config import PROJECT_DIR

APP_NAME = "Zadoo"
CONFIG_VERSION = 1
ACCESS_CODE_ITERATIONS = 260_000
ACCESS_CODE_MAX_LENGTH = 10
FALSE_VALUES = {"0", "false", "no", "off", "disabled"}

PERMISSION_KEYS = (
    "mouse",
    "keyboard",
    "clipboard_pull",
    "clipboard_push",
    "system_audio",
    "mic",
    "camera",
    "terminal",
    "snapshots",
    "advanced_video",
    "tunnel_refresh",
    "remote_alerts",
)

DEFAULT_PERMISSIONS = {key: False for key in PERMISSION_KEYS}


def _program_data_dir() -> Path:
    if sys.platform.startswith("win"):
        base = os.environ.get("ProgramData") or r"C:\ProgramData"
        return Path(base) / APP_NAME
    return PROJECT_DIR / ".zadoo"


def settings_dir() -> Path:
    override = os.environ.get("ZADOO_SETTINGS_DIR")
    return Path(override).expanduser().resolve() if override else _program_data_dir()


def settings_path() -> Path:
    override = os.environ.get("ZADOO_SETTINGS_PATH")
    return Path(override).expanduser().resolve() if override else settings_dir() / "config.json"


def _now() -> float:
    return time.time()


def _truthy(value: Any) -> bool:
    return str(value or "").strip().lower() not in FALSE_VALUES


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _unb64(text: str) -> bytes:
    return base64.b64decode(str(text or "").encode("ascii"), validate=True)


def _dpapi_encrypt(text: str) -> str:
    if not text:
        return ""
    if not sys.platform.startswith("win"):
        return "plain:" + _b64(text.encode("utf-8"))
    try:
        import win32crypt

        encrypted = win32crypt.CryptProtectData(
            text.encode("utf-8"),
            "Zadoo",
            None,
            None,
            None,
            0x4,  # CRYPTPROTECT_LOCAL_MACHINE
        )
        return "dpapi:" + _b64(encrypted)
    except Exception as exc:
        raise RuntimeError("Windows DPAPI encryption failed") from exc


def _dpapi_decrypt(value: str) -> str:
    value = str(value or "")
    if not value:
        return ""
    if value.startswith("plain:"):
        return _unb64(value[6:]).decode("utf-8", "replace")
    if not value.startswith("dpapi:"):
        return ""
    if not sys.platform.startswith("win"):
        return ""
    try:
        import win32crypt

        decrypted = win32crypt.CryptUnprotectData(_unb64(value[6:]), None, None, None, 0)[1]
        return decrypted.decode("utf-8", "replace")
    except Exception:
        return ""


def hash_access_code(code: str, *, salt: str | None = None) -> dict[str, Any]:
    clean = str(code or "").strip()
    if not clean:
        raise ValueError("access code is required")
    if len(clean) > ACCESS_CODE_MAX_LENGTH:
        raise ValueError(f"access code must be {ACCESS_CODE_MAX_LENGTH} characters or fewer")
    salt = salt or _b64(secrets.token_bytes(16))
    digest = hashlib.pbkdf2_hmac(
        "sha256",
        clean.encode("utf-8"),
        _unb64(salt),
        ACCESS_CODE_ITERATIONS,
    )
    return {
        "algorithm": "pbkdf2_sha256",
        "iterations": ACCESS_CODE_ITERATIONS,
        "salt": salt,
        "hash": _b64(digest),
    }


def verify_access_code(code: str, record: dict[str, Any] | None) -> bool:
    clean = str(code or "").strip()
    if not clean or not isinstance(record, dict):
        return False
    try:
        iterations = int(record.get("iterations") or ACCESS_CODE_ITERATIONS)
        salt = str(record.get("salt") or "")
        expected = str(record.get("hash") or "")
        actual = hashlib.pbkdf2_hmac("sha256", clean.encode("utf-8"), _unb64(salt), iterations)
        return hmac.compare_digest(_b64(actual), expected)
    except Exception:
        return False


def _empty_alerts() -> dict[str, dict[str, Any]]:
    return {
        key: {"title": "", "message": "", "enabled": False}
        for key in ("A", "B", "C", "D")
    }


def _default_data() -> dict[str, Any]:
    return {
        "version": CONFIG_VERSION,
        "setup_complete": False,
        "created_at": _now(),
        "updated_at": _now(),
        "env_alerts_imported": False,
        "access_code": None,
        "access_code_plain": "",
        "email_to": "",
        "resend_api_key": "",
        "permissions": dict(DEFAULT_PERMISSIONS),
        "alerts": _empty_alerts(),
    }


def _clean_permissions(raw: Any) -> dict[str, bool]:
    raw = raw if isinstance(raw, dict) else {}
    return {key: bool(raw.get(key, DEFAULT_PERMISSIONS[key])) for key in PERMISSION_KEYS}


def normalize_settings(data: dict[str, Any] | None) -> dict[str, Any]:
    base = _default_data()
    raw_data = data if isinstance(data, dict) else {}
    if isinstance(data, dict):
        for key in list(base.keys()):
            if key in data:
                base[key] = data[key]
    base["version"] = CONFIG_VERSION
    base["setup_complete"] = bool(base.get("setup_complete", False))
    base["access_code_plain"] = str(base.get("access_code_plain") or "").strip()[:ACCESS_CODE_MAX_LENGTH]
    permissions_source = raw_data.get("permissions") if "permissions" in raw_data else None
    base["permissions"] = _clean_permissions(permissions_source)
    alerts = _empty_alerts()
    for key, item in (base.get("alerts") or {}).items():
        code = str(key or "").upper()
        if code not in alerts or not isinstance(item, dict):
            continue
        title = str(item.get("title") or "").strip()
        message = str(item.get("message") or "").strip()
        alerts[code] = {
            "title": title,
            "message": message,
            "enabled": bool(item.get("enabled", bool(title or message))) and bool(title or message),
        }
    base["alerts"] = alerts
    return base


class SettingsStore:
    def __init__(self, path: Path | None = None):
        self.path = path or settings_path()
        self._data: dict[str, Any] | None = None
        self._mtime_ns: int | None = None

    def load(self, *, reload: bool = False) -> dict[str, Any]:
        try:
            mtime_ns = self.path.stat().st_mtime_ns
        except Exception:
            mtime_ns = None
        if self._data is not None and not reload and mtime_ns == self._mtime_ns:
            return self._data
        try:
            raw = json.loads(self.path.read_text(encoding="utf-8"))
        except Exception:
            raw = None
        data = normalize_settings(raw)
        self._data = data
        self._mtime_ns = mtime_ns
        if not data.get("env_alerts_imported"):
            self.import_env_alerts(save=False)
        return self._data

    def save(self, data: dict[str, Any] | None = None) -> dict[str, Any]:
        clean = normalize_settings(data if data is not None else self.load())
        clean["updated_at"] = _now()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(clean, indent=2, sort_keys=True), encoding="utf-8")
        tmp.replace(self.path)
        self._data = clean
        try:
            self._mtime_ns = self.path.stat().st_mtime_ns
        except Exception:
            self._mtime_ns = None
        return clean

    def configured(self) -> bool:
        return bool(self.load().get("setup_complete"))

    def import_env_alerts(self, *, save: bool = True) -> dict[str, Any]:
        data = self.load()
        if data.get("env_alerts_imported"):
            return data
        changed = False
        for key in ("A", "B", "C", "D"):
            combined = os.getenv(f"ALERT_{key}")
            title = os.getenv(f"ALERT_{key}_TITLE")
            message = os.getenv(f"ALERT_{key}_MESSAGE")
            if (not title and not message) and combined:
                sep = "|" if "|" in combined else "::" if "::" in combined else None
                if sep:
                    title, message = combined.split(sep, 1)
                else:
                    title, message = combined, ""
            title = str(title or "").strip()
            message = str(message or "").strip()
            if title or message:
                data["alerts"][key] = {"title": title, "message": message, "enabled": True}
                changed = True
        data["env_alerts_imported"] = True
        if save and changed:
            self.save(data)
        return data

    def set_access_code(self, code: str) -> None:
        data = self.load()
        clean = str(code or "").strip()
        data["access_code_plain"] = clean[:ACCESS_CODE_MAX_LENGTH]
        data["access_code"] = hash_access_code(clean)
        self.save(data)

    def verify_access_code(self, code: str) -> bool:
        clean = str(code or "").strip()
        data = self.load()
        plain = str(data.get("access_code_plain") or "").strip()
        if plain:
            return bool(clean) and secrets.compare_digest(clean, plain)
        return verify_access_code(clean, data.get("access_code"))

    def set_resend_api_key(self, api_key: str) -> None:
        data = self.load()
        data["resend_api_key"] = _dpapi_encrypt(str(api_key or "").strip())
        self.save(data)

    def get_resend_api_key(self) -> str:
        return _dpapi_decrypt(str(self.load().get("resend_api_key") or ""))

    def public_view(self, *, include_secret: bool = False) -> dict[str, Any]:
        data = self.load()
        view = {
            "setup_complete": bool(data.get("setup_complete")),
            "access_code": data.get("access_code_plain") or "",
            "email_to": data.get("email_to") or "",
            "has_resend_api_key": bool(self.get_resend_api_key()),
            "permissions": data.get("permissions") or dict(DEFAULT_PERMISSIONS),
            "alerts": data.get("alerts") or _empty_alerts(),
        }
        if include_secret:
            view["resend_api_key"] = self.get_resend_api_key()
        return view

    def apply_setup(self, payload: dict[str, Any], *, require_code: str | None = None) -> dict[str, Any]:
        data = self.load()
        if data.get("setup_complete") and not self.verify_access_code(str(require_code or "")):
            raise PermissionError("Invalid access code")
        new_code = str(payload.get("access_code") or "").strip()
        if new_code:
            if len(new_code) > ACCESS_CODE_MAX_LENGTH:
                raise ValueError(f"access code must be {ACCESS_CODE_MAX_LENGTH} characters or fewer")
            data["access_code_plain"] = new_code
            data["access_code"] = hash_access_code(new_code)
        elif not data.get("access_code") and not data.get("access_code_plain"):
            raise ValueError("access code is required")

        if "email_to" in payload:
            data["email_to"] = str(payload.get("email_to") or "").strip()
        if bool(payload.get("clear_resend_api_key", False)):
            data["resend_api_key"] = ""
        elif "resend_api_key" in payload:
            api_key = str(payload.get("resend_api_key") or "").strip()
            if api_key:
                data["resend_api_key"] = _dpapi_encrypt(api_key)

        alerts = _empty_alerts()
        for key, item in (payload.get("alerts") or {}).items():
            code = str(key or "").upper()
            if code not in alerts or not isinstance(item, dict):
                continue
            title = str(item.get("title") or "").strip()
            message = str(item.get("message") or "").strip()
            alerts[code] = {"title": title, "message": message, "enabled": bool(item.get("enabled", bool(title or message))) and bool(title or message)}
        data["alerts"] = alerts
        if "permissions" in payload:
            data["permissions"] = _clean_permissions(payload.get("permissions"))
        data["setup_complete"] = True
        return self.save(data)


_STORE: SettingsStore | None = None


def get_settings_store() -> SettingsStore:
    global _STORE
    if _STORE is None:
        _STORE = SettingsStore()
    return _STORE

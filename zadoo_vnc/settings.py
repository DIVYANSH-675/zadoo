"""Persistent installed-app settings for Zadoo."""
from __future__ import annotations

import base64
import copy
import json
import math
import msvcrt
import os
import secrets
import sys
import threading
import time
import urllib.parse
import winreg
from contextlib import contextmanager
from pathlib import Path
from typing import Any

import win32crypt

from . import __version__


def _clean_http_origin(value: Any, field: str) -> str:
    if not isinstance(value, str):
        raise ValueError(f"{field} must be a string")
    clean = value.strip().rstrip("/")
    parsed = urllib.parse.urlparse(clean)
    try:
        _ = parsed.port
    except ValueError as exc:
        raise ValueError(f"{field} must be an HTTP(S) origin without a path query or fragment") from exc
    if (
        parsed.scheme not in {"http", "https"}
        or not parsed.hostname
        or parsed.username is not None
        or parsed.password is not None
        or parsed.netloc.endswith(":")
        or any(char.isspace() for char in clean)
        or "?" in clean
        or "#" in clean
        or parsed.params
        or parsed.query
        or parsed.fragment
    ):
        raise ValueError(f"{field} must be an HTTP(S) origin without a path query or fragment")
    if parsed.path not in {"", "/"}:
        raise ValueError(f"{field} must be an HTTP(S) origin without a path query or fragment")
    return clean


APP_NAME = "Zadoo"
CONFIG_VERSION = 3
ACCESS_CODE_MAX_LENGTH = 10
DEFAULT_CLOUD_API_BASE = _clean_http_origin(
    os.getenv("ZADOO_CLOUD_API_BASE", "https://zadoo-web.vercel.app"),
    "ZADOO_CLOUD_API_BASE",
)

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
ALERT_CODES = ("A", "B", "C", "D")

# Default: everything allowed except the shell/terminal (opt-in for safety).
DEFAULT_PERMISSIONS = {key: (key != "terminal") for key in PERMISSION_KEYS}


def _program_data_dir() -> Path:
    if not sys.platform.startswith("win"):
        raise RuntimeError(f"Zadoo settings require Windows; current platform is {sys.platform}")
    program_data = os.environ.get("PROGRAMDATA", "").strip()
    if not program_data:
        raise RuntimeError("ProgramData environment variable is not set")
    return Path(program_data) / APP_NAME


def settings_dir() -> Path:
    override = os.environ.get("ZADOO_SETTINGS_DIR")
    if override is None:
        return _program_data_dir()
    if not override.strip():
        raise RuntimeError("ZADOO_SETTINGS_DIR must not be empty")
    return Path(override).expanduser().resolve()


def settings_path() -> Path:
    override = os.environ.get("ZADOO_SETTINGS_PATH")
    if override is None:
        return settings_dir() / "config.json"
    if not override.strip():
        raise RuntimeError("ZADOO_SETTINGS_PATH must not be empty")
    return Path(override).expanduser().resolve()


# Compatibility name used by the SaaS client; the package owns the version.
APP_VERSION = __version__


def machine_id() -> str:
    """Return the stable Windows MachineGuid used for device identity."""
    if not sys.platform.startswith("win"):
        raise RuntimeError(f"Machine identity requires Windows; current platform is {sys.platform}")
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Cryptography") as key:
            value, _ = winreg.QueryValueEx(key, "MachineGuid")
    except OSError as exc:
        raise RuntimeError(f"Unable to read Windows MachineGuid: {exc}") from exc
    if not isinstance(value, str) or not value.strip():
        raise RuntimeError("Windows MachineGuid is empty")
    return value.strip()


def _b64(data: bytes) -> str:
    return base64.b64encode(data).decode("ascii")


def _unb64(text: str) -> bytes:
    if not isinstance(text, str):
        raise ValueError("Encrypted setting payload must be text")
    return base64.b64decode(text.encode("ascii"), validate=True)


def _dpapi_encrypt(text: str) -> str:
    if not text:
        return ""
    try:
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
        raise RuntimeError(f"Windows DPAPI encryption failed: {exc}") from exc


def _dpapi_decrypt(value: str) -> str:
    if not isinstance(value, str):
        raise ValueError("Encrypted setting must be text")
    if not value:
        return ""
    if not value.startswith("dpapi:"):
        raise ValueError("Encrypted setting must use the dpapi: format")
    try:
        decrypted = win32crypt.CryptUnprotectData(_unb64(value[6:]), None, None, None, 0)[1]
        return decrypted.decode("utf-8")
    except Exception as exc:
        raise RuntimeError(f"Windows DPAPI decryption failed: {exc}") from exc


def _validate_access_code(code: str) -> str:
    if not isinstance(code, str):
        raise ValueError("access code must be a string")
    clean = code.strip()
    if not clean:
        raise ValueError("access code is required")
    if len(clean) > ACCESS_CODE_MAX_LENGTH:
        raise ValueError(f"access code must be {ACCESS_CODE_MAX_LENGTH} characters or fewer")
    return clean


def _empty_alerts() -> dict[str, dict[str, Any]]:
    return {
        key: {"title": "", "message": "", "enabled": False}
        for key in ALERT_CODES
    }


def _nonnegative_int_field(data: dict[str, Any], field: str, context: str) -> int:
    value = data.get(field)
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ValueError(f"{context} {field} must be a non-negative integer")
    return value


def _clean_entitlement(value: Any, context: str, *, allow_empty: bool = False) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{context} must be an object")
    if allow_empty and not value:
        return {}
    required = {
        "allowed", "revoked", "reason", "planCode", "includedMinutesRemaining",
        "walletMinutesRemaining", "concurrencyLimit", "activeSessions", "graceUntil",
    }
    unknown = sorted(set(value) - required - {"minutesRemaining"})
    missing = sorted(required - set(value))
    if unknown:
        raise ValueError(f"{context} has unknown fields: {', '.join(unknown)}")
    if missing:
        raise ValueError(f"{context} is missing fields: {', '.join(missing)}")
    for field in ("allowed", "revoked"):
        if type(value[field]) is not bool:
            raise ValueError(f"{context} {field} must be a boolean")
    for field in (
        "includedMinutesRemaining", "walletMinutesRemaining", "concurrencyLimit", "activeSessions",
    ):
        _nonnegative_int_field(value, field, context)
    if "minutesRemaining" in value:
        _nonnegative_int_field(value, "minutesRemaining", context)
    for field in ("reason", "planCode", "graceUntil"):
        if not isinstance(value[field], (str, type(None))):
            raise ValueError(f"{context} {field} must be a string or null")
    if value["revoked"] and not value["reason"]:
        raise ValueError(f"{context} revoked=true requires reason")
    return copy.deepcopy(value)


def _clean_credits(value: Any, context: str, *, allow_empty: bool = False) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError(f"{context} must be an object")
    if allow_empty and not value:
        return {}
    required = {
        "includedMinutesRemaining", "walletMinutes", "totalMinutesRemaining",
        "allowed", "reason", "planCode",
    }
    unknown = sorted(set(value) - required)
    missing = sorted(required - set(value))
    if unknown:
        raise ValueError(f"{context} has unknown fields: {', '.join(unknown)}")
    if missing:
        raise ValueError(f"{context} is missing fields: {', '.join(missing)}")
    included = _nonnegative_int_field(value, "includedMinutesRemaining", context)
    wallet = _nonnegative_int_field(value, "walletMinutes", context)
    total = _nonnegative_int_field(value, "totalMinutesRemaining", context)
    if total != included + wallet:
        raise ValueError(f"{context} totalMinutesRemaining does not equal included plus wallet")
    if type(value["allowed"]) is not bool:
        raise ValueError(f"{context} allowed must be a boolean")
    for field in ("reason", "planCode"):
        if not isinstance(value[field], (str, type(None))):
            raise ValueError(f"{context} {field} must be a string or null")
    return copy.deepcopy(value)


def _clean_activation(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValueError("activation must be an object")
    if not value:
        return {}
    required = {"activation_id", "poll_secret", "code", "connect_url", "expires_at", "created_at"}
    if set(value) != required:
        raise ValueError(f"activation must contain exactly: {', '.join(sorted(required))}")
    for field in required - {"created_at"}:
        if not isinstance(value[field], str) or not value[field].strip():
            raise ValueError(f"activation {field} must be a non-empty string")
    created_at = value["created_at"]
    if isinstance(created_at, bool) or not isinstance(created_at, (int, float)) or not math.isfinite(created_at):
        raise ValueError("activation created_at must be a finite number")
    return copy.deepcopy(value)


def _machine_name() -> str:
    name = os.environ.get("COMPUTERNAME", "").strip()
    if not name:
        raise RuntimeError("Windows computer name is empty")
    return name


def _default_data() -> dict[str, Any]:
    return {
        "version": CONFIG_VERSION,
        "setup_complete": False,
        "created_at": time.time(),
        "updated_at": time.time(),
        "access_code": "",
        "email_to": "",
        "permissions": dict(DEFAULT_PERMISSIONS),
        "alerts": _empty_alerts(),
        "cloud_api_base": DEFAULT_CLOUD_API_BASE,
        "workspace_id": "",
        "device_id": "",
        "device_name": _machine_name(),
        "device_token": "",
        "activation": {},
        "entitlement_cache": {},
        "autostart_enabled": True,
        "show_settings_in_taskbar": False,
        # User profile (fetched from cloud after sign-in)
        "user_name": "",
        "user_email": "",
        # Credits cache (fetched from cloud)
        "credits_cache": {},
    }


def _clean_permissions(raw: Any) -> dict[str, bool]:
    if not isinstance(raw, dict):
        raise ValueError("permissions must be an object")
    if set(raw) != set(PERMISSION_KEYS):
        raise ValueError(f"permissions must contain exactly: {', '.join(PERMISSION_KEYS)}")
    if any(type(raw[key]) is not bool for key in PERMISSION_KEYS):
        raise ValueError("every permission value must be a boolean")
    return {key: raw[key] for key in PERMISSION_KEYS}


def _clean_alerts(raw: Any) -> dict[str, dict[str, Any]]:
    if not isinstance(raw, dict) or set(raw) != set(ALERT_CODES):
        raise ValueError(f"alerts must contain exactly: {', '.join(ALERT_CODES)}")
    alerts = {}
    for code in ALERT_CODES:
        item = raw[code]
        if not isinstance(item, dict):
            raise ValueError(f"alert {code} must be an object")
        if set(item) != {"title", "message", "enabled"}:
            raise ValueError(f"alert {code} must contain title, message, and enabled")
        if not isinstance(item["title"], str) or not isinstance(item["message"], str):
            raise ValueError(f"alert {code} title and message must be strings")
        if type(item["enabled"]) is not bool:
            raise ValueError(f"alert {code} enabled must be a boolean")
        title = item["title"].strip()
        message = item["message"].strip()
        if item["enabled"] and not (title or message):
            raise ValueError(f"enabled alert {code} must have a title or message")
        alerts[code] = {"title": title, "message": message, "enabled": item["enabled"]}
    return alerts


def normalize_settings(data: dict[str, Any] | None) -> dict[str, Any]:
    base = _default_data()
    if data is not None:
        if not isinstance(data, dict):
            raise ValueError("Zadoo settings must be a JSON object")
        expected = set(base)
        unknown = sorted(set(data) - expected)
        missing = sorted(expected - set(data))
        if unknown:
            raise ValueError(f"Unknown Zadoo settings: {', '.join(unknown)}")
        if missing:
            raise ValueError(f"Missing Zadoo settings: {', '.join(missing)}")
        if data.get("version") != CONFIG_VERSION:
            raise ValueError(f"Zadoo settings version must be {CONFIG_VERSION}")
        base.update(data)
    for key in ("setup_complete", "autostart_enabled", "show_settings_in_taskbar"):
        if type(base[key]) is not bool:
            raise ValueError(f"{key} must be a boolean")
    for key in (
        "access_code", "email_to", "cloud_api_base", "workspace_id", "device_id",
        "device_name", "device_token", "user_name", "user_email",
    ):
        if not isinstance(base[key], str):
            raise ValueError(f"{key} must be a string")
        base[key] = base[key].strip()
    base["cloud_api_base"] = _clean_http_origin(base["cloud_api_base"], "cloud_api_base")
    if len(base["device_name"]) > 80:
        raise ValueError("device_name must be 80 characters or fewer")
    if not base["device_name"]:
        raise ValueError("device_name is required")
    if base["setup_complete"] and not base["access_code"]:
        raise ValueError("configured settings require an access code")
    base["activation"] = _clean_activation(base["activation"])
    base["entitlement_cache"] = _clean_entitlement(
        base["entitlement_cache"], "entitlement_cache", allow_empty=True
    )
    base["credits_cache"] = _clean_credits(base["credits_cache"], "credits_cache", allow_empty=True)
    for key in ("created_at", "updated_at"):
        if (
            isinstance(base[key], bool)
            or not isinstance(base[key], (int, float))
            or not math.isfinite(base[key])
        ):
            raise ValueError(f"{key} must be a number")
    base["permissions"] = _clean_permissions(base["permissions"])
    base["alerts"] = _clean_alerts(base["alerts"])
    return base


class SettingsStore:
    def __init__(self, path: Path | None = None):
        self.path = path if path is not None else settings_path()
        self._data: dict[str, Any] | None = None
        self._mtime_ns: int | None = None
        # Re-entrant so a locked atomic_update can call load()/save() (also locked).
        self._lock = threading.RLock()

    @contextmanager
    def _write_lock(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        lock_path = self.path.with_suffix(self.path.suffix + ".lock")
        with lock_path.open("a+b") as lock_file:
            lock_file.seek(0, os.SEEK_END)
            if lock_file.tell() == 0:
                lock_file.write(b"\0")
                lock_file.flush()
            lock_file.seek(0)
            try:
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_LOCK, 1)
            except OSError as exc:
                raise RuntimeError(f"Unable to lock Zadoo settings at {lock_path}: {exc}") from exc
            try:
                yield
            finally:
                lock_file.seek(0)
                msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)

    def load(self, *, reload: bool = False) -> dict[str, Any]:
        with self._lock:
            try:
                mtime_ns = self.path.stat().st_mtime_ns
            except FileNotFoundError:
                mtime_ns = None
            if self._data is not None and not reload and mtime_ns == self._mtime_ns:
                return copy.deepcopy(self._data)
            if self.path.exists():
                try:
                    raw = json.loads(self.path.read_text(encoding="utf-8"))
                except (OSError, json.JSONDecodeError) as exc:
                    raise RuntimeError(f"Unable to load Zadoo settings from {self.path}: {exc}") from exc
            else:
                raw = None
            data = normalize_settings(raw)
            self._data = data
            self._mtime_ns = mtime_ns
            return copy.deepcopy(self._data)

    def _save_unlocked(self, data: dict[str, Any]) -> dict[str, Any]:
        clean = normalize_settings(data)
        clean["updated_at"] = time.time()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(clean, indent=2, sort_keys=True), encoding="utf-8")
        tmp.replace(self.path)
        self._data = clean
        self._mtime_ns = self.path.stat().st_mtime_ns
        return copy.deepcopy(clean)

    def save(self, data: dict[str, Any]) -> dict[str, Any]:
        with self._lock, self._write_lock():
            return self._save_unlocked(data)

    def atomic_update(self, mutator) -> dict[str, Any]:
        """Cross-process locked load -> mutate -> atomic replace."""
        with self._lock, self._write_lock():
            data = self.load(reload=True)
            mutator(data)
            return self._save_unlocked(data)

    def configured(self) -> bool:
        return self.load()["setup_complete"]

    def get_access_code(self) -> str:
        return _dpapi_decrypt(self.load()["access_code"])

    def verify_access_code(self, code: str) -> bool:
        if not isinstance(code, str):
            raise ValueError("access code must be a string")
        clean = code.strip()
        expected = self.get_access_code()
        return bool(clean and expected) and secrets.compare_digest(clean, expected)

    def get_device_token(self) -> str:
        return _dpapi_decrypt(self.load()["device_token"])

    def clear_account(self) -> None:
        def _clear(data):
            for key in ("device_token", "workspace_id", "device_id", "user_name", "user_email"):
                data[key] = ""
            for key in ("activation", "entitlement_cache", "credits_cache"):
                data[key] = {}
        self.atomic_update(_clear)

    def update_cloud_status(self, payload: dict[str, Any]) -> dict[str, Any]:
        if not isinstance(payload, dict):
            raise ValueError("cloud status payload must be an object")
        allowed = {"workspace_id", "device_id", "device_name", "device_token", "activation"}
        unknown = sorted(set(payload) - allowed)
        if unknown:
            raise ValueError(f"Unknown cloud status fields: {', '.join(unknown)}")
        values = {}
        for key in ("workspace_id", "device_id", "device_name"):
            if key in payload:
                if not isinstance(payload[key], str):
                    raise ValueError(f"{key} must be a string")
                values[key] = payload[key].strip()
        if "device_token" in payload:
            if not isinstance(payload["device_token"], str) or not payload["device_token"].strip():
                raise ValueError("device_token must be a non-empty string")
            values["device_token"] = _dpapi_encrypt(payload["device_token"].strip())
        if "activation" in payload:
            if not isinstance(payload["activation"], dict):
                raise ValueError("activation must be an object")
            values["activation"] = copy.deepcopy(payload["activation"])
        return self.atomic_update(lambda data: data.update(values))

    def owner_view(self) -> dict[str, Any]:
        data = self.load()
        return {
            "setup_complete": data["setup_complete"],
            "access_code": _dpapi_decrypt(data["access_code"]),
            "email_to": data["email_to"],
            "permissions": data["permissions"],
            "alerts": data["alerts"],
            "cloud": {
                "cloud_api_base": data["cloud_api_base"],
                "workspace_id": data["workspace_id"],
                "device_id": data["device_id"],
                "device_name": data["device_name"],
                "has_device_token": bool(_dpapi_decrypt(data["device_token"])),
                "activation": data["activation"],
                "entitlement_cache": data["entitlement_cache"],
                "autostart_enabled": data["autostart_enabled"],
                "show_settings_in_taskbar": data["show_settings_in_taskbar"],
            },
        }

    def apply_setup(self, payload: dict[str, Any], *, require_code: str | None = None) -> dict[str, Any]:
        if not isinstance(payload, dict):
            raise ValueError("setup payload must be an object")
        allowed = {
            "access_code", "email_to", "alerts", "permissions", "device_name",
            "autostart_enabled", "show_settings_in_taskbar",
        }
        unknown = sorted(set(payload) - allowed)
        if unknown:
            raise ValueError(f"Unknown setup fields: {', '.join(unknown)}")
        raw_code = payload.get("access_code", "")
        if not isinstance(raw_code, str):
            raise ValueError("access_code must be a string")
        new_code = raw_code.strip()
        encrypted_code = _dpapi_encrypt(_validate_access_code(new_code)) if new_code else ""
        values = {}
        if "email_to" in payload:
            if not isinstance(payload["email_to"], str):
                raise ValueError("email_to must be a string")
            values["email_to"] = payload["email_to"].strip()
        if "alerts" in payload:
            values["alerts"] = _clean_alerts(payload["alerts"])
        if "permissions" in payload:
            values["permissions"] = _clean_permissions(payload["permissions"])
        if "device_name" in payload:
            if not isinstance(payload["device_name"], str):
                raise ValueError("device_name must be a string")
            values["device_name"] = payload["device_name"].strip()
        for key in ("autostart_enabled", "show_settings_in_taskbar"):
            if key in payload:
                if type(payload[key]) is not bool:
                    raise ValueError(f"{key} must be a boolean")
                values[key] = payload[key]

        def _apply(data):
            if data["setup_complete"] and (
                not isinstance(require_code, str) or not self.verify_access_code(require_code)
            ):
                raise PermissionError("Invalid access code")
            if encrypted_code:
                data["access_code"] = encrypted_code
            elif not data["access_code"]:
                raise ValueError("access code is required")
            data.update(values)
            data["setup_complete"] = True

        return self.atomic_update(_apply)


_STORE: SettingsStore | None = None


def get_settings_store() -> SettingsStore:
    global _STORE
    if _STORE is None:
        _STORE = SettingsStore()
    return _STORE

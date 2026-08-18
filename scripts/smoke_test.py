"""Smoke checks for the Zadoo VNC runtime."""
from __future__ import annotations

import argparse
import asyncio
import base64
import gzip
import json
import os
import queue
import re
import secrets
import sys
import tempfile
import threading
import time
import tomllib
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from itertools import pairwise
from pathlib import Path
from types import SimpleNamespace

ROOT = Path(__file__).resolve().parents[1]
SOURCE_GLOBS = (
    "zadoo_vnc/**/*.py",
    "zadoo_vnc/templates/*.html",
    "scripts/*.py",
    "installer/*.iss",
    ".github/workflows/*.yml",
    "*.py",
    "*.md",
    "*.toml",
    "*.txt",
    ".env.example",
)
MOJIBAKE_RE = re.compile(r"[\u00c2\u00c3\u00e2\u00f0][^\x00-\x7f]+")
GITHUB_TOKEN_RE = re.compile(r"\bgh[pousr]_[A-Za-z0-9]{20,}\b")


def _legacy_auth_strings():
    return tuple(
        "".join(chr(part) for part in value)
        for value in (
            (84, 69, 82, 77, 73, 78, 65, 84, 79, 82),
            (65, 68, 86, 69, 78, 84, 85, 82, 69, 83),
            (73, 78, 78, 79, 86, 65, 84, 73, 79, 78),
            (67, 72, 65, 76, 76, 69, 78, 71, 69, 82),
        )
    )


FORBIDDEN_SOURCE_STRINGS = (
    "re_" + "GF7",
    "resend_" + "api_key_default",
    "email_" + "to_default",
    "iskssj07" + "@gmail.com",
    "Set-Clipboard" + " -Value @\"\"",
    "Popen(" + "'clip'",
    *_legacy_auth_strings(),
)


def fail(message: str) -> None:
    print(f"FAIL: {message}")
    raise SystemExit(1)


def ok(message: str) -> None:
    print(f"OK: {message}")


def set_cookie_headers(headers) -> list[str]:
    return list(headers.get_all("Set-Cookie") or [])


def cookie_header_from_set_cookie(headers) -> str:
    cookies = []
    for value in set_cookie_headers(headers):
        first = str(value).split(";", 1)[0].strip()
        if first:
            cookies.append(first)
    return "; ".join(cookies)


def assert_camera_payload(body: bytes) -> None:
    try:
        payload = json.loads(body.decode("utf-8"))
    except Exception as exc:
        fail(f"camera payload is not valid JSON: {exc}")
    if not isinstance(payload, dict) or payload.get("success") is not True:
        fail("camera payload missing success=true")
    devices = payload.get("devices")
    if not isinstance(devices, list):
        fail("camera payload devices is not a list")
    for index, device in enumerate(devices):
        if not isinstance(device, dict):
            fail(f"camera device {index} is not an object")
        for key in ("id", "name", "label", "device_path"):
            if not isinstance(device.get(key), str) or not device.get(key).strip():
                fail(f"camera device {index} missing non-empty {key!r}")


def assert_mic_payload(body: bytes) -> None:
    try:
        payload = json.loads(body.decode("utf-8"))
    except Exception as exc:
        fail(f"mic payload is not valid JSON: {exc}")
    if not isinstance(payload, dict) or payload.get("success") is not True:
        fail("mic payload missing success=true")
    devices = payload.get("devices")
    if not isinstance(devices, list):
        fail("mic payload devices is not a list")
    for index, device in enumerate(devices):
        if not isinstance(device, dict):
            fail(f"mic device {index} is not an object")
        for key in ("id", "name", "label"):
            if not isinstance(device.get(key), str) or not device.get(key).strip():
                fail(f"mic device {index} missing non-empty {key!r}")


def source_files():
    seen = set()
    for glob in SOURCE_GLOBS:
        for path in ROOT.glob(glob):
            if path.is_file() and path not in seen:
                seen.add(path)
                yield path


def assert_imports() -> None:
    sys.path.insert(0, str(ROOT))
    test_auth_code = "SMK" + secrets.token_hex(3).upper()
    temp_settings = tempfile.TemporaryDirectory(prefix="zadoo-smoke-")
    old_env = {
        key: os.environ.get(key)
        for key in (
            "ZADOO_ACCESS_CODE",
            "ZADOO_SETTINGS_PATH",
            "ZADOO_SETTINGS_DIR",
            "ZADOO_ALLOWED_ORIGINS",
            "ZADOO_AUTH_MAX_FAILURES",
            "ZADOO_AUTH_WINDOW_SECONDS",
            "ZADOO_AUTH_LOCKOUT_SECONDS",
            "ZADOO_CLIPBOARD_IMAGE_MAX_BYTES",
            "ZADOO_CLIPBOARD_TEXT_MAX_BYTES",
            "ZADOO_CLOUD_HEARTBEAT_GRACE_SECONDS",
            "ZADOO_MAX_VIEWERS",
            "ZADOO_STREAM_START_PROFILE",
        )
    }
    os.environ["ZADOO_ACCESS_CODE"] = test_auth_code
    os.environ["ZADOO_ALLOW_DIRECT_ACCESS"] = "1"  # tests drive the viewer locally (no tunnel)
    os.environ["ZADOO_SETTINGS_PATH"] = str(Path(temp_settings.name) / "config.json")
    os.environ["ZADOO_AUTH_MAX_FAILURES"] = "3"
    os.environ["ZADOO_AUTH_WINDOW_SECONDS"] = "60"
    os.environ["ZADOO_AUTH_LOCKOUT_SECONDS"] = "120"
    os.environ["ZADOO_CLIPBOARD_IMAGE_MAX_BYTES"] = "8"
    os.environ["ZADOO_CLIPBOARD_TEXT_MAX_BYTES"] = "8"
    os.environ["ZADOO_CLOUD_HEARTBEAT_GRACE_SECONDS"] = "120"
    os.environ["ZADOO_MAX_VIEWERS"] = "5"
    os.environ["ZADOO_STREAM_START_PROFILE"] = "half-120-q54"
    for key in (
        "ZADOO_ALLOWED_ORIGINS",
        "ZADOO_SETTINGS_DIR",
    ):
        os.environ.pop(key, None)
    from websockets.datastructures import Headers

    import zadoo_vnc.settings as settings_mod
    from zadoo_vnc import __version__, camera_discovery, screen_capture
    from zadoo_vnc.app import require_windows_x64
    from zadoo_vnc.assets import load_binary, load_static, load_template
    from zadoo_vnc.config import env_bool
    from zadoo_vnc.diagnostics import build_diagnostic_bundle, redact_diagnostic_text
    from zadoo_vnc.logging_utils import _BoundedLogFile, _prune_logs
    from zadoo_vnc.saas import ZadooCloudClient
    from zadoo_vnc.server import VNCServer
    from zadoo_vnc.streaming import STREAM_LADDER

    require_windows_x64()
    project_metadata = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    if project_metadata["project"].get("dynamic") != ["version"]:
        fail("pyproject.toml must declare the package version as dynamic")
    if project_metadata["tool"]["setuptools"]["dynamic"].get("version") != {"attr": "zadoo_vnc.__version__"}:
        fail("setuptools version metadata does not use zadoo_vnc.__version__")
    if __version__ != settings_mod.APP_VERSION:
        fail(f"cloud agent version {settings_mod.APP_VERSION!r} does not match package version {__version__!r}")
    env_example = (ROOT / ".env.example").read_text(encoding="utf-8")
    for lock_name, required_packages in {
        "requirements-runtime.lock": ("pillow==12.3.0", "python-dotenv==1.2.2", "websockets==15.0.1"),
        "requirements-build.lock": ("pip==26.1.2", "setuptools==83.0.0", "pyinstaller==6.21.0"),
        "ci_requirements.lock": ("pip==26.1.2", "pip-audit==2.10.1"),
    }.items():
        lock_text = (ROOT / lock_name).read_text(encoding="utf-8").lower()
        if "--hash=sha256:" not in lock_text:
            fail(f"{lock_name} does not enforce package hashes")
        for package in required_packages:
            if package not in lock_text:
                fail(f"{lock_name} is missing {package}")
    build_script = (ROOT / "scripts" / "build_windows.ps1").read_text(encoding="utf-8")
    if build_script.count("--require-hashes") < 2:
        fail("Windows build does not enforce both dependency lock files")
    for required_build_marker in (
        "InnoSetupInstallerSha256",
        '$script:InnoPath = $installedCompiler',
        "release-manifest.json",
        "SHA256SUMS.txt",
        "Write-ReleaseMetadata",
    ):
        if required_build_marker not in build_script:
            fail(f"Windows build is missing release safeguard {required_build_marker}")
    local_inno = '$env:LOCALAPPDATA\\Programs\\Inno Setup 7\\ISCC.exe'
    path_inno = '"ISCC.exe"'
    if build_script.index(local_inno) > build_script.index(path_inno, build_script.index("function Resolve-Inno")):
        fail("Windows build must prefer the pinned Inno installation over PATH shims")
    installer_script = (ROOT / "installer" / "zadoo.iss").read_text(encoding="utf-8")
    for required_uninstall_path in ("{localappdata}\\Zadoo", "{commonappdata}\\Zadoo"):
        if required_uninstall_path not in installer_script:
            fail(f"Windows uninstaller does not cover settings path {required_uninstall_path}")
    documented_env = {
        "EMAIL_TO",
        "HIDE_CONSOLE",
        "RESEND_API_KEY",
        "RESEND_FROM",
        "ZADOO_ACCESS_CODE",
        "ZADOO_ALLOWED_ORIGINS",
        "ZADOO_ALLOW_DIRECT_ACCESS",
        "ZADOO_AUTH_LOCKOUT_SECONDS",
        "ZADOO_AUTH_MAX_FAILURES",
        "ZADOO_AUTH_WINDOW_SECONDS",
        "ZADOO_CLIPBOARD_IMAGE_MAX_BYTES",
        "ZADOO_CLIPBOARD_TEXT_MAX_BYTES",
        "ZADOO_CLOUD_HEARTBEAT_GRACE_SECONDS",
        "ZADOO_CLOUD_API_BASE",
        "ZADOO_CLOUDFLARED_PATH",
        "ZADOO_CLOUDFLARED_PROTOCOL",
        "ZADOO_DISABLE_TUNNEL",
        "ZADOO_LOG_BACKUP_COUNT",
        "ZADOO_DIAGNOSTIC_MAX_LOG_BYTES",
        "ZADOO_DIAGNOSTIC_MAX_LOG_FILES",
        "ZADOO_LOG_LEVEL",
        "ZADOO_LOG_MAX_BYTES",
        "ZADOO_LOG_RETENTION_DAYS",
        "ZADOO_MAX_VIEWERS",
        "ZADOO_SETTINGS_DIR",
        "ZADOO_SETTINGS_PATH",
        "ZADOO_SIGN_PFX_PASSWORD",
        "ZADOO_STREAM_START_PROFILE",
        "ZADOO_TARGET_KBPS",
    }
    for name in documented_env:
        if re.search(rf"^#?\s*{name}=", env_example, flags=re.MULTILINE) is None:
            fail(f".env.example does not document {name}")
    with tempfile.TemporaryDirectory(prefix="zadoo-log-smoke-") as temp_log_dir:
        log_dir = Path(temp_log_dir)
        log_path = log_dir / "zadoo_20260714.log"
        bounded_log = _BoundedLogFile(log_path, max_bytes=16, backup_count=2)
        bounded_log.write("first-line\n")
        bounded_log.write("second-line\n")
        bounded_log.flush()
        bounded_log.close()
        if not log_path.is_file() or not Path(f"{log_path}.1").is_file():
            fail("bounded log did not rotate at the configured byte limit")
        old_log = log_dir / "zadoo_20000101.log"
        old_log.write_text("old", encoding="utf-8")
        os.utime(old_log, (1, 1))
        _prune_logs(log_dir, retention_days=1)
        if old_log.exists():
            fail("expired log was not pruned")
    settings_override = os.environ["ZADOO_SETTINGS_PATH"]
    os.environ["ZADOO_SETTINGS_PATH"] = ""
    try:
        settings_mod.settings_path()
    except RuntimeError as exc:
        if str(exc) != "ZADOO_SETTINGS_PATH must not be empty":
            fail(f"empty settings path returned the wrong error: {exc}")
    else:
        fail("empty settings path override did not fail")
    finally:
        os.environ["ZADOO_SETTINGS_PATH"] = settings_override
    original_program_data = os.environ.get("PROGRAMDATA")
    original_local_app_data = os.environ.get("LOCALAPPDATA")
    with tempfile.TemporaryDirectory(prefix="zadoo-migration-smoke-") as migration_root:
        migration_root = Path(migration_root)
        os.environ.pop("ZADOO_SETTINGS_PATH", None)
        os.environ["PROGRAMDATA"] = str(migration_root / "ProgramData")
        os.environ["LOCALAPPDATA"] = str(migration_root / "LocalAppData")
        legacy_path = migration_root / "ProgramData" / "Zadoo" / "config.json"
        legacy_path.parent.mkdir(parents=True)
        legacy_code = "MIGRATE10"
        legacy_data = settings_mod.normalize_settings(None)
        legacy_data["setup_complete"] = True
        legacy_data["access_code"] = "dpapi:" + settings_mod._b64(
            settings_mod.win32crypt.CryptProtectData(
                legacy_code.encode("utf-8"), "Zadoo", None, None, None, 0x4
            )
        )
        legacy_path.write_text(json.dumps(legacy_data), encoding="utf-8")
        destination = settings_mod.settings_path()
        settings_mod._migrate_legacy_settings(destination)
        migrated_store = settings_mod.SettingsStore(destination)
        if migrated_store.get_access_code() != legacy_code:
            fail("legacy ProgramData access code did not survive per-user migration")
        migrated_text = destination.read_text(encoding="utf-8")
        if "dpapi-user:" not in migrated_text or '"dpapi:' in migrated_text:
            fail("legacy ProgramData secrets were not re-encrypted with current-user DPAPI")
        if not legacy_path.exists():
            fail("legacy settings migration removed the rollback copy")
        descriptor = settings_mod.win32security.GetNamedSecurityInfo(
            str(destination),
            settings_mod.win32security.SE_FILE_OBJECT,
            settings_mod.win32security.DACL_SECURITY_INFORMATION,
        )
        dacl = descriptor.GetSecurityDescriptorDacl()
        if dacl is None or dacl.GetAceCount() != 3:
            fail("migrated settings file does not have the expected protected three-principal ACL")
    os.environ["ZADOO_SETTINGS_PATH"] = settings_override
    if original_program_data is None:
        os.environ.pop("PROGRAMDATA", None)
    else:
        os.environ["PROGRAMDATA"] = original_program_data
    if original_local_app_data is None:
        os.environ.pop("LOCALAPPDATA", None)
    else:
        os.environ["LOCALAPPDATA"] = original_local_app_data
    os.environ["ZADOO_SMOKE_EMPTY_BOOL"] = ""
    try:
        env_bool("ZADOO_SMOKE_EMPTY_BOOL")
    except ValueError as exc:
        if str(exc) != "ZADOO_SMOKE_EMPTY_BOOL must be a boolean; got ''":
            fail(f"empty boolean returned the wrong error: {exc}")
    else:
        fail("empty boolean configuration did not fail")
    finally:
        os.environ.pop("ZADOO_SMOKE_EMPTY_BOOL")
    os.environ["ZADOO_AUTH_MAX_FAILURES"] = "bad"
    try:
        VNCServer(6173)
    except ValueError as exc:
        if str(exc) != "ZADOO_AUTH_MAX_FAILURES must be an integer; got 'bad'":
            fail(f"invalid authentication limit returned the wrong error: {exc}")
    else:
        fail("invalid authentication limit did not fail during server construction")
    finally:
        os.environ["ZADOO_AUTH_MAX_FAILURES"] = "3"
    os.environ["ZADOO_ALLOWED_ORIGINS"] = "/relative"
    try:
        VNCServer(6173)
    except ValueError as exc:
        if not str(exc).startswith("Invalid ZADOO_ALLOWED_ORIGINS: ZADOO_ALLOWED_ORIGINS entry must be"):
            fail(f"invalid allowed origin returned the wrong error: {exc}")
    else:
        fail("invalid allowed origin did not fail during server construction")
    finally:
        os.environ.pop("ZADOO_ALLOWED_ORIGINS")
    settings_mod._STORE = None

    for name in ("index.html", "terminal.html", "host_controls.html", "benchmark.html"):
        if not load_template(name).strip():
            fail(f"template did not load: {name}")

    for asset_name in ("brand-header.png", "splash.png", "trigger-icon.png"):
        if not load_binary(asset_name):
            fail(f"asset is empty: {asset_name}")

    for asset_name, prefix in (
        ("vendor/codemirror-5.65.21.min.css", b"/*"),
        ("vendor/codemirror-5.65.21.min.js", b"/**"),
        ("vendor/xterm-5.3.0.min.css", b"/**"),
        ("vendor/xterm-5.3.0-fit-0.8.0.min.js", b"/**"),
    ):
        if not load_static(asset_name).startswith(prefix):
            fail(f"static asset is missing or invalid: {asset_name}")

    isolated_store = settings_mod.get_settings_store()
    if isolated_store.configured():
        fail("smoke settings store should start unconfigured")
    try:
        isolated_store.apply_setup({"unknown": True})
    except ValueError as exc:
        if str(exc) != "Unknown setup fields: unknown":
            fail(f"unknown setup field returned the wrong error: {exc}")
    else:
        fail("unknown setup field did not fail")
    try:
        isolated_store.update_cloud_status({"unknown": True})
    except ValueError as exc:
        if str(exc) != "Unknown cloud status fields: unknown":
            fail(f"unknown cloud field returned the wrong error: {exc}")
    else:
        fail("unknown cloud status field did not fail")
    os.environ.pop("ZADOO_ACCESS_CODE")
    try:
        VNCServer(6173)._announce_auth_codes()
        fail("unconfigured server started without ZADOO_ACCESS_CODE")
    except RuntimeError as exc:
        if str(exc) != "ZADOO_ACCESS_CODE is required when installed Settings are not configured":
            fail(f"missing access code returned wrong error: {exc}")
    finally:
        os.environ["ZADOO_ACCESS_CODE"] = test_auth_code
    if any(item.get("enabled") for item in isolated_store.owner_view().get("alerts", {}).values()):
        fail("settings store should not create default alert text")
    try:
        isolated_store.apply_setup({"access_code": "TOO-LONG-CODE"})
        fail("access-code storage accepted a code longer than 10 characters")
    except ValueError:
        pass
    isolated_store.apply_setup({"access_code": "SMOKE"})
    external_store = settings_mod.SettingsStore(isolated_store.path)
    external_data = external_store.load(reload=True)
    external_data["setup_complete"] = True
    external_data["permissions"] = dict.fromkeys(settings_mod.PERMISSION_KEYS, True)
    time.sleep(0.02)
    external_store.save(external_data)
    if not all(isolated_store.load()["permissions"].values()):
        fail("settings store did not reload changes saved by another process")
    isolated_copy = isolated_store.load()
    isolated_copy["permissions"]["mouse"] = False
    if isolated_store.load()["permissions"]["mouse"] is not True:
        fail("settings load returned mutable cached state")

    first_store = settings_mod.SettingsStore(isolated_store.path)
    second_store = settings_mod.SettingsStore(isolated_store.path)
    first_entered = threading.Event()
    update_errors = []

    def first_update():
        try:
            def mutate(data):
                data["user_name"] = "First Writer"
                first_entered.set()
                time.sleep(0.2)
            first_store.atomic_update(mutate)
        except Exception as exc:
            update_errors.append(str(exc))

    def second_update():
        first_entered.wait(timeout=2)
        try:
            second_store.atomic_update(lambda data: data.__setitem__("user_email", "second@example.com"))
        except Exception as exc:
            update_errors.append(str(exc))

    threads = [threading.Thread(target=first_update), threading.Thread(target=second_update)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(timeout=5)
    if any(thread.is_alive() for thread in threads) or update_errors:
        fail(f"cross-store settings transaction failed: {update_errors}")
    concurrent_data = isolated_store.load(reload=True)
    if concurrent_data["user_name"] != "First Writer" or concurrent_data["user_email"] != "second@example.com":
        fail(f"cross-store settings update was lost: {concurrent_data}")

    before_failed_update = isolated_store.load(reload=True)
    try:
        isolated_store.atomic_update(lambda data: data.__setitem__("device_name", "x" * 81))
    except ValueError as exc:
        if str(exc) != "device_name must be 80 characters or fewer":
            fail(f"oversized device name returned the wrong error: {exc}")
    else:
        fail("oversized device name did not fail")
    if isolated_store.load(reload=True) != before_failed_update:
        fail("failed settings transaction changed stored state")

    malformed_settings = settings_mod.normalize_settings(None)
    malformed_settings["activation"] = {"activation_id": "only"}
    try:
        settings_mod.normalize_settings(malformed_settings)
    except ValueError as exc:
        expected = "activation must contain exactly: activation_id, code, connect_url, created_at, expires_at, poll_secret"
        if str(exc) != expected:
            fail(f"malformed activation returned the wrong error: {exc}")
    else:
        fail("malformed activation did not fail")

    incomplete_settings = settings_mod.normalize_settings(None)
    del incomplete_settings["email_to"]
    try:
        settings_mod.normalize_settings(incomplete_settings)
    except ValueError as exc:
        if str(exc) != "Missing Zadoo settings: email_to":
            fail(f"incomplete settings returned the wrong error: {exc}")
    else:
        fail("incomplete settings did not fail")

    malformed_settings = settings_mod.normalize_settings(None)
    malformed_settings["entitlement_cache"] = {"allowed": True}
    try:
        settings_mod.normalize_settings(malformed_settings)
    except ValueError as exc:
        expected = (
            "entitlement_cache is missing fields: activeSessions, concurrencyLimit, graceUntil, "
            "includedMinutesRemaining, planCode, reason, revoked, walletMinutesRemaining"
        )
        if str(exc) != expected:
            fail(f"malformed entitlement cache returned the wrong error: {exc}")
    else:
        fail("malformed entitlement cache did not fail")

    malformed_settings = settings_mod.normalize_settings(None)
    malformed_settings["credits_cache"] = {
        "includedMinutesRemaining": 5,
        "walletMinutes": 7,
        "totalMinutesRemaining": 13,
        "allowed": True,
        "reason": None,
        "planCode": None,
    }
    try:
        settings_mod.normalize_settings(malformed_settings)
    except ValueError as exc:
        if str(exc) != "credits_cache totalMinutesRemaining does not equal included plus wallet":
            fail(f"malformed credits cache returned the wrong error: {exc}")
    else:
        fail("malformed credits cache did not fail")

    isolated_store.save(settings_mod.normalize_settings(None))

    class FakeCloudStore:
        def __init__(self):
            self.data = {"entitlement_cache": {"sentinel": True}}
            self.update_count = 0

        @staticmethod
        def get_device_token():
            return "smoke-token"

        def atomic_update(self, update):
            self.update_count += 1
            update(self.data)

    fake_cloud_store = FakeCloudStore()
    cloud = ZadooCloudClient(fake_cloud_store)
    cloud._request = lambda *_args, **_kwargs: {
        "success": False,
        "entitlement": {"allowed": True, "revoked": False},
    }
    cloud.start_session()
    cloud.session_heartbeat("smoke-session")
    if fake_cloud_store.update_count or fake_cloud_store.data["entitlement_cache"] != {"sentinel": True}:
        fail("failed cloud session response overwrote the entitlement cache")
    invalid_utf8 = cloud._response_object(b"\xff", "/smoke", 502)
    if invalid_utf8.get("error", "").split(":", 1)[0] != "Cloud API returned invalid UTF-8 for /smoke (HTTP 502)":
        fail(f"cloud invalid UTF-8 returned the wrong error: {invalid_utf8}")
    missing_success = cloud._validated_result({"error": "upstream"}, "/smoke", 400, http_error=True)
    if missing_success != {
        "success": False,
        "error": "Cloud API response for /smoke is missing boolean success (HTTP 400)",
    }:
        fail(f"cloud response without success returned the wrong error: {missing_success}")
    contradictory = cloud._validated_result({"success": True}, "/smoke", 500, http_error=True)
    if contradictory != {
        "success": False,
        "error": "Cloud API returned success=true for /smoke with HTTP 500",
    }:
        fail(f"contradictory cloud HTTP response returned the wrong error: {contradictory}")
    expired = cloud._validated_result(
        {"success": False, "status": "expired"},
        "/api/agent/activate/poll",
        410,
        http_error=True,
    )
    if expired != {"success": False, "status": "expired", "error": "Activation code expired"}:
        fail(f"expired activation returned the wrong error: {expired}")
    try:
        cloud._required_credits({
            "credits": {
                "includedMinutesRemaining": 5,
                "walletMinutes": 7,
                "totalMinutesRemaining": 13,
                "allowed": True,
                "reason": None,
                "planCode": None,
            }
        }, "Credits")
    except RuntimeError as exc:
        if str(exc) != "Credits credits totalMinutesRemaining does not equal included plus wallet":
            fail(f"invalid credits total returned the wrong error: {exc}")
    else:
        fail("invalid credits total did not fail")

    session_server = VNCServer(6173)
    session_server.settings_store = SimpleNamespace(get_device_token=lambda: "smoke-token")
    import zadoo_vnc.saas as saas_mod

    original_start_session = saas_mod.ZadooCloudClient.start_session
    entitlement = {
        "allowed": False,
        "revoked": False,
        "reason": "No active plan or wallet balance",
        "planCode": None,
        "includedMinutesRemaining": 0,
        "walletMinutesRemaining": 0,
        "concurrencyLimit": 1,
        "activeSessions": 0,
        "graceUntil": None,
    }
    try:
        saas_mod.ZadooCloudClient.start_session = lambda _client, _url: {
            "success": False,
            "status_code": 402,
            "error": entitlement["reason"],
            "entitlement": entitlement,
        }
        if session_server._cloud_start_remote_session() != {"success": True, "no_credits": True}:
            fail("zero-credit session response did not enter grace")

        entitlement = dict(
            entitlement,
            reason="Concurrency limit reached",
            planCode="PAYG",
            walletMinutesRemaining=10,
            activeSessions=1,
        )
        blocked = session_server._cloud_start_remote_session()
        if blocked["success"] is not False or blocked["error"] != "Concurrency limit reached":
            fail(f"concurrency block was incorrectly converted to grace: {blocked}")
    finally:
        saas_mod.ZadooCloudClient.start_session = original_start_session

    loads = [profile.target_fps / (profile.scale_div ** 2) for profile in STREAM_LADDER]
    if any(current < following for current, following in pairwise(loads)):
        fail(f"adaptive stream ladder is not monotonic: {loads}")

    raw_camera_path = r"\\?\usb#vid_0000&pid_0000#smoke"
    cameras = camera_discovery._normalize_camera_devices([
        {"name": "Smoke Camera", "device_path": raw_camera_path},
    ])
    if len(cameras) != 1 or camera_discovery.camera_open_target(cameras[0]) != f"@device_pnp_{raw_camera_path}":
        fail(f"camera DirectShow path normalization failed: {cameras}")
    try:
        camera_discovery._normalize_camera_devices([
            {"name": "No Path", "device_path": ""},
        ])
    except ValueError as exc:
        if str(exc) != "Camera 0 has no valid DirectShow device path":
            fail(f"missing camera path returned the wrong error: {exc}")
    else:
        fail("camera without a DirectShow device path did not fail")

    try:
        server = VNCServer(6173)
        if server.port != 6173:
            fail("VNCServer did not instantiate with the fixed port")
        if not getattr(server, "adaptive_stream", None) or server.adaptive_stream.profile.name != "half-120-q54":
            fail("adaptive stream did not select the requested half-120-q54 startup profile")
        if server.current_fps != 120 or server.current_quality != 85:
            fail(f"startup profile did not apply fps/locked quality: fps={server.current_fps} quality={server.current_quality}")
        if server._quality_locked_by_user:
            fail("startup quality lock should be OFF so the adaptive resolution ladder can run on slow links")
        initial_profile = server.adaptive_stream.profile_index
        changed = server.adaptive_stream.observe_server(
            frame_bytes=20_000,
            max_write_buffer=0,
            skipped_total=0,
            inflight_sends=0,
            video_clients=1,
            frame_age_ms=0,
            capture_stats={
                "quality": server.current_quality,
                "last_capture_ms": 0.5,
                "last_encode_ms": 2.0,
                "current_fps": 0.1,
                "target_fps": server.current_fps,
                "perf_scale_div": 2,
            },
        )
        if changed or server.adaptive_stream.profile_index != initial_profile:
            fail("static desktop changed the adaptive stream profile")
        server.adaptive_stream.last_eval_at = 0
        server._apply_quality(70)
        if not server._quality_locked_by_user:
            fail("changing quality should lock it so the user's manual choice is honored")
        try:
            server._apply_quality(101)
        except ValueError as exc:
            if str(exc) != "JPEG quality must be between 10 and 100":
                fail(f"out-of-range quality returned the wrong error: {exc}")
        else:
            fail("out-of-range quality did not fail")
        try:
            server._apply_quality(True)
        except ValueError as exc:
            if str(exc) != "JPEG quality must be an integer":
                fail(f"boolean quality returned the wrong error: {exc}")
        else:
            fail("boolean quality did not fail")
        invalid_client_stats = {
            "display_fps": float("nan"),
            "decode_ms": 0,
            "draw_ms": 0,
            "recv_kbps": 0,
            "dropped_blobs": 0,
            "rtt_ms": 0,
            "receive_delay_ms": 0,
        }
        try:
            server.adaptive_stream.record_client_stats(invalid_client_stats)
        except ValueError as exc:
            if str(exc) != "Client stream statistic display_fps must be finite and non-negative":
                fail(f"NaN stream metric returned the wrong error: {exc}")
        else:
            fail("NaN stream metric did not fail")
        region_capturer = screen_capture.ScreenCapturer()
        bgra_red = screen_capture.np.zeros((32, 32, 4), dtype=screen_capture.np.uint8)
        bgra_red[:, :, :] = (0, 0, 255, 255)
        from imagecodecs._jpeg8 import jpeg8_decode

        decoded_red = jpeg8_decode(region_capturer._encode_frame(bgra_red))
        if decoded_red.shape != (32, 32, 3) or decoded_red[16, 16, 0] < 240 or decoded_red[16, 16, 2] > 15:
            fail(f"native BGRA screen encoding changed channel order: {decoded_red[16, 16].tolist()}")
        region_capturer.is_running = True
        region_capturer.active_capture_method = "bettercam"
        region_capturer.capture_stats["sequence"] = 1
        if not region_capturer.get_capture_stats()["is_working"]:
            fail("static desktop with a valid frame was reported as not working")
        region_capturer.configure_performance(True, "center_0.5", 1, False)
        if region_capturer.get_active_region_norm() != (0.25, 0.25, 0.75, 0.75):
            fail(f"center performance region changed: {region_capturer.get_active_region_norm()}")
        region_capturer.configure_performance(
            True,
            "custom",
            1,
            False,
            {"x0": 0.2, "y0": 0.1, "x1": 0.6, "y1": 0.5},
        )
        previous_performance = (
            region_capturer.perf_enabled,
            region_capturer.perf_region,
            region_capturer.perf_scale_div,
            region_capturer.perf_grayscale,
            dict(region_capturer._custom_rect_norm),
        )
        try:
            region_capturer.configure_performance(
                True,
                "custom",
                1,
                False,
                {"x0": 0.8, "y0": 0.1, "x1": 0.2, "y1": 0.5},
            )
        except ValueError as exc:
            if str(exc) != "Custom performance region must have positive width and height":
                fail(f"invalid custom region returned the wrong error: {exc}")
        else:
            fail("invalid custom region did not fail")
        current_performance = (
            region_capturer.perf_enabled,
            region_capturer.perf_region,
            region_capturer.perf_scale_div,
            region_capturer.perf_grayscale,
            dict(region_capturer._custom_rect_norm),
        )
        if current_performance != previous_performance:
            fail("invalid performance configuration partially changed capture state")
        server.screen_capturer = region_capturer
        mapped = server._map_view_norm_to_screen_norm(0.5, 0.5)
        if tuple(round(v, 4) for v in mapped) != (0.4, 0.3):
            fail(f"custom ROI mouse mapping failed: {mapped}")
        mapped_corner = server._map_view_norm_to_screen_norm(1.0, 0.0)
        if tuple(round(v, 4) for v in mapped_corner) != (0.6, 0.1):
            fail(f"custom ROI corner mapping failed: {mapped_corner}")
        try:
            server._map_view_norm_to_screen_norm(1.01, 0.5)
        except ValueError as exc:
            if str(exc) != "x must be a number from 0 to 1":
                fail(f"invalid input coordinate returned the wrong error: {exc}")
        else:
            fail("out-of-range input coordinate did not fail")
        server._set_manual_performance({
            "enabled": True,
            "region": "custom",
            "scale_div": 1,
            "grayscale": True,
            "rect_norm": {"x0": 0.2, "y0": 0.1, "x1": 0.6, "y1": 0.5},
        })
        if region_capturer.perf_region != "custom" or region_capturer.perf_scale_div != 2:
            fail("adaptive profile overwrote or weakened the manual performance region")
        if not region_capturer.perf_grayscale:
            fail("adaptive profile removed manual grayscale")
        server._set_manual_performance({
            "enabled": False,
            "region": "full",
            "scale_div": 1,
            "grayscale": False,
        })
        if server._map_view_norm_to_screen_norm(0.5, 0.5) != (0.5, 0.5):
            fail("disabled performance mode should not remap mouse coordinates")
        server.screen_capturer = None

        class FakeBetterCam:
            def __init__(self):
                self.grab_calls = 0
                self.stop_calls = 0
                self._duplicator = SimpleNamespace(texture=object(), duplicator=object())
                self._stagesurf = SimpleNamespace(texture=object(), width=8, height=8)

            width = 8
            height = 8

            def grab(self, region=None):
                if region is not None:
                    fail(f"default BetterCam capture unexpectedly used region {region}")
                self.grab_calls += 1
                return screen_capture.np.zeros((8, 8, 3), dtype=screen_capture.np.uint8)

            def stop(self):
                self.stop_calls += 1

        fake_bettercam = FakeBetterCam()
        capturer = screen_capture.ScreenCapturer(fps=60, quality=60)
        capturer.bettercam_camera = fake_bettercam
        capturer._initial_frame_pending = False
        capturer.set_streaming_active(False)
        capturer.set_streaming_active(True)
        if not capturer._initial_frame_pending:
            fail("resuming screen streaming did not request a fresh initial frame")
        capturer._initial_frame_pending = False
        capturer._grab_screen_bettercam()
        capturer.fps = 120
        capturer._grab_screen_bettercam()
        if fake_bettercam.grab_calls != 2:
            fail("BetterCam direct capture did not poll for changed frames")
        capturer._release_bettercam()
        if fake_bettercam.stop_calls != 1:
            fail("BetterCam disposal did not stop the camera")
        if fake_bettercam._duplicator.texture is not None or fake_bettercam._duplicator.duplicator is not None:
            fail("BetterCam disposal retained duplicator resources")
        if fake_bettercam._stagesurf.texture is not None:
            fail("BetterCam disposal retained the staging surface")

        class FakePanic(BaseException):
            pass

        terminal_attempts = 0
        expected_proc = object()
        terminal_spawn_kwargs = {}
        original_module_path = os.environ.get("PSMODULEPATH")
        injected_module_path = str(ROOT / "launcher" / "PowerShell" / "Modules")
        preserved_module_path = str(ROOT / "Documents" / "WindowsPowerShell" / "Modules")
        test_module_path = os.pathsep.join((injected_module_path, preserved_module_path))

        def transient_conpty(*_args, **_kwargs):
            nonlocal terminal_attempts, terminal_spawn_kwargs
            terminal_attempts += 1
            terminal_spawn_kwargs = _kwargs
            if terminal_attempts == 1:
                raise FakePanic("called Result::unwrap() on HRESULT(0x800700BB)")
            return expected_proc

        media = sys.modules["zadoo_vnc.media"]
        original_pty_process = media.PtyProcess
        media.PtyProcess = SimpleNamespace(spawn=transient_conpty)
        os.environ["PSMODULEPATH"] = test_module_path
        try:
            spawned_proc = asyncio.run(media._spawn_conpty(["powershell.exe"], str(ROOT), (34, 120)))
            restored_module_path = os.environ.get("PSMODULEPATH")
        finally:
            media.PtyProcess = original_pty_process
            if original_module_path is None:
                os.environ.pop("PSMODULEPATH", None)
            else:
                os.environ["PSMODULEPATH"] = original_module_path
        if spawned_proc is not expected_proc or terminal_attempts != 2:
            fail("ConPTY transient startup race did not retry exactly once")
        terminal_env = terminal_spawn_kwargs.get("env", {})
        terminal_module_paths = terminal_env.get("PSMODULEPATH", "").split(os.pathsep)
        normalized_terminal_paths = [
            os.path.normcase(os.path.normpath(path)) for path in terminal_module_paths if path
        ]
        if not normalized_terminal_paths or any(
            f"{os.sep}windowspowershell{os.sep}" not in path
            for path in normalized_terminal_paths
        ):
            fail("ConPTY inherited a non-Windows-PowerShell module path")
        if os.path.normcase(os.path.normpath(injected_module_path)) in normalized_terminal_paths:
            fail("ConPTY inherited a launcher-specific PowerShell 7 module path")
        if os.path.normcase(os.path.normpath(preserved_module_path)) not in normalized_terminal_paths:
            fail("ConPTY removed the user's Windows PowerShell module path")
        if restored_module_path != test_module_path:
            fail("ConPTY did not restore the host PowerShell module path after spawning")

        async def route_checks():
            class FakeConnection:
                def __init__(self, host="127.0.0.1"):
                    self.remote_address = (host, 50000)

            class FakeHttpRequest:
                def __init__(self, path, headers, body=b""):
                    self.path = path
                    self.headers = headers
                    self.body = body

            async def request(target, path, headers, body=b"", remote_host="127.0.0.1"):
                return await target.process_request(
                    FakeConnection(remote_host),
                    FakeHttpRequest(path, headers, body),
                )

            loop = asyncio.get_running_loop()
            stopped = loop.create_future()
            stopped.set_result(True)
            media_done = loop.create_future()
            media_done.set_result(None)
            server._raise_for_completed_required_tasks(
                {media_done, stopped}, stopped, None
            )
            fatal = RuntimeError("heartbeat fail-closed smoke")
            try:
                server._raise_for_completed_required_tasks({stopped}, stopped, fatal)
                fail("fatal shutdown error was ignored")
            except RuntimeError as exc:
                if exc is not fatal:
                    fail(f"fatal shutdown returned the wrong error: {exc}")
            try:
                server._raise_for_completed_required_tasks({media_done}, stopped, None)
                fail("unexpected media task completion was ignored")
            except RuntimeError as exc:
                if str(exc) != "A required media broadcast task stopped unexpectedly":
                    fail(f"media task completion returned the wrong error: {exc}")

            headers = Headers()
            headers["Host"] = "localhost:6173"
            if server._runtime_admin_code_valid(headers):
                fail("local admin authorization accepted a missing access code")
            admin_headers = headers.copy()
            admin_headers["X-Zadoo-Code"] = test_auth_code
            if not server._runtime_admin_code_valid(admin_headers):
                fail("local admin authorization rejected the configured access code")
            spoofed_admin = await request(
                server,
                "/api/runtime/status",
                headers,
                remote_host="203.0.113.20",
            )
            if spoofed_admin.status_code != 403:
                fail(f"remote peer spoofed local admin access: {spoofed_admin.status_code}")
            tunnel_admin_headers = Headers()
            tunnel_admin_headers["Host"] = "device.example.com"
            tunnel_admin_headers["CF-Ray"] = "smoke"
            tunnel_admin = await request(server, "/api/runtime/status", tunnel_admin_headers)
            if tunnel_admin.status_code != 403:
                fail(f"tunnel reached local admin API: {tunnel_admin.status_code}")
            unauth_response = await request(server, "/benchmark.html", headers)
            if unauth_response.status_code != 403:
                fail(f"unauthenticated benchmark returned {unauth_response.status_code}, expected 403")
            static_response = await request(
                server,
                "/static/vendor/codemirror-5.65.21.min.js",
                headers,
            )
            if static_response.status_code != 200 or static_response.headers["Cache-Control"] != "public, max-age=31536000, immutable":
                fail("versioned static asset was not served with immutable caching")
            expected_security_headers = {
                "Content-Security-Policy": "base-uri 'none'; object-src 'none'; frame-ancestors 'self'",
                "Cross-Origin-Resource-Policy": "same-origin",
                "Referrer-Policy": "no-referrer",
                "X-Content-Type-Options": "nosniff",
                "X-Frame-Options": "SAMEORIGIN",
            }
            for name, expected in expected_security_headers.items():
                if static_response.headers.get(name) != expected:
                    fail(f"response security header {name} was missing or invalid")
            gzip_headers = headers.copy()
            gzip_headers["Accept-Encoding"] = "br, gzip;q=1"
            gzip_index = await request(server, "/", gzip_headers)
            if (
                gzip_index.headers.get("Content-Encoding") != "gzip"
                or gzip_index.headers.get("Vary") != "Accept-Encoding"
                or gzip.decompress(gzip_index.body).decode("utf-8") != load_template("index.html")
            ):
                fail("HTML gzip negotiation returned invalid content or headers")
            gzip_static = await request(
                server, "/static/vendor/codemirror-5.65.21.min.js", gzip_headers
            )
            if (
                gzip_static.headers.get("Content-Encoding") != "gzip"
                or gzip.decompress(gzip_static.body)
                != load_static("vendor/codemirror-5.65.21.min.js")
            ):
                fail("static gzip negotiation returned invalid content")
            no_gzip_headers = headers.copy()
            no_gzip_headers["Accept-Encoding"] = "gzip;q=0"
            no_gzip_index = await request(server, "/", no_gzip_headers)
            if no_gzip_index.headers.get("Content-Encoding") is not None:
                fail("gzip;q=0 unexpectedly returned compressed content")
            explicit_no_gzip_headers = headers.copy()
            explicit_no_gzip_headers["Accept-Encoding"] = "gzip;q=0, *;q=1"
            explicit_no_gzip_index = await request(server, "/", explicit_no_gzip_headers)
            if explicit_no_gzip_index.headers.get("Content-Encoding") is not None:
                fail("explicit gzip refusal was overridden by wildcard encoding")
            asset_response = await request(server, "/brand-header.png", headers)
            if asset_response.headers.get("Cache-Control") != "public, max-age=31536000, immutable":
                fail("content-versioned image was not served with immutable caching")
            index_template = load_template("index.html")
            for asset_version in ("b0e399b76691", "a8ff7f23c7aa", "7cc869410b1c"):
                if f"?v={asset_version}" not in index_template:
                    fail(f"index template is missing image content version {asset_version}")
            for mobile_marker in (
                'id="view-controls"',
                "remoteView.mode",
                "canvas.addEventListener('pointerdown'",
                "touchDistance(points)",
                'callRpc("billing.topup_order"',
                'callRpc("billing.topup_verify"',
            ):
                if mobile_marker not in index_template:
                    fail(f"mobile/RPC frontend is missing {mobile_marker}")
            if "/api/local/topup-order" in index_template or "/api/local/topup-verify" in index_template:
                fail("frontend still calls legacy state-changing payment HTTP routes")
            host_controls_template = load_template("host_controls.html")
            if "alert.trigger" not in host_controls_template or "/api/alert" in host_controls_template:
                fail("host controls did not migrate alerts to authenticated WebSocket RPC")
            snapshot_prefix_response = await request(server, "/snapshot-extra?fmt=png", headers)
            if snapshot_prefix_response.status_code != 404:
                fail(f"snapshot prefix route returned {snapshot_prefix_response.status_code}, expected 404")
            query_auth_response = await request(server, f"/api/auth?code={test_auth_code}", headers)
            if query_auth_response.status_code != 400:
                fail(f"query auth returned {query_auth_response.status_code}, expected 400")
            legacy_headers = Headers()
            legacy_headers["Host"] = "localhost:6173"
            legacy_headers["X-Zadoo-Code"] = _legacy_auth_strings()[0]
            legacy_response = await request(server, "/api/auth", legacy_headers)
            if legacy_response.status_code != 401:
                fail(f"legacy hardcoded auth code returned {legacy_response.status_code}, expected 401")
            cross_origin_headers = Headers()
            cross_origin_headers["Host"] = "localhost:6173"
            cross_origin_headers["Origin"] = "https://evil.example"
            cross_origin_response = await request(server, "/", cross_origin_headers)
            if cross_origin_response.status_code != 403:
                fail(f"cross-origin HTTP request returned {cross_origin_response.status_code}, expected 403")
            auth_request_headers = Headers()
            auth_request_headers["Host"] = "localhost:6173"
            auth_request_headers["Origin"] = "http://localhost:6173"
            auth_request_headers["X-Zadoo-Code"] = test_auth_code
            auth_response = await request(server, "/api/auth", auth_request_headers)
            if auth_response.status_code != 200:
                fail(f"auth route returned {auth_response.status_code}, expected 200")
            auth_payload = json.loads(auth_response.body.decode("utf-8"))
            csrf_token = auth_payload.get("csrf_token")
            cookie = cookie_header_from_set_cookie(auth_response.headers)
            if not cookie or "zadoo_auth=" not in cookie or not csrf_token:
                fail("auth route did not set zadoo_auth cookie")
            if auth_payload.get("limits", {}).get("clipboard_image_max_bytes") != 8:
                fail("auth route did not expose the configured clipboard image limit")
            if auth_payload.get("limits", {}).get("max_viewers") != 5:
                fail("auth route did not expose the five-viewer limit")
            minimum_image_frame = ((server._clipboard_image_max_bytes + 2) // 3) * 4 + 4096
            if server._websocket_max_size < minimum_image_frame:
                fail("websocket frame limit cannot carry the configured clipboard image limit")
            headers["Cookie"] = cookie
            offline_credits = await request(server, "/api/local/credits", headers)
            offline_payload = json.loads(offline_credits.body.decode("utf-8"))
            if offline_credits.status_code != 200 or offline_payload != {
                "success": False,
                "offline": True,
                "error": "Device not signed in",
            }:
                fail(f"offline credits returned unexpected response: {offline_credits.status_code} {offline_payload}")
            for path, expected_error in {
                "/snapshot?fmt=png&fmt=jpeg": "Duplicate snapshot parameters: fmt",
                "/snapshot?unknown=1": "Unsupported snapshot parameters: unknown",
                "/snapshot?x0=0&y0=0&x1=nan&y1=1": (
                    "Snapshot region coordinates must be finite values between 0 and 1"
                ),
                "/snapshot?x0=0.5&y0=0&x1=0.5&y1=1": (
                    "Snapshot region must have positive width and height"
                ),
                "/snapshot?max_w=0": "Snapshot max_w must be a positive integer",
            }.items():
                response = await request(server, path, headers)
                if response.status_code != 400 or response.body.decode("utf-8") != expected_error:
                    fail(f"snapshot validation returned {response.status_code}: {response.body!r}")
            query_access_headers = Headers()
            query_access_headers["Host"] = "localhost:6173"
            query_access_headers["X-Zadoo-Code"] = test_auth_code
            query_access_response = await request(server, "/api/auth?access=lockdown", query_access_headers)
            if query_access_response.status_code != 400:
                fail(f"query access auth returned {query_access_response.status_code}, expected 400")
            query_access_payload = json.loads(query_access_response.body.decode("utf-8"))
            if query_access_payload.get("error") != "Authentication access selectors are not supported":
                fail(f"query access auth returned wrong error: {query_access_payload}")
            ws_bad_origin_headers = Headers()
            ws_bad_origin_headers["Host"] = "localhost:6173"
            ws_bad_origin_headers["Origin"] = "https://evil.example"
            ws_bad_origin_headers["Connection"] = "Upgrade"
            ws_bad_origin_headers["Upgrade"] = "websocket"
            ws_bad_origin_headers["Cookie"] = headers["Cookie"]
            ws_bad_origin = await request(server, "/video", ws_bad_origin_headers)
            if ws_bad_origin.status_code != 403:
                fail(f"cross-origin websocket returned {ws_bad_origin.status_code}, expected 403")
            throttled_server = VNCServer(6174)
            bad_headers = Headers()
            bad_headers["Host"] = "localhost:6174"
            bad_headers["X-Forwarded-For"] = "203.0.113.10"
            bad_headers["X-Zadoo-Code"] = "BADCODE"
            for _ in range(3):
                await request(throttled_server, "/api/auth", bad_headers)
            locked_headers = Headers()
            locked_headers["Host"] = "localhost:6174"
            locked_headers["X-Forwarded-For"] = "203.0.113.10"
            locked_headers["X-Zadoo-Code"] = test_auth_code
            locked_response = await request(throttled_server, "/api/auth", locked_headers)
            if locked_response.status_code != 429:
                fail(f"auth lockout returned {locked_response.status_code}, expected 429")
            csrf_headers = Headers()
            csrf_headers["Host"] = "localhost:6173"
            csrf_headers["Cookie"] = headers["Cookie"]
            csrf_headers["X-Zadoo-CSRF"] = csrf_token
            checks = {
                "/api/list-cameras": 200,
                "/api/list-mics": 200,
            }
            for path, expected_status in checks.items():
                response = await request(server, path, headers)
                if response is None:
                    fail(f"route returned websocket pass-through unexpectedly: {path}")
                if response.status_code != expected_status:
                    fail(f"route {path} returned {response.status_code}, expected {expected_status}")
                if path == "/api/list-cameras":
                    assert_camera_payload(response.body)
                elif path == "/api/list-mics":
                    assert_mic_payload(response.body)
            for path in (
                "/api/public-url",
                "/api/refresh-tunnel",
                "/api/set-quality?value=80",
                "/api/set-fps?value=20",
            ):
                response = await request(server, path, csrf_headers)
                if response.status_code != 404:
                    fail(f"removed duplicate route {path} returned {response.status_code}, expected 404")
            rpc_required_error = "State-changing action requires authenticated WebSocket RPC"
            for legacy_mutation in (
                "/api/settings/reload",
                "/api/runtime/stop",
                "/api/runtime/refresh-tunnel",
                "/api/local/topup-order",
                "/api/local/topup-verify",
                "/api/alert?code=A",
            ):
                legacy_response = await request(server, legacy_mutation, csrf_headers)
                legacy_payload = json.loads(legacy_response.body.decode("utf-8"))
                if legacy_response.status_code != 405 or legacy_payload.get("error") != rpc_required_error:
                    fail(f"legacy mutation {legacy_mutation} did not fail with the exact RPC error")
            global_market = await request(server, "/api/local/payment-market", headers)
            global_payload = json.loads(global_market.body.decode("utf-8"))
            if global_payload["market"] != "GLOBAL" or global_payload["currency"] != "USD":
                fail(f"default payment market was not global USD: {global_payload}")
            india_headers = headers.copy()
            india_headers["CF-IPCountry"] = "IN"
            india_market = await request(server, "/api/local/payment-market", india_headers)
            india_payload = json.loads(india_market.body.decode("utf-8"))
            if india_payload["market"] != "INDIA" or india_payload["currency"] != "INR":
                fail(f"India payment market was not INR: {india_payload}")
            removed_clipboard_response = await request(server, "/api/set-clipboard-image", csrf_headers)
            if removed_clipboard_response.status_code != 404:
                fail(
                    "removed /api/set-clipboard-image returned "
                    f"{removed_clipboard_response.status_code}, expected 404"
                )
            previous_fps = server.current_fps
            if server._apply_quality(100) != 100 or server.current_quality != 100:
                fail(f"quality control set current_quality={server.current_quality}, expected 100")
            if server.current_fps != previous_fps:
                fail(f"quality control changed current_fps from {previous_fps} to {server.current_fps}")
            class DummyCapturer:
                def __init__(self):
                    self.fps = 0
                    self.quality = 100
                    self.perf_enabled = None
                    self.perf_region = None
                    self.perf_scale_div = None
                    self.perf_grayscale = None
                def configure_performance(self, enabled, region, scale_div, grayscale, _rect_norm=None):
                    self.perf_enabled = enabled
                    self.perf_region = region
                    self.perf_scale_div = scale_div
                    self.perf_grayscale = grayscale
                def set_streaming_active(self, _active):
                    pass
            dummy = DummyCapturer()
            server.screen_capturer = dummy
            if getattr(server, "adaptive_stream", None):
                server.adaptive_stream.profile_index = 6
                server.adaptive_stream.measured_fps_cap = None
            server._apply_stream_profile("smoke_quality_lock")
            if server.current_quality != 100 or dummy.quality != 100:
                fail("adaptive stream profile changed user-selected quality")
            if dummy.perf_scale_div != 1 or dummy.perf_enabled:
                fail(f"quality 100 did not force full-resolution scale: enabled={dummy.perf_enabled} scale={dummy.perf_scale_div}")
            if getattr(server, "adaptive_stream", None):
                server.adaptive_stream.profile_index = 0
                server.adaptive_stream.measured_fps_cap = None
                server._apply_stream_profile("smoke_fps_cap_start")
            changed = server.adaptive_stream.observe_server(
                frame_bytes=6_000_000,
                max_write_buffer=0,
                skipped_total=0,
                inflight_sends=0,
                video_clients=1,
                frame_age_ms=0,
                capture_stats={
                    "quality": 100,
                    "last_capture_ms": 12.0,
                    "last_encode_ms": 12.0,
                    "current_fps": 40.0,
                    "target_fps": 240,
                    "perf_scale_div": 1,
                },
            )
            server._apply_stream_profile("smoke_fps_cap")
            if not changed or not server.adaptive_stream.measured_fps_cap:
                fail("full-resolution overload did not create a measured FPS cap")
            if dummy.perf_scale_div != 1 or dummy.quality != 100:
                fail("measured FPS cap changed quality or full-resolution scale")
            if dummy.fps >= 120:
                fail(f"measured FPS cap did not reduce target fps: {dummy.fps}")
            server.screen_capturer = None
            try:
                server._capture_stats_payload()
            except RuntimeError as exc:
                if str(exc) != "Screen capturer is not configured":
                    fail(f"capture stats returned wrong missing-capture error: {exc}")
            else:
                fail("capture stats accepted a missing screen capturer")
            stream = server._stream_status_payload()
            if not stream["measured_fps_cap"] or stream["fps_cap_reason"] != "host_capture_encode":
                fail("stream status missing measured FPS cap or host downshift reason")
            benchmark_response = await request(server, "/benchmark.html", headers)
            if benchmark_response.status_code != 200 or b"Stream Benchmark" not in benchmark_response.body:
                fail("benchmark.html route did not return the benchmark page")

            class FakeWebSocketRequest:
                def __init__(self, request_headers):
                    self.headers = request_headers

            class FakeWebSocket:
                def __init__(self, request_headers, messages, remote_address=("smoke", 1)):
                    self.request = FakeWebSocketRequest(request_headers)
                    self.sent = []
                    self.messages = list(messages)
                    self.remote_address = remote_address
                async def send(self, data):
                    self.sent.append(data)
                async def close(self, code=1000, reason=""):
                    pass
                def __aiter__(self):
                    return self
                async def __anext__(self):
                    if self.messages:
                        return self.messages.pop(0)
                    raise StopAsyncIteration

            malformed_rpc = FakeWebSocket(headers, ["{}"])
            await server.rpc_handler(malformed_rpc)
            malformed_payload = json.loads(malformed_rpc.sent[-1])
            if malformed_payload != {
                "id": None,
                "success": False,
                "error": "RPC request must contain exactly id, action, and params",
            }:
                fail(f"malformed RPC returned the wrong exact error: {malformed_payload}")

            viewer_entries = []
            for index in range(5):
                viewer_headers = Headers()
                viewer_headers["Host"] = "localhost:6173"
                viewer_headers["Cookie"] = f"zadoo_auth=viewer-{index}"
                viewer_ws = FakeWebSocket(viewer_headers, [])
                token = server._register_viewer_connection(viewer_ws, "/video")
                viewer_entries.append((token, viewer_ws))
            same_viewer_headers = Headers()
            same_viewer_headers["Host"] = "localhost:6173"
            same_viewer_headers["Cookie"] = "zadoo_auth=viewer-0"
            same_viewer_ws = FakeWebSocket(same_viewer_headers, [])
            same_token = server._register_viewer_connection(same_viewer_ws, "/input")
            if len(server._viewer_connections_by_session) != 5:
                fail("multiple sockets from one authenticated viewer consumed extra viewer slots")
            sixth_headers = Headers()
            sixth_headers["Host"] = "localhost:6173"
            sixth_headers["Cookie"] = "zadoo_auth=viewer-5"
            sixth_ws = FakeWebSocket(sixth_headers, [])
            try:
                server._register_viewer_connection(sixth_ws, "/video")
                fail("sixth authenticated viewer exceeded the configured viewer limit")
            except RuntimeError as exc:
                if str(exc) != "Viewer limit reached (5)":
                    fail(f"sixth viewer returned the wrong exact error: {exc}")
            server._unregister_viewer_connection(same_token, same_viewer_ws)
            for token, viewer_ws in viewer_entries:
                server._unregister_viewer_connection(token, viewer_ws)
            if server._viewer_connections_by_session:
                fail("viewer slots were not released after all sockets closed")

            started, failures, delay, elapsed = server._heartbeat_retry_policy(
                None, 0, 100.0, "offline"
            )
            if (started, failures, delay, elapsed) != (100.0, 1, 5, 0.0):
                fail("heartbeat retry policy did not start with a five-second backoff")
            _, _, delay, elapsed = server._heartbeat_retry_policy(
                started, failures, 219.0, "offline"
            )
            if delay != 1.0 or elapsed != 119.0:
                fail("heartbeat retry policy did not clamp its final retry to the fail-closed deadline")
            try:
                server._heartbeat_retry_policy(started, failures, 220.0, "offline")
                fail("heartbeat retry policy did not fail closed after 120 seconds")
            except RuntimeError as exc:
                if str(exc) != "Cloud heartbeat failed for 120 seconds: offline":
                    fail(f"heartbeat fail-closed policy returned the wrong exact error: {exc}")

            second_auth_response = await request(server, "/api/auth", auth_request_headers)
            second_headers = Headers()
            second_headers["Host"] = "localhost:6173"
            second_headers["Cookie"] = cookie_header_from_set_cookie(second_auth_response.headers)
            locked_ws = FakeWebSocket(headers, [])
            locked_ws_peer = FakeWebSocket(headers, [])
            unlocked_ws = FakeWebSocket(second_headers, [])
            server._set_grace_lock(locked_ws, True)
            server._set_grace_lock(locked_ws_peer, True)
            if server._is_ws_action_authorized(locked_ws, "click"):
                fail("grace lock did not block its authenticated session")
            if not server._is_ws_action_authorized(locked_ws, "get_stream_status"):
                fail("grace lock blocked a view-only action")
            if not server._is_ws_action_authorized(unlocked_ws, "click"):
                fail("one viewer's grace lock blocked another authenticated session")
            server._set_grace_lock(locked_ws, False)
            if server._is_ws_action_authorized(locked_ws_peer, "click"):
                fail("unlocking one video connection cleared a peer grace lock")
            server._set_grace_lock(locked_ws_peer, False)

            recorded_actions = []
            original_process_event = server.process_event
            try:
                server.process_event = lambda event, _websocket=None: recorded_actions.append(event.get("action"))
                full_ws = FakeWebSocket(headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.input_event_handler(full_ws)
                if recorded_actions != ["click"]:
                    fail(f"full /input did not process allowed control action: {recorded_actions}")
                server.screen_capturer = dummy
                quality_ws = FakeWebSocket(headers, [json.dumps({"action": "set_quality", "value": 80})])
                await server.video_stream_handler(quality_ws)
                if server.current_quality != 80:
                    fail(f"video quality control set {server.current_quality}, expected 80")
            finally:
                server.process_event = original_process_event

            class FakeSendWebSocket:
                remote_address = ("smoke-send", 1)
                def __init__(self):
                    self.sent = []
                async def send(self, data):
                    self.sent.append(data)

            alert_response = await request(server, "/api/alert?code=A", csrf_headers)
            if alert_response.status_code != 405:
                fail(f"legacy alert route returned {alert_response.status_code}, expected 405")

            settings_code = "SMOKE" + secrets.token_hex(2).upper()
            no_permissions = dict.fromkeys(settings_mod.PERMISSION_KEYS, False)
            full_permissions = dict.fromkeys(settings_mod.PERMISSION_KEYS, True)
            setup_payload = {
                "access_code": settings_code,
                "email_to": "alerts@example.invalid",
                "permissions": no_permissions,
                "alerts": {
                    "A": {"enabled": True, "title": "Smoke Alert", "message": "Settings-backed alert"},
                    "B": {"enabled": False, "title": "", "message": ""},
                    "C": {"enabled": False, "title": "", "message": ""},
                    "D": {"enabled": False, "title": "", "message": ""},
                },
            }
            isolated_store.apply_setup(setup_payload)
            server._load_alert_presets_from_settings()
            server.auth_sessions = {}
            if not isolated_store.verify_access_code(settings_code):
                fail("settings access code did not verify after save")
            if isolated_store.verify_access_code("wrong"):
                fail("settings access code accepted a wrong code")
            owner_settings = isolated_store.owner_view()
            if owner_settings.get("access_code") != settings_code:
                fail("settings access code was not visible in the local settings view")
            try:
                isolated_store.apply_setup(
                    {**setup_payload, "access_code": ""},
                    require_code="wrong",
                )
                fail("settings edit accepted the wrong admin code")
            except PermissionError:
                pass

            def auth_with_settings_code():
                auth_headers = Headers()
                auth_headers["Host"] = "localhost:6173"
                auth_headers["X-Zadoo-Code"] = settings_code
                return auth_headers

            denied_auth_response = await request(server, "/api/auth", auth_with_settings_code())
            if denied_auth_response.status_code != 200:
                fail("settings code authentication failed")
            denied_payload = json.loads(denied_auth_response.body.decode("utf-8"))
            if any(denied_payload.get("permissions", {}).values()):
                fail(f"settings auth payload did not use single permission matrix: {denied_payload}")
            denied_ws_headers = Headers()
            denied_ws_headers["Host"] = "localhost:6173"
            denied_ws_headers["Cookie"] = cookie_header_from_set_cookie(denied_auth_response.headers)
            recorded_actions = []
            original_process_event = server.process_event
            try:
                server.process_event = lambda event, _websocket=None: recorded_actions.append(event.get("action"))
                denied_ws = FakeWebSocket(denied_ws_headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.input_event_handler(denied_ws)
                if recorded_actions:
                    fail(f"settings disabled permissions processed control actions: {recorded_actions}")
                if not any("Forbidden" in str(item) for item in denied_ws.sent):
                    fail("settings disabled permissions did not report forbidden control action")
            finally:
                server.process_event = original_process_event

            denied_terminal_response = await request(server, "/terminal.html", denied_ws_headers)
            if denied_terminal_response.status_code != 403:
                fail(f"disabled terminal route returned {denied_terminal_response.status_code}, expected 403")
            if server._is_ws_authorized("/terminal", denied_ws_headers):
                fail("disabled terminal websocket was authorized")

            isolated_store.apply_setup(
                {**setup_payload, "access_code": settings_code, "permissions": full_permissions},
                require_code=settings_code,
            )
            server._load_alert_presets_from_settings()
            server.auth_sessions = {}
            allowed_auth_response = await request(server, "/api/auth", auth_with_settings_code())
            allowed_payload = json.loads(allowed_auth_response.body.decode("utf-8"))
            if allowed_auth_response.status_code != 200 or not all(allowed_payload.get("permissions", {}).values()):
                fail(f"settings allowed auth payload was wrong: {allowed_payload}")
            allowed_cookie = cookie_header_from_set_cookie(allowed_auth_response.headers)
            allowed_csrf = allowed_payload.get("csrf_token")
            allowed_csrf_headers = Headers()
            allowed_csrf_headers["Host"] = "localhost:6173"
            allowed_csrf_headers["Cookie"] = allowed_cookie
            allowed_csrf_headers["X-Zadoo-CSRF"] = allowed_csrf

            allowed_ws_headers = Headers()
            allowed_ws_headers["Host"] = "localhost:6173"
            allowed_ws_headers["Cookie"] = allowed_cookie
            allowed_terminal_response = await request(server, "/terminal.html", allowed_ws_headers)
            if allowed_terminal_response.status_code != 200:
                fail(f"enabled terminal route returned {allowed_terminal_response.status_code}, expected 200")
            if not server._is_ws_authorized("/terminal", allowed_ws_headers):
                fail("enabled terminal websocket was not authorized")

            recorded_actions = []
            original_process_event = server.process_event
            try:
                server.process_event = lambda event, _websocket=None: recorded_actions.append(event.get("action"))
                allowed_ws = FakeWebSocket(allowed_ws_headers, [json.dumps({"action": "click", "x": 0.5, "y": 0.5})])
                await server.input_event_handler(allowed_ws)
                if recorded_actions != ["click"]:
                    fail(f"settings enabled permissions did not process allowed control action: {recorded_actions}")
            finally:
                server.process_event = original_process_event

            alert_video_ws = FakeSendWebSocket()
            alert_input_ws = FakeSendWebSocket()
            server.video_clients = {alert_video_ws}
            server.input_clients = {alert_input_ws}
            alert_rpc = FakeWebSocket(
                allowed_ws_headers,
                [json.dumps({"id": "alert-1", "action": "alert.trigger", "params": {"code": "A"}})],
            )
            await server.rpc_handler(alert_rpc)
            if not alert_rpc.sent:
                fail("settings alert RPC did not return a response")
            alert_rpc_payload = json.loads(alert_rpc.sent[-1])
            if alert_rpc_payload != {"id": "alert-1", "success": True, "result": {"ok": True}}:
                fail(f"settings alert RPC response was invalid: {alert_rpc_payload}")
            await asyncio.sleep(0.05)
            for name, ws in {"video": alert_video_ws, "input": alert_input_ws}.items():
                if not ws.sent:
                    fail(f"settings alert route did not send to {name} websocket")
                alert_payload = json.loads(ws.sent[-1])
                if alert_payload.get("type") != "controller_alert" or not alert_payload.get("id"):
                    fail(f"settings alert payload for {name} websocket was invalid: {alert_payload}")
            server.video_clients = set()
            server.input_clients = set()

            local_rpc_headers = Headers()
            local_rpc_headers["Host"] = "localhost:6173"
            local_rpc = FakeWebSocket(
                local_rpc_headers,
                [
                    json.dumps(
                        {
                            "id": "reload-1",
                            "action": "settings.reload",
                            "params": {"admin_code": settings_code},
                        }
                    )
                ],
                remote_address=("127.0.0.1", 6173),
            )
            await server.rpc_handler(local_rpc)
            local_rpc_payload = json.loads(local_rpc.sent[-1])
            if local_rpc_payload.get("success") is not True or not isinstance(
                local_rpc_payload.get("result", {}).get("settings"), dict
            ):
                fail(f"authenticated local settings RPC failed: {local_rpc_payload}")

            diagnostic_marker = "diagnostic-" + "redaction-value"
            redacted = redact_diagnostic_text(
                "Authorization: Bearer token-value; zadoo_auth=cookie-value; "
                "person@example.invalid dpapi-user:QUJD C:\\Users\\PrivateUser\\file.log "
                "workspace_id=old-workspace " + diagnostic_marker,
                [diagnostic_marker],
            )
            for forbidden in (
                "token-value",
                "cookie-value",
                "person@example.invalid",
                "dpapi-user:QUJD",
                "PrivateUser",
                "old-workspace",
                diagnostic_marker,
            ):
                if forbidden in redacted:
                    fail(f"diagnostic redaction leaked {forbidden!r}")
            with tempfile.TemporaryDirectory(prefix="zadoo-diagnostic-smoke-") as diagnostic_dir:
                bundle = build_diagnostic_bundle(
                    Path(diagnostic_dir) / "diagnostics.zip",
                    store=isolated_store,
                    runtime_status={
                        "success": True,
                        "public_url": "https://smoke-secret.trycloudflare.com",
                    },
                )
                with zipfile.ZipFile(bundle) as archive:
                    members = set(archive.namelist())
                    if "diagnostics.json" not in members:
                        fail("diagnostic bundle is missing diagnostics.json")
                    combined = b"\n".join(archive.read(name) for name in members).decode(
                        "utf-8", errors="replace"
                    )
                for forbidden in (settings_code, "alerts@example.invalid", "smoke-secret"):
                    if forbidden in combined:
                        fail(f"diagnostic bundle leaked {forbidden!r}")
                if str(isolated_store.path) in combined:
                    fail("diagnostic bundle exposed the absolute settings path")

            image_ws = FakeSendWebSocket()
            too_large_png = base64.b64encode(b"123456789").decode("ascii")
            server._handle_set_clipboard_image({"mime": "image/png", "data_base64": too_large_png}, image_ws)
            await asyncio.sleep(0.05)
            if not image_ws.sent:
                fail("clipboard image limit did not send a result")
            image_payload = json.loads(image_ws.sent[-1])
            if image_payload.get("success") is not False or "size limit" not in image_payload.get("error", ""):
                fail(f"clipboard image limit returned unexpected payload: {image_payload}")

            signature_cases = (
                ("image/png", b"8BPS\x00\x01", "image data does not match declared MIME type image/png"),
                ("image/png", b"\xff\xd8\xff\xe0", "image data does not match declared MIME type image/png"),
                ("image/jpeg", b"\x89PNG\r\n\x1a\n", "image data does not match declared MIME type image/jpeg"),
            )
            for mime, payload, expected_error in signature_cases:
                try:
                    server._validate_clipboard_image_signature(mime, payload)
                    fail(f"clipboard image signature accepted invalid {mime} payload")
                except ValueError as exc:
                    if str(exc) != expected_error:
                        fail(f"clipboard image signature returned wrong error: {exc}")
            server._validate_clipboard_image_signature("image/png", b"\x89PNG\r\n\x1a\n")
            server._validate_clipboard_image_signature("image/jpeg", b"\xff\xd8\xff\xe0")
            server._validate_clipboard_image_signature("image/jpg", b"\xff\xd8\xff\xe0")

            mic_a, mic_b = object(), object()
            server.mic_clients = {mic_a: queue.Queue(maxsize=5), mic_b: queue.Queue(maxsize=5)}
            server._broadcast_realtime_frame(server.mic_clients, b"same-frame")
            if (
                server.mic_clients[mic_a].get_nowait() != b"same-frame"
                or server.mic_clients[mic_b].get_nowait() != b"same-frame"
            ):
                fail("per-viewer media queues did not receive the same frame")
            stop_calls = []
            original_stop_mic = server._stop_mic_capture
            try:
                server._stop_mic_capture = lambda: stop_calls.append("stop")
                if server._remove_mic_client(mic_a):
                    fail("mic capture stopped while another mic client remained")
                if stop_calls:
                    fail("mic stop called before last client disconnected")
                if not server._remove_mic_client(mic_b) or stop_calls != ["stop"]:
                    fail("mic capture did not stop after last client disconnected")
            finally:
                server._stop_mic_capture = original_stop_mic

        asyncio.run(route_checks())
    finally:
        for key, value in old_env.items():
            if value is None:
                os.environ.pop(key, None)
            else:
                os.environ[key] = value
        settings_mod._STORE = None
        temp_settings.cleanup()

    ok("imports, templates, assets, server instantiation, and in-process routes")


def assert_source_clean() -> None:
    for path in source_files():
        text = path.read_text(encoding="utf-8")
        for forbidden in FORBIDDEN_SOURCE_STRINGS:
            if forbidden in text:
                fail(f"forbidden source string {forbidden!r} found in {path.relative_to(ROOT)}")
        if GITHUB_TOKEN_RE.search(text):
            fail(f"GitHub token-like value found in {path.relative_to(ROOT)}")
        match = MOJIBAKE_RE.search(text)
        if match:
            fail(f"mojibake marker {match.group(0)!r} found in {path.relative_to(ROOT)}")
    ok("source has no forbidden secrets or mojibake markers")


def fetch(base_url: str, path: str, headers: dict[str, str] | None = None):
    url = base_url.rstrip("/") + path
    try:
        request = urllib.request.Request(url, headers=headers or {})
        with urllib.request.urlopen(request, timeout=8) as response:
            return response.status, response.headers.get("Content-Type", ""), response.read()
    except urllib.error.HTTPError as e:
        return e.code, e.headers.get("Content-Type", ""), e.read()


def assert_live(base_url: str) -> None:
    status, content_type, body = fetch(base_url, "/")
    if status == 403:
        fail(
            "live root returned 403 — direct (non-tunnel) access is blocked by default. "
            "Start the target server with ZADOO_ALLOW_DIRECT_ACCESS=1 to run live localhost "
            "smoke tests, e.g.  set ZADOO_ALLOW_DIRECT_ACCESS=1 && python -m zadoo_vnc"
        )
    if status != 200 or "text/html" not in content_type:
        fail(f"live root failed: status={status} content_type={content_type}")

    for path in ("/brand-header.png", "/splash.png", "/trigger-icon.png"):
        status, content_type, body = fetch(base_url, path)
        if status != 200 or "image/png" not in content_type or not body:
            fail(f"live asset failed: {path} status={status} content_type={content_type}")

    status, content_type, body = fetch(base_url, "/benchmark.html")
    if status != 403:
        fail(f"live unauthenticated benchmark returned status={status}, expected 403")

    live_auth_code = os.environ.get("ZADOO_SMOKE_AUTH_CODE") or os.environ.get("ZADOO_ACCESS_CODE")
    if not live_auth_code:
        fail("live checks require ZADOO_ACCESS_CODE or ZADOO_SMOKE_AUTH_CODE because no hardcoded auth default exists")

    legacy_url = base_url.rstrip("/") + "/api/auth"
    legacy_request = urllib.request.Request(legacy_url, headers={"X-Zadoo-Code": _legacy_auth_strings()[0]})
    try:
        with urllib.request.urlopen(legacy_request, timeout=8) as response:
            legacy_status = response.status
            response.read()
    except urllib.error.HTTPError as e:
        legacy_status = e.code
    if legacy_status != 401:
        fail(f"live legacy auth returned status={legacy_status}, expected 401")

    auth_url = base_url.rstrip("/") + "/api/auth"
    live_auth_headers = {"X-Zadoo-Code": live_auth_code}
    auth_request = urllib.request.Request(auth_url, headers=live_auth_headers)
    try:
        with urllib.request.urlopen(auth_request, timeout=8) as response:
            auth_status = response.status
            auth_cookies = response.headers.get_all("Set-Cookie") or []
            auth_cookie = "; ".join(
                cookie.split(";", 1)[0].strip()
                for cookie in auth_cookies
                if cookie.split(";", 1)[0].strip()
            )
            response.read()
    except urllib.error.HTTPError as e:
        auth_status = e.code
        auth_cookies = e.headers.get_all("Set-Cookie") or []
        auth_cookie = "; ".join(
            cookie.split(";", 1)[0].strip()
            for cookie in auth_cookies
            if cookie.split(";", 1)[0].strip()
        )
    if auth_status != 200 or "zadoo_auth=" not in auth_cookie:
        fail(f"live auth failed: status={auth_status}")
    auth_headers = {"Cookie": auth_cookie.split(";", 1)[0]}

    status, content_type, body = fetch(base_url, "/api/list-cameras", auth_headers)
    if status != 200:
        fail(f"live list-cameras failed: status={status}")
    assert_camera_payload(body)

    status, content_type, body = fetch(base_url, "/api/list-mics", auth_headers)
    if status != 200:
        fail(f"live list-mics failed: status={status}")
    assert_mic_payload(body)

    status, content_type, body = fetch(base_url, "/benchmark.html", auth_headers)
    if status != 200 or "text/html" not in content_type or b"Stream Benchmark" not in body:
        fail(f"live benchmark route failed: status={status} content_type={content_type}")

    ok(f"live HTTP checks passed for {base_url}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--live", help="Base URL for live HTTP route checks, for example http://localhost:6173")
    args = parser.parse_args()

    assert_imports()
    assert_source_clean()
    if args.live:
        assert_live(args.live)


if __name__ == "__main__":
    main()

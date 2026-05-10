# Zadoo VNC

Zadoo VNC is a Windows-focused remote screen, input, clipboard, and Cloudflare tunnel runtime. The compatibility entrypoint is `zadoo_vnc_single.py`; the package entrypoint is `zadoo-vnc`.

## Setup

Use Python 3.11 on Windows.

```powershell
python -m pip install -r requirements.txt
```

Optional feature groups are defined in `pyproject.toml`:

```powershell
python -m pip install -e .[media,email,perf,ssh]
```

## Run

```powershell
python zadoo_vnc_single.py
```

The app starts a local web UI, screen capture, input WebSockets, and one Cloudflare tunnel when `cloudflared.exe` is available or auto-download is explicitly enabled.

## Configuration

Copy `.env.example` to `.env` and fill only the values you need.

- `RESEND_API_KEY`, `RESEND_FROM`, and `EMAIL_TO` or `GMAIL_TO` enable tunnel email notifications.
- `CODE_FULL`, `CODE_LIMITED`, `CODE_PARTIAL`, `CODE_LOCKDOWN`, and `CUSTOM_PASSWORD` control server-side access sessions. If none are set, the app prints one temporary full-access code at startup.
- `ZADOO_DISABLE_TUNNEL=1` keeps the UI local-only.
- `ZADOO_CLOUDFLARED_PATH` points at a managed `cloudflared.exe`; `ZADOO_AUTO_DOWNLOAD_CLOUDFLARED=1` permits downloading it when no local binary exists.
- `ZADOO_INSTALL_STARTUP_TASK=1` and `ZADOO_ENABLE_PROCESS_PROTECTOR=1` opt into persistence/restart behavior for packaged deployments.
- `ALERT_A`, `ALERT_B`, `ALERT_C`, and `ALERT_D` customize host alert presets.

No API keys or access codes are intentionally bundled. If a previous key or code was exposed in source or logs, rotate it before using public links.

## Smoke Tests

```powershell
python -m compileall zadoo_vnc zadoo_vnc_single.py scripts
python scripts/smoke_test.py
python scripts/smoke_test.py --live http://localhost:6173
```

The live smoke test assumes the app is already running.

## Notes

The core dependency set covers screen sharing, remote input, text clipboard, templates, and local/tunnel web serving. Audio, webcam, email, SSH PTY, and faster capture backends are optional extras.

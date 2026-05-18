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

The app starts a local web UI, screen capture, input WebSockets, and one Cloudflare tunnel when `cloudflared.exe` is available or can be downloaded.

## Configuration

Copy `.env.example` to `.env` and fill only the values you need.

- `RESEND_API_KEY`, `RESEND_FROM`, and `EMAIL_TO` or `GMAIL_TO` enable tunnel email notifications.
- `CODE_FULL`, `CODE_LIMITED`, `CODE_PARTIAL`, `CODE_LOCKDOWN`, and `CUSTOM_PASSWORD` control server-side access sessions. If none are set, the app prints one temporary full-access code at startup.
- `ZADOO_ALLOWED_ORIGINS` adds comma-separated extra HTTP/WebSocket origins. Same-host browser origins are allowed by default.
- `ZADOO_ALLOW_QUERY_AUTH=1` temporarily re-enables legacy `/api/auth?code=...`; the UI uses the safer `X-Zadoo-Code` header by default.
- `ZADOO_AUTH_MAX_FAILURES`, `ZADOO_AUTH_WINDOW_SECONDS`, and `ZADOO_AUTH_LOCKOUT_SECONDS` tune in-memory login throttling.
- `ZADOO_DISABLE_STARTUP_TASK=1` prevents frozen builds from creating the Windows logon task.
- `ZADOO_CLOUDFLARED_DOWNLOAD_TIMEOUT` and `ZADOO_SKIP_CLOUDFLARED_SIGNATURE_CHECK` control cloudflared download and signature verification behavior.
- `ZADOO_CLIPBOARD_TEXT_MAX_BYTES` and `ZADOO_CLIPBOARD_IMAGE_MAX_BYTES` cap remote clipboard payload sizes.
- `ZADOO_DETECT_GPU_NAMES=1` enables optional PowerShell GPU-name detection; it is off by default to keep startup responsive.
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

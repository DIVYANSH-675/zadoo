# Zadoo VNC

Zadoo VNC is a Windows-focused remote screen, input, clipboard, and Cloudflare tunnel runtime. The compatibility entrypoint is `zadoo_vnc_single.py`; the package entrypoint is `zadoo-vnc`.

## Setup

Use Python 3.11 on Windows.

```powershell
python -m pip install -r requirements.txt
```

Optional feature groups are defined in `pyproject.toml`:

```powershell
python -m pip install -e .[media,email,perf,ssh,ocr]
```

## Run

```powershell
python zadoo_vnc_single.py
```

The app starts a local web UI, screen capture, input WebSockets, and one Cloudflare tunnel when `cloudflared.exe` is available or can be downloaded.

## Configuration

Copy `.env.example` to `.env` and fill only the values you need.

- `RESEND_API_KEY`, `RESEND_FROM`, and `EMAIL_TO` or `GMAIL_TO` enable tunnel email notifications.
- `CODE_FULL`, `CODE_LIMITED`, `CODE_PARTIAL`, `CODE_LOCKDOWN`, and `CUSTOM_PASSWORD` control server-side access sessions.
- `ALERT_A`, `ALERT_B`, `ALERT_C`, and `ALERT_D` customize host alert presets.

No Resend API key is intentionally bundled. If a previous key was exposed in source or logs, rotate it before using email notifications.

## Smoke Tests

```powershell
python -m compileall zadoo_vnc zadoo_vnc_single.py scripts
python scripts/smoke_test.py
python scripts/smoke_test.py --live http://localhost:6173
```

The live smoke test assumes the app is already running.

## Notes

The core dependency set covers screen sharing, remote input, text clipboard, templates, and local/tunnel web serving. Audio, webcam, OCR, email, SSH PTY, and faster capture backends are optional extras.

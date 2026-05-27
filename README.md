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
- `ZADOO_CAPTURE_METHOD=bettercam` or `dxcam` selects the screen capture backend. BetterCam is the default when installed.
- `ZADOO_ADAPTIVE_STREAM=1` keeps screen sharing on the adaptive JPEG WebSocket path.
- `ZADOO_STREAM_START_PROFILE=720p120` is the stable default. Use `720p240` or `1080p240` to opt in to higher-FPS startup profiles when the host, browser, and network can keep up.
- `ZADOO_STREAM_MODE=adaptive_jpeg_ws` is the active video transport. WebRTC/H.264 is not enabled by this build.
- `ALERT_A`, `ALERT_B`, `ALERT_C`, and `ALERT_D` customize host alert presets.

## Screen Sharing Performance

Live screen sharing uses one explicit capture backend at a time. The UI lets the user switch directly between BetterCam and DXCam, and BetterCam is chosen by default when available to avoid DXGI device conflicts between the two libraries. Set `ZADOO_CAPTURE_METHOD=dxcam` on hosts where DXCam is known to be stable. DXCam `0.3.0` is driven through its ring-buffer API, `start(region, target_fps, video_mode)` plus `get_latest_frame()`, because that is the high-throughput path. This installed DXCam API does not expose the researched `processor_backend="cv2"` argument.

Install `imagecodecs` with the requirements file so JPEG encoding uses the fast path. The default `720p120` profile applies quality and scaling automatically until the user manually changes the quality slider. For maximum FPS, set `ZADOO_STREAM_START_PROFILE=720p240`; use `1080p240` only on fast local networks and capable hardware.

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

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
- `ZADOO_FFMPEG_PATH` and `ZADOO_MEDIAMTX_PATH` enable the optional Turbo WebRTC stream. The app probes `ddagrab`, `gfxcapture`, `gdigrab`, `h264_mf`, `h264_qsv`, `h264_amf`, `h264_nvenc`, and `libx264`, benchmarks a safe profile, and falls back to JPEG WebSocket when WebRTC is not available. Run `powershell -ExecutionPolicy Bypass -File scripts\setup_turbo_stream.ps1` to install the local toolchain under `tools/`.
- `ZADOO_TURBO_PROFILE`, `ZADOO_TURBO_CAPTURE`, and `ZADOO_TURBO_ENCODER` can force a specific profile or backend for diagnostics. Default laptop target is `540p30`; upgrades to `540p60` or `720p30` happen only after a local benchmark passes.
- `ZADOO_TURBO_TRANSPORT=rtsp` is the default fast path: FFmpeg publishes H.264 to MediaMTX over RTSP and browsers play it back through MediaMTX WebRTC. `whip` and `auto` are available for diagnostics.
- `ZADOO_INSTALL_STARTUP_TASK=1` and `ZADOO_ENABLE_PROCESS_PROTECTOR=1` opt into persistence/restart behavior for packaged deployments.
- `ALERT_A`, `ALERT_B`, `ALERT_C`, and `ALERT_D` customize host alert presets.

No API keys or access codes are intentionally bundled. If a previous key or code was exposed in source or logs, rotate it before using public links.

## Turbo WebRTC Notes

The fastest path needs both `ffmpeg.exe` and `mediamtx.exe`.

```powershell
powershell -ExecutionPolicy Bypass -File scripts\setup_turbo_stream.ps1
python zadoo_vnc_single.py
```

The installer downloads FFmpeg and MediaMTX into `tools/`, which is ignored by git. You can also point at your own binaries:

```powershell
$env:ZADOO_FFMPEG_PATH="C:\tools\ffmpeg\bin\ffmpeg.exe"
$env:ZADOO_MEDIAMTX_PATH="C:\tools\mediamtx\mediamtx.exe"
python zadoo_vnc_single.py
```

Local/LAN playback uses MediaMTX RTSP on port `8554`, MediaMTX WebRTC on port `8889`, and the app remains on port `6173`. For public internet WebRTC, expose the MediaMTX WebRTC HTTP and media ports and configure STUN/TURN or MediaMTX `webrtcAdditionalHosts`; an HTTP-only tunnel can still show the JPEG fallback but cannot reliably carry WebRTC media.

## Smoke Tests

```powershell
python -m compileall zadoo_vnc zadoo_vnc_single.py scripts
python scripts/smoke_test.py
python scripts/smoke_test.py --live http://localhost:6173
```

The live smoke test assumes the app is already running.

## Notes

The core dependency set covers screen sharing, remote input, text clipboard, templates, and local/tunnel web serving. Audio, webcam, email, SSH PTY, and faster capture backends are optional extras.

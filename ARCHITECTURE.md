# Zadoo host architecture

Zadoo is a Windows 10/11 x64 host runtime. The hosted dashboard and billing service are maintained in the separate `zadoo-web` repository.

## Entrypoints and lifecycle

- `python -m zadoo_vnc` and the `zadoo-vnc` console script call `zadoo_vnc.app:main`.
- `--open` and `--settings` open the native Settings window; no arguments start the runtime.
- The runtime starts screen capture, binds one `websockets` 15 server to port `6173`, verifies the first encoded frame, then starts the signed, bundled `cloudflared.exe`.
- Installed settings and logs live under `%ProgramData%\Zadoo`.

## Runtime modules

| Module | Responsibility |
|---|---|
| `app.py` | Startup, entitlement gate, capture and server lifecycle |
| `server.py` | WebSocket server, background tasks, tunnel launch, cloud heartbeat |
| `routes.py` | HTTP handshake responses, authentication, permissions, snapshots, cloud proxies |
| `media.py` | Screen broadcast, system audio, microphone, webcam, terminal |
| `input_control.py` | Mouse, keyboard, Unicode typing, clipboard, alerts and host hotkeys |
| `screen_capture.py` | BetterCam capture, adaptive scaling and `imagecodecs` JPEG encoding |
| `streaming.py` | Always-on adaptive stream ladder and diagnostics |
| `camera_discovery.py` | Windows camera enumeration and stable identifiers |
| `tunnel.py` | Signed Cloudflare quick tunnel and optional Resend notification |
| `settings.py` | Normalized JSON settings, DPAPI secrets and access-code verification |
| `settings_window.py` | Native owner configuration and runtime controls |
| `saas.py` | Device activation, entitlement, session metering and heartbeat calls |

## Request flow

1. Cloudflare forwards the public HTTPS request to `localhost:6173`.
2. `RoutesMixin.process_request` validates origin, route, session cookie, permission and CSRF requirements.
3. `/api/auth` verifies the access code and creates an in-memory one-hour session.
4. Permissions are stored in that session, avoiding settings-file I/O on mouse, keyboard and stream actions.
5. WebSocket routes dispatch to `/video`, `/input`, `/audio`, `/mic`, `/webcam` or `/terminal`.

Direct LAN access is denied by default; localhost access from the host is allowed. `ZADOO_ALLOW_DIRECT_ACCESS=1` is intended only for development and live smoke tests.

## Screen stream

The shipping video transport is adaptive JPEG frames over WebSocket:

1. BetterCam desktop duplication produces RGB NumPy frames.
2. The capture thread applies the selected region, optional grayscale and integer downscaling.
3. `imagecodecs` encodes JPEG directly from the contiguous array.
4. The broadcast loop keeps at most one in-flight send per client and skips stale frames instead of building delay.
5. Browser and server statistics drive the adaptive FPS, quality and scale ladder.

MSS seeds the initial screen, then BetterCam desktop duplication returns only changed frames. Static desktops are polled but not re-encoded or retransmitted. A capture or encode exception stops the capture thread and is exposed as an exact error.

## Input and clipboard

- Mouse buttons, movement and wheel events use Win32 `SendInput` only.
- Text input uses Win32 `SendInput`, including `KEYEVENTF_UNICODE`, without modifying the clipboard.
- Text and image clipboard sync use the native Windows clipboard through `pywin32`.
- The authenticated UI receives the configured clipboard byte limit; the WebSocket frame cap is derived from that limit and remains bounded.
- Each input message is permission-checked in memory. Processing errors are returned as `input_error` WebSocket messages.

## Media and terminal

- System audio uses `soundcard` loopback capture.
- Microphone input uses `sounddevice`.
- One capture pipeline publishes into an independent five-frame drop-oldest queue per audio or microphone viewer.
- Webcam capture uses the selected DirectShow device through PyAV at 640x480 and 30 fps.
- Terminal sessions use PowerShell through ConPTY (`pywinpty`).
- CodeMirror 5.65.21 and xterm 5.3.0 with fit 0.8.0 are pinned, bundled locally and served with immutable caching.

## Billing

The local billing proxy forwards trusted country headers to `zadoo-web`. India viewers receive INR wallet pricing; all other viewers receive USD pricing. The browser submits minor currency units, while the hosted service remains authoritative for the market, minimum, rate and Razorpay order.

## Tunnel and notifications

The build downloads cloudflared 2026.6.0 from its versioned release, verifies its pinned SHA-256 and Authenticode signature, then bundles it as `cloudflared.exe`. Runtime never downloads or searches for alternate binaries. It verifies the Authenticode signature once per process and reports exact startup/timeout/exit errors through the local APIs and Settings UI. Resend email is optional and runs only after a public URL exists.

## Build and verification

`scripts/build_windows.ps1` requires Python 3.11.9 and Node.js 24.18.0 x64. It creates one x64 build environment, installs hash-locked dependencies, validates imports, runs source checks, verifies downloaded build tools, builds the installer and optional portable executable, signs artifacts, and emits a SHA-256 release manifest. Every invocation must select a PFX or explicitly request unsigned output with `-NoSelfSign`; the build never creates or trusts certificates. GitHub CI repeats the unsigned build, dependency audits, manifest verification, and CycloneDX SBOM generation on Windows 2025 x64.

The standard verification set is:

```powershell
python -m compileall -q zadoo_vnc scripts
python -m ruff check .
python scripts\smoke_test.py
python scripts\check_template_js.py
python scripts\check_workflow_pins.py
```

## Supported configuration overrides

The native Settings window is the user-facing configuration surface. Source-development overrides include `ZADOO_ACCESS_CODE`, `ZADOO_STREAM_START_PROFILE`, `ZADOO_TARGET_KBPS`, `ZADOO_CLOUDFLARED_PATH`, `ZADOO_CLOUDFLARED_PROTOCOL`, `ZADOO_ALLOWED_ORIGINS` and explicit test/disable switches referenced in code. Invalid numeric and capture-profile values fail with a named error.

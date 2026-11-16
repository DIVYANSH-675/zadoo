# Zadoo

Zadoo is a free Windows 10/11 x64 remote-access host. View and control your computer from a browser using a local access code. Screen sharing, input, clipboard, snapshots, audio, camera, terminal, and alerts are included.

## License

Zadoo's original code is licensed under [Apache 2.0](LICENSE). Third-party components retain their own licenses; see [NOTICE](NOTICE) and the notices bundled with each build. Licensing does not imply that a Windows binary is signed. See the [code signing policy](https://zadoo-web.vercel.app/code-signing) for the current approval status.

## Run from source

Use Python 3.11.9 x64:

```powershell
py -3.11-64 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --require-hashes -r requirements-build.lock
.\.venv\Scripts\python.exe -m pip install --require-hashes -r requirements-runtime.lock
.\.venv\Scripts\python.exe -m pip install --no-deps --no-build-isolation -e .
.\.venv\Scripts\python.exe -m zadoo_vnc --settings
```

Create an access code (up to 10 characters), review permissions, then choose **Start Zadoo**. Open the public link shown in Settings on another device and enter the code. Terminal access is opt-in.

No account or payment is required. The optional sign-in on the Device page links your device to the separate [zadoo-web](https://github.com/DIVYANSH-675/zadoo-web) dashboard. A dashboard outage does not block or stop remote access.

## Workspace controls

The Windows settings window uses responsive navigation, scrollable pages and plain-language permission explanations. Settings save automatically. The browser's fixed control dock keeps Fit/Width, zoom, Reset and the hand control together; its control panel supports small screens, keyboard navigation and reduced motion. The live screen has no animated visual effects.

Support the developer is optional. It opens the supplied Google Pay QR locally and never unlocks or restricts a feature. Verify the displayed recipient in your payment app. The dark Windows installer has a cartoon Zadoo companion, welcome/install/finish scenes and a short connection checklist. Motion uses brief, capped native-timer bursts, pauses when the window is inactive, and starts from the Windows animation preference; a checkbox can turn it off. Installation progress comes from actual file operations, not an animation timer. Mascot artwork is installer-only and adds no work to the streaming runtime.

For local development, copy `.env.example` to `.env`, set `ZADOO_ACCESS_CODE`, `ZADOO_DISABLE_TUNNEL=1`, and `ZADOO_ALLOW_DIRECT_ACCESS=1`, then run `python -m zadoo_vnc`. Never commit real credentials.

## Streaming

Current browsers negotiate incremental JPEG updates: only changed desktop regions are encoded and drawn, with full frames for initialization, format changes, and recovery. The clear profile starts at full resolution and JPEG quality 85 with 4:4:4 colour for text/desktop updates. Large motion uses faster 4:2:0 colour sampling, followed by a full-detail refresh when motion stops; an explicit high-quality selection keeps full colour detail. Slow hosts, browsers, or connections can still adapt down; actual FPS depends on the desktop source and available resources. Full-screen video has fewer unchanged pixels and does not get the same bandwidth savings as text-and-window activity.

The viewer displays measured FPS and provides an **Auto** quality button to return from a manual selection to adaptation. Still desktops intentionally send fewer updates. Input listeners survive reconnects without duplication; mouse bursts keep only the latest pending position, while the first movement and clicks remain immediate. Drag/key releases are sent when focus leaves the viewer to avoid stuck input. The dark cyan controls support keyboard navigation and reduced motion.

Queues remain bounded. Sequence checks request a fresh full frame after a missed update, and mixed sessions automatically use ordinary JPEG while an older client is present.

## Security and limits

Access codes and optional device tokens use current-user Windows DPAPI. Settings and logs have protected per-user ACLs. Browser mutations use authenticated same-origin WebSocket RPC. Up to five distinct authenticated viewers can connect, bounding host resource usage; there is no usage timer.

The public connection uses a Cloudflare HTTPS tunnel. Traffic passes through Cloudflare; this is not a claim of application-level end-to-end encryption. Share the URL and code only with trusted people. Stop the host or change its access code to end access.

## Reproducible Windows x64 build

```powershell
.\scripts\build_windows.ps1 -NoSelfSign -InstallMissingTools
```

Build inputs are hash-locked. Python 3.11.9 x64, Node.js 24.18.0 x64, Inno Setup 7.0.2 x64, and the pinned signed cloudflared binary are verified. Output is under `dist\portable` and `dist\installer`, with a SHA-256 manifest. Trusted-signed releases require a real code-signing certificate. The build never creates or trusts a substitute certificate.

The owner has separately approved [Zadoo 1.1.5 as an unsigned public preview](https://github.com/DIVYANSH-675/zadoo/releases/tag/v1.1.5-preview.1). Its filename and download page explicitly say **unsigned preview**. Windows may show an unknown-publisher or SmartScreen warning; do not disable Windows security protections. The checksum identifies the reviewed file but does not authenticate its publisher or guarantee software safety. Download and verification details are on the [official website](https://zadoo-web.vercel.app/download). The trusted-signed release workflow remains unchanged.

## Verify

```powershell
.\.venv\Scripts\python.exe -m ruff check zadoo_vnc scripts
.\.venv\Scripts\python.exe -m compileall -q zadoo_vnc scripts
.\.venv\Scripts\python.exe -m vulture zadoo_vnc scripts --min-confidence 80
.\.venv\Scripts\python.exe scripts\smoke_test.py
.\.venv\Scripts\python.exe scripts\settings_ui_test.py
.\.venv\Scripts\python.exe scripts\installer_ui_test.py
.\.venv\Scripts\python.exe scripts\frame_dispatch_test.py
.\.venv\Scripts\python.exe scripts\stream_policy_test.py
.\.venv\Scripts\python.exe scripts\delta_frames_test.py
node scripts\input_controls_test.mjs
.\.venv\Scripts\python.exe scripts\check_template_js.py
.\.venv\Scripts\python.exe scripts\check_workflow_pins.py
```

For browser integration tests install the `test` extra and Chromium, start an isolated host, set `ZADOO_SMOKE_AUTH_CODE`, and run the live smoke and `scripts/terminal_e2e_test.py`. `scripts/viewer_ui_test.py --code <test-code>` covers connection errors, keyboard controls, Auto quality, zoom/reset, optional support, reconnects and responsive layouts; its synthetic input is intercepted rather than sent to the host. `scripts/runtime_e2e_test.py` includes this browser check. `scripts/settings_ui_test.py --preview` opens an isolated native preview without starting sharing or changing Windows preferences. `scripts/preview_installer.ps1` builds a separate, inert visual wizard that cannot install anything. See `OPERATIONS.md` for installation, diagnostics, signing, and release checks.

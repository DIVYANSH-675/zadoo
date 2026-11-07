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

No account or payment is required. The optional Dashboard tab links your device to the separate [zadoo-web](https://github.com/DIVYANSH-675/zadoo-web) dashboard. A dashboard outage does not block or stop remote access.

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

Build inputs are hash-locked. Python 3.11.9 x64, Node.js 24.18.0 x64, Inno Setup 7.0.2 x64, and the pinned signed cloudflared binary are verified. Output is under `dist\portable` and `dist\installer`, with a SHA-256 manifest. Unsigned artifacts are for testing; public releases require a real code-signing certificate. The build never creates or trusts a substitute certificate.

## Verify

```powershell
.\.venv\Scripts\python.exe -m ruff check zadoo_vnc scripts
.\.venv\Scripts\python.exe -m compileall -q zadoo_vnc scripts
.\.venv\Scripts\python.exe -m vulture zadoo_vnc scripts --min-confidence 80
.\.venv\Scripts\python.exe scripts\smoke_test.py
.\.venv\Scripts\python.exe scripts\frame_dispatch_test.py
.\.venv\Scripts\python.exe scripts\stream_policy_test.py
.\.venv\Scripts\python.exe scripts\delta_frames_test.py
node scripts\input_controls_test.mjs
.\.venv\Scripts\python.exe scripts\check_template_js.py
.\.venv\Scripts\python.exe scripts\check_workflow_pins.py
```

For browser integration tests install the `test` extra and Chromium, start an isolated host, set `ZADOO_SMOKE_AUTH_CODE`, and run the live smoke and `scripts/terminal_e2e_test.py`. `scripts/viewer_ui_test.py --code <test-code>` covers connection errors, keyboard controls, Auto quality, reconnects and responsive layouts; its synthetic input is intercepted rather than sent to the host. `scripts/runtime_e2e_test.py` includes this browser check. See `OPERATIONS.md` for installation, diagnostics, signing, and release checks.

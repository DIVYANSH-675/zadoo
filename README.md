# Zadoo

Zadoo is a Windows-focused remote screen, input, clipboard, media, terminal, alert, and Cloudflare tunnel runtime. The source package remains `zadoo_vnc`; the compatibility entrypoint is `zadoo_vnc_single.py`.

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

The app starts a local web UI on the fixed port `6173`. Installed builds open the small native `Zadoo Settings` window on first launch; create one visible access code, at most 10 characters, and choose the permissions that code is allowed to use. The Cloudflare tunnel stays disabled until setup is complete.

## Configuration

Installed builds store user settings in:

```text
%ProgramData%\Zadoo\config.json
```

Settings include the access-code hash, Resend API key encrypted with Windows DPAPI, recipient email, alert slots, the single permission matrix, SaaS activation state, startup state, taskbar behavior, and setup completion.

- One access code is used for all sessions. It is limited to 10 characters, and the old value is never revealed.
- Permissions control mouse, keyboard, clipboard pull, clipboard push, system audio, mic, camera, terminal, snapshots, advanced video controls, tunnel refresh, and remote alerts.
- Alert A-D slots have no defaults. Blank slots are disabled.
- The Resend area shows `Email not Set` until both the API key and recipient email are configured.
- The Settings window can start and stop the runtime, start device activation with the hosted SaaS, refresh entitlement, toggle autostart, and choose whether minimized Settings remains visible on the taskbar.

`.env.example` is now for development/test fallback only. User-facing configuration should be done from the native Zadoo Settings window.

## Hosted SaaS

The hosted billing and dashboard app lives in `web/`.

```powershell
cd web
npm install
npm run prisma:generate
npm run prisma:migrate
npm run dev
```

It includes an original Zadoo website, pricing, Auth.js Google/email OTP login, dashboard, device activation, billing history, Razorpay checkout, signed webhooks, entitlement APIs, and agent session usage endpoints. Configure `web\.env.local` from `web\.env.example` with Postgres, Auth.js, Razorpay, and email provider credentials.

## Screen Sharing Performance

Live screen sharing uses one explicit capture backend at a time. The UI lets the user switch directly between BetterCam and DXCam, and BetterCam is chosen by default when available to avoid DXGI device conflicts between the two libraries. Set `ZADOO_CAPTURE_METHOD=dxcam` on hosts where DXCam is known to be stable. DXCam `0.3.0` is driven through its ring-buffer API, `start(region, target_fps, video_mode)` plus `get_latest_frame()`, because that is the high-throughput path. This installed DXCam API does not expose the researched `processor_backend="cv2"` argument.

Install `imagecodecs` with the requirements file so JPEG encoding uses the fast path. The stream starts at the bandwidth-safe `540p60` profile, and the adaptive controller raises or lowers both resolution and FPS to fit the link: it upshifts toward high FPS on fast/local connections and downshifts resolution first (then FPS) on slow links so text stays legible. Quality and scaling adapt automatically unless the user manually changes the quality slider (which pins quality and disables resolution adaptation for that session). For maximum FPS on a fast local network, set `ZADOO_STREAM_START_PROFILE=720p240` or `1080p240`. For a known slow link, set `ZADOO_TARGET_KBPS` (for example `2000` for a 2 Mbps line) so the controller proactively caps egress instead of waiting for queues to back up.

No API keys or access codes are intentionally bundled. If a previous key or code was exposed in source or logs, rotate it before using public links.

## Windows Packaging

The packaging entrypoint is:

```powershell
.\scripts\build_windows.ps1 -Arch x64
.\scripts\build_windows.ps1 -Arch All -PfxPath C:\path\codesign.pfx
```

The script creates isolated build environments under `.build_envs\py311-x64` and `.build_envs\py311-x86`, installs `requirements.txt` for x64 or `requirements-x86.txt` for x86 plus `build_requirements.txt`, validates imports, bundles the matching signed `cloudflared.exe`, builds a PyInstaller one-folder app for the installer, builds portable one-file EXEs, signs the EXEs/installers, and compiles Inno Setup installers.

Required local tools:

- Python 3.11 x64 for x64 builds.
- Python 3.11 x86 for x86 builds.
- Inno Setup 6 for installers.
- Windows SDK `signtool.exe` and a PFX for release signing. Without a PFX, local builds use a self-signed `CN=Zadoo Local Build` certificate.
- `C:\Users\divya\Real\app_icon.ico` for the EXE and installer icon.

Artifacts are written under `dist\onedir`, `dist\portable`, and `dist\installer`. The x86 build excludes PyAV, pywinpty, DXCam, and BetterCam because those packages do not provide stable Python 3.11 win32 support here; x86 uses the stable fallback capture/camera paths and terminal reports unavailable when PTY support is missing.

## Smoke Tests

```powershell
python -m compileall zadoo_vnc zadoo_vnc_single.py scripts
python scripts/smoke_test.py
python scripts/check_template_js.py
python scripts/smoke_test.py --live http://localhost:6173
```

The live smoke test assumes the app is already running. For configured installs, set `ZADOO_SMOKE_AUTH_CODE` and, if needed, `ZADOO_SMOKE_SHARE_TOKEN` before running the live check.

## Notes

The core dependency set covers screen sharing, remote input, text clipboard, templates, and local/tunnel web serving. Audio, webcam, email, SSH PTY, and faster capture backends are optional extras.

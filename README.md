# Zadoo

Zadoo is a Windows x64 remote screen, input, clipboard, media, terminal, alert, and Cloudflare tunnel runtime.

See `OPERATIONS.md` for exact source setup, health checks, signed release verification, rollback,
and the production checklist. See `CONTRIBUTING.md` for validation and dependency-lock maintenance.

## Setup

Use Python 3.11.9 on Windows.

```powershell
python -m pip install -e .
```

## Run

```powershell
python -m zadoo_vnc
```

Open the native Settings window first, sign in, create one visible access code of at most 10 characters, and choose its permissions:

```powershell
python -m zadoo_vnc --settings
```

Running without arguments then starts the host runtime on fixed port `6173`. For local source testing without SaaS/tunnel access, explicitly set `ZADOO_DISABLE_TUNNEL=1` and `ZADOO_ACCESS_CODE`.

## Configuration

Installed builds store user settings in:

```text
%ProgramData%\Zadoo\config.json
```

Settings include the DPAPI-encrypted access code, recipient email, alert slots, the single permission matrix, SaaS activation state, startup state, taskbar behavior, and setup completion.

- One access code is used for all sessions. It is limited to 10 characters and is visible only in the local Settings window.
- Permissions control mouse, keyboard, clipboard pull, clipboard push, system audio, mic, camera, terminal, snapshots, advanced video controls, tunnel refresh, and remote alerts.
- Alert A-D slots have no defaults. Blank slots are disabled.
- Email notification requires `RESEND_API_KEY`, `RESEND_FROM`, and a recipient set through `Email To` or `EMAIL_TO`; status names any missing field.
- The Settings window can start and stop the runtime, start device activation with the hosted SaaS, refresh entitlement, toggle autostart, and choose whether minimized Settings remains visible on the taskbar.

`.env.example` documents source-development overrides only. User-facing configuration belongs in the native Zadoo Settings window.

## Hosted SaaS

The hosted billing and dashboard app lives in the separate `zadoo-web` repository.

```powershell
cd ..\zadoo-web
npm install
npm run prisma:generate
npm run prisma:migrate
npm run dev
```

It includes an original Zadoo website, pricing, Auth.js Google/email OTP login, dashboard, device activation, billing history, Razorpay checkout, signed webhooks, entitlement APIs, and agent session usage endpoints. Configure that repository from its own environment example with Postgres, Auth.js, Razorpay, and email provider credentials.

## Screen Sharing Performance

Live screen sharing uses BetterCam desktop duplication directly. MSS provides the one initial frame; afterward unchanged desktops are neither re-encoded nor retransmitted.

The required `imagecodecs` dependency performs JPEG encoding. The stream starts at `half-60-q52` (half-size, 60 FPS, JPEG quality 52), and the always-on adaptive controller adjusts scale, quality, and FPS to fit the host, browser, and link. A manual quality selection pins quality for that session. For maximum FPS on a fast local network, set `ZADOO_STREAM_START_PROFILE=full-240-q52` or `half-240-q54`. For a known slow link, set `ZADOO_TARGET_KBPS` (for example `2000` for a 2 Mbps line).

No API keys or access codes are intentionally bundled. If a previous key or code was exposed in source or logs, rotate it before using public links.

## Windows Packaging

The packaging entrypoint is:

```powershell
.\scripts\build_windows.ps1
.\scripts\build_windows.ps1 -PfxPath C:\path\codesign.pfx
.\scripts\build_windows.ps1 -NoSelfSign
```

The script creates an isolated build environment under `.build_envs\py311-x64`, installs the fully pinned runtime and build dependency graphs from `requirements-runtime.lock` and `requirements-build.lock` with SHA-256 enforcement, validates imports, downloads the pinned x64 `cloudflared.exe` and verifies its SHA-256 plus Authenticode signature, verifies the pinned Inno Setup installer and compiler version, SHA-256, architecture, and signatures, builds a PyInstaller one-folder app for the installer, builds a portable one-file EXE, signs the EXEs/installers, and compiles the Inno Setup installer. A build must explicitly supply either `-PfxPath` for release signing or `-NoSelfSign` for unsigned local artifacts; it never creates or trusts certificates.

Required local tools:

- Python 3.11.9 x64; pass `-PythonPath` when it cannot be resolved through the x64 `py` launcher or its standard install path.
- Node.js 24.18.0 x64 for template JavaScript validation; pass `-NodePath` when it is not on `PATH`.
- Inno Setup 7.0.2 x64 for installers.
- Windows SDK 10.0.26100.7705 `signtool.exe` and a PFX for release signing.
- The repository `app_icon.ico` for the EXE and installer icon.

Final artifacts are written under `dist\portable` and `dist\installer`. Each fresh build removes stale `dist` output and writes `dist\release-manifest.json` plus `dist\SHA256SUMS.txt` with artifact sizes, x64 architecture, Authenticode status, and SHA-256 hashes. Pass `-KeepOneDir` to retain `dist\onedir` and PyInstaller work files.

## Continuous Integration

`.github/workflows/windows-x64-ci.yml` runs on every pull request and push to `main` using a GitHub-hosted Windows 2025 x64 runner. It audits both hash-locked dependency graphs, runs the full source gate, produces an unsigned installer and portable executable, verifies their release manifest, generates a CycloneDX runtime SBOM, and uploads short-lived CI artifacts clearly labeled as unsigned. GitHub Actions are pinned to full commit SHAs and checked by `scripts/check_workflow_pins.py`.

Unsigned CI artifacts are for testing only. A distributable release must be built with `-PfxPath`, must report `signed: true` in `release-manifest.json`, and must have `Valid` Authenticode status for every executable in the manifest.

## Smoke Tests

```powershell
python -m pip install -e ".[test]"
python -m playwright install chromium
python -m compileall zadoo_vnc scripts
python scripts/smoke_test.py
python scripts/check_template_js.py
python scripts/check_workflow_pins.py
python scripts/smoke_test.py --live http://localhost:6173
python scripts/terminal_e2e_test.py --code YOUR_CODE
```

The live smoke test assumes the app is already running. For configured installs, set `ZADOO_SMOKE_AUTH_CODE` before running the live check.

## Notes

The installation includes every runtime feature: screen sharing, input, clipboard, media, email, terminal, and accelerated capture.

# Zadoo operations

Zadoo is free. It runs on Windows 10/11 x64 and requires a local access code. There is no paid activation, minute balance, or billing setup.

## Setup and run

Follow the source/build commands in `README.md`. Installed builds use `%LOCALAPPDATA%\Zadoo\config.json`; keep that file private. Access codes and optional dashboard tokens are protected by the current Windows user's DPAPI.

Open Settings, set an access code, choose permissions, and select **Start Zadoo**. The public link appears after cloudflared starts. A dashboard account is optional. Source test hosts can use `ZADOO_DISABLE_TUNNEL=1` and `ZADOO_ALLOW_DIRECT_ACCESS=1`; direct LAN access is denied by default.

The host uses port 6173. Check local health with:

```powershell
Invoke-RestMethod http://127.0.0.1:6173/api/runtime/status -TimeoutSec 3
```

Check `success`, `running`, `public_url`, `tunnel_error`, and `tunnel_block_reason`. Optional dashboard failure does not gate access. If cloudflared exits, the tunnel monitor reports bounded retries and fails closed after the configured 60–120 seconds.

Stop from Settings. The authenticated RPC returns `Zadoo runtime stopping`. Confirm the process exits and port 6173 is released.

## Logs and support

Logs are under `%LOCALAPPDATA%\Zadoo\logs` and are bounded/rotated. **Export Diagnostics** creates a bounded ZIP and redacts codes, tokens, identifiers, email, cookies, DPAPI data, user paths and tunnel links. Independently inspect the ZIP before sharing.

## Signing and publishing

Unsigned local test build:

```powershell
.\scripts\build_windows.ps1 -NoSelfSign -InstallMissingTools
```

For distribution use a real code-signing PFX. Supply its password only in the process environment:

```powershell
$credential = Get-Credential -UserName codesign -Message "Enter the PFX password"
$env:ZADOO_SIGN_PFX_PASSWORD = $credential.GetNetworkCredential().Password
.\scripts\build_windows.ps1 -PfxPath C:\secure\zadoo-codesign.pfx -InstallMissingTools
Remove-Item Env:\ZADOO_SIGN_PFX_PASSWORD
$credential = $null
```

Verify every SHA-256 and Authenticode status in `dist/release-manifest.json`. Public artifacts must report `signed: true` and `Valid`. GitHub's protected `release` environment needs `ZADOO_SIGN_PFX_BASE64` and `ZADOO_SIGN_PFX_PASSWORD`. Push an annotated tag matching the package version, then dispatch `windows-x64-release.yml` from that tag. The workflow rejects unsigned output, mismatched versions, and duplicate releases.

## Release-candidate verification

- On a separate Windows x64 machine, verify the installer signature/hash, perform a fresh install, and confirm the startup task and Settings launch.
- Upgrade a configured version 1.0 host. Confirm the access code, permissions, alerts and optional device link remain; no credit panel or sign-in gate may appear.
- Connect through the public link from a second device without linking an account. Verify screen, input, clipboard, snapshots, terminal, camera, microphone and system audio with the intended permissions.
- Open five distinct authenticated viewer sessions. Verify the sixth reports `Viewer limit reached (5)` and disconnected viewers release their slots.
- Make the optional account site unavailable. Existing and new free remote sessions must continue. A failed tunnel process must still produce its exact error and bounded shutdown.
- On iOS/Android verify fit, pan, pinch, zoom, orientation, permission errors and reconnect. Inspect browser console/network and host logs.
- Stop the runtime, verify no orphan process/listener, and test uninstall with both keep/delete-data choices on disposable snapshots.

## Rollback

Keep the previous signed installer. Stop the host, preserve protected settings, install the previous build, and re-check permissions and connectivity. Version 4 removes usage caches; a rollback to a paid-era host may require restoring its earlier settings backup. Do not copy DPAPI secrets between Windows users.

The application is Windows-native; Docker is not an applicable host deployment.

# Zadoo operations

Zadoo is free. It runs on Windows 10/11 x64 and requires a local access code. There is no paid activation, minute balance, or billing setup.

## Setup and run

Follow the source/build commands in `README.md`. Installed builds use `%LOCALAPPDATA%\Zadoo\config.json`; keep that file private. Access codes and optional dashboard tokens are protected by the current Windows user's DPAPI.

Open Settings, set an access code, choose permissions, and select **Start Zadoo**. The public link appears after cloudflared starts. A dashboard account is optional. Source test hosts can use `ZADOO_DISABLE_TUNNEL=1` and `ZADOO_ALLOW_DIRECT_ACCESS=1`; direct LAN access is denied by default.

Silent uninstall (`/SILENT` or `/VERYSILENT`) always keeps settings and logs, including any legacy ProgramData copy, and does not display a data-deletion prompt. Interactive uninstall defaults to keeping data; deletion requires an explicit Yes. The application files, shortcuts and startup task are still removed.

The host uses port 6173. Check local health with:

```powershell
Invoke-RestMethod http://127.0.0.1:6173/api/runtime/status -TimeoutSec 3
```

Check `success`, `running`, `public_url`, `tunnel_error`, and `tunnel_block_reason`. Optional dashboard failure does not gate access. If cloudflared exits, the tunnel monitor reports bounded retries and fails closed after the configured 60–120 seconds.

Stop from Settings. The authenticated RPC returns `Zadoo runtime stopping`. Confirm the process exits and port 6173 is released.

## Logs and support

Logs are under `%LOCALAPPDATA%\Zadoo\logs` and are bounded/rotated. **Export Diagnostics** creates a bounded ZIP and redacts codes, tokens, identifiers, email, cookies, DPAPI data, user paths and tunnel links. Independently inspect the ZIP before sharing.

## Signing and publishing

Zadoo's original code is licensed under Apache-2.0; bundled dependencies retain their own licenses. The build embeds the project license and notice, sets `ProductName=Zadoo` and matching product/file versions, and rejects mismatched Windows metadata.

The initial free SignPath Foundation application was not approved. Zadoo does not currently have an approved trusted code-signing provider. The trusted-signed release workflow still requires a separately obtained trusted PFX; it is not connected to SignPath. Do not add a fake certificate, mark an unsigned manifest as signed, or describe an unsigned preview as a trusted-signed release.

If pursuing the SignPath route again:

1. The owner completes [the official application](https://signpath.org/apply), provides truthful reputation evidence and contact details, and personally reviews the provider agreements. MFA is required on GitHub and SignPath.
2. Wait for eligibility/certificate approval and configure the project's named author, reviewer and release-approver roles. Review dependency licensing and the application's network/privacy disclosures with the provider.
3. Configure provider-approved artifact restrictions for only Zadoo-owned executables, enforcing the product name and the release version. Keep upstream executables and DLLs under their original publishers.
4. Replace the PFX release step with the approved GitHub-hosted build/signing integration: sign the host and portable executables, package the signed host, then sign the installer. Configure required manual approval, scoped secrets and origin verification; do not store credentials in source or chat.
5. Verify signatures, trusted timestamps, packaged contents and installation/upgrade/uninstallation on a clean Windows x64 system before publishing. The public website discovers a valid signed release and matching digest automatically.

See the [SignPath requirements](https://signpath.org/terms) and [GitHub integration documentation](https://docs.signpath.io/trusted-build-systems/github). Provider approval is discretionary; applying does not guarantee free signing.

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

Verify every SHA-256 and Authenticode status in `dist/release-manifest.json`. Trusted-signed release artifacts must report `signed: true` and `Valid`. GitHub's protected `release` environment needs `ZADOO_SIGN_PFX_BASE64` and `ZADOO_SIGN_PFX_PASSWORD`. Push an annotated tag matching the package version, then dispatch `windows-x64-release.yml` from that tag. The workflow rejects unsigned output, mismatched versions, and duplicate releases.

### Approved unsigned preview

The owner explicitly authorized a separate, clearly labelled public preview after the signing application was rejected. [Release `v1.1.5-preview.1`](https://github.com/DIVYANSH-675/zadoo/releases/tag/v1.1.5-preview.1) points to source commit `d0d5b36221abcddc186841e71743f6084099cdf8` and contains `Zadoo-1.1.5-x64-Unsigned-Preview-Setup.exe` (68,666,872 bytes; SHA-256 `8e876885f44f00b0bd0a2072802028dd6b747bd497f0df3a17c4dca5f2bf205b`). This is the locally built and tested installer, not a signed artifact or a claim of CI build provenance.

The preview manifest records `channel: unsigned-preview`, `signed: false` and `authenticode_status: NotSigned`. The website pins this exact prerelease, source commit, size and digest, and uses `/api/download/windows-preview` separately from the unchanged trusted-signed route. A different unsigned binary requires an explicit review and pin update; it is never accepted automatically as a signed fallback.

The website and release notes disclose that Windows may show an unknown-publisher or SmartScreen warning. Do not disable Windows security protections. A checksum confirms file integrity, not publisher identity or software safety.

## Release-candidate verification

The Windows CI workflow builds the immutable 1.1.4 baseline and runs `scripts/installer_e2e_test.ps1` on a disposable GitHub-hosted runner. It verifies a fresh install, same-version repair, upgrade, every installed payload hash, startup task and shortcuts, installed HTTP/auth/assets/shutdown, and both silent uninstall modes with current-user and legacy data preserved. The harness refuses developer PCs, self-hosted runners, existing Zadoo installations/data, and unexpected test paths. It does not replace interactive, hardware, browser or trusted-signature testing.

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

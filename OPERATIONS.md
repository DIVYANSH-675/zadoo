# Zadoo operations guide

This guide covers the Windows x64 host runtime in this repository. The hosted billing and
dashboard service is maintained separately in `zadoo-web`.

## Supported environment

- Windows 10 or Windows 11 x64.
- Python 3.11.9 x64 for source runs and builds.
- Node.js 24.18.0 x64 for JavaScript validation during builds.
- A current Chromium, Edge, Firefox, or Safari browser for viewers.
- Port `6173` available on the host. Direct LAN access is denied by default.

Zadoo is a native Windows host and is not deployed as a Docker container. Screen capture,
Win32 input, DPAPI, audio, camera, ConPTY, startup tasks, and Authenticode require Windows.

## Reproducible source setup

Run these commands from the repository root in PowerShell:

```powershell
py -3.11-64 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --no-cache-dir --require-hashes -r requirements-build.lock
.\.venv\Scripts\python.exe -m pip install --no-cache-dir --require-hashes -r requirements-runtime.lock
.\.venv\Scripts\python.exe -m pip install --no-cache-dir --no-deps --no-build-isolation -e .
Copy-Item .env.example .env
```

Edit `.env` and set a development-only `ZADOO_ACCESS_CODE`. To run without the hosted SaaS
or public tunnel, also set `ZADOO_DISABLE_TUNNEL=1`. Never commit `.env` or a real access code.

```powershell
.\.venv\Scripts\python.exe -m zadoo_vnc
```

Open `http://127.0.0.1:6173` on the host. Set `ZADOO_ALLOW_DIRECT_ACCESS=1` only when a
deliberate development or test client must connect directly rather than through Cloudflare.

## Installed setup

1. Install a signed `Zadoo-<version>-x64-Setup.exe` as an administrator.
2. Upgrades migrate an existing `%ProgramData%\Zadoo\config.json` once into the current
   Windows user's protected profile; keep the legacy copy until the upgrade is verified.
3. Open Zadoo Settings, sign in to the hosted service, and finish device activation.
4. Create the local access code and review every permission before sharing a public link.
5. Start Zadoo from Settings and confirm that a public URL appears.
6. Open the public URL from a second device and verify authentication, screen video, and only
   the permissions intended for that viewer.

Installed configuration is stored in `%LOCALAPPDATA%\Zadoo\config.json`. Access codes and
device tokens use current-user DPAPI and the settings tree has an explicit protected ACL. The
file cannot be decrypted by another Windows account and must not be copied as a backup or
deployment method.

## Health and graceful shutdown

The local runtime status endpoint is restricted to loopback requests and does not require the
viewer session cookie:

```powershell
$status = Invoke-RestMethod http://127.0.0.1:6173/api/runtime/status -TimeoutSec 3
$status | Format-List success,running,port,tunnel_enabled,tunnel_block_reason,tunnel_error,public_url
```

A healthy local process returns `success=True`, `running=True`, and `port=6173`. When tunneling
is enabled, treat a non-empty `tunnel_block_reason`, `tunnel_error`, or missing `public_url` as a
degraded public-access state even though the local process is running.

Stop the runtime with **Stop** in Zadoo Settings so its authenticated local WebSocket RPC can
close capture, media, terminal, tunnel, and heartbeat workers cleanly. Legacy state-changing
HTTP routes intentionally return HTTP 405 with
`State-changing action requires authenticated WebSocket RPC`.

The expected response message is `Zadoo runtime stopping`; port `6173` should then stop
listening. Use Task Manager termination only when the authenticated stop path cannot respond.

## Logs and local monitoring

Runtime logs are under `%LOCALAPPDATA%\Zadoo\logs`. Defaults keep the current log plus two
5 MiB backups for each active day and prune log files older than 14 days. Source-only overrides
are documented in `.env.example`:

- `ZADOO_LOG_LEVEL`
- `ZADOO_LOG_RETENTION_DAYS`
- `ZADOO_LOG_MAX_BYTES`
- `ZADOO_LOG_BACKUP_COUNT`
- `ZADOO_DIAGNOSTIC_MAX_LOG_BYTES`
- `ZADOO_DIAGNOSTIC_MAX_LOG_FILES`

Use **Runtime > Export Diagnostics** in Zadoo Settings before escalating an incident. The ZIP
contains bounded log tails and a system/runtime summary; access codes, tokens, IDs, email
addresses, DPAPI blobs, cookies, bearer values, and public tunnel links are redacted. Never
attach `config.json`, PFX files, or API keys.

## Release build and verification

Unsigned builds are suitable only for local testing:

```powershell
.\scripts\build_windows.ps1 -NoSelfSign
```

Distributable builds require a real code-signing PFX and password supplied only in the process
environment:

```powershell
$credential = Get-Credential -UserName "codesign" -Message "Enter the PFX password; the user name is ignored"
$env:ZADOO_SIGN_PFX_PASSWORD = $credential.GetNetworkCredential().Password
.\scripts\build_windows.ps1 -PfxPath C:\secure\zadoo-codesign.pfx
Remove-Item Env:\ZADOO_SIGN_PFX_PASSWORD
$credential = $null
```

For a public GitHub release, configure the protected `release` environment with required
reviewers and these repository/environment secrets:

- `ZADOO_SIGN_PFX_BASE64`: base64 of the code-signing PFX.
- `ZADOO_SIGN_PFX_PASSWORD`: the PFX password.

Create and push an annotated tag exactly matching `v<zadoo_vnc.__version__>`, then dispatch
**Publish signed Windows x64 release** from that tag:

```powershell
$version = .\.venv\Scripts\python.exe -c "from zadoo_vnc import __version__; print(__version__)"
git tag -a "v$version" -m "Zadoo $version"
git push origin "v$version"
gh workflow run windows-x64-release.yml --ref "v$version"
```

The workflow refuses branches, version-mismatched tags, missing/invalid secrets, unsigned
artifacts, invalid signatures, hash mismatches, and duplicate releases.

The build starts from a clean `dist` directory. Before release, verify:

```powershell
$manifest = Get-Content -Raw dist\release-manifest.json | ConvertFrom-Json
if (-not $manifest.signed) { throw "Release artifacts are not signed" }
foreach ($artifact in $manifest.artifacts) {
    $path = Join-Path dist ($artifact.file -replace "/", "\")
    $hash = (Get-FileHash -Algorithm SHA256 -LiteralPath $path).Hash.ToLowerInvariant()
    if ($hash -ne $artifact.sha256) { throw "SHA-256 mismatch: $path" }
    if ((Get-AuthenticodeSignature -LiteralPath $path).Status -ne "Valid") {
        throw "Invalid Authenticode signature: $path"
    }
}
```

`SHA256SUMS.txt` is the human-readable checksum list. GitHub CI also generates
`zadoo-runtime.sbom.cdx.json`; CI artifacts are explicitly unsigned and must not be published as
a release.

## Rollback

1. Keep the previous signed installer and its verified checksum before rollout.
2. Gracefully stop Zadoo.
3. Uninstall the current build. Preserve `%LOCALAPPDATA%\Zadoo` and the legacy ProgramData copy
   unless the configuration itself is known to be corrupt or compromised.
4. Install the previous signed build.
5. Open Settings and verify activation, permissions, runtime status, and the public URL.
6. Perform a second-device viewer smoke test.

If a configuration schema migration prevents rollback, restore through Zadoo Settings rather
than editing encrypted JSON by hand.

## Release-candidate verification runbook

Run this final pass on a clean Windows 10 or 11 x64 machine or disposable VM. Use signed release
artifacts and staging/test accounts; do not use the unsigned CI or local-development artifacts.

### Fresh install and upgrade

1. Copy the installer, `release-manifest.json`, `SHA256SUMS.txt`, and SBOM to the test machine.
2. Run the manifest/hash/signature verification above, then start the installer and accept the
   Windows elevation prompt:

   ```powershell
   $installer = Get-Item .\Zadoo-*-x64-Setup.exe
   Start-Process -FilePath $installer.FullName -Verb RunAs -Wait
   Get-AuthenticodeSignature 'C:\Program Files\Zadoo\Zadoo.exe' |
       Format-List Status,SignerCertificate
   Get-ScheduledTask -TaskName Zadoo | Format-List TaskName,State
   ```

3. Configure a staging device in Settings, start the runtime, run the health command above, and
   stop it from Settings. Confirm the process exits and `Test-NetConnection 127.0.0.1 -Port 6173`
   reports `TcpTestSucceeded=False` after shutdown.
4. Install the previous signed release, configure it, and record the visible Settings values.
   Run the candidate installer over that release. Confirm the same activation and permission
   choices remain visible, the access code still authenticates, and the legacy ProgramData copy
   remains available for rollback when a migration occurred.
5. Uninstall the candidate. Confirm the scheduled task and Program Files directory are removed.
   Exercise both answers to the uninstall data-retention prompt on separate disposable snapshots.

### Public service and multi-viewer flow

1. Use a staging Zadoo account and the payment provider's test mode to complete sign-in,
   activation, checkout, webhook processing, entitlement refresh, and billing-history display.
2. Start the installed host with tunneling enabled. The local health endpoint must show a public
   URL and empty tunnel error fields. Open that URL over mobile data or another external network;
   direct LAN access is intentionally denied by default.
3. Authenticate two through five distinct browser sessions. Verify each can reconnect without
   consuming a second slot, then confirm a sixth distinct session receives exactly
   `Viewer limit reached (5)`. Close viewers and confirm their slots become available.
4. In a staging-only outage window, make the hosted heartbeat endpoint unreachable. Verify the
   5/10/20/30-second retries, recovery when connectivity returns before the configured grace
   deadline, and fail-closed shutdown at the deadline with
   `Cloud heartbeat failed for 120 seconds: ...` at the default setting.

### Physical media and mobile flow

1. Use a host with a real camera, microphone, speakers, and two displays if multi-monitor support
   is required. Grant only the permissions under test and keep private material off screen.
2. From a second device, verify screen video, mouse/keyboard input, text and image clipboard,
   snapshot, terminal, camera video, selected-microphone audio, and system-audio loopback. Keep
   each media stream open for at least 60 seconds, switch devices once, and confirm disconnecting
   the final viewer stops the corresponding capture worker without log errors.
3. On current iOS Safari and Android Chrome, exercise fit-width, fit, pan, pinch zoom, reset,
   orientation changes, the on-screen controls, reconnect, and permission-denied error states.
   Confirm remote pointer coordinates still match after every transform.
4. Inspect the browser console/network panel and `%LOCALAPPDATA%\Zadoo\logs`; there must be no
   unexpected external assets, unhandled exceptions, credential values, or orphaned workers.
   Export diagnostics and independently inspect the ZIP for redaction before approval.

## Production checklist

- [ ] Pull request checks pass on the exact release commit.
- [ ] Runtime and build dependency audits report no known vulnerabilities.
- [ ] Installer and portable executable are x64 and Authenticode status is `Valid`.
- [ ] Manifest and independent SHA-256 checks agree.
- [ ] CycloneDX runtime SBOM is retained with the release.
- [ ] `.env`, access codes, device tokens, PFX files, and passwords are absent from Git.
- [ ] Fresh install and in-place upgrade both pass on supported Windows x64.
- [ ] Local status, public tunnel, authentication, permissions, terminal, and graceful stop pass.
- [ ] Two-to-five viewer admission, the sixth-viewer rejection, and heartbeat fail-closed behavior pass.
- [ ] A diagnostic export opens successfully and contains no known secrets or personal identifiers.
- [ ] Browser console and network checks show no unexpected errors or external assets.
- [ ] Previous signed installer and rollback procedure are available.

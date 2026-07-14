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
2. Leave the ProgramData deletion prompt unanswered during upgrades; it appears only when
   uninstalling.
3. Open Zadoo Settings, sign in to the hosted service, and finish device activation.
4. Create the local access code and review every permission before sharing a public link.
5. Start Zadoo from Settings and confirm that a public URL appears.
6. Open the public URL from a second device and verify authentication, screen video, and only
   the permissions intended for that viewer.

Installed configuration is stored in `%ProgramData%\Zadoo\config.json`. It contains DPAPI
encrypted values and must not be copied to another machine as a backup or deployment method.

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

Stop the runtime through its authenticated local endpoint so capture, media, terminal, tunnel,
and heartbeat workers can close cleanly:

```powershell
$code = Read-Host "Local Zadoo access code"
Invoke-RestMethod http://127.0.0.1:6173/api/runtime/stop -Headers @{"X-Zadoo-Code" = $code} -TimeoutSec 10
```

The expected response message is `Zadoo runtime stopping`; port `6173` should then stop
listening. Use Task Manager termination only when the authenticated stop path cannot respond.

## Logs and local monitoring

Runtime logs are under `%ProgramData%\Zadoo\logs`. Defaults keep the current log plus two
5 MiB backups for each active day and prune log files older than 14 days. Source-only overrides
are documented in `.env.example`:

- `ZADOO_LOG_LEVEL`
- `ZADOO_LOG_RETENTION_DAYS`
- `ZADOO_LOG_MAX_BYTES`
- `ZADOO_LOG_BACKUP_COUNT`

Before escalating an incident, preserve the relevant bounded logs, the application version,
Windows version, status payload with access codes removed, and exact reproduction time. Do not
attach `config.json`, device tokens, access codes, PFX files, or API keys.

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
3. Uninstall the current build. Choose **No** when asked whether to delete ProgramData settings
   and logs unless the configuration itself is known to be corrupt or compromised.
4. Install the previous signed build.
5. Open Settings and verify activation, permissions, runtime status, and the public URL.
6. Perform a second-device viewer smoke test.

If a configuration schema migration prevents rollback, restore through Zadoo Settings rather
than editing encrypted JSON by hand.

## Production checklist

- [ ] Pull request checks pass on the exact release commit.
- [ ] Runtime and build dependency audits report no known vulnerabilities.
- [ ] Installer and portable executable are x64 and Authenticode status is `Valid`.
- [ ] Manifest and independent SHA-256 checks agree.
- [ ] CycloneDX runtime SBOM is retained with the release.
- [ ] `.env`, access codes, device tokens, PFX files, and passwords are absent from Git.
- [ ] Fresh install and in-place upgrade both pass on supported Windows x64.
- [ ] Local status, public tunnel, authentication, permissions, terminal, and graceful stop pass.
- [ ] Browser console and network checks show no unexpected errors or external assets.
- [ ] Previous signed installer and rollback procedure are available.

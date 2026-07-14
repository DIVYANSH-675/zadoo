# Security policy

## Supported versions

Security fixes are applied to the current `1.x` release line. Older builds are unsupported.

## Reporting a vulnerability

Do not open a public issue for a suspected vulnerability. Use the repository's private
[security advisory form](https://github.com/DIVYANSH-675/zadoo/security/advisories/new) and include:

- the affected version and Windows version;
- reproducible steps or a minimal proof of concept;
- the expected and observed behavior;
- the security impact; and
- any suggested mitigation.

Do not include live access codes, API keys, PFX files, or user data. Rotate any credential that
may have been exposed while reproducing the issue.

## Runtime safeguards

- Installed settings live under the current Windows user's `%LOCALAPPDATA%\Zadoo` tree with a
  protected ACL. Access codes and device tokens use current-user Windows DPAPI.
- Browser state mutations use an authenticated same-origin WebSocket RPC. The desktop runtime
  administration actions additionally require a loopback connection and the local access code.
- Viewer authentication is rate limited, permissions are enforced per authenticated session,
  and at most five simultaneous viewer sessions are admitted by default.
- Hosted-service heartbeat failures retry with bounded backoff for up to 120 seconds, then the
  runtime fails closed.
- **Export Diagnostics** bounds included log tails and redacts access codes, tokens, IDs, email,
  cookies, bearer values, DPAPI blobs, and public tunnel URLs.
- Public releases must be x64, hash-verified, SBOM-backed, and Authenticode signed. The manual
  release workflow refuses to publish unsigned or invalidly signed artifacts.

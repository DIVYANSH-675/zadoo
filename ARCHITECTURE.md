# Zadoo architecture

Zadoo is a Windows x64 host with a browser viewer. It is free and has no payment provider, minute counter, purchase flow, or paid entitlement gate.

## Host lifecycle

1. Load and validate protected per-user settings.
2. Require a local access code and initialize capture.
3. Bind the HTTP/WebSocket server on port 6173.
4. Start the signed bundled cloudflared tunnel unless explicitly disabled for local testing.
5. Authenticate viewers and enforce per-feature permissions and the five-viewer resource cap.
6. Stop through authenticated local RPC; close media, terminal, capture, and tunnel workers.

The optional device dashboard receives presence and public-link updates only when a device has been linked. A failed dashboard request is logged and does not stop the host. Tunnel-process failures retain bounded retry/failure reporting independent of account state.

## Components

| Module | Responsibility |
|---|---|
| `app.py` | Windows/x64 checks, configuration, capture and server lifecycle |
| `server.py` | WebSocket lifecycle, viewer resource cap, tunnel monitoring, optional device presence |
| `routes.py` | Authentication, permissions, assets, snapshots, local administration and RPC |
| `media.py` | Screen/audio/camera/terminal streams and bounded media queues |
| `streaming.py` | Adaptive capture profile |
| `settings.py` | Versioned settings, migration, DPAPI, ACLs and atomic updates |
| `settings_window.py` | Native access, permissions, alerts, runtime and optional dashboard UI |
| `saas.py` | Optional device linking and presence client |
| `tunnel.py` | Verified cloudflared process and optional link email |
| `diagnostics.py` | Bounded redacted local support exports |

## Free-version upgrade

Settings version 4 migrates version 3 without changing access codes, permissions, alerts, device links, or encrypted secrets. Obsolete entitlement and credit caches are discarded. Browser payment panels, credit polling, usage locks and session-metering calls are removed.

The separate website retains optional Google sign-in and device linking. Legacy host API responses allow free access; payment endpoints return HTTP 410 and never call a provider. Historical database rows are retained without a destructive migration.

## Performance and transport

Templates and static text assets use cached gzip; images are optimized; CodeMirror is loaded on demand. Media queues drop old frames to bound latency. Terminal creation runs off the event loop. Windows PowerShell inherits only compatible module paths.

Screen, input, clipboard, terminal and media traffic pass through the host's Cloudflare tunnel. The account website does not relay those streams. HTTPS transport does not imply application-level end-to-end encryption.

## Build and release

Hash-locked dependencies, verified build tools, full-SHA GitHub actions, x64 PE checks, SHA-256 manifests, dependency audits and an SBOM are enforced. Public release automation requires a matching version tag and valid real Authenticode signatures.

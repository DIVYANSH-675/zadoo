# Zadoo — Full Architecture

Zadoo is a **Windows remote-screen / remote-control product** with a usage-metered SaaS billing backend. A host PC runs a native **Settings app + streaming runtime**; viewers open a **public Cloudflare-tunnel link** in a browser to watch the host screen, control it, hear its audio, use a terminal/webcam, and **buy session credits** via Razorpay. Billing, accounts, device linking and entitlements live in a separate **Next.js cloud app** on Vercel + Postgres.

This document describes the **entire** design: topology, components, data model, identity/auth, every request path, the streaming and audio pipelines, the payment flow, configuration, build/deploy, and known limitations.

---

## 0. Glossary

| Term | Meaning |
|---|---|
| **Runtime** | The Python streaming server (`zadoo_vnc`), packaged as `Zadoo.exe`. Binds `localhost:6173`. |
| **Settings app** | The native Tkinter window (`settings_window.py`) that signs in, starts/stops the runtime, and shows the public link. Same `Zadoo.exe`, launched with `--settings`. |
| **Cloud** | The Next.js app (`web/`) at `https://zadoo-web.vercel.app`. Handles accounts, devices, billing. |
| **Public link** | A `*.trycloudflare.com` URL created by `cloudflared` that tunnels to `localhost:6173`. |
| **Viewer** | A browser opening the public link to watch/control the host. |
| **Device token** | An opaque bearer token minted by the cloud when a device is signed in; stored DPAPI-encrypted on the host; used for all `agent/*` cloud calls. |
| **Access code** | A short host-set password viewers type to authenticate to the *local* runtime (separate from the cloud account). |
| **Entitlement / wallet** | Cloud-side metering: included plan minutes + pay-as-you-go wallet minutes. |

---

## 1. Two repositories / two deployables

```mermaid
flowchart LR
  subgraph R1["zadoo.git — Windows runtime (Python)"]
    A1["zadoo_vnc package"]
    A2["Zadoo.exe (PyInstaller)"]
    A3["Zadoo-Setup.exe (Inno Setup)"]
    A1 --> A2 --> A3
  end
  subgraph R2["zadoo-web.git — Cloud (Next.js)"]
    B1["web/ app + lib + prisma"]
    B2["Vercel deployment"]
    B3["Neon/Postgres"]
    B1 --> B2 --> B3
  end
  A2 -. "agent/* REST (device token)" .-> B2
```

* **`zadoo.git`** builds the host installer. Source: `zadoo_vnc/`, `zadoo_vnc_single.py` (entry), `scripts/build_windows.ps1`, `installer/zadoo.iss`.
* **`zadoo-web.git`** is `web/`: Next.js 15 App Router, NextAuth v5, Prisma, Razorpay. Auto-deploys to Vercel on push.

---

## 2. Top-level topology

```mermaid
flowchart TB
  subgraph HOST["Host Windows PC"]
    SW["Settings App (Tkinter)<br/>settings_window.py"]
    RT["Runtime / VNCServer<br/>localhost:6173 (websockets)"]
    CAP["Screen capture<br/>(dxcam/bettercam/mss)"]
    AUD["System audio + Mic<br/>(soundcard/sounddevice)"]
    INP["Input inject<br/>(win32 mouse/kbd/clipboard)"]
    CF["cloudflared.exe<br/>(quick tunnel)"]
    CFG["C:\\ProgramData\\Zadoo\\config.json<br/>(DPAPI-encrypted token)"]
    SW --> RT
    SW <--> CFG
    RT --> CAP & AUD & INP
    RT --> CF
  end

  subgraph EDGE["Cloudflare edge"]
    TUN["*.trycloudflare.com"]
  end

  subgraph CLOUD["Vercel — zadoo-web"]
    NX["Next.js API + pages"]
    PG[("Postgres")]
    NX --> PG
  end

  RZP["Razorpay"]
  V["Viewer browser<br/>(index.html)"]

  CF <-->|HTTP/2 tunnel| TUN
  V <-->|HTTPS / WSS| TUN
  TUN <-->|to localhost:6173| RT
  SW -->|"activate / poll (device token)"| NX
  RT -->|"agent/* proxy (Bearer device token)"| NX
  V -->|"Razorpay checkout.js"| RZP
  RZP -->|"signed webhook"| NX
```

Key point: **viewers never talk to the cloud directly.** The payment panel calls the **local** runtime (`/api/local/*`), which **proxies** to the cloud using the host's device token. The cloud talks to Razorpay; Razorpay's checkout widget runs in the viewer's browser.

---

## 3. Identity & auth — three independent layers

```mermaid
flowchart TD
  subgraph L1["1) Cloud user identity (NextAuth)"]
    U["Google OAuth / email OTP"] --> US["Session cookie<br/>requireUser()"]
  end
  subgraph L2["2) Device identity (agent)"]
    D["Activation handshake"] --> DT["DeviceToken (Bearer)<br/>requireAgent()"]
  end
  subgraph L3["3) Viewer ↔ runtime (local)"]
    AC["Access code"] --> CK["Local auth cookie + role<br/>(handle_auth)"]
  end
```

1. **User identity** — `web/app/api/auth/[...nextauth]` (Google + email OTP). `requireUser()` (`web/lib/auth.ts`) gates dashboard/device/checkout routes. NextAuth tables: `User`, `Account`, `Session`, `VerificationToken`.
2. **Device identity** — the host proves itself with a **DeviceToken** bearer (`requireAgent()` in `web/lib/agent.ts`): `tokenHash = hmacToken(token)` looked up in `DeviceToken`; rejected if revoked or `device.status==REVOKED`. Used for all `/api/agent/*` calls.
3. **Viewer → runtime** — the local server issues its own **role + cookie** when a viewer POSTs the host's **access code** to `/api/auth` (`handle_auth`). Roles map to a **permission matrix** (which features are allowed). This is entirely local; the cloud is not involved.

---

## 4. Runtime component architecture (`zadoo_vnc`)

```mermaid
flowchart TB
  ENTRY["zadoo_vnc_single.py<br/>(--settings | --open | run)"]
  ENTRY --> APP["app.py — orchestration"]
  ENTRY --> SW["settings_window.py — Tkinter UI"]

  APP --> SRV["server.py — VNCServer<br/>(websockets.serve, queues, clients)"]
  SRV -. mixin .-> ROUTES["routes.py — RoutesMixin<br/>process_request gate + dispatch + cloud proxy"]
  SRV -. mixin .-> MEDIA["media.py — MediaMixin<br/>audio/mic/webcam/terminal/video broadcast"]
  SRV -. mixin .-> INPUT["input_control.py — InputControlMixin<br/>mouse/kbd/clipboard"]

  SRV --> CAPM["screen_capture.py — ScreenCapturer<br/>dxcam ring buffer / bettercam / mss → JPEG"]
  SRV --> STR["streaming.py — AdaptiveStreamController<br/>STREAM_LADDER, observe_server, apply_to"]
  APP --> TUN["tunnel.py — CloudflareTunnelManager"]
  SW --> SAAS["saas.py — ZadooCloudClient<br/>activate/poll/entitlement/credits/heartbeat"]
  SW --> WST["windows_startup.py — schtasks task"]

  subgraph SHARED["Shared services"]
    SET["settings.py — SettingsStore<br/>config.json + DPAPI + atomic_update"]
    CFGP["config.py — env/paths"]
    NET["network.py — get_local_ip"]
    DPI["dpi.py / win32_input.py"]
    DEP["dependencies.py — optional import probes"]
    LOG["logging_utils.py"]
    CAMD["camera_discovery.py"]
  end
  APP --> SET
  ROUTES --> SET
  SAAS --> SET
  MEDIA --> CAMD
  INPUT --> DPI

  subgraph TPL["templates/"]
    IDX["index.html — viewer SPA<br/>(canvas, AudioWorklet, payment panel, terminal)"]
    TRM["terminal.html"]
    BM["benchmark.html"]
    HC["host_controls.html"]
  end
  ROUTES --> IDX
```

**Module responsibilities**

| Module | Role |
|---|---|
| `zadoo_vnc_single.py` | Entry shim. Dispatches `--settings` / `--open` / default run. |
| `app.py` | Startup orchestration: free port 6173, decide tunnel (signed-in + not revoked + not disabled), build `VNCServer`, `ScreenCapturer`, `CloudflareTunnelManager`, run `asyncio` server. |
| `server.py` | `VNCServer`: `websockets.serve`, holds `current_fps/current_quality`, the `AdaptiveStreamController`, client sets (`video/audio/mic/input/cursor`), realtime queues (`audio_queue`, `mic_queue`), `frame_ready_event`. Composed of the three mixins. |
| `routes.py` | `RoutesMixin`: `process_request` security gate, full HTTP/WS route dispatch, `/api/local/*` → cloud **proxy** (`_proxy_cloud_get/post`), local auth (`handle_auth`, access codes), feature/permission map. |
| `media.py` | `MediaMixin`: system-audio loopback + mic capture, `/audio` `/mic` send loops, webcam (`/webcam`), terminal PTY (`/ssh`, `winpty`), **video broadcast loop**, `_clarify_mono_audio` AGC. |
| `input_control.py` | `InputControlMixin`: `/input` WS → mouse/keyboard/clipboard via win32, cursor broadcast. |
| `screen_capture.py` | `ScreenCapturer`: backend select (dxcam ring buffer / bettercam / mss / fast-ctypes), scale/grayscale, JPEG encode (`imagecodecs`/opencv/Pillow). |
| `streaming.py` | Adaptive bitrate controller: `STREAM_LADDER`, `observe_server` (latency/backlog/bitrate heuristics), `apply_to` (push fps/quality/scale to capturer). |
| `tunnel.py` | `CloudflareTunnelManager`: spawn `cloudflared tunnel`, parse public URL, signature-verify (cached), optional Resend email. |
| `saas.py` | `ZadooCloudClient`: HTTP client to the cloud `agent/*` API (activation, poll, entitlement, credits, profile, heartbeat, go-offline). |
| `settings.py` | `SettingsStore`: `C:\ProgramData\Zadoo\config.json`, DPAPI encrypt for secrets, `atomic_update`, normalized defaults, access-code hashing. |
| `settings_window.py` | `ZadooSettingsWindow`: tabs (Access/Account/Permissions/Alerts/Runtime), autosave, sign-in/activation, Start/Stop runtime, public-link display, credits, DPI-aware sizing. |
| `windows_startup.py` | Create/remove the `schtasks` logon task. |
| `dpi.py`, `win32_input.py` | DPI awareness + raw win32 input. |
| `config.py`, `network.py`, `dependencies.py`, `logging_utils.py`, `camera_discovery.py`, `assets.py`, `process_utils.py` | Env/paths, local IP, optional-dependency probes, logging, camera enumeration, asset paths, process helpers. |

---

## 5. Cloud component architecture (`web/`)

```mermaid
flowchart TB
  subgraph PAGES["App pages (App Router)"]
    LAND["/ landing"] 
    PRICE["/pricing"]
    LOGIN["/login"]
    DASH["/dashboard (+ billing, devices)"]
    ONB["/onboarding"]
    CONN["/connect-device"]
    DL["/download"]
    ENT["/enterprise"]
  end

  subgraph API["app/api"]
    AUTH["auth/[...nextauth]"]
    subgraph AGENT["agent/* (device-token)"]
      ACT["activate/start, activate/poll"]
      CRED["credits, entitlement, profile, heartbeat"]
      SESS["session/start, session/heartbeat, session/end"]
      WAL["wallet/topup-order, wallet/topup-verify"]
    end
    DEV["devices/claim, devices/remove (user)"]
    CHK["checkout, checkout/verify (user)"]
    WH["webhooks/razorpay"]
    PROF["profile"]
    ENTAPI["enterprise"]
  end

  subgraph LIB["lib/"]
    L_AUTH["auth.ts — requireUser"]
    L_AGENT["agent.ts — requireAgent (DeviceToken)"]
    L_BILL["billing.ts — entitlementDecision, debitSessionUsage"]
    L_PAY["payments.ts — settleRazorpayPayment, verify signature"]
    L_PRICE["pricing.ts — plans, walletMinutesForAmount, markets"]
    L_SEC["security.ts — hmacToken, json, timingSafeEqualText"]
    L_REG["region.ts — marketFromCountry"]
    L_DB["db.ts — Prisma client"]
  end

  PG[("Postgres (Prisma)")]
  RZP["Razorpay API"]

  AGENT --> L_AGENT --> L_DB --> PG
  WAL --> L_PAY --> RZP
  WH --> L_PAY
  SESS --> L_BILL --> L_DB
  CHK --> L_PAY
  DEV --> L_AUTH
  CHK --> L_AUTH
  AUTH --> L_DB
```

---

## 6. Data model (Prisma / Postgres)

```mermaid
erDiagram
  User ||--o{ Account : has
  User ||--o{ Session : has
  User ||--o{ Membership : member
  Workspace ||--o{ Membership : has
  Workspace ||--o{ Device : owns
  Workspace ||--o{ ActivationCode : issues
  Workspace ||--o{ Entitlement : has
  Workspace ||--o{ WalletLedger : has
  Workspace ||--o{ RemoteSession : has
  Workspace ||--o{ Payment : has
  Workspace ||--o{ EnterpriseLead : has
  Device ||--o{ DeviceToken : has
  Device ||--o{ ActivationCode : claims
  Device ||--o{ RemoteSession : runs
  Entitlement ||--o{ RemoteSession : meters
  RemoteSession ||--o{ UsageLedger : logs

  User {
    string id PK
    string email UK
  }
  Workspace {
    string id PK
    Market market
    Currency currency
  }
  Membership {
    string id PK
    WorkspaceRole role
  }
  Device {
    string id PK
    DeviceStatus status
    string publicUrl
  }
  DeviceToken {
    string id PK
    string tokenHash UK
    datetime revokedAt
  }
  ActivationCode {
    string id PK
    string codeHash UK
    ActivationStatus status
    datetime expiresAt
  }
  Entitlement {
    string id PK
    PlanCode planCode
    int includedMinutesTotal
    int includedMinutesUsed
    int concurrencyLimit
  }
  WalletLedger {
    string id PK
    LedgerKind kind
    int minutes
    int amountMinor
  }
  RemoteSession {
    string id PK
    SessionStatus status
    int includedMinutesUsed
    int walletMinutesUsed
  }
  UsageLedger {
    string id PK
    int minutes
    int includedMinutes
    int walletMinutes
  }
  Payment {
    string id PK
    PaymentStatus status
    string providerOrderId
    string providerPaymentId
    int amountMinor
  }
  WebhookEvent {
    string id PK
    string eventId
    string payloadHash
  }
  EnterpriseLead {
    string id PK
    string email
  }
```

**Enums:** `WorkspaceRole(OWNER/ADMIN/MEMBER)`, `Market(INDIA/GLOBAL)`, `Currency(INR/USD)`, `PaymentProvider(RAZORPAY/MANUAL)`, `PaymentStatus(PENDING/PAID/FAILED/CANCELED/REFUNDED)`, `PlanCode(PAYG/TRIAL_15/PASS_14D/MONTHLY/YEARLY/ENTERPRISE)`, `EntitlementStatus(ACTIVE/EXPIRED/REVOKED/GRACE)`, `DeviceStatus(PENDING/ACTIVE/REVOKED)`, `ActivationStatus(PENDING/CLAIMED/EXPIRED)`, `SessionStatus(ACTIVE/ENDED/BLOCKED)`, `LedgerKind(CREDIT/DEBIT/ADJUSTMENT)`.

Notable constraints: `DeviceToken.tokenHash @unique`, `Payment @@unique([provider, providerPaymentId])`, `WebhookEvent @@unique([provider, eventId])`.

---

## 7. Local route & permission map (runtime)

`routes.py` maps each path to a **feature**; `handle_auth` grants a **role** whose **permission set** decides which features the viewer may use. WebSocket upgrades are auth-checked before the stream handler runs; state-changing HTTP routes in `CSRF_HTTP_FEATURES` require a CSRF token.

| Path | Feature | Notes |
|---|---|---|
| `/`, `/api/auth`, `*.png`, `/api/settings/status` | `public` | always reachable |
| `/video`, `/input`, `/api/public-url`, `/api/stream-stats`, `/benchmark.html` | `view` | viewing/telemetry |
| `/audio` | `system_audio` | system-audio stream |
| `/mic`, `/api/list-mics` | `mic` | host mic stream |
| `/snapshot` | `snapshots` | single frame |
| `/api/set-quality`, `/api/set-fps` | `advanced_video` | CSRF |
| `/api/refresh-tunnel` | `tunnel_refresh` | CSRF |
| `/terminal.html`, `/ssh` | `terminal` | PTY |
| `/webcam`, `/api/list-cameras` | `camera` | webcam |
| `/host-controls`, `/api/alert` | `remote_alerts` | CSRF |
| `/api/local/credits|profile|topup-order|topup-verify` | `billing` | **proxied to cloud** |
| `/api/settings/*`, `/api/runtime/*` | `settings` | admin (Settings window), CSRF; exempt from public-only gate |

### Request gate (every request)

```mermaid
flowchart TD
  R["process_request(path, headers)"] --> O{"origin allowed?"}
  O -- no --> F1["/api/* → JSON 403 · else text 403"]
  O -- yes --> ADM{"admin route?<br/>/api/runtime|settings"}
  ADM -- no --> PUB{"tunnel OR localhost OR<br/>ZADOO_ALLOW_DIRECT_ACCESS?"}
  PUB -- no --> F2["/api/* → JSON 'use public link' · else HTML page"]
  PUB -- yes --> WS{"WebSocket upgrade?"}
  ADM -- yes --> WS
  WS -- yes --> WSA{"_is_ws_authorized?"}
  WSA -- no --> F3["403"]
  WSA -- yes --> PASS["handler runs"]
  WS -- no --> FEAT{"feature auth ok?"}
  FEAT -- no --> F4["/api/* → JSON 403 · else text 403"]
  FEAT -- yes --> CSRF{"CSRF required & valid?"}
  CSRF -- no --> F5["403"]
  CSRF -- yes --> DISP["route dispatch"]
```

> **Design fix baked in:** every `/api/*` failure returns **JSON** (not the HTML "open via public link" page), so the browser's `resp.json()` never throws `Unexpected token '<'`. Unmatched `/api/*` → JSON 404.

---

## 8. Sequence — device activation & sign-in

```mermaid
sequenceDiagram
  participant SW as Settings App
  participant Cloud
  participant Browser as Host browser
  participant Store as config.json (DPAPI)

  SW->>Cloud: POST /api/agent/activate/start (device name, machine id)
  Cloud-->>SW: { code, pollSecret, connectUrl }
  SW->>Browser: open connectUrl
  Browser->>Cloud: sign in (NextAuth) + POST /api/devices/claim (requireUser)
  Cloud->>Cloud: link Device→Workspace, ActivationCode=CLAIMED
  loop every ~2.5s until claimed/expired
    SW->>Cloud: POST /api/agent/activate/poll (pollSecret)
    Cloud-->>SW: pending | { status: claimed, deviceToken, workspaceId, deviceId }
  end
  SW->>Store: store DPAPI(deviceToken) + workspace/device ids
  SW->>Cloud: fetch profile / credits / entitlement
```

The minted **DeviceToken** is the credential for every subsequent `agent/*` call (heartbeat, credits, top-up proxy, session).

---

## 9. Sequence — start runtime, tunnel, public link

```mermaid
sequenceDiagram
  participant SW as Settings App
  participant Runtime
  participant Cloudflared
  participant Edge as Cloudflare

  SW->>SW: balance>0? signed-in? (else block/prompt)
  SW->>Runtime: launch Zadoo.exe (runtime) [after graceful stop if already running]
  Runtime->>Runtime: bind localhost:6173 (websockets.serve)
  Runtime->>Cloudflared: spawn `cloudflared tunnel --no-autoupdate --protocol http2 --url http://localhost:6173`
  Note over Cloudflared: HTTP/2 avoids QUIC/UDP-7844 backoff on restrictive nets;<br/>signature verified once (cached); stale cleanup in background
  Cloudflared-->>Runtime: parse "*.trycloudflare.com" from stdout
  Runtime-->>SW: public_url (polled into the Settings link)
  SW->>SW: show clickable public link
```

---

## 10. Sequence — viewer connects & authenticates

```mermaid
sequenceDiagram
  participant V as Viewer browser
  participant Edge as Cloudflare
  participant RT as Runtime

  V->>Edge: GET public link
  Edge->>RT: GET / (CF-Connecting-IP present ⇒ tunnel request)
  RT-->>V: index.html (SPA)
  V->>RT: POST /api/auth?code=ACCESS_CODE
  RT->>RT: _match_auth_code → role + permission matrix
  RT-->>V: Set-Cookie auth token + role; CSRF token
  V->>RT: open WS /video /audio /input (auth-checked) per granted features
```

---

## 11. Sequence — video streaming + adaptive control

```mermaid
sequenceDiagram
  participant CAP as ScreenCapturer (ring buffer)
  participant BC as Video broadcast loop (media.py)
  participant ASC as AdaptiveStreamController
  participant V as Viewer canvas

  loop per captured frame
    CAP->>CAP: grab → scale/grayscale → JPEG encode
    CAP->>BC: frame_ready_event.set()
  end
  loop broadcast
    BC->>BC: per client: skip if write-buffer>256KB or send in-flight
    BC->>V: WS binary JPEG frame
  end
  V-->>ASC: client stats (display_fps, decode_ms, recv_kbps, rtt, dropped)
  BC-->>ASC: server stats (frame_bytes, max_write_buffer, skipped, frame_age)
  loop every ~1s
    ASC->>ASC: observe_server() → up/down-shift profile or fps cap
    ASC->>CAP: apply_to() → set fps / quality / scale_div / grayscale
  end
```

### Adaptive ladder (`STREAM_LADDER`, top = highest bandwidth)

| # | Profile | target FPS | JPEG quality | scale_div |
|---|---|---|---|---|
| 0 | 1080p240 | 240 | 52 | 1 |
| 1 | 720p240 | 240 | 54 | 2 |
| 2 | 1080p120 | 120 | 65 | 1 |
| 3 | 900p120 | 120 | 56 | 1 |
| 4 | 720p120 | 120 | 54 | 2 |
| 5 | 720p60 | 60 | 60 | 2 |
| 6 | **540p60 (default start)** | 60 | 52 | 2 |
| 7 | 540p30 | 30 | 54 | 2 |
| 8 | 540p15 | 15 | 50 | 2 |
| 9 | 360p30 | 30 | 46 | 3 |
| 10 | 360p15 | 15 | 42 | 3 |
| 11 | 360p10 | 10 | 40 | 3 |

```mermaid
stateDiagram-v2
  [*] --> Start540p60
  Start540p60 --> Lower : overloaded (backlog / bitrate-over / stale / slow) — faster + multi-rung when severe
  Lower --> Lower : still overloaded
  Lower --> Higher : 6 good intervals & ≥8s & under bitrate target
  Higher --> Start540p60 : reaches top of needed tier
  note right of Lower
    MJPEG has no inter-frame compression.
    Resolution ladder is the real bandwidth lever.
    Quality lock is OFF by default so the ladder works.
    Optional ZADOO_TARGET_KBPS for a hard cap (e.g. 2000).
  end note
```

* The default boots **unlocked** (`_quality_locked_by_user=False`) so the **resolution ladder** is active; the lock turns on only if the viewer manually sets a quality.
* Optional **WebRTC turbo path** (FFmpeg → MediaMTX → H.264) exists conceptually but is **not wired in this build** (`webrtc_configured=False`), so the live path is **MJPEG-over-WebSocket**.

---

## 12. Sequence — audio (system + mic) pipeline

```mermaid
sequenceDiagram
  participant SC as soundcard loopback / sounddevice mic
  participant Q as audio_queue / mic_queue (maxsize 5 ≈100ms, drop-oldest)
  participant SL as send loop (drop frame if socket buffer >16KB)
  participant WS as /audio or /mic WebSocket
  participant WK as AudioWorklet "pcm-player"
  participant SPK as Speakers

  SC->>SC: 48kHz, 20ms frames → mono s16le PCM (+ AGC clarity)
  SC->>Q: put_nowait (drop oldest on Full → bounded latency)
  loop
    Q->>SL: get frame
    SL->>WS: send (unless write-buffer backed up → drop)
  end
  WS->>WK: ArrayBuffer (transferred, zero-copy)
  WK->>WK: ring buffer + linear-interp resample (srcRate→ctxRate)
  WK->>SPK: 128-frame quanta, ~40ms jitter buffer, 160ms cap
```

**Robustness properties**
* PCM is **uncompressed 48 kHz mono s16le (~768 kbps)** over WS, played by an **AudioWorklet** (dedicated audio thread, not the main thread → glitch-free).
* The worklet **resamples** to the device AudioContext rate (fixes pitch distortion at 44.1 kHz devices) and keeps a **~40 ms adaptive jitter buffer** (primes before play, caps at ~160 ms, re-primes on underrun).
* Server-side latency is bounded by **small realtime queues (~100 ms, drop-oldest)** and a **socket-buffer drop** (>16 KB) so voice cannot accumulate seconds of lag.

---

## 13. Sequence — input control

```mermaid
sequenceDiagram
  participant V as Viewer
  participant RT as /input WS (input_control.py)
  participant OS as Windows (win32)
  V->>RT: JSON {type: move|click|key|scroll|clipboard}
  RT->>OS: SendInput / clipboard ops (DPI-aware coords)
  RT-->>V: cursor broadcast (optional)
```

---

## 14. Sequence — payment / top-up (full)

```mermaid
sequenceDiagram
  participant V as Viewer (payment panel)
  participant RT as Runtime /api/local/topup-*
  participant Cloud
  participant RZP as Razorpay

  V->>RT: POST /api/local/topup-order?amountRupees=N
  RT->>Cloud: POST /api/agent/wallet/topup-order (Bearer device token)
  alt no device token
    RT-->>V: {success:false, error:"Device not signed in"}  // → "host not signed in"
  else
    Cloud->>RZP: orders.create(amount,currency)
    Cloud-->>RT: { order, razorpayKeyId, paymentId, minutesGranted }
    RT-->>V: same JSON
    V->>RZP: open checkout.js (key, order_id) → user pays
    RZP-->>V: { razorpay_payment_id, order_id, signature }
    V->>RT: POST /api/local/topup-verify?...signature...
    RT->>Cloud: POST /api/agent/wallet/topup-verify
    Cloud->>Cloud: verify HMAC signature → settleRazorpayPayment → credit wallet minutes
    Cloud-->>RT: { success, minutesGranted, totalMinutesRemaining }
    RT-->>V: success → unlock controls, update timer
  end
  RZP-->>Cloud: webhook payment.captured/order.paid (signed) → idempotent settle (safety net)
```

The client reads every response with a **guarded JSON parser**: non-JSON (gate/tunnel/CDN HTML) is mapped to a clear actionable message instead of `Unexpected token '<'`, and the "open via public link" hint is shown **only** for raw-IP/LAN access — never on the cloud link.

---

## 15. Sequence — session metering, grace, lock

```mermaid
sequenceDiagram
  participant RT as Runtime
  participant Cloud
  participant V as Viewer credit UI

  RT->>Cloud: POST /api/agent/session/start (entitlementDecision)
  loop ~every minute
    RT->>Cloud: POST /api/agent/session/heartbeat (debitSessionUsage)
    Cloud-->>RT: minutesRemaining (included + wallet)
  end
  V->>V: poll /api/local/credits (banner/timer)
  alt minutes → 0
    V->>V: grace countdown → controls lock overlay → session-end overlay
    V->>V: "Add Credits" → top-up flow → resume
  end
  RT->>Cloud: POST /api/agent/session/end (on stop / sign-out)
```

---

## 16. Settings / configuration storage

* File: **`C:\ProgramData\Zadoo\config.json`** managed by `SettingsStore` (`settings.py`).
* Secrets (`device_token`, `resend_api_key`) are **DPAPI-encrypted** (`dpapi:` prefix); if pywin32/DPAPI is unavailable it **degrades gracefully** to obfuscated `plain:` storage (so settings still save) rather than crashing.
* Writes go through `atomic_update` (lock-held read-modify-write) for the concurrent UI/heartbeat paths.
* Key fields: `access_code(_plain)`, `email_to`, `device_name`, `autostart_enabled`, `show_settings_in_taskbar`, `permissions{}`, `alerts{}`, `device_token`, `workspace_id`, `device_id`, `cloud_api_base`, `public_url`, `entitlement_cache`, `credits_cache`, `session_blocked`.

---

## 17. Tunnel internals (`tunnel.py`)

* Command: `cloudflared tunnel --no-autoupdate [--protocol http2] --url http://localhost:6173`.
* **`--protocol http2`** by default avoids cloudflared's QUIC-over-UDP-7844 handshake + exponential backoff on UDP-restricted networks (the main cause of slow public-URL generation). Override via `ZADOO_CLOUDFLARED_PROTOCOL`.
* **Signature verification is cached** per `(path, mtime)` so restarts don't re-run a slow PowerShell `Get-AuthenticodeSignature` each time.
* **Stale-process cleanup runs in the background** (excluding the just-started PID) so it never blocks URL generation.
* Optional **Resend** email of the public link when configured.

---

## 18. Build & deployment

```mermaid
flowchart LR
  SRC["zadoo_vnc + zadoo_vnc_single.py"] --> PYI["PyInstaller --onedir<br/>(.build_envs\\py311-x64)"]
  PYI --> ONE["dist\\onedir\\x64\\Zadoo\\Zadoo.exe + _internal"]
  ONE --> ISCC["Inno Setup (installer\\zadoo.iss)"]
  ISCC --> INST["dist\\installer\\Zadoo-1.0.0-x64-Setup.exe"]
  ISCC -. cleanup .-> RMV["onedir/build pruned<br/>→ only the installer remains"]
```

* `scripts/build_windows.ps1 -Arch x64 -NoSelfSign -SkipPortable`: creates the py3.11 venv, installs `requirements.txt` + `build_requirements.txt`, runs PyInstaller (onedir), then Inno Setup, then **deletes the loose one-folder app** so only `Zadoo-1.0.0-x64-Setup.exe` remains (`-KeepOneDir` opts out).
* The installer installs to `Program Files\Zadoo`, registers a `schtasks` logon task, creates **"Zadoo Settings" / "Open Zadoo"** shortcuts, and opens Settings on Finish. Uninstall removes the task and offers to delete `ProgramData\Zadoo`.
* The runtime needs **Python 3.11 (x64)** at build time and bundles `cloudflared.exe`; FFmpeg/MediaMTX are optional (turbo path).
* The **cloud** (`web/`) auto-deploys to **Vercel** from `zadoo-web.git`; `next build` fails on TS errors (no `ignoreBuildErrors`).

---

## 19. Environment variables (selected)

| Var | Where | Purpose |
|---|---|---|
| `ZADOO_CLOUD_API_BASE` | runtime | Cloud base (default `https://zadoo-web.vercel.app`). |
| `ZADOO_ACCESS_CODE` / `CODE_*` | runtime | Local access code(s). |
| `ZADOO_DISABLE_TUNNEL` | runtime | Local-only (no public link). |
| `ZADOO_ALLOW_DIRECT_ACCESS` | runtime | Allow non-tunnel LAN access. |
| `ZADOO_CLOUDFLARED_PROTOCOL` | runtime | `http2`(default)/`quic`/`auto`. |
| `ZADOO_STREAM_START_PROFILE` | runtime | Start ladder rung (default `540p60`). |
| `ZADOO_TARGET_KBPS` | runtime | Hard egress cap (0=off; e.g. `2000` for 2 Mbps). |
| `ZADOO_TURBO_*`, `ZADOO_FFMPEG_PATH`, `ZADOO_MEDIAMTX_PATH` | runtime | Optional WebRTC turbo path. |
| `RESEND_API_KEY`, `EMAIL_TO` | runtime | Public-link email. |
| `DATABASE_URL` | cloud | Postgres. |
| `NEXTAUTH_URL`, `AUTH_SECRET`, `NEXTAUTH_SECRET` | cloud | NextAuth. |
| `AGENT_TOKEN_PEPPER` | cloud | Pepper for `hmacToken` (device-token hashing). |
| `GOOGLE_CLIENT_ID/SECRET` | cloud | Google OAuth. |
| `RAZORPAY_KEY_ID/SECRET/WEBHOOK_SECRET` | cloud | Payments. |

---

## 20. Tech stack

* **Runtime:** Python 3.11, `websockets` (asyncio), `dxcam`/`bettercam`/`mss`/`fast-ctypes-screenshots`, `imagecodecs`/`opencv`/`Pillow` (JPEG), `soundcard`/`sounddevice` (audio), `pywin32`/`keyboard`/`pyautogui` (input), `pywinpty` (terminal), `paramiko`, Tkinter (Settings UI), PyInstaller + Inno Setup, `cloudflared`.
* **Cloud:** Next.js 15 (App Router), React 19, NextAuth v5 (Google + email OTP), Prisma + Postgres (Neon), Razorpay, Vercel.
* **Client (in-browser):** vanilla JS SPA (`index.html`), Canvas for video, **AudioWorklet** for audio, Razorpay `checkout.js`, `xterm.js` (terminal).

---

## 21. Known limitations / risks (from the architecture audit)

These are documented for completeness (not all fixed):

* **Billing:** the Razorpay **webhook can't settle top-ups** (top-up orders carry no `zadooPaymentId` note) → a captured payment is lost if the browser closes before `/topup-verify`; payment **verify isn't bound to the stored `providerOrderId`** → a valid cheap-order signature can settle a different expensive PENDING payment (signature-replay); settlement doesn't re-verify the **amount actually paid**; `session/start` and `debitSessionUsage` have **TOCTOU races** under concurrent calls.
* **Access/admin:** the default access code **`ZADOO123`** can grant a full session when a device is signed-in but `setup_complete` is false; the admin/local gate **trusts the `Host` header**, so `ZADOO_ALLOW_DIRECT_ACCESS=1` + a spoofed `Host` can reach `/api/settings/*` and `/api/runtime/*`.
* **Audio:** multiple simultaneous listeners share **one queue** (frames are split between them); system audio takes the **left channel only** (right-panned audio dropped).
* **Misc:** startup blocks up to ~8 s on a synchronous cloud `entitlement()` call; settings UI autosave runs `schtasks` on the Tk thread; some settings writes bypass `atomic_update`.

---

*Generated as a living architecture reference for Zadoo (runtime `zadoo.git` + cloud `zadoo-web.git`).*

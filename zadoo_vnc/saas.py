"""Small HTTP client for the hosted Zadoo SaaS APIs."""
from __future__ import annotations

import json
import platform
import time
import urllib.error
import urllib.request
from typing import Any

from .settings import APP_VERSION, DEFAULT_CLOUD_API_BASE, SettingsStore, get_settings_store, machine_id


class ZadooCloudClient:
    def __init__(self, store: SettingsStore | None = None):
        self.store = store or get_settings_store()

    def _base_url(self) -> str:
        data = self.store.load(reload=True)
        return str(data.get("cloud_api_base") or DEFAULT_CLOUD_API_BASE).strip().rstrip("/")

    def _request(self, method: str, path: str, payload: dict[str, Any] | None = None, *, token: str | None = None, timeout: float = 8.0) -> dict[str, Any]:
        body = None
        headers = {"Accept": "application/json"}
        if payload is not None:
            body = json.dumps(payload).encode("utf-8")
            headers["Content-Type"] = "application/json"
        if token:
            headers["Authorization"] = f"Bearer {token}"
        request = urllib.request.Request(
            self._base_url() + path,
            data=body,
            headers=headers,
            method=method.upper(),
        )
        try:
            with urllib.request.urlopen(request, timeout=timeout) as response:
                raw = response.read()
        except urllib.error.HTTPError as exc:
            raw = exc.read()
            try:
                data = json.loads(raw.decode("utf-8", "replace"))
            except Exception:
                data = {"success": False, "error": str(exc)}
            data.setdefault("status_code", exc.code)
            return data
        try:
            return json.loads(raw.decode("utf-8", "replace"))
        except Exception:
            return {"success": False, "error": "Cloud API returned invalid JSON"}

    def start_activation(self) -> dict[str, Any]:
        data = self.store.load(reload=True)
        payload = {
            "deviceName": data.get("device_name") or platform.node() or "Windows PC",
            "machineId": machine_id(),
            "platform": platform.system().lower() or "windows",
            "version": APP_VERSION,
        }
        result = self._request("POST", "/api/agent/activate/start", payload)
        if result.get("success"):
            data["activation"] = {
                "activation_id": result.get("activationId"),
                "poll_secret": result.get("pollSecret"),
                "code": result.get("code"),
                "connect_url": result.get("connectUrl") or f"{self._base_url()}/connect-device?code={result.get('code')}",
                "expires_at": result.get("expiresAt"),
                "created_at": time.time(),
            }
            self.store.save(data)
        return result

    def poll_activation(self) -> dict[str, Any]:
        data = self.store.load(reload=True)
        activation = data.get("activation") or {}
        payload = {
            "activationId": activation.get("activation_id") or activation.get("activationId"),
            "pollSecret": activation.get("poll_secret") or activation.get("pollSecret"),
        }
        if not payload["activationId"] or not payload["pollSecret"]:
            return {"success": False, "error": "Start activation first"}
        result = self._request("POST", "/api/agent/activate/poll", payload)
        if result.get("success") and result.get("status") == "claimed":
            self.store.update_cloud_status({
                "workspace_id": result.get("workspaceId"),
                "device_id": result.get("deviceId"),
                "device_name": result.get("deviceName") or data.get("device_name"),
                "device_token": result.get("deviceToken"),
                "activation": {},
            })
        return result

    def entitlement(self) -> dict[str, Any]:
        token = self.store.get_device_token()
        if not token:
            return {"success": False, "error": "Device is not activated"}
        result = self._request("GET", "/api/agent/entitlement", token=token)
        if result.get("success"):
            data = self.store.load()
            data["entitlement_cache"] = result.get("entitlement") or {}
            data["billing_status"] = {
                "last_checked_at": time.time(),
                "allowed": bool((result.get("entitlement") or {}).get("allowed")),
                "reason": (result.get("entitlement") or {}).get("reason"),
            }
            self.store.save(data)
        return result

    def heartbeat(self, public_url: str | None = None) -> dict[str, Any]:
        token = self.store.get_device_token()
        if not token:
            return {"success": False, "error": "Device is not activated"}
        result = self._request("POST", "/api/agent/heartbeat", {"publicUrl": public_url or "", "version": APP_VERSION}, token=token)
        if result.get("success"):
            data = self.store.load()
            data["entitlement_cache"] = result.get("entitlement") or {}
            data["billing_status"] = {
                "last_checked_at": time.time(),
                "allowed": bool((result.get("entitlement") or {}).get("allowed")),
                "reason": (result.get("entitlement") or {}).get("reason"),
            }
            self.store.save(data)
        return result

    def go_offline(self) -> dict[str, Any]:
        """Tell the cloud this device is stopping so the dashboard shows it Offline
        immediately (clears the public URL) instead of waiting for heartbeats to lapse."""
        token = self.store.get_device_token()
        if not token:
            return {"success": True, "skipped": True}
        return self._request("POST", "/api/agent/heartbeat", {"offline": True}, token=token, timeout=3.0)

    def start_session(self, public_url: str | None = None) -> dict[str, Any]:
        token = self.store.get_device_token()
        if not token:
            return {"success": True, "skipped": True}
        result = self._request("POST", "/api/agent/session/start", {"publicUrl": public_url or ""}, token=token)
        entitlement = result.get("entitlement") if isinstance(result, dict) else None
        if isinstance(entitlement, dict):
            data = self.store.load()
            data["entitlement_cache"] = entitlement
            self.store.save(data)
        return result

    def session_heartbeat(self, session_id: str, minutes: int = 1) -> dict[str, Any]:
        token = self.store.get_device_token()
        if not token or not session_id:
            return {"success": True, "skipped": True}
        result = self._request("POST", "/api/agent/session/heartbeat", {"sessionId": session_id, "minutes": int(minutes or 1)}, token=token)
        entitlement = result.get("entitlement") if isinstance(result, dict) else None
        if isinstance(entitlement, dict):
            data = self.store.load()
            data["entitlement_cache"] = entitlement
            # Set session_blocked flag for the tunnel/server to detect
            allowed = bool(entitlement.get("allowed", True))
            data["session_blocked"] = not allowed
            data["session_block_reason"] = str(entitlement.get("reason") or "") if not allowed else ""
            self.store.save(data)
        return result

    def end_session(self, session_id: str) -> dict[str, Any]:
        token = self.store.get_device_token()
        if not token or not session_id:
            return {"success": True, "skipped": True}
        return self._request("POST", "/api/agent/session/end", {"sessionId": session_id}, token=token)

    def fetch_profile(self) -> dict[str, Any]:
        """Fetch user profile (name, email, avatar URL) from cloud and cache locally."""
        token = self.store.get_device_token()
        if not token:
            return {"success": False, "error": "Device not signed in"}
        result = self._request("GET", "/api/agent/profile", token=token)
        if result.get("success"):
            profile = result.get("profile") or {}
            data = self.store.load()
            data["user_name"] = str(profile.get("name") or "").strip()
            data["user_email"] = str(profile.get("email") or "").strip()
            data["user_image_url"] = str(profile.get("image") or "").strip()
            self.store.save(data)
        return result

    def fetch_credits(self) -> dict[str, Any]:
        """Fetch current credits balance from cloud and cache locally."""
        token = self.store.get_device_token()
        if not token:
            return {"success": False, "error": "Device not signed in"}
        result = self._request("GET", "/api/agent/credits", token=token)
        if result.get("success"):
            data = self.store.load()
            data["credits_cache"] = result.get("credits") or {}
            self.store.save(data)
        return result


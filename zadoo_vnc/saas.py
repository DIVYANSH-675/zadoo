"""Small HTTP client for the optional Zadoo device registry."""
from __future__ import annotations

import json
import platform
import time
import urllib.error
import urllib.request
from typing import Any

from .settings import (
    APP_VERSION,
    SettingsStore,
    get_settings_store,
    machine_id,
)


class ZadooCloudClient:
    MAX_RESPONSE_BYTES = 1_000_000

    def __init__(self, store: SettingsStore | None = None):
        self.store = store if store is not None else get_settings_store()

    def _base_url(self) -> str:
        return self.store.load(reload=True)["cloud_api_base"]

    @staticmethod
    def _required_string(payload: dict[str, Any], field: str, context: str) -> str:
        value = payload.get(field)
        if not isinstance(value, str) or not value.strip():
            raise RuntimeError(f"{context} response is missing {field}")
        return value.strip()

    @staticmethod
    def _required_object(payload: dict[str, Any], field: str, context: str) -> dict[str, Any]:
        value = payload.get(field)
        if not isinstance(value, dict):
            raise RuntimeError(f"{context} response is missing {field}")
        return value

    @staticmethod
    def _response_object(raw: bytes, path: str, status: int) -> dict[str, Any]:
        try:
            text = raw.decode("utf-8")
        except UnicodeDecodeError as exc:
            return {
                "success": False,
                "error": f"Cloud API returned invalid UTF-8 for {path} (HTTP {status}): {exc}",
            }
        try:
            data = json.loads(text)
        except json.JSONDecodeError as exc:
            return {
                "success": False,
                "error": f"Cloud API returned invalid JSON for {path} (HTTP {status}): {exc.msg}",
            }
        if not isinstance(data, dict):
            return {
                "success": False,
                "error": f"Cloud API returned non-object JSON for {path} (HTTP {status})",
            }
        return data

    @classmethod
    def _read_response_body(cls, response, path: str, status: int) -> bytes:
        raw = response.read(cls.MAX_RESPONSE_BYTES + 1)
        if len(raw) > cls.MAX_RESPONSE_BYTES:
            raise RuntimeError(
                f"Cloud API response for {path} exceeds {cls.MAX_RESPONSE_BYTES} bytes (HTTP {status})"
            )
        return raw

    @staticmethod
    def _validated_result(
        data: dict[str, Any],
        path: str,
        status: int,
        *,
        http_error: bool,
    ) -> dict[str, Any]:
        if type(data.get("success")) is not bool:
            return {
                "success": False,
                "error": f"Cloud API response for {path} is missing boolean success (HTTP {status})",
            }
        if http_error and data["success"]:
            return {
                "success": False,
                "error": f"Cloud API returned success=true for {path} with HTTP {status}",
            }
        if (
            path == "/api/agent/activate/poll"
            and not data["success"]
            and data.get("status") == "expired"
            and "error" not in data
        ):
            data["error"] = "Activation code expired"
        if not data["success"] and (not isinstance(data.get("error"), str) or not data["error"].strip()):
            return {
                "success": False,
                "error": f"Cloud API failure for {path} is missing error (HTTP {status})",
            }
        return data

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
                status = response.status
                raw = self._read_response_body(response, path, status)
        except urllib.error.HTTPError as exc:
            raw = self._read_response_body(exc, path, exc.code)
            data = self._response_object(raw, path, exc.code)
            data = self._validated_result(data, path, exc.code, http_error=True)
            data["status_code"] = exc.code
            return data
        data = self._response_object(raw, path, status)
        return self._validated_result(data, path, status, http_error=False)

    def start_activation(self) -> dict[str, Any]:
        data = self.store.load(reload=True)
        system = platform.system().lower()
        if system != "windows":
            raise RuntimeError(f"Device activation requires Windows; current platform is {system}")
        payload = {
            "deviceName": data["device_name"],
            "machineId": machine_id(),
            "platform": system,
            "version": APP_VERSION,
        }
        result = self._request("POST", "/api/agent/activate/start", payload)
        if result["success"]:
            activation = {
                "activation_id": self._required_string(result, "activationId", "Activation start"),
                "poll_secret": self._required_string(result, "pollSecret", "Activation start"),
                "code": self._required_string(result, "code", "Activation start"),
                "connect_url": self._required_string(result, "connectUrl", "Activation start"),
                "expires_at": self._required_string(result, "expiresAt", "Activation start"),
                "created_at": time.time(),
            }
            self.store.atomic_update(lambda current: current.__setitem__("activation", activation))
        return result

    def poll_activation(self) -> dict[str, Any]:
        data = self.store.load(reload=True)
        activation = data["activation"]
        payload = {
            "activationId": activation.get("activation_id"),
            "pollSecret": activation.get("poll_secret"),
        }
        if not payload["activationId"] or not payload["pollSecret"]:
            return {"success": False, "error": "Start activation first"}
        result = self._request("POST", "/api/agent/activate/poll", payload)
        if not result["success"]:
            return result
        status = self._required_string(result, "status", "Activation poll")
        if status not in {"pending", "claimed"}:
            raise RuntimeError(f"Activation poll response has unsupported status: {status}")
        if status == "claimed":
            self.store.update_cloud_status({
                "workspace_id": self._required_string(result, "workspaceId", "Activation poll"),
                "device_id": self._required_string(result, "deviceId", "Activation poll"),
                "device_name": self._required_string(result, "deviceName", "Activation poll"),
                "device_token": self._required_string(result, "deviceToken", "Activation poll"),
                "activation": {},
            })
        return result

    def heartbeat(self, public_url: str | None = None) -> dict[str, Any]:
        if not (token := self.store.get_device_token()):
            return {"success": False, "error": "Device is not activated"}
        return self._request("POST", "/api/agent/heartbeat", {"publicUrl": public_url or "", "version": APP_VERSION}, token=token)

    def go_offline(self) -> dict[str, Any]:
        """Tell the cloud this device is stopping so the dashboard shows it Offline
        immediately (clears the public URL) instead of waiting for heartbeats to lapse."""
        if not (token := self.store.get_device_token()):
            return {"success": False, "error": "Device is not activated"}
        return self._request("POST", "/api/agent/heartbeat", {"offline": True}, token=token, timeout=3.0)

    def fetch_profile(self) -> dict[str, Any]:
        """Fetch the user profile from cloud and cache it locally."""
        if not (token := self.store.get_device_token()):
            return {"success": False, "error": "Device not signed in"}
        result = self._request("GET", "/api/agent/profile", token=token)
        if result["success"]:
            profile = self._required_object(result, "profile", "Profile")
            name = self._required_string(profile, "name", "Profile")
            email = self._required_string(profile, "email", "Profile")
            self.store.atomic_update(lambda data: data.update(user_name=name, user_email=email))
        return result


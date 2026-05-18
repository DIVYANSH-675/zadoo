"""Windows camera discovery helpers."""
from __future__ import annotations

import hashlib
import json
import logging
import os
import re
import sys
from dataclasses import dataclass
from typing import Any

from .process_utils import _run_hidden

log = logging.getLogger(__name__)
logging.getLogger("comtypes").setLevel(logging.WARNING)
logging.getLogger("comtypes.client").setLevel(logging.WARNING)
logging.getLogger("comtypes.client._code_cache").setLevel(logging.WARNING)


@dataclass(frozen=True)
class CameraDevice:
    id: str
    name: str
    label: str
    open_name: str
    source: str
    index: int | None = None
    device_path: str | None = None

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {
            "id": self.id,
            "name": self.name,
            "label": self.label,
            "open_name": self.open_name,
            "source": self.source,
        }
        if self.index is not None:
            data["index"] = self.index
        if self.device_path:
            data["device_path"] = self.device_path
        return data


def enumerate_camera_devices() -> list[dict[str, Any]]:
    """Return video capture devices as JSON-ready objects."""
    devices: list[dict[str, Any]] = []
    if sys.platform.startswith("win"):
        devices = _enumerate_windows_cameras()
    return devices


def normalize_camera_devices(raw_devices: list[Any]) -> list[dict[str, Any]]:
    """Normalize legacy strings or device dicts into the public camera shape."""
    entries: list[dict[str, Any]] = []
    for index, item in enumerate(raw_devices or []):
        if isinstance(item, dict):
            name = _clean_name(item.get("name") or item.get("label") or item.get("open_name"))
            if not name:
                continue
            device_path = _clean_name(item.get("device_path"))
            open_name = _clean_name(item.get("open_name")) or name
            source = _clean_name(item.get("source")) or "camera"
            raw_id = _clean_name(item.get("id")) or _camera_id(source, index, name, device_path)
            entries.append(
                {
                    "id": raw_id,
                    "name": name,
                    "label": _clean_name(item.get("label")) or name,
                    "open_name": open_name,
                    "source": source,
                    "index": item.get("index", index),
                    **({"device_path": device_path} if device_path else {}),
                }
            )
            continue
        name = _clean_name(item)
        if name:
            entries.append(_camera_dict(name=name, index=index, source="legacy"))
    return _label_duplicates(_dedupe_devices(entries))


def resolve_camera_selection(camera_id: str | None, device_name: str | None, devices: list[dict[str, Any]] | None = None) -> dict[str, Any] | None:
    """Find the selected camera object by id or by any legacy name field."""
    normalized = normalize_camera_devices(devices or enumerate_camera_devices())
    wanted_id = (camera_id or "").strip()
    wanted_name = (device_name or "").strip()
    if wanted_id:
        for device in normalized:
            if wanted_id == str(device.get("id", "")):
                return device
    if wanted_name:
        for device in normalized:
            values = (
                device.get("open_name"),
                device.get("name"),
                device.get("label"),
                device.get("device_path"),
                device.get("id"),
            )
            if any(wanted_name == str(value) for value in values if value):
                return device
    return None


def camera_open_candidates(device: dict[str, Any] | None) -> list[str]:
    """Return DirectShow open strings to try for one selected camera."""
    if not device:
        return []
    candidates: list[str] = []
    keys = ["open_name", "name"]
    if device.get("source") in {"pygrabber", "directshow"}:
        device_path = _ffmpeg_directshow_path(device.get("device_path"))
        if device_path:
            candidates.append(device_path)
    for key in keys:
        value = _clean_name(device.get(key))
        if value and value not in candidates:
            candidates.append(value)
    return candidates


def _enumerate_windows_cameras() -> list[dict[str, Any]]:
    names = _pygrabber_names()
    metadata = _directshow_metadata()
    devices = _merge_names_and_metadata(names, metadata)
    if not devices:
        devices = _powershell_pnp_devices()
    return normalize_camera_devices(devices)


def _pygrabber_names() -> list[str]:
    initialized = False
    try:
        try:
            from comtypes import CoInitialize, CoUninitialize
            CoInitialize()
            initialized = True
        except Exception:
            initialized = False
        from pygrabber.dshow_graph import FilterGraph

        graph = FilterGraph()
        return [_clean_name(name) for name in graph.get_input_devices() if _clean_name(name)]
    except Exception as exc:
        log.warning("pygrabber camera enumeration failed: %s", exc)
        return []
    finally:
        if initialized:
            try:
                CoUninitialize()
            except Exception:
                pass


def _directshow_metadata() -> list[dict[str, Any]]:
    try:
        from comtypes import GUID, CoInitialize, CoUninitialize
        from comtypes import client
        from comtypes.persist import IPropertyBag
        from pygrabber.dshow_ids import DeviceCategories, clsids
        try:
            from pygrabber.dshow_core import ICreateDevEnum
        except Exception:
            from pygrabber.dshow_structures import ICreateDevEnum
    except Exception as exc:
        log.debug("DirectShow metadata imports unavailable: %s", exc)
        return []

    initialized = False
    try:
        try:
            CoInitialize()
            initialized = True
        except Exception:
            initialized = False
        system_device_enum = client.CreateObject(clsids.CLSID_SystemDeviceEnum, interface=ICreateDevEnum)
        video_category = (
            getattr(DeviceCategories, "CLSID_VideoInputDeviceCategory", None)
            or getattr(DeviceCategories, "VideoInputDevice")
        )
        enumerator = system_device_enum.CreateClassEnumerator(
            GUID(video_category),
            dwFlags=0,
        )
        if not enumerator:
            return []
        results: list[dict[str, Any]] = []
        moniker, count = enumerator.Next(1)
        index = 0
        while count > 0:
            item = _read_directshow_moniker(moniker, IPropertyBag)
            if item.get("name"):
                item["index"] = index
                item["source"] = "directshow"
                results.append(item)
            index += 1
            moniker, count = enumerator.Next(1)
        return results
    except Exception as exc:
        log.debug("DirectShow metadata enumeration failed: %s", exc)
        return []
    finally:
        if initialized:
            try:
                CoUninitialize()
            except Exception:
                pass


def _read_directshow_moniker(moniker: Any, property_bag_type: Any) -> dict[str, Any]:
    try:
        bag = moniker.BindToStorage(0, 0, property_bag_type._iid_).QueryInterface(property_bag_type)
    except Exception:
        return {}
    item: dict[str, Any] = {}
    for prop, field in (("FriendlyName", "name"), ("Description", "description"), ("DevicePath", "device_path")):
        try:
            value = _clean_name(bag.Read(prop, pErrorLog=None))
            if value:
                item[field] = value
        except Exception:
            continue
    return item


def _merge_names_and_metadata(names: list[str], metadata: list[dict[str, Any]]) -> list[dict[str, Any]]:
    devices: list[dict[str, Any]] = []
    used_metadata: set[int] = set()
    for index, name in enumerate(names):
        meta_index = _matching_metadata_index(name, index, metadata, used_metadata)
        meta = metadata[meta_index] if meta_index is not None else {}
        if meta_index is not None:
            used_metadata.add(meta_index)
        devices.append(
            _camera_dict(
                name=name,
                index=index,
                source="pygrabber",
                device_path=_clean_name(meta.get("device_path")),
            )
        )
    for index, meta in enumerate(metadata):
        if index in used_metadata:
            continue
        name = _clean_name(meta.get("name"))
        if name:
            devices.append(
                _camera_dict(
                    name=name,
                    index=meta.get("index", len(devices)),
                    source="directshow",
                    device_path=_clean_name(meta.get("device_path")),
                )
            )
    return devices


def _matching_metadata_index(name: str, index: int, metadata: list[dict[str, Any]], used: set[int]) -> int | None:
    if index < len(metadata) and index not in used and _clean_name(metadata[index].get("name")) == name:
        return index
    for meta_index, item in enumerate(metadata):
        if meta_index in used:
            continue
        if _clean_name(item.get("name")) == name:
            return meta_index
    return None


def _powershell_pnp_devices() -> list[dict[str, Any]]:
    ps = os.path.join(
        os.environ.get("SystemRoot", "C:\\Windows"),
        "System32",
        "WindowsPowerShell",
        "v1.0",
        "powershell.exe",
    )
    commands = [
        "$ErrorActionPreference='SilentlyContinue'; Get-PnpDevice -Class Camera | Select-Object FriendlyName,InstanceId,Status | ConvertTo-Json -Compress",
        "$ErrorActionPreference='SilentlyContinue'; Get-PnpDevice -Class Image | Select-Object FriendlyName,InstanceId,Status | ConvertTo-Json -Compress",
        "$ErrorActionPreference='SilentlyContinue'; Get-CimInstance Win32_PnPEntity | Where-Object { $_.PNPClass -eq 'Camera' -or $_.PNPClass -eq 'Image' -or $_.Name -match 'camera|webcam|video' } | Select-Object Name,PNPDeviceID,PNPClass | ConvertTo-Json -Compress",
    ]
    devices: list[dict[str, Any]] = []
    for command in commands:
        try:
            result = _run_hidden([ps, "-NoProfile", "-Command", command], timeout=8)
        except Exception as exc:
            log.debug("camera PnP command failed: %s", exc)
            continue
        output = (result.stdout or "").strip()
        if not output:
            continue
        for item in _json_items(output):
            name = _clean_name(item.get("FriendlyName") or item.get("Name"))
            if not name or _is_audio_name(name):
                continue
            device_path = _clean_name(item.get("InstanceId") or item.get("PNPDeviceID"))
            devices.append(
                _camera_dict(
                    name=name,
                    index=len(devices),
                    source="windows-pnp",
                    device_path=device_path,
                )
            )
    return devices


def _json_items(output: str) -> list[dict[str, Any]]:
    try:
        parsed = json.loads(output)
    except Exception:
        return []
    if isinstance(parsed, dict):
        return [parsed]
    if isinstance(parsed, list):
        return [item for item in parsed if isinstance(item, dict)]
    return []


def _camera_dict(name: str, index: int | None, source: str, device_path: str | None = None) -> dict[str, Any]:
    cleaned_name = _clean_name(name)
    cleaned_path = _clean_name(device_path)
    return CameraDevice(
        id=_camera_id(source, index, cleaned_name, cleaned_path),
        name=cleaned_name,
        label=cleaned_name,
        open_name=cleaned_name,
        source=source,
        index=index,
        device_path=cleaned_path,
    ).to_dict()


def _camera_id(source: str, index: int | None, name: str, device_path: str | None) -> str:
    identity = device_path or f"{source}:{index}:{name}"
    digest = hashlib.sha1(identity.encode("utf-8", errors="ignore")).hexdigest()[:12]
    return f"cam-{digest}"


def _dedupe_devices(devices: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for device in devices:
        key = str(device.get("id") or device.get("device_path") or f"{device.get('source')}:{device.get('index')}:{device.get('name')}")
        if key in seen:
            continue
        seen.add(key)
        result.append(device)
    return result


def _label_duplicates(devices: list[dict[str, Any]]) -> list[dict[str, Any]]:
    totals: dict[str, int] = {}
    for device in devices:
        name = str(device.get("name") or "")
        totals[name] = totals.get(name, 0) + 1
    counts: dict[str, int] = {}
    labeled: list[dict[str, Any]] = []
    for device in devices:
        copy = dict(device)
        name = str(copy.get("name") or "")
        counts[name] = counts.get(name, 0) + 1
        copy["label"] = f"{name} ({counts[name]})" if totals.get(name, 0) > 1 else name
        labeled.append(copy)
    return labeled


def _clean_name(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _is_audio_name(name: str) -> bool:
    return bool(re.search(r"\b(microphone|mic|audio|speaker|line[- ]?in|line[- ]?out|headset)\b", name, re.IGNORECASE))


def _ffmpeg_directshow_path(value: Any) -> str:
    path = _clean_name(value)
    if not path:
        return ""
    if path.startswith("@device_"):
        return path
    if path.startswith("\\\\?\\"):
        return f"@device_pnp_{path}"
    return ""

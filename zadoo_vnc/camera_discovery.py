"""Windows camera discovery helpers."""
from __future__ import annotations

import hashlib
import sys
from typing import Any


def enumerate_camera_devices() -> list[dict[str, Any]]:
    """Return video capture devices as JSON-ready objects."""
    if not sys.platform.startswith("win"):
        raise RuntimeError(f"Camera discovery requires Windows; current platform is {sys.platform}")
    return _normalize_camera_devices(_directshow_metadata())


def _normalize_camera_devices(raw_devices: list[Any]) -> list[dict[str, Any]]:
    entries: list[dict[str, Any]] = []
    for index, item in enumerate(raw_devices):
        if not isinstance(item, dict):
            raise ValueError(f"Camera {index} must be an object")
        name = _clean_name(item.get("name"))
        if not name:
            raise ValueError(f"Camera {index} has no name")
        device_path = _ffmpeg_directshow_path(item.get("device_path"))
        if not device_path:
            raise ValueError(f"Camera {index} has no valid DirectShow device path")
        entries.append(
            {
                "id": _camera_id(device_path),
                "name": name,
                "label": name,
                "device_path": device_path,
            }
        )
    return _label_duplicates(_dedupe_devices(entries))


def camera_open_target(device: dict[str, Any]) -> str:
    """Return the exact DirectShow target for one selected camera."""
    if not isinstance(device, dict):
        raise ValueError("Camera must be an object")
    device_path = _ffmpeg_directshow_path(device.get("device_path"))
    if not device_path:
        raise ValueError(f"Camera {device.get('id')!r} has no valid DirectShow device path")
    return device_path


def _directshow_metadata() -> list[dict[str, Any]]:
    from comtypes import GUID, CoInitialize, CoUninitialize, client
    from comtypes.persist import IPropertyBag
    from pygrabber.dshow_core import ICreateDevEnum
    from pygrabber.dshow_ids import DeviceCategories, clsids

    CoInitialize()
    try:
        system_device_enum = client.CreateObject(clsids.CLSID_SystemDeviceEnum, interface=ICreateDevEnum)
        enumerator = system_device_enum.CreateClassEnumerator(
            GUID(DeviceCategories.VideoInputDevice),
            dwFlags=0,
        )
        if not enumerator:
            return []
        results: list[dict[str, Any]] = []
        moniker, count = enumerator.Next(1)
        while count > 0:
            results.append(_read_directshow_moniker(moniker, IPropertyBag))
            moniker, count = enumerator.Next(1)
        return results
    finally:
        CoUninitialize()


def _read_directshow_moniker(moniker: Any, property_bag_type: Any) -> dict[str, Any]:
    try:
        bag = moniker.BindToStorage(0, 0, property_bag_type._iid_).QueryInterface(property_bag_type)
    except Exception as exc:
        raise RuntimeError(f"Camera metadata binding failed: {exc}") from exc
    item: dict[str, Any] = {}
    for prop, field in (("FriendlyName", "name"), ("DevicePath", "device_path")):
        try:
            value = _clean_name(bag.Read(prop, pErrorLog=None))
        except Exception as exc:
            raise RuntimeError(f"Camera property {prop} read failed: {exc}") from exc
        if not value:
            raise RuntimeError(f"Camera property {prop} is empty")
        item[field] = value
    return item


def _camera_id(device_path: str) -> str:
    digest = hashlib.sha256(device_path.encode("utf-8")).hexdigest()[:16]
    return f"cam-{digest}"


def _dedupe_devices(devices: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    seen: set[str] = set()
    for device in devices:
        key = str(device["id"])
        if key in seen:
            continue
        seen.add(key)
        result.append(device)
    return result


def _label_duplicates(devices: list[dict[str, Any]]) -> list[dict[str, Any]]:
    totals: dict[str, int] = {}
    for device in devices:
        name = device["name"]
        totals[name] = totals.get(name, 0) + 1
    counts: dict[str, int] = {}
    labeled: list[dict[str, Any]] = []
    for device in devices:
        copy = dict(device)
        name = copy["name"]
        counts[name] = counts.get(name, 0) + 1
        copy["label"] = f"{name} ({counts[name]})" if totals[name] > 1 else name
        labeled.append(copy)
    return labeled


def _clean_name(value: Any) -> str:
    return value.strip() if isinstance(value, str) else ""


def _ffmpeg_directshow_path(value: Any) -> str:
    path = _clean_name(value)
    if not path:
        return ""
    if path.startswith("@device_"):
        return path
    if path.startswith("\\\\?\\"):
        return f"@device_pnp_{path}"
    return ""

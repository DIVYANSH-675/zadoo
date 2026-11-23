"""Bounded desktop-duplication adapter for the pinned BetterCam 1.0.0 backend."""
from __future__ import annotations

import time

from bettercam.bettercam import BetterCam
from bettercam.core import Device, Output
from bettercam.util.io import enum_dxgi_adapters
from comtypes import COMError


class CaptureUnavailable(RuntimeError):
    """A desktop transition requires fresh capture resources, not a busy loop."""


RECOVERABLE_HRESULTS = {
    0x80070005: "E_ACCESSDENIED",
    0x887A0004: "DXGI_ERROR_UNSUPPORTED",
    0x887A0005: "DXGI_ERROR_DEVICE_REMOVED",
    0x887A0007: "DXGI_ERROR_DEVICE_RESET",
    0x887A0022: "DXGI_ERROR_NOT_CURRENTLY_AVAILABLE",
    0x887A0026: "DXGI_ERROR_ACCESS_LOST",
    0x887A0028: "DXGI_ERROR_SESSION_DISCONNECTED",
}


def recoverable_com_error(error: COMError) -> bool:
    return (error.hresult & 0xFFFFFFFF) in RECOVERABLE_HRESULTS


def capture_error_text(error: Exception) -> str:
    if isinstance(error, COMError):
        code = error.hresult & 0xFFFFFFFF
        name = RECOVERABLE_HRESULTS.get(code, "COM_ERROR")
        return f"{name} (0x{code:08X}): {error}"
    return str(error)


def release_camera_resources(camera):
    """Let comtypes release each reference once, including partial construction."""
    duplicator = getattr(camera, "_duplicator", None)
    if duplicator is not None:
        duplicator.texture = None
        duplicator.duplicator = None
    surface = getattr(camera, "_stagesurf", None)
    if surface is not None:
        surface.texture = None
        surface.width = 0
        surface.height = 0


class RecoverableBetterCam(BetterCam):
    def _on_output_change(self):
        # BetterCam's implementation retries COM failures forever without a delay
        # or stop check. Return control to ScreenCapturer's bounded recovery loop.
        raise CaptureUnavailable("DXGI_ERROR_ACCESS_LOST (0x887A0026): desktop capture access was lost")

    def release(self):
        # The parent destructor also calls this after a failed __init__.
        if hasattr(self, "_BetterCam__frame_available"):
            self.stop()
        release_camera_resources(self)


def create_camera():
    """Use fresh DXGI/D3D objects, never the factory's cached lost device."""
    # DXGI lists the desktop-primary adapter/output first. Re-enumerating only at
    # creation/recovery also handles display topology changes without per-frame work.
    for adapter in enum_dxgi_adapters():
        device = Device(adapter)
        for output_pointer in device.enum_outputs():
            output = Output(output_pointer)
            if output.attached_to_desktop and min(output.resolution) > 0:
                camera = RecoverableBetterCam(output=output, device=device, region=None, output_color="BGRA")
                # Preserve BetterCam.create's initialization settling time.
                time.sleep(0.1)
                return camera
    raise CaptureUnavailable("DXGI_ERROR_NOT_FOUND (0x887A0002): no active desktop display is available")

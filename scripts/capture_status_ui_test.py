"""Offline browser regression: capture errors must be visible, exact, and safe."""
from __future__ import annotations

import os
import sys
from pathlib import Path

_DLL_DIRECTORY = os.add_dll_directory(str(Path(sys.prefix) / "Scripts")) if os.name == "nt" else None
from playwright.sync_api import expect, sync_playwright  # noqa: E402
from viewer_layout_test import open_controls, serve_asset  # noqa: E402


def main():
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(viewport={"width": 390, "height": 844}, reduced_motion="reduce")
        context.route("**/*", serve_asset)
        page = context.new_page()
        errors = []
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.goto("http://zadoo-preview.test/", wait_until="load")
        open_controls(page)
        page.locator("#connection-details > summary").click()
        message = "Screen capture recovering: E_ACCESSDENIED (0x80070005): <img src=x onerror=alert(1)>"
        stats = {"is_working": False, "capture_state": "recovering", "error": message,
                 "stream_status": "test", "stream_profile": "test"}
        page.evaluate("stats => handleCaptureStats(stats)", stats)
        expect(page.locator("#capture-notice")).to_be_visible()
        expect(page.locator("#capture-notice")).to_have_text(message)
        expect(page.locator("#status-text")).to_have_text(message)
        expect(page.locator("#status-indicator")).to_have_class("status-indicator warning")
        expect(page.locator("#capture-notice img")).to_have_count(0)
        for selector in ("#capture-notice", "#sidebar"):
            if not page.locator(selector).evaluate("e => e.scrollWidth <= e.clientWidth"):
                raise AssertionError(f"Long capture error overflows {selector}")
        # The ordinary FPS timer must not erase the visible capture failure.
        page.wait_for_timeout(1100)
        expect(page.locator("#capture-notice")).to_have_text(message)
        stats.update({"capture_state": "failed", "error": "Screen capture recovery failed after 30 seconds: exact cause"})
        page.evaluate("stats => handleCaptureStats(stats)", stats)
        expect(page.locator("#capture-notice")).to_have_text(stats["error"])
        stats.update({"is_working": True, "capture_state": "ready", "error": None})
        page.evaluate("stats => handleCaptureStats(stats)", stats)
        expect(page.locator("#capture-notice")).not_to_be_visible()
        expect(page.locator("#status-text")).to_have_text("Ready")
        if errors:
            raise AssertionError(errors)
        browser.close()
    print("CAPTURE_STATUS_UI: PASS exact errors, recovery, safe text, mobile layout, cleared after success")


if __name__ == "__main__":
    main()

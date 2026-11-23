"""Offline visual checks of the actual viewer assets; never connects to a host."""
from __future__ import annotations

import json
import mimetypes
import os
import sys
from pathlib import Path
from urllib.parse import urlsplit

_DLL_DIRECTORY = os.add_dll_directory(str(Path(sys.prefix) / "Scripts")) if os.name == "nt" else None
from playwright.sync_api import expect, sync_playwright  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT.parent / "outputs" / "zadoo-1.1.9-ui"
SIZES = [(320, 568), (390, 844), (768, 1024), (667, 375), (844, 390), (1280, 900)]
TOUCH_SIZES = [(320, 568), (390, 844), (844, 390)]


def check(value, message):
    if not value:
        raise AssertionError(message)


def serve_asset(route):
    path = urlsplit(route.request.url).path
    if path == "/":
        asset = ROOT / "zadoo_vnc/templates/index.html"
    elif path.startswith("/static/"):
        asset = ROOT / "zadoo_vnc" / path.lstrip("/")
    elif path in ("/splash.png", "/brand-header.png", "/trigger-icon.png", "/favicon.ico"):
        asset = ROOT / ("app_icon.ico" if path == "/favicon.ico" else path.lstrip("/"))
    else:
        route.fulfill(status=404, body="Offline layout preview: no host connected")
        return
    check(asset.resolve().is_relative_to(ROOT), "Preview asset escaped the repository")
    route.fulfill(body=asset.read_bytes(), content_type=mimetypes.guess_type(asset.name)[0] or "application/octet-stream")


def open_controls(page):
    # Layout-only state: no authentication, sharing, network session or remote input.
    page.evaluate("() => { document.getElementById('splash-screen').hidden = true; document.body.classList.remove('splash-active'); toggleSidebar(true); }")
    page.wait_for_function("document.querySelector('.sidebar-logo').naturalWidth === 768")


def check_controls(page, width, height, *, touch=False):
    page.locator("#sidebar .sidebar-content").evaluate("e => { e.scrollTop = 0; }")
    sidebar = page.locator("#sidebar")
    check(sidebar.evaluate("e => e.scrollWidth <= e.clientWidth"), f"Sidebar overflow: {width}px")
    box = sidebar.bounding_box()
    check(box and box["x"] >= -1 and box["x"] + box["width"] <= width + 1, "Sidebar extends outside the viewport")
    check(abs(box["width"] - 300) < 1, "The approved sidebar must stay 300 px wide")
    check("free" not in sidebar.inner_text().lower(), "Do not show a free badge in the sidebar")
    check(page.locator("#sidebar #auth-code").count() == 0, "Host code must not be in the controls sidebar")
    check(page.locator("#sidebar .deck-fold").count() == 3, "Keep all three advanced sections")
    check(page.locator("#sidebar .deck-action").count() == 15, "Keep all fifteen tools")
    for selector in ("#view-mode", "#view-zoom-out", "#view-zoom-in", "#view-reset", ".deck-close", ".deck-action"):
        for control in page.locator(selector).all():
            bounds = control.bounding_box()
            check(bounds and min(bounds["width"], bounds["height"]) >= (44 if touch else 30), f"Small control: {selector}")
            check(control.evaluate("e => e.scrollWidth <= e.clientWidth"), f"Clipped control: {selector}")
    prefix = "touch-controls" if touch else "controls"
    sidebar.screenshot(path=str(OUTPUT / f"{prefix}-{width}x{height}.png"))
    for details, child in (("connection-details", "#sidebar-fps"), ("capture-details", "#perf-region"), ("public-link-details", "#public-link-section")):
        summary = page.locator(f"#{details} > summary")
        expect(page.locator(child)).not_to_be_visible()
        summary.press("Enter")
        expect(page.locator(child)).to_be_visible()
        check(sidebar.evaluate("e => e.scrollWidth <= e.clientWidth"), f"Expanded {details} overflows")
        if touch:
            for control in page.locator(f"#{details} button:visible, #{details} select:visible, #{details} .focus-gray:visible").all():
                bounds = control.bounding_box()
                check(bounds and bounds["height"] >= 44, f"Small touch control in {details}")
        summary.press("Enter")
        expect(page.locator(child)).not_to_be_visible()


def main():
    OUTPUT.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(reduced_motion="reduce")
        context.route("**/*", serve_asset)
        page = context.new_page()
        errors = []
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.goto("http://zadoo-preview.test/", wait_until="load")
        expect(page.locator(".splash-art img")).to_have_attribute("src", "/splash.png?v=a8ff7f23c7aa")
        page.wait_for_function("document.querySelector('.splash-art img').naturalWidth === 1200")
        for width, height in SIZES:
            page.set_viewport_size({"width": width, "height": height})
            page.locator("#splash-screen").evaluate("e => { e.scrollTop = 0; }")
            check(page.locator("#splash-screen").evaluate("e => e.scrollWidth <= e.clientWidth"), f"Login overflow: {width}x{height}")
            art = page.locator(".splash-art").bounding_box()
            field = page.locator("#auth-code").bounding_box()
            check(art and field, "Missing wooden-box input")
            check(art["x"] + art["width"] * .14 <= field["x"]
                  and field["x"] + field["width"] <= art["x"] + art["width"] * .84
                  and art["y"] + art["height"] * .72 <= field["y"]
                  and field["y"] + field["height"] <= art["y"] + art["height"] * .94,
                  f"Code input is not on the wooden box at {width}x{height}: {field}")
            check(field["height"] >= 44, "Code field must be touch-friendly")
            expect(page.get_by_role("button", name="Connect", exact=True)).to_be_enabled()
            page.screenshot(path=str(OUTPUT / f"login-{width}x{height}.png"))
        open_controls(page)
        for width, height in SIZES:
            page.set_viewport_size({"width": width, "height": height})
            check_controls(page, width, height)
        context.close()
        for width, height in TOUCH_SIZES:
            touch = browser.new_context(viewport={"width": width, "height": height}, is_mobile=True, has_touch=True, reduced_motion="reduce")
            touch.route("**/*", serve_asset)
            page = touch.new_page()
            page.on("pageerror", lambda error: errors.append(str(error)))
            page.goto("http://zadoo-preview.test/", wait_until="load")
            check(page.evaluate("matchMedia('(pointer: coarse)').matches"), "Touch test must use a coarse pointer")
            open_controls(page)
            check_controls(page, width, height, touch=True)
            touch.close()
        check(not errors, errors)
        browser.close()
    print(json.dumps({"offline_layouts": SIZES, "touch_layouts": TOUCH_SIZES, "original_wooden_box_and_logo": "pass", "input_on_wood": "pass", "compact_sidebar_and_native_disclosures": "pass", "sidebar_touch_targets": "pass", "browser_errors": errors}, indent=2))


if __name__ == "__main__":
    main()

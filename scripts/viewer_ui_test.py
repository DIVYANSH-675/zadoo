"""Real-browser UI/reconnect tests; synthetic input is intercepted, never typed on the host."""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

_DLL_DIRECTORY = os.add_dll_directory(str(Path(sys.prefix) / "Scripts")) if os.name == "nt" else None

# Playwright's native dependency needs the DLL directory registered first.
from playwright.sync_api import expect, sync_playwright  # noqa: E402


def check(condition, message):
    if not condition:
        raise AssertionError(message)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", default="http://127.0.0.1:6173")
    parser.add_argument("--code", required=True)
    args = parser.parse_args()
    screenshots = Path(__file__).resolve().parents[2] / "outputs" / "smooth-zadoo-e2e"
    screenshots.mkdir(parents=True, exist_ok=True)
    errors = []
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch()
        context = browser.new_context(viewport={"width": 1280, "height": 900})
        context.add_init_script("""(() => {
            window.__testListeners = {};
            window.__testInput = [];
            const add = EventTarget.prototype.addEventListener;
            EventTarget.prototype.addEventListener = function(name, ...rest) {
                if (this === document && ['mousedown','mouseup','mousemove','keydown','keyup','wheel'].includes(name))
                    window.__testListeners[name] = (window.__testListeners[name] || 0) + 1;
                return add.call(this, name, ...rest);
            };
            const send = WebSocket.prototype.send;
            WebSocket.prototype.send = function(data) {
                if (new URL(this.url).pathname === '/input' && typeof data === 'string') {
                    const value = JSON.parse(data);
                    if (['move','drag','click','key','scroll'].includes(value.action)) {
                        window.__testInput.push(value); return;
                    }
                }
                return send.call(this, data);
            };
        })();""")
        page = context.new_page()
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.goto(args.base, wait_until="domcontentloaded")
        page.get_by_role("button", name="Connect", exact=True).click()
        expect(page.locator("#auth-error")).to_have_text("Access code is required")
        page.screenshot(path=str(screenshots / "viewer-login.png"))
        page.get_by_label("Access code", exact=True).fill("WRONG-E2E")
        page.get_by_role("button", name="Connect", exact=True).click()
        expect(page.locator("#auth-error")).to_have_text("Invalid code")
        page.get_by_label("Access code", exact=True).fill(args.code)
        page.get_by_role("button", name="Connect", exact=True).click()
        page.wait_for_function("window.__zadooAuthenticated && videoMetrics.displayedFrames > 0")
        page.locator("#sidebar-trigger").press("Enter")
        expect(page.locator("#sidebar-trigger")).to_have_attribute("aria-expanded", "true")
        expect(page.locator(".deck-close")).to_be_focused()
        check(page.locator("#sidebar").evaluate("e => getComputedStyle(e).transitionProperty === 'transform, visibility'"), "Panel must animate position, not layout")
        expect(page.locator("#sidebar-fps")).to_be_visible()
        page.locator("#quality-slider").press("End")
        expect(page.locator("#quality-auto")).to_have_attribute("aria-pressed", "false")
        page.locator("#quality-auto").click()
        expect(page.locator("#quality-auto")).to_have_attribute("aria-pressed", "true")
        check("Manual quality" not in page.locator("#quality-help").inner_text(), "Auto quality hint did not update")
        page.locator(".deck-close").scroll_into_view_if_needed()
        page.locator("#sidebar").screenshot(path=str(screenshots / "viewer-controls.png"))
        page.locator(".deck-close").press("Escape")
        expect(page.locator("#sidebar-trigger")).to_have_attribute("aria-expanded", "false")
        expect(page.locator("#sidebar-trigger")).to_be_focused()
        check(page.locator("#sidebar").evaluate("e => e.inert"), "Hidden controls must not receive focus")
        listeners = page.evaluate("window.__testListeners")
        for _ in range(5):
            page.evaluate("() => { window.inputSocket.close(); }")
            page.wait_for_function("window.inputSocket && window.inputSocket.readyState === WebSocket.OPEN")
        check(page.evaluate("window.__testListeners") == listeners, "Reconnect added document listeners")
        page.locator("#screen").press("a")
        sent = page.evaluate("window.__testInput.filter(e => e.action === 'key')")
        check(sent == [{"action": "key", "key": "a", "state": "down"}, {"action": "key", "key": "a", "state": "up"}], sent)
        page.locator("#sidebar-trigger").press("Enter")
        page.locator("#quality-auto").press("Tab")
        check(page.evaluate("window.__testInput.filter(e => e.action === 'key').length") == 2, "Local control keys leaked to host")
        for width in [320, 390, 820]:
            page.set_viewport_size({"width": width, "height": 844})
            page.wait_for_function("() => { const r = document.querySelector('#sidebar').getBoundingClientRect(); return r.x >= -1 && r.right <= innerWidth + 1; }")
            box = page.locator("#sidebar").bounding_box()
            check(box and box["x"] >= -1 and box["x"] + box["width"] <= width + 1, box)
            check(page.locator("#sidebar").evaluate("e => e.scrollWidth <= e.clientWidth"), f"Panel overflows at {width}px")
        page.emulate_media(reduced_motion="reduce")
        check(page.locator("#sidebar").evaluate("e => getComputedStyle(e).transitionDuration === '0s'"), "Reduced motion is not respected")
        check(not errors, errors)
        browser.close()
    print(json.dumps({"authentication_form": "pass", "keyboard_panel_navigation": "pass", "quality_manual_auto": "pass", "reconnects_without_duplicate_input": 5, "mobile_tablet_widths": [320, 390, 820], "reduced_motion": "pass", "browser_errors": errors}, indent=2))


if __name__ == "__main__":
    main()

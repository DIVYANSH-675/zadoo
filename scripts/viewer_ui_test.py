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
            window.__testLocalRequests = [];
            Object.defineProperty(navigator, 'clipboard', {value: {writeText: async value => { window.__testClipboard = value; }}});
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
                    if (value.action === 'get_public_url') window.__testLocalRequests.push(value.action);
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
        page.wait_for_function("document.querySelector('.splash-art img').naturalWidth === 1200")
        expect(page.locator(".splash-art img")).to_have_attribute("src", "/splash.png?v=a8ff7f23c7aa")
        check(not page.locator("#support-qr").get_attribute("src"), "QR must load only when support is opened")
        for width, height in [(320, 568), (390, 844), (768, 1024), (667, 375), (844, 390), (1280, 900)]:
            page.set_viewport_size({"width": width, "height": height})
            page.locator("#splash-screen").evaluate("e => { e.scrollTop = 0; }")
            check(page.locator("#splash-screen").evaluate("e => e.scrollWidth <= e.clientWidth"), f"Login overflows at {width}px")
            story = page.locator(".splash-container").bounding_box()
            check(story and story["y"] >= 0, f"Login header is unreachable at {width}x{height}: {story}")
            page.locator("#auth-code").scroll_into_view_if_needed()
            box = page.locator("#auth-code").bounding_box()
            check(box and box["x"] >= 0 and box["x"] + box["width"] <= width, box)
            if width == 390:
                page.screenshot(path=str(screenshots / "viewer-login-mobile.png"))
        page.screenshot(path=str(screenshots / "viewer-login-desktop.png"))
        page.locator("#splash-screen [data-open-support]").click()
        expect(page.locator("#support-dialog")).to_be_visible()
        page.wait_for_function("document.querySelector('#support-qr').naturalWidth === 1440")
        page.locator("#support-copy-upi").click()
        expect(page.locator("#support-status")).to_contain_text("UPI ID copied")
        check(page.evaluate("window.__testClipboard") == "divyanshgupta675@oksbi", "Incorrect support recipient")
        page.locator("#support-dialog .dialog-close").press("Escape")
        expect(page.locator("#support-dialog")).not_to_be_visible()
        expect(page.locator("#splash-screen [data-open-support]")).to_be_focused()
        page.get_by_role("button", name="Connect", exact=True).click()
        expect(page.locator("#auth-error")).to_have_text("Access code is required")
        page.screenshot(path=str(screenshots / "viewer-login.png"))
        page.get_by_label("Access code", exact=True).fill("WRONG-E2E")
        page.get_by_role("button", name="Connect", exact=True).click()
        expect(page.locator("#auth-error")).to_have_text("Invalid code")
        page.get_by_label("Access code", exact=True).fill(args.code)
        page.get_by_role("button", name="Connect", exact=True).click()
        page.wait_for_function("window.__zadooAuthenticated && videoMetrics.displayedFrames > 0")
        hand = page.locator("#sidebar-trigger").bounding_box()
        check(hand and hand["x"] > 1200 and min(hand["width"], hand["height"]) >= 44, "Hand must stay on the right edge")
        expect(page.locator("#view-controls")).not_to_be_visible()
        check(page.locator("#sidebar #auth-code").count() == 0, "Code entry must remain outside the controls sidebar")
        page.locator("#sidebar-trigger").press("Enter")
        expect(page.locator("#sidebar-trigger")).to_have_attribute("aria-expanded", "true")
        expect(page.locator(".deck-close")).to_be_focused()
        page.wait_for_function("document.querySelector('.sidebar-logo').naturalWidth === 768")
        check(page.locator("#sidebar #view-controls").count() == 1, "View controls must live inside the sidebar")
        panel = page.locator("#sidebar").bounding_box()
        check(panel and abs(panel["width"] - 300) < 1, "The approved sidebar must stay compact")
        check("free" not in page.locator("#sidebar").inner_text().lower(), "Do not show a free badge in the sidebar")
        for button in ["#view-mode", "#view-zoom-out", "#view-zoom-in", "#view-reset"]:
            box = page.locator(button).bounding_box()
            check(box and min(box["width"], box["height"]) >= 30, f"Small desktop target: {button}")
        page.locator("#view-zoom-in").click()
        expect(page.locator("#view-zoom-label")).to_have_text("125%")
        page.locator("#view-zoom-out").click()
        expect(page.locator("#view-zoom-label")).to_have_text("100%")
        page.locator("#view-zoom-in").click()
        page.locator("#view-reset").click()
        page.wait_for_function("remoteView.zoom === 1 && remoteView.panX === 0 && remoteView.panY === 0")
        page.locator("#view-mode").click()
        expect(page.locator("#view-mode")).to_have_text("Width")
        page.locator("#view-mode").click()
        expect(page.locator("#view-mode")).to_have_text("Fit")
        check(page.locator("#sidebar").evaluate("e => getComputedStyle(e).transitionProperty === 'transform, visibility'"), "Panel must animate position, not layout")
        expect(page.locator("#sidebar-fps")).not_to_be_visible()
        page.locator("#connection-details > summary").press("Enter")
        expect(page.locator("#sidebar-fps")).to_be_visible()
        page.locator("#connection-details > summary").press("Enter")
        expect(page.locator("#sidebar-fps")).not_to_be_visible()
        page.locator("#capture-details > summary").press("Enter")
        expect(page.locator("#perf-region")).to_be_visible()
        page.locator("#capture-details > summary").press("Enter")
        expect(page.locator("#perf-region")).not_to_be_visible()
        page.evaluate("""() => {
            window.__testPublicResponses = [];
            window.inputSocket.addEventListener('message', event => {
                const value = JSON.parse(event.data);
                if (value.type === 'public_url_response') window.__testPublicResponses.push(value);
            });
        }""")
        previous_requests = page.evaluate("window.__testLocalRequests.length")
        page.locator("#public-link-toggle").press("Enter")
        expect(page.locator("#public-link-section")).to_be_visible()
        page.wait_for_function("previous => window.__testLocalRequests.length > previous", arg=previous_requests)
        page.wait_for_function("window.__testPublicResponses.length > 0")
        response = page.evaluate("window.__testPublicResponses.at(-1)")
        expect(page.locator("#link-text")).to_have_text(response["url"] if response["success"] else response["error"])
        page.locator("#public-link-toggle").press("Enter")
        expect(page.locator("#public-link-section")).not_to_be_visible()
        page.locator("#device-settings-toggle").click()
        expect(page.locator("#device-settings-panel")).to_be_visible()
        expect(page.locator("#camera-select-trigger")).to_be_visible()
        expect(page.locator("#mic-select")).to_be_visible()
        check(page.locator("#sidebar").evaluate("e => e.scrollWidth <= e.clientWidth"), "Device settings overflow")
        page.locator("#device-settings-toggle").click()
        expect(page.locator("#device-settings-panel")).not_to_be_visible()
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
            controls = page.locator("#view-controls").bounding_box()
            check(controls and controls["x"] >= 0 and controls["x"] + controls["width"] <= width, f"View controls overflow at {width}px: {controls}")
            expect(page.locator("#view-zoom-label")).to_be_visible()
        page.set_viewport_size({"width": 390, "height": 844})
        page.locator("#sidebar [data-open-support]").click()
        page.wait_for_function("document.querySelector('#support-qr').complete")
        page.locator("#support-dialog").screenshot(path=str(screenshots / "viewer-support-mobile.png"))
        page.locator("#support-dialog .dialog-close").press("Escape")
        expect(page.locator("#sidebar [data-open-support]")).to_be_focused()
        check(page.evaluate("window.__testInput.filter(e => e.action === 'key').length") == 2, "Support dialog keys leaked to host")
        page.locator(".deck-close").click()
        page.locator("#sidebar-trigger").click()
        expect(page.locator("#sidebar-trigger")).to_have_attribute("aria-expanded", "true")
        check(page.locator("#sidebar .sidebar-content").evaluate("e => e.scrollTop === 0"), "Reopening controls must return to Fit and zoom")
        fit = page.locator("#view-mode").bounding_box()
        check(fit and 0 <= fit["y"] <= 200, "Fit must be within reach when the hand opens")
        page.locator(".deck-close").scroll_into_view_if_needed()
        page.locator("#sidebar").evaluate("e => { e.scrollTop = 0; }")
        page.locator("#sidebar").screenshot(path=str(screenshots / "viewer-controls-mobile.png"))
        page.emulate_media(reduced_motion="reduce")
        check(page.locator("#sidebar").evaluate("e => getComputedStyle(e).transitionDuration === '0s'"), "Reduced motion is not respected")
        check(not errors, errors)
        browser.close()
    print(json.dumps({"authentication_form": "pass", "keyboard_panel_navigation": "pass", "native_disclosures_and_public_link_refresh": "pass", "compact_sidebar_without_free_badge": "pass", "fit_zoom_reset": "pass", "optional_support": "pass", "quality_manual_auto": "pass", "reconnects_without_duplicate_input": 5, "mobile_tablet_widths": [320, 390, 768, 820, 844, 1280], "reduced_motion": "pass", "browser_errors": errors}, indent=2))


if __name__ == "__main__":
    main()

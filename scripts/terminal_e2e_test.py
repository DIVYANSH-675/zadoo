"""Playwright E2E test for the Zadoo terminal keyboard path."""
from __future__ import annotations

import argparse
import time

from playwright.sync_api import sync_playwright


def _norm(text: str) -> str:
    return (text or "").replace("\r\n", "\n").replace("\r", "\n")


def _exact_count(text: str, expected: str) -> int:
    return sum(1 for line in _norm(text).split("\n") if line.strip() == expected)


def _screen_text(screen) -> str:
    try:
        return screen.inner_text(timeout=4000)
    except Exception:
        return ""


def _wait_count(screen, expected: str, min_count: int, seconds: float = 8.0):
    deadline = time.time() + seconds
    last = ""
    while time.time() < deadline:
        last = _screen_text(screen)
        if _exact_count(last, expected) >= min_count:
            return True, last
        time.sleep(0.2)
    return False, last


def _run_case(label, name, page, screen, actions, expected, timeout=8):
    before = _exact_count(_screen_text(screen), expected)
    for kind, value in actions:
        if kind == "type":
            page.keyboard.type(value, delay=8)
        elif kind == "press":
            page.keyboard.press(value)
        elif kind == "sleep":
            time.sleep(value)
    ok, last = _wait_count(screen, expected, before + 1, timeout)
    print(f"{label}.{name}: {'PASS' if ok else 'FAIL'} expected_line={expected}")
    if not ok:
        print(f"TAIL: {last[-1600:]!r}")
    return ok


def _run_matrix(label, page, screen, token):
    print(f"===== MATRIX {label} TOKEN {token} =====")
    results = [
        _run_case(label, "basic_enter", page, screen, [("type", f"echo BASIC_{token}"), ("press", "Enter")], f"BASIC_{token}"),
        _run_case(label, "alphabet_lower", page, screen, [("type", f"echo ALPHA_LOWER_{token}_abcdefghijklmnopqrstuvwxyz"), ("press", "Enter")], f"ALPHA_LOWER_{token}_abcdefghijklmnopqrstuvwxyz"),
        _run_case(label, "alphabet_upper", page, screen, [("type", f"echo ALPHA_UPPER_{token}_ABCDEFGHIJKLMNOPQRSTUVWXYZ"), ("press", "Enter")], f"ALPHA_UPPER_{token}_ABCDEFGHIJKLMNOPQRSTUVWXYZ"),
        _run_case(label, "backspace", page, screen, [("type", f"echo BS_{token}_ABCX"), ("press", "Backspace"), ("type", "D"), ("press", "Enter")], f"BS_{token}_ABCD"),
        _run_case(label, "delete", page, screen, [("type", f"echo DEL_{token}_ABXCD"), ("press", "ArrowLeft"), ("press", "ArrowLeft"), ("press", "ArrowLeft"), ("press", "Delete"), ("press", "Enter")], f"DEL_{token}_ABCD"),
        _run_case(label, "arrow_left_insert", page, screen, [("type", f"echo LEFT_{token}_ACD"), ("press", "ArrowLeft"), ("press", "ArrowLeft"), ("type", "B"), ("press", "Enter")], f"LEFT_{token}_ABCD"),
        _run_case(label, "arrow_right_insert", page, screen, [("type", f"echo RIGHT_{token}_ABD"), ("press", "ArrowLeft"), ("type", "C"), ("press", "ArrowRight"), ("type", "E"), ("press", "Enter")], f"RIGHT_{token}_ABCDE"),
        _run_case(label, "home_prefix", page, screen, [("type", f"HOME_{token}"), ("press", "Home"), ("type", "echo "), ("press", "Enter")], f"HOME_{token}"),
        _run_case(label, "end_suffix", page, screen, [("type", f"echo END_{token}_A"), ("press", "Home"), ("press", "End"), ("type", "B"), ("press", "Enter")], f"END_{token}_AB"),
    ]

    hist = f"HIST_{token}"
    results.append(_run_case(label, "history_seed", page, screen, [("type", f"echo {hist}"), ("press", "Enter")], hist))
    before = _exact_count(_screen_text(screen), hist)
    page.keyboard.press("ArrowUp")
    page.keyboard.press("Enter")
    ok, last = _wait_count(screen, hist, before + 1)
    print(f"{label}.history_up: {'PASS' if ok else 'FAIL'} expected_line={hist}")
    if not ok:
        print(f"TAIL: {last[-1600:]!r}")
    results.append(ok)

    old = f"DOWN_OLD_{token}"
    new = f"DOWN_NEW_{token}"
    results.append(_run_case(label, "history_down_seed_old", page, screen, [("type", f"echo {old}"), ("press", "Enter")], old))
    results.append(_run_case(label, "history_down_seed_new", page, screen, [("type", f"echo {new}"), ("press", "Enter")], new))
    before = _exact_count(_screen_text(screen), new)
    for key in ("ArrowUp", "ArrowUp", "ArrowDown", "Enter"):
        page.keyboard.press(key)
        time.sleep(0.1)
    ok, last = _wait_count(screen, new, before + 1)
    print(f"{label}.history_down: {'PASS' if ok else 'FAIL'} expected_line={new}")
    if not ok:
        print(f"TAIL: {last[-1600:]!r}")
    results.append(ok)

    ctrl = f"CTRLC_{token}"
    before = _exact_count(_screen_text(screen), ctrl)
    page.keyboard.type("Start-Sleep -Seconds 10", delay=8)
    page.keyboard.press("Enter")
    time.sleep(0.8)
    page.keyboard.press("Control+C")
    time.sleep(0.4)
    page.keyboard.type(f"echo {ctrl}", delay=8)
    page.keyboard.press("Enter")
    ok, last = _wait_count(screen, ctrl, before + 1, 4)
    print(f"{label}.ctrl_c_interrupt: {'PASS' if ok else 'FAIL'} expected_line={ctrl}")
    if not ok:
        print(f"TAIL: {last[-1600:]!r}")
    results.append(ok)

    print(f"{label}.SUMMARY {sum(1 for r in results if r)} / {len(results)}")
    return all(results)


def _authenticate_in_browser(page, base: str, code: str):
    page.goto(f"{base}/", wait_until="domcontentloaded", timeout=15000)
    page.wait_for_timeout(1200)
    if page.locator("#auth-code").count():
        page.locator("#auth-code").fill(code)
        page.keyboard.press("Enter")
        page.wait_for_timeout(2500)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--url", default="http://127.0.0.1:6173")
    parser.add_argument("--code", default="ZADOO123")
    args = parser.parse_args()

    base = args.url.rstrip("/")
    token = str(int(time.time() * 1000))[-8:]
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1300, "height": 900})

        direct = context.new_page()
        _authenticate_in_browser(direct, base, args.code)
        direct.goto(f"{base}/terminal.html", wait_until="domcontentloaded", timeout=15000)
        direct.wait_for_selector(".xterm-helper-textarea", state="attached", timeout=10000)
        direct.locator(".xterm").click(force=True)
        ok_direct = _run_matrix("DIRECT", direct, direct.locator(".xterm-screen"), token)

        page = context.new_page()
        _authenticate_in_browser(page, base, args.code)
        page.evaluate("toggleTerminal()")
        page.wait_for_selector("#terminal-iframe", timeout=10000)
        page.wait_for_timeout(2500)
        frame = page.frame_locator("#terminal-iframe")
        frame.locator(".xterm-helper-textarea").wait_for(state="attached", timeout=10000)
        ok_panel = _run_matrix("PANEL", page, frame.locator(".xterm-screen"), token)

        browser.close()
    return 0 if ok_direct and ok_panel else 1


if __name__ == "__main__":
    raise SystemExit(main())

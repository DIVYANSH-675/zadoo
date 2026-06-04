"""Syntax-check inline JavaScript blocks in Zadoo HTML templates with Node.js."""
from __future__ import annotations

import shutil
import subprocess
import tempfile
from html.parser import HTMLParser
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = ROOT / "zadoo_vnc" / "templates"


class ScriptCollector(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self._collecting = False
        self._current: list[str] = []
        self.scripts: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() != "script":
            return
        attr_map = {str(k).lower(): str(v or "") for k, v in attrs}
        if attr_map.get("src"):
            return
        script_type = attr_map.get("type", "").lower()
        if script_type and script_type not in {"text/javascript", "application/javascript", "module"}:
            return
        self._collecting = True
        self._current = []

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() == "script" and self._collecting:
            self.scripts.append("".join(self._current))
            self._collecting = False
            self._current = []

    def handle_data(self, data: str) -> None:
        if self._collecting:
            self._current.append(data)


def main() -> None:
    node = shutil.which("node")
    if not node:
        print("SKIP: Node.js not found; template JavaScript syntax check skipped")
        return
    failures: list[str] = []
    with tempfile.TemporaryDirectory(prefix="zadoo-js-") as temp_dir:
        temp_root = Path(temp_dir)
        for html_path in sorted(TEMPLATE_DIR.glob("*.html")):
            parser = ScriptCollector()
            parser.feed(html_path.read_text(encoding="utf-8"))
            for index, script in enumerate(parser.scripts, start=1):
                js_path = temp_root / f"{html_path.stem}-{index}.js"
                js_path.write_text(script, encoding="utf-8")
                result = subprocess.run([node, "--check", str(js_path)], capture_output=True, text=True)
                if result.returncode != 0:
                    failures.append(f"{html_path.name} script {index}: {result.stderr.strip() or result.stdout.strip()}")
    if failures:
        for failure in failures:
            print(f"FAIL: {failure}")
        raise SystemExit(1)
    print("OK: template JavaScript syntax")


if __name__ == "__main__":
    main()

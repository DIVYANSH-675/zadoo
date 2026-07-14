"""Reject unsafe or floating third-party references in GitHub workflows."""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_ROOT = ROOT / ".github" / "workflows"
USES_RE = re.compile(r"^\s*-?\s*uses:\s*([^\s#]+)", re.MULTILINE)
FULL_SHA_RE = re.compile(r"^[^@]+@[0-9a-fA-F]{40}$")


def fail(message: str) -> None:
    raise SystemExit(f"FAIL: {message}")


def main() -> None:
    workflows = sorted((*WORKFLOW_ROOT.glob("*.yml"), *WORKFLOW_ROOT.glob("*.yaml")))
    if not workflows:
        fail("no GitHub Actions workflows found")
    checked = 0
    for path in workflows:
        text = path.read_text(encoding="utf-8")
        relative = path.relative_to(ROOT)
        if "pull_request_target:" in text:
            fail(f"{relative} uses privileged pull_request_target")
        permissions = re.search(
            r"(?m)^permissions:\s*$\n\s+contents:\s*(read|write)\s*$", text
        )
        if not permissions:
            fail(f"{relative} must declare explicit contents permissions")
        if permissions.group(1) == "write":
            if re.search(r"(?m)^\s+(?:push|pull_request):\s*$", text):
                fail(f"{relative} grants contents: write to an automatic trigger")
            required_release_guards = (
                "on:\n  workflow_dispatch:\n\npermissions:",
                "environment: release",
                'if ($env:GITHUB_REF_TYPE -ne "tag")',
                'if ($manifest.signed -ne $true)',
                "gh release create",
            )
            if path.name != "windows-x64-release.yml" or any(
                marker not in text for marker in required_release_guards
            ):
                fail(f"{relative} has contents: write without the signed tagged-release guards")
        for reference in USES_RE.findall(text):
            if reference.startswith("./"):
                continue
            checked += 1
            if not FULL_SHA_RE.fullmatch(reference):
                fail(f"{relative} has an action that is not pinned to a full commit SHA: {reference}")
    if checked == 0:
        fail("no third-party action references found")
    print(f"OK: {len(workflows)} workflow(s), {checked} action reference(s) pinned to full commit SHAs")


if __name__ == "__main__":
    main()

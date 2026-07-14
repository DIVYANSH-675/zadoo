# Contributing

Zadoo targets Windows 10/11 x64 and Python 3.11.9. Keep runtime behavior, public routes, settings
compatibility, and exact user-facing error messages stable unless a change is explicitly approved.

## Before opening a pull request

```powershell
python -m ruff check .
python -m compileall -q zadoo_vnc scripts
python scripts\smoke_test.py
python scripts\check_template_js.py
python scripts\check_workflow_pins.py
```

For UI, authentication, or terminal changes, also run the live and Playwright checks documented
in `README.md`. For packaging changes, run a complete `-NoSelfSign` build and inspect
`dist\release-manifest.json`.

## Updating dependencies

Direct versions live in `pyproject.toml`, `build_requirements.txt`, and `ci_requirements.txt`.
Every direct change must regenerate its hash-locked graph using Python 3.11.9 and pip-tools 7.5.3:

```powershell
py -3.11-64 -m venv .lock-tools
.\.lock-tools\Scripts\python.exe -m pip install pip-tools==7.5.3
.\.lock-tools\Scripts\python.exe -m piptools compile --generate-hashes --strip-extras --no-emit-index-url --no-emit-trusted-host --output-file requirements-runtime.lock pyproject.toml
.\.lock-tools\Scripts\python.exe -m piptools compile --allow-unsafe --generate-hashes --no-emit-index-url --no-emit-trusted-host --output-file requirements-build.lock build_requirements.txt
.\.lock-tools\Scripts\python.exe -m piptools compile --allow-unsafe --generate-hashes --no-emit-index-url --no-emit-trusted-host --output-file ci_requirements.lock ci_requirements.txt
```

Audit both shipped and build dependencies after regeneration. Do not suppress a vulnerability
without documenting why it is unreachable and obtaining review.

## Pull request evidence

Include the root cause, user impact, files changed, validation commands, and before/after
performance measurements for hot-path work. State hardware-dependent or signed-release checks
that could not be performed. Never paste real secrets into issues, logs, screenshots, or commits.

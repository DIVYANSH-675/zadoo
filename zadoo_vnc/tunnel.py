"""Cloudflare tunnel and optional email notification management."""
from __future__ import annotations

import logging
import os
import queue
import re
import subprocess
import threading
import time
from pathlib import Path

import resend

from .config import resource_path, windows_system_executable
from .network import get_local_ip
from .settings import get_settings_store


class CloudflareTunnelManager:
    def __init__(self, primary_port):
        self.primary_port = primary_port
        self.primary_tunnel_process = None
        self.primary_public_url = None
        configured_path = os.getenv("ZADOO_CLOUDFLARED_PATH")
        if configured_path is not None and not configured_path.strip():
            raise ValueError("ZADOO_CLOUDFLARED_PATH must not be empty")
        self.cloudflared_path = str(
            Path(configured_path.strip()).expanduser().resolve()
            if configured_path is not None
            else resource_path("cloudflared.exe")
        )
        self.protocol = os.getenv("ZADOO_CLOUDFLARED_PROTOCOL", "http2").strip().lower()
        if self.protocol not in {"http2", "quic"}:
            raise ValueError(
                "ZADOO_CLOUDFLARED_PROTOCOL must be http2 or quic; "
                f"got {self.protocol!r}"
            )
        self.last_error = None

        store = get_settings_store()
        settings = store.load()
        self.resend_api_key = os.getenv("RESEND_API_KEY", "").strip()
        self.resend_from = os.getenv("RESEND_FROM", "").strip()
        self.email_to = str(settings["email_to"] or os.getenv("EMAIL_TO") or "").strip()

        self.last_email_message = self._email_configuration_error()
        self._notified_url = None
        self._tunnel_lock = threading.RLock()
        self._signature_cache = None

    def _verify_cloudflared_signature(self):
        path = Path(self.cloudflared_path)
        if not path.is_file():
            raise FileNotFoundError(f"cloudflared.exe not found: {path}")
        # Get-AuthenticodeSignature spawns PowerShell (~1-2s). The binary does not change
        # within a session, so cache the verdict per (path, mtime) — this keeps tunnel
        # restarts/refreshes from paying that cost every time.
        mtime = path.stat().st_mtime_ns
        cache = self._signature_cache
        if cache and cache[0] == self.cloudflared_path and cache[1] == mtime:
            return
        quoted_path = self.cloudflared_path.replace("'", "''")
        command = f"(Get-AuthenticodeSignature -LiteralPath '{quoted_path}').Status"
        result = subprocess.run(
            [
                windows_system_executable("WindowsPowerShell", "v1.0", "powershell.exe"),
                "-NoProfile",
                "-Command",
                command,
            ],
            capture_output=True,
            text=True,
            timeout=10,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        status = (result.stdout or "").strip()
        if result.returncode != 0:
            detail = (result.stderr or status or f"exit code {result.returncode}").strip()
            raise RuntimeError(f"cloudflared signature check failed: {detail}")
        if status.lower() != "valid":
            raise RuntimeError(f"cloudflared signature check failed: {status or 'no status returned'}")
        self._signature_cache = (self.cloudflared_path, mtime)

    def start_tunnel(self, port):
        """Start a Cloudflare tunnel for a specific port."""
        self._verify_cloudflared_signature()
        process = None
        startup_complete = threading.Event()
        try:
            cmd = [
                self.cloudflared_path,
                "tunnel",
                "--no-autoupdate",
                "--protocol",
                self.protocol,
                "--url",
                f"http://localhost:{port}",
            ]
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=subprocess.CREATE_NO_WINDOW,
                encoding="utf-8",
                errors="strict",
            )

            logging.info("Started cloudflared for port %s with PID %s", port, process.pid)

            output_queue = queue.Queue()
            def log_output(pipe, out_q):
                try:
                    with pipe:
                        for line in iter(pipe.readline, ""):
                            ln = line.strip()
                            if not ln:
                                continue
                            logging.info("cloudflared[%s]: %s", port, ln)
                            if not startup_complete.is_set():
                                out_q.put_nowait(ln)
                except Exception as e:
                    if not startup_complete.is_set():
                        out_q.put_nowait(e)

            threading.Thread(target=log_output, args=(process.stdout, output_queue), daemon=True).start()

            print(f" Looking for public URL for port {port}...")
            deadline = time.monotonic() + 20
            while time.monotonic() < deadline:
                try:
                    output = output_queue.get(timeout=1.0)
                    if isinstance(output, Exception):
                        raise RuntimeError(f"cloudflared output read failed: {output}") from output
                    ln = output
                    if "trycloudflare.com" in ln:
                        url_match = re.search(r"https://[a-zA-Z0-9-]+\.trycloudflare\.com", ln)
                        if url_match:
                            url = url_match.group(0)
                            startup_complete.set()
                            logging.info("Found public URL for port %s: %s", port, url)
                            print(f" Found public URL for port {port}: {url}")
                            return process, url
                except queue.Empty:
                    pass
                if process.poll() is not None:
                    raise RuntimeError(
                        f"cloudflared exited with code {process.returncode} before providing a public URL"
                    )
            raise TimeoutError("cloudflared did not provide a public URL within 20 seconds")

        except Exception:
            startup_complete.set()
            if process and process.poll() is None:
                process.kill()
                process.wait(timeout=5)
            raise

    def start_primary_tunnel(self):
        """Start the primary tunnel once and return the current public URL."""
        with self._tunnel_lock:
            if self.primary_tunnel_process and self.primary_tunnel_process.poll() is None:
                if not self.primary_public_url:
                    raise RuntimeError("Cloudflare tunnel process is running without a public URL")
                logging.info("Cloudflare tunnel already running for port %s", self.primary_port)
                return self.primary_public_url

            print(f" Starting Cloudflare tunnel for port {self.primary_port}...")
            self.last_error = None
            try:
                self.primary_tunnel_process, self.primary_public_url = self.start_tunnel(self.primary_port)
                print("=" * 80)
                print(" PRIMARY PORT PUBLIC URL READY!")
                print(f" Port {self.primary_port}: {self.primary_public_url}")
                print("=" * 80)
                self.notify_public_url(self.primary_public_url)
                return self.primary_public_url
            except Exception as exc:
                self.last_error = str(exc)
                self.primary_tunnel_process = None
                self.primary_public_url = None
                logging.error("Tunnel startup failed: %s", self.last_error, exc_info=True)
                raise RuntimeError(f"Cloudflare tunnel startup failed: {exc}") from exc

    def refresh_tunnel(self):
        """Restart the primary tunnel and notify once."""
        with self._tunnel_lock:
            self._stop_process()
            self.primary_public_url = None
            self._notified_url = None
            return self.start_primary_tunnel()

    def get_current_url(self):
        """Get the URL for the active tunnel."""
        return self.primary_public_url

    def notify_public_url(self, url):
        """Send an email notification when Resend configuration is available."""
        try:
            store = get_settings_store()
            settings = store.load(reload=True)
            self.resend_api_key = os.getenv("RESEND_API_KEY", "").strip()
            self.resend_from = os.getenv("RESEND_FROM", "").strip()
            self.email_to = str(settings["email_to"] or os.getenv("EMAIL_TO") or "").strip()
            configuration_error = self._email_configuration_error()
            if configuration_error:
                self.last_email_message = configuration_error
                return False

            if url == self._notified_url:
                self.last_email_message = "Email already sent for current URL"
                return True
            private_ip = get_local_ip()

            subject = "Zadoo Public Link"
            html_body = (
                f"<p>Private IP: <strong>{private_ip}:{self.primary_port}</strong><br/>"
                f"Public URL: <a href='{url}'>{url}</a></p>"
            )

            resend.api_key = self.resend_api_key
            resend.Emails.send({
                "from": self.resend_from,
                "to": self.email_to,
                "subject": subject,
                "html": html_body,
            })

            self._notified_url = url
            self.last_email_message = f"Email sent to {self.email_to}"
            print(f" Email sent to: {self.email_to}")
            return True

        except Exception as e:
            logging.error("notify_public_url failed: %s", e, exc_info=True)
            self.last_email_message = f"Email failed: {e}"
            return False

    def _email_configuration_error(self):
        missing = []
        if not self.resend_api_key:
            missing.append("RESEND_API_KEY")
        if not self.resend_from:
            missing.append("RESEND_FROM")
        if not self.email_to:
            missing.append("Email To")
        return f"Email notification disabled; missing: {', '.join(missing)}" if missing else None

    def _stop_process(self):
        process = self.primary_tunnel_process
        if process and process.poll() is None:
            process.kill()
            process.wait(timeout=5)
        self.primary_tunnel_process = None

    def cleanup(self):
        """Clean up the tunnel process owned by this manager."""
        print(" Cleaning up owned tunnel process...")

        self._stop_process()
        self.primary_public_url = None
        print(" Owned tunnel cleanup complete")

"""Cloudflare tunnel and optional email notification management."""
from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
import time
import urllib.request
from pathlib import Path

from .config import RUNTIME_DIR, env_int, resource_path
from .dependencies import resend
from .logging_utils import _log_fallback
from .network import get_local_ip
from .settings import get_settings_store, settings_dir


class CloudflareTunnelManager:
    DEFAULT_RESEND_FROM = "onboarding@resend.dev"
    EMAIL_NOT_SET_MESSAGE = "Email not Set"

    def __init__(self, primary_port):
        self.primary_port = primary_port
        self.current_port = primary_port
        self.primary_tunnel_process = None
        self.primary_public_url = None
        self.cloudflared_path = None
        self.last_notified_url = None

        # Resend is optional and user-configured through installed settings.
        # Environment values remain a developer fallback before setup.
        store = get_settings_store()
        settings = store.load()
        self.resend_api_key = store.get_resend_api_key() or os.getenv("RESEND_API_KEY")
        self.resend_from = self.DEFAULT_RESEND_FROM
        self.email_to = str(settings.get("email_to") or os.getenv("EMAIL_TO") or "").strip()

        self.last_email_status = None
        self.last_email_message = None if self.resend_api_key and self.email_to else self.EMAIL_NOT_SET_MESSAGE
        self._notified_states = set()
        self._tunnel_lock = threading.RLock()

        # Email port configuration (which port to include in email).
        self.email_port = None

    def _verify_cloudflared_signature(self):
        if os.name != "nt" or os.getenv("ZADOO_SKIP_CLOUDFLARED_SIGNATURE_CHECK", "").strip().lower() in {"1", "true", "yes", "on"}:
            return True
        if not self.cloudflared_path or not os.path.exists(self.cloudflared_path):
            return False
        try:
            quoted_path = self.cloudflared_path.replace("'", "''")
            command = f"(Get-AuthenticodeSignature -LiteralPath '{quoted_path}').Status"
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", command],
                capture_output=True,
                text=True,
                timeout=10,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode == 0 and (result.stdout or "").strip().lower() == "valid":
                return True
            logging.error("cloudflared signature verification failed: %s %s", result.stdout.strip(), result.stderr.strip())
        except Exception:
            logging.error("cloudflared signature verification failed", exc_info=True)
        return False

    def _cleanup_stale_cloudflared(self, port):
        if os.name != "nt":
            return
        try:
            target = f"http://localhost:{int(port)}"
        except Exception:
            target = f"http://localhost:{port}"
        command = (
            "Get-CimInstance Win32_Process -Filter \"name = 'cloudflared.exe'\" | "
            "Select-Object ProcessId,CommandLine | ConvertTo-Json -Compress"
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", command],
                capture_output=True,
                text=True,
                timeout=5,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            if result.returncode != 0 or not (result.stdout or "").strip():
                return
            import json

            data = json.loads(result.stdout)
            processes = data if isinstance(data, list) else [data]
            for item in processes:
                command_line = str(item.get("CommandLine") or "")
                pid = int(item.get("ProcessId") or 0)
                if pid > 0 and target in command_line and " tunnel " in f" {command_line} ":
                    try:
                        subprocess.run(
                            ["taskkill", "/PID", str(pid), "/T", "/F"],
                            capture_output=True,
                            timeout=5,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                        )
                        logging.info("Stopped stale cloudflared process PID %s for %s", pid, target)
                    except Exception:
                        logging.debug("Failed to stop stale cloudflared PID %s", pid, exc_info=True)
        except Exception:
            logging.debug("Stale cloudflared cleanup failed", exc_info=True)

    def _cloudflared_download_url(self):
        forced = os.getenv("ZADOO_CLOUDFLARED_ARCH", "").strip().lower()
        arch = "386" if forced in {"x86", "386", "32"} or sys.maxsize <= 2**32 else "amd64"
        return f"https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-{arch}.exe"

    def _cloudflared_candidates(self):
        env_path = os.getenv("ZADOO_CLOUDFLARED_PATH", "").strip()
        if env_path:
            yield Path(env_path)
        yield resource_path("cloudflared.exe")
        yield RUNTIME_DIR / "cloudflared.exe"
        yield Path(os.getcwd()) / "cloudflared.exe"
        yield settings_dir() / "bin" / "cloudflared.exe"

    def _find_cloudflared(self):
        for path in self._cloudflared_candidates():
            try:
                if path.exists() and path.is_file():
                    return str(path)
            except Exception:
                continue
        return None

    def download_cloudflared(self):
        """Download cloudflared if not present."""
        existing = self._find_cloudflared()
        if existing:
            self.cloudflared_path = existing
            print(" Using existing cloudflared.exe")
            return True

        try:
            print(" Downloading cloudflared...")
            url = self._cloudflared_download_url()
            timeout = env_int("ZADOO_CLOUDFLARED_DOWNLOAD_TIMEOUT", 30, 1, 600)
            target = settings_dir() / "bin" / "cloudflared.exe"
            target.parent.mkdir(parents=True, exist_ok=True)

            with urllib.request.urlopen(url, timeout=timeout) as response, open(target, "wb") as out_file:
                out_file.write(response.read())

            self.cloudflared_path = str(target)
            print(" Downloaded cloudflared.exe")
            return True

        except Exception as e:
            print(f" Failed to download cloudflared: {e}")
            return False

    def start_tunnel(self, port):
        """Start a Cloudflare tunnel for a specific port."""
        if not self.cloudflared_path or not os.path.exists(self.cloudflared_path):
            if not self.download_cloudflared():
                return None, None
        if not self._verify_cloudflared_signature():
            print(" Refusing to run cloudflared.exe because signature verification failed")
            return None, None

        try:
            self._cleanup_stale_cloudflared(port)
            cmd = [self.cloudflared_path, "tunnel", "--url", f"http://localhost:{port}"]
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                encoding="utf-8",
                errors="replace",
            )

            logging.info("Started cloudflared for port %s with PID %s", port, process.pid)

            output_queue = queue.Queue()

            def log_output(pipe, out_q):
                try:
                    with pipe:
                        for line in iter(pipe.readline, ""):
                            ln = (line or "").strip()
                            if not ln:
                                continue
                            logging.info("cloudflared[%s]: %s", port, ln)
                            try:
                                out_q.put_nowait(ln)
                            except Exception:
                                pass
                except Exception as e:
                    logging.error("Error reading cloudflared output for port %s: %s", port, e)

            threading.Thread(target=log_output, args=(process.stdout, output_queue), daemon=True).start()

            print(f" Looking for public URL for port {port}...")
            start_time = time.time()
            while time.time() - start_time < 20:
                try:
                    ln = output_queue.get(timeout=1.0)
                    if "trycloudflare.com" in ln:
                        import re

                        url_match = re.search(r"https?://[a-zA-Z0-9-]+\.trycloudflare\.com", ln)
                        if url_match:
                            url = url_match.group(0)
                            logging.info("Found public URL for port %s: %s", port, url)
                            print(f" Found public URL for port {port}: {url}")
                            return process, url
                except Exception:
                    pass
                if process.poll() is not None:
                    logging.error("cloudflared process for port %s terminated unexpectedly.", port)
                    break
            logging.warning("Could not find public URL for port %s within 20 seconds.", port)
            return process, None

        except Exception as e:
            logging.error(" Failed to start tunnel for port %s: %s", port, e, exc_info=True)
            print(f" Failed to start tunnel for port {port}: {e}")
            return None, None

    def start_primary_tunnel(self):
        """Start the primary tunnel once and return the current public URL."""
        with self._tunnel_lock:
            if self.primary_tunnel_process and self.primary_tunnel_process.poll() is None:
                logging.info("Cloudflare tunnel already running for port %s", self.primary_port)
                return self.primary_public_url

            print(f" Starting Cloudflare tunnel for port {self.primary_port}...")
            self.current_port = self.primary_port
            self.primary_tunnel_process, self.primary_public_url = self.start_tunnel(self.primary_port)

            if self.primary_public_url:
                print("=" * 80)
                print(" PRIMARY PORT PUBLIC URL READY!")
                print(f" Port {self.primary_port}: {self.primary_public_url}")
                print("=" * 80)
                try:
                    self.notify_public_url(self.primary_public_url, self.email_port)
                except Exception:
                    logging.warning("Failed to send email notification for tunnel", exc_info=True)
                return self.primary_public_url

            try:
                _log_fallback("tunnel.start_primary", "private_ip_notification", "public_url_unavailable")
                self.notify_public_url(None, self.email_port)
            except Exception:
                logging.warning("Failed to send private IP notification for tunnel", exc_info=True)
            return None

    def refresh_tunnel(self, new_port=None):
        """Restart the primary tunnel on a new port and notify once."""
        try:
            with self._tunnel_lock:
                if new_port is not None:
                    print(f" Refreshing tunnel to new port {new_port}...")
                    self.primary_port = new_port
                    self.current_port = new_port
                else:
                    print(f" Refreshing tunnel on current port {self.primary_port}...")

                if self.primary_tunnel_process:
                    print(f" Stopping old tunnel process (PID: {self.primary_tunnel_process.pid})")
                    try:
                        self.primary_tunnel_process.terminate()
                        self.primary_tunnel_process.wait(timeout=5)
                        print(" Old tunnel process stopped")
                    except Exception as e:
                        print(f" Error stopping tunnel: {e}")
                        try:
                            self.primary_tunnel_process.kill()
                            print(" Old tunnel process force killed")
                        except Exception:
                            print(" Could not kill old tunnel process")
                else:
                    print(" No old tunnel process to stop")

                self.primary_tunnel_process = None
                self.primary_public_url = None
                self._notified_states.clear()

                print(" Starting new tunnel...")
                result = self.start_primary_tunnel()
                print(f" start_primary_tunnel() returned: {result}")
                return result
        except Exception as e:
            print(f" Error in refresh_tunnel: {e}")
            logging.error("Failed to refresh primary tunnel", exc_info=True)
            return None

    def get_current_url(self):
        """Get the URL for the active tunnel."""
        return self.primary_public_url

    def notify_public_url(self, url, port):
        """Send an email notification when Resend configuration is available."""
        try:
            store = get_settings_store()
            settings = store.load(reload=True)
            self.resend_api_key = store.get_resend_api_key() or os.getenv("RESEND_API_KEY")
            self.email_to = str(settings.get("email_to") or os.getenv("EMAIL_TO") or "").strip()
            if not (resend and self.resend_api_key and self.resend_from and self.email_to):
                self.last_notified_url = url
                self.last_email_status = False
                self.last_email_message = self.EMAIL_NOT_SET_MESSAGE
                return False

            private_ip = get_local_ip()
            notify_port = port or self.primary_port
            state_key = url or f"private:{private_ip}:{notify_port}"
            if state_key in self._notified_states:
                self.last_email_status = True
                self.last_email_message = "Email already sent for current URL"
                return True

            if url:
                subject = "Zadoo Public Link"
                html_body = (
                    f"<p>Private IP: <strong>{private_ip}:{notify_port}</strong><br/>"
                    f"Public URL: <a href='{url}'>{url}</a></p>"
                )
            else:
                subject = "Zadoo Access"
                html_body = (
                    f"<p>Private IP: <strong>{private_ip}:{notify_port}</strong><br/>"
                    "Public URL: <em>not available</em></p>"
                )

            resend.api_key = self.resend_api_key
            resend.Emails.send({
                "from": self.resend_from,
                "to": self.email_to,
                "subject": subject,
                "html": html_body,
            })

            self._notified_states.add(state_key)
            self.last_notified_url = url
            self.last_email_status = True
            self.last_email_message = f"Email sent to {self.email_to}"
            print(f" Email sent to: {self.email_to}")
            return True

        except Exception as e:
            logging.error("notify_public_url failed: %s", e, exc_info=True)
            self.last_email_status = False
            self.last_email_message = "notify_public_url failed"
            return False

    def cleanup(self):
        """Clean up the tunnel process owned by this manager."""
        print(" Cleaning up owned tunnel process...")

        process = self.primary_tunnel_process
        if process:
            try:
                print(" Stopping Primary tunnel...")
                process.terminate()
                process.wait(timeout=5)
                print(" Primary tunnel stopped")
            except Exception:
                try:
                    process.kill()
                    print(" Primary tunnel force killed")
                except Exception:
                    print(" Could not stop Primary tunnel")

        self.primary_tunnel_process = None
        self.primary_public_url = None
        print(" Owned tunnel cleanup complete")

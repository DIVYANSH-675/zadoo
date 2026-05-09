"""Cloudflare tunnel and optional email notification management."""
from __future__ import annotations

import logging
import os
import queue
import subprocess
import threading
import time
import urllib.request

from .dependencies import resend
from .network import get_local_ip


class CloudflareTunnelManager:
    def __init__(self, primary_port):
        self.primary_port = primary_port
        self.current_port = primary_port
        self.primary_tunnel_process = None
        self.primary_public_url = None
        self.cloudflared_path = None
        self.last_notified_url = None

        # Resend is optional and entirely env-driven. No API key or default
        # recipient is bundled in source.
        self.resend_api_key = os.getenv("RESEND_API_KEY")
        self.resend_from = os.getenv("RESEND_FROM", "onboarding@resend.dev")
        self.email_to = os.getenv("EMAIL_TO") or os.getenv("GMAIL_TO")

        self.last_email_status = None
        self.last_email_message = None
        self._notified_states = set()
        self._tunnel_lock = threading.RLock()

        # Email port configuration (which port to include in email).
        self.email_port = None

    def download_cloudflared(self):
        """Download cloudflared if not present."""
        self.cloudflared_path = os.path.join(os.getcwd(), "cloudflared.exe")

        if os.path.exists(self.cloudflared_path):
            print("✅ Using existing cloudflared.exe")
            return True

        try:
            print("📥 Downloading cloudflared...")
            url = "https://github.com/cloudflare/cloudflared/releases/latest/download/cloudflared-windows-amd64.exe"

            with urllib.request.urlopen(url) as response, open(self.cloudflared_path, "wb") as out_file:
                out_file.write(response.read())

            print("✅ Downloaded cloudflared.exe")
            return True

        except Exception as e:
            print(f"❌ Failed to download cloudflared: {e}")
            return False

    def start_tunnel(self, port):
        """Start a Cloudflare tunnel for a specific port."""
        if not self.cloudflared_path or not os.path.exists(self.cloudflared_path):
            if not self.download_cloudflared():
                return None, None

        try:
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

            print(f"🔍 Looking for public URL for port {port}...")
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
                            print(f"✅ Found public URL for port {port}: {url}")
                            return process, url
                except Exception:
                    pass
                if process.poll() is not None:
                    logging.error("cloudflared process for port %s terminated unexpectedly.", port)
                    break
            logging.warning("Could not find public URL for port %s within 20 seconds.", port)
            return process, None

        except Exception as e:
            logging.error("❌ Failed to start tunnel for port %s: %s", port, e, exc_info=True)
            print(f"❌ Failed to start tunnel for port {port}: {e}")
            return None, None

    def start_primary_tunnel(self):
        """Start the primary tunnel once and return the current public URL."""
        with self._tunnel_lock:
            if self.primary_tunnel_process and self.primary_tunnel_process.poll() is None:
                logging.info("Cloudflare tunnel already running for port %s", self.primary_port)
                return self.primary_public_url

            print(f"🚀 Starting Cloudflare tunnel for port {self.primary_port}...")
            self.current_port = self.primary_port
            self.primary_tunnel_process, self.primary_public_url = self.start_tunnel(self.primary_port)

            if self.primary_public_url:
                print("🌍" * 80)
                print("🌍 PRIMARY PORT PUBLIC URL READY!")
                print(f"🔗 Port {self.primary_port}: {self.primary_public_url}")
                print("🌍" * 80)
                try:
                    self.notify_public_url(self.primary_public_url, self.email_port)
                except Exception:
                    logging.warning("Failed to send email notification for tunnel", exc_info=True)
                return self.primary_public_url

            try:
                self.notify_public_url(None, self.email_port)
            except Exception:
                logging.warning("Failed to send private IP notification for tunnel", exc_info=True)
            return None

    def refresh_tunnel(self, new_port=None):
        """Restart the primary tunnel on a new port and notify once."""
        try:
            with self._tunnel_lock:
                if new_port is not None:
                    print(f"🔁 Refreshing tunnel to new port {new_port}...")
                    self.primary_port = new_port
                    self.current_port = new_port
                else:
                    print(f"🔁 Refreshing tunnel on current port {self.primary_port}...")

                if self.primary_tunnel_process:
                    print(f"🛑 Stopping old tunnel process (PID: {self.primary_tunnel_process.pid})")
                    try:
                        self.primary_tunnel_process.terminate()
                        self.primary_tunnel_process.wait(timeout=5)
                        print("✅ Old tunnel process stopped")
                    except Exception as e:
                        print(f"⚠️ Error stopping tunnel: {e}")
                        try:
                            self.primary_tunnel_process.kill()
                            print("✅ Old tunnel process force killed")
                        except Exception:
                            print("❌ Could not kill old tunnel process")
                else:
                    print("ℹ️ No old tunnel process to stop")

                self.primary_tunnel_process = None
                self.primary_public_url = None
                self._notified_states.clear()

                print("🚀 Starting new tunnel...")
                result = self.start_primary_tunnel()
                print(f"🚀 start_primary_tunnel() returned: {result}")
                return result
        except Exception as e:
            print(f"❌ Error in refresh_tunnel: {e}")
            logging.error("Failed to refresh primary tunnel", exc_info=True)
            return None

    def get_current_url(self):
        """Get the URL for the active tunnel."""
        return self.primary_public_url

    def notify_public_url(self, url, port):
        """Send an email notification when Resend configuration is available."""
        try:
            if not (resend and self.resend_api_key and self.resend_from and self.email_to):
                self.last_notified_url = url
                self.last_email_status = False
                self.last_email_message = "Email disabled (set RESEND_API_KEY and EMAIL_TO)"
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
            print(f"✅ Email sent to: {self.email_to}")
            return True

        except Exception as e:
            logging.error("notify_public_url failed: %s", e, exc_info=True)
            self.last_email_status = False
            self.last_email_message = "notify_public_url failed"
            return False

    def cleanup(self):
        """Clean up tunnel processes."""
        print("🧹 Cleaning up all tunnel processes...")

        for process_name, process in [("Primary", self.primary_tunnel_process)]:
            if process:
                try:
                    print(f"🛑 Stopping {process_name} tunnel...")
                    process.terminate()
                    process.wait(timeout=5)
                    print(f"✅ {process_name} tunnel stopped")
                except Exception:
                    try:
                        process.kill()
                        print(f"✅ {process_name} tunnel force killed")
                    except Exception:
                        print(f"⚠️ Could not stop {process_name} tunnel")

        self.primary_tunnel_process = None
        self.primary_public_url = None
        print("🧹 All tunnels cleaned up")

"""Browser-based OAuth sign-in (authorization code + PKCE) over a loopback HTTPS listener.

Used by ``office-connect login --browser`` for tenants that block the
device-code flow. MSAL builds the authorize URL (with PKCE + state); we open
it in the user's browser and catch the redirect on a short-lived HTTPS server
bound to the loopback interface only.

TLS for ``localhost``:

1. Explicit ``cert_file`` / ``key_file`` if given.
2. A cached cert under ``~/.config/office-connect/tls/`` if still valid.
3. ``mkcert`` (if installed) — produces a cert the browser already trusts.
4. A self-signed cert generated with ``cryptography`` — works, but the browser
   shows a one-time certificate warning that has to be clicked through.

The redirect URI (``https://localhost:8443/auth/callback`` by default) must be
registered on the Azure AD app under **Mobile and desktop applications**
(public client) so the code can be redeemed without a client secret.
"""

from __future__ import annotations

import datetime as _dt
import html
import ipaddress
import logging
import os
import shutil
import socket
import ssl
import subprocess
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Callable
from urllib.parse import parse_qs, urlsplit

logger = logging.getLogger(__name__)

DEFAULT_REDIRECT_HOST = "localhost"
DEFAULT_REDIRECT_PORT = 8443
DEFAULT_REDIRECT_PATH = "/auth/callback"
DEFAULT_TLS_DIR = "~/.config/office-connect/tls"
DEFAULT_TIMEOUT_S = 300.0

_CERT_NAME = "localhost.pem"
_KEY_NAME = "localhost-key.pem"
_SELF_SIGNED_DAYS = 825
_RENEW_BEFORE = _dt.timedelta(days=14)


class BrowserLoginError(RuntimeError):
    """Raised when the browser sign-in cannot complete."""


def build_redirect_uri(host: str = DEFAULT_REDIRECT_HOST,
                       port: int = DEFAULT_REDIRECT_PORT,
                       path: str = DEFAULT_REDIRECT_PATH) -> str:
    if not path.startswith("/"):
        path = "/" + path
    netloc = host if port == 443 else f"{host}:{port}"
    return f"https://{netloc}{path}"


# ── TLS certificate ──────────────────────────────────────────────────────

def _cert_still_valid(cert_path: Path) -> bool:
    try:
        from cryptography import x509
        cert = x509.load_pem_x509_certificate(cert_path.read_bytes())
    except Exception:
        return False
    not_after = cert.not_valid_after_utc
    return not_after - _dt.datetime.now(_dt.timezone.utc) > _RENEW_BEFORE


def _try_mkcert(cert_path: Path, key_path: Path) -> bool:
    mkcert = shutil.which("mkcert")
    if not mkcert:
        return False
    try:
        subprocess.run(
            [mkcert, "-cert-file", str(cert_path), "-key-file", str(key_path),
             "localhost", "127.0.0.1", "::1"],
            check=True, capture_output=True, timeout=60,
        )
    except (subprocess.SubprocessError, OSError) as exc:
        logger.warning("[LOGIN] mkcert failed, falling back to self-signed: %s", exc)
        return False
    return cert_path.is_file() and key_path.is_file()


def _write_self_signed(cert_path: Path, key_path: Path) -> None:
    from cryptography import x509
    from cryptography.hazmat.primitives import hashes, serialization
    from cryptography.hazmat.primitives.asymmetric import ec
    from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

    key = ec.generate_private_key(ec.SECP256R1())
    name = x509.Name([x509.NameAttribute(NameOID.COMMON_NAME, "localhost")])
    now = _dt.datetime.now(_dt.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(name)
        .issuer_name(name)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - _dt.timedelta(minutes=5))
        .not_valid_after(now + _dt.timedelta(days=_SELF_SIGNED_DAYS))
        .add_extension(x509.SubjectAlternativeName([
            x509.DNSName("localhost"),
            x509.IPAddress(ipaddress.ip_address("127.0.0.1")),
            x509.IPAddress(ipaddress.ip_address("::1")),
        ]), critical=False)
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]), critical=False)
        .sign(key, hashes.SHA256())
    )
    key_bytes = key.private_bytes(
        serialization.Encoding.PEM,
        serialization.PrivateFormat.PKCS8,
        serialization.NoEncryption(),
    )
    fd = os.open(key_path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "wb") as fh:
        fh.write(key_bytes)
    cert_path.write_bytes(cert.public_bytes(serialization.Encoding.PEM))


def ensure_localhost_cert(tls_dir: str | os.PathLike = DEFAULT_TLS_DIR) -> tuple[Path, Path, str]:
    """Return ``(cert_path, key_path, source)`` for a localhost TLS cert.

    ``source`` is ``"cached"``, ``"mkcert"`` or ``"self-signed"``.
    """
    d = Path(tls_dir).expanduser()
    d.mkdir(parents=True, exist_ok=True)
    try:
        os.chmod(d, 0o700)
    except OSError:
        pass
    cert_path, key_path = d / _CERT_NAME, d / _KEY_NAME
    if cert_path.is_file() and key_path.is_file() and _cert_still_valid(cert_path):
        return cert_path, key_path, "cached"
    if _try_mkcert(cert_path, key_path):
        try:
            os.chmod(key_path, 0o600)
        except OSError:
            pass
        return cert_path, key_path, "mkcert"
    _write_self_signed(cert_path, key_path)
    return cert_path, key_path, "self-signed"


# ── Loopback callback server ─────────────────────────────────────────────

_PAGE = """<!doctype html><html><head><meta charset="utf-8">
<title>office-connect sign-in</title>
<style>body{{font-family:system-ui,sans-serif;max-width:32rem;margin:4rem auto;padding:0 1rem}}</style>
</head><body><h2>{title}</h2><p>{body}</p></body></html>"""


def _is_loopback(addr: str) -> bool:
    try:
        ip = ipaddress.ip_address(addr.split("%", 1)[0])
    except ValueError:
        return False
    if ip.version == 6 and ip.ipv4_mapped is not None:
        ip = ip.ipv4_mapped
    return ip.is_loopback


class _CallbackState:
    def __init__(self, path: str):
        self.path = path
        self.params: dict[str, str] | None = None
        self.event = threading.Event()


def _make_handler(state: _CallbackState) -> type[BaseHTTPRequestHandler]:
    class Handler(BaseHTTPRequestHandler):
        def log_message(self, fmt: str, *args: Any) -> None:  # silence stderr noise
            logger.debug("[LOGIN] callback %s", fmt % args)

        def _send(self, status: int, title: str, body: str) -> None:
            payload = _PAGE.format(title=html.escape(title), body=html.escape(body)).encode()
            self.send_response(status)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(payload)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("Referrer-Policy", "no-referrer")
            self.end_headers()
            self.wfile.write(payload)

        def do_GET(self) -> None:  # noqa: N802 (stdlib API)
            if not _is_loopback(self.client_address[0]):
                self._send(403, "Forbidden", "Loopback only.")
                return
            parts = urlsplit(self.path)
            if parts.path != state.path:
                self._send(404, "Not found", "")
                return
            if state.event.is_set():
                self._send(409, "Already handled",
                           "The sign-in response was already received. You can close this tab.")
                return
            params = {k: v[0] for k, v in parse_qs(parts.query).items() if v}
            state.params = params
            state.event.set()
            if "error" in params:
                self._send(400, "Sign-in failed",
                           params.get("error_description") or params["error"])
            else:
                self._send(200, "Signed in",
                           "office-connect received the sign-in. You can close this tab "
                           "and return to the terminal.")

    return Handler


class _LoopbackHTTPSServer(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def _start_servers(port: int, ssl_ctx: ssl.SSLContext, state: _CallbackState) -> list[_LoopbackHTTPSServer]:
    """Bind on 127.0.0.1 and (best effort) ::1 — never on a public interface."""
    handler = _make_handler(state)
    servers: list[_LoopbackHTTPSServer] = []
    errors: list[str] = []
    for family, addr in ((socket.AF_INET, "127.0.0.1"), (socket.AF_INET6, "::1")):
        cls = type("_S", (_LoopbackHTTPSServer,), {"address_family": family})
        try:
            srv = cls((addr, port), handler)
        except OSError as exc:
            errors.append(f"{addr}:{port}: {exc}")
            continue
        srv.socket = ssl_ctx.wrap_socket(srv.socket, server_side=True)
        threading.Thread(target=srv.serve_forever, name=f"login-cb-{addr}", daemon=True).start()
        servers.append(srv)
    if not servers:
        raise BrowserLoginError(
            "could not start the local HTTPS callback listener: " + "; ".join(errors)
            + " — is another process using the port? Try --redirect-port."
        )
    return servers


def run_browser_login(
    msal_app: Any,
    scopes: list[str],
    *,
    host: str = DEFAULT_REDIRECT_HOST,
    port: int = DEFAULT_REDIRECT_PORT,
    path: str = DEFAULT_REDIRECT_PATH,
    cert_file: str | None = None,
    key_file: str | None = None,
    tls_dir: str = DEFAULT_TLS_DIR,
    open_browser: bool = True,
    timeout: float = DEFAULT_TIMEOUT_S,
    login_hint: str | None = None,
    echo: Callable[[str], None] = print,
) -> dict:
    """Run the auth-code + PKCE flow and return MSAL's token result dict."""
    if not path.startswith("/"):
        path = "/" + path
    redirect_uri = build_redirect_uri(host, port, path)

    if cert_file and key_file:
        cert_path, key_path, source = Path(cert_file).expanduser(), Path(key_file).expanduser(), "provided"
    else:
        cert_path, key_path, source = ensure_localhost_cert(tls_dir)
    ssl_ctx = ssl.create_default_context(ssl.Purpose.CLIENT_AUTH)
    ssl_ctx.load_cert_chain(str(cert_path), str(key_path))

    flow = msal_app.initiate_auth_code_flow(
        scopes, redirect_uri=redirect_uri, login_hint=login_hint or None,
    )
    if "auth_uri" not in flow:
        err = flow.get("error_description") or flow.get("error") or flow
        raise BrowserLoginError(f"could not build the authorize URL: {err}")

    state = _CallbackState(path)
    servers = _start_servers(port, ssl_ctx, state)
    try:
        echo(f"Listening for the sign-in redirect on {redirect_uri}")
        if source == "self-signed":
            echo("Note: using a self-signed localhost certificate — your browser will "
                 "show a certificate warning after sign-in; accept it to continue "
                 "(install `mkcert` to avoid this).")
        echo(f"\nOpen this URL to sign in (if the browser does not open by itself):\n\n  {flow['auth_uri']}\n")
        if open_browser:
            try:
                webbrowser.open(flow["auth_uri"], new=1)
            except Exception as exc:  # pragma: no cover - platform dependent
                logger.debug("[LOGIN] webbrowser.open failed: %s", exc)
        echo("Waiting for sign-in… (Ctrl-C to cancel)")
        if not state.event.wait(timeout):
            raise BrowserLoginError(f"timed out after {int(timeout)}s waiting for the browser redirect")
    finally:
        for srv in servers:
            srv.shutdown()
            srv.server_close()

    params = state.params or {}
    if "error" in params:
        raise BrowserLoginError(params.get("error_description") or params["error"])
    result = msal_app.acquire_token_by_auth_code_flow(flow, params)
    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or result
        raise BrowserLoginError(f"token redemption failed: {err}")
    return result

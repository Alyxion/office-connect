"""Tests for the browser (auth-code + PKCE) login with a loopback HTTPS callback."""

from __future__ import annotations

import socket
import ssl
import threading
import time
import urllib.error
import urllib.request

import pytest

from office_con.auth import browser_login as bl


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def test_build_redirect_uri():
    assert bl.build_redirect_uri() == "https://localhost:8443/auth/callback"
    assert bl.build_redirect_uri("localhost", 443, "cb") == "https://localhost/cb"


def test_self_signed_cert_generated_and_cached(tmp_path, monkeypatch):
    monkeypatch.setattr(bl.shutil, "which", lambda _name: None)  # force self-signed
    cert, key, source = bl.ensure_localhost_cert(tmp_path)
    assert source == "self-signed"
    assert (key.stat().st_mode & 0o777) == 0o600
    ctx = ssl.create_default_context(ssl.Purpose.CLIENT_AUTH)
    ctx.load_cert_chain(str(cert), str(key))  # valid pair
    _, _, source2 = bl.ensure_localhost_cert(tmp_path)
    assert source2 == "cached"


def test_is_loopback():
    assert bl._is_loopback("127.0.0.1")
    assert bl._is_loopback("::1")
    assert bl._is_loopback("::ffff:127.0.0.1")
    assert not bl._is_loopback("192.168.1.10")
    assert not bl._is_loopback("garbage")


class _FakeMsal:
    def __init__(self):
        self.redeemed_with = None
        self.redirect_uri = None

    def initiate_auth_code_flow(self, scopes, redirect_uri=None, login_hint=None):
        self.redirect_uri = redirect_uri
        return {"auth_uri": "https://login.example/authorize?x=1", "state": "S"}

    def acquire_token_by_auth_code_flow(self, flow, params):
        self.redeemed_with = params
        if params.get("state") != flow["state"]:
            return {"error": "state_mismatch"}
        return {"access_token": "AT", "refresh_token": "RT", "id_token_claims": {}}


def _hit(url: str, timeout: float = 5.0):
    ctx = ssl._create_unverified_context()
    deadline = time.time() + timeout
    while True:
        try:
            return urllib.request.urlopen(url, context=ctx, timeout=2)
        except urllib.error.HTTPError as e:
            return e
        except urllib.error.URLError:
            if time.time() > deadline:
                raise
            time.sleep(0.05)


def _run_in_thread(fake, port, tmp_path, monkeypatch, **kw):
    monkeypatch.setattr(bl.shutil, "which", lambda _name: None)
    out: dict = {}

    def target():
        try:
            out["result"] = bl.run_browser_login(
                fake, ["User.Read"], port=port, tls_dir=str(tmp_path),
                open_browser=False, echo=lambda _m: None, **kw,
            )
        except Exception as exc:  # noqa: BLE001
            out["error"] = exc

    t = threading.Thread(target=target, daemon=True)
    t.start()
    return t, out


def test_callback_roundtrip(tmp_path, monkeypatch):
    port = _free_port()
    fake = _FakeMsal()
    t, out = _run_in_thread(fake, port, tmp_path, monkeypatch, timeout=10)

    assert _hit(f"https://127.0.0.1:{port}/other").code == 404
    resp = _hit(f"https://127.0.0.1:{port}/auth/callback?code=C&state=S")
    assert resp.status == 200
    t.join(5)
    assert "error" not in out, out.get("error")
    assert out["result"]["access_token"] == "AT"
    assert fake.redeemed_with == {"code": "C", "state": "S"}
    assert fake.redirect_uri == f"https://localhost:{port}/auth/callback"


def test_callback_error_param(tmp_path, monkeypatch):
    port = _free_port()
    t, out = _run_in_thread(_FakeMsal(), port, tmp_path, monkeypatch, timeout=10)
    resp = _hit(f"https://127.0.0.1:{port}/auth/callback?error=access_denied&error_description=nope")
    assert resp.code == 400
    t.join(5)
    assert isinstance(out.get("error"), bl.BrowserLoginError)
    assert "nope" in str(out["error"])


def test_timeout(tmp_path, monkeypatch):
    port = _free_port()
    t, out = _run_in_thread(_FakeMsal(), port, tmp_path, monkeypatch, timeout=0.3)
    t.join(5)
    assert isinstance(out.get("error"), bl.BrowserLoginError)
    assert "timed out" in str(out["error"])


def test_port_in_use(tmp_path, monkeypatch):
    monkeypatch.setattr(bl.shutil, "which", lambda _name: None)
    s4 = socket.socket()
    s4.bind(("127.0.0.1", 0))
    s4.listen()
    port = s4.getsockname()[1]
    s6 = socket.socket(socket.AF_INET6)
    try:
        s6.bind(("::1", port))
        s6.listen()
    except OSError:
        pass
    try:
        with pytest.raises(bl.BrowserLoginError, match="callback listener"):
            bl.run_browser_login(_FakeMsal(), ["x"], port=port, tls_dir=str(tmp_path),
                                 open_browser=False, echo=lambda _m: None, timeout=1)
    finally:
        s4.close()
        s6.close()


def test_cli_browser_flag_writes_keyfile(tmp_path, monkeypatch):
    from office_con import mcp_server

    captured = {}

    def fake_run(msal_app, scopes, **kw):
        captured.update(kw)
        return {"access_token": "AT", "refresh_token": "RT",
                "id_token_claims": {"preferred_username": "user@example.com"}, "scope": "User.Read"}

    monkeypatch.setattr(mcp_server, "run_browser_login", fake_run)
    import msal
    monkeypatch.setattr(msal, "PublicClientApplication", lambda *a, **k: object())

    async def fake_verify(_path):
        return {"display_name": "U", "mail": "user@example.com", "job_title": None,
                "inbox_total": 0, "inbox_unread": 0}

    monkeypatch.setattr(mcp_server, "_verify_login_async", fake_verify)
    keyfile = tmp_path / "token.json"
    mcp_server._cli_login([
        "--keyfile", str(keyfile), "--app-config", str(tmp_path / "config.json"),
        "--client-id", "cid", "--tenant-id", "tid", "--browser", "--no-open",
    ])
    import json
    data = json.loads(keyfile.read_text())
    assert data["access_token"] == "AT" and data["refresh_token"] == "RT"
    assert data["email"] == "user@example.com"
    assert captured["port"] == 8443 and captured["open_browser"] is False

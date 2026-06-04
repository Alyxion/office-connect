"""Tests for rotated/redacting logging and HTTP 429 throttle handling."""

from __future__ import annotations

import logging

import pytest

from office_con.logging_setup import (
    configure_logging,
    is_server_environment,
    _RedactingFilter,
)
from office_con.msgraph.ms_graph_handler import (
    _redact_url,
    _parse_retry_after,
    _THROTTLE_WAIT_CAP_S,
    MsGraphInstance,
)


# ── URL redaction ─────────────────────────────────────────────────────────

def test_redact_url_strips_query_with_search_term():
    url = 'https://graph.microsoft.com/v1.0/me/messages?$search="Jochen Munz"&$select=id'
    assert _redact_url(url) == "https://graph.microsoft.com/v1.0/me/messages"


def test_redact_url_strips_fragment_and_handles_plain():
    assert _redact_url("https://x/y#frag") == "https://x/y"
    assert _redact_url("https://x/y") == "https://x/y"


# ── Retry-After parsing ─────────────────────────────────────────────────────

def test_retry_after_honored_and_clamped():
    assert _parse_retry_after({"Retry-After": "3"}, 0) == 3.0
    # A hostile/huge Retry-After is clamped so we never re-introduce a long hang.
    assert _parse_retry_after({"Retry-After": "9999"}, 0) == _THROTTLE_WAIT_CAP_S


def test_retry_after_fallback_backoff_when_absent_or_garbage():
    assert _parse_retry_after({}, 0) == 0.5
    assert _parse_retry_after({}, 2) == 2.0
    assert _parse_retry_after({"Retry-After": "soon"}, 1) == 1.0  # garbage → backoff


# ── Redacting log filter ────────────────────────────────────────────────────

@pytest.mark.parametrize("raw, must_not_contain, must_contain", [
    ("Authorization Bearer eyJabc.def.ghi xyz", "eyJabc", "Bearer <redacted>"),
    ("token eyJ0eXAiOiJKV1Qabc.payloadpart.sig", "payloadpart", "<redacted-jwt>"),
    ("mail to alice@example.com done", "example.com", "<redacted-email>"),
])
def test_redacting_filter_masks_secrets(raw, must_not_contain, must_contain):
    rec = logging.LogRecord("n", logging.INFO, __file__, 1, raw, (), None)
    assert _RedactingFilter().filter(rec) is True
    out = rec.getMessage()
    assert must_not_contain not in out
    assert must_contain in out


def test_configure_logging_writes_redacted_rotating_file(tmp_path):
    log_path = tmp_path / "logs" / "oc.log"
    result = configure_logging(log_file=str(log_path), level="INFO")
    try:
        assert result == log_path
        logging.getLogger("office_con.x").warning(
            "Bearer eyJsecrettoken.aa.bb hit user@example.com"
        )
        for h in logging.getLogger("office_con").handlers:
            h.flush()
        text = log_path.read_text()
        assert "eyJsecrettoken" not in text
        assert "user@example.com" not in text
        assert "<redacted-email>" in text
        # 0600 perms.
        assert (log_path.stat().st_mode & 0o777) == 0o600
    finally:
        # Detach handler so the temp file isn't held across tests.
        configure_logging(log_file="none")


def test_configure_logging_none_disables_file(tmp_path, monkeypatch):
    monkeypatch.delenv("OFFICE_CONNECT_LOG_FILE", raising=False)
    assert configure_logging(log_file="none") is None


# ── server-environment privacy gate ─────────────────────────────────────────

@pytest.mark.parametrize("env_key, env_val", [
    ("OFFICE_CONNECT_SERVER", "1"),
    ("WEBSITE_INSTANCE_ID", "abc123"),
    ("FUNCTIONS_WORKER_RUNTIME", "python"),
    ("WEBSITE_URL", "https://myapp.azurewebsites.net"),
    ("KUBERNETES_SERVICE_HOST", "10.0.0.1"),
])
def test_server_environment_detected(env_key, env_val, monkeypatch):
    for k in ("OFFICE_CONNECT_SERVER", "WEBSITE_INSTANCE_ID",
              "FUNCTIONS_WORKER_RUNTIME", "WEBSITE_URL", "KUBERNETES_SERVICE_HOST"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv(env_key, env_val)
    assert is_server_environment() is True


def test_local_environment_not_flagged(monkeypatch):
    for k in ("OFFICE_CONNECT_SERVER", "WEBSITE_INSTANCE_ID",
              "FUNCTIONS_WORKER_RUNTIME", "WEBSITE_URL", "KUBERNETES_SERVICE_HOST"):
        monkeypatch.delenv(k, raising=False)
    monkeypatch.setenv("WEBSITE_URL", "http://localhost:8080")  # local dev
    assert is_server_environment() is False


def test_server_env_disables_file_logging_even_with_explicit_path(tmp_path, monkeypatch):
    """A server deployment must NOT write a log file even if a path is forced."""
    monkeypatch.setenv("OFFICE_CONNECT_SERVER", "1")
    log_path = tmp_path / "logs" / "oc.log"
    result = configure_logging(log_file=str(log_path), level="INFO")
    assert result is None
    assert not log_path.exists()
    # And library logs still propagate to the host app's logging.
    assert logging.getLogger("office_con").propagate is True


# ── 429 retry loop in run_async ─────────────────────────────────────────────

class _Resp:
    def __init__(self, status_code, headers=None):
        self.status_code = status_code
        self.headers = headers or {}


@pytest.mark.asyncio
async def test_run_async_retries_on_429_then_succeeds(monkeypatch):
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.cache_dict["access_token"] = "tok"

    # First two transport calls are throttled, third succeeds.
    seq = [_Resp(429, {"Retry-After": "0"}), _Resp(429, {"Retry-After": "0"}), _Resp(200)]
    calls = {"n": 0}

    async def fake_super(self, *, url, method="GET", json=None, token=None, add_headers=None):
        i = calls["n"]; calls["n"] += 1
        return seq[i]

    # Patch the base-class run_async that MsGraphInstance.run_async delegates to.
    monkeypatch.setattr(
        "office_con.web_user_instance.WebUserInstance.run_async", fake_super
    )

    resp = await g.run_async(url="https://graph.microsoft.com/v1.0/search/query",
                             method="POST", json={}, token="tok")
    assert resp.status_code == 200
    assert calls["n"] == 3  # original + 2 retries


@pytest.mark.asyncio
async def test_run_async_gives_up_after_max_429(monkeypatch):
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.cache_dict["access_token"] = "tok"
    calls = {"n": 0}

    async def always_429(self, *, url, method="GET", json=None, token=None, add_headers=None):
        calls["n"] += 1
        return _Resp(429, {"Retry-After": "0"})

    monkeypatch.setattr(
        "office_con.web_user_instance.WebUserInstance.run_async", always_429
    )

    resp = await g.run_async(url="https://graph.microsoft.com/v1.0/search/query", token="tok")
    # Returns the throttled response (handler surfaces a clear message); does
    # not raise or loop forever. original + _MAX_THROTTLE_RETRIES(3) = 4 calls.
    assert resp.status_code == 429
    assert calls["n"] == 4

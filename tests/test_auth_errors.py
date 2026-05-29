"""Unit tests for self-healing auth failure surfacing.

When the Microsoft 365 session is dead (expired/revoked token that cannot be
refreshed), the Graph layer must raise ``AuthExpiredError`` instead of letting
handlers silently return empty results, and the MCP layer must turn that into a
clear, actionable "run office-connect login" message. No credentials needed.
"""

from __future__ import annotations

import pytest

from office_con.mcp_server import _auth_error_text, _handle_tool
from office_con.msgraph.ms_graph_handler import AuthExpiredError, MsGraphInstance


class _Resp:
    def __init__(self, status_code: int):
        self.status_code = status_code

    def json(self):
        return {}


def _dead_graph(monkeypatch) -> MsGraphInstance:
    """A graph whose Graph calls always 401 and whose refresh always fails."""
    g = MsGraphInstance(can_refresh=True, client_id="x", tenant_id="common")
    g.cache_dict["access_token"] = "dead"
    g.cache_dict["refresh_token"] = "also-dead"
    g.msg_endpoint = "https://graph.microsoft.com/v1.0/"

    async def fake_parent(self, **kwargs):
        return _Resp(401)

    # Patch the inherited transport so every request comes back 401.
    monkeypatch.setattr(MsGraphInstance.__bases__[0], "run_async", fake_parent)

    async def fail_refresh():
        return None

    async def cached_token():
        return "dead"

    g.refresh_token_async = fail_refresh  # type: ignore[assignment]
    g.get_access_token_async = cached_token  # type: ignore[assignment]
    return g


@pytest.mark.asyncio
async def test_run_async_raises_on_unrecoverable_401(monkeypatch):
    g = _dead_graph(monkeypatch)
    with pytest.raises(AuthExpiredError):
        await g.run_async(url=g.msg_endpoint + "me", token="dead")


@pytest.mark.asyncio
async def test_run_async_raises_when_no_refresh_capability(monkeypatch):
    g = _dead_graph(monkeypatch)
    g.can_refresh = False
    with pytest.raises(AuthExpiredError):
        await g.run_async(url=g.msg_endpoint + "me", token="dead")


@pytest.mark.asyncio
async def test_check_connection_surfaces_auth_failure(monkeypatch):
    """o365_check_connection on a dead session bubbles AuthExpiredError, which
    the call_tool wrapper renders as the re-auth message."""
    g = _dead_graph(monkeypatch)
    with pytest.raises(AuthExpiredError):
        await _handle_tool(g, "o365_check_connection", {})


@pytest.mark.asyncio
async def test_profile_no_longer_swallows_dead_auth(monkeypatch):
    """A 401 during profile fetch must raise, not return an empty profile."""
    g = _dead_graph(monkeypatch)
    with pytest.raises(AuthExpiredError):
        await _handle_tool(g, "o365_get_profile", {})


def test_auth_error_text_names_the_fix():
    text = _auth_error_text("token revoked")
    assert "office-connect login" in text
    assert "token revoked" in text

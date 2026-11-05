"""End-to-end tests that drive the real office-connect stdio MCP server through
our own client (``office_con.testing.mcp_client``) against mock data.

Unlike the handler-level mock tests, these exercise the FULL stack: a spawned
``python -m office_con.mcp_server`` subprocess, the JSON-RPC stdio transport,
tool registration / permission filtering, and ``call_tool`` dispatch. This is
the layer where the bugs from the recent mail work lived (get_mail expanding
attachment bytes inline; the contentId $select 400; date-range search), so these
tests guard exactly those tool-layer regressions.

Each test spawns its own short-lived server (~1-2s). They need no credentials
and make no network calls — everything is served by MockGraphTransport via the
keyfile's ``"mock": true`` flag plus ``LLMING_MOCK_USERS=1`` (set by
``mock_server_client``).
"""

import asyncio
import base64
import json
import urllib.error
import urllib.request

import pytest

from office_con.testing.mcp_client import MCPClient, mock_server_client


def _json(client: MCPClient, result: dict):
    return json.loads(client.text_content(result))


async def _first_attachment_mail(client: MCPClient) -> dict:
    """Find an attachment-bearing mail by scanning list metadata (mock ids are
    random per process, so they must be discovered through the server)."""
    listing = _json(client, await client.call_tool("o365_list_mail", {"limit": 40}))
    for m in listing["elements"]:
        if m.get("has_attachments"):
            return m
    raise AssertionError("no attachment-bearing mail in mock inbox")


@pytest.mark.asyncio
async def test_initialize_and_list_tools():
    async with mock_server_client() as client:
        assert client.server_info.get("name")  # handshake completed
        tools = await client.list_tools()
        names = {t["name"] for t in tools}
        # A representative spread across tiers must be exposed at 'all'.
        assert {"o365_get_mail", "o365_search_mail", "o365_send_mail",
                "o365_peek_mail_attachment"} <= names


@pytest.mark.asyncio
async def test_permission_tier_filters_tools():
    """read_only must hide write tools that 'all' exposes."""
    async with mock_server_client(permission_level="all") as client:
        all_names = {t["name"] for t in await client.list_tools()}
    async with mock_server_client(permission_level="read_only") as client:
        ro_names = {t["name"] for t in await client.list_tools()}
    assert "o365_send_mail" in all_names
    assert "o365_send_mail" not in ro_names          # write tool hidden
    assert "o365_get_mail" in ro_names               # read tool still present
    assert ro_names < all_names


@pytest.mark.asyncio
async def test_check_connection_reports_authenticated():
    async with mock_server_client() as client:
        result = _json(client, await client.call_tool("o365_check_connection"))
        assert result["connected"] is True
        assert "@" in result["email"]


@pytest.mark.asyncio
async def test_get_mail_lists_attachment_metadata_not_bytes():
    """The headline fix: get_mail must list attachments as metadata
    (attachment_id + size) WITHOUT downloading the bytes — verified through the
    real MCP layer, not just the handler."""
    async with mock_server_client() as client:
        mail_row = await _first_attachment_mail(client)
        mail = _json(client, await client.call_tool(
            "o365_get_mail", {"email_id": mail_row["email_id"]}))
        assert mail["has_attachments"] is True
        assert mail["attachments"], "attachment metadata should be populated"
        for att in mail["attachments"]:
            assert att["attachment_id"], "id needed for lazy peek/download"
            assert att["name"]
            assert att["size"] is not None
            assert att["content_bytes"] is None, "bytes must NOT be inlined"


@pytest.mark.asyncio
async def test_peek_attachment_fetches_by_id_and_degrades_gracefully():
    """Lazy attachment read: peek fetches bytes by id. The mock's fake xlsx is
    not a real spreadsheet, so this also asserts peek degrades to a clean
    'unreadable' result instead of a generic tool failure."""
    async with mock_server_client() as client:
        mail_row = await _first_attachment_mail(client)
        mail = _json(client, await client.call_tool(
            "o365_get_mail", {"email_id": mail_row["email_id"]}))
        att_id = mail["attachments"][0]["attachment_id"]
        peek = _json(client, await client.call_tool(
            "o365_peek_mail_attachment",
            {"email_id": mail_row["email_id"], "attachment_id": att_id}))
        # Reached the bytes (byte_size present) rather than "no inline content".
        assert peek.get("byte_size", 0) > 0
        assert peek.get("kind") in {"pdf", "xlsx", "docx", "unreadable", "unsupported"}


@pytest.mark.asyncio
async def test_search_mail_with_date_range_uses_search_query():
    """o365_search_mail with since/until routes through /search/query and still
    returns the OfficeMailList shape end-to-end."""
    async with mock_server_client() as client:
        result = _json(client, await client.call_tool("o365_search_mail", {
            "body": "the", "since": "2020-01-01", "until": "2030-12-31", "limit": 5,
        }))
        assert "elements" in result
        for m in result["elements"]:
            assert m["email_id"]
            assert "subject" in m


@pytest.mark.asyncio
async def test_search_mail_requires_criteria():
    """An empty search is refused (not silently run as a full-mailbox dump)."""
    async with mock_server_client() as client:
        text = await client.call_tool_text("o365_search_mail", {})
        assert "Refused" in text or "criteria" in text.lower()


@pytest.mark.asyncio
async def test_get_mail_attachment_base64_round_trips_to_draft():
    """Without a file server, get_mail_attachment returns base64 that decodes to
    real bytes and can be re-attached to a new draft."""
    async with mock_server_client() as client:
        mail_row = await _first_attachment_mail(client)
        mail = _json(client, await client.call_tool(
            "o365_get_mail", {"email_id": mail_row["email_id"]}))
        att = mail["attachments"][0]
        dl = _json(client, await client.call_tool("o365_get_mail_attachment", {
            "email_id": mail_row["email_id"], "attachment_id": att["attachment_id"]}))
        assert dl["delivery"] == "base64"
        raw = base64.b64decode(dl["content_base64"])
        assert len(raw) > 0
        # re-attach to a draft by content_base64
        draft = _json(client, await client.call_tool("o365_create_mail_draft", {
            "to": ["receipts@example.com"], "subject": "Receipt", "body": "see attached",
            "attachments": [{"name": dl["name"], "content_type": dl["content_type"],
                             "content_base64": dl["content_base64"]}]}))
        assert draft  # draft created


@pytest.mark.asyncio
async def test_get_mail_attachment_url_requires_file_server():
    """delivery='url' is refused with guidance when the file server is off."""
    async with mock_server_client() as client:
        mail_row = await _first_attachment_mail(client)
        mail = _json(client, await client.call_tool(
            "o365_get_mail", {"email_id": mail_row["email_id"]}))
        att_id = mail["attachments"][0]["attachment_id"]
        text = await client.call_tool_text("o365_get_mail_attachment", {
            "email_id": mail_row["email_id"], "attachment_id": att_id, "delivery": "url"})
        assert "file server" in text.lower()


@pytest.mark.asyncio
async def test_get_mail_attachment_url_delivery_fetchable():
    """With --file-server, delivery defaults to an unguessable loopback URL whose
    bytes (fetched over real HTTP) match the base64 path."""
    async with mock_server_client(extra_args=["--file-server"]) as client:
        mail_row = await _first_attachment_mail(client)
        mail = _json(client, await client.call_tool(
            "o365_get_mail", {"email_id": mail_row["email_id"]}))
        att_id = mail["attachments"][0]["attachment_id"]

        res = _json(client, await client.call_tool("o365_get_mail_attachment", {
            "email_id": mail_row["email_id"], "attachment_id": att_id}))
        assert res["delivery"] == "url"
        assert res["url"].startswith("http://127.0.0.1:")
        assert "/f/" in res["url"] and len(res["url"].split("/f/")[1].split("/")[0]) >= 32

        def _fetch(url):
            r = urllib.request.urlopen(url)
            return r.read()
        fetched = await asyncio.to_thread(_fetch, res["url"])

        # Even with the file server on, an explicit delivery='base64' is honored
        # (a network-isolated fetcher cannot reach the loopback URL) — the bytes
        # match what the URL serves.
        b64 = _json(client, await client.call_tool("o365_get_mail_attachment", {
            "email_id": mail_row["email_id"], "attachment_id": att_id, "delivery": "base64"}))
        assert b64["delivery"] == "base64"
        assert base64.b64decode(b64["content_base64"]) == fetched

        # A tampered/guessed token must 404.
        bad = res["url"].rsplit("/f/", 1)[0] + "/f/" + "A" * 43 + "/x"
        with pytest.raises(urllib.error.HTTPError) as exc:
            await asyncio.to_thread(_fetch, bad)
        assert exc.value.code == 404

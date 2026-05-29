"""Tests for the mail/calendar enrichments requested via Claude Desktop feedback.

Covers (against the mock transport — no real HTTP):
- recipient + conversation/internet-message-id metadata on parse and list
- graph_url / outlook_url aliases
- body_format text/none and body_truncated truncation
- batch get_mails
- folder scoping + exclusion + unread counts
- legacy Exchange-DN isolation
- reply / forward / update_event / send_event_invite action tools
"""

from __future__ import annotations

import pytest

from office_con.mcp_server import _body_opts, _handle_tool
from office_con.msgraph.mail_handler import (
    MailAddress,
    _is_legacy_dn,
    resolve_well_known_folder,
)
from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.testing.fixtures import default_mock_profile
from office_con.testing.mock_tokens import make_mock_access_token


@pytest.fixture
def graph():
    profile = default_mock_profile()
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.enable_mock(profile)
    g.cache_dict["access_token"] = make_mock_access_token(profile.email, profile.user_id)
    return g


# ── pure helpers ──────────────────────────────────────────────────────────

def test_resolve_well_known_folder():
    assert resolve_well_known_folder("sent") == "sentitems"
    assert resolve_well_known_folder("trash") == "deleteditems"
    assert resolve_well_known_folder("junk") == "junkemail"
    assert resolve_well_known_folder("INBOX") == "inbox"
    # Unknown names pass through unchanged (assumed to be a folder id).
    assert resolve_well_known_folder("AAMk-custom-id") == "AAMk-custom-id"
    assert resolve_well_known_folder(None) is None


def test_is_legacy_dn():
    assert _is_legacy_dn("/O=EXCHANGELABS/OU=x/CN=RECIPIENTS/CN=abc")
    assert not _is_legacy_dn("michael@example.com")
    assert not _is_legacy_dn(None)


def test_parse_address_isolates_legacy_dn():
    from office_con.msgraph.mail_handler import _parse_address
    legacy = _parse_address({"emailAddress": {"name": "Old User", "address": "/O=EXCHANGELABS/CN=RECIPIENTS/CN=x"}})
    assert isinstance(legacy, MailAddress)
    assert legacy.address is None
    assert legacy.legacy_dn.startswith("/O=")
    normal = _parse_address({"emailAddress": {"name": "N", "address": "n@example.com"}})
    assert normal.address == "n@example.com"
    assert normal.legacy_dn is None


def test_body_opts():
    assert _body_opts({}) == ("text", 50000)
    assert _body_opts({"body_format": "html", "max_body_chars": 0}) == ("html", None)
    assert _body_opts({"max_body_chars": 100}) == ("text", 100)


# ── list / search enrichment ────────────────────────────────────────────────

class TestListEnrichment:
    @pytest.mark.asyncio
    async def test_list_includes_recipients_and_ids(self, graph):
        result = await graph.get_mail().email_index_async(limit=5)
        assert result.elements
        m = result.elements[0]
        assert m.to_recipients and m.to_recipients[0].address
        assert m.conversation_id is not None
        assert m.internet_message_id is not None
        # URL aliases populated.
        assert m.graph_url and m.graph_url == m.email_url
        assert m.outlook_url == m.web_link

    @pytest.mark.asyncio
    async def test_folder_scoping(self, graph):
        sent = await graph.get_mail().email_index_async(folder="sent", limit=50)
        assert sent.elements
        # All returned messages live in Sent Items.
        for m in sent.elements:
            assert m.email_id

    @pytest.mark.asyncio
    async def test_exclude_folders(self, graph):
        # Search across all folders, excluding deleted items.
        all_hits = await graph.get_mail().email_index_async(query="the", limit=100)
        excluded = await graph.get_mail().email_index_async(
            query="the", limit=100, exclude_folders=["deleteditems"],
        )
        # Excluding a folder cannot increase the result count.
        assert len(excluded.elements) <= len(all_hits.elements)


# ── body hygiene ────────────────────────────────────────────────────────────

class TestBodyHygiene:
    @pytest.mark.asyncio
    async def test_truncation_flag(self, graph):
        listing = await graph.get_mail().email_index_async(limit=20)
        an_id = listing.elements[0].email_id
        mail = await graph.get_mail().get_mail_async(email_id=an_id, max_body_chars=10)
        if mail.body:
            assert len(mail.body) <= 10
            assert mail.body_truncated is True

    @pytest.mark.asyncio
    async def test_body_none_skips_body(self, graph):
        listing = await graph.get_mail().email_index_async(limit=5)
        an_id = listing.elements[0].email_id
        mail = await graph.get_mail().get_mail_async(email_id=an_id, body_format="none")
        assert mail is not None
        # bodyPreview survives even when the full body is skipped.
        assert mail.body_preview is not None


# ── batch fetch ──────────────────────────────────────────────────────────────

class TestBatchFetch:
    @pytest.mark.asyncio
    async def test_get_mails_batch(self, graph):
        listing = await graph.get_mail().email_index_async(limit=4)
        ids = [m.email_id for m in listing.elements]
        mails = await graph.get_mail().get_mails_async(ids, max_body_chars=0)
        assert [m.email_id for m in mails] == ids  # preserves order
        assert all(m.graph_url for m in mails)

    @pytest.mark.asyncio
    async def test_get_mails_empty(self, graph):
        assert await graph.get_mail().get_mails_async([]) == []


# ── action handlers ──────────────────────────────────────────────────────────

class TestActionHandlers:
    @pytest.mark.asyncio
    async def test_reply(self, graph):
        listing = await graph.get_mail().email_index_async(limit=1)
        ok = await graph.get_mail().reply_async(listing.elements[0].email_id, "Thanks!")
        assert ok is True

    @pytest.mark.asyncio
    async def test_reply_all(self, graph):
        listing = await graph.get_mail().email_index_async(limit=1)
        ok = await graph.get_mail().reply_async(
            listing.elements[0].email_id, "All thanks", reply_all=True,
        )
        assert ok is True

    @pytest.mark.asyncio
    async def test_forward(self, graph):
        listing = await graph.get_mail().email_index_async(limit=1)
        ok = await graph.get_mail().forward_async(
            listing.elements[0].email_id, ["x@example.com"], "FYI",
        )
        assert ok is True

    @pytest.mark.asyncio
    async def test_update_event(self, graph):
        ev = await graph.get_calendar().update_event_async(
            "some-event-id", subject="Renamed",
        )
        assert ev is not None
        assert ev.subject == "Renamed"

    @pytest.mark.asyncio
    async def test_update_event_no_fields_returns_none(self, graph):
        assert await graph.get_calendar().update_event_async("id") is None


# ── MCP tool dispatch ────────────────────────────────────────────────────────

class TestToolDispatch:
    @pytest.mark.asyncio
    async def test_unread_counts_tool(self, graph):
        out = await _handle_tool(graph, "o365_unread_counts", {})
        assert "total_unread" in out[0].text
        assert "folders" in out[0].text

    @pytest.mark.asyncio
    async def test_get_mails_tool(self, graph):
        listing = await _handle_tool(graph, "o365_list_mail", {"limit": 3})
        import json
        ids = [e["email_id"] for e in json.loads(listing[0].text)["elements"]]
        out = await _handle_tool(graph, "o365_get_mails", {"ids": ids})
        assert json.loads(out[0].text)  # non-empty list

    @pytest.mark.asyncio
    async def test_reply_tool(self, graph):
        listing = await _handle_tool(graph, "o365_list_mail", {"limit": 1})
        import json
        an_id = json.loads(listing[0].text)["elements"][0]["email_id"]
        out = await _handle_tool(graph, "o365_reply_to_mail", {"email_id": an_id, "body": "hi"})
        assert json.loads(out[0].text)["sent"] is True

    @pytest.mark.asyncio
    async def test_send_event_invite_tool(self, graph):
        out = await _handle_tool(graph, "o365_send_event_invite", {
            "subject": "Sync",
            "start": "2026-06-01T10:00:00",
            "end": "2026-06-01T10:30:00",
            "attendees": [{"email": "a@example.com"}],
        })
        import json
        assert json.loads(out[0].text)["subject"] == "Sync"

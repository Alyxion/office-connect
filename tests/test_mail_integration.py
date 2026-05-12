"""Integration tests for mail handlers — run against real MS Graph.

Skipped automatically when no test_config.json or token file is present.
Configure by creating ``tests/test_config.json``::

    {
        "token_file": "~/Downloads/token_export.json"
    }

Run with: poetry run pytest tests/test_mail_integration.py -v -s
"""

from __future__ import annotations

import pytest

from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.msgraph.mail_handler import (
    OfficeMailHandler,
    MailFolderHandler,
    OfficeMail,
    OfficeMailList,
    FolderInfo,
    OfficeMailCategory,
)
from office_con.testing.mock_config import get_test_config, get_token_data


def _make_graph() -> MsGraphInstance:
    data = get_token_data()
    if data is None:
        pytest.skip("No token file configured (tests/test_config.json)")

    graph = MsGraphInstance(
        scopes=None,
        endpoint="https://graph.microsoft.com/v1.0/",
    )
    graph.cache_dict = data
    graph.email = data.get("email", "")
    graph.client_id = data.get("client_id", "")
    graph.client_secret = data.get("client_secret", "")
    graph.tenant_id = data.get("tenant_id", "")
    return graph


_has_token = get_token_data() is not None
_skip = pytest.mark.skipif(not _has_token, reason="No token file — set tests/test_config.json")


# ═══════════════════════════════════════════════════════════════════
# MailFolderHandler — real Graph
# ═══════════════════════════════════════════════════════════════════

@_skip
class TestMailFoldersIntegration:

    @pytest.mark.asyncio
    async def test_get_folders(self):
        graph = _make_graph()
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        assert isinstance(folders, list)
        assert len(folders) > 0
        assert all(isinstance(f, FolderInfo) for f in folders)
        print(f"\n  Found {len(folders)} folders:")
        for f in folders:
            print(f"    {f.name} (id={f.id[:20]}…, total={f.total}, unread={f.unread})")

    @pytest.mark.asyncio
    async def test_inbox_exists(self):
        graph = _make_graph()
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        names_lower = [f.name.lower() for f in folders]
        assert "inbox" in names_lower, f"Inbox not found in: {[f.name for f in folders]}"

    @pytest.mark.asyncio
    async def test_get_single_folder(self):
        graph = _make_graph()
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        assert folders, "No folders to test"
        first = folders[0]
        single = await handler.get_folder_async(first.id)
        assert single is not None
        assert isinstance(single, FolderInfo)
        assert single.id == first.id
        print(f"\n  Single folder: {single.name} (total={single.total}, unread={single.unread})")

    @pytest.mark.asyncio
    async def test_folder_has_parent_id(self):
        graph = _make_graph()
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        for f in folders:
            assert f.parent_id is not None or f.parent_id is None  # field exists
            assert isinstance(f.id, str)
            assert isinstance(f.name, str)


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler.email_index_async — real Graph
# ═══════════════════════════════════════════════════════════════════

@_skip
class TestEmailIndexIntegration:

    @pytest.mark.asyncio
    async def test_list_inbox(self):
        graph = _make_graph()
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=5)
        assert isinstance(result, OfficeMailList)
        assert result.total_mails >= 0
        print(f"\n  Inbox: {result.total_mails} total, fetched {len(result.elements)}")
        for m in result.elements:
            print(f"    [{m.email_id[:12]}…] {m.from_name}: {m.subject}")

    @pytest.mark.asyncio
    async def test_list_with_pagination(self):
        graph = _make_graph()
        mail = graph.get_mail()
        page1 = await mail.email_index_async(limit=2, skip=0)
        page2 = await mail.email_index_async(limit=2, skip=2)
        assert isinstance(page1, OfficeMailList)
        assert isinstance(page2, OfficeMailList)
        if page1.elements and page2.elements:
            ids1 = {m.email_id for m in page1.elements}
            ids2 = {m.email_id for m in page2.elements}
            assert ids1.isdisjoint(ids2), "Pagination returned overlapping messages"

    @pytest.mark.asyncio
    async def test_list_by_folder_id(self):
        graph = _make_graph()
        folders = await graph.get_mail_folders().get_folders_async()
        sent = next((f for f in folders if "sent" in f.name.lower()), None)
        if sent is None:
            pytest.skip("No Sent folder found")
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=3, folder_id=sent.id)
        assert isinstance(result, OfficeMailList)
        print(f"\n  Sent Items: {result.total_mails} total, fetched {len(result.elements)}")

    @pytest.mark.asyncio
    async def test_message_fields(self):
        graph = _make_graph()
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=1)
        if not result.elements:
            pytest.skip("No messages in inbox")
        m = result.elements[0]
        assert isinstance(m, OfficeMail)
        assert m.email_id
        assert m.local_timestamp
        assert isinstance(m.is_read, bool)
        assert isinstance(m.has_attachments, bool)


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler.email_index_async(query=...) — search, real Graph
# ═══════════════════════════════════════════════════════════════════

@_skip
class TestSearchIntegration:

    @pytest.mark.asyncio
    async def test_search(self):
        graph = _make_graph()
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=5, query="meeting")
        assert isinstance(result, OfficeMailList)
        print(f"\n  Search 'meeting': {len(result.elements)} results")
        for m in result.elements:
            print(f"    {m.from_name}: {m.subject}")

    @pytest.mark.asyncio
    async def test_search_returns_mail_objects(self):
        graph = _make_graph()
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=3, query="re:")
        for m in result.elements:
            assert isinstance(m, OfficeMail)
            assert m.email_id


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler.get_mail_async — single message, real Graph
# ═══════════════════════════════════════════════════════════════════

@_skip
class TestGetMailIntegration:

    @pytest.mark.asyncio
    async def test_get_mail_by_id(self):
        graph = _make_graph()
        mail = graph.get_mail()
        listing = await mail.email_index_async(limit=1)
        if not listing.elements:
            pytest.skip("No messages")
        msg_id = listing.elements[0].email_id
        result = await mail.get_mail_async(email_id=msg_id)
        assert result is not None
        assert isinstance(result, OfficeMail)
        assert result.email_id == msg_id
        assert result.body is not None
        print(f"\n  Message: {result.subject}")
        print(f"    Body type: {result.body_type}, length: {len(result.body or '')}")
        print(f"    Attachments: {len(result.attachments)}")

    @pytest.mark.asyncio
    async def test_get_mail_by_url(self):
        graph = _make_graph()
        mail = graph.get_mail()
        listing = await mail.email_index_async(limit=1)
        if not listing.elements:
            pytest.skip("No messages")
        msg = listing.elements[0]
        if not msg.email_url:
            pytest.skip("No email_url on message")
        result = await mail.get_mail_async(email_url=msg.email_url)
        assert result is not None
        assert isinstance(result, OfficeMail)


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler.get_categories_async — real Graph
# ═══════════════════════════════════════════════════════════════════

@_skip
class TestCategoriesIntegration:

    @pytest.mark.asyncio
    async def test_get_categories(self):
        graph = _make_graph()
        mail = graph.get_mail()
        cats = await mail.get_categories_async()
        assert isinstance(cats, list)
        print(f"\n  Found {len(cats)} categories:")
        for c in cats:
            assert isinstance(c, OfficeMailCategory)
            print(f"    {c.name} (color={c.color}, preset={c.preset_color})")


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler.get_user_profile_async — real Graph
# ═══════════════════════════════════════════════════════════════════

@_skip
class TestProfileIntegration:

    @pytest.mark.asyncio
    async def test_get_profile(self):
        graph = _make_graph()
        mail = graph.get_mail()
        profile = await mail.get_user_profile_async()
        assert profile is not None
        assert "displayName" in profile
        print(f"\n  Profile: {profile.get('displayName')} <{profile.get('mail')}>")

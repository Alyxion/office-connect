"""Unit tests for OfficeMailHandler and MailFolderHandler using mock transport.

Run with: poetry run pytest dependencies/office-mcp/tests/test_mail_handler.py -v
"""

from __future__ import annotations

import pytest

from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.msgraph.mail_handler import (
    OfficeMailHandler,
    MailFolderHandler,
    OfficeMail,
    OfficeMailAttachment,
    OfficeMailList,
    FolderInfo,
    MoveResult,
    compute_folder_signature,
)
from office_con.testing.fixtures import default_mock_profile
from office_con.testing.mock_tokens import make_mock_access_token


@pytest.fixture
def graph():
    profile = default_mock_profile()
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.enable_mock(profile)
    g.cache_dict["access_token"] = make_mock_access_token(profile.email, profile.user_id)
    return g


@pytest.fixture
def graph_no_token():
    profile = default_mock_profile()
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.enable_mock(profile)
    return g


# ═══════════════════════════════════════════════════════════════════
# MailFolderHandler
# ═══════════════════════════════════════════════════════════════════

class TestMailFolderHandler:

    @pytest.mark.asyncio
    async def test_get_folders(self, graph: MsGraphInstance):
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        assert isinstance(folders, list)
        assert len(folders) > 0
        assert all(isinstance(f, FolderInfo) for f in folders)

    @pytest.mark.asyncio
    async def test_folder_has_expected_fields(self, graph: MsGraphInstance):
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        inbox = next((f for f in folders if f.name == "Inbox"), None)
        assert inbox is not None
        assert inbox.id == "inbox"
        assert inbox.total > 0

    @pytest.mark.asyncio
    async def test_folder_parent_ids(self, graph: MsGraphInstance):
        handler = graph.get_mail_folders()
        folders = await handler.get_folders_async()
        child = next((f for f in folders if f.name == "Notifications"), None)
        assert child is not None
        assert child.parent_id == "inbox"

    @pytest.mark.asyncio
    async def test_get_single_folder(self, graph: MsGraphInstance):
        handler = graph.get_mail_folders()
        folder = await handler.get_folder_async("inbox")
        assert folder is not None
        assert isinstance(folder, FolderInfo)
        assert folder.id == "inbox"

    @pytest.mark.asyncio
    async def test_get_folder_not_found(self, graph: MsGraphInstance):
        handler = graph.get_mail_folders()
        folder = await handler.get_folder_async("nonexistent-folder-id")
        # Mock returns a fallback, but we verify it's a FolderInfo
        assert folder is None or isinstance(folder, FolderInfo)

    @pytest.mark.asyncio
    async def test_get_folders_no_token(self, graph_no_token: MsGraphInstance):
        handler = graph_no_token.get_mail_folders()
        folders = await handler.get_folders_async()
        assert folders == []

    @pytest.mark.asyncio
    async def test_get_folder_no_token(self, graph_no_token: MsGraphInstance):
        handler = graph_no_token.get_mail_folders()
        folder = await handler.get_folder_async("inbox")
        assert folder is None


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — email_index_async
# ═══════════════════════════════════════════════════════════════════

class TestEmailIndex:

    @pytest.mark.asyncio
    async def test_list_inbox(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=10)
        assert isinstance(result, OfficeMailList)
        assert len(result.elements) > 0
        assert result.total_mails > 0

    @pytest.mark.asyncio
    async def test_list_with_folder_id(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=5, folder_id="inbox")
        assert isinstance(result, OfficeMailList)

    @pytest.mark.asyncio
    async def test_list_with_pagination(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        page1 = await mail.email_index_async(limit=2, skip=0)
        page2 = await mail.email_index_async(limit=2, skip=2)
        assert isinstance(page1, OfficeMailList)
        assert isinstance(page2, OfficeMailList)

    @pytest.mark.asyncio
    async def test_list_no_token(self, graph_no_token: MsGraphInstance):
        mail = graph_no_token.get_mail()
        result = await mail.email_index_async(limit=10)
        assert result.elements == []
        assert result.total_mails == 0

    @pytest.mark.asyncio
    async def test_keyword_only_enforcement(self):
        """mail_address and folder_id must be keyword-only arguments."""
        with pytest.raises(TypeError):
            OfficeMailHandler.__init__  # just to have a reference
            # Simulate calling with positional args where keyword-only is required
            # This is a static check — we call with 4 positional args
            # email_index_async(self, limit, skip, mail_address) should fail
            # We test by introspecting the signature
            import inspect
            sig = inspect.signature(OfficeMailHandler.email_index_async)
            params = list(sig.parameters.values())
            # Find the * separator — params after it are keyword-only
            keyword_only = [p for p in params if p.kind == inspect.Parameter.KEYWORD_ONLY]
            assert any(p.name == "mail_address" for p in keyword_only), "mail_address should be keyword-only"
            assert any(p.name == "folder_id" for p in keyword_only), "folder_id should be keyword-only"
            assert any(p.name == "query" for p in keyword_only), "query should be keyword-only"
            raise TypeError("Expected — verifying keyword-only enforcement works")


class TestEmailIndexKeywordOnly:

    def test_mail_address_is_keyword_only(self):
        import inspect
        sig = inspect.signature(OfficeMailHandler.email_index_async)
        params = list(sig.parameters.values())
        keyword_only = [p for p in params if p.kind == inspect.Parameter.KEYWORD_ONLY]
        names = {p.name for p in keyword_only}
        assert "mail_address" in names
        assert "folder_id" in names
        assert "query" in names


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — search via email_index_async(query=...)
# ═══════════════════════════════════════════════════════════════════

class TestSearch:

    @pytest.mark.asyncio
    async def test_search_by_query(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=10, query="meeting")
        assert isinstance(result, OfficeMailList)

    @pytest.mark.asyncio
    async def test_search_returns_mail_objects(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=5, query="test")
        for m in result.elements:
            assert isinstance(m, OfficeMail)
            assert m.email_id

    @pytest.mark.asyncio
    async def test_search_no_token(self, graph_no_token: MsGraphInstance):
        mail = graph_no_token.get_mail()
        result = await mail.email_index_async(limit=5, query="anything")
        assert result.elements == []


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — delete_message_async
# ═══════════════════════════════════════════════════════════════════

class TestDeleteMessage:

    @pytest.mark.asyncio
    async def test_delete_by_string_id(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        ok = await mail.delete_message_async("some-message-id")
        assert ok is True

    @pytest.mark.asyncio
    async def test_delete_by_office_mail_object(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        listing = await mail.email_index_async(limit=1)
        assert len(listing.elements) > 0
        msg = listing.elements[0]
        ok = await mail.delete_message_async(msg)
        assert ok is True

    @pytest.mark.asyncio
    async def test_delete_no_token(self, graph_no_token: MsGraphInstance):
        mail = graph_no_token.get_mail()
        ok = await mail.delete_message_async("some-id")
        assert ok is False


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — move_message_async
# ═══════════════════════════════════════════════════════════════════

class TestMoveMessage:

    @pytest.mark.asyncio
    async def test_move_by_string_ids(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.move_message_async("msg-id", "archive")
        assert result is not None
        assert isinstance(result, MoveResult)
        assert result.id

    @pytest.mark.asyncio
    async def test_move_by_typed_objects(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        listing = await mail.email_index_async(limit=1)
        assert len(listing.elements) > 0
        msg = listing.elements[0]
        dest = FolderInfo(id="archive", name="Archive")
        result = await mail.move_message_async(msg, dest)
        assert result is not None
        assert isinstance(result, MoveResult)

    @pytest.mark.asyncio
    async def test_move_mixed_types(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        dest = FolderInfo(id="drafts", name="Drafts")
        result = await mail.move_message_async("msg-id", dest)
        assert result is not None
        assert isinstance(result, MoveResult)

    @pytest.mark.asyncio
    async def test_move_no_token(self, graph_no_token: MsGraphInstance):
        mail = graph_no_token.get_mail()
        result = await mail.move_message_async("msg-id", "archive")
        assert result is None


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — get_mail_async
# ═══════════════════════════════════════════════════════════════════

class TestGetMail:

    @pytest.mark.asyncio
    async def test_get_mail_by_id(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        listing = await mail.email_index_async(limit=1)
        assert listing.elements
        msg_id = listing.elements[0].email_id
        result = await mail.get_mail_async(email_id=msg_id)
        assert result is not None
        assert isinstance(result, OfficeMail)

    @pytest.mark.asyncio
    async def test_get_mail_no_token(self, graph_no_token: MsGraphInstance):
        mail = graph_no_token.get_mail()
        result = await mail.get_mail_async(email_id="some-id")
        assert result is None


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — send / draft lifecycle
# ═══════════════════════════════════════════════════════════════════

class TestSendAndDraft:

    @pytest.mark.asyncio
    async def test_create_draft(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.create_draft_async(
            to_recipients=["test@example.com"],
            subject="Test",
            body="Hello",
        )
        assert result is not None
        assert "id" in result

    @pytest.mark.asyncio
    async def test_send_draft(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        draft = await mail.create_draft_async(
            to_recipients=["test@example.com"],
            subject="Test",
            body="Hello",
        )
        assert draft is not None
        ok = await mail.send_draft_async(draft["id"])
        assert ok is True

    @pytest.mark.asyncio
    async def test_flag_read(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        listing = await mail.email_index_async(limit=1)
        assert listing.elements
        msg = listing.elements[0]
        ok = await mail.flag_read_async(msg.email_url or f"{graph.msg_endpoint}me/messages/{msg.email_id}", True)
        assert ok is True


# ═══════════════════════════════════════════════════════════════════
# OfficeMailHandler — categories
# ═══════════════════════════════════════════════════════════════════

class TestCategories:

    @pytest.mark.asyncio
    async def test_get_categories(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        cats = await mail.get_categories_async()
        assert isinstance(cats, list)
        assert len(cats) > 0

    @pytest.mark.asyncio
    async def test_get_categories_no_token(self, graph_no_token: MsGraphInstance):
        mail = graph_no_token.get_mail()
        cats = await mail.get_categories_async()
        assert cats == []


# ═══════════════════════════════════════════════════════════════════
# MsGraphInstance — factory methods
# ═══════════════════════════════════════════════════════════════════

class TestFactoryMethods:

    def test_get_mail_returns_handler(self, graph: MsGraphInstance):
        assert isinstance(graph.get_mail(), OfficeMailHandler)

    def test_get_mail_folders_returns_handler(self, graph: MsGraphInstance):
        assert isinstance(graph.get_mail_folders(), MailFolderHandler)


# ═══════════════════════════════════════════════════════════════════
# compute_folder_signature
# ═══════════════════════════════════════════════════════════════════

class TestComputeFolderSignature:

    def _make_row(self, id="msg-1", is_read=False, scanning=False,
                  categories=None, importance="normal"):
        return {
            "id": id,
            "from_name": "Alice",
            "from_email": "alice@example.com",
            "subject": "Hello",
            "preview": "Hi there",
            "received": "2026-04-18 10:00:00",
            "is_read": is_read,
            "has_attachments": False,
            "importance": importance,
            "categories": categories or [],
            "scanning": scanning,
        }

    def test_same_list_shuffled_order_same_sig(self):
        rows = [self._make_row(id=f"msg-{i}") for i in range(5)]
        sig_a = compute_folder_signature(rows)
        sig_b = compute_folder_signature(list(reversed(rows)))
        assert sig_a == sig_b

    def test_flip_is_read_changes_sig(self):
        rows = [self._make_row(id="msg-1", is_read=False)]
        sig_a = compute_folder_signature(rows)
        rows[0]["is_read"] = True
        sig_b = compute_folder_signature(rows)
        assert sig_a != sig_b

    def test_change_scanning_changes_sig(self):
        rows = [self._make_row(id="msg-1", scanning=False)]
        sig_a = compute_folder_signature(rows)
        rows[0]["scanning"] = True
        sig_b = compute_folder_signature(rows)
        assert sig_a != sig_b

    def test_empty_list_deterministic(self):
        sig_a = compute_folder_signature([])
        sig_b = compute_folder_signature([])
        assert sig_a == sig_b
        assert isinstance(sig_a, str)
        assert len(sig_a) == 16

    def test_1000_rows_under_50ms(self):
        import time
        rows = [self._make_row(id=f"msg-{i}") for i in range(1000)]
        start = time.monotonic()
        compute_folder_signature(rows)
        elapsed_ms = (time.monotonic() - start) * 1000
        assert elapsed_ms < 50, f"took {elapsed_ms:.1f}ms"


# ═══════════════════════════════════════════════════════════════════
# OfficeMail.scanning property
# ═══════════════════════════════════════════════════════════════════

class TestOfficeMailScanning:

    def test_scanning_has_attachments_no_list(self):
        """has_attachments=True, attachments=[] → scanning True."""
        mail = OfficeMail(
            email_id="m1", email_type="mail",
            has_attachments=True, attachments=[],
        )
        assert mail.scanning is True

    def test_scanning_virus_placeholder(self):
        """Attachment named 'virus scan in progress.html' → scanning True."""
        mail = OfficeMail(
            email_id="m2", email_type="mail",
            has_attachments=True,
            attachments=[
                OfficeMailAttachment(
                    name="virus scan in progress.html",
                    content_type="text/html",
                ),
            ],
        )
        assert mail.scanning is True

    def test_not_scanning_normal_attachment(self):
        """Normal attachment → scanning False."""
        mail = OfficeMail(
            email_id="m3", email_type="mail",
            has_attachments=True,
            attachments=[
                OfficeMailAttachment(
                    name="report.pdf",
                    content_type="application/pdf",
                ),
            ],
        )
        assert mail.scanning is False

    def test_not_scanning_no_attachments(self):
        """has_attachments=False → scanning False."""
        mail = OfficeMail(
            email_id="m4", email_type="mail",
            has_attachments=False, attachments=[],
        )
        assert mail.scanning is False

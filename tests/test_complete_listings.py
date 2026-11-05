"""Unit tests for the completeness-critical Graph operations.

Guards the regression where a mailbox-wide attachment listing put
``hasAttachments`` before ``receivedDateTime`` in ``$filter`` while ordering by
``receivedDateTime``: Graph answered 400 ``InefficientFilter``, the caller got an
empty list, and every receipt mail was silently skipped. Complete listings must
page through ``@odata.nextLink`` and raise on failure; conditional drive writes
must surface a lost race as ``DriveItemConflictError``.
"""

from __future__ import annotations

from datetime import date, datetime
from urllib.parse import unquote

import pytest

from office_con.msgraph.errors import DriveItemConflictError, GraphRequestError
from office_con.msgraph.files_handler import FilesHandler
from office_con.msgraph.mail_handler import OfficeMailAttachment
from office_con.msgraph.ms_graph_handler import MsGraphInstance


class _FakeResponse:
    def __init__(self, status_code: int, payload: dict | None = None, text: str = "") -> None:
        self.status_code = status_code
        self._payload = payload or {}
        self.text = text
        self.content = b""

    def json(self) -> dict:
        return self._payload


def _graph_with_sequence(responses: list[_FakeResponse]):
    """Build a graph whose run_async returns one prepared response per call."""
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.cache_dict["access_token"] = "fake-token"
    g.email = "me@example.com"
    calls: list[dict] = []

    async def fake_run_async(*, url, method="GET", json=None, data=None, token=None, add_headers=None):
        calls.append({"url": url, "method": method, "json": json, "data": data, "headers": add_headers})
        return responses.pop(0)

    g.run_async = fake_run_async  # type: ignore[assignment]
    return g, calls


def _message(mid: str) -> dict:
    return {"id": mid, "subject": f"Invoice {mid}", "receivedDateTime": "2026-10-13T07:49:00Z",
            "hasAttachments": True, "from": {"emailAddress": {"address": "billing@example.com"}}}


@pytest.mark.asyncio
async def test_list_messages_orders_filter_and_follows_next_link():
    g, calls = _graph_with_sequence([
        _FakeResponse(200, {"value": [_message("a")], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"}),
        _FakeResponse(200, {"value": [_message("b")]}),
    ])
    result = await g.get_mail().list_messages_async(received_since=date(2025, 9, 1), has_attachments=True)

    assert [m.email_id for m in result.elements] == ["a", "b"]
    first = unquote(calls[0]["url"])
    assert "/me/messages?" in first  # mailbox-wide, not inbox only
    assert "$filter=receivedDateTime ge 2025-09-01T00:00:00Z and hasAttachments eq true" in first
    assert "$orderby=receivedDateTime desc" in first
    assert calls[1]["url"] == "https://graph.microsoft.com/v1.0/next"


@pytest.mark.asyncio
async def test_list_messages_raises_instead_of_returning_empty():
    g, _ = _graph_with_sequence([_FakeResponse(400, text='{"error":{"code":"InefficientFilter"}}')])
    with pytest.raises(GraphRequestError) as exc:
        await g.get_mail().list_messages_async(received_since=datetime(2025, 9, 1), has_attachments=True)
    assert exc.value.status_code == 400


@pytest.mark.asyncio
async def test_list_messages_subject_prefix_in_folder():
    g, calls = _graph_with_sequence([_FakeResponse(200, {"value": [_message("s")]})])
    await g.get_mail().list_messages_async(folder="sent", subject_prefix="Genehmigt")
    url = unquote(calls[0]["url"])
    assert "/me/mailFolders/sentitems/messages?" in url
    assert "startswith(subject,'Genehmigt')" in url
    assert "$orderby" not in url  # no date filter → no ordering (would need the property in $filter)


@pytest.mark.asyncio
async def test_patch_draft_sends_only_given_fields():
    g, calls = _graph_with_sequence([_FakeResponse(200, {"id": "d1"})])
    assert await g.get_mail().patch_draft_async("d1", body="<p>new</p>")
    assert calls[0]["method"] == "PATCH"
    assert calls[0]["json"] == {"body": {"contentType": "HTML", "content": "<p>new</p>"}}


@pytest.mark.asyncio
async def test_save_draft_uses_upload_session_for_large_attachments():
    g, calls = _graph_with_sequence([
        _FakeResponse(201, {"id": "d1", "webLink": "https://outlook/d1"}),
        _FakeResponse(201, {"uploadUrl": "https://upload.example/session"}),
    ])
    uploaded: list[tuple[str, int]] = []

    async def fake_upload(upload_url, content):
        uploaded.append((upload_url, len(content)))
        return {"id": "att"}

    g.upload_session_async = fake_upload  # type: ignore[assignment]
    big = OfficeMailAttachment(name="scan.pdf", content_type="application/pdf", content_bytes=b"x" * (4 * 1024 * 1024))
    draft = await g.get_mail().save_draft_async(["to@example.com"], "Subject", "<p>Hi</p>", attachments=[big])

    assert draft is not None and draft.id == "d1" and draft.failed_attachments == []
    assert calls[1]["url"].endswith("/me/messages/d1/attachments/createUploadSession")
    assert uploaded == [("https://upload.example/session", 4 * 1024 * 1024)]


@pytest.mark.asyncio
async def test_save_draft_reports_failed_attachments():
    g, _ = _graph_with_sequence([
        _FakeResponse(201, {"id": "d1"}),
        _FakeResponse(500, text="boom"),
    ])
    small = OfficeMailAttachment(name="a.pdf", content_type="application/pdf", content_bytes=b"pdf")
    draft = await g.get_mail().save_draft_async(["to@example.com"], "S", "B", attachments=[small])
    assert draft is not None and draft.failed_attachments == ["a.pdf"]


@pytest.mark.asyncio
async def test_calendar_events_follow_next_link():
    event = {"id": "e", "subject": "Trip", "start": {"dateTime": "2026-10-12T08:00:00"},
             "end": {"dateTime": "2026-10-16T18:00:00"}, "isCancelled": True}
    g, calls = _graph_with_sequence([
        _FakeResponse(200, {"value": [event], "@odata.nextLink": "https://graph.microsoft.com/v1.0/cal-next"}),
        _FakeResponse(200, {"value": [dict(event, id="f", isCancelled=False)]}),
    ])
    result = await g.get_calendar().get_events_async(datetime(2026, 1, 1), datetime(2026, 12, 31), calendar_id="c1", limit=500)
    assert [e.id for e in result.events] == ["e", "f"]
    assert result.events[0].is_cancelled and not result.events[1].is_cancelled
    assert calls[1]["url"] == "https://graph.microsoft.com/v1.0/cal-next"


@pytest.mark.asyncio
async def test_conditional_upload_raises_conflict():
    g, calls = _graph_with_sequence([_FakeResponse(412, text="precondition failed")])
    with pytest.raises(DriveItemConflictError):
        await FilesHandler(g).upload_file_by_path_async("tools/data/manifest.json", b"{}", if_match='"etag-1"')
    assert calls[0]["headers"]["If-Match"] == '"etag-1"'
    assert "root:/tools/data/manifest.json:/content?@microsoft.graph.conflictBehavior=replace" in calls[0]["url"]


@pytest.mark.asyncio
async def test_create_only_upload_uses_fail_conflict_behavior():
    item = {"id": "i1", "name": "lease.json", "eTag": '"e1"', "file": {}}
    g, calls = _graph_with_sequence([_FakeResponse(201, item)])
    created = await FilesHandler(g).upload_file_by_path_async("tools/lease.json", b"{}", create_only=True)
    assert created.e_tag == '"e1"'
    assert calls[0]["url"].endswith("conflictBehavior=fail")


@pytest.mark.asyncio
async def test_item_by_path_missing_is_none():
    g, _ = _graph_with_sequence([_FakeResponse(404)])
    assert await FilesHandler(g).get_item_by_path_async("tools/nothing.json") is None


@pytest.mark.asyncio
async def test_search_items_pages_while_more_results():
    def page(ids: list[str], more: bool) -> _FakeResponse:
        return _FakeResponse(200, {"value": [{"hitsContainers": [{
            "total": 3, "moreResultsAvailable": more,
            "hits": [{"resource": {"id": i, "name": f"{i}.pdf", "parentReference": {"driveId": "d"}}} for i in ids],
        }]}]})

    g, calls = _graph_with_sequence([page(["a", "b"], True), page(["c"], False)])
    result = await FilesHandler(g).search_items_async("invoice", limit=400)
    assert [i.id for i in result.items] == ["a", "b", "c"]
    assert [c["json"]["requests"][0]["from"] for c in calls] == [0, 200]

"""Unit tests for FilesHandler.search_items_async.

Guards the regression where file search only hit the user's personal OneDrive
(``/me/drive/root/search``) and therefore never found files living in SharePoint
document libraries. The default (no ``drive_id``) path must use the tenant-wide
Microsoft Search API (``POST /search/query`` with ``entityTypes: ["driveItem"]``)
and surface each hit's ``drive_id`` so callers can open SharePoint results.
"""

from __future__ import annotations

import pytest

from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.msgraph.files_handler import FilesHandler


class _FakeResponse:
    def __init__(self, status_code: int, payload: dict) -> None:
        self.status_code = status_code
        self._payload = payload

    def json(self) -> dict:
        return self._payload


def _graph_with_capture(payload: dict, status_code: int = 200):
    """Build a graph whose run_async records its call and returns ``payload``."""
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.cache_dict["access_token"] = "fake-token"
    calls: list[dict] = []

    async def fake_run_async(*, url, method="GET", json=None, token=None, add_headers=None):
        calls.append({"url": url, "method": method, "json": json})
        return _FakeResponse(status_code, payload)

    g.run_async = fake_run_async  # type: ignore[assignment]
    return g, calls


def _graph_with_sequence(responses: list[_FakeResponse]):
    """Build a graph whose run_async returns one prepared response per call."""
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.cache_dict["access_token"] = "fake-token"
    calls: list[dict] = []

    async def fake_run_async(*, url, method="GET", json=None, token=None, add_headers=None):
        calls.append({"url": url, "method": method, "json": json})
        return responses.pop(0)

    g.run_async = fake_run_async  # type: ignore[assignment]
    return g, calls


_UNIFIED_HIT = {
    "value": [{
        "hitsContainers": [{
            "total": 1,
            "hits": [{
                "resource": {
                    "id": "ITEM123",
                    "name": "Playbook_ProjectAlpha.docx",
                    "size": 4096,
                    "webUrl": "https://contoso.sharepoint.com/sites/x/Playbook.docx",
                    "file": {"mimeType": "application/vnd.openxmlformats"},
                    "parentReference": {"driveId": "DRIVE_SP_99", "path": "/drive/root:/Docs"},
                },
            }],
        }],
    }],
}


@pytest.mark.asyncio
async def test_search_no_drive_uses_unified_tenant_wide_endpoint():
    graph, calls = _graph_with_capture(_UNIFIED_HIT)
    handler = FilesHandler(graph)

    result = await handler.search_items_async("ProjectAlpha", limit=10)

    # One call, to the unified Search endpoint, as a POST with driveItem entity type.
    assert len(calls) == 1
    assert calls[0]["url"].endswith("search/query")
    assert calls[0]["method"] == "POST"
    body = calls[0]["json"]["requests"][0]
    assert body["entityTypes"] == ["driveItem"]
    assert body["query"]["queryString"] == "ProjectAlpha"
    assert body["size"] == 10

    # Hit is parsed and the SharePoint drive_id is surfaced for follow-up fetches.
    assert result.total_items == 1
    assert len(result.items) == 1
    item = result.items[0]
    assert item.id == "ITEM123"
    assert item.name == "Playbook_ProjectAlpha.docx"
    assert item.drive_id == "DRIVE_SP_99"
    assert item.parent_path == "/drive/root:/Docs"


@pytest.mark.asyncio
async def test_search_with_drive_id_uses_scoped_endpoint():
    graph, calls = _graph_with_capture({"value": []})
    handler = FilesHandler(graph)

    await handler.search_items_async("report", drive_id="DRIVE_ABC", limit=5)

    assert len(calls) == 1
    assert calls[0]["method"] == "GET"
    assert "drives/DRIVE_ABC/root/search(q='report')" in calls[0]["url"]
    assert "search/query" not in calls[0]["url"]


@pytest.mark.asyncio
async def test_search_empty_result_is_empty_not_fallback():
    graph, _ = _graph_with_capture({"value": [{"hitsContainers": [{"total": 0, "hits": []}]}]})
    handler = FilesHandler(graph)

    result = await handler.search_items_async("nomatch")

    assert result.total_items == 0
    assert result.items == []


@pytest.mark.asyncio
async def test_search_falls_back_to_personal_onedrive_when_unified_search_fails():
    fallback_payload = {
        "value": [{
            "id": "OD_ITEM",
            "name": "report.docx",
            "parentReference": {"driveId": "DRIVE_PERSONAL"},
        }]
    }
    graph, calls = _graph_with_sequence([
        _FakeResponse(403, {"error": {"code": "accessDenied"}}),
        _FakeResponse(200, fallback_payload),
    ])
    handler = FilesHandler(graph)

    result = await handler.search_items_async("report", limit=5)

    assert len(calls) == 2
    assert calls[0]["url"].endswith("search/query")
    assert calls[0]["method"] == "POST"
    assert calls[1]["method"] == "GET"
    assert "me/drive/root/search(q='report')" in calls[1]["url"]
    assert result.total_items == 1
    assert result.items[0].id == "OD_ITEM"
    assert result.items[0].drive_id == "DRIVE_PERSONAL"

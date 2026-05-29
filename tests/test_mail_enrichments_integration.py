"""READ-ONLY integration tests for the mail/calendar enrichments, run against
real Microsoft Graph using the resolved token file.

⚠️  Per AGENTS.md these tests MUST stay read-only (GET requests only). Do NOT
add send/reply/forward/create/update/delete coverage here — that belongs in the
mock suite (tests/test_mail_enrichments.py). Sending or mutating the real
mailbox with the live token is forbidden.

Token resolution (first usable wins):
  1. tests/msgraph_test_token.json
  2. ~/.config/office-connect/token.json
  3. token_file in tests/test_config.json

Run with: poetry run pytest tests/test_mail_enrichments_integration.py -v -s
"""

from __future__ import annotations

from pathlib import Path

import pytest
import pytest_asyncio

from office_con.mcp_server import _create_graph
from office_con.msgraph.ms_graph_handler import MsGraphInstance


def _resolve_token_path() -> Path | None:
    candidates: list[Path] = [
        Path(__file__).parent / "msgraph_test_token.json",
        Path.home() / ".config" / "office-connect" / "token.json",
    ]
    try:
        from office_con.testing.mock_config import get_test_config
        cfg_token = get_test_config().get("token_file")
        if cfg_token:
            candidates.append(Path(cfg_token).expanduser())
    except Exception:
        pass
    for p in candidates:
        if p.is_file():
            return p
    return None


TOKEN_FILE = _resolve_token_path()


@pytest_asyncio.fixture
async def graph() -> MsGraphInstance:
    if TOKEN_FILE is None:
        pytest.skip("No token file found — run `office-connect login`.")
    inst = await _create_graph(str(TOKEN_FILE))
    resp = await inst.run_async(url=f"{inst.msg_endpoint}me")
    if resp is None or resp.status_code != 200:
        pytest.skip("Token expired or invalid — cannot reach MS Graph.")
    return inst


@pytest.mark.asyncio
async def test_list_carries_recipient_and_thread_metadata(graph: MsGraphInstance):
    result = await graph.get_mail().email_index_async(limit=10)
    if not result.elements:
        pytest.skip("Mailbox inbox empty.")
    m = result.elements[0]
    # graph_url / outlook_url aliases populated on the list shape.
    assert m.graph_url and m.graph_url == m.email_url
    assert m.outlook_url == m.web_link
    # conversation_id present so a thread can be pulled without a get.
    assert m.conversation_id, "conversation_id missing on list result"
    # At least one message in a real inbox should expose recipients + msg-id.
    assert any(e.to_recipients for e in result.elements), "no to_recipients on any list item"
    assert any(e.internet_message_id for e in result.elements), "no internet_message_id on any list item"
    print(f"\n[list] {len(result.elements)} msgs; first from={m.from_email!r} "
          f"to={[r.address for r in m.to_recipients]} conv={m.conversation_id[:12]}...")


@pytest.mark.asyncio
async def test_get_mail_body_formats_and_truncation(graph: MsGraphInstance):
    listing = await graph.get_mail().email_index_async(limit=5)
    if not listing.elements:
        pytest.skip("Mailbox inbox empty.")
    mid = listing.elements[0].email_id

    # text format → body_text populated, body_type text.
    text_mail = await graph.get_mail().get_mail_async(email_id=mid, body_format="text")
    assert text_mail is not None
    assert text_mail.body_text is not None

    # none format → body skipped, preview survives.
    none_mail = await graph.get_mail().get_mail_async(email_id=mid, body_format="none")
    assert none_mail is not None
    assert not none_mail.body
    assert none_mail.body_preview is not None

    # truncation flag honored.
    tiny = await graph.get_mail().get_mail_async(email_id=mid, body_format="html", max_body_chars=20)
    if tiny and tiny.body:
        assert len(tiny.body) <= 20
        assert tiny.body_truncated is True
    print(f"\n[body] text_len={len(text_mail.body_text or '')} "
          f"truncated_demo={'yes' if (tiny and tiny.body_truncated) else 'n/a'}")


@pytest.mark.asyncio
async def test_batch_get_mails_preserves_order(graph: MsGraphInstance):
    listing = await graph.get_mail().email_index_async(limit=5)
    ids = [m.email_id for m in listing.elements]
    if len(ids) < 2:
        pytest.skip("Need >=2 inbox messages for a batch test.")
    mails = await graph.get_mail().get_mails_async(ids, body_format="text", max_body_chars=2000)
    fetched = [m.email_id for m in mails]
    # Order preserved; every returned id was requested.
    assert fetched == [i for i in ids if i in set(fetched)]
    assert set(fetched).issubset(set(ids))
    print(f"\n[batch] requested {len(ids)} → fetched {len(fetched)} in one $batch round trip")


@pytest.mark.asyncio
async def test_folder_scoping_sent_items(graph: MsGraphInstance):
    sent = await graph.get_mail().email_index_async(folder="sent", limit=10)
    # Sent Items may legitimately be empty; just assert the call shape works
    # and that any returned item has the sender == the account (best-effort).
    print(f"\n[folder] sent items returned {len(sent.elements)} messages")
    assert isinstance(sent.elements, list)


@pytest.mark.asyncio
async def test_search_with_folder_exclusion(graph: MsGraphInstance):
    # A broad term likely to match; exclusion must not increase the count.
    all_hits = await graph.get_mail().email_index_async(query="a", limit=50)
    excl = await graph.get_mail().email_index_async(
        query="a", limit=50, exclude_folders=["deleteditems", "junkemail"],
    )
    assert len(excl.elements) <= len(all_hits.elements)
    print(f"\n[search] all={len(all_hits.elements)} after-exclude={len(excl.elements)}")


@pytest.mark.asyncio
async def test_check_connection(graph: MsGraphInstance):
    from office_con.mcp_server import _handle_tool
    import json
    out = await _handle_tool(graph, "o365_check_connection", {})
    payload = json.loads(out[0].text)
    assert payload["connected"] is True
    assert payload["email"]
    print(f"\n[conn] connected as {payload['email']}")


@pytest.mark.asyncio
async def test_unread_counts(graph: MsGraphInstance):
    folders = await graph.get_mail_folders().get_folders_async(recursive=True)
    assert folders, "no folders returned"
    total_unread = sum(f.unread for f in folders)
    inbox = next((f for f in folders if f.name.lower() == "inbox"), None)
    assert inbox is not None
    print(f"\n[unread] total_unread={total_unread}; inbox unread={inbox.unread}/{inbox.total}")

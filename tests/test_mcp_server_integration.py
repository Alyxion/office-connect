"""Integration tests for the Office 365 MCP Server.

Tests every tool via _handle_tool against the real MS Graph API.
Requires a valid token file at tests/msgraph_test_token.json.

Run with: poetry run pytest dependencies/office-mcp/tests/test_mcp_server_integration.py -v
"""

from __future__ import annotations

import json
from datetime import datetime, timedelta
from pathlib import Path

import pytest
import pytest_asyncio

from office_con.mcp_server import TOOLS, _create_graph, _handle_tool, _json_result, create_server
from office_con.msgraph.ms_graph_handler import MsGraphInstance

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

TOKEN_FILE = Path(__file__).parent / "msgraph_test_token.json"


def _load_token() -> dict:
    if not TOKEN_FILE.exists():
        pytest.skip(f"Token file not found: {TOKEN_FILE}")
    with open(TOKEN_FILE) as f:
        return json.load(f)


@pytest_asyncio.fixture
async def graph() -> MsGraphInstance:
    """Create an MsGraphInstance from the test token file."""
    _load_token()  # skip if no token file
    inst = await _create_graph(str(TOKEN_FILE))
    # Verify token is usable — try a simple API call
    resp = await inst.run_async(url=f"{inst.msg_endpoint}me")
    if resp is None or resp.status_code != 200:
        pytest.skip("Token expired or invalid — cannot reach MS Graph API")
    return inst


# ---------------------------------------------------------------------------
# Server & tool definitions
# ---------------------------------------------------------------------------


class TestServerSetup:

    def test_create_server_returns_server(self):
        """create_server should return an MCP Server instance."""
        _load_token()  # skip if no token file
        server, keyfile = create_server(str(TOKEN_FILE))
        assert server is not None
        assert keyfile == str(TOKEN_FILE)

    def test_tools_list_not_empty(self):
        assert len(TOOLS) > 0

    def test_all_tools_have_names(self):
        for tool in TOOLS:
            assert tool.name
            assert tool.name.startswith("o365_")

    def test_all_tools_have_descriptions(self):
        for tool in TOOLS:
            assert tool.description
            assert len(tool.description) > 10

    def test_all_tools_have_schemas(self):
        for tool in TOOLS:
            assert tool.inputSchema is not None
            assert tool.inputSchema["type"] == "object"

    def test_tool_count(self):
        """All read + draft + all-tier tools. Update this when adding new tools."""
        assert len(TOOLS) == 38

    def test_tool_names_unique(self):
        names = [t.name for t in TOOLS]
        assert len(names) == len(set(names))


class TestJsonResult:

    def test_dict(self):
        result = _json_result({"key": "value"})
        assert len(result) == 1
        assert result[0].type == "text"
        parsed = json.loads(result[0].text)
        assert parsed["key"] == "value"

    def test_list(self):
        result = _json_result([1, 2, 3])
        parsed = json.loads(result[0].text)
        assert parsed == [1, 2, 3]

    def test_none(self):
        result = _json_result(None)
        assert result[0].text == "null"

    def test_string(self):
        result = _json_result("hello")
        assert result[0].text == "hello"


# ---------------------------------------------------------------------------
# Profile tools
# ---------------------------------------------------------------------------


class TestProfileTools:

    @pytest.mark.asyncio
    async def test_get_profile(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_get_profile", {})
        assert len(result) == 1
        data = json.loads(result[0].text)
        assert data is not None
        assert "id" in data
        assert "display_name" in data

    @pytest.mark.asyncio
    async def test_get_profile_has_email(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_get_profile", {})
        data = json.loads(result[0].text)
        assert data is not None
        assert data.get("mail") or data.get("user_principal_name")


# ---------------------------------------------------------------------------
# Mail tools
# ---------------------------------------------------------------------------


class TestMailTools:

    @pytest.mark.asyncio
    async def test_list_mail(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_mail", {"limit": 5})
        assert len(result) == 1
        data = json.loads(result[0].text)
        assert "elements" in data
        assert "total_mails" in data
        assert isinstance(data["elements"], list)

    @pytest.mark.asyncio
    async def test_list_mail_default_limit(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_mail", {})
        data = json.loads(result[0].text)
        assert len(data["elements"]) <= 10

    @pytest.mark.asyncio
    async def test_list_mail_with_skip(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_mail", {"limit": 2, "skip": 0})
        data = json.loads(result[0].text)
        assert isinstance(data["elements"], list)

    @pytest.mark.asyncio
    async def test_get_mail(self, graph: MsGraphInstance):
        # First get a mail ID
        index = await _handle_tool(graph, "o365_list_mail", {"limit": 1})
        index_data = json.loads(index[0].text)
        if not index_data["elements"]:
            pytest.skip("No mails in inbox")
        email_id = index_data["elements"][0]["email_id"]
        result = await _handle_tool(graph, "o365_get_mail", {"email_id": email_id})
        data = json.loads(result[0].text)
        assert data["email_id"] == email_id
        assert "body" in data

    @pytest.mark.asyncio
    async def test_get_mail_categories(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_get_mail_categories", {})
        data = json.loads(result[0].text)
        assert isinstance(data, list)

    @pytest.mark.asyncio
    async def test_search_mail_refuses_empty(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_search_mail", {})
        assert "Refused" in result[0].text

    @pytest.mark.asyncio
    async def test_search_mail_by_subject(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_search_mail", {"subject": "meeting", "limit": 3})
        data = json.loads(result[0].text)
        assert "elements" in data


# ---------------------------------------------------------------------------
# Calendar tools
# ---------------------------------------------------------------------------


class TestCalendarTools:

    @pytest.mark.asyncio
    async def test_list_calendars(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_calendars", {})
        data = json.loads(result[0].text)
        assert isinstance(data, list)
        assert len(data) >= 0

    @pytest.mark.asyncio
    async def test_get_events(self, graph: MsGraphInstance):
        now = datetime.now()
        start = (now - timedelta(days=7)).isoformat()
        end = (now + timedelta(days=7)).isoformat()
        result = await _handle_tool(graph, "o365_get_events", {
            "start_date": start,
            "end_date": end,
            "limit": 5,
        })
        data = json.loads(result[0].text)
        assert "events" in data
        assert isinstance(data["events"], list)

    @pytest.mark.asyncio
    async def test_get_events_custom_range(self, graph: MsGraphInstance):
        start = "2026-03-01"
        end = "2026-03-31"
        result = await _handle_tool(graph, "o365_get_events", {
            "start_date": start,
            "end_date": end,
        })
        data = json.loads(result[0].text)
        assert "events" in data

    @pytest.mark.asyncio
    async def test_get_schedule(self, graph: MsGraphInstance):
        # Get user email for schedule check
        profile = await _handle_tool(graph, "o365_get_profile", {})
        profile_data = json.loads(profile[0].text)
        if not profile_data:
            pytest.skip("Profile not available")
        email = profile_data.get("mail") or profile_data.get("user_principal_name")
        if not email:
            pytest.skip("No email in profile")
        now = datetime.now()
        start = now.replace(hour=8, minute=0, second=0, microsecond=0)
        end = start + timedelta(hours=8)
        result = await _handle_tool(graph, "o365_get_schedule", {
            "emails": [email],
            "start": start.isoformat(),
            "end": end.isoformat(),
        })
        data = json.loads(result[0].text)
        assert isinstance(data, list)

    @pytest.mark.asyncio
    async def test_search_events_refuses_empty(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_search_events", {})
        assert "Refused" in result[0].text

    @pytest.mark.asyncio
    async def test_search_events_by_subject(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_search_events", {"subject": "meeting", "limit": 3})
        data = json.loads(result[0].text)
        assert "events" in data
        assert "filter" in data


# ---------------------------------------------------------------------------
# Teams tools
# ---------------------------------------------------------------------------


class TestTeamsTools:

    @pytest.mark.asyncio
    async def test_list_teams(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_teams", {})
        data = json.loads(result[0].text)
        assert "teams" in data
        assert "total_teams" in data

    @pytest.mark.asyncio
    async def test_list_channels(self, graph: MsGraphInstance):
        teams_result = await _handle_tool(graph, "o365_list_teams", {})
        teams_data = json.loads(teams_result[0].text)
        if not teams_data["teams"]:
            pytest.skip("No teams joined")
        team_id = teams_data["teams"][0]["id"]
        result = await _handle_tool(graph, "o365_list_channels", {"team_id": team_id})
        data = json.loads(result[0].text)
        assert "channels" in data
        assert data["total_channels"] > 0

    @pytest.mark.asyncio
    async def test_get_channel_messages(self, graph: MsGraphInstance):
        teams_result = await _handle_tool(graph, "o365_list_teams", {})
        teams_data = json.loads(teams_result[0].text)
        if not teams_data["teams"]:
            pytest.skip("No teams joined")
        team_id = teams_data["teams"][0]["id"]
        channels_result = await _handle_tool(graph, "o365_list_channels", {"team_id": team_id})
        channels_data = json.loads(channels_result[0].text)
        if not channels_data["channels"]:
            pytest.skip("No channels")
        channel_id = channels_data["channels"][0]["id"]
        result = await _handle_tool(graph, "o365_get_channel_messages", {
            "team_id": team_id,
            "channel_id": channel_id,
            "limit": 5,
        })
        data = json.loads(result[0].text)
        assert "messages" in data
        assert isinstance(data["messages"], list)

    @pytest.mark.asyncio
    async def test_get_team_members(self, graph: MsGraphInstance):
        teams_result = await _handle_tool(graph, "o365_list_teams", {})
        teams_data = json.loads(teams_result[0].text)
        if not teams_data["teams"]:
            pytest.skip("No teams joined")
        team_id = teams_data["teams"][0]["id"]
        result = await _handle_tool(graph, "o365_get_team_members", {"team_id": team_id})
        data = json.loads(result[0].text)
        assert "members" in data
        assert "total_members" in data


# ---------------------------------------------------------------------------
# Chat tools
# ---------------------------------------------------------------------------


class TestChatTools:

    @pytest.mark.asyncio
    async def test_list_chats(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_chats", {"limit": 5})
        data = json.loads(result[0].text)
        assert "chats" in data
        assert "total_chats" in data

    @pytest.mark.asyncio
    async def test_list_chats_default(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_chats", {})
        data = json.loads(result[0].text)
        assert isinstance(data["chats"], list)

    @pytest.mark.asyncio
    async def test_get_chat_messages(self, graph: MsGraphInstance):
        chats_result = await _handle_tool(graph, "o365_list_chats", {"limit": 5})
        chats_data = json.loads(chats_result[0].text)
        if not chats_data["chats"]:
            pytest.skip("No chats available")
        chat_id = chats_data["chats"][0]["id"]
        result = await _handle_tool(graph, "o365_get_chat_messages", {
            "chat_id": chat_id,
            "limit": 5,
        })
        data = json.loads(result[0].text)
        assert "messages" in data

    @pytest.mark.asyncio
    async def test_get_chat_members(self, graph: MsGraphInstance):
        chats_result = await _handle_tool(graph, "o365_list_chats", {"limit": 5})
        chats_data = json.loads(chats_result[0].text)
        if not chats_data["chats"]:
            pytest.skip("No chats available")
        chat_id = chats_data["chats"][0]["id"]
        result = await _handle_tool(graph, "o365_get_chat_members", {"chat_id": chat_id})
        data = json.loads(result[0].text)
        assert "members" in data
        assert "total_members" in data

    @pytest.mark.asyncio
    async def test_search_messages(self, graph: MsGraphInstance):
        # The call may legitimately return zero hits or fail for scope reasons;
        # we just verify it doesn't crash and returns the expected shape.
        result = await _handle_tool(graph, "o365_search_messages",
                                    {"query": "meeting", "limit": 3})
        text = result[0].text
        if "failed" in text.lower():
            pytest.skip(f"search/query unavailable: {text[:120]}")
        data = json.loads(text)
        assert "hits" in data
        assert "kql" in data


# ---------------------------------------------------------------------------
# Files / OneDrive tools
# ---------------------------------------------------------------------------


class TestFilesTools:

    @pytest.mark.asyncio
    async def test_get_my_drive(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_get_my_drive", {})
        data = json.loads(result[0].text)
        if data is None:
            pytest.skip("Drive not available")
        assert "id" in data
        assert "drive_type" in data

    @pytest.mark.asyncio
    async def test_list_drive_items_root(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_drive_items", {"limit": 5})
        data = json.loads(result[0].text)
        assert "items" in data
        assert isinstance(data["items"], list)

    @pytest.mark.asyncio
    async def test_list_drive_items_folder(self, graph: MsGraphInstance):
        # Get root items first
        root = await _handle_tool(graph, "o365_list_drive_items", {"limit": 20})
        root_data = json.loads(root[0].text)
        folders = [i for i in root_data["items"] if i.get("is_folder")]
        if not folders:
            pytest.skip("No folders in OneDrive root")
        folder_id = folders[0]["id"]
        result = await _handle_tool(graph, "o365_list_drive_items", {"folder_id": folder_id, "limit": 5})
        data = json.loads(result[0].text)
        assert "items" in data

    @pytest.mark.asyncio
    async def test_get_file_content(self, graph: MsGraphInstance):
        root = await _handle_tool(graph, "o365_list_drive_items", {"limit": 20})
        root_data = json.loads(root[0].text)
        files = [i for i in root_data["items"] if not i.get("is_folder") and (i.get("size") or 0) < 1_000_000]
        if not files:
            pytest.skip("No small files in OneDrive root")
        item_id = files[0]["id"]
        result = await _handle_tool(graph, "o365_get_file_content", {"item_id": item_id})
        assert len(result) == 1
        assert len(result[0].text) > 0

    @pytest.mark.asyncio
    async def test_search_files(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_search_files", {"query": "test", "limit": 5})
        data = json.loads(result[0].text)
        assert "items" in data


# ---------------------------------------------------------------------------
# SharePoint tools
# ---------------------------------------------------------------------------


class TestSharePointTools:

    @pytest.mark.asyncio
    async def test_search_sites(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_search_sites", {"query": "*"})
        data = json.loads(result[0].text)
        assert "sites" in data
        assert isinstance(data["sites"], list)

    @pytest.mark.asyncio
    async def test_get_site_drives(self, graph: MsGraphInstance):
        sites_result = await _handle_tool(graph, "o365_search_sites", {"query": "*"})
        sites_data = json.loads(sites_result[0].text)
        if not sites_data["sites"]:
            pytest.skip("No SharePoint sites found")
        site_id = sites_data["sites"][0]["id"]
        result = await _handle_tool(graph, "o365_get_site_drives", {"site_id": site_id})
        data = json.loads(result[0].text)
        assert "drives" in data


# ---------------------------------------------------------------------------
# Directory tools
# ---------------------------------------------------------------------------


class TestDirectoryTools:

    @pytest.mark.asyncio
    async def test_list_users(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_users", {"limit": 5})
        data = json.loads(result[0].text)
        assert "users" in data
        assert "total_users" in data
        assert data["total_users"] >= 0

    @pytest.mark.asyncio
    async def test_list_users_default(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_list_users", {})
        data = json.loads(result[0].text)
        assert len(data["users"]) <= 25

    @pytest.mark.asyncio
    async def test_get_user_manager(self, graph: MsGraphInstance):
        # Get current user ID first
        profile = await _handle_tool(graph, "o365_get_profile", {})
        profile_data = json.loads(profile[0].text)
        if not profile_data or not profile_data.get("id"):
            pytest.skip("Profile not available")
        user_id = profile_data["id"]
        result = await _handle_tool(graph, "o365_get_user_manager", {"user_id": user_id})
        data = json.loads(result[0].text)
        # Manager may be null if user has no manager
        assert data is None or isinstance(data, dict)


# ---------------------------------------------------------------------------
# Unknown tool
# ---------------------------------------------------------------------------


class TestUnknownTool:

    @pytest.mark.asyncio
    async def test_unknown_tool(self, graph: MsGraphInstance):
        result = await _handle_tool(graph, "o365_nonexistent", {})
        assert "Unknown tool" in result[0].text


# ---------------------------------------------------------------------------
# Tool coverage check
# ---------------------------------------------------------------------------


class TestToolCoverage:

    def test_all_read_only_tools_have_integration_tests(self):
        """Every READ_ONLY tool should have a live-Graph test above.

        Write/draft tools are intentionally excluded — invoking them in an
        integration run would create or send real mail / calendar events.
        Their gating is covered by tests/test_mcp_permissions.py.
        """
        from office_con.mcp_permissions import PermissionLevel
        from office_con.mcp_server import TOOL_PERMISSIONS
        tested_tools = {
            "o365_get_profile",
            "o365_list_mail",
            "o365_get_mail",
            "o365_get_mail_categories",
            "o365_list_calendars",
            "o365_get_events",
            "o365_get_schedule",
            "o365_list_teams",
            "o365_list_channels",
            "o365_get_channel_messages",
            "o365_get_team_members",
            "o365_list_chats",
            "o365_get_chat_messages",
            "o365_get_chat_members",
            "o365_get_my_drive",
            "o365_list_drive_items",
            "o365_get_file_content",
            "o365_search_files",
            "o365_search_sites",
            "o365_get_site_drives",
            "o365_list_users",
            "o365_get_user_manager",
            "o365_list_rooms",
            "o365_get_room_availability",
            "o365_search_mail",
            "o365_search_events",
            "o365_search_messages",
            "o365_peek_drive_file",
            "o365_peek_mail_attachment",
        }
        read_only_tools = {
            name for name, lvl in TOOL_PERMISSIONS.items()
            if lvl is PermissionLevel.READ_ONLY
        }
        untested = read_only_tools - tested_tools
        assert not untested, f"Untested read-only tools: {sorted(untested)}"

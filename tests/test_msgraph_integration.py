"""Read-only integration tests for office-mcp MS Graph handlers.

Requires a valid token file at tests/msgraph_test_token.json.
Run with: poetry run pytest dependencies/office-mcp/tests/test_msgraph_integration.py -v
"""

import json
import os
from pathlib import Path
from datetime import datetime, timedelta

import pytest
import pytest_asyncio

from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.msgraph.profile_handler import ProfileHandler, UserProfile
from office_con.msgraph.mail_handler import OfficeMailHandler, OfficeMail, OfficeMailList
from office_con.msgraph.calendar_handler import CalendarHandler, CalendarEvent, CalendarEventList
from office_con.msgraph.directory_handler import DirectoryHandler, DirectoryUser, DirectoryUserList
from office_con.msgraph.teams_handler import TeamsHandler, Team, TeamList, Channel, ChannelList, ChannelMessage, ChannelMessageList, TeamMember, TeamMemberList
from office_con.msgraph.chat_handler import ChatHandler, Chat, ChatList, ChatMessage, ChatMessageList, ChatMember, ChatMemberList
from office_con.msgraph.files_handler import FilesHandler, Drive, DriveList, DriveItem, DriveItemList, SharePointSite, SharePointSiteList

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
    """Create an MsGraphInstance seeded with the test token."""
    data = _load_token()
    inst = MsGraphInstance(
        scopes=None,
        app=data.get("app", "test"),
        session_id=data.get("session_id"),
        can_refresh=True,
        endpoint="https://graph.microsoft.com/v1.0/",
    )
    inst.email = data.get("email")
    original_token = data["access_token"]
    # Seed tokens into memory cache directly (no Redis needed for tests)
    inst.cache_dict["access_token"] = original_token
    inst.cache_dict["refresh_token"] = data.get("refresh_token")
    # Try refreshing to get a fresh token (needs O365_CLIENT_SECRET env var)
    try:
        refreshed = await inst.refresh_token_async()
        if refreshed:
            inst.cache_dict["access_token"] = refreshed
    except Exception:
        pass
    # Ensure original token is preserved if refresh cleared it
    if not inst.cache_dict.get("access_token"):
        inst.cache_dict["access_token"] = original_token
    return inst


# ---------------------------------------------------------------------------
# Token & instance basics
# ---------------------------------------------------------------------------

class TestTokenBasics:

    @pytest.mark.asyncio
    async def test_access_token_in_cache(self, graph: MsGraphInstance):
        token = inst_token(graph)
        assert token is not None
        assert len(token) > 100

    @pytest.mark.asyncio
    async def test_is_token_still_valid(self, graph: MsGraphInstance):
        token = inst_token(graph)
        # Token may be expired if not refreshed, but the method should not crash
        result = graph.is_token_still_valid(token)
        assert isinstance(result, bool)

    @pytest.mark.asyncio
    async def test_time_until_expiration(self, graph: MsGraphInstance):
        token = inst_token(graph)
        ttl = graph.time_until_token_expiration(token)
        assert isinstance(ttl, (int, float))

    @pytest.mark.asyncio
    async def test_run_async_get(self, graph: MsGraphInstance):
        """run_async should return a response wrapper for a simple GET."""
        resp = await graph.run_async(url=f"{graph.msg_endpoint}me")
        assert resp is not None
        assert hasattr(resp, "status_code")
        assert hasattr(resp, "json")

    @pytest.mark.asyncio
    async def test_run_async_invalid_url(self, graph: MsGraphInstance):
        """run_async should handle a 404 gracefully."""
        resp = await graph.run_async(url=f"{graph.msg_endpoint}me/nonexistent_endpoint_xyz")
        assert resp is not None
        assert resp.status_code >= 400


# ---------------------------------------------------------------------------
# Profile
# ---------------------------------------------------------------------------

class TestProfile:

    @pytest.mark.asyncio
    async def test_get_profile_async(self, graph: MsGraphInstance):
        profile_handler = await graph.get_profile_async()
        assert isinstance(profile_handler, ProfileHandler)
        assert profile_handler.me is not None

    @pytest.mark.asyncio
    async def test_profile_has_user_id(self, graph: MsGraphInstance):
        handler = await graph.get_profile_async()
        me = handler.me
        assert me.id, "Profile should have a user ID"
        assert len(me.id) > 10

    @pytest.mark.asyncio
    async def test_profile_has_display_name(self, graph: MsGraphInstance):
        handler = await graph.get_profile_async()
        me = handler.me
        assert me.display_name, "Profile should have a display name"

    @pytest.mark.asyncio
    async def test_profile_has_email(self, graph: MsGraphInstance):
        handler = await graph.get_profile_async()
        me = handler.me
        assert me.mail or me.user_principal_name, "Profile should have mail or UPN"

    @pytest.mark.asyncio
    async def test_profile_populates_instance(self, graph: MsGraphInstance):
        """get_profile_async should populate graph.email, graph.user_id, etc."""
        await graph.get_profile_async()
        assert graph.email is not None
        assert graph.user_id is not None
        assert graph.given_name is not None
        assert graph.full_name is not None

    @pytest.mark.asyncio
    async def test_me_async_returns_user_profile(self, graph: MsGraphInstance):
        handler = ProfileHandler(graph)
        me = await handler.me_async()
        assert isinstance(me, UserProfile)
        assert me.id

    @pytest.mark.asyncio
    async def test_me_property_cached(self, graph: MsGraphInstance):
        handler = await graph.get_profile_async()
        me1 = handler.me
        me2 = handler.me
        assert me1 is me2, "me property should return cached instance"


# ---------------------------------------------------------------------------
# Mail (read-only)
# ---------------------------------------------------------------------------

class TestMail:

    @pytest.mark.asyncio
    async def test_email_index(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=5)
        assert isinstance(result, OfficeMailList)
        assert isinstance(result.elements, list)
        assert result.total_mails >= 0

    @pytest.mark.asyncio
    async def test_email_index_has_mails(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=5)
        if result.total_mails > 0:
            assert len(result.elements) > 0
            first = result.elements[0]
            assert isinstance(first, OfficeMail)
            assert first.email_id

    @pytest.mark.asyncio
    async def test_email_fields(self, graph: MsGraphInstance):
        mail = graph.get_mail()
        result = await mail.email_index_async(limit=1)
        if result.elements:
            m = result.elements[0]
            assert m.email_id is not None
            assert m.email_type is not None
            assert m.local_timestamp is not None
            assert isinstance(m.is_read, bool)
            assert isinstance(m.has_attachments, bool)
            assert isinstance(m.categories, list)

    @pytest.mark.asyncio
    async def test_get_single_mail(self, graph: MsGraphInstance):
        handler = graph.get_mail()
        index = await handler.email_index_async(limit=1)
        if not index.elements:
            pytest.skip("No mails in inbox")
        first = index.elements[0]
        full_mail = await handler.get_mail_async(email_url=first.email_url)
        assert full_mail is not None
        assert full_mail.email_id == first.email_id
        # Full mail should have body
        assert full_mail.body is not None

    @pytest.mark.asyncio
    async def test_get_mail_without_attachments(self, graph: MsGraphInstance):
        handler = graph.get_mail()
        index = await handler.email_index_async(limit=1)
        if not index.elements:
            pytest.skip("No mails in inbox")
        first = index.elements[0]
        mail = await handler.get_mail_async(email_url=first.email_url, attachments=False)
        assert mail is not None

    @pytest.mark.asyncio
    async def test_get_categories(self, graph: MsGraphInstance):
        handler = graph.get_mail()
        categories = await handler.get_categories_async()
        assert isinstance(categories, list)
        # User should have at least default categories
        if categories:
            cat = categories[0]
            assert cat.id
            assert cat.name

    @pytest.mark.asyncio
    async def test_get_user_profile_via_mail(self, graph: MsGraphInstance):
        handler = graph.get_mail()
        profile = await handler.get_user_profile_async()
        assert profile is not None
        assert "id" in profile
        assert "displayName" in profile

    @pytest.mark.asyncio
    async def test_email_index_pagination(self, graph: MsGraphInstance):
        handler = graph.get_mail()
        page1 = await handler.email_index_async(limit=2, skip=0)
        page2 = await handler.email_index_async(limit=2, skip=2)
        if page1.total_mails > 2:
            assert len(page1.elements) <= 2
            if page2.elements:
                assert page1.elements[0].email_id != page2.elements[0].email_id


# ---------------------------------------------------------------------------
# Calendar (read-only)
# ---------------------------------------------------------------------------

class TestCalendar:

    @pytest.mark.asyncio
    async def test_get_calendars(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        calendars = await cal.get_calendars_async()
        assert isinstance(calendars, list)
        assert len(calendars) > 0, "User should have at least one calendar"

    @pytest.mark.asyncio
    async def test_get_default_calendar_id(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        cal_id = await cal.get_default_calendar_id_async()
        assert cal_id is not None
        assert len(cal_id) > 10

    @pytest.mark.asyncio
    async def test_get_events_this_month(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        events = await cal.get_events_this_month_async(limit=10)
        assert isinstance(events, CalendarEventList)
        assert isinstance(events.events, list)

    @pytest.mark.asyncio
    async def test_get_events_custom_range(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        now = datetime.now()
        start = now - timedelta(days=7)
        end = now + timedelta(days=7)
        events = await cal.get_events_async(start_date=start, end_date=end, limit=10)
        assert isinstance(events, CalendarEventList)

    @pytest.mark.asyncio
    async def test_calendar_event_fields(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        now = datetime.now()
        events = await cal.get_events_async(
            start_date=now - timedelta(days=30),
            end_date=now + timedelta(days=30),
            limit=5,
        )
        if events.events:
            e = events.events[0]
            assert isinstance(e, CalendarEvent)
            assert e.id
            assert e.subject
            assert isinstance(e.start_time, datetime)
            assert isinstance(e.end_time, datetime)
            assert isinstance(e.is_all_day, bool)

    @pytest.mark.asyncio
    async def test_get_user_timezone(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        tz = await cal.get_user_timezone_async()
        assert isinstance(tz, str)
        assert len(tz) > 0

    @pytest.mark.asyncio
    async def test_get_schedule(self, graph: MsGraphInstance):
        cal = graph.get_calendar()
        now = datetime.now()
        start = now.replace(hour=8, minute=0, second=0, microsecond=0)
        end = start + timedelta(hours=8)
        schedule = await cal.get_schedule_async(
            emails=[graph.email],
            start=start,
            end=end,
        )
        assert isinstance(schedule, list)


# ---------------------------------------------------------------------------
# Directory (read-only)
# ---------------------------------------------------------------------------

class TestDirectory:

    @pytest.mark.asyncio
    async def test_get_users(self, graph: MsGraphInstance):
        directory = graph.get_directory()
        result = await directory.get_users_async(limit=5)
        assert isinstance(result, DirectoryUserList)
        assert result.total_users > 0
        assert len(result.users) > 0

    @pytest.mark.asyncio
    async def test_directory_user_fields(self, graph: MsGraphInstance):
        directory = graph.get_directory()
        result = await directory.get_users_async(limit=3)
        if result.users:
            u = result.users[0]
            assert isinstance(u, DirectoryUser)
            assert u.id
            assert u.display_name or u.email

    @pytest.mark.asyncio
    async def test_get_user_photo(self, graph: MsGraphInstance):
        """Fetch a user photo — may return None if no photo is set."""
        await graph.get_profile_async()
        directory = graph.get_directory()
        photo = await directory.get_user_photo_async(graph.user_id)
        # Photo can be None if user has no photo set
        if photo is not None:
            assert isinstance(photo, bytes)
            assert len(photo) > 0

    @pytest.mark.asyncio
    async def test_get_user_manager(self, graph: MsGraphInstance):
        await graph.get_profile_async()
        directory = graph.get_directory()
        manager = await directory.get_user_manager_async(graph.user_id)
        # Manager may be None if user has no manager
        if manager is not None:
            assert isinstance(manager, dict)
            assert "id" in manager or "displayName" in manager


# ---------------------------------------------------------------------------
# Teams (read-only)
# ---------------------------------------------------------------------------

class TestTeams:

    @pytest.mark.asyncio
    async def test_get_joined_teams(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        result = await handler.get_joined_teams_async()
        assert isinstance(result, TeamList)
        assert isinstance(result.teams, list)
        assert result.total_teams >= 0

    @pytest.mark.asyncio
    async def test_team_fields(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        result = await handler.get_joined_teams_async()
        if result.teams:
            t = result.teams[0]
            assert isinstance(t, Team)
            assert t.id
            assert t.display_name

    @pytest.mark.asyncio
    async def test_get_channels(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        teams = await handler.get_joined_teams_async()
        if not teams.teams:
            pytest.skip("No teams joined")
        result = await handler.get_channels_async(teams.teams[0].id)
        assert isinstance(result, ChannelList)
        assert result.total_channels > 0  # every team has at least General

    @pytest.mark.asyncio
    async def test_channel_fields(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        teams = await handler.get_joined_teams_async()
        if not teams.teams:
            pytest.skip("No teams joined")
        channels = await handler.get_channels_async(teams.teams[0].id)
        if channels.channels:
            c = channels.channels[0]
            assert isinstance(c, Channel)
            assert c.id
            assert c.display_name

    @pytest.mark.asyncio
    async def test_get_channel_messages(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        teams = await handler.get_joined_teams_async()
        if not teams.teams:
            pytest.skip("No teams joined")
        channels = await handler.get_channels_async(teams.teams[0].id)
        if not channels.channels:
            pytest.skip("No channels")
        result = await handler.get_channel_messages_async(
            teams.teams[0].id, channels.channels[0].id, limit=5
        )
        assert isinstance(result, ChannelMessageList)
        assert isinstance(result.messages, list)

    @pytest.mark.asyncio
    async def test_channel_message_fields(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        teams = await handler.get_joined_teams_async()
        if not teams.teams:
            pytest.skip("No teams joined")
        channels = await handler.get_channels_async(teams.teams[0].id)
        if not channels.channels:
            pytest.skip("No channels")
        msgs = await handler.get_channel_messages_async(
            teams.teams[0].id, channels.channels[0].id, limit=3
        )
        if msgs.messages:
            m = msgs.messages[0]
            assert isinstance(m, ChannelMessage)
            assert m.id

    @pytest.mark.asyncio
    async def test_get_team_members(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        teams = await handler.get_joined_teams_async()
        if not teams.teams:
            pytest.skip("No teams joined")
        result = await handler.get_team_members_async(teams.teams[0].id)
        assert isinstance(result, TeamMemberList)
        assert result.total_members >= 0

    @pytest.mark.asyncio
    async def test_team_member_fields(self, graph: MsGraphInstance):
        handler = graph.get_teams()
        teams = await handler.get_joined_teams_async()
        if not teams.teams:
            pytest.skip("No teams joined")
        members = await handler.get_team_members_async(teams.teams[0].id)
        if members.members:
            m = members.members[0]
            assert isinstance(m, TeamMember)
            assert m.id
            assert m.display_name


# ---------------------------------------------------------------------------
# Chats (read-only)
# ---------------------------------------------------------------------------

class TestChats:

    @pytest.mark.asyncio
    async def test_get_chats(self, graph: MsGraphInstance):
        handler = graph.get_chat()
        result = await handler.get_chats_async(limit=10)
        assert isinstance(result, ChatList)
        assert isinstance(result.chats, list)
        assert result.total_chats >= 0

    @pytest.mark.asyncio
    async def test_chat_fields(self, graph: MsGraphInstance):
        handler = graph.get_chat()
        result = await handler.get_chats_async(limit=5)
        if result.chats:
            c = result.chats[0]
            assert isinstance(c, Chat)
            assert c.id
            assert c.chat_type in ("oneOnOne", "group", "meeting", None)

    @pytest.mark.asyncio
    async def test_get_chat_messages(self, graph: MsGraphInstance):
        handler = graph.get_chat()
        chats = await handler.get_chats_async(limit=5)
        if not chats.chats:
            pytest.skip("No chats available")
        result = await handler.get_chat_messages_async(chats.chats[0].id, limit=5)
        assert isinstance(result, ChatMessageList)
        assert isinstance(result.messages, list)

    @pytest.mark.asyncio
    async def test_chat_message_fields(self, graph: MsGraphInstance):
        handler = graph.get_chat()
        chats = await handler.get_chats_async(limit=5)
        if not chats.chats:
            pytest.skip("No chats available")
        msgs = await handler.get_chat_messages_async(chats.chats[0].id, limit=3)
        if msgs.messages:
            m = msgs.messages[0]
            assert isinstance(m, ChatMessage)
            assert m.id

    @pytest.mark.asyncio
    async def test_get_chat_members(self, graph: MsGraphInstance):
        handler = graph.get_chat()
        chats = await handler.get_chats_async(limit=5)
        if not chats.chats:
            pytest.skip("No chats available")
        result = await handler.get_chat_members_async(chats.chats[0].id)
        assert isinstance(result, ChatMemberList)
        assert result.total_members >= 0

    @pytest.mark.asyncio
    async def test_chat_member_fields(self, graph: MsGraphInstance):
        handler = graph.get_chat()
        chats = await handler.get_chats_async(limit=5)
        if not chats.chats:
            pytest.skip("No chats available")
        members = await handler.get_chat_members_async(chats.chats[0].id)
        if members.members:
            m = members.members[0]
            assert isinstance(m, ChatMember)
            assert m.id


# ---------------------------------------------------------------------------
# Files & OneDrive (read-only)
# ---------------------------------------------------------------------------

class TestFiles:

    @pytest.mark.asyncio
    async def test_get_my_drive(self, graph: MsGraphInstance):
        handler = graph.get_files()
        drive = await handler.get_my_drive_async()
        assert drive is not None
        assert isinstance(drive, Drive)
        assert drive.id
        assert drive.drive_type

    @pytest.mark.asyncio
    async def test_get_my_drives(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.get_my_drives_async()
        assert isinstance(result, DriveList)
        assert result.total_drives >= 0

    @pytest.mark.asyncio
    async def test_get_root_items(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.get_root_items_async(limit=10)
        assert isinstance(result, DriveItemList)
        assert isinstance(result.items, list)

    @pytest.mark.asyncio
    async def test_drive_item_fields(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.get_root_items_async(limit=5)
        if result.items:
            item = result.items[0]
            assert isinstance(item, DriveItem)
            assert item.id
            assert item.name
            assert isinstance(item.is_folder, bool)

    @pytest.mark.asyncio
    async def test_get_folder_items(self, graph: MsGraphInstance):
        handler = graph.get_files()
        root = await handler.get_root_items_async(limit=20)
        folders = [i for i in root.items if i.is_folder]
        if not folders:
            pytest.skip("No folders in OneDrive root")
        result = await handler.get_folder_items_async(folders[0].id, limit=10)
        assert isinstance(result, DriveItemList)

    @pytest.mark.asyncio
    async def test_get_item(self, graph: MsGraphInstance):
        handler = graph.get_files()
        root = await handler.get_root_items_async(limit=1)
        if not root.items:
            pytest.skip("No items in OneDrive root")
        item = await handler.get_item_async(root.items[0].id)
        assert item is not None
        assert item.id == root.items[0].id

    @pytest.mark.asyncio
    async def test_search_items(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.search_items_async("test", limit=5)
        assert isinstance(result, DriveItemList)
        # Search may return empty, that's fine

    @pytest.mark.asyncio
    async def test_get_file_content(self, graph: MsGraphInstance):
        """Download a small file's content."""
        handler = graph.get_files()
        root = await handler.get_root_items_async(limit=20)
        files = [i for i in root.items if not i.is_folder and (i.size or 0) < 1_000_000]
        if not files:
            pytest.skip("No small files in OneDrive root")
        content = await handler.get_file_content_async(files[0].id)
        assert content is not None
        assert isinstance(content, bytes)
        assert len(content) > 0


# ---------------------------------------------------------------------------
# SharePoint (read-only)
# ---------------------------------------------------------------------------

class TestSharePoint:

    @pytest.mark.asyncio
    async def test_search_sites(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.search_sites_async("*")
        assert isinstance(result, SharePointSiteList)
        assert isinstance(result.sites, list)

    @pytest.mark.asyncio
    async def test_site_fields(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.search_sites_async("*")
        if result.sites:
            s = result.sites[0]
            assert isinstance(s, SharePointSite)
            assert s.id
            assert s.web_url

    @pytest.mark.asyncio
    async def test_get_followed_sites(self, graph: MsGraphInstance):
        handler = graph.get_files()
        result = await handler.get_followed_sites_async()
        assert isinstance(result, SharePointSiteList)

    @pytest.mark.asyncio
    async def test_get_site_drives(self, graph: MsGraphInstance):
        handler = graph.get_files()
        sites = await handler.search_sites_async("*")
        if not sites.sites:
            pytest.skip("No SharePoint sites found")
        result = await handler.get_site_drives_async(sites.sites[0].id)
        assert isinstance(result, DriveList)


# ---------------------------------------------------------------------------
# WebUserInstance core methods
# ---------------------------------------------------------------------------

class TestWebUserInstance:

    @pytest.mark.asyncio
    async def test_identifier(self, graph: MsGraphInstance):
        await graph.get_profile_async()
        ident = graph.identifier
        assert ident is not None
        assert "@" in ident  # should be email-based

    @pytest.mark.asyncio
    async def test_get_access_token_async_from_cache(self, graph: MsGraphInstance):
        """get_access_token_async should return token from memory cache."""
        token = await graph.get_access_token_async()
        assert token is not None
        assert len(token) > 100

    @pytest.mark.asyncio
    async def test_features(self, graph: MsGraphInstance):
        assert graph.FEATURE_MAIL in graph.features
        assert graph.FEATURE_CALENDAR in graph.features
        assert graph.FEATURE_PROFILE in graph.features


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def inst_token(graph: MsGraphInstance) -> str:
    return graph.cache_dict.get("access_token")

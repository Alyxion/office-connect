"""Integration tests — run against real MS Graph with a token file.

Skipped automatically when no test_config.json or token file is present.
Configure by creating ``tests/test_config.json``::

    {
        "token_file": "~/Downloads/token_export.json",
        "expected_rooms": ["Metzingen", "Chicago"],
        "expected_teams": ["My Team"],
        "expected_presence_users": ["colleague@example.com"]
    }
"""

from __future__ import annotations

import pytest

from office_con.testing.mock_config import get_test_config, get_token_data


def _make_graph():
    """Create a real MsGraphInstance from token file."""
    from office_con.msgraph.ms_graph_handler import MsGraphInstance

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


_cfg = get_test_config()
_has_token = get_token_data() is not None
_skip = pytest.mark.skipif(not _has_token, reason="No token file — set tests/test_config.json")


# ── Places ──────────────────────────────────────────────────────

@_skip
class TestPlacesIntegration:

    @pytest.mark.asyncio
    async def test_get_rooms(self):
        from office_con.msgraph.places_handler import PlacesHandler
        graph = _make_graph()
        handler = PlacesHandler(graph)
        rooms = await handler.get_rooms_async()
        assert isinstance(rooms, list)
        print(f"\n  Found {len(rooms)} rooms:")
        for r in rooms:
            print(f"    {r.get('displayName')} (capacity={r.get('capacity')}, building={r.get('building')})")

        expected = _cfg.get("expected_rooms", [])
        if expected:
            names = " ".join(r.get("displayName", "") for r in rooms).lower()
            for keyword in expected:
                assert keyword.lower() in names, f"Expected room containing '{keyword}' not found"

    @pytest.mark.asyncio
    async def test_get_room_lists(self):
        from office_con.msgraph.places_handler import PlacesHandler
        graph = _make_graph()
        handler = PlacesHandler(graph)
        lists = await handler.get_room_lists_async()
        assert isinstance(lists, list)
        print(f"\n  Found {len(lists)} room lists:")
        for rl in lists:
            print(f"    {rl.get('displayName')} ({rl.get('emailAddress')})")


# ── Presence ────────────────────────────────────────────────────

@_skip
class TestPresenceIntegration:

    @pytest.mark.asyncio
    async def test_my_presence(self):
        from office_con.msgraph.presence_handler import PresenceHandler
        graph = _make_graph()
        handler = PresenceHandler(graph)
        result = await handler.get_my_presence_async()
        assert "availability" in result
        print(f"\n  My presence: {result.get('availability')} / {result.get('activity')}")

    @pytest.mark.asyncio
    async def test_batch_presence(self):
        users = _cfg.get("expected_presence_users", [])
        if not users:
            pytest.skip("No expected_presence_users in config")
        from office_con.msgraph.presence_handler import PresenceHandler
        graph = _make_graph()
        handler = PresenceHandler(graph)
        result = await handler.get_presences_async(users)
        assert isinstance(result, list)
        for p in result:
            print(f"  {p.get('id')}: {p.get('availability')}")


# ── Teams ───────────────────────────────────────────────────────

@_skip
class TestTeamsIntegration:

    @pytest.mark.asyncio
    async def test_joined_teams(self):
        from office_con.msgraph.teams_handler import TeamsHandler
        graph = _make_graph()
        handler = TeamsHandler(graph)
        result = await handler.get_joined_teams_async()
        # TeamsHandler returns a TeamList model with .teams list
        teams = result.teams if hasattr(result, "teams") else result
        assert len(teams) > 0
        print(f"\n  Found {len(teams)} teams:")
        for t in teams:
            name = t.display_name if hasattr(t, "display_name") else t.get("displayName", "?")
            print(f"    {name}")

        expected = _cfg.get("expected_teams", [])
        if expected:
            names = " ".join(
                t.display_name if hasattr(t, "display_name") else t.get("displayName", "")
                for t in teams
            ).lower()
            for keyword in expected:
                assert keyword.lower() in names, f"Expected team containing '{keyword}' not found"

    @pytest.mark.asyncio
    async def test_team_channels(self):
        from office_con.msgraph.teams_handler import TeamsHandler
        graph = _make_graph()
        handler = TeamsHandler(graph)
        result = await handler.get_joined_teams_async()
        teams = result.teams if hasattr(result, "teams") else result
        if not teams:
            pytest.skip("No teams")
        first = teams[0]
        team_id = first.id if hasattr(first, "id") else first.get("id")
        team_name = first.display_name if hasattr(first, "display_name") else first.get("displayName", "?")
        channels = await handler.get_channels_async(team_id)
        channels_list = channels.channels if hasattr(channels, "channels") else channels
        print(f"\n  Team '{team_name}' has {len(channels_list)} channels:")
        for c in channels_list:
            name = c.display_name if hasattr(c, "display_name") else c.get("displayName", "?")
            print(f"    #{name}")


# ── Tasks ───────────────────────────────────────────────────────

@_skip
class TestTasksIntegration:

    @pytest.mark.asyncio
    async def test_task_lists(self):
        from office_con.msgraph.tasks_handler import TasksHandler
        graph = _make_graph()
        handler = TasksHandler(graph)
        lists = await handler.get_task_lists_async()
        assert isinstance(lists, list)
        print(f"\n  Found {len(lists)} task lists:")
        for tl in lists:
            print(f"    {tl.get('displayName')} (id={tl.get('id', '')[:12]}...)")

    @pytest.mark.asyncio
    async def test_tasks_in_list(self):
        from office_con.msgraph.tasks_handler import TasksHandler
        graph = _make_graph()
        handler = TasksHandler(graph)
        lists = await handler.get_task_lists_async()
        if not lists:
            pytest.skip("No task lists")
        list_id = lists[0].get("id")
        tasks = await handler.get_tasks_async(list_id)
        assert isinstance(tasks, list)
        print(f"\n  List '{lists[0].get('displayName')}' has {len(tasks)} tasks:")
        for t in tasks[:5]:
            print(f"    [{t.get('status')}] {t.get('title')}")


# ── People ──────────────────────────────────────────────────────

@_skip
class TestPeopleIntegration:

    @pytest.mark.asyncio
    async def test_relevant_people(self):
        from office_con.msgraph.people_handler import PeopleHandler
        graph = _make_graph()
        handler = PeopleHandler(graph)
        people = await handler.get_relevant_people_async(limit=10)
        assert isinstance(people, list)
        print(f"\n  Top {len(people)} relevant people:")
        for p in people:
            emails = [e.get("address") for e in p.get("emailAddresses", [])]
            print(f"    {p.get('displayName')} ({', '.join(emails)})")

    @pytest.mark.asyncio
    async def test_contacts(self):
        from office_con.msgraph.people_handler import PeopleHandler
        graph = _make_graph()
        handler = PeopleHandler(graph)
        contacts = await handler.get_contacts_async(limit=10)
        assert isinstance(contacts, list)
        print(f"\n  Found {len(contacts)} contacts")


# ── Mailbox Settings ────────────────────────────────────────────

@_skip
class TestMailboxSettingsIntegration:

    @pytest.mark.asyncio
    async def test_mailbox_settings(self):
        from office_con.msgraph.mailbox_settings_handler import MailboxSettingsHandler
        graph = _make_graph()
        handler = MailboxSettingsHandler(graph)
        settings = await handler.get_mailbox_settings_async()
        assert isinstance(settings, dict)
        print(f"\n  Timezone: {settings.get('timeZone')}")
        print(f"  Language: {settings.get('language', {}).get('displayName')}")
        print(f"  Date format: {settings.get('dateFormat')}")

    @pytest.mark.asyncio
    async def test_working_hours(self):
        from office_con.msgraph.mailbox_settings_handler import MailboxSettingsHandler
        graph = _make_graph()
        handler = MailboxSettingsHandler(graph)
        wh = await handler.get_working_hours_async()
        assert isinstance(wh, dict)
        days = wh.get("daysOfWeek", [])
        print(f"\n  Working days: {', '.join(days)}")
        print(f"  Start: {wh.get('startTime')}, End: {wh.get('endTime')}")


# ── Online Meetings ─────────────────────────────────────────────

@_skip
class TestOnlineMeetingsIntegration:

    @pytest.mark.asyncio
    async def test_meetings(self):
        from office_con.msgraph.online_meetings_handler import OnlineMeetingsHandler
        graph = _make_graph()
        handler = OnlineMeetingsHandler(graph)
        meetings = await handler.get_meetings_async(limit=5)
        assert isinstance(meetings, list)
        print(f"\n  Found {len(meetings)} online meetings:")
        for m in meetings:
            print(f"    {m.get('subject', '(no subject)')}")

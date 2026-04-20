"""Unit tests for new MS Graph handlers using the mock transport.

Tests presence, tasks, people, places, mailbox settings, and online meetings
handlers against MockGraphTransport without any real HTTP calls.

Run with: poetry run pytest dependencies/office-mcp/tests/test_new_handlers.py -v
"""

from __future__ import annotations

import pytest

from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.msgraph.presence_handler import PresenceHandler
from office_con.msgraph.tasks_handler import TasksHandler
from office_con.msgraph.people_handler import PeopleHandler
from office_con.msgraph.places_handler import PlacesHandler
from office_con.msgraph.mailbox_settings_handler import MailboxSettingsHandler
from office_con.msgraph.online_meetings_handler import OnlineMeetingsHandler
from office_con.testing.fixtures import default_mock_profile
from office_con.testing.mock_tokens import make_mock_access_token


@pytest.fixture
def graph():
    profile = default_mock_profile()
    g = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
    g.enable_mock(profile)
    # Seed a mock access token so handlers pass the token gate
    g.cache_dict["access_token"] = make_mock_access_token(profile.email, profile.user_id)
    return g


# ---------------------------------------------------------------------------
# Presence
# ---------------------------------------------------------------------------

class TestPresenceHandler:

    @pytest.mark.asyncio
    async def test_get_my_presence(self, graph: MsGraphInstance):
        handler = PresenceHandler(graph)
        result = await handler.get_my_presence_async()
        assert isinstance(result, dict)
        assert result["availability"] == "Available"
        assert result["activity"] == "Available"
        assert "id" in result

    @pytest.mark.asyncio
    async def test_get_user_presence(self, graph: MsGraphInstance):
        handler = PresenceHandler(graph)
        user_id = "some-user-id-123"
        result = await handler.get_user_presence_async(user_id)
        assert isinstance(result, dict)
        assert result["availability"] == "Available"
        assert result["activity"] == "Available"
        assert result["id"] == user_id

    @pytest.mark.asyncio
    async def test_get_presences(self, graph: MsGraphInstance):
        handler = PresenceHandler(graph)
        ids = ["user-a", "user-b", "user-c"]
        result = await handler.get_presences_async(ids)
        assert isinstance(result, list)
        assert len(result) == 3
        for i, p in enumerate(result):
            assert p["id"] == ids[i]
            assert p["availability"] == "Available"
            assert p["activity"] == "Available"

    @pytest.mark.asyncio
    async def test_get_presences_empty(self, graph: MsGraphInstance):
        handler = PresenceHandler(graph)
        result = await handler.get_presences_async([])
        assert isinstance(result, list)
        assert len(result) == 0


# ---------------------------------------------------------------------------
# Tasks (To Do)
# ---------------------------------------------------------------------------

class TestTasksHandler:

    @pytest.mark.asyncio
    async def test_get_task_lists(self, graph: MsGraphInstance):
        handler = TasksHandler(graph)
        result = await handler.get_task_lists_async()
        assert isinstance(result, list)
        assert len(result) == 3
        names = [tl["displayName"] for tl in result]
        assert "Tasks" in names
        assert "Einkaufsliste" in names
        assert "Arbeitsprojekte" in names
        for tl in result:
            assert "id" in tl
            assert "displayName" in tl

    @pytest.mark.asyncio
    async def test_get_tasks_default_list(self, graph: MsGraphInstance):
        handler = TasksHandler(graph)
        result = await handler.get_tasks_async("tasklist-tasks")
        assert isinstance(result, list)
        assert len(result) == 3
        for task in result:
            assert "id" in task
            assert "title" in task
            assert "status" in task
            assert "importance" in task
            assert task["status"] in ("notStarted", "inProgress", "completed")

    @pytest.mark.asyncio
    async def test_get_tasks_work_list(self, graph: MsGraphInstance):
        handler = TasksHandler(graph)
        result = await handler.get_tasks_async("tasklist-work")
        assert isinstance(result, list)
        assert len(result) == 5
        # Check that at least one has a due date
        has_due = any(t.get("dueDateTime") is not None for t in result)
        assert has_due

    @pytest.mark.asyncio
    async def test_get_tasks_nonexistent_list(self, graph: MsGraphInstance):
        handler = TasksHandler(graph)
        result = await handler.get_tasks_async("nonexistent-list")
        assert isinstance(result, list)
        assert len(result) == 0

    @pytest.mark.asyncio
    async def test_get_task_single(self, graph: MsGraphInstance):
        handler = TasksHandler(graph)
        result = await handler.get_task_async("tasklist-tasks", "task-001")
        assert result is not None
        assert isinstance(result, dict)
        assert result["id"] == "task-001"
        assert result["title"] == "Quartalsbericht vorbereiten"
        assert result["importance"] == "high"

    @pytest.mark.asyncio
    async def test_get_task_not_found(self, graph: MsGraphInstance):
        handler = TasksHandler(graph)
        result = await handler.get_task_async("tasklist-tasks", "nonexistent-task")
        # Handler returns None on 404
        assert result is None


# ---------------------------------------------------------------------------
# People
# ---------------------------------------------------------------------------

class TestPeopleHandler:

    @pytest.mark.asyncio
    async def test_get_relevant_people(self, graph: MsGraphInstance):
        handler = PeopleHandler(graph)
        result = await handler.get_relevant_people_async()
        assert isinstance(result, list)
        assert len(result) == 5
        for person in result:
            assert "id" in person
            assert "displayName" in person
            assert "emailAddresses" in person
            assert len(person["emailAddresses"]) > 0
            assert "address" in person["emailAddresses"][0]
        # Check known names
        names = [p["displayName"] for p in result]
        assert "Max Mustermann" in names
        assert "Anna Schmidt" in names

    @pytest.mark.asyncio
    async def test_search_people(self, graph: MsGraphInstance):
        handler = PeopleHandler(graph)
        # search_people_async hits me/people?$search=... which routes to _people_response
        result = await handler.search_people_async("Max")
        assert isinstance(result, list)
        assert len(result) > 0

    @pytest.mark.asyncio
    async def test_get_contacts(self, graph: MsGraphInstance):
        handler = PeopleHandler(graph)
        result = await handler.get_contacts_async()
        assert isinstance(result, list)
        assert len(result) == 3
        for contact in result:
            assert "id" in contact
            assert "givenName" in contact
            assert "surname" in contact
            assert "emailAddresses" in contact
            assert len(contact["emailAddresses"]) > 0
        # Check known contacts
        surnames = [c["surname"] for c in result]
        assert "Fischer" in surnames
        assert "Hoffmann" in surnames
        assert "Zimmermann" in surnames

    @pytest.mark.asyncio
    async def test_contacts_have_business_phones(self, graph: MsGraphInstance):
        handler = PeopleHandler(graph)
        result = await handler.get_contacts_async()
        for contact in result:
            assert "businessPhones" in contact
            assert len(contact["businessPhones"]) > 0
            assert contact["businessPhones"][0].startswith("+49")


# ---------------------------------------------------------------------------
# Places
# ---------------------------------------------------------------------------

class TestPlacesHandler:

    @pytest.mark.asyncio
    async def test_get_rooms(self, graph: MsGraphInstance):
        handler = PlacesHandler(graph)
        result = await handler.get_rooms_async()
        assert isinstance(result, list)
        assert len(result) == 4
        for room in result:
            assert "id" in room
            assert "displayName" in room
            assert "emailAddress" in room
            assert "capacity" in room
            assert isinstance(room["capacity"], int)
            assert room["capacity"] > 0
        # Check known room names
        names = [r["displayName"] for r in result]
        assert "Raum Stuttgart" in names
        assert "Raum Heidelberg" in names
        assert "Raum Reutlingen" in names
        assert "Raum Tuebingen" in names

    @pytest.mark.asyncio
    async def test_rooms_have_building_info(self, graph: MsGraphInstance):
        handler = PlacesHandler(graph)
        result = await handler.get_rooms_async()
        buildings = {r["building"] for r in result}
        assert "Hauptgebaeude" in buildings
        assert "Neubau" in buildings

    @pytest.mark.asyncio
    async def test_get_room_lists(self, graph: MsGraphInstance):
        handler = PlacesHandler(graph)
        result = await handler.get_room_lists_async()
        assert isinstance(result, list)
        assert len(result) == 2
        for rl in result:
            assert "id" in rl
            assert "displayName" in rl
            assert "emailAddress" in rl
        names = [rl["displayName"] for rl in result]
        assert "Hauptgebaeude" in names
        assert "Neubau" in names


# ---------------------------------------------------------------------------
# Mailbox Settings
# ---------------------------------------------------------------------------

class TestMailboxSettingsHandler:

    @pytest.mark.asyncio
    async def test_get_mailbox_settings(self, graph: MsGraphInstance):
        handler = MailboxSettingsHandler(graph)
        result = await handler.get_mailbox_settings_async()
        assert isinstance(result, dict)
        assert result["timeZone"] == "Europe/Berlin"
        assert result["language"]["locale"] == "de-DE"
        assert result["dateFormat"] == "dd.MM.yyyy"
        assert result["timeFormat"] == "HH:mm"
        # Check working hours are included
        wh = result["workingHours"]
        assert "monday" in wh["daysOfWeek"]
        assert "friday" in wh["daysOfWeek"]
        assert "saturday" not in wh["daysOfWeek"]
        assert wh["startTime"] == "08:00:00.0000000"
        assert wh["endTime"] == "17:00:00.0000000"
        # Check auto-reply settings are included
        ars = result["automaticRepliesSetting"]
        assert ars["status"] == "disabled"

    @pytest.mark.asyncio
    async def test_get_automatic_replies(self, graph: MsGraphInstance):
        handler = MailboxSettingsHandler(graph)
        result = await handler.get_automatic_replies_async()
        assert isinstance(result, dict)
        assert result["status"] == "disabled"
        assert "externalAudience" in result
        assert "internalReplyMessage" in result
        assert "externalReplyMessage" in result
        assert "scheduledStartDateTime" in result
        assert "scheduledEndDateTime" in result

    @pytest.mark.asyncio
    async def test_get_working_hours(self, graph: MsGraphInstance):
        handler = MailboxSettingsHandler(graph)
        result = await handler.get_working_hours_async()
        assert isinstance(result, dict)
        assert "daysOfWeek" in result
        assert len(result["daysOfWeek"]) == 5
        assert result["startTime"] == "08:00:00.0000000"
        assert result["endTime"] == "17:00:00.0000000"
        assert result["timeZone"]["name"] == "Europe/Berlin"


# ---------------------------------------------------------------------------
# Online Meetings
# ---------------------------------------------------------------------------

class TestOnlineMeetingsHandler:

    @pytest.mark.asyncio
    async def test_get_meetings(self, graph: MsGraphInstance):
        handler = OnlineMeetingsHandler(graph)
        result = await handler.get_meetings_async()
        assert isinstance(result, list)
        assert len(result) == 2
        for meeting in result:
            assert "id" in meeting
            assert "subject" in meeting
            assert "startDateTime" in meeting
            assert "endDateTime" in meeting
            assert "joinWebUrl" in meeting
            assert meeting["joinWebUrl"].startswith("https://")

    @pytest.mark.asyncio
    async def test_get_meeting_subjects(self, graph: MsGraphInstance):
        handler = OnlineMeetingsHandler(graph)
        result = await handler.get_meetings_async()
        subjects = [m["subject"] for m in result]
        assert "Sprint Planning KW15" in subjects
        assert "Projektstatus Besprechung" in subjects

    @pytest.mark.asyncio
    async def test_get_single_meeting(self, graph: MsGraphInstance):
        handler = OnlineMeetingsHandler(graph)
        result = await handler.get_meeting_async("meeting-001")
        assert result is not None
        assert isinstance(result, dict)
        assert result["id"] == "meeting-001"
        assert result["subject"] == "Sprint Planning KW15"
        assert "participants" in result
        assert "organizer" in result["participants"]

    @pytest.mark.asyncio
    async def test_get_meeting_not_found(self, graph: MsGraphInstance):
        handler = OnlineMeetingsHandler(graph)
        result = await handler.get_meeting_async("nonexistent-meeting-id")
        # Handler returns None on 404
        assert result is None

    @pytest.mark.asyncio
    async def test_meetings_have_chat_info(self, graph: MsGraphInstance):
        handler = OnlineMeetingsHandler(graph)
        result = await handler.get_meetings_async()
        for meeting in result:
            assert "chatInfo" in meeting
            assert "threadId" in meeting["chatInfo"]

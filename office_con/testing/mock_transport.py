"""URL-routing mock transport for MS Graph API calls.

Intercepts ``run_async()`` requests and returns synthetic
``AsyncResponseWrapper``-compatible responses based on URL patterns.
"""

from __future__ import annotations

import json as _json
import logging
import re
from typing import TYPE_CHECKING
from urllib.parse import urlparse, parse_qs

if TYPE_CHECKING:
    from office_con.testing.mock_data import MockUserProfile

logger = logging.getLogger(__name__)


class _MockResponse:
    """Minimal response object matching AsyncResponseWrapper interface."""

    def __init__(self, status_code: int, data: dict | bytes):
        self.status_code = status_code
        if isinstance(data, bytes):
            self.content = data
            self.text = None
        else:
            text = _json.dumps(data)
            self.content = text.encode()
            self.text = text
        self.headers = {"Content-Type": "application/json"}
        self.url = ""

    def json(self):
        return _json.loads(self.text)

    def raise_for_status(self):
        if self.status_code >= 400:
            raise Exception(f"HTTP Error: {self.status_code}")


def _make_response(status: int, data: dict | bytes) -> _MockResponse:
    return _MockResponse(status, data)


def _extract_path(url: str) -> str:
    """Extract the MS Graph path after /v1.0/ (e.g. 'me', 'me/calendars')."""
    parsed = urlparse(url)
    path = parsed.path
    # Strip /v1.0/ prefix
    v1_idx = path.find("/v1.0/")
    if v1_idx >= 0:
        path = path[v1_idx + 6:]
    # Also strip leading /
    return path.lstrip("/")


def _uid() -> str:
    import uuid
    return str(uuid.uuid4())


class MockGraphTransport:
    """Intercepts MS Graph HTTP calls and returns synthetic responses."""

    def __init__(self, profile: MockUserProfile):
        self._profile = profile

    async def handle_request(self, url: str, method: str, json_body: dict | None) -> _MockResponse:
        """Route URL to mock handler."""
        path = _extract_path(url)
        parsed = urlparse(url)
        qs = parse_qs(parsed.query)

        # ── /me ──────────────────────────────────────────────
        if path == "me":
            return self._profile_response()

        # ── /me/photo/$value ─────────────────────────────────
        if path == "me/photo/$value" or path.startswith("me/photo"):
            return self._photo_response()

        # ── /me/calendars/{id}/events (POST) ─────────────────
        if re.match(r"me/calendars/[^/]+/events$", path) and method == "POST":
            return self._create_event_response(json_body)

        # ── /me/calendars ────────────────────────────────────
        if path == "me/calendars":
            return self._calendars_response()

        # ── /me/calendars/{id}/calendarView ──────────────────
        if "calendarView" in path:
            return self._calendar_events_response(url)

        # ── /me/calendar/getSchedule ─────────────────────────
        if path == "me/calendar/getSchedule":
            return self._schedule_response(json_body)

        # ── /me/outlook/masterCategories ─────────────────────
        if path == "me/outlook/masterCategories":
            return self._categories_response()

        # ── /me/mailboxSettings (and sub-paths) ──────────────
        if path == "me/mailboxSettings/automaticRepliesSetting":
            return self._automatic_replies_response()
        if path == "me/mailboxSettings/workingHours":
            return self._working_hours_response()
        if path == "me/mailboxSettings":
            return self._mailbox_settings_response()

        # ── /me/mailFolders ──────────────────────────────────
        if path.startswith("me/mailFolders"):
            return self._mail_folder_response(path, method, json_body, qs)

        # ── /me/messages ─────────────────────────────────────
        if path.startswith("me/messages"):
            return self._messages_response(path, method, json_body)

        # ── /me/sendMail ─────────────────────────────────────
        if path == "me/sendMail":
            return _make_response(202, {})

        # ── /me/joinedTeams ──────────────────────────────────
        if path == "me/joinedTeams":
            return self._joined_teams_response()

        # ── /teams/{id}/channels ─────────────────────────────
        if re.match(r"teams/[^/]+/channels$", path):
            team_id = path.split("/")[1]
            return self._team_channels_response(team_id)

        # ── /teams/{id}/members ──────────────────────────────
        if re.match(r"teams/[^/]+/members$", path):
            team_id = path.split("/")[1]
            return self._team_members_response(team_id)

        # ── /me/chats ────────────────────────────────────────
        if path == "me/chats":
            return self._chats_response()

        # ── /me/presence ─────────────────────────────────────
        if path == "me/presence":
            return self._my_presence_response()

        # ── /me/todo/lists ───────────────────────────────────
        if re.match(r"me/todo/lists/[^/]+/tasks/[^/]+$", path):
            parts = path.split("/")
            return self._single_task_response(parts[3], parts[5])
        if re.match(r"me/todo/lists/[^/]+/tasks", path):
            list_id = path.split("/")[3]
            return self._tasks_response(list_id)
        if path == "me/todo/lists":
            return self._task_lists_response()

        # ── /me/people ───────────────────────────────────────
        if path.startswith("me/people"):
            return self._people_response()

        # ── /me/contacts ─────────────────────────────────────
        if path.startswith("me/contacts"):
            return self._contacts_response()

        # ── /me/onlineMeetings ───────────────────────────────
        if re.match(r"me/onlineMeetings/[^/]+$", path):
            meeting_id = path.split("/")[2]
            return self._single_meeting_response(meeting_id)
        if path.startswith("me/onlineMeetings"):
            return self._online_meetings_response()

        # ── /places ──────────────────────────────────────────
        if path == "places/microsoft.graph.room":
            return self._rooms_response()
        if path == "places/microsoft.graph.roomList":
            return self._room_lists_response()

        # ── /communications/getPresencesByUserId ─────────────
        if path == "communications/getPresencesByUserId":
            return self._presences_by_user_id_response(json_body)

        # ── /users/{id}/presence ─────────────────────────────
        if re.match(r"users/[^/]+/presence$", path):
            user_id = path.split("/")[1]
            return self._user_presence_response(user_id)

        # ── /users/* ─────────────────────────────────────────
        if path.startswith("users/") or path == "users":
            return self._directory_response(path, qs)

        logger.warning("[MOCK] Unhandled Graph URL: %s %s", method, url)
        return _make_response(200, {"value": []})

    # ── Profile ──────────────────────────────────────────────

    def _profile_response(self) -> _MockResponse:
        p = self._profile
        return _make_response(200, {
            "id": p.user_id,
            "displayName": p.full_name,
            "givenName": p.given_name,
            "surname": p.surname,
            "mail": p.email,
            "userPrincipalName": p.email,
            "jobTitle": p.job_title,
            "department": p.department,
            "officeLocation": p.office_location,
            "businessPhones": [],
            "mobilePhone": None,
            "preferredLanguage": "de-DE",
        })

    def _photo_response(self) -> _MockResponse:
        # 1x1 transparent PNG
        return _make_response(404, {"error": {"code": "ImageNotFound", "message": "No photo"}})

    # ── Calendar ─────────────────────────────────────────────

    def _calendars_response(self) -> _MockResponse:
        return _make_response(200, {
            "value": [{
                "id": "mock-calendar-default",
                "name": "Calendar",
                "isDefaultCalendar": True,
                "canEdit": True,
                "color": "auto",
            }],
        })

    def _calendar_events_response(self, url: str) -> _MockResponse:
        return _make_response(200, {"value": self._profile.calendar_events})

    def _create_event_response(self, json_body: dict | None) -> _MockResponse:
        event = dict(json_body or {})
        event.setdefault("id", _uid())
        event.setdefault("webLink", "https://outlook.office.com/calendar/mock/" + event["id"])
        return _make_response(201, event)

    def _schedule_response(self, json_body: dict | None) -> _MockResponse:
        emails = (json_body or {}).get("schedules", [])
        schedules = []
        for email in emails:
            schedules.append({
                "scheduleId": email,
                "availabilityView": "0000000000",
                "scheduleItems": [],
                "workingHours": {
                    "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
                    "startTime": "08:00:00.0000000",
                    "endTime": "17:00:00.0000000",
                    "timeZone": {"name": "Europe/Berlin"},
                },
            })
        return _make_response(200, {"value": schedules})

    # ── Categories ───────────────────────────────────────────

    def _categories_response(self) -> _MockResponse:
        categories = [
            {"displayName": "Red Category", "color": "preset0"},
            {"displayName": "Orange Category", "color": "preset1"},
            {"displayName": "Yellow Category", "color": "preset2"},
            {"displayName": "Green Category", "color": "preset3"},
            {"displayName": "Blue Category", "color": "preset4"},
            {"displayName": "Purple Category", "color": "preset5"},
        ]
        return _make_response(200, {"value": categories})

    # ── Mail ─────────────────────────────────────────────────

    def _mailbox_settings_response(self) -> _MockResponse:
        return _make_response(200, {
            "timeZone": "Europe/Berlin",
            "language": {"locale": "de-DE", "displayName": "Deutsch"},
            "dateFormat": "dd.MM.yyyy",
            "timeFormat": "HH:mm",
            "automaticRepliesSetting": {
                "status": "disabled",
                "externalAudience": "none",
                "internalReplyMessage": "",
                "externalReplyMessage": "",
                "scheduledStartDateTime": {
                    "dateTime": "2026-04-06T00:00:00.0000000",
                    "timeZone": "Europe/Berlin",
                },
                "scheduledEndDateTime": {
                    "dateTime": "2026-04-07T00:00:00.0000000",
                    "timeZone": "Europe/Berlin",
                },
            },
            "workingHours": {
                "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
                "startTime": "08:00:00.0000000",
                "endTime": "17:00:00.0000000",
                "timeZone": {"name": "Europe/Berlin"},
            },
        })

    def _mail_folder_response(self, path: str, method: str,
                              json_body: dict | None, qs: dict) -> _MockResponse:
        # Parse folder ID from path: me/mailFolders/{folder_id}/...
        # parts: ["me", "mailFolders", "{folder_id}", "messages"|"childFolders"]
        parts = path.split("/")
        folder_id = parts[2] if len(parts) > 2 else None

        # GET me/mailFolders/{folder_id}/childFolders
        if len(parts) >= 4 and parts[3] == "childFolders" and method == "GET":
            child_folders = [
                f for f in self._profile.mail_folders
                if f.get("parentFolderId") == folder_id
            ]
            return _make_response(200, {"value": child_folders})

        # GET me/mailFolders/{folder_id}/messages
        if "messages" in path and method == "GET":
            # Filter messages by _folder_id
            if folder_id:
                messages = [
                    m for m in self._profile.mail_messages
                    if m.get("_folder_id") == folder_id
                ]
            else:
                messages = self._profile.mail_messages

            total = len(messages)
            top = int(qs.get("$top", [str(total)])[0])
            skip = int(qs.get("$skip", ["0"])[0])
            page = messages[skip:skip + top]

            result: dict = {"value": page}
            # Include count if requested
            if "$count" in qs and qs["$count"][0].lower() == "true":
                result["@odata.count"] = total
            else:
                result["@odata.count"] = total

            return _make_response(200, result)

        # GET me/mailFolders (list all folders)
        if path == "me/mailFolders" and method == "GET":
            return _make_response(200, {"value": self._profile.mail_folders})

        # GET me/mailFolders/{folder_id} (single folder info)
        if folder_id and method == "GET":
            for f in self._profile.mail_folders:
                if f.get("id") == folder_id:
                    return _make_response(200, f)
            # Fallback for legacy behaviour
            folder_messages = [
                m for m in self._profile.mail_messages
                if m.get("_folder_id") == folder_id
            ]
            return _make_response(200, {
                "id": folder_id,
                "displayName": folder_id.capitalize(),
                "totalItemCount": len(folder_messages),
                "unreadItemCount": sum(1 for m in folder_messages if not m.get("isRead")),
            })

        # Fallback: folder info for inbox
        return _make_response(200, {
            "id": "mock-inbox-id",
            "displayName": "Inbox",
            "totalItemCount": len(self._profile.mail_messages),
            "unreadItemCount": sum(1 for m in self._profile.mail_messages if not m.get("isRead")),
        })

    def _messages_response(self, path: str, method: str,
                           json_body: dict | None) -> _MockResponse:
        # POST me/messages/{id}/reply
        if method == "POST" and path.endswith("/reply"):
            return _make_response(202, {})
        # POST me/messages/{id}/replyAll
        if method == "POST" and path.endswith("/replyAll"):
            return _make_response(202, {})
        # POST me/messages/{id}/send
        if method == "POST" and path.endswith("/send"):
            return _make_response(202, {})
        if method == "PATCH":
            # Mark read, set categories, etc.
            return _make_response(200, json_body or {})
        if method == "POST":
            # Create draft
            draft = dict(json_body or {})
            draft.setdefault("id", _uid())
            draft.setdefault("webLink", "https://outlook.office.com/mock")
            return _make_response(201, draft)
        # GET single message by ID — extract ID from path like me/messages/{id}
        parts = path.rstrip("/").split("/")
        msg_id = parts[-1] if len(parts) >= 2 else ""
        for msg in self._profile.mail_messages:
            if msg.get("id") == msg_id:
                return _make_response(200, msg)
        # Fallback: return first message
        return _make_response(200, self._profile.mail_messages[0] if self._profile.mail_messages else {})

    # ── Teams ────────────────────────────────────────────────

    def _joined_teams_response(self) -> _MockResponse:
        return _make_response(200, {"value": self._profile.teams})

    def _team_channels_response(self, team_id: str) -> _MockResponse:
        channels = [
            {
                "id": f"{team_id}-general",
                "displayName": "General",
                "description": "General discussion",
                "membershipType": "standard",
            },
            {
                "id": f"{team_id}-random",
                "displayName": "Random",
                "description": "Off-topic chat",
                "membershipType": "standard",
            },
        ]
        return _make_response(200, {"value": channels})

    def _team_members_response(self, team_id: str) -> _MockResponse:
        # Return a subset of directory users as team members
        members = []
        for user in self._profile.directory_users[:5]:
            members.append({
                "id": _uid(),
                "displayName": user.get("displayName", ""),
                "email": user.get("mail", ""),
                "roles": [],
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
            })
        return _make_response(200, {"value": members})

    # ── Chats ────────────────────────────────────────────────

    def _chats_response(self) -> _MockResponse:
        return _make_response(200, {"value": self._profile.chats})

    # ── Directory ────────────────────────────────────────────

    def _directory_response(self, path: str, qs: dict) -> _MockResponse:
        # /users/{id}/photo/$value
        if "photo/$value" in path:
            # Extract user ID from path
            parts = path.split("/")
            if len(parts) >= 2:
                user_id = parts[1]
                photo_bytes = self._profile.user_photos.get(user_id)
                if photo_bytes:
                    resp = _make_response(200, photo_bytes)
                    resp.headers["Content-Type"] = "image/svg+xml"
                    return resp
            return _make_response(404, {"error": {"code": "ImageNotFound"}})
        # /users/{id}/photo (metadata, not $value)
        if "photo" in path:
            return _make_response(404, {"error": {"code": "ImageNotFound"}})
        # /users?$top=...
        if path == "users":
            users = self._profile.directory_users
            top = int(qs.get("$top", [str(len(users))])[0])
            return _make_response(200, {"value": users[:top]})
        # /users/{id}
        user_id = path.split("/")[1] if "/" in path else ""
        for u in self._profile.directory_users:
            if u.get("id") == user_id or u.get("mail", "").lower() == user_id.lower():
                return _make_response(200, u)
        return _make_response(404, {"error": {"code": "Request_ResourceNotFound"}})

    # ── Presence ─────────────────────────────────────────────

    def _my_presence_response(self) -> _MockResponse:
        return _make_response(200, {
            "id": self._profile.user_id,
            "availability": "Available",
            "activity": "Available",
        })

    def _user_presence_response(self, user_id: str) -> _MockResponse:
        return _make_response(200, {
            "id": user_id,
            "availability": "Available",
            "activity": "Available",
        })

    def _presences_by_user_id_response(self, json_body: dict | None) -> _MockResponse:
        ids = (json_body or {}).get("ids", [])
        presences = []
        for uid in ids:
            presences.append({
                "id": uid,
                "availability": "Available",
                "activity": "Available",
            })
        return _make_response(200, {"value": presences})

    # ── Tasks (To Do) ────────────────────────────────────────

    _TASK_LISTS = [
        {"id": "tasklist-tasks", "displayName": "Tasks", "isOwner": True, "isShared": False,
         "wellknownListName": "defaultList"},
        {"id": "tasklist-shopping", "displayName": "Einkaufsliste", "isOwner": True, "isShared": False,
         "wellknownListName": "none"},
        {"id": "tasklist-work", "displayName": "Arbeitsprojekte", "isOwner": True, "isShared": True,
         "wellknownListName": "none"},
    ]

    _TASKS_BY_LIST: dict[str, list[dict]] = {
        "tasklist-tasks": [
            {"id": "task-001", "title": "Quartalsbericht vorbereiten", "status": "notStarted",
             "importance": "high", "isReminderOn": True,
             "dueDateTime": {"dateTime": "2026-04-10T17:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": "Q1 Bericht fuer Management erstellen"}},
            {"id": "task-002", "title": "Reisekosten einreichen", "status": "inProgress",
             "importance": "normal", "isReminderOn": False,
             "dueDateTime": {"dateTime": "2026-04-08T17:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": ""}},
            {"id": "task-003", "title": "Teammeeting planen", "status": "completed",
             "importance": "normal", "isReminderOn": False,
             "dueDateTime": {"dateTime": "2026-04-05T12:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": ""}},
        ],
        "tasklist-shopping": [
            {"id": "task-010", "title": "Druckerpapier bestellen", "status": "notStarted",
             "importance": "normal", "isReminderOn": False,
             "dueDateTime": None,
             "body": {"contentType": "text", "content": "A4, 80g/m2, 5 Pakete"}},
            {"id": "task-011", "title": "Kaffee fuer Kueche", "status": "notStarted",
             "importance": "low", "isReminderOn": False,
             "dueDateTime": None,
             "body": {"contentType": "text", "content": ""}},
            {"id": "task-012", "title": "Whiteboardmarker", "status": "notStarted",
             "importance": "low", "isReminderOn": False,
             "dueDateTime": None,
             "body": {"contentType": "text", "content": "Schwarz, Rot, Blau, Gruen"}},
        ],
        "tasklist-work": [
            {"id": "task-020", "title": "API-Dokumentation aktualisieren", "status": "inProgress",
             "importance": "high", "isReminderOn": True,
             "dueDateTime": {"dateTime": "2026-04-15T17:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": "Neue Endpoints dokumentieren"}},
            {"id": "task-021", "title": "Code-Review fuer Feature-Branch", "status": "notStarted",
             "importance": "high", "isReminderOn": False,
             "dueDateTime": {"dateTime": "2026-04-09T12:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": ""}},
            {"id": "task-022", "title": "Unit-Tests erweitern", "status": "notStarted",
             "importance": "normal", "isReminderOn": False,
             "dueDateTime": {"dateTime": "2026-04-12T17:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": "Coverage auf 80% erhoehen"}},
            {"id": "task-023", "title": "Performance-Optimierung DB-Queries", "status": "notStarted",
             "importance": "normal", "isReminderOn": False,
             "dueDateTime": {"dateTime": "2026-04-20T17:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": ""}},
            {"id": "task-024", "title": "Deployment-Pipeline pruefen", "status": "completed",
             "importance": "normal", "isReminderOn": False,
             "dueDateTime": {"dateTime": "2026-04-04T17:00:00.0000000", "timeZone": "Europe/Berlin"},
             "body": {"contentType": "text", "content": ""}},
        ],
    }

    def _task_lists_response(self) -> _MockResponse:
        return _make_response(200, {"value": self._TASK_LISTS})

    def _tasks_response(self, list_id: str) -> _MockResponse:
        tasks = self._TASKS_BY_LIST.get(list_id, [])
        return _make_response(200, {"value": tasks})

    def _single_task_response(self, list_id: str, task_id: str) -> _MockResponse:
        tasks = self._TASKS_BY_LIST.get(list_id, [])
        for task in tasks:
            if task["id"] == task_id:
                return _make_response(200, task)
        return _make_response(404, {"error": {"code": "itemNotFound", "message": "Task not found"}})

    # ── People ───────────────────────────────────────────────

    def _people_response(self) -> _MockResponse:
        people = [
            {
                "id": "person-001",
                "displayName": "Max Mustermann",
                "givenName": "Max",
                "surname": "Mustermann",
                "emailAddresses": [{"address": "max.mustermann@example.com", "rank": 1}],
                "department": "Vertrieb",
                "jobTitle": "Vertriebsleiter",
                "companyName": "Example GmbH",
                "personType": {"class": "Person", "subclass": "OrganizationUser"},
            },
            {
                "id": "person-002",
                "displayName": "Anna Schmidt",
                "givenName": "Anna",
                "surname": "Schmidt",
                "emailAddresses": [{"address": "anna.schmidt@example.com", "rank": 1}],
                "department": "Engineering",
                "jobTitle": "Software-Entwicklerin",
                "companyName": "Example GmbH",
                "personType": {"class": "Person", "subclass": "OrganizationUser"},
            },
            {
                "id": "person-003",
                "displayName": "Thomas Weber",
                "givenName": "Thomas",
                "surname": "Weber",
                "emailAddresses": [{"address": "thomas.weber@example.com", "rank": 1}],
                "department": "Marketing",
                "jobTitle": "Marketing-Manager",
                "companyName": "Example GmbH",
                "personType": {"class": "Person", "subclass": "OrganizationUser"},
            },
            {
                "id": "person-004",
                "displayName": "Lisa Mueller",
                "givenName": "Lisa",
                "surname": "Mueller",
                "emailAddresses": [{"address": "lisa.mueller@example.com", "rank": 1}],
                "department": "Personal",
                "jobTitle": "HR-Leiterin",
                "companyName": "Example GmbH",
                "personType": {"class": "Person", "subclass": "OrganizationUser"},
            },
            {
                "id": "person-005",
                "displayName": "Stefan Braun",
                "givenName": "Stefan",
                "surname": "Braun",
                "emailAddresses": [{"address": "stefan.braun@example.com", "rank": 1}],
                "department": "Finanzen",
                "jobTitle": "Controller",
                "companyName": "Example GmbH",
                "personType": {"class": "Person", "subclass": "OrganizationUser"},
            },
        ]
        return _make_response(200, {"value": people})

    def _contacts_response(self) -> _MockResponse:
        contacts = [
            {
                "id": "contact-001",
                "givenName": "Klaus",
                "surname": "Fischer",
                "displayName": "Klaus Fischer",
                "emailAddresses": [{"address": "klaus.fischer@partner.de", "name": "Klaus Fischer"}],
                "businessPhones": ["+49 711 1234567"],
                "companyName": "Partner AG",
                "jobTitle": "Geschaeftsfuehrer",
                "department": "Geschaeftsleitung",
            },
            {
                "id": "contact-002",
                "givenName": "Maria",
                "surname": "Hoffmann",
                "displayName": "Maria Hoffmann",
                "emailAddresses": [{"address": "maria.hoffmann@kunde.de", "name": "Maria Hoffmann"}],
                "businessPhones": ["+49 711 7654321"],
                "companyName": "Kunde GmbH",
                "jobTitle": "Einkaufsleiterin",
                "department": "Einkauf",
            },
            {
                "id": "contact-003",
                "givenName": "Peter",
                "surname": "Zimmermann",
                "displayName": "Peter Zimmermann",
                "emailAddresses": [{"address": "peter.zimmermann@lieferant.de", "name": "Peter Zimmermann"}],
                "businessPhones": ["+49 711 9876543"],
                "companyName": "Lieferant KG",
                "jobTitle": "Technischer Leiter",
                "department": "Technik",
            },
        ]
        return _make_response(200, {"value": contacts})

    # ── Places ───────────────────────────────────────────────

    def _rooms_response(self) -> _MockResponse:
        rooms = [
            {
                "id": "room-stuttgart",
                "displayName": "Raum Stuttgart",
                "emailAddress": "raum.stuttgart@example.com",
                "capacity": 12,
                "building": "Hauptgebaeude",
                "floorNumber": 2,
                "isWheelChairAccessible": True,
                "audioDeviceName": "Jabra Speak 750",
                "videoDeviceName": "Logitech Rally",
                "phone": "+49 711 1000001",
            },
            {
                "id": "room-heidelberg",
                "displayName": "Raum Heidelberg",
                "emailAddress": "raum.heidelberg@example.com",
                "capacity": 8,
                "building": "Hauptgebaeude",
                "floorNumber": 1,
                "isWheelChairAccessible": True,
                "audioDeviceName": "Jabra Speak 510",
                "videoDeviceName": None,
                "phone": "+49 711 1000002",
            },
            {
                "id": "room-reutlingen",
                "displayName": "Raum Reutlingen",
                "emailAddress": "raum.reutlingen@example.com",
                "capacity": 20,
                "building": "Neubau",
                "floorNumber": 3,
                "isWheelChairAccessible": True,
                "audioDeviceName": "Poly Studio",
                "videoDeviceName": "Poly Studio X50",
                "phone": "+49 711 1000003",
            },
            {
                "id": "room-tuebingen",
                "displayName": "Raum Tuebingen",
                "emailAddress": "raum.tuebingen@example.com",
                "capacity": 4,
                "building": "Neubau",
                "floorNumber": 1,
                "isWheelChairAccessible": False,
                "audioDeviceName": None,
                "videoDeviceName": None,
                "phone": None,
            },
        ]
        return _make_response(200, {"value": rooms})

    def _room_lists_response(self) -> _MockResponse:
        room_lists = [
            {
                "id": "roomlist-hauptgebaeude",
                "displayName": "Hauptgebaeude",
                "emailAddress": "hauptgebaeude@example.com",
            },
            {
                "id": "roomlist-neubau",
                "displayName": "Neubau",
                "emailAddress": "neubau@example.com",
            },
        ]
        return _make_response(200, {"value": room_lists})

    # ── Mailbox Settings (extended) ──────────────────────────

    def _automatic_replies_response(self) -> _MockResponse:
        return _make_response(200, {
            "status": "disabled",
            "externalAudience": "none",
            "internalReplyMessage": "",
            "externalReplyMessage": "",
            "scheduledStartDateTime": {
                "dateTime": "2026-04-06T00:00:00.0000000",
                "timeZone": "Europe/Berlin",
            },
            "scheduledEndDateTime": {
                "dateTime": "2026-04-07T00:00:00.0000000",
                "timeZone": "Europe/Berlin",
            },
        })

    def _working_hours_response(self) -> _MockResponse:
        return _make_response(200, {
            "daysOfWeek": ["monday", "tuesday", "wednesday", "thursday", "friday"],
            "startTime": "08:00:00.0000000",
            "endTime": "17:00:00.0000000",
            "timeZone": {"name": "Europe/Berlin"},
        })

    # ── Online Meetings ──────────────────────────────────────

    _ONLINE_MEETINGS = [
        {
            "id": "meeting-001",
            "subject": "Sprint Planning KW15",
            "startDateTime": "2026-04-07T09:00:00Z",
            "endDateTime": "2026-04-07T10:00:00Z",
            "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/mock-meeting-001",
            "videoTeleconferenceId": "123456789",
            "chatInfo": {"threadId": "19:meeting_mock001@thread.v2"},
            "participants": {
                "organizer": {
                    "upn": "mock@example.com",
                    "identity": {"user": {"displayName": "Mock User"}},
                },
            },
        },
        {
            "id": "meeting-002",
            "subject": "Projektstatus Besprechung",
            "startDateTime": "2026-04-08T14:00:00Z",
            "endDateTime": "2026-04-08T15:30:00Z",
            "joinWebUrl": "https://teams.microsoft.com/l/meetup-join/mock-meeting-002",
            "videoTeleconferenceId": "987654321",
            "chatInfo": {"threadId": "19:meeting_mock002@thread.v2"},
            "participants": {
                "organizer": {
                    "upn": "max.mustermann@example.com",
                    "identity": {"user": {"displayName": "Max Mustermann"}},
                },
            },
        },
    ]

    def _online_meetings_response(self) -> _MockResponse:
        return _make_response(200, {"value": self._ONLINE_MEETINGS})

    def _single_meeting_response(self, meeting_id: str) -> _MockResponse:
        for meeting in self._ONLINE_MEETINGS:
            if meeting["id"] == meeting_id:
                return _make_response(200, meeting)
        return _make_response(404, {"error": {"code": "itemNotFound", "message": "Meeting not found"}})

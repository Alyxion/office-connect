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

        # ── /me/mailboxSettings ──────────────────────────────
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
            "automaticRepliesSetting": {"status": "disabled"},
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

"""URL-routing mock transport for MS Graph API calls.

Intercepts ``run_async()`` requests and returns synthetic
``AsyncResponseWrapper``-compatible responses based on URL patterns.
"""

from __future__ import annotations

import json as _json
import logging
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

        # ── /me/calendars ────────────────────────────────────
        if path == "me/calendars":
            return self._calendars_response()

        # ── /me/calendars/{id}/calendarView ──────────────────
        if "calendarView" in path:
            return self._calendar_events_response(url)

        # ── /me/calendar/getSchedule ─────────────────────────
        if path == "me/calendar/getSchedule":
            return self._schedule_response(json_body)

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

    # ── Mail ─────────────────────────────────────────────────

    def _mailbox_settings_response(self) -> _MockResponse:
        return _make_response(200, {
            "timeZone": "Europe/Berlin",
            "language": {"locale": "de-DE", "displayName": "Deutsch"},
            "automaticRepliesSetting": {"status": "disabled"},
        })

    def _mail_folder_response(self, path: str, method: str,
                              json_body: dict | None, qs: dict) -> _MockResponse:
        # GET .../mailFolders/inbox/messages
        if "messages" in path and method == "GET":
            messages = self._profile.mail_messages
            top = int(qs.get("$top", [str(len(messages))])[0])
            skip = int(qs.get("$skip", ["0"])[0])
            page = messages[skip:skip + top]
            return _make_response(200, {
                "value": page,
                "@odata.count": len(messages),
            })
        # Folder info
        return _make_response(200, {
            "id": "mock-inbox-id",
            "displayName": "Inbox",
            "totalItemCount": len(self._profile.mail_messages),
            "unreadItemCount": sum(1 for m in self._profile.mail_messages if not m.get("isRead")),
        })

    def _messages_response(self, path: str, method: str,
                           json_body: dict | None) -> _MockResponse:
        if method == "PATCH":
            # Mark read, set categories, etc.
            return _make_response(200, json_body or {})
        if method == "POST":
            return _make_response(201, {"id": "mock-draft-id", "webLink": "https://outlook.office.com/mock"})
        # GET single message
        return _make_response(200, self._profile.mail_messages[0] if self._profile.mail_messages else {})

    # ── Directory ────────────────────────────────────────────

    def _directory_response(self, path: str, qs: dict) -> _MockResponse:
        # /users/{id}/photo/$value
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

"""Default synthetic MS Graph data for mock users.

Every function returns dicts in the exact JSON shape that the real
MS Graph API would return, so that ``ProfileHandler``, ``CalendarHandler``,
``MailHandler``, and ``DirectoryHandler`` parse them unchanged.
"""

from __future__ import annotations

import uuid
from datetime import datetime, timedelta, timezone


def _uid() -> str:
    return str(uuid.uuid4())


def _dt(dt: datetime) -> str:
    """ISO-8601 without timezone (MS Graph dateTime format)."""
    return dt.strftime("%Y-%m-%dT%H:%M:%S.0000000")


def _now() -> datetime:
    return datetime.now(timezone.utc)


# ── Calendar ─────────────────────────────────────────────────────

def default_calendar_events() -> list[dict]:
    """~8 synthetic calendar events spanning today and tomorrow."""
    today = _now().replace(hour=0, minute=0, second=0, microsecond=0)
    tomorrow = today + timedelta(days=1)

    def _event(subject: str, start: datetime, end: datetime,
               organizer_name: str = "Max Mustermann",
               organizer_email: str = "max.mustermann@example.com",
               is_all_day: bool = False,
               location: str = "",
               online_meeting_url: str = "") -> dict:
        ev = {
            "id": _uid(),
            "subject": subject,
            "start": {"dateTime": _dt(start), "timeZone": "Europe/Berlin"},
            "end": {"dateTime": _dt(end), "timeZone": "Europe/Berlin"},
            "isAllDay": is_all_day,
            "organizer": {
                "emailAddress": {"name": organizer_name, "address": organizer_email},
            },
            "attendees": [
                {
                    "emailAddress": {"name": organizer_name, "address": organizer_email},
                    "status": {"response": "accepted"},
                    "type": "required",
                },
            ],
            "bodyPreview": "",
            "location": {"displayName": location},
            "isOnlineMeeting": bool(online_meeting_url),
            "onlineMeeting": {"joinUrl": online_meeting_url} if online_meeting_url else None,
            "showAs": "busy",
            "responseStatus": {"response": "accepted"},
            "sensitivity": "normal",
            "importance": "normal",
        }
        return ev

    return [
        _event("Daily Standup", today.replace(hour=9), today.replace(hour=9, minute=30),
               location="Teams", online_meeting_url="https://teams.microsoft.com/mock-standup"),
        _event("Project Review", today.replace(hour=10), today.replace(hour=11),
               organizer_name="Anna Schmidt", organizer_email="anna.schmidt@example.com",
               location="Meeting Room A"),
        _event("Lunch Break", today.replace(hour=12), today.replace(hour=13)),
        _event("1:1 with Manager", today.replace(hour=14), today.replace(hour=14, minute=30),
               organizer_name="Klaus Weber", organizer_email="klaus.weber@example.com",
               online_meeting_url="https://teams.microsoft.com/mock-1on1"),
        _event("Sprint Planning", today.replace(hour=15), today.replace(hour=16),
               organizer_name="Lisa Braun", organizer_email="lisa.braun@example.com",
               location="Conference Room B"),
        _event("Company All-Hands", tomorrow.replace(hour=10), tomorrow.replace(hour=11, minute=30),
               organizer_name="CEO Office", organizer_email="ceo.office@example.com",
               location="Main Hall", online_meeting_url="https://teams.microsoft.com/mock-allhands"),
        _event("Team Workshop", tomorrow.replace(hour=13), tomorrow.replace(hour=16),
               organizer_name="Anna Schmidt", organizer_email="anna.schmidt@example.com",
               location="Innovation Lab"),
        _event("National Holiday", today, today + timedelta(days=1),
               is_all_day=True, organizer_name="System", organizer_email="system@example.com"),
    ]


# ── Mail ─────────────────────────────────────────────────────────

def default_mail_inbox() -> list[dict]:
    """~5 synthetic inbox messages in MS Graph JSON shape."""
    now = _now()

    def _msg(subject: str, sender_name: str, sender_email: str,
             body: str, minutes_ago: int = 0, is_read: bool = False,
             has_attachments: bool = False) -> dict:
        received = now - timedelta(minutes=minutes_ago)
        return {
            "id": _uid(),
            "subject": subject,
            "from": {"emailAddress": {"name": sender_name, "address": sender_email}},
            "toRecipients": [{"emailAddress": {"name": "Mock User", "address": "mock@example.com"}}],
            "receivedDateTime": received.isoformat(),
            "sentDateTime": received.isoformat(),
            "isRead": is_read,
            "isDraft": False,
            "importance": "normal",
            "hasAttachments": has_attachments,
            "bodyPreview": body[:200],
            "body": {"contentType": "text", "content": body},
            "categories": [],
            "flag": {"flagStatus": "notFlagged"},
            "webLink": f"https://outlook.office.com/mail/mock/{_uid()}",
            "conversationId": _uid(),
        }

    return [
        _msg("Q1 Sales Report Ready", "Anna Schmidt", "anna.schmidt@example.com",
             "Hi,\n\nThe Q1 sales report is ready for review. Please check the attached PDF and let me know if you have any questions.\n\nBest regards,\nAnna",
             minutes_ago=15, has_attachments=True),
        _msg("Meeting Follow-up: Project Alpha", "Klaus Weber", "klaus.weber@example.com",
             "Team,\n\nThanks for the productive meeting today. Here are the action items we agreed on:\n1. Complete API integration by Friday\n2. Update documentation\n3. Schedule demo with stakeholders\n\nPlease update your tasks accordingly.\n\nBest,\nKlaus",
             minutes_ago=45, is_read=True),
        _msg("Re: Vacation Request", "HR Department", "hr@example.com",
             "Your vacation request for March 24-28 has been approved. Enjoy your time off!",
             minutes_ago=120, is_read=True),
        _msg("New IT Security Policy", "IT Department", "it@example.com",
             "Dear colleagues,\n\nPlease review the updated IT security policy attached to this email. All employees must acknowledge the new policy by end of month.\n\nIT Department",
             minutes_ago=240, has_attachments=True),
        _msg("Lunch plans?", "Lisa Braun", "lisa.braun@example.com",
             "Hey! Want to grab lunch at the Italian place today? I heard they have a new menu.",
             minutes_ago=30),
    ]


# ── Directory ────────────────────────────────────────────────────

def default_company_directory() -> list[dict]:
    """~15 synthetic users with manager links (org tree) in MS Graph JSON shape."""
    # CEO
    ceo_id = _uid()
    # VPs
    vp_tech_id = _uid()
    vp_sales_id = _uid()
    vp_hr_id = _uid()

    def _user(uid: str, given: str, surname: str, email: str,
              title: str, dept: str, manager_id: str | None = None,
              location: str = "Headquarters") -> dict:
        return {
            "id": uid,
            "displayName": f"{given} {surname}",
            "givenName": given,
            "surname": surname,
            "mail": email,
            "userPrincipalName": email,
            "jobTitle": title,
            "department": dept,
            "officeLocation": location,
            "mobilePhone": "+49 170 " + uid[:7].replace("-", ""),
            "businessPhones": ["+49 7123 " + uid[:4].replace("-", "")],
            "accountEnabled": True,
            "manager": {"id": manager_id} if manager_id else None,
        }

    return [
        _user(ceo_id, "Heinrich", "Fischer", "heinrich.fischer@example.com",
              "CEO", "Executive Board"),
        _user(vp_tech_id, "Klaus", "Weber", "klaus.weber@example.com",
              "VP Technology", "IT / Digital", manager_id=ceo_id),
        _user(vp_sales_id, "Sabine", "Mueller", "sabine.mueller@example.com",
              "VP Sales", "Sales", manager_id=ceo_id),
        _user(vp_hr_id, "Petra", "Schneider", "petra.schneider@example.com",
              "VP Human Resources", "HR", manager_id=ceo_id),
        _user(_uid(), "Anna", "Schmidt", "anna.schmidt@example.com",
              "Sales Manager", "Sales", manager_id=vp_sales_id),
        _user(_uid(), "Max", "Mustermann", "max.mustermann@example.com",
              "Software Engineer", "IT / Digital", manager_id=vp_tech_id),
        _user(_uid(), "Lisa", "Braun", "lisa.braun@example.com",
              "Scrum Master", "IT / Digital", manager_id=vp_tech_id),
        _user(_uid(), "Thomas", "Keller", "thomas.keller@example.com",
              "Data Scientist", "IT / Digital", manager_id=vp_tech_id),
        _user(_uid(), "Julia", "Richter", "julia.richter@example.com",
              "Sales Representative", "Sales", manager_id=vp_sales_id),
        _user(_uid(), "Markus", "Bauer", "markus.bauer@example.com",
              "Key Account Manager", "Sales", manager_id=vp_sales_id,
              location="Stuttgart"),
        _user(_uid(), "Christine", "Wagner", "christine.wagner@example.com",
              "HR Business Partner", "HR", manager_id=vp_hr_id),
        _user(_uid(), "Stefan", "Hoffmann", "stefan.hoffmann@example.com",
              "DevOps Engineer", "IT / Digital", manager_id=vp_tech_id),
        _user(_uid(), "Katrin", "Schwarz", "katrin.schwarz@example.com",
              "UX Designer", "IT / Digital", manager_id=vp_tech_id),
        _user(_uid(), "Frank", "Zimmermann", "frank.zimmermann@example.com",
              "Product Owner", "IT / Digital", manager_id=vp_tech_id),
        _user(_uid(), "Monika", "Krueger", "monika.krueger@example.com",
              "Marketing Manager", "Marketing", manager_id=ceo_id),
    ]

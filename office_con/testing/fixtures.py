"""Default synthetic MS Graph data for mock users.

Every function returns dicts in the exact JSON shape that the real
MS Graph API would return, so that ``ProfileHandler``, ``CalendarHandler``,
``MailHandler``, and ``DirectoryHandler`` parse them unchanged.
"""

from __future__ import annotations

import base64
import uuid
from datetime import datetime, timedelta, timezone

from office_con.testing.mock_data import MockUserProfile, generate_avatar_svg, load_face_photo


def _uid() -> str:
    return str(uuid.uuid4())


def _dt(dt: datetime) -> str:
    """ISO-8601 without timezone (MS Graph dateTime format)."""
    return dt.strftime("%Y-%m-%dT%H:%M:%S.0000000")


def _now() -> datetime:
    return datetime.now(timezone.utc)


# ── Mail Folders ────────────────────────────────────────────────────

def default_mail_folders() -> list[dict]:
    """~10 synthetic mail folders with realistic Outlook structure."""
    inbox_id = "inbox"
    drafts_id = "drafts"
    sent_id = "sentitems"
    deleted_id = "deleteditems"
    archive_id = "archive"
    junk_id = "junkemail"
    outbox_id = "outbox"
    notifications_id = "inbox-notifications"
    done_id = "inbox-done"

    def _folder(fid: str, name: str, parent_id: str | None = None,
                child_count: int = 0, unread: int = 0,
                total: int = 0) -> dict:
        return {
            "id": fid,
            "displayName": name,
            "parentFolderId": parent_id or "root",
            "childFolderCount": child_count,
            "unreadItemCount": unread,
            "totalItemCount": total,
        }

    return [
        _folder(inbox_id, "Inbox", child_count=2, unread=5, total=18),
        _folder(drafts_id, "Drafts", total=3),
        _folder(sent_id, "Sent Items", total=5),
        _folder(deleted_id, "Deleted Items", total=3),
        _folder(archive_id, "Archive", total=0),
        _folder(junk_id, "Junk Email", total=0),
        _folder(outbox_id, "Outbox", total=0),
        _folder(notifications_id, "Notifications", parent_id=inbox_id, total=3),
        _folder(done_id, "Done", parent_id=inbox_id, total=2),
    ]


# ── Calendar ─────────────────────────────────────────────────────

def default_calendar_events() -> list[dict]:
    """Synthetic calendar events spanning the current month and next 2 months (~60-80 events)."""
    now = _now()
    today = now.replace(hour=0, minute=0, second=0, microsecond=0)
    # Monday of the current week as anchor
    monday = today - timedelta(days=today.weekday())

    # First day of current month and last day of month+2
    first_of_month = today.replace(day=1)
    # 3 months of coverage: current month + next 2
    if first_of_month.month <= 10:
        end_of_range = first_of_month.replace(month=first_of_month.month + 2, day=28)
    else:
        end_of_range = first_of_month.replace(
            year=first_of_month.year + 1,
            month=(first_of_month.month + 2 - 1) % 12 + 1,
            day=28,
        )

    def _event(
        subject: str,
        start: datetime,
        end: datetime,
        organizer_name: str = "Max Mustermann",
        organizer_email: str = "max.mustermann@example.com",
        is_all_day: bool = False,
        location: str = "",
        online_meeting_url: str = "",
        show_as: str = "busy",
        sensitivity: str = "normal",
        attendees: list[dict] | None = None,
        body_preview: str = "",
    ) -> dict:
        if attendees is None:
            attendees = [
                {
                    "emailAddress": {"name": organizer_name, "address": organizer_email},
                    "status": {"response": "accepted"},
                    "type": "required",
                },
            ]
        ev = {
            "id": _uid(),
            "subject": subject,
            "start": {"dateTime": _dt(start), "timeZone": "Europe/Berlin"},
            "end": {"dateTime": _dt(end), "timeZone": "Europe/Berlin"},
            "isAllDay": is_all_day,
            "organizer": {
                "emailAddress": {"name": organizer_name, "address": organizer_email},
            },
            "attendees": attendees,
            "bodyPreview": body_preview,
            "location": {"displayName": location},
            "isOnlineMeeting": bool(online_meeting_url),
            "onlineMeeting": {"joinUrl": online_meeting_url} if online_meeting_url else None,
            "showAs": show_as,
            "responseStatus": {"response": "accepted"},
            "sensitivity": sensitivity,
            "importance": "normal",
        }
        return ev

    events: list[dict] = []

    # Helper: iterate weekdays across the 3-month window
    def _iter_weekdays(target_weekday: int, start: datetime = first_of_month,
                       end: datetime = end_of_range):
        """Yield dates for *target_weekday* (0=Mon) between start and end."""
        cur = start
        # Advance to the first occurrence
        while cur.weekday() != target_weekday:
            cur += timedelta(days=1)
        while cur <= end:
            yield cur
            cur += timedelta(weeks=1)

    # ── Recurring: Daily Standup 9:00-9:30 Mon-Fri (Teams call) ────────
    for wd in range(5):  # Mon-Fri
        for day in _iter_weekdays(wd):
            events.append(
                _event("Daily Standup",
                       day.replace(hour=9, minute=0),
                       day.replace(hour=9, minute=30),
                       location="Microsoft Teams",
                       online_meeting_url="https://teams.microsoft.com/l/meetup-join/standup",
                       body_preview="Quick sync — what did you do yesterday, what are you doing today, any blockers?"),
            )

    # ── Recurring: Weekly Team Meeting Wed 10:00-11:00 ─────────────────
    for wed in _iter_weekdays(2):  # Wednesday
        events.append(
            _event("Weekly Team Meeting",
                   wed.replace(hour=10, minute=0),
                   wed.replace(hour=11, minute=0),
                   organizer_name="Klaus Weber",
                   organizer_email="klaus.weber@example.com",
                   location="Conference Room A",
                   online_meeting_url="https://teams.microsoft.com/l/meetup-join/weekly-team",
                   body_preview="Agenda: status updates, blockers, upcoming milestones."),
        )

    # ── Recurring: Biweekly 1:1 with Manager Thu 14:00-14:30 ──────────
    biweekly_toggle = True
    for thu in _iter_weekdays(3):  # Thursday
        if biweekly_toggle:
            events.append(
                _event("1:1 with Manager",
                       thu.replace(hour=14, minute=0),
                       thu.replace(hour=14, minute=30),
                       organizer_name="Klaus Weber",
                       organizer_email="klaus.weber@example.com",
                       online_meeting_url="https://teams.microsoft.com/l/meetup-join/1on1-manager",
                       body_preview="Career growth, current sprint, feedback."),
            )
        biweekly_toggle = not biweekly_toggle

    # ── Recurring: Monthly All-Hands — first Monday of each month ─────
    seen_months: set[tuple[int, int]] = set()
    for mon_day in _iter_weekdays(0):  # Monday
        key = (mon_day.year, mon_day.month)
        if key not in seen_months and mon_day.day <= 7:
            seen_months.add(key)
            events.append(
                _event("Monthly All-Hands",
                       mon_day.replace(hour=15, minute=0),
                       mon_day.replace(hour=16, minute=0),
                       organizer_name="Heinrich Fischer",
                       organizer_email="heinrich.fischer@example.com",
                       location="Auditorium / Teams Live",
                       online_meeting_url="https://teams.microsoft.com/l/meetup-join/all-hands",
                       body_preview="Company updates, Q&A with leadership."),
            )

    # ── Recurring: Lunch block Mon-Fri (free) ─────────────────────────
    # Only add for the current week and next 2 weeks to avoid clutter
    for week_offset in range(3):
        for wd in range(5):
            day = monday + timedelta(weeks=week_offset, days=wd)
            if first_of_month <= day <= end_of_range:
                events.append(
                    _event("Lunch",
                           day.replace(hour=12, minute=0),
                           day.replace(hour=13, minute=0),
                           show_as="free",
                           body_preview="Blocked for lunch."),
                )

    # ── One-off meetings scattered throughout the 3 months ────────────

    # Week 0 (current week from Monday anchor)
    events.append(
        _event("Project Alpha Design Review",
               (monday + timedelta(days=1)).replace(hour=14, minute=0),
               (monday + timedelta(days=1)).replace(hour=15, minute=0),
               organizer_name="Anna Schmidt",
               organizer_email="anna.schmidt@example.com",
               location="Meeting Room B",
               body_preview="Review the updated wireframes and UX flow for Project Alpha."),
    )
    events.append(
        _event("Customer Demo — Acme Corp",
               (monday + timedelta(days=2)).replace(hour=13, minute=0),
               (monday + timedelta(days=2)).replace(hour=14, minute=0),
               organizer_name="Petra Schneider",
               organizer_email="petra.schneider@example.com",
               location="Presentation Room",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/acme-demo",
               show_as="tentative",
               body_preview="Product demo for Acme Corp — sales engineering lead presenting."),
    )
    events.append(
        _event("Sprint Planning",
               (monday + timedelta(days=0)).replace(hour=11, minute=0),
               (monday + timedelta(days=0)).replace(hour=12, minute=0),
               organizer_name="Lisa Braun",
               organizer_email="lisa.braun@example.com",
               location="Innovation Lab",
               body_preview="Plan the next sprint: capacity, priorities, story pointing."),
    )

    # Week 1
    events.append(
        _event("External Partner Meeting — TechVentures",
               (monday + timedelta(weeks=1, days=1)).replace(hour=13, minute=0),
               (monday + timedelta(weeks=1, days=1)).replace(hour=14, minute=30),
               organizer_name="Heinrich Fischer",
               organizer_email="heinrich.fischer@example.com",
               location="External — TechVentures HQ",
               body_preview="Quarterly partnership review with TechVentures team."),
    )
    events.append(
        _event("Sprint Retrospective",
               (monday + timedelta(weeks=1, days=4)).replace(hour=15, minute=0),
               (monday + timedelta(weeks=1, days=4)).replace(hour=16, minute=0),
               organizer_name="Lisa Braun",
               organizer_email="lisa.braun@example.com",
               location="Innovation Lab",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/retro",
               body_preview="What went well, what can improve, action items."),
    )
    events.append(
        _event("Interview Panel — Senior Developer",
               (monday + timedelta(weeks=1, days=2)).replace(hour=11, minute=0),
               (monday + timedelta(weeks=1, days=2)).replace(hour=12, minute=0),
               organizer_name="Anna Schmidt",
               organizer_email="anna.schmidt@example.com",
               location="Meeting Room C",
               show_as="tentative",
               body_preview="Technical interview for the Senior Developer position. Bring scoring rubrics."),
    )
    events.append(
        _event("Budget Planning Q3",
               (monday + timedelta(weeks=1, days=0)).replace(hour=10, minute=0),
               (monday + timedelta(weeks=1, days=0)).replace(hour=12, minute=0),
               organizer_name="Heinrich Fischer",
               organizer_email="heinrich.fischer@example.com",
               location="Executive Boardroom",
               body_preview="Review Q3 budget allocations across all departments."),
    )

    # Week 2
    events.append(
        _event("Cloud Architecture Workshop",
               (monday + timedelta(weeks=2, days=1)).replace(hour=9, minute=30),
               (monday + timedelta(weeks=2, days=1)).replace(hour=12, minute=0),
               organizer_name="Petra Schneider",
               organizer_email="petra.schneider@example.com",
               location="Training Room 1",
               body_preview="Hands-on workshop: migrating on-prem workloads to Azure."),
    )
    events.append(
        _event("Customer Demo — GlobalTech",
               (monday + timedelta(weeks=2, days=3)).replace(hour=14, minute=0),
               (monday + timedelta(weeks=2, days=3)).replace(hour=15, minute=30),
               organizer_name="Petra Schneider",
               organizer_email="petra.schneider@example.com",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/globaltech-demo",
               show_as="tentative",
               body_preview="Product demo for GlobalTech, focusing on integration capabilities."),
    )
    events.append(
        _event("Interview Panel — UX Designer",
               (monday + timedelta(weeks=2, days=4)).replace(hour=10, minute=0),
               (monday + timedelta(weeks=2, days=4)).replace(hour=11, minute=0),
               organizer_name="Anna Schmidt",
               organizer_email="anna.schmidt@example.com",
               location="Meeting Room C",
               body_preview="Portfolio review and design challenge for UX Designer candidate."),
    )

    # Week 3
    events.append(
        _event("Security Training — Annual Refresh",
               (monday + timedelta(weeks=3, days=2)).replace(hour=13, minute=0),
               (monday + timedelta(weeks=3, days=2)).replace(hour=15, minute=0),
               organizer_name="Klaus Weber",
               organizer_email="klaus.weber@example.com",
               location="Auditorium",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/security-training",
               body_preview="Mandatory annual security awareness training. Bring your laptop."),
    )
    events.append(
        _event("Sprint Review",
               (monday + timedelta(weeks=3, days=4)).replace(hour=14, minute=0),
               (monday + timedelta(weeks=3, days=4)).replace(hour=15, minute=0),
               organizer_name="Lisa Braun",
               organizer_email="lisa.braun@example.com",
               location="Innovation Lab",
               body_preview="Demo completed stories to stakeholders."),
    )

    # Week 4-5
    events.append(
        _event("Quarterly Business Review",
               (monday + timedelta(weeks=4, days=1)).replace(hour=10, minute=0),
               (monday + timedelta(weeks=4, days=1)).replace(hour=12, minute=0),
               organizer_name="Heinrich Fischer",
               organizer_email="heinrich.fischer@example.com",
               location="Executive Boardroom",
               body_preview="Review key metrics, revenue targets, and strategic initiatives for the quarter."),
    )
    events.append(
        _event("Product Roadmap Alignment",
               (monday + timedelta(weeks=4, days=3)).replace(hour=15, minute=0),
               (monday + timedelta(weeks=4, days=3)).replace(hour=16, minute=30),
               organizer_name="Anna Schmidt",
               organizer_email="anna.schmidt@example.com",
               location="Meeting Room A",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/roadmap",
               body_preview="Align engineering and product on the upcoming quarter's roadmap."),
    )
    events.append(
        _event("External: Industry Conference Call",
               (monday + timedelta(weeks=5, days=0)).replace(hour=16, minute=0),
               (monday + timedelta(weeks=5, days=0)).replace(hour=17, minute=0),
               organizer_name="Petra Schneider",
               organizer_email="petra.schneider@example.com",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/conference",
               show_as="tentative",
               body_preview="Pre-conference coordination call with organizers."),
    )
    events.append(
        _event("Innovation Hackathon Kickoff",
               (monday + timedelta(weeks=5, days=2)).replace(hour=9, minute=0),
               (monday + timedelta(weeks=5, days=2)).replace(hour=10, minute=0),
               organizer_name="Lisa Braun",
               organizer_email="lisa.braun@example.com",
               location="Innovation Lab",
               body_preview="Hackathon theme announcement and team formation."),
    )

    # ── Out of Office events ──────────────────────────────────────────

    # Vacation days (all-day, OOF)
    vacation_start = monday + timedelta(weeks=3, days=0)  # a Monday
    for d in range(5):  # Mon-Fri vacation week
        vday = vacation_start + timedelta(days=d)
        events.append(
            _event("Vacation",
                   vday,
                   vday + timedelta(days=1),
                   is_all_day=True,
                   show_as="oof",
                   sensitivity="normal",
                   body_preview="Out of office — on vacation."),
        )

    # Doctor appointment (partial day OOF)
    events.append(
        _event("Doctor Appointment",
               (monday + timedelta(weeks=2, days=0)).replace(hour=8, minute=0),
               (monday + timedelta(weeks=2, days=0)).replace(hour=10, minute=0),
               show_as="oof",
               sensitivity="normal",
               body_preview="Personal appointment — will be back by 10:00."),
    )

    # Dentist appointment
    events.append(
        _event("Dentist",
               (monday + timedelta(weeks=4, days=4)).replace(hour=15, minute=0),
               (monday + timedelta(weeks=4, days=4)).replace(hour=16, minute=0),
               show_as="oof",
               sensitivity="normal",
               body_preview="Leaving early for dentist appointment."),
    )

    # ── All-day events ────────────────────────────────────────────────

    # Company Holiday
    next_friday = monday + timedelta(weeks=1, days=4)
    events.append(
        _event("Company Holiday",
               next_friday,
               next_friday + timedelta(days=1),
               is_all_day=True,
               organizer_name="HR Department",
               organizer_email="hr@example.com",
               body_preview="Office closed — public holiday."),
    )

    # Team Offsite
    offsite_day = monday + timedelta(weeks=5, days=3)
    events.append(
        _event("Team Offsite",
               offsite_day,
               offsite_day + timedelta(days=2),
               is_all_day=True,
               organizer_name="Klaus Weber",
               organizer_email="klaus.weber@example.com",
               location="Riverside Conference Center",
               body_preview="2-day team offsite: strategy, team building, and workshops."),
    )

    # Conference attendance
    conf_day = monday + timedelta(weeks=6, days=1)
    events.append(
        _event("Tech Summit 2026",
               conf_day,
               conf_day + timedelta(days=3),
               is_all_day=True,
               organizer_name="Petra Schneider",
               organizer_email="petra.schneider@example.com",
               location="Convention Center",
               show_as="oof",
               body_preview="Attending Tech Summit 2026 — booth duty and keynote sessions."),
    )

    # ── Tentative meetings ────────────────────────────────────────────
    events.append(
        _event("Potential Client Call — Newco",
               (monday + timedelta(weeks=2, days=2)).replace(hour=16, minute=0),
               (monday + timedelta(weeks=2, days=2)).replace(hour=16, minute=45),
               organizer_name="Heinrich Fischer",
               organizer_email="heinrich.fischer@example.com",
               online_meeting_url="https://teams.microsoft.com/l/meetup-join/newco-call",
               show_as="tentative",
               body_preview="Exploratory call with potential new client Newco."),
    )
    events.append(
        _event("Optional: Yoga Session",
               (monday + timedelta(weeks=1, days=3)).replace(hour=17, minute=0),
               (monday + timedelta(weeks=1, days=3)).replace(hour=18, minute=0),
               organizer_name="Lisa Braun",
               organizer_email="lisa.braun@example.com",
               location="Wellness Room",
               show_as="tentative",
               body_preview="Optional team yoga session — all welcome!"),
    )

    return events


# ── Mail ─────────────────────────────────────────────────────────

def _strip_html(html: str) -> str:
    """Naive HTML tag stripper for bodyPreview generation."""
    import re
    text = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL)
    text = re.sub(r"<[^>]+>", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&#\d+;", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


# ── Shared HTML helpers ─────────────────────────────────────────

_FONT = "'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Arial, sans-serif"

_HTML_HEAD = (
    "<html><head><style>"
    f"body {{ font-family: {_FONT}; font-size: 14px; line-height: 1.5; color: #333; "
    "margin: 0; padding: 20px; }}"
    "a { color: #0078d4; text-decoration: none; }"
    "a:hover { text-decoration: underline; }"
    "p { margin: 0 0 12px 0; }"
    "h1, h2, h3 { margin: 0 0 12px 0; }"
    "ul, ol { margin: 0 0 12px 0; padding-left: 24px; }"
    "li { margin-bottom: 4px; }"
    "code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; "
    f"font-family: 'Cascadia Code', Consolas, monospace; font-size: 13px; }}"
    "pre { background: #f6f8fa; padding: 16px; border-radius: 6px; "
    f"font-family: 'Cascadia Code', Consolas, monospace; font-size: 12px; "
    "overflow-x: auto; border: 1px solid #e1e4e8; }"
    "table { border-collapse: collapse; }"
    "td, th { padding: 8px 12px; }"
    "</style></head><body>"
)

_HTML_TAIL = "</body></html>"


def _sig(name: str, title: str, company: str, phone: str,
         color: str = "#0078d4", initials: str | None = None) -> str:
    """Generate a professional HTML email signature with optional logo div."""
    logo = ""
    if initials:
        logo = (
            f'<td style="vertical-align:top;padding-right:12px;">'
            f'<div style="width:40px;height:40px;border-radius:6px;background:{color};'
            f"color:#fff;font-weight:bold;font-size:16px;font-family:{_FONT};"
            f'text-align:center;line-height:40px;">{initials}</div></td>'
        )
    return (
        '<br><br>'
        '<table cellpadding="0" cellspacing="0" border="0" '
        f'style="font-family:{_FONT};font-size:12px;color:#666;">'
        f'<tr>{logo}'
        '<td style="vertical-align:top;">'
        f'<div style="font-size:13px;font-weight:600;color:#333;">{name}</div>'
        f'<div>{title}</div>'
        f'<div style="color:#888;">{company}</div>'
        f'<div style="color:#888;">{phone}</div>'
        '</td></tr></table>'
    )


def _fake_attachment(name: str, content_type: str, text_content: str = "") -> dict:
    """Create a fake attachment dict in MS Graph JSON shape."""
    if not text_content:
        text_content = f"This is a placeholder for {name}."
    content_bytes = base64.b64encode(text_content.encode()).decode()
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentType": content_type,
        "contentBytes": content_bytes,
        "size": len(text_content),
        "isInline": False,
    }


def default_mail_inbox(messages_per_folder: int = 5) -> list[dict]:
    """Synthetic mail messages across multiple folders in MS Graph JSON shape.

    Each message includes a ``_folder_id`` field so the transport can filter
    by folder.  Returns ~34 messages across Inbox, Sent Items, Drafts,
    Deleted Items, and Inbox sub-folders (Notifications, Done).
    """
    now = _now()

    def _msg(subject: str, sender_name: str, sender_email: str,
             body: str, folder_id: str = "inbox",
             minutes_ago: int = 0, is_read: bool = False,
             is_draft: bool = False, has_attachments: bool = False,
             importance: str = "normal") -> dict:
        received = now - timedelta(minutes=minutes_ago)
        ts = received.strftime("%Y-%m-%dT%H:%M:%SZ")
        # Auto-detect HTML content
        is_html = body.lstrip().startswith("<") or "<html" in body.lower()
        content_type = "html" if is_html else "text"
        preview = _strip_html(body)[:200] if is_html else body[:200]
        return {
            "id": _uid(),
            "subject": subject,
            "from": {"emailAddress": {"name": sender_name, "address": sender_email}},
            "toRecipients": [{"emailAddress": {"name": "Mock User", "address": "mock@example.com"}}],
            "receivedDateTime": ts,
            "sentDateTime": ts,
            "isRead": is_read,
            "isDraft": is_draft,
            "importance": importance,
            "hasAttachments": has_attachments,
            "bodyPreview": preview,
            "body": {"contentType": content_type, "content": body},
            "categories": [],
            "flag": {"flagStatus": "notFlagged"},
            "webLink": f"https://outlook.office.com/mail/mock/{_uid()}",
            "conversationId": _uid(),
            "_folder_id": folder_id,
        }

    messages: list[dict] = []

    # ── Inbox messages (18) ─────────────────────────────────

    # 1. CEO all-hands announcement (internal, HTML)
    messages.append(
        _msg(
            "All-Hands Meeting: Q2 Strategy & Company Update",
            "Heinrich Fischer", "heinrich.fischer@example.com",
            _HTML_HEAD
            + '<div style="border-left:4px solid #0078d4;padding-left:16px;margin-bottom:20px;">'
            + '<h2 style="color:#0078d4;margin:0;">All-Hands Meeting Announcement</h2>'
            + '</div>'
            + "<p>Dear Team,</p>"
            + "<p>I am pleased to invite you to our quarterly All-Hands meeting to discuss our "
            + "Q2 strategy, recent wins, and the road ahead.</p>"
            + '<div style="background:#f0f6ff;border-radius:8px;padding:16px;margin:16px 0;">'
            + '<p style="margin:0;"><strong>Date:</strong> Friday, April 10 at 14:00 CET</p>'
            + '<p style="margin:4px 0 0 0;"><strong>Location:</strong> Main Auditorium + Teams Live</p>'
            + '<p style="margin:4px 0 0 0;"><strong>Duration:</strong> 90 minutes</p>'
            + "</div>"
            + "<p><strong>Agenda highlights:</strong></p>"
            + "<ul>"
            + "<li>Q1 financial results and Q2 targets</li>"
            + "<li>Product roadmap update from Engineering</li>"
            + "<li>New partnership announcements</li>"
            + "<li>Employee recognition awards</li>"
            + "<li>Open Q&amp;A session</li>"
            + "</ul>"
            + "<p>Please submit your questions in advance via the shared form.</p>"
            + _sig("Heinrich Fischer", "Chief Executive Officer", "Example GmbH",
                   "+49 7123 100", color="#1a365d", initials="EG")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=35, is_read=False, importance="high",
        ),
    )

    # 2. Sprint review notes from Scrum Master (internal, HTML)
    messages.append(
        _msg(
            "Sprint 24 Review Notes & Action Items",
            "Lisa Braun", "lisa.braun@example.com",
            _HTML_HEAD
            + "<p>Hi team,</p>"
            + "<p>Great sprint review today! Here is a summary of what we covered:</p>"
            + '<h3 style="color:#0078d4;">Completed Stories</h3>'
            + "<ul>"
            + "<li><strong>ENG-1042:</strong> API Gateway rate limiting &mdash; <span style='color:#28a745;'>Done</span></li>"
            + "<li><strong>ENG-1038:</strong> Dashboard chart widgets &mdash; <span style='color:#28a745;'>Done</span></li>"
            + "<li><strong>ENG-1055:</strong> User profile caching &mdash; <span style='color:#28a745;'>Done</span></li>"
            + "</ul>"
            + '<h3 style="color:#e36209;">Carried Over</h3>'
            + "<ul>"
            + "<li><strong>ENG-1060:</strong> CI/CD staging environment (80% complete)</li>"
            + "<li><strong>ENG-1063:</strong> Search indexing optimization (blocked by infra)</li>"
            + "</ul>"
            + '<h3 style="color:#0078d4;">Action Items</h3>'
            + '<table style="width:100%;border:1px solid #e1e4e8;margin:12px 0;">'
            + '<tr style="background:#f6f8fa;"><th style="text-align:left;border-bottom:1px solid #e1e4e8;">Owner</th>'
            + '<th style="text-align:left;border-bottom:1px solid #e1e4e8;">Action</th>'
            + '<th style="text-align:left;border-bottom:1px solid #e1e4e8;">Due</th></tr>'
            + '<tr><td>Tobias N.</td><td>Finish staging pipeline config</td><td>Wed</td></tr>'
            + '<tr><td>Stefan H.</td><td>Unblock infra ticket with Ops</td><td>Thu</td></tr>'
            + '<tr><td>Sandra K.</td><td>Update component library docs</td><td>Fri</td></tr>'
            + "</table>"
            + "<p>Sprint 25 planning is Monday at 10:00. Please groom your backlog items before then.</p>"
            + _sig("Lisa Braun", "Scrum Master", "Example GmbH", "+49 170 5551234")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=90, is_read=True,
        ),
    )

    # 3. Customer escalation forward from Sales Manager (internal, importance=high)
    messages.append(
        _msg(
            "URGENT: Customer Escalation - Acme Corp Integration Failure",
            "Anna Schmidt", "anna.schmidt@example.com",
            _HTML_HEAD
            + '<div style="background:#fff3cd;border:1px solid #ffc107;border-radius:6px;padding:12px;margin-bottom:16px;">'
            + '<strong style="color:#856404;">Priority: High</strong> &mdash; Customer impact ongoing'
            + "</div>"
            + "<p>Hi team,</p>"
            + "<p>I just got off a call with Acme Corp. Their integration with our API has been failing "
            + "since this morning. They are unable to process orders and this is directly affecting "
            + "their revenue.</p>"
            + "<p><strong>Details:</strong></p>"
            + "<ul>"
            + "<li><strong>Customer:</strong> Acme Corp (Enterprise Tier, ARR: 240k)</li>"
            + "<li><strong>Issue:</strong> HTTP 503 errors on <code>/api/v2/orders</code> endpoint</li>"
            + "<li><strong>Since:</strong> ~08:30 CET today</li>"
            + "<li><strong>Impact:</strong> Complete order processing outage on their side</li>"
            + "</ul>"
            + "<p>Can someone from Engineering please investigate immediately? I have a follow-up "
            + "call with their CTO at 15:00.</p>"
            + _sig("Anna Schmidt", "Sales Manager", "Example GmbH", "+49 170 5559876")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=25, is_read=False, importance="high",
            has_attachments=True,
        ),
    )
    messages[-1]["attachments"] = [
        _fake_attachment(
            "escalation_timeline.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "Timestamp,Event,Status\n08:30,First 503 error detected,Open\n"
            "09:00,Customer notified support,Investigating\n"
            "09:45,Root cause identified: rate limiter config,In Progress\n",
        ),
    ]

    # 4. GitHub PR review notification (external, GitHub styling)
    messages.append(
        _msg(
            "[example-org/api-gateway] PR #347: Fix rate limiter token bucket overflow",
            "GitHub", "noreply@github.com",
            _HTML_HEAD
            + '<div style="background:#24292e;color:#fff;padding:16px 20px;border-radius:6px 6px 0 0;margin:-20px -20px 20px -20px;">'
            + '<span style="font-size:20px;font-weight:600;">Pull Request Review</span>'
            + "</div>"
            + '<p><a href="#">max.mustermann</a> requested your review on '
            + '<a href="#"><strong>#347 Fix rate limiter token bucket overflow</strong></a></p>'
            + '<div style="background:#f6f8fa;border:1px solid #e1e4e8;border-radius:6px;padding:12px;margin:12px 0;">'
            + "<p style='margin:0 0 8px 0;font-weight:600;'>Changes in this PR:</p>"
            + '<ul style="margin:0;">'
            + "<li><code>src/ratelimiter/bucket.py</code> &mdash; fix overflow on int32 boundary</li>"
            + "<li><code>tests/test_bucket.py</code> &mdash; add regression test</li>"
            + "<li><code>CHANGELOG.md</code> &mdash; document fix</li>"
            + "</ul>"
            + "</div>"
            + '<table style="margin:12px 0;">'
            + '<tr><td style="color:#28a745;font-weight:600;">+47</td>'
            + '<td style="color:#cb2431;font-weight:600;padding-left:8px;">-12</td>'
            + '<td style="padding-left:12px;color:#586069;">across 3 files</td></tr>'
            + "</table>"
            + '<p><a href="#" style="background:#28a745;color:#fff;padding:8px 20px;border-radius:6px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">Review changes</a></p>'
            + '<p style="color:#586069;font-size:12px;margin-top:20px;border-top:1px solid #e1e4e8;padding-top:12px;">'
            + "You are receiving this because you were requested as a reviewer.<br>"
            + '<a href="#">Unsubscribe</a> | <a href="#">View on GitHub</a></p>'
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=55, is_read=False,
        ),
    )

    # 5. Weekly tech newsletter (external, newsletter layout)
    kw = now.isocalendar()[1]
    messages.append(
        _msg(
            "TechCrunch Weekly: AI Agents, Cloud Costs & Developer Tools",
            "TechCrunch Weekly", "newsletter@techcrunch-weekly.com",
            _HTML_HEAD
            + '<div style="background:linear-gradient(135deg,#0a9b4a,#067a3a);color:#fff;padding:24px;'
            + 'border-radius:8px;text-align:center;margin-bottom:24px;">'
            + f'<h1 style="margin:0;font-size:24px;">TechCrunch Weekly &mdash; KW {kw}</h1>'
            + '<p style="margin:8px 0 0 0;opacity:0.9;">Your curated tech digest</p>'
            + "</div>"
            + '<div style="border-left:3px solid #0a9b4a;padding-left:16px;margin-bottom:20px;">'
            + '<h3 style="margin:0 0 8px 0;">AI Agents Are Reshaping Enterprise Software</h3>'
            + "<p>The latest wave of AI-powered agents is transforming how companies handle customer "
            + "support, code review, and data analysis. We look at the top 5 platforms leading the charge.</p>"
            + '<a href="#" style="color:#0a9b4a;font-weight:600;">Read more &rarr;</a>'
            + "</div>"
            + '<div style="border-left:3px solid #e36209;padding-left:16px;margin-bottom:20px;">'
            + '<h3 style="margin:0 0 8px 0;">Cloud Costs: Why FinOps Is the New DevOps</h3>'
            + "<p>As cloud spending spirals, engineering teams are adopting FinOps practices. "
            + "Here is a practical guide to getting started.</p>"
            + '<a href="#" style="color:#e36209;font-weight:600;">Read more &rarr;</a>'
            + "</div>"
            + '<div style="border-left:3px solid #0078d4;padding-left:16px;margin-bottom:20px;">'
            + '<h3 style="margin:0 0 8px 0;">5 Developer Tools You Should Try This Month</h3>'
            + "<p>From AI-assisted IDEs to next-gen database clients, our curated picks for April.</p>"
            + '<a href="#" style="color:#0078d4;font-weight:600;">Read more &rarr;</a>'
            + "</div>"
            + '<div style="text-align:center;margin-top:24px;padding-top:16px;border-top:1px solid #e1e4e8;">'
            + '<p style="color:#586069;font-size:12px;">You received this because you subscribed to TechCrunch Weekly.<br>'
            + '<a href="#" style="color:#586069;">Unsubscribe</a> | '
            + '<a href="#" style="color:#586069;">View in browser</a></p>'
            + "</div>"
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=180, is_read=True,
        ),
    )

    # 6. AWS billing alert (external, AWS styling)
    messages.append(
        _msg(
            "AWS Billing Alert: Monthly charges exceed $2,500 threshold",
            "Amazon Web Services", "billing@aws.amazon.com",
            _HTML_HEAD
            + '<div style="background:#232f3e;padding:16px 20px;margin:-20px -20px 20px -20px;">'
            + '<span style="color:#ff9900;font-size:20px;font-weight:700;">aws</span>'
            + "</div>"
            + '<div style="background:#fff8e1;border-left:4px solid #ff9900;padding:12px 16px;margin-bottom:16px;border-radius:0 4px 4px 0;">'
            + '<strong style="color:#b7791f;">Billing Alert:</strong> Your estimated charges have exceeded the configured threshold.'
            + "</div>"
            + '<table style="width:100%;border:1px solid #e1e4e8;margin:16px 0;">'
            + '<tr style="background:#f6f8fa;"><th style="text-align:left;border-bottom:1px solid #e1e4e8;">Service</th>'
            + '<th style="text-align:right;border-bottom:1px solid #e1e4e8;">MTD Cost</th></tr>'
            + '<tr><td>Amazon EC2</td><td style="text-align:right;">$1,247.83</td></tr>'
            + '<tr><td>Amazon RDS</td><td style="text-align:right;">$682.40</td></tr>'
            + '<tr><td>Amazon S3</td><td style="text-align:right;">$341.17</td></tr>'
            + '<tr><td>AWS Lambda</td><td style="text-align:right;">$189.52</td></tr>'
            + '<tr><td>Other Services</td><td style="text-align:right;">$156.08</td></tr>'
            + '<tr style="background:#f6f8fa;font-weight:700;"><td style="border-top:2px solid #232f3e;">Total Estimated</td>'
            + '<td style="text-align:right;border-top:2px solid #232f3e;">$2,617.00</td></tr>'
            + "</table>"
            + '<p>Threshold configured: <strong>$2,500.00</strong></p>'
            + '<p><a href="#" style="background:#ff9900;color:#232f3e;padding:8px 20px;border-radius:4px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">View Billing Dashboard</a></p>'
            + '<p style="color:#586069;font-size:12px;margin-top:20px;border-top:1px solid #e1e4e8;padding-top:12px;">'
            + "Amazon Web Services, Inc. is a subsidiary of Amazon.com, Inc.<br>"
            + "This is an automated notification from your AWS account.</p>"
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=240, is_read=True, has_attachments=True,
        ),
    )
    messages[-1]["attachments"] = [
        _fake_attachment(
            "invoice_march_2026.pdf",
            "application/pdf",
            "AWS Invoice - March 2026\n"
            "Account: 1234-5678-9012\n"
            "Total: $2,617.00\n"
            "EC2: $1,247.83 | RDS: $682.40 | S3: $341.17 | Lambda: $189.52 | Other: $156.08\n",
        ),
    ]

    # 7. Meeting follow-up with action items (internal, HTML table)
    messages.append(
        _msg(
            "Meeting Follow-up: Project Alpha Architecture Review",
            "Klaus Weber", "klaus.weber@example.com",
            _HTML_HEAD
            + "<p>Team,</p>"
            + "<p>Thanks for the productive architecture review today. Below are the decisions made "
            + "and action items assigned:</p>"
            + '<h3 style="color:#0078d4;">Decisions</h3>'
            + "<ol>"
            + "<li>We will migrate to event-driven architecture for the order pipeline</li>"
            + "<li>PostgreSQL stays as primary DB; Redis for caching layer</li>"
            + "<li>API versioning via URL path (<code>/v2/</code>), not headers</li>"
            + "</ol>"
            + '<h3 style="color:#0078d4;">Action Items</h3>'
            + '<table style="width:100%;border:1px solid #e1e4e8;margin:12px 0;">'
            + '<tr style="background:#0078d4;color:#fff;">'
            + "<th style='text-align:left;'>Who</th><th style='text-align:left;'>What</th>"
            + "<th style='text-align:left;'>Deadline</th><th style='text-align:left;'>Status</th></tr>"
            + '<tr><td>Tobias Neumann</td><td>Draft event schema for order events</td><td>Apr 7</td>'
            + '<td><span style="color:#e36209;">Pending</span></td></tr>'
            + '<tr style="background:#f6f8fa;"><td>Stefan Hoffmann</td><td>Set up Kafka cluster in staging</td><td>Apr 9</td>'
            + '<td><span style="color:#e36209;">Pending</span></td></tr>'
            + '<tr><td>Sandra Koch</td><td>Update frontend to consume v2 endpoints</td><td>Apr 11</td>'
            + '<td><span style="color:#e36209;">Pending</span></td></tr>'
            + '<tr style="background:#f6f8fa;"><td>Max Mustermann</td><td>Write migration guide for v1 consumers</td><td>Apr 14</td>'
            + '<td><span style="color:#e36209;">Pending</span></td></tr>'
            + "</table>"
            + "<p>Next review meeting is in two weeks. Please update your items in Jira.</p>"
            + _sig("Klaus Weber", "VP Technology", "Example GmbH",
                   "+49 7123 200", color="#1a365d", initials="EG")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=320, is_read=True,
            has_attachments=True,
        ),
    )
    messages[-1]["attachments"] = [
        _fake_attachment(
            "action_items.docx",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "Project Alpha Architecture Review - Action Items\n\n"
            "1. Tobias Neumann - Draft event schema for order events (Apr 7)\n"
            "2. Stefan Hoffmann - Set up Kafka cluster in staging (Apr 9)\n"
            "3. Sandra Koch - Update frontend to consume v2 endpoints (Apr 11)\n"
            "4. Max Mustermann - Write migration guide for v1 consumers (Apr 14)\n",
        ),
    ]

    # 8. HR benefits enrollment reminder (internal, styled)
    messages.append(
        _msg(
            "Reminder: Benefits Enrollment Window Closes April 15",
            "Christine Wagner", "christine.wagner@example.com",
            _HTML_HEAD
            + '<div style="background:linear-gradient(135deg,#6b46c1,#805ad5);color:#fff;padding:20px;border-radius:8px;'
            + 'text-align:center;margin-bottom:20px;">'
            + '<h2 style="margin:0;">Benefits Enrollment 2026</h2>'
            + '<p style="margin:8px 0 0 0;opacity:0.9;">Open enrollment closes in 13 days</p>'
            + "</div>"
            + "<p>Dear colleagues,</p>"
            + "<p>This is a friendly reminder that the annual benefits enrollment window closes on "
            + "<strong>April 15, 2026</strong>. If you do not make any changes, your current elections "
            + "will roll over automatically.</p>"
            + '<h3 style="color:#6b46c1;">Available Plans</h3>'
            + "<ul>"
            + "<li><strong>Health Insurance:</strong> Basic, Plus, Premium tiers</li>"
            + "<li><strong>Dental &amp; Vision:</strong> New enhanced coverage option</li>"
            + "<li><strong>Retirement:</strong> Company matches up to 6% of salary</li>"
            + "<li><strong>Wellness:</strong> Gym subsidy increased to 50/month</li>"
            + "</ul>"
            + '<p><a href="#" style="background:#6b46c1;color:#fff;padding:10px 24px;border-radius:6px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">Review Your Benefits</a></p>'
            + "<p>If you have questions, please reach out to the HR team or book an office hour slot.</p>"
            + _sig("Christine Wagner", "HR Business Partner", "Example GmbH", "+49 7123 310")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=480, is_read=True,
        ),
    )

    # 9. External partner proposal (external, professional)
    messages.append(
        _msg(
            "Partnership Proposal: Joint Cloud Migration Workshop",
            "Julia Meyer", "julia.meyer@partnercompany.de",
            _HTML_HEAD
            + "<p>Dear team at Example GmbH,</p>"
            + "<p>Following our conversation at the Stuttgart Tech Conference last week, I would like "
            + "to formally propose a joint workshop on cloud migration strategies for mid-market enterprises.</p>"
            + '<h3 style="color:#0078d4;">Proposed Format</h3>'
            + '<table style="width:100%;border:1px solid #e1e4e8;margin:12px 0;">'
            + '<tr><td style="width:120px;font-weight:600;background:#f6f8fa;">Format</td>'
            + "<td>Half-day workshop (4 hours)</td></tr>"
            + '<tr><td style="font-weight:600;background:#f6f8fa;">Target Audience</td>'
            + "<td>CTOs and Engineering Leads from DACH region</td></tr>"
            + '<tr><td style="font-weight:600;background:#f6f8fa;">Location</td>'
            + "<td>Stuttgart or virtual (hybrid option)</td></tr>"
            + '<tr><td style="font-weight:600;background:#f6f8fa;">Proposed Date</td>'
            + "<td>Late May or early June 2026</td></tr>"
            + '<tr><td style="font-weight:600;background:#f6f8fa;">Investment</td>'
            + "<td>Shared costs, estimated 3,500 per company</td></tr>"
            + "</table>"
            + "<p>I have attached a detailed proposal document with agenda, speaker suggestions, "
            + "and budget breakdown. I would love to schedule a 30-minute call to discuss further.</p>"
            + "<p>Looking forward to your thoughts!</p>"
            + _sig("Julia Meyer", "Head of Partnerships", "Partner Company GmbH",
                   "+49 711 9988770", color="#2d5f2d", initials="PC")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=600, is_read=True, has_attachments=True,
        ),
    )
    messages[-1]["attachments"] = [
        _fake_attachment(
            "Q1_Sales_Report_2026.pdf",
            "application/pdf",
            "Q1 Sales Report 2026 - Example GmbH\n\n"
            "Revenue: 1,850,000 EUR (+18% YoY)\n"
            "New Customers: 34\n"
            "Churn Rate: 2.1%\n"
            "Top Segment: Enterprise (62%)\n",
        ),
        _fake_attachment(
            "IT_Security_Policy_v3.pdf",
            "application/pdf",
            "IT Security Policy v3.0 - Example GmbH\n\n"
            "1. Password Requirements: Minimum 12 characters, MFA required\n"
            "2. Data Classification: Public, Internal, Confidential, Restricted\n"
            "3. Incident Response: Report within 4 hours to security@example.com\n",
        ),
    ]

    # 10. Jira ticket assignment notification (external, Jira styling)
    messages.append(
        _msg(
            "[JIRA] (ENG-1072) Assigned to you: Implement webhook retry logic",
            "Jira", "support@atlassian.com",
            _HTML_HEAD
            + '<div style="background:#0052cc;padding:12px 20px;margin:-20px -20px 20px -20px;">'
            + '<span style="color:#fff;font-size:18px;font-weight:600;">'
            + '<span style="background:#fff;color:#0052cc;padding:2px 6px;border-radius:3px;'
            + 'font-size:12px;margin-right:8px;">JIRA</span>Ticket Assigned</span>'
            + "</div>"
            + '<p>Frank Zimmermann assigned <strong><a href="#">ENG-1072</a></strong> to you:</p>'
            + '<div style="background:#f4f5f7;border-radius:6px;padding:16px;margin:12px 0;">'
            + '<p style="margin:0 0 8px 0;font-size:16px;font-weight:600;">'
            + '<a href="#">ENG-1072: Implement webhook retry logic with exponential backoff</a></p>'
            + '<table style="font-size:13px;">'
            + '<tr><td style="color:#5e6c84;width:80px;">Type:</td><td>Story</td></tr>'
            + '<tr><td style="color:#5e6c84;">Priority:</td>'
            + '<td><span style="color:#ff5630;">High</span></td></tr>'
            + '<tr><td style="color:#5e6c84;">Sprint:</td><td>Sprint 25</td></tr>'
            + '<tr><td style="color:#5e6c84;">Points:</td><td>5</td></tr>'
            + "</table>"
            + "</div>"
            + "<p><strong>Description:</strong></p>"
            + "<p>Implement retry logic for outgoing webhooks. Failed deliveries should be retried "
            + "with exponential backoff (1s, 2s, 4s, 8s, max 60s). After 10 consecutive failures, "
            + "disable the webhook and notify the owner.</p>"
            + '<p><a href="#" style="background:#0052cc;color:#fff;padding:6px 16px;border-radius:3px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">View in Jira</a></p>'
            + '<p style="color:#5e6c84;font-size:11px;margin-top:16px;border-top:1px solid #e1e4e8;padding-top:12px;">'
            + "This message was sent by Atlassian Jira. "
            + '<a href="#">Manage notifications</a></p>'
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=110, is_read=False,
        ),
    )

    # 11. Figma comment notification (external)
    messages.append(
        _msg(
            "[Figma] Katrin Schwarz commented on 'Dashboard v2 Mockups'",
            "Figma", "team@figma.com",
            _HTML_HEAD
            + '<div style="background:#1e1e1e;padding:16px 20px;margin:-20px -20px 20px -20px;">'
            + '<span style="color:#fff;font-size:18px;font-weight:600;">'
            + '<span style="color:#a259ff;margin-right:6px;">&#9670;</span>Figma</span>'
            + "</div>"
            + '<p><strong>Katrin Schwarz</strong> left a comment on '
            + '<a href="#" style="font-weight:600;">Dashboard v2 Mockups</a>:</p>'
            + '<div style="background:#f7f7f7;border-left:3px solid #a259ff;padding:12px 16px;'
            + 'margin:12px 0;border-radius:0 6px 6px 0;">'
            + '<p style="margin:0;font-style:italic;">"I updated the chart color palette to match '
            + "our new brand guidelines. Can you check if the contrast ratios still pass WCAG AA? "
            + 'The sidebar gradient might need adjustment too."</p>'
            + "</div>"
            + '<p><a href="#" style="background:#a259ff;color:#fff;padding:8px 20px;border-radius:6px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">Open in Figma</a></p>'
            + '<p style="color:#586069;font-size:12px;margin-top:20px;border-top:1px solid #e1e4e8;padding-top:12px;">'
            + 'You are receiving this because you are a collaborator on this file. '
            + '<a href="#">Notification settings</a></p>'
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=150, is_read=True,
        ),
    )

    # 12. Casual lunch invite (short text, internal)
    messages.append(
        _msg(
            "Lunch today?",
            "Tobias Neumann", "tobias.neumann@example.com",
            "Hey! A few of us are going to the new Thai place on Schlossstrasse. "
            "Want to join? We're heading out at 12:15.\n\n"
            "- Tobias",
            folder_id="inbox", minutes_ago=40, is_read=False,
        ),
    )

    # 13. IT maintenance notice in German (internal, HTML)
    messages.append(
        _msg(
            "Geplante Wartung: VPN & Netzwerkinfrastruktur am Samstag",
            "Stefan Hoffmann", "stefan.hoffmann@example.com",
            _HTML_HEAD
            + '<div style="background:#d63384;color:#fff;padding:16px 20px;border-radius:8px;margin-bottom:16px;">'
            + '<h2 style="margin:0;font-size:18px;">&#9888; Geplante Wartungsarbeiten</h2>'
            + "</div>"
            + "<p>Hallo zusammen,</p>"
            + "<p>am kommenden <strong>Samstag, 5. April</strong>, finden geplante Wartungsarbeiten "
            + "an unserer Netzwerkinfrastruktur statt.</p>"
            + '<div style="background:#f8d7da;border:1px solid #f5c6cb;border-radius:6px;padding:12px 16px;margin:12px 0;">'
            + '<p style="margin:0;"><strong>Zeitfenster:</strong> 06:00 - 10:00 Uhr (MESZ)</p>'
            + '<p style="margin:4px 0 0 0;"><strong>Betroffene Systeme:</strong> VPN, internes WLAN, Druckserver</p>'
            + "</div>"
            + "<p><strong>Wichtige Hinweise:</strong></p>"
            + "<ul>"
            + "<li>Der VPN-Zugang wird in diesem Zeitraum <strong>nicht verfügbar</strong> sein</li>"
            + "<li>Bitte speichert alle offenen Arbeiten vor Freitagabend lokal ab</li>"
            + "<li>Das WLAN im Erdgeschoss ist nicht betroffen</li>"
            + "<li>Nach Abschluss der Wartung wird eine Bestätigung versendet</li>"
            + "</ul>"
            + "<p>Bei Fragen wendet euch bitte an das IT-Support-Team.</p>"
            + _sig("Stefan Hoffmann", "DevOps Engineer", "Example GmbH", "+49 7123 450")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=520, is_read=True,
        ),
    )

    # 14. Quarterly OKR review (internal, importance=high)
    messages.append(
        _msg(
            "Q1 OKR Review: Results & Q2 Planning",
            "Heinrich Fischer", "heinrich.fischer@example.com",
            _HTML_HEAD
            + '<div style="border-left:4px solid #e53e3e;padding-left:16px;margin-bottom:20px;">'
            + '<h2 style="color:#e53e3e;margin:0;">Q1 OKR Review &mdash; Action Required</h2>'
            + '<p style="color:#718096;margin:4px 0 0 0;">Please complete your self-assessment by Friday</p>'
            + "</div>"
            + "<p>Team,</p>"
            + "<p>Q1 has wrapped up and it is time to review our OKR progress. Below is the "
            + "company-level summary:</p>"
            + '<table style="width:100%;border:1px solid #e1e4e8;margin:16px 0;">'
            + '<tr style="background:#2d3748;color:#fff;">'
            + "<th style='text-align:left;'>Objective</th><th style='text-align:center;width:100px;'>Target</th>"
            + "<th style='text-align:center;width:100px;'>Actual</th>"
            + "<th style='text-align:center;width:80px;'>Score</th></tr>"
            + '<tr><td>Revenue Growth</td><td style="text-align:center;">+15%</td>'
            + '<td style="text-align:center;">+18%</td>'
            + '<td style="text-align:center;color:#28a745;font-weight:700;">1.0</td></tr>'
            + '<tr style="background:#f6f8fa;"><td>Customer Satisfaction (NPS)</td><td style="text-align:center;">50</td>'
            + '<td style="text-align:center;">54</td>'
            + '<td style="text-align:center;color:#28a745;font-weight:700;">0.9</td></tr>'
            + '<tr><td>Platform Uptime</td><td style="text-align:center;">99.9%</td>'
            + '<td style="text-align:center;">99.7%</td>'
            + '<td style="text-align:center;color:#e36209;font-weight:700;">0.7</td></tr>'
            + '<tr style="background:#f6f8fa;"><td>New Feature Releases</td><td style="text-align:center;">8</td>'
            + '<td style="text-align:center;">6</td>'
            + '<td style="text-align:center;color:#e36209;font-weight:700;">0.6</td></tr>'
            + "</table>"
            + "<p><strong>Next steps:</strong></p>"
            + "<ol>"
            + "<li>Complete your team-level OKR self-assessment in the shared spreadsheet</li>"
            + "<li>Schedule 1:1 review with your manager by end of next week</li>"
            + "<li>Draft Q2 OKR proposals for discussion at the leadership offsite</li>"
            + "</ol>"
            + _sig("Heinrich Fischer", "Chief Executive Officer", "Example GmbH",
                   "+49 7123 100", color="#1a365d", initials="EG")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=720, is_read=True, importance="high",
        ),
    )

    # 15. Local meetup invitation (external, event styling)
    messages.append(
        _msg(
            "Stuttgart Tech Meetup: Building Scalable APIs with Python",
            "Stuttgart Tech Events", "events@meetup-stuttgart.de",
            _HTML_HEAD
            + '<div style="background:linear-gradient(135deg,#ed4245,#f57242);color:#fff;padding:24px;'
            + 'border-radius:8px;text-align:center;margin-bottom:20px;">'
            + '<p style="margin:0;font-size:12px;text-transform:uppercase;letter-spacing:2px;opacity:0.8;">Stuttgart Tech Meetup</p>'
            + '<h1 style="margin:8px 0;font-size:22px;">Building Scalable APIs with Python</h1>'
            + '<p style="margin:0;font-size:16px;">Donnerstag, 17. April 2026 | 18:30 Uhr</p>'
            + "</div>"
            + '<div style="background:#f6f8fa;border-radius:6px;padding:16px;margin-bottom:16px;">'
            + '<p style="margin:0 0 8px 0;font-weight:600;">Event Details:</p>'
            + '<table style="font-size:13px;">'
            + "<tr><td style='width:80px;color:#586069;'>Ort:</td><td>Coworking Space Hub, Stuttgarter Str. 42</td></tr>"
            + "<tr><td style='color:#586069;'>Sprache:</td><td>English &amp; Deutsch</td></tr>"
            + "<tr><td style='color:#586069;'>Kosten:</td><td>Free (Sponsored by TechHub e.V.)</td></tr>"
            + "</table>"
            + "</div>"
            + '<h3 style="color:#ed4245;">Agenda</h3>'
            + "<ul>"
            + "<li><strong>18:30</strong> &mdash; Doors open, networking &amp; pizza</li>"
            + "<li><strong>19:00</strong> &mdash; Talk: FastAPI at Scale &mdash; Lessons from Production</li>"
            + "<li><strong>19:45</strong> &mdash; Lightning talks (5 min each, sign up on the night)</li>"
            + "<li><strong>20:15</strong> &mdash; Open discussion &amp; drinks</li>"
            + "</ul>"
            + '<p style="text-align:center;"><a href="#" style="background:#ed4245;color:#fff;padding:10px 28px;'
            + 'border-radius:6px;text-decoration:none;font-weight:600;display:inline-block;">RSVP Now</a></p>'
            + '<p style="color:#586069;font-size:12px;text-align:center;margin-top:20px;border-top:1px solid #e1e4e8;padding-top:12px;">'
            + 'Stuttgart Tech Events | <a href="#">Unsubscribe</a></p>'
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=960, is_read=True,
        ),
    )

    # 16. Vacation approval from HR (short, internal)
    messages.append(
        _msg(
            "Re: Vacation Request April 21-25 - Approved",
            "Petra Schneider", "petra.schneider@example.com",
            _HTML_HEAD
            + "<p>Hi,</p>"
            + "<p>Your vacation request for <strong>April 21-25, 2026</strong> has been approved. "
            + "Enjoy your time off!</p>"
            + "<p>Please ensure your out-of-office reply is set and any handover tasks are "
            + "documented before you leave.</p>"
            + _sig("Petra Schneider", "VP Human Resources", "Example GmbH", "+49 7123 300")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=1100, is_read=True,
        ),
    )

    # 17. Code review feedback (internal, HTML with code block)
    messages.append(
        _msg(
            "Code Review: auth_middleware.py - Token Refresh Logic",
            "Max Mustermann", "max.mustermann@example.com",
            _HTML_HEAD
            + "<p>Hey,</p>"
            + "<p>I reviewed your changes to the token refresh logic in <code>auth_middleware.py</code>. "
            + "Overall looks good, but I have a couple of suggestions:</p>"
            + '<h3 style="color:#0078d4;">1. Race Condition in Token Refresh</h3>'
            + "<p>The current implementation could lead to multiple simultaneous refresh requests. "
            + "Consider adding a lock:</p>"
            + "<pre>"
            + "import asyncio\n\n"
            + "_refresh_lock = asyncio.Lock()\n\n"
            + "async def refresh_token(session):\n"
            + "    async with _refresh_lock:\n"
            + "        # Check again inside the lock in case another\n"
            + "        # coroutine already refreshed it\n"
            + "        if not session.token_expired:\n"
            + "            return session.access_token\n"
            + "        new_token = await _do_refresh(session)\n"
            + "        return new_token"
            + "</pre>"
            + '<h3 style="color:#0078d4;">2. Error Handling</h3>'
            + '<p>The <code>except Exception</code> on line 47 is too broad. Let us catch '
            + "<code>TokenRefreshError</code> specifically and let other exceptions propagate:</p>"
            + "<pre>"
            + "try:\n"
            + "    token = await refresh_token(session)\n"
            + "except TokenRefreshError as exc:\n"
            + "    logger.warning(\"Token refresh failed: %s\", exc)\n"
            + "    raise AuthenticationError(\"Session expired\") from exc"
            + "</pre>"
            + "<p>Other than that, the retry backoff logic is clean. Ship it once those are addressed!</p>"
            + _sig("Max Mustermann", "Senior Software Engineer", "Example GmbH", "+49 170 5551111")
            + _HTML_TAIL,
            folder_id="inbox", minutes_ago=200, is_read=True,
        ),
    )

    # 18. Personal thank you note (short text, internal)
    messages.append(
        _msg(
            "Thanks for the help yesterday!",
            "Sandra Koch", "sandra.koch@example.com",
            "Hey,\n\n"
            "Just wanted to say thanks for helping me debug that CSS grid issue yesterday. "
            "The dashboard layout is working perfectly now and the client was really impressed "
            "with how fast we turned it around.\n\n"
            "I owe you a coffee! :)\n\n"
            "Cheers,\nSandra",
            folder_id="inbox", minutes_ago=130, is_read=True,
        ),
    )

    # ── Sent Items (5) ──────────────────────────────────────

    messages.append(
        _msg(
            "Re: URGENT: Customer Escalation - Acme Corp Integration Failure",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Anna,</p>"
            + "<p>I am looking into this right now. The 503 errors appear to be caused by the rate "
            + "limiter deployment from this morning. I have identified the issue and am rolling back "
            + "the config change.</p>"
            + "<p>ETA for resolution: ~30 minutes. I will update the incident channel.</p>"
            + _sig("Mock User", "Software Engineer", "Example GmbH", "+49 170 5550000")
            + _HTML_TAIL,
            folder_id="sentitems", minutes_ago=18, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Re: Sprint 24 Review Notes & Action Items",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Thanks Lisa, great summary!</p>"
            + "<p>I will have the staging pipeline config ready by Wednesday. Already started "
            + "working on it this afternoon.</p>"
            + _sig("Mock User", "Software Engineer", "Example GmbH", "+49 170 5550000")
            + _HTML_TAIL,
            folder_id="sentitems", minutes_ago=80, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Re: Lunch today?",
            "Mock User", "mock@example.com",
            "Count me in! See you in the lobby at 12:15.",
            folder_id="sentitems", minutes_ago=35, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Fw: Partnership Proposal - Cloud Migration Workshop",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Klaus,</p>"
            + "<p>Forwarding this proposal from Partner Company for your review. "
            + "I think it could be a good fit for our Q3 marketing push. The budget "
            + "seems reasonable and the format aligns with what we discussed at the offsite.</p>"
            + "<p>Let me know your thoughts when you get a chance.</p>"
            + _sig("Mock User", "Software Engineer", "Example GmbH", "+49 170 5550000")
            + _HTML_TAIL,
            folder_id="sentitems", minutes_ago=550, is_read=True, has_attachments=True,
        ),
    )
    messages[-1]["attachments"] = [
        _fake_attachment(
            "partnership_proposal_cloud_migration.pdf",
            "application/pdf",
            "Joint Cloud Migration Workshop Proposal\n\n"
            "Format: Half-day workshop (4 hours)\n"
            "Target: CTOs and Engineering Leads, DACH region\n"
            "Location: Stuttgart or virtual\n"
            "Proposed Date: Late May / early June 2026\n"
            "Investment: 3,500 EUR per company\n",
        ),
    ]

    messages.append(
        _msg(
            "Re: Q1 OKR Review: Results & Q2 Planning",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Heinrich,</p>"
            + "<p>I have completed my self-assessment and updated the shared spreadsheet. "
            + "Our team's key highlights:</p>"
            + "<ul>"
            + "<li>API response time improvement: target 200ms, achieved 185ms</li>"
            + "<li>Test coverage: target 85%, achieved 88%</li>"
            + "<li>Deployment frequency: target weekly, achieved twice weekly</li>"
            + "</ul>"
            + "<p>Happy to discuss Q2 proposals in our 1:1 on Thursday.</p>"
            + _sig("Mock User", "Software Engineer", "Example GmbH", "+49 170 5550000")
            + _HTML_TAIL,
            folder_id="sentitems", minutes_ago=700, is_read=True,
        ),
    )

    # ── Drafts (3) ──────────────────────────────────────────

    messages.append(
        _msg(
            "Re: Code Review: auth_middleware.py - Token Refresh Logic",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Max,</p>"
            + "<p>Good catches, both of them. I have already fixed the race condition with the "
            + "asyncio lock approach you suggested. For the exception handling, I am thinking "
            + "we should also add</p>"
            + _HTML_TAIL,
            folder_id="drafts", minutes_ago=120, is_draft=True,
        ),
    )

    messages.append(
        _msg(
            "Cloud Migration Workshop - Initial Thoughts",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Dear Ms. Meyer,</p>"
            + "<p>Thank you for the detailed proposal regarding the joint cloud migration workshop. "
            + "We have reviewed it internally and are very interested in moving forward.</p>"
            + "<p>A few points we would like to discuss:</p>"
            + "<ul>"
            + "<li>We would prefer a hybrid format with strong virtual participation support</li>"
            + "<li>We can contribute two speakers from our Engineering and DevOps teams</li>"
            + "</ul>"
            + _HTML_TAIL,
            folder_id="drafts", minutes_ago=400, is_draft=True,
        ),
    )

    messages.append(
        _msg(
            "Engineering Team Update - Week of April 6",
            "Mock User", "mock@example.com",
            _HTML_HEAD
            + "<p>Hi everyone,</p>"
            + "<p>Here is the weekly engineering update:</p>"
            + '<h3 style="color:#0078d4;">Completed</h3>'
            + "<ul>"
            + "<li>Rate limiter token bucket fix (PR #347 merged)</li>"
            + "<li>Dashboard v2 mockups approved by design</li>"
            + "</ul>"
            + '<h3 style="color:#e36209;">In Progress</h3>'
            + "<ul>"
            + "<li>Webhook retry logic implementation</li>"
            + "</ul>"
            + _HTML_TAIL,
            folder_id="drafts", minutes_ago=60, is_draft=True,
        ),
    )

    # ── Inbox > Notifications subfolder (3) ─────────────────

    messages.append(
        _msg(
            "[GitHub] CI passed: example-org/api-gateway (main)",
            "GitHub Actions", "noreply@github.com",
            _HTML_HEAD
            + '<div style="background:#24292e;color:#fff;padding:12px 20px;margin:-20px -20px 16px -20px;">'
            + '<span style="font-size:16px;font-weight:600;">GitHub Actions</span>'
            + "</div>"
            + '<p><span style="color:#28a745;font-weight:700;">&#10003; All checks passed</span> for '
            + "<strong>main</strong> branch</p>"
            + '<div style="background:#f6f8fa;border:1px solid #e1e4e8;border-radius:6px;padding:12px;margin:12px 0;">'
            + '<table style="font-size:13px;width:100%;">'
            + '<tr><td style="color:#28a745;">&#10003;</td><td>build (ubuntu-latest, 3.11)</td>'
            + '<td style="color:#586069;text-align:right;">2m 14s</td></tr>'
            + '<tr><td style="color:#28a745;">&#10003;</td><td>lint (ruff, mypy)</td>'
            + '<td style="color:#586069;text-align:right;">47s</td></tr>'
            + '<tr><td style="color:#28a745;">&#10003;</td><td>test (pytest, coverage 91%)</td>'
            + '<td style="color:#586069;text-align:right;">3m 02s</td></tr>'
            + "</table>"
            + "</div>"
            + '<p style="color:#586069;font-size:12px;">Triggered by push to main (a1b2c3d) by max.mustermann</p>'
            + _HTML_TAIL,
            folder_id="inbox-notifications", minutes_ago=50, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "[JIRA] ENG-1060 status changed: In Progress -> In Review",
            "Jira", "support@atlassian.com",
            _HTML_HEAD
            + '<div style="background:#0052cc;padding:10px 16px;margin:-20px -20px 16px -20px;">'
            + '<span style="color:#fff;font-size:14px;font-weight:600;">'
            + '<span style="background:#fff;color:#0052cc;padding:2px 6px;border-radius:3px;'
            + 'font-size:11px;margin-right:6px;">JIRA</span>Status Update</span>'
            + "</div>"
            + "<p><strong>Tobias Neumann</strong> changed the status of "
            + '<a href="#"><strong>ENG-1060</strong></a>:</p>'
            + '<p><span style="background:#0052cc;color:#fff;padding:2px 8px;border-radius:3px;font-size:12px;">'
            + "In Progress</span> &rarr; "
            + '<span style="background:#00875a;color:#fff;padding:2px 8px;border-radius:3px;font-size:12px;">'
            + "In Review</span></p>"
            + '<p style="color:#5e6c84;font-size:13px;">CI/CD staging environment - ready for code review</p>'
            + _HTML_TAIL,
            folder_id="inbox-notifications", minutes_ago=95, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Calendar Reminder: Sprint 25 Planning tomorrow at 10:00",
            "Microsoft Outlook", "noreply@microsoft.com",
            _HTML_HEAD
            + '<div style="background:#0078d4;padding:12px 20px;margin:-20px -20px 16px -20px;">'
            + '<span style="color:#fff;font-size:16px;font-weight:600;">&#128197; Calendar Reminder</span>'
            + "</div>"
            + '<div style="background:#f0f6ff;border-radius:6px;padding:16px;margin:12px 0;">'
            + '<p style="margin:0 0 8px 0;font-size:16px;font-weight:600;">Sprint 25 Planning</p>'
            + '<table style="font-size:13px;">'
            + "<tr><td style='width:70px;color:#586069;'>When:</td><td>Monday, April 6 at 10:00 - 11:30</td></tr>"
            + "<tr><td style='color:#586069;'>Where:</td><td>Conference Room A + Teams</td></tr>"
            + "<tr><td style='color:#586069;'>Organizer:</td><td>Lisa Braun</td></tr>"
            + "</table>"
            + "</div>"
            + '<p><a href="#">Accept</a> | <a href="#">Tentative</a> | <a href="#">Decline</a></p>'
            + _HTML_TAIL,
            folder_id="inbox-notifications", minutes_ago=140, is_read=True,
        ),
    )

    # ── Inbox > Done subfolder (2) ──────────────────────────

    messages.append(
        _msg(
            "Re: Office Key Card Replacement - Done",
            "Melanie Schreiber", "melanie.schreiber@example.com",
            _HTML_HEAD
            + "<p>Hi,</p>"
            + "<p>Your replacement key card is ready for pickup at the IT help desk on the "
            + "ground floor. Please bring your employee ID for verification.</p>"
            + "<p>Opening hours: Monday - Friday, 08:00 - 17:00</p>"
            + _sig("Melanie Schreiber", "IT Support Specialist", "Example GmbH", "+49 7123 460")
            + _HTML_TAIL,
            folder_id="inbox-done", minutes_ago=2000, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Re: Expense Report March 2026 - Approved & Processed",
            "Anja Beyer", "anja.beyer@example.com",
            _HTML_HEAD
            + "<p>Hi,</p>"
            + "<p>Your expense report for March 2026 has been approved and processed. "
            + "The reimbursement of <strong>347.50 EUR</strong> will be included in your "
            + "next payroll cycle (April 30).</p>"
            + '<table style="width:100%;border:1px solid #e1e4e8;margin:12px 0;font-size:13px;">'
            + '<tr style="background:#f6f8fa;"><th style="text-align:left;">Item</th>'
            + '<th style="text-align:right;">Amount</th></tr>'
            + "<tr><td>Train ticket Stuttgart-Munich</td><td style='text-align:right;'>89.00</td></tr>"
            + "<tr><td>Hotel (1 night)</td><td style='text-align:right;'>145.00</td></tr>"
            + "<tr><td>Client dinner</td><td style='text-align:right;'>78.50</td></tr>"
            + "<tr><td>Taxi</td><td style='text-align:right;'>35.00</td></tr>"
            + '<tr style="background:#f6f8fa;font-weight:700;border-top:2px solid #333;">'
            + "<td>Total</td><td style='text-align:right;'>347.50 EUR</td></tr>"
            + "</table>"
            + _sig("Anja Beyer", "Financial Controller", "Example GmbH", "+49 7123 510")
            + _HTML_TAIL,
            folder_id="inbox-done", minutes_ago=3000, is_read=True,
        ),
    )

    # ── Deleted Items (3) ───────────────────────────────────

    messages.append(
        _msg(
            "TechCrunch Weekly: Last Week in AI, Quantum & Startups",
            "TechCrunch Weekly", "newsletter@techcrunch-weekly.com",
            _HTML_HEAD
            + '<div style="background:linear-gradient(135deg,#0a9b4a,#067a3a);color:#fff;padding:20px;'
            + 'border-radius:8px;text-align:center;margin-bottom:20px;">'
            + f'<h2 style="margin:0;">TechCrunch Weekly &mdash; KW {kw - 1}</h2>'
            + "</div>"
            + '<p style="color:#586069;">Last week\'s top stories you may have missed:</p>'
            + "<ul>"
            + "<li>Quantum computing startup raises $200M Series C</li>"
            + "<li>EU AI Act: What it means for European startups</li>"
            + "<li>The rise of vertical SaaS in healthcare</li>"
            + "</ul>"
            + '<p style="color:#586069;font-size:12px;text-align:center;">'
            + '<a href="#">Unsubscribe</a> | <a href="#">View in browser</a></p>'
            + _HTML_TAIL,
            folder_id="deleteditems", minutes_ago=10800, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Your Figma Pro trial ends in 3 days",
            "Figma", "team@figma.com",
            _HTML_HEAD
            + '<div style="background:#1e1e1e;padding:16px 20px;margin:-20px -20px 20px -20px;">'
            + '<span style="color:#fff;font-size:18px;font-weight:600;">'
            + '<span style="color:#a259ff;margin-right:6px;">&#9670;</span>Figma</span>'
            + "</div>"
            + '<p>Your Figma Professional trial expires on <strong>April 5, 2026</strong>.</p>'
            + "<p>Upgrade now to keep access to:</p>"
            + "<ul>"
            + "<li>Unlimited projects and files</li>"
            + "<li>Advanced prototyping features</li>"
            + "<li>Team libraries and shared components</li>"
            + "</ul>"
            + '<p><a href="#" style="background:#a259ff;color:#fff;padding:10px 24px;border-radius:6px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">Upgrade to Pro</a></p>'
            + _HTML_TAIL,
            folder_id="deleteditems", minutes_ago=7200, is_read=True,
        ),
    )

    messages.append(
        _msg(
            "Webinar Recording: Kubernetes Best Practices 2026",
            "DevOps Weekly", "noreply@devops-weekly.io",
            _HTML_HEAD
            + "<p>Hi there,</p>"
            + "<p>Thanks for registering for our webinar! Here is the recording link:</p>"
            + '<p><a href="#" style="background:#326ce5;color:#fff;padding:8px 20px;border-radius:4px;'
            + 'text-decoration:none;font-weight:600;display:inline-block;">Watch Recording</a></p>'
            + "<p>Key topics covered:</p>"
            + "<ul>"
            + "<li>Multi-cluster management patterns</li>"
            + "<li>GitOps with ArgoCD at scale</li>"
            + "<li>Cost optimization for Kubernetes workloads</li>"
            + "</ul>"
            + '<p style="color:#586069;font-size:12px;">'
            + '<a href="#">Unsubscribe</a> from DevOps Weekly emails</p>'
            + _HTML_TAIL,
            folder_id="deleteditems", minutes_ago=14400, is_read=True,
        ),
    )

    return messages


# ── Directory ────────────────────────────────────────────────────

_AVATAR_COLORS = [
    "#4A90D9", "#D94A4A", "#4AD97A", "#D9A04A", "#7A4AD9",
    "#D94A9A", "#4AD9D9", "#8B6914", "#2E8B57", "#B22222",
    "#4169E1", "#FF8C00", "#6A5ACD", "#20B2AA", "#CD5C5C",
    "#3CB371", "#DAA520", "#5F9EA0", "#BC8F8F", "#6495ED",
    "#F4A460", "#66CDAA", "#DB7093", "#8FBC8F", "#778899",
]


def default_company_directory(count: int = 25) -> list[dict]:
    """~25 synthetic users with org hierarchy in MS Graph JSON shape.

    Each user includes a ``_gender`` field (``"male"`` / ``"female"``)
    for avatar generation.
    """
    # ── Stable IDs for manager references ───────────────────
    ceo_id = _uid()
    vp_tech_id = _uid()
    vp_sales_id = _uid()
    vp_hr_id = _uid()
    vp_finance_id = _uid()
    vp_marketing_id = _uid()
    dir_eng_id = _uid()
    dir_ops_id = _uid()

    def _user(uid: str, given: str, surname: str, email: str,
              title: str, dept: str, gender: str,
              manager_id: str | None = None,
              location: str = "Headquarters",
              phone_suffix: str = "") -> dict:
        mobile_suffix = phone_suffix or uid[:7].replace("-", "")
        office_suffix = phone_suffix or uid[:4].replace("-", "")
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
            "mobilePhone": f"+49 170 {mobile_suffix}",
            "businessPhones": [f"+49 7123 {office_suffix}"],
            "accountEnabled": True,
            "manager": {"id": manager_id} if manager_id else None,
            "_gender": gender,
        }

    users = [
        # ── Executive ───────────────────────────────────────
        _user(ceo_id, "Heinrich", "Fischer", "heinrich.fischer@example.com",
              "CEO", "Executive Board", "male"),
        _user(vp_tech_id, "Klaus", "Weber", "klaus.weber@example.com",
              "VP Technology", "Engineering", "male", manager_id=ceo_id),
        _user(vp_sales_id, "Sabine", "Mueller", "sabine.mueller@example.com",
              "VP Sales", "Sales", "female", manager_id=ceo_id),
        _user(vp_hr_id, "Petra", "Schneider", "petra.schneider@example.com",
              "VP Human Resources", "HR", "female", manager_id=ceo_id),
        _user(vp_finance_id, "Werner", "Hartmann", "werner.hartmann@example.com",
              "VP Finance", "Finance", "male", manager_id=ceo_id),
        _user(vp_marketing_id, "Monika", "Krueger", "monika.krueger@example.com",
              "VP Marketing", "Marketing", "female", manager_id=ceo_id),

        # ── Directors ───────────────────────────────────────
        _user(dir_eng_id, "Frank", "Zimmermann", "frank.zimmermann@example.com",
              "Director of Engineering", "Engineering", "male", manager_id=vp_tech_id),
        _user(dir_ops_id, "Claudia", "Lehmann", "claudia.lehmann@example.com",
              "Director of Operations", "Operations", "female", manager_id=vp_tech_id),

        # ── Engineering (Managers + ICs) ────────────────────
        _user(_uid(), "Lisa", "Braun", "lisa.braun@example.com",
              "Scrum Master", "Engineering", "female", manager_id=dir_eng_id),
        _user(_uid(), "Max", "Mustermann", "max.mustermann@example.com",
              "Senior Software Engineer", "Engineering", "male", manager_id=dir_eng_id),
        _user(_uid(), "Thomas", "Keller", "thomas.keller@example.com",
              "Data Scientist", "Engineering", "male", manager_id=dir_eng_id),
        _user(_uid(), "Stefan", "Hoffmann", "stefan.hoffmann@example.com",
              "DevOps Engineer", "Engineering", "male", manager_id=dir_eng_id),
        _user(_uid(), "Katrin", "Schwarz", "katrin.schwarz@example.com",
              "UX Designer", "Engineering", "female", manager_id=dir_eng_id),
        _user(_uid(), "Tobias", "Neumann", "tobias.neumann@example.com",
              "Backend Developer", "Engineering", "male", manager_id=dir_eng_id),
        _user(_uid(), "Sandra", "Koch", "sandra.koch@example.com",
              "Frontend Developer", "Engineering", "female", manager_id=dir_eng_id),

        # ── Operations ──────────────────────────────────────
        _user(_uid(), "Bernd", "Vogel", "bernd.vogel@example.com",
              "Systems Administrator", "Operations", "male", manager_id=dir_ops_id),
        _user(_uid(), "Melanie", "Schreiber", "melanie.schreiber@example.com",
              "IT Support Specialist", "Operations", "female", manager_id=dir_ops_id),

        # ── Sales ───────────────────────────────────────────
        _user(_uid(), "Anna", "Schmidt", "anna.schmidt@example.com",
              "Sales Manager", "Sales", "female", manager_id=vp_sales_id),
        _user(_uid(), "Julia", "Richter", "julia.richter@example.com",
              "Sales Representative", "Sales", "female", manager_id=vp_sales_id),
        _user(_uid(), "Markus", "Bauer", "markus.bauer@example.com",
              "Key Account Manager", "Sales", "male", manager_id=vp_sales_id,
              location="Stuttgart"),

        # ── HR ──────────────────────────────────────────────
        _user(_uid(), "Christine", "Wagner", "christine.wagner@example.com",
              "HR Business Partner", "HR", "female", manager_id=vp_hr_id),
        _user(_uid(), "Jens", "Lorenz", "jens.lorenz@example.com",
              "Recruiter", "HR", "male", manager_id=vp_hr_id),

        # ── Finance ─────────────────────────────────────────
        _user(_uid(), "Anja", "Beyer", "anja.beyer@example.com",
              "Financial Controller", "Finance", "female", manager_id=vp_finance_id),
        _user(_uid(), "Ralf", "Seidel", "ralf.seidel@example.com",
              "Accountant", "Finance", "male", manager_id=vp_finance_id),

        # ── Marketing ───────────────────────────────────────
        _user(_uid(), "Daniela", "Engel", "daniela.engel@example.com",
              "Marketing Manager", "Marketing", "female", manager_id=vp_marketing_id),
    ]

    return users[:count]


# ── Default mock profile ────────────────────────────────────────

def default_mock_profile() -> MockUserProfile:
    """Return a fully-populated ``MockUserProfile`` with all defaults wired together.

    Includes generated SVG avatar bytes for every directory user.
    """
    directory = default_company_directory()
    mail = default_mail_inbox()
    folders = default_mail_folders()
    events = default_calendar_events()

    # Update folder counts to match actual messages
    folder_msg_counts: dict[str, dict[str, int]] = {}
    for msg in mail:
        fid = msg.get("_folder_id", "inbox")
        if fid not in folder_msg_counts:
            folder_msg_counts[fid] = {"total": 0, "unread": 0}
        folder_msg_counts[fid]["total"] += 1
        if not msg.get("isRead", False):
            folder_msg_counts[fid]["unread"] += 1

    for folder in folders:
        fid = folder["id"]
        if fid in folder_msg_counts:
            folder["totalItemCount"] = folder_msg_counts[fid]["total"]
            folder["unreadItemCount"] = folder_msg_counts[fid]["unread"]

    # Load real face photos by gender, fallback to SVG avatars
    user_photos: dict[str, bytes] = {}
    male_idx = female_idx = 0
    for user in directory:
        gender = user.get("_gender", "male")
        photo = load_face_photo(gender, male_idx if gender == "male" else female_idx)
        if gender == "male":
            male_idx += 1
        else:
            female_idx += 1
        if photo:
            user_photos[user["id"]] = photo
        else:
            given = user.get("givenName", "?")
            surname = user.get("surname", "?")
            initials = f"{given[0]}{surname[0]}".upper()
            color = _AVATAR_COLORS[(male_idx + female_idx) % len(_AVATAR_COLORS)]
            user_photos[user["id"]] = generate_avatar_svg(initials, color)

    # Default teams
    teams = [
        {
            "id": "team-engineering",
            "displayName": "Engineering",
            "description": "Engineering department team",
        },
        {
            "id": "team-general",
            "displayName": "General",
            "description": "Company-wide team",
        },
    ]

    # Default chats
    chats = [
        {
            "id": "chat-001",
            "topic": "Project Alpha Discussion",
            "chatType": "group",
            "createdDateTime": _now().isoformat(),
        },
        {
            "id": "chat-002",
            "topic": None,
            "chatType": "oneOnOne",
            "createdDateTime": (_now() - timedelta(days=1)).isoformat(),
        },
    ]

    # Default drives
    drives = [
        {
            "id": "drive-onedrive",
            "name": "OneDrive",
            "driveType": "personal",
            "owner": {"user": {"displayName": "Mock User"}},
        },
    ]

    return MockUserProfile(
        email="mock@example.com",
        user_id="mock-user-id-00000",
        given_name="Mock",
        surname="User",
        full_name="Mock User",
        job_title="Software Engineer",
        department="Engineering",
        office_location="Headquarters",
        calendar_events=events,
        mail_messages=mail,
        mail_folders=folders,
        directory_users=directory,
        teams=teams,
        chats=chats,
        drives=drives,
        user_photos=user_photos,
    )

from datetime import datetime, timedelta
from typing import Any, List, Optional, Dict
from pydantic import BaseModel, Field

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class CalendarAttendee(BaseModel):
    """An attendee of a calendar event."""
    name: Optional[str] = Field(default=None, description="Attendee display name")
    email: Optional[str] = Field(default=None, description="Attendee email address")
    status: Optional[str] = Field(default=None, description="Response status: accepted, declined, tentative, none")
    type: Optional[str] = Field(default=None, description="Attendee type: required, optional, resource")


class CalendarEvent(BaseModel):
    """A single Outlook calendar event."""
    id: str = Field(description="MS Graph event id")
    subject: str = Field(description="Event subject / title")
    body_preview: Optional[str] = Field(default=None, description="Short plain-text preview of the body")
    body: Optional[str] = Field(default=None, description="Full event body content")
    body_type: Optional[str] = Field(default=None, description="Body content type: 'html' or 'text'")
    start_time: datetime = Field(description="Event start time (UTC)")
    end_time: datetime = Field(description="Event end time (UTC)")
    location: Optional[str] = Field(default=None, description="Event location display name")
    is_all_day: bool = Field(default=False, description="Whether this is an all-day event")
    organizer_name: Optional[str] = Field(default=None, description="Organizer display name")
    organizer_email: Optional[str] = Field(default=None, description="Organizer email address")
    attendees: List[CalendarAttendee] = Field(default_factory=list, description="Event attendees")
    is_online_meeting: bool = Field(default=False, description="Whether an online meeting link is attached")
    online_meeting_url: Optional[str] = Field(default=None, description="Teams / online meeting join URL")
    sensitivity: Optional[str] = Field(default=None, description="Sensitivity: normal, personal, private, confidential")
    show_as: Optional[str] = Field(default=None, description="Free/busy status: free, tentative, busy, oof, workingElsewhere, unknown")
    importance: Optional[str] = Field(default=None, description="Importance: low, normal, high")


class CalendarEventList(BaseModel):
    """Paginated list of calendar events."""
    events: List[CalendarEvent] = Field(default_factory=list, description="Calendar events in this result")
    total_events: int = Field(default=0, description="Total number of events in the queried range")


class CalendarHandler:
    """Handler for Microsoft Graph Calendar API operations."""

    def __init__(self, wui: "MsGraphInstance"):
        self.msg = wui

    def parse_event(self, event: Dict) -> CalendarEvent:
        """Parse a calendar event from the Microsoft Graph API response."""
        start_time_str = event.get('start', {}).get('dateTime')
        end_time_str = event.get('end', {}).get('dateTime')

        start_time = datetime.fromisoformat(start_time_str.replace('Z', '+00:00')) if start_time_str else datetime.now()
        end_time = datetime.fromisoformat(end_time_str.replace('Z', '+00:00')) if end_time_str else datetime.now()

        organizer = event.get('organizer', {}).get('emailAddress', {})
        organizer_name = organizer.get('name')
        organizer_email = organizer.get('address')

        location = event.get('location', {}).get('displayName')

        attendees = []
        for attendee_data in event.get('attendees', []):
            email_address = attendee_data.get('emailAddress', {})
            attendee = CalendarAttendee(
                name=email_address.get('name'),
                email=email_address.get('address'),
                status=attendee_data.get('status', {}).get('response'),
                type=attendee_data.get('type')
            )
            attendees.append(attendee)

        is_online_meeting = bool(event.get('isOnlineMeeting', False))
        online_meeting_url = (event.get('onlineMeeting', {}) or {}).get('joinUrl')

        body = event.get('body', {}) or {}
        return CalendarEvent(
            id=event.get('id', ''),
            subject=event.get('subject') or 'No Subject',
            body_preview=event.get('bodyPreview'),
            body=body.get('content'),
            body_type=body.get('contentType'),
            start_time=start_time,
            end_time=end_time,
            location=location,
            is_all_day=event.get('isAllDay', False),
            organizer_name=organizer_name,
            organizer_email=organizer_email,
            attendees=attendees,
            is_online_meeting=is_online_meeting,
            online_meeting_url=online_meeting_url,
            sensitivity=event.get('sensitivity'),
            show_as=event.get('showAs'),
            importance=event.get('importance')
        )

    # ── async API ─────────────────────────────────────────────────────────

    async def get_calendars_async(self) -> List[Dict]:
        """Get the list of calendars for the current user."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return []
        url = f"{self.msg.msg_endpoint}me/calendars"
        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return []
        try:
            return response.json().get('value', [])
        except Exception:
            return []

    async def get_default_calendar_id_async(self) -> Optional[str]:
        """Get the ID of the default calendar."""
        calendars = await self.get_calendars_async()
        if not calendars:
            return None
        for calendar in calendars:
            if calendar.get('isDefaultCalendar', False):
                return calendar.get('id')
        if calendars:
            return calendars[0].get('id')
        return None

    async def get_events_async(
        self,
        start_date: Optional[datetime] = None,
        end_date: Optional[datetime] = None,
        calendar_id: Optional[str] = None,
        limit: int = 50,
    ) -> CalendarEventList:
        """Get calendar events for a specified time range."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return CalendarEventList()

        if start_date is None:
            now = datetime.now()
            start_date = datetime(now.year, now.month, 1)
        if end_date is None:
            if start_date.month == 12:
                end_date = datetime(start_date.year + 1, 1, 1) - timedelta(days=1)
            else:
                end_date = datetime(start_date.year, start_date.month + 1, 1) - timedelta(days=1)
            end_date = end_date.replace(hour=23, minute=59, second=59)

        if calendar_id is None:
            calendar_id = await self.get_default_calendar_id_async()
            if calendar_id is None:
                return CalendarEventList()

        url = f"{self.msg.msg_endpoint}me/calendars/{calendar_id}/calendarView"
        params = {
            "startDateTime": start_date.isoformat(),
            "endDateTime": end_date.isoformat(),
            "$top": str(limit),
            "$orderby": "start/dateTime",
            "$select": "id,subject,bodyPreview,body,start,end,location,organizer,attendees,isAllDay,isOnlineMeeting,onlineMeeting,sensitivity,showAs,importance",
        }
        param_str = "&".join([f"{k}={v}" for k, v in params.items()])
        url = f"{url}?{param_str}"

        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return CalendarEventList()

        data = response.json()
        events_data = data.get('value', [])
        events = [self.parse_event(event) for event in events_data]
        return CalendarEventList(events=events, total_events=len(events))

    async def get_events_this_month_async(self, calendar_id: Optional[str] = None, limit: int = 50) -> CalendarEventList:
        """Get calendar events for the current month."""
        now = datetime.now()
        start_date = datetime(now.year, now.month, 1)
        if now.month == 12:
            end_date = datetime(now.year + 1, 1, 1) - timedelta(days=1)
        else:
            end_date = datetime(now.year, now.month + 1, 1) - timedelta(days=1)
        end_date = end_date.replace(hour=23, minute=59, second=59)
        return await self.get_events_async(start_date, end_date, calendar_id, limit)

    async def create_event_async(
        self,
        subject: str,
        start_time: datetime,
        end_time: datetime,
        body: Optional[str] = None,
        is_html: bool = False,
        location: Optional[str] = None,
        attendees: Optional[List[Dict[str, str]]] = None,
        is_all_day: bool = False,
        calendar_id: Optional[str] = None,
    ) -> Optional[CalendarEvent]:
        """Create a new calendar event."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None

        if calendar_id is None:
            calendar_id = await self.get_default_calendar_id_async()
            if calendar_id is None:
                return None

        event = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body or ""
            },
            "start": {
                "dateTime": start_time.isoformat(),
                "timeZone": "UTC"
            },
            "end": {
                "dateTime": end_time.isoformat(),
                "timeZone": "UTC"
            },
            "isAllDay": is_all_day
        }

        if location:
            event["location"] = {"displayName": location}

        if attendees:
            event["attendees"] = [
                {
                    "emailAddress": {
                        "address": attendee["email"],
                        "name": attendee.get("name", attendee["email"])
                    },
                    "type": "required"
                }
                for attendee in attendees
            ]

        url = f"{self.msg.msg_endpoint}me/calendars/{calendar_id}/events"
        response = await self.msg.run_async(url=url, method="POST", json=event, token=access_token)
        if response is None or response.status_code != 201:
            return None

        try:
            return self.parse_event(response.json())
        except Exception:
            return None

    async def get_user_timezone_async(self) -> str:
        """Return the logged-in user's timezone."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return "W. Europe Standard Time"
        url = f"{self.msg.msg_endpoint}me/mailboxSettings"
        response = await self.msg.run_async(url=url, token=access_token)
        if response is not None and response.status_code == 200:
            tz = response.json().get("timeZone")
            if tz:
                return tz
        return "W. Europe Standard Time"

    async def get_schedule_async(
        self,
        emails: List[str],
        start: datetime,
        end: datetime,
        interval: int = 30,
        timezone: str = "UTC",
    ) -> List[Dict[str, Any]]:
        """Query free/busy availability for one or more users via getSchedule."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return []
        url = f"{self.msg.msg_endpoint}me/calendar/getSchedule"
        body = {
            "schedules": emails,
            "startTime": {
                "dateTime": start.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": timezone,
            },
            "endTime": {
                "dateTime": end.strftime("%Y-%m-%dT%H:%M:%S"),
                "timeZone": timezone,
            },
            "availabilityViewInterval": interval,
        }
        response = await self.msg.run_async(url=url, method="POST", json=body, token=access_token)
        if response is None or response.status_code != 200:
            return []
        return response.json().get("value", [])

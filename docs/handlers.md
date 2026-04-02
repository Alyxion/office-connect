# MS Graph Handlers Reference

This document provides a comprehensive reference for the Microsoft Graph API
handler layer in `office-mcp`. Each handler wraps a specific MS Graph domain
(mail, calendar, teams, chats, files, directory, profile) behind a typed,
fully async Python interface. All responses are returned as Pydantic models.

## Architecture Overview

The handler layer follows a **factory pattern** centred on
`MsGraphInstance`, which inherits from `WebUserInstance`. The instance
manages OAuth tokens (access and refresh), Redis/memory caching, and
user profile state. Individual domain handlers are instantiated on demand
via factory methods.

```text
MsGraphInstance (factory + auth)
|
+-- get_mail()       -> OfficeMailHandler      (read + write)
+-- get_calendar()   -> CalendarHandler         (read + write)
+-- get_teams()      -> TeamsHandler            (read-only)
+-- get_chat()       -> ChatHandler             (read-only)
+-- get_files()      -> FilesHandler            (read-only)
+-- get_directory()  -> DirectoryHandler        (read-only)
+-- get_profile_async() -> ProfileHandler       (read-only)
```

Every handler stores a back-reference to the `MsGraphInstance` that created
it (`self.msg`). This gives each handler access to the shared access token,
the Graph endpoint URL, and the `run_async()` helper that executes HTTP
requests via `aiohttp`.

### Read-Only vs Read-Write Summary

| Handler | Access | Notes |
|---------|--------|-------|
| `ProfileHandler` | Read-only | Fetches user profile from `/me`. |
| `OfficeMailHandler` | **Read + Write** | Send mail, create/update/send drafts, flag read, set categories. |
| `CalendarHandler` | **Read + Write** | Create events, query free/busy schedules. |
| `TeamsHandler` | Read-only | List joined teams, channels, channel messages, members. |
| `ChatHandler` | Read-only | List chats, chat messages, chat members. |
| `FilesHandler` | Read-only | Browse OneDrive, download files, search, SharePoint sites. |
| `DirectoryHandler` | Read-only | List/search Azure AD users, fetch manager, fetch photo. |

## MsGraphInstance

`MsGraphInstance` is the central entry point. It extends `WebUserInstance`
(which provides OAuth token management, Redis caching, and encrypted token
storage) and adds MS Graph-specific authentication plus handler factories.

**Module:** `office_mcp.msgraph.ms_graph_handler`

### Constructor

```python
MsGraphInstance(
    scopes: list[str] | None = None,
    *,
    cache_dict: dict | None = None,
    redis_url: str | None = None,
    mongodb_url: str | None = None,
    auth_url: str | None = None,
    app: str = "office",
    session_id: str | None = None,
    can_refresh: bool = True,
    client_id: str | None = None,
    client_secret: str | None = None,
    endpoint: str | None = None,
    tenant_id: str | None = None,
    select_account: bool = False,
)
```

| Parameter | Default | Description |
|-----------|---------|-------------|
| `scopes` | `None` | OAuth scopes to request (e.g. `["Mail.Read", "Calendars.Read"]`). |
| `cache_dict` | `None` | Optional dict for in-memory token caching. A new dict is created if `None`. |
| `redis_url` | `None` | Redis connection URL for persistent token storage. |
| `mongodb_url` | `None` | MongoDB connection string for auxiliary data. |
| `auth_url` | `None` | OAuth redirect URL. Used during the authorization code flow. |
| `app` | `"office"` | Application identifier (used as Redis key prefix). |
| `session_id` | `None` | Session identifier. Drives encryption key derivation for token storage in Redis. |
| `can_refresh` | `True` | Whether token refresh is enabled. |
| `client_id` | `None` | Azure AD client ID. Falls back to env var `O365_CLIENT_ID`. |
| `client_secret` | `None` | Azure AD client secret. Falls back to env var `O365_CLIENT_SECRET`. |
| `endpoint` | `None` | MS Graph base URL. Falls back to env var `O365_ENDPOINT`. |
| `tenant_id` | `None` | Azure AD tenant ID. Falls back to env var `O365_TENANT_ID` (default `"common"`). |
| `select_account` | `False` | When `True`, adds `prompt=select_account` to the authorization URL so the user can pick an account. |

#### The `cache_dict`

`cache_dict` is a plain Python dictionary shared between the instance and
its handlers. It acts as a fast in-memory token cache with two reserved keys:

- `"access_token"` -- the current MS Graph bearer token.
- `"refresh_token"` -- the current refresh token.

When Redis is configured, tokens are also persisted there (encrypted with a
key derived from `session_id` and `O365_SALT`). The memory cache is
always checked first for performance.

### Factory Methods

| Method | Returns | Notes |
|--------|---------|-------|
| `get_mail()` | `OfficeMailHandler` | Synchronous factory. |
| `get_calendar()` | `CalendarHandler` | Synchronous factory. |
| `get_teams()` | `TeamsHandler` | Synchronous factory. Import is deferred. |
| `get_chat()` | `ChatHandler` | Synchronous factory. Import is deferred. |
| `get_files()` | `FilesHandler` | Synchronous factory. Import is deferred. |
| `get_directory()` | `DirectoryHandler` | Synchronous factory. Import is deferred. |
| `get_profile_async()` | `ProfileHandler` | **Async.** Fetches the user's `/me` profile on first call and caches the `UserProfile` in `self.me`. |

#### Example

```python
from office_mcp.msgraph.ms_graph_handler import MsGraphInstance

graph = MsGraphInstance(
    scopes=["Mail.Read", "Calendars.ReadWrite"],
    client_id="<client-id>",
    client_secret="<client-secret>",
    tenant_id="<tenant-id>",
    endpoint="https://graph.microsoft.com/v1.0/",
)

# Inject tokens (normally done via OAuth flow)
graph.cache_dict["access_token"] = "<bearer-token>"

# Obtain handlers
mail_handler = graph.get_mail()
cal_handler  = graph.get_calendar()
teams_handler = graph.get_teams()
```

### Token Refresh

`MsGraphInstance` provides two async methods for token lifecycle management:

`refresh_async() -> str | None`
:   Checks whether the current access token is within `min_expiry` seconds
    (default 15 minutes) of expiration. If so, calls `refresh_token_async()`.
    Returns the new token or `None`.

`refresh_token_async() -> str | None`
:   Posts the stored refresh token to the Azure AD `/oauth2/v2.0/token`
    endpoint. On success, stores the new access and refresh tokens in both
    the memory cache and Redis. Returns the new access token or `None`.

`acquire_token_async(code, redirect_url)`
:   Exchanges an authorization code for tokens, fetches the user profile,
    and caches everything. Used at the end of the OAuth authorization code
    flow.

## Profile Handler

**Module:** `office_mcp.msgraph.profile_handler`

**Access:** Read-only

The `ProfileHandler` fetches the authenticated user's profile from the
`/me` endpoint. It is typically obtained via
`await graph.get_profile_async()` rather than instantiated directly.

### Methods

`me` (property)
:   Returns the cached `UserProfile` or `None`. Does not make any
    network call. Use `me_async()` to fetch from the API.

`me_async() -> UserProfile`
:   Fetches the user's profile from MS Graph (`GET /me`) if not already
    cached. Returns a `UserProfile` instance (may have empty fields if
    the request fails).

### UserProfile Model

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Azure AD user UUID. |
| `display_name` | `str` | Full display name (alias `displayName`). |
| `given_name` | `str \| None` | First name (alias `givenName`). |
| `surname` | `str` | Last name. |
| `mail` | `str \| None` | Primary email address. |
| `user_principal_name` | `str` | UPN (alias `userPrincipalName`). |
| `job_title` | `str \| None` | Job title (alias `jobTitle`). |
| `office_location` | `str \| None` | Office location (alias `officeLocation`). |
| `business_phones` | `List[str]` | Business phone numbers (alias `businessPhones`). |
| `mobile_phone` | `str \| None` | Mobile phone number (alias `mobilePhone`). |
| `preferred_language` | `str \| None` | Preferred language (alias `preferredLanguage`). |

### Example

```python
profile = await graph.get_profile_async()
me = profile.me
if me:
    print(f"Hello, {me.display_name} ({me.mail})")
    print(f"Job: {me.job_title}")
```

## Mail Handler

**Module:** `office_mcp.msgraph.mail_handler`

**Access:** Read + Write

`OfficeMailHandler` provides full Outlook mail operations: listing inbox
messages, reading individual emails (with attachments), composing and
sending messages, managing drafts, flagging read state, and category
management.

Obtain via `graph.get_mail()`.

### Methods

#### Inbox and Reading

`email_index_async(limit=40, skip=0, mail_address=None) -> OfficeMailList`
:   Fetches the inbox index. By default returns the authenticated user's
    inbox. Pass `mail_address` to read another user's inbox (requires
    appropriate permissions). Supports pagination via `limit` and `skip`.
    Returns an `OfficeMailList` with `elements` and `total_mails`.

`get_mail_async(email_id=None, email_url=None, attachments=True) -> OfficeMail | None`
:   Retrieves a single email by ID or full URL. When `attachments=True`
    (the default), the `$expand=attachments` query parameter is included
    so attachment content is downloaded inline.

`get_user_profile_async() -> dict | None`
:   Convenience method that fetches `/me` and returns the raw JSON dict.

#### Composing and Sending

`send_message_async(to_recipients, subject, body, is_html=False, save_to_sent_items=True, is_draft=False, attachments=None) -> bool`
:   Sends an email or creates a draft. When `is_draft=True`, creates a
    draft message (`POST /me/messages`). Otherwise sends via
    `POST /me/sendMail`. Supports file attachments.

`create_draft_async(to_recipients, subject, body, is_html=False, cc_recipients=None, bcc_recipients=None, attachments=None) -> dict | None`
:   Creates a draft message and returns `{"id": "...", "webLink": "..."}`.
    If `attachments` are provided, they are added after the draft is
    created.

`update_draft_async(message_id, to_recipients, subject, body, is_html=False, cc_recipients=None, bcc_recipients=None, attachments=None) -> dict | None`
:   Updates an existing draft (`PATCH /me/messages/{id}`). If
    `attachments` is not `None`, all existing attachments are removed
    first and then the new ones are added.

`send_draft_async(message_id) -> bool`
:   Sends an existing draft by ID (`POST /me/messages/{id}/send`).

#### Flags and Categories

`flag_read_async(email_url, read_state: bool) -> bool`
:   Sets the `isRead` property on a message (`PATCH`).

`set_mail_categories_async(email_url, categories: list[str]) -> bool`
:   Sets the category list on a message (`PATCH`).

`get_categories_async(mail_address=None) -> list[OfficeMailCategory]`
:   Returns the master category list for the mailbox.

`ensure_category_exists_async(*, name, color="preset0", mail_address=None) -> bool`
:   Creates a category if it does not already exist. The `color` parameter
    uses Outlook preset strings (`"preset0"` through `"preset24"`).
    See `OfficeCategoryColor` for named constants.

### Pydantic Models

#### OfficeMail

| Field | Type | Description |
|-------|------|-------------|
| `email_id` | `str` | Graph message ID. |
| `email_url` | `str \| None` | Full Graph API URL for the message. |
| `flag_state` | `Literal["flagged", "notFlagged", "done"]` | Outlook flag status. |
| `importance` | `str \| None` | `"low"`, `"normal"`, or `"high"`. |
| `is_read` | `bool` | Whether the message has been read. |
| `email_type` | `str` | OData type (e.g. `"#microsoft.graph.message"`). |
| `local_timestamp` | `str \| None` | Received time converted to local timezone (`YYYY-MM-DD HH:MM:SS`). |
| `from_name` | `str \| None` | Sender display name. |
| `from_email` | `str \| None` | Sender email address. |
| `subject` | `str \| None` | Email subject line. |
| `body_preview` | `str \| None` | Short plain-text preview of the body. |
| `body` | `str \| None` | Full body content (HTML or text). |
| `body_type` | `str \| None` | `"HTML"` or `"Text"`. |
| `has_attachments` | `bool` | Whether the message has attachments. |
| `web_link` | `str \| None` | Outlook Web App deep link. |
| `categories` | `List[str]` | List of category names applied to the message. |
| `confidential_level` | `str \| None` | Sensitivity level: `"normal"`, `"personal"`, `"private"`, `"confidential"`. |
| `attachments` | `List[OfficeMailAttachment]` | Parsed attachment objects (see below). |
| `zip_data` | `bytes \| None` | Optional zipped attachment bundle. |

#### OfficeMailList

| Field | Type | Description |
|-------|------|-------------|
| `elements` | `List[OfficeMail]` | The email items. |
| `total_mails` | `int` | Total count from `@odata.count`. |

#### OfficeMailAttachment

| Field | Type | Description |
|-------|------|-------------|
| `name` | `str` | Filename. |
| `content_type` | `str` | MIME type. |
| `content_bytes` | `bytes \| None` | Decoded file content. |
| `content_id` | `str \| None` | Content-ID header (for inline/embedded images). |
| `is_embedded` | `bool` | `True` if the attachment is referenced inline in the HTML body (e.g. via `cid:`). |

#### OfficeMailCategory

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Category ID. |
| `name` | `str` | Display name. |
| `preset_color` | `str` | Outlook preset color code (e.g. `"preset0"`). |
| `color` | `str` | Resolved HTML color string (e.g. `"red"`, `"darkblue"`). |

#### OfficeCategoryColor

A helper class providing named constants for Outlook's 25 preset colors:

```python
from office_mcp.msgraph.mail_handler import OfficeCategoryColor

OfficeCategoryColor.RED         # "preset0"
OfficeCategoryColor.BLUE        # "preset7"
OfficeCategoryColor.DARK_GREEN  # "preset19"
```

### Example

```python
mail = graph.get_mail()

# List inbox
inbox = await mail.email_index_async(limit=10)
for m in inbox.elements:
    status = "READ" if m.is_read else "UNREAD"
    print(f"[{status}] {m.from_name}: {m.subject}")

# Read a specific email with attachments
email = await mail.get_mail_async(email_id=inbox.elements[0].email_id)
if email:
    print(email.body)
    for att in email.attachments:
        if not att.is_embedded:
            print(f"  Attachment: {att.name} ({att.content_type})")

# Send an email
sent = await mail.send_message_async(
    to_recipients=["colleague@example.com"],
    subject="Meeting notes",
    body="<p>Here are the notes from today.</p>",
    is_html=True,
)

# Draft workflow: create, update, send
draft = await mail.create_draft_async(
    to_recipients=["boss@example.com"],
    subject="Q3 Report",
    body="Please find the report attached.",
)
if draft:
    await mail.update_draft_async(
        message_id=draft["id"],
        to_recipients=["boss@example.com"],
        subject="Q3 Report (updated)",
        body="Revised version attached.",
    )
    await mail.send_draft_async(draft["id"])
```

## Mail Filter

**Module:** `office_mcp.msgraph.mail_filter`

**Access:** N/A (local processing, no network calls)

The mail filter system provides a rule engine for classifying and filtering
`OfficeMail` objects locally. Filters match against sender addresses,
subject lines, and body content. Each rule can have a priority prefix, and
inclusive rules override exclusive rules when they have equal or higher
priority.

### Classes

#### OfficeMailFilter

Defines a single named filter with six rule lists:

| Field | Type | Description |
|-------|------|-------------|
| `name` | `str` | Filter name (used in results for attribution). |
| `senders_excluded` | `List[str]` | Sender patterns to exclude (supports `*` wildcards via `fnmatch`). |
| `subjects_excluded` | `List[str]` | Subject substrings to exclude (case-insensitive). |
| `body_content_excluded` | `List[str]` | Body substrings to exclude (case-insensitive). |
| `senders_included` | `List[str] \| None` | Sender patterns to include (overrides exclusion at equal/higher priority). |
| `subjects_included` | `List[str] \| None` | Subject substrings to include. |
| `body_content_included` | `List[str] \| None` | Body substrings to include. |

**Priority syntax:** Each entry in the rule lists can have an optional
priority prefix, e.g. `"P1:sender@domain.com"`. The priority is an
integer from 0 (highest) to 100 (lowest/default). When a mail matches
both an inclusive and an exclusive rule, the rule with the lower priority
number wins. On a tie, inclusive wins.

**Loading from files:**

```python
# From a JSON file
f = OfficeMailFilter.from_json_file("filters/spam.json")

# From a dict
f = OfficeMailFilter.from_dict({
    "name": "newsletters",
    "senders_excluded": ["*@newsletter.example.com"],
    "subjects_excluded": ["Unsubscribe"],
})
```

`apply(mail: OfficeMail) -> OfficeMailFilterResults`
:   Evaluates the filter against a single email and returns detailed results.

#### OfficeMailFilterList

A collection of `OfficeMailFilter` objects. Supports merging via `+`
operator and `combine()` method.

```python
from office_mcp.msgraph.mail_filter import OfficeMailFilterList

filters = OfficeMailFilterList.from_json_files([
    "filters/spam.json",
    "filters/newsletters.json",
])
# Or combine programmatically
combined = filter_a + filter_b  # OfficeMailFilter + OfficeMailFilter
combined = list_a + list_b      # OfficeMailFilterList + OfficeMailFilterList
```

`apply(mail: OfficeMail) -> OfficeMailFilterResults`
:   Applies all filters and returns combined results.

#### OfficeMailFilterResults

| Field | Type | Description |
|-------|------|-------------|
| `email` | `OfficeMail` | The email that was tested. |
| `any_filter_hit` | `bool` | `True` if any filter rule matched. |
| `matched_filters` | `List[str]` | Names of filters that matched. |
| `matched_reasons` | `Dict[str, List[OfficeMailFilterReason]]` | Detailed reasons keyed by filter name. |
| `excluded` | `bool` | `True` if the email should be excluded (blocked). |

`get_reason_text(lang="en") -> List[str]`
:   Returns human-readable explanation strings. Supports `"en"` and
    `"de"` languages.

#### OfficeMailFilterReason

| Field | Type | Description |
|-------|------|-------------|
| `filter_name` | `str` | Name of the filter that produced this reason. |
| `reason_type` | `Literal["sender", "subject", "body"]` | Which part of the email matched. |
| `value` | `str` | The matched value (email address, subject, or body text). |
| `inclusive` | `bool` | `True` if this is an inclusive (allow) rule. |
| `priority` | `int` | Priority of the winning rule (0 = highest). |

### Example

```python
from office_mcp.msgraph.mail_filter import OfficeMailFilter

spam_filter = OfficeMailFilter(
    name="spam",
    senders_excluded=["*@spam.example.com", "P2:noreply@*"],
    senders_included=["P1:noreply@important.example.com"],
    subjects_excluded=["Win a prize"],
)

mail = graph.get_mail()
inbox = await mail.email_index_async(limit=50)
for email in inbox.elements:
    result = spam_filter.apply(email)
    if result.excluded:
        reasons = result.get_reason_text(lang="en")
        print(f"BLOCKED: {email.subject}")
        for r in reasons:
            print(f"  -> {r}")
```

## Calendar Handler

**Module:** `office_mcp.msgraph.calendar_handler`

**Access:** Read + Write

The `CalendarHandler` provides operations for reading calendar events,
creating new events, querying free/busy schedules, and retrieving user
timezone settings.

Obtain via `graph.get_calendar()`.

### Methods

#### Reading Events

`get_calendars_async() -> List[Dict]`
:   Returns the raw list of calendar objects for the current user.

`get_default_calendar_id_async() -> str | None`
:   Returns the ID of the user's default calendar (the one marked
    `isDefaultCalendar`, or the first calendar).

`get_events_async(start_date=None, end_date=None, calendar_id=None, limit=50) -> CalendarEventList`
:   Fetches calendar events within a date range using the `calendarView`
    endpoint (which expands recurring events). If `start_date` is
    `None`, defaults to the first day of the current month. If
    `end_date` is `None`, defaults to the last day of the same month
    as `start_date`. Events are ordered by start time.

`get_events_this_month_async(calendar_id=None, limit=50) -> CalendarEventList`
:   Convenience method that calls `get_events_async` with the current
    month's date range.

#### Creating Events

`create_event_async(subject, start_time, end_time, body=None, is_html=False, location=None, attendees=None, is_all_day=False, calendar_id=None) -> CalendarEvent | None`
:   Creates a new calendar event. The `attendees` parameter is a list of
    dicts with `"email"` and optional `"name"` keys. All attendees are
    added as `"required"` type. Returns the created event or `None` on
    failure.

#### Availability

`get_schedule_async(emails, start, end, interval=30, timezone="UTC") -> List[Dict]`
:   Queries the `getSchedule` endpoint for free/busy availability of one
    or more users. The `interval` is in minutes. Returns the raw schedule
    data from MS Graph.

`get_user_timezone_async() -> str`
:   Returns the logged-in user's mailbox timezone setting. Falls back to
    `"W. Europe Standard Time"` if unavailable.

### Pydantic Models

#### CalendarEvent

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Event ID. |
| `subject` | `str` | Event subject (defaults to `"No Subject"`). |
| `body_preview` | `str \| None` | Short text preview of the event body. |
| `body` | `str \| None` | Full body content. |
| `body_type` | `str \| None` | `"HTML"` or `"Text"`. |
| `start_time` | `datetime` | Event start time. |
| `end_time` | `datetime` | Event end time. |
| `location` | `str \| None` | Location display name. |
| `is_all_day` | `bool` | Whether this is an all-day event. |
| `organizer_name` | `str \| None` | Organizer display name. |
| `organizer_email` | `str \| None` | Organizer email address. |
| `attendees` | `List[CalendarAttendee]` | List of attendees (see below). |
| `is_online_meeting` | `bool` | Whether a Teams meeting link is attached. |
| `online_meeting_url` | `str \| None` | Teams join URL. |
| `sensitivity` | `str \| None` | `"normal"`, `"personal"`, `"private"`, or `"confidential"`. |
| `show_as` | `str \| None` | `"free"`, `"tentative"`, `"busy"`, `"oof"`, `"workingElsewhere"`, or `"unknown"`. |
| `importance` | `str \| None` | `"low"`, `"normal"`, or `"high"`. |

#### CalendarAttendee

| Field | Type | Description |
|-------|------|-------------|
| `name` | `str \| None` | Attendee display name. |
| `email` | `str \| None` | Attendee email address. |
| `status` | `str \| None` | Response status: `"accepted"`, `"declined"`, `"tentative"`, `"notResponded"`, etc. |
| `type` | `str \| None` | `"required"`, `"optional"`, or `"resource"`. |

#### CalendarEventList

| Field | Type | Description |
|-------|------|-------------|
| `events` | `List[CalendarEvent]` | The event items. |
| `total_events` | `int` | Total number of events returned. |

### Example

```python
from datetime import datetime, timedelta

cal = graph.get_calendar()

# This week's events
now = datetime.now()
events = await cal.get_events_async(
    start_date=now,
    end_date=now + timedelta(days=7),
    limit=25,
)
for e in events.events:
    online = " [Teams]" if e.is_online_meeting else ""
    print(f"{e.start_time:%H:%M} - {e.subject}{online}")
    for a in e.attendees:
        print(f"    {a.name} ({a.status})")

# Create a new event
new_event = await cal.create_event_async(
    subject="Project Review",
    start_time=datetime(2026, 3, 10, 14, 0),
    end_time=datetime(2026, 3, 10, 15, 0),
    body="<p>Agenda: Q1 review</p>",
    is_html=True,
    location="Conference Room B",
    attendees=[
        {"email": "alice@example.com", "name": "Alice"},
        {"email": "bob@example.com", "name": "Bob"},
    ],
)

# Check free/busy for multiple users
schedules = await cal.get_schedule_async(
    emails=["alice@example.com", "bob@example.com"],
    start=datetime(2026, 3, 10, 8, 0),
    end=datetime(2026, 3, 10, 18, 0),
    interval=30,
    timezone="Europe/Berlin",
)
```

## Teams Handler

**Module:** `office_mcp.msgraph.teams_handler`

**Access:** Read-only

The `TeamsHandler` provides read access to Microsoft Teams data: joined
teams, channels within teams, channel messages, and team membership.

Obtain via `graph.get_teams()`.

### Methods

`get_joined_teams_async() -> TeamList`
:   Lists all teams the authenticated user has joined
    (`GET /me/joinedTeams`).

`get_channels_async(team_id: str) -> ChannelList`
:   Lists channels in a team (`GET /teams/{team_id}/channels`).

`get_channel_messages_async(team_id: str, channel_id: str, limit=20) -> ChannelMessageList`
:   Fetches recent messages from a channel
    (`GET /teams/{team_id}/channels/{channel_id}/messages`).

`get_team_members_async(team_id: str) -> TeamMemberList`
:   Lists members of a team (`GET /teams/{team_id}/members`).

### Pydantic Models

#### Team

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Team (group) ID. |
| `display_name` | `str \| None` | Team display name. |
| `description` | `str \| None` | Team description. |
| `visibility` | `str \| None` | `"public"`, `"private"`, or `"hiddenMembership"`. |
| `is_archived` | `bool` | Whether the team is archived. |
| `web_url` | `str \| None` | Deep-link to the team in the Teams client. |

#### TeamList

| Field | Type | Description |
|-------|------|-------------|
| `teams` | `List[Team]` | The team items. |
| `total_teams` | `int` | Number of teams. |

#### Channel

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Channel ID. |
| `display_name` | `str \| None` | Channel name. |
| `description` | `str \| None` | Channel description. |
| `membership_type` | `str \| None` | `"standard"`, `"private"`, or `"shared"`. |
| `web_url` | `str \| None` | Deep-link to the channel. |
| `is_favorite_by_default` | `bool \| None` | Whether the channel auto-favorites for new members. |

#### ChannelList

| Field | Type | Description |
|-------|------|-------------|
| `channels` | `List[Channel]` | The channel items. |
| `total_channels` | `int` | Number of channels. |

#### ChannelMessage

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Message ID. |
| `created_at` | `datetime \| None` | When the message was created (UTC). |
| `subject` | `str \| None` | Thread subject (root message only). |
| `body_content` | `str \| None` | Message body (text or HTML). |
| `body_type` | `str \| None` | `"text"` or `"html"`. |
| `sender` | `ChannelMessageFrom \| None` | Sender information. |
| `importance` | `str \| None` | `"normal"`, `"high"`, or `"urgent"`. |
| `web_url` | `str \| None` | Deep-link to the message. |

#### ChannelMessageFrom

| Field | Type | Description |
|-------|------|-------------|
| `display_name` | `str \| None` | Sender display name. |
| `email` | `str \| None` | Sender email. |
| `user_id` | `str \| None` | Azure AD user ID. |

#### ChannelMessageList

| Field | Type | Description |
|-------|------|-------------|
| `messages` | `List[ChannelMessage]` | The message items. |
| `total_messages` | `int` | Number of messages. |

#### TeamMember

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Membership ID. |
| `display_name` | `str \| None` | Member display name. |
| `email` | `str \| None` | Member email. |
| `user_id` | `str \| None` | Azure AD user ID. |
| `roles` | `List[str]` | Roles: `"owner"`, `"member"`, `"guest"`. |

#### TeamMemberList

| Field | Type | Description |
|-------|------|-------------|
| `members` | `List[TeamMember]` | The member items. |
| `total_members` | `int` | Number of members. |

### Example

```python
teams = graph.get_teams()

# List joined teams
team_list = await teams.get_joined_teams_async()
for t in team_list.teams:
    archived = " (archived)" if t.is_archived else ""
    print(f"{t.display_name}{archived}")

    # Channels in each team
    channels = await teams.get_channels_async(t.id)
    for ch in channels.channels:
        print(f"  #{ch.display_name} [{ch.membership_type}]")

    # Messages in the first channel
    if channels.channels:
        msgs = await teams.get_channel_messages_async(
            t.id, channels.channels[0].id, limit=5
        )
        for m in msgs.messages:
            sender = m.sender.display_name if m.sender else "Unknown"
            print(f"    [{sender}]: {m.body_content[:80] if m.body_content else ''}")

# Team members
members = await teams.get_team_members_async(team_list.teams[0].id)
for m in members.members:
    print(f"  {m.display_name} ({', '.join(m.roles)})")
```

## Chat Handler

**Module:** `office_mcp.msgraph.chat_handler`

**Access:** Read-only

**Required scopes:** `Chat.Read` or `Chat.ReadWrite`

The `ChatHandler` provides read access to the authenticated user's 1:1,
group, and meeting chats in Microsoft Teams.

Obtain via `graph.get_chat()`.

### Methods

`get_chats_async(limit=50) -> ChatList`
:   Lists the current user's chats (`GET /me/chats`). Includes 1:1,
    group, and meeting chats.

`get_chat_messages_async(chat_id: str, limit=20) -> ChatMessageList`
:   Fetches recent messages from a specific chat
    (`GET /me/chats/{chat_id}/messages`).

`get_chat_members_async(chat_id: str) -> ChatMemberList`
:   Lists the members of a chat (`GET /me/chats/{chat_id}/members`).

### Pydantic Models

#### Chat

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Chat ID. |
| `topic` | `str \| None` | Chat topic (typically set for group chats). |
| `chat_type` | `str \| None` | `"oneOnOne"`, `"group"`, or `"meeting"`. |
| `created_at` | `datetime \| None` | When the chat was created. |
| `last_updated_at` | `datetime \| None` | When the chat was last updated. |
| `web_url` | `str \| None` | Deep-link to the chat in Teams. |
| `tenant_id` | `str \| None` | Azure AD tenant ID. |

#### ChatList

| Field | Type | Description |
|-------|------|-------------|
| `chats` | `List[Chat]` | The chat items. |
| `total_chats` | `int` | Number of chats. |

#### ChatMessage

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Message ID. |
| `created_at` | `datetime \| None` | When the message was created (UTC). |
| `body_content` | `str \| None` | Message body content (text or HTML). |
| `body_type` | `str \| None` | `"text"` or `"html"`. |
| `sender` | `ChatMessageFrom \| None` | Sender information. |
| `importance` | `str \| None` | `"normal"`, `"high"`, or `"urgent"`. |
| `message_type` | `str \| None` | `"message"`, `"chatEvent"`, `"typing"`, etc. |
| `web_url` | `str \| None` | Deep-link to the message. |

#### ChatMessageFrom

| Field | Type | Description |
|-------|------|-------------|
| `display_name` | `str \| None` | Sender display name. |
| `email` | `str \| None` | Sender email. |
| `user_id` | `str \| None` | Azure AD user ID. |

#### ChatMessageList

| Field | Type | Description |
|-------|------|-------------|
| `messages` | `List[ChatMessage]` | The message items. |
| `total_messages` | `int` | Number of messages. |

#### ChatMember

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Membership ID. |
| `display_name` | `str \| None` | Member display name. |
| `email` | `str \| None` | Member email. |
| `user_id` | `str \| None` | Azure AD user ID. |
| `roles` | `List[str]` | Roles in the chat. |

#### ChatMemberList

| Field | Type | Description |
|-------|------|-------------|
| `members` | `List[ChatMember]` | The member items. |
| `total_members` | `int` | Number of members. |

### Example

```python
chat = graph.get_chat()

# List recent chats
chat_list = await chat.get_chats_async(limit=10)
for c in chat_list.chats:
    label = c.topic or c.chat_type or c.id[:20]
    print(f"Chat: {label} (type={c.chat_type})")

# Read messages from a specific chat
if chat_list.chats:
    msgs = await chat.get_chat_messages_async(chat_list.chats[0].id, limit=5)
    for m in msgs.messages:
        sender = m.sender.display_name if m.sender else "System"
        body = (m.body_content or "")[:100]
        print(f"  [{sender}]: {body}")

# List chat members
if chat_list.chats:
    members = await chat.get_chat_members_async(chat_list.chats[0].id)
    for m in members.members:
        print(f"  - {m.display_name} ({m.email})")
```

## Files Handler

**Module:** `office_mcp.msgraph.files_handler`

**Access:** Read-only

**Required scopes:** `Files.Read.All` (or `Files.ReadWrite.All`),
`Sites.Read.All`

The `FilesHandler` provides read access to OneDrive files and folders,
SharePoint document libraries, file downloads, and search.

Obtain via `graph.get_files()`.

### Methods

#### Drives

`get_my_drives_async() -> DriveList`
:   Lists all drives accessible to the current user (personal OneDrive plus
    shared drives) via `GET /me/drives`.

`get_my_drive_async() -> Drive | None`
:   Returns the user's default OneDrive via `GET /me/drive`.

#### Drive Items (Files and Folders)

`get_root_items_async(drive_id=None, limit=50) -> DriveItemList`
:   Lists items in the root of a drive. If `drive_id` is `None`, uses
    the user's default OneDrive.

`get_folder_items_async(item_id, drive_id=None, limit=50) -> DriveItemList`
:   Lists children of a folder identified by `item_id`.

`get_item_async(item_id, drive_id=None) -> DriveItem | None`
:   Returns metadata for a single drive item (file or folder).

`get_file_content_async(item_id, drive_id=None) -> bytes | None`
:   Downloads the binary content of a file. Returns `None` for folders
    or on failure.

`search_items_async(query, drive_id=None, limit=25) -> DriveItemList`
:   Full-text search for files and folders by name or content using the
    Graph `search(q='...')` function.

#### SharePoint Sites

`get_followed_sites_async() -> SharePointSiteList`
:   Lists SharePoint sites the user follows (`GET /me/followedSites`).

`search_sites_async(query: str) -> SharePointSiteList`
:   Searches for SharePoint sites by keyword
    (`GET /sites?search={query}`).

`get_site_drives_async(site_id: str) -> DriveList`
:   Lists document libraries in a SharePoint site
    (`GET /sites/{site_id}/drives`).

### Pydantic Models

#### Drive

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Drive ID. |
| `name` | `str \| None` | Drive name. |
| `drive_type` | `str \| None` | `"personal"`, `"business"`, or `"documentLibrary"`. |
| `owner_name` | `str \| None` | Owner display name. |
| `quota_total` | `int \| None` | Total quota in bytes. |
| `quota_used` | `int \| None` | Used quota in bytes. |
| `web_url` | `str \| None` | Browser URL to the drive root. |

#### DriveList

| Field | Type | Description |
|-------|------|-------------|
| `drives` | `List[Drive]` | The drive items. |
| `total_drives` | `int` | Number of drives. |

#### DriveItem

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Drive item ID. |
| `name` | `str \| None` | File or folder name. |
| `size` | `int \| None` | Size in bytes. |
| `web_url` | `str \| None` | Browser URL. |
| `created_at` | `datetime \| None` | Creation timestamp (UTC). |
| `modified_at` | `datetime \| None` | Last modified timestamp (UTC). |
| `created_by` | `DriveItemUser \| None` | Who created the item. |
| `modified_by` | `DriveItemUser \| None` | Who last modified the item. |
| `mime_type` | `str \| None` | MIME type (files only). |
| `is_folder` | `bool` | `True` if the item is a folder. |
| `folder_child_count` | `int \| None` | Number of children (folders only). |
| `download_url` | `str \| None` | Pre-authenticated download URL (short-lived, from `@microsoft.graph.downloadUrl`). |

#### DriveItemList

| Field | Type | Description |
|-------|------|-------------|
| `items` | `List[DriveItem]` | The item list. |
| `total_items` | `int` | Number of items. |

#### DriveItemUser

| Field | Type | Description |
|-------|------|-------------|
| `display_name` | `str \| None` | Display name. |
| `email` | `str \| None` | Email address. |
| `user_id` | `str \| None` | Azure AD user ID. |

#### SharePointSite

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Site ID (format: `host,site-collection-id,web-id`). |
| `display_name` | `str \| None` | Site display name. |
| `name` | `str \| None` | Site URL name. |
| `web_url` | `str \| None` | Full URL to the site. |
| `description` | `str \| None` | Site description. |
| `created_at` | `datetime \| None` | When the site was created. |

#### SharePointSiteList

| Field | Type | Description |
|-------|------|-------------|
| `sites` | `List[SharePointSite]` | The site items. |
| `total_sites` | `int` | Number of sites. |

### Example

```python
files = graph.get_files()

# Default OneDrive info
drive = await files.get_my_drive_async()
if drive:
    used_gb = (drive.quota_used or 0) / (1024 ** 3)
    total_gb = (drive.quota_total or 0) / (1024 ** 3)
    print(f"OneDrive: {drive.name} ({used_gb:.1f} / {total_gb:.1f} GB)")

# Browse root
root = await files.get_root_items_async(limit=20)
for item in root.items:
    kind = "DIR " if item.is_folder else "FILE"
    print(f"  [{kind}] {item.name}  ({item.size or 0} bytes)")

# Navigate into a folder
folder = next((i for i in root.items if i.is_folder), None)
if folder:
    children = await files.get_folder_items_async(folder.id)
    for child in children.items:
        print(f"    {child.name}")

# Download a file
file_item = next((i for i in root.items if not i.is_folder), None)
if file_item:
    content = await files.get_file_content_async(file_item.id)
    if content:
        print(f"Downloaded {len(content)} bytes")

# Search
results = await files.search_items_async("quarterly report")
for item in results.items:
    print(f"  Found: {item.name} ({item.web_url})")

# SharePoint sites
sites = await files.get_followed_sites_async()
for s in sites.sites:
    print(f"Site: {s.display_name} ({s.web_url})")
    drives = await files.get_site_drives_async(s.id)
    for d in drives.drives:
        print(f"  Library: {d.name}")
```

## Directory Handler

**Module:** `office_mcp.msgraph.directory_handler`

**Access:** Read-only

The `DirectoryHandler` provides access to the Azure AD / Entra ID
organizational directory: listing users, retrieving manager relationships,
and downloading user profile photos.

Obtain via `graph.get_directory()`.

### Methods

`get_users_async(limit=100) -> DirectoryUserList`
:   Returns the first page of users (up to 100). Selects rich fields
    including job title, department, manager, office location, and phone.
    Uses `$expand=manager($select=id)` to include the manager UUID in
    a single call.

`get_all_users_async() -> DirectoryUserList`
:   Fetches **all** users by following `@odata.nextLink` pagination.
    Serialized with an async lock to prevent concurrent full-directory
    fetches. Additionally expands license and account-enabled status.

`get_user_manager_async(user_id: str) -> dict | None`
:   Returns the raw manager JSON for a user
    (`GET /users/{user_id}/manager`).

`get_user_photo_async(user_id: str) -> bytes | None`
:   Downloads the user's profile photo as raw bytes
    (`GET /users/{user_id}/photo/$value`). Returns `None` if no
    photo is set or the request fails.

### Pydantic Models

#### DirectoryUser

| Field | Type | Description |
|-------|------|-------------|
| `id` | `str` | Azure AD user UUID. |
| `display_name` | `str \| None` | Full display name. |
| `email` | `str \| None` | Email address (falls back to UPN if `mail` is null). |
| `job_title` | `str \| None` | Job title (e.g. `"Sales Manager"`). |
| `department` | `str \| None` | Department (e.g. `"Sales"`). |
| `manager_id` | `str \| None` | Manager's Azure AD UUID. |
| `account_enabled` | `bool \| None` | Whether the account is enabled. |
| `surname` | `str \| None` | Last name. |
| `given_name` | `str \| None` | First name. |
| `office_location` | `str \| None` | Office location (e.g. `"Office K"`). |
| `mobile_phone` | `str \| None` | Mobile phone number. |

#### DirectoryUserList

| Field | Type | Description |
|-------|------|-------------|
| `users` | `List[DirectoryUser]` | The user items. |
| `total_users` | `int` | Number of users. |

### Example

```python
directory = graph.get_directory()

# First page of users
user_list = await directory.get_users_async(limit=50)
for u in user_list.users:
    enabled = "active" if u.account_enabled else "disabled"
    print(f"{u.display_name} ({u.email}) - {u.job_title} [{enabled}]")

# Get a user's manager
if user_list.users:
    user = user_list.users[0]
    if user.manager_id:
        manager = await directory.get_user_manager_async(user.id)
        if manager:
            print(f"{user.display_name}'s manager: {manager.get('displayName')}")

# Download profile photo
if user_list.users:
    photo_bytes = await directory.get_user_photo_async(user_list.users[0].id)
    if photo_bytes:
        with open("photo.jpg", "wb") as f:
            f.write(photo_bytes)
        print(f"Saved profile photo ({len(photo_bytes)} bytes)")

# Full directory export (paginated, may be slow for large tenants)
all_users = await directory.get_all_users_async()
print(f"Total directory users: {all_users.total_users}")
```

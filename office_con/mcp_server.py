"""Office 365 MCP Server — read-only access to Microsoft 365 via MS Graph.

Exposes tools for reading mail, calendar, teams, chats, files, directory,
and profile data. Authenticates via a JSON key file containing OAuth tokens.

Usage:
    python -m office_con.mcp_server --keyfile path/to/token.json

Key file format:
    {
        "app": "MyApp",
        "session_id": "...",
        "email": "user@example.com",
        "access_token": "eyJ...",
        "refresh_token": "1.AUs..."
    }
"""

from __future__ import annotations

import argparse
import asyncio
import json
import logging
import os
import sys
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from office_con.msgraph.ms_graph_handler import MsGraphInstance

logger = logging.getLogger(__name__)


def _write_secure_json(path: str, data: dict) -> None:
    """Write JSON to a file with restrictive permissions (0600)."""
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(data, f, indent=2)


# ---------------------------------------------------------------------------
# Graph instance bootstrap
# ---------------------------------------------------------------------------


def export_keyfile(
    path: str,
    *,
    access_token: str,
    refresh_token: str,
    client_id: str,
    client_secret: str,
    tenant_id: str = "common",
    app: str = "office-mcp",
    session_id: str | None = None,
    email: str | None = None,
) -> None:
    """Export a complete keyfile with all fields needed for token refresh."""
    data = {
        "app": app,
        "session_id": session_id or "",
        "email": email or "",
        "access_token": access_token,
        "refresh_token": refresh_token,
        "client_id": client_id,
        "client_secret": client_secret,
        "tenant_id": tenant_id,
    }
    _write_secure_json(path, data)


async def _create_graph(keyfile: str) -> MsGraphInstance:
    """Create an authenticated MsGraphInstance from a key file."""
    from pathlib import Path as _Path
    data = await asyncio.to_thread(lambda: json.loads(_Path(keyfile).read_text()))

    inst = MsGraphInstance(
        scopes=None,
        app=data.get("app", "office-mcp"),
        session_id=data.get("session_id"),
        can_refresh=True,
        endpoint=data.get("endpoint", "https://graph.microsoft.com/v1.0/"),
        client_id=data.get("client_id"),
        client_secret=data.get("client_secret"),
        tenant_id=data.get("tenant_id"),
    )
    inst.email = data.get("email")
    inst.cache_dict["access_token"] = data["access_token"]
    if data.get("refresh_token"):
        inst.cache_dict["refresh_token"] = data["refresh_token"]

    # Try refreshing the token
    try:
        refreshed = await inst.refresh_token_async()
        if refreshed:
            inst.cache_dict["access_token"] = refreshed
            # Persist refreshed tokens back to the keyfile
            data["access_token"] = refreshed
            new_refresh = inst.cache_dict.get("refresh_token")
            if new_refresh:
                data["refresh_token"] = new_refresh
            _write_secure_json(keyfile, data)
    except Exception:
        pass
    # Ensure we still have a token
    if not inst.cache_dict.get("access_token"):
        inst.cache_dict["access_token"] = data["access_token"]

    return inst


# ---------------------------------------------------------------------------
# Tool definitions
# ---------------------------------------------------------------------------

TOOLS = [
    # ── Profile ───────────────────────────────────────────────────────
    Tool(
        name="o365_get_profile",
        description="Get the current user's profile (name, email, job title, department, phone, location).",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    # ── Mail ──────────────────────────────────────────────────────────
    Tool(
        name="o365_list_mail",
        description="List recent emails from the user's inbox.",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {"type": "integer", "description": "Max emails to return (default 10)", "default": 10},
                "skip": {"type": "integer", "description": "Number of emails to skip for pagination", "default": 0},
            },
            "required": [],
        },
    ),
    Tool(
        name="o365_get_mail",
        description="Get a single email by ID, including full body and attachments metadata.",
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "The email message ID"},
            },
            "required": ["email_id"],
        },
    ),
    Tool(
        name="o365_get_mail_categories",
        description="List the user's Outlook mail categories.",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    # ── Calendar ──────────────────────────────────────────────────────
    Tool(
        name="o365_list_calendars",
        description="List the user's calendars.",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    Tool(
        name="o365_get_events",
        description="Get calendar events within a date range.",
        inputSchema={
            "type": "object",
            "properties": {
                "start_date": {"type": "string", "description": "Start date (ISO 8601, e.g. 2026-03-01)"},
                "end_date": {"type": "string", "description": "End date (ISO 8601, e.g. 2026-03-31)"},
                "limit": {"type": "integer", "description": "Max events to return (default 25)", "default": 25},
            },
            "required": ["start_date", "end_date"],
        },
    ),
    Tool(
        name="o365_get_schedule",
        description="Get free/busy availability for one or more users.",
        inputSchema={
            "type": "object",
            "properties": {
                "emails": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Email addresses to check availability for",
                },
                "start": {"type": "string", "description": "Start datetime (ISO 8601)"},
                "end": {"type": "string", "description": "End datetime (ISO 8601)"},
            },
            "required": ["emails", "start", "end"],
        },
    ),
    # ── Teams ─────────────────────────────────────────────────────────
    Tool(
        name="o365_list_teams",
        description="List Microsoft Teams the user has joined.",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    Tool(
        name="o365_list_channels",
        description="List channels in a team.",
        inputSchema={
            "type": "object",
            "properties": {
                "team_id": {"type": "string", "description": "Team ID"},
            },
            "required": ["team_id"],
        },
    ),
    Tool(
        name="o365_get_channel_messages",
        description="Get recent messages from a team channel.",
        inputSchema={
            "type": "object",
            "properties": {
                "team_id": {"type": "string", "description": "Team ID"},
                "channel_id": {"type": "string", "description": "Channel ID"},
                "limit": {"type": "integer", "description": "Max messages (default 20)", "default": 20},
            },
            "required": ["team_id", "channel_id"],
        },
    ),
    Tool(
        name="o365_get_team_members",
        description="List members of a team.",
        inputSchema={
            "type": "object",
            "properties": {
                "team_id": {"type": "string", "description": "Team ID"},
            },
            "required": ["team_id"],
        },
    ),
    # ── Chats ─────────────────────────────────────────────────────────
    Tool(
        name="o365_list_chats",
        description="List the user's recent chats (1:1, group, meeting).",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {"type": "integer", "description": "Max chats to return (default 25)", "default": 25},
            },
            "required": [],
        },
    ),
    Tool(
        name="o365_get_chat_messages",
        description="Get recent messages from a chat.",
        inputSchema={
            "type": "object",
            "properties": {
                "chat_id": {"type": "string", "description": "Chat ID"},
                "limit": {"type": "integer", "description": "Max messages (default 20)", "default": 20},
            },
            "required": ["chat_id"],
        },
    ),
    Tool(
        name="o365_get_chat_members",
        description="List members of a chat.",
        inputSchema={
            "type": "object",
            "properties": {
                "chat_id": {"type": "string", "description": "Chat ID"},
            },
            "required": ["chat_id"],
        },
    ),
    # ── Files / OneDrive ──────────────────────────────────────────────
    Tool(
        name="o365_get_my_drive",
        description="Get the user's default OneDrive info.",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    Tool(
        name="o365_list_drive_items",
        description="List files and folders in a drive location.",
        inputSchema={
            "type": "object",
            "properties": {
                "folder_id": {"type": "string", "description": "Folder ID (omit for root)", "default": ""},
                "drive_id": {"type": "string", "description": "Drive ID (omit for default OneDrive)", "default": ""},
                "limit": {"type": "integer", "description": "Max items (default 25)", "default": 25},
            },
            "required": [],
        },
    ),
    Tool(
        name="o365_get_file_content",
        description="Download a file's content as text (UTF-8). For binary files, returns base64.",
        inputSchema={
            "type": "object",
            "properties": {
                "item_id": {"type": "string", "description": "Drive item ID"},
                "drive_id": {"type": "string", "description": "Drive ID (omit for default OneDrive)", "default": ""},
            },
            "required": ["item_id"],
        },
    ),
    Tool(
        name="o365_search_files",
        description="Search for files by name or content in OneDrive.",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Search query"},
                "limit": {"type": "integer", "description": "Max results (default 10)", "default": 10},
            },
            "required": ["query"],
        },
    ),
    # ── SharePoint ────────────────────────────────────────────────────
    Tool(
        name="o365_search_sites",
        description="Search for SharePoint sites.",
        inputSchema={
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Search query (use * for all sites)"},
            },
            "required": ["query"],
        },
    ),
    Tool(
        name="o365_get_site_drives",
        description="List document libraries in a SharePoint site.",
        inputSchema={
            "type": "object",
            "properties": {
                "site_id": {"type": "string", "description": "SharePoint site ID"},
            },
            "required": ["site_id"],
        },
    ),
    # ── Directory ─────────────────────────────────────────────────────
    Tool(
        name="o365_list_users",
        description="List users in the organization directory.",
        inputSchema={
            "type": "object",
            "properties": {
                "limit": {"type": "integer", "description": "Max users (default 25)", "default": 25},
            },
            "required": [],
        },
    ),
    Tool(
        name="o365_get_user_manager",
        description="Get a user's manager.",
        inputSchema={
            "type": "object",
            "properties": {
                "user_id": {"type": "string", "description": "Azure AD user ID"},
            },
            "required": ["user_id"],
        },
    ),

    # ── Rooms / Places ────────────────────────────────────────────────
    Tool(
        name="o365_list_rooms",
        description=(
            "List meeting rooms. Returns room name, capacity, building, and floor.\n"
            "Use the room name with o365_get_room_availability to check bookings."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "filter": {"type": "string", "description": "Filter rooms by name substring (case-insensitive, optional)"},
            },
            "required": [],
        },
    ),
    Tool(
        name="o365_get_room_availability",
        description=(
            "Get the availability schedule for one or more meeting rooms today.\n"
            "Returns time slots with free/busy status. Booking subject names are hidden by default.\n"
            "Pass room names (from o365_list_rooms) as a list."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "rooms": {
                    "type": "array",
                    "items": {"type": "string"},
                    "description": "Room names (or name substrings) to check",
                },
                "date": {"type": "string", "description": "Date in YYYY-MM-DD format (default: today)"},
            },
            "required": ["rooms"],
        },
    ),
]


# ---------------------------------------------------------------------------
# Tool execution
# ---------------------------------------------------------------------------


def _json_result(obj: Any) -> list[TextContent]:
    """Serialize a Pydantic model or dict to JSON TextContent."""
    if hasattr(obj, "model_dump"):
        text = json.dumps(obj.model_dump(), default=str, indent=2)
    elif isinstance(obj, dict):
        text = json.dumps(obj, default=str, indent=2)
    elif isinstance(obj, list):
        text = json.dumps(obj, default=str, indent=2)
    elif obj is None:
        text = "null"
    else:
        text = str(obj)
    return [TextContent(type="text", text=text)]


async def _handle_tool(graph: MsGraphInstance, name: str, args: dict[str, Any],
                       *, show_room_booking_names: bool = False) -> list[TextContent]:
    """Route a tool call to the appropriate handler."""

    # ── Profile ───────────────────────────────────────────────────────
    if name == "o365_get_profile":
        handler = await graph.get_profile_async()
        return _json_result(handler.me)

    # ── Mail ──────────────────────────────────────────────────────────
    if name == "o365_list_mail":
        mail = graph.get_mail()
        result = await mail.email_index_async(
            limit=args.get("limit", 10),
            skip=args.get("skip", 0),
        )
        return _json_result(result)

    if name == "o365_get_mail":
        mail = graph.get_mail()
        result = await mail.get_mail_async(email_id=args["email_id"])
        return _json_result(result)

    if name == "o365_get_mail_categories":
        mail = graph.get_mail()
        result = await mail.get_categories_async()
        return _json_result([c.model_dump() for c in result])

    # ── Calendar ──────────────────────────────────────────────────────
    if name == "o365_list_calendars":
        cal = graph.get_calendar()
        result = await cal.get_calendars_async()
        return _json_result(result)

    if name == "o365_get_events":
        from datetime import datetime
        cal = graph.get_calendar()
        start = datetime.fromisoformat(args["start_date"])
        end = datetime.fromisoformat(args["end_date"])
        result = await cal.get_events_async(
            start_date=start,
            end_date=end,
            limit=args.get("limit", 25),
        )
        return _json_result(result)

    if name == "o365_get_schedule":
        from datetime import datetime
        cal = graph.get_calendar()
        result = await cal.get_schedule_async(
            emails=args["emails"],
            start=datetime.fromisoformat(args["start"]),
            end=datetime.fromisoformat(args["end"]),
        )
        return _json_result(result)

    # ── Teams ─────────────────────────────────────────────────────────
    if name == "o365_list_teams":
        teams = graph.get_teams()
        result = await teams.get_joined_teams_async()
        return _json_result(result)

    if name == "o365_list_channels":
        teams = graph.get_teams()
        result = await teams.get_channels_async(args["team_id"])
        return _json_result(result)

    if name == "o365_get_channel_messages":
        teams = graph.get_teams()
        result = await teams.get_channel_messages_async(
            args["team_id"], args["channel_id"], limit=args.get("limit", 20),
        )
        return _json_result(result)

    if name == "o365_get_team_members":
        teams = graph.get_teams()
        result = await teams.get_team_members_async(args["team_id"])
        return _json_result(result)

    # ── Chats ─────────────────────────────────────────────────────────
    if name == "o365_list_chats":
        chat = graph.get_chat()
        result = await chat.get_chats_async(limit=args.get("limit", 25))
        return _json_result(result)

    if name == "o365_get_chat_messages":
        chat = graph.get_chat()
        result = await chat.get_chat_messages_async(
            args["chat_id"], limit=args.get("limit", 20),
        )
        return _json_result(result)

    if name == "o365_get_chat_members":
        chat = graph.get_chat()
        result = await chat.get_chat_members_async(args["chat_id"])
        return _json_result(result)

    # ── Files / OneDrive ──────────────────────────────────────────────
    if name == "o365_get_my_drive":
        files = graph.get_files()
        result = await files.get_my_drive_async()
        return _json_result(result)

    if name == "o365_list_drive_items":
        files = graph.get_files()
        folder_id = args.get("folder_id", "")
        drive_id = args.get("drive_id", "") or None
        limit = args.get("limit", 25)
        if folder_id:
            result = await files.get_folder_items_async(folder_id, drive_id=drive_id, limit=limit)
        else:
            result = await files.get_root_items_async(drive_id=drive_id, limit=limit)
        return _json_result(result)

    if name == "o365_get_file_content":
        import base64
        files = graph.get_files()
        drive_id = args.get("drive_id", "") or None
        content = await files.get_file_content_async(args["item_id"], drive_id=drive_id)
        if content is None:
            return [TextContent(type="text", text="File not found or is a folder.")]
        try:
            text = content.decode("utf-8")
            return [TextContent(type="text", text=text)]
        except UnicodeDecodeError:
            b64 = base64.b64encode(content).decode("ascii")
            return [TextContent(type="text", text=f"[binary file, {len(content)} bytes, base64]\n{b64}")]

    if name == "o365_search_files":
        files = graph.get_files()
        result = await files.search_items_async(
            args["query"], limit=args.get("limit", 10),
        )
        return _json_result(result)

    # ── SharePoint ────────────────────────────────────────────────────
    if name == "o365_search_sites":
        files = graph.get_files()
        result = await files.search_sites_async(args["query"])
        return _json_result(result)

    if name == "o365_get_site_drives":
        files = graph.get_files()
        result = await files.get_site_drives_async(args["site_id"])
        return _json_result(result)

    # ── Directory ─────────────────────────────────────────────────────
    if name == "o365_list_users":
        directory = graph.get_directory()
        result = await directory.get_users_async(limit=args.get("limit", 25))
        return _json_result(result)

    if name == "o365_get_user_manager":
        directory = graph.get_directory()
        result = await directory.get_user_manager_async(args["user_id"])
        return _json_result(result)

    # ── Rooms / Places ────────────────────────────────────────────────
    if name == "o365_list_rooms":
        from office_con.msgraph.places_handler import PlacesHandler
        ph = PlacesHandler(graph)
        rooms = await ph.get_rooms_async()
        name_filter = args.get("filter", "").lower()
        if name_filter:
            rooms = [r for r in rooms if name_filter in r.get("displayName", "").lower()]
        result = [
            {
                "name": r.get("displayName", ""),
                "email": r.get("emailAddress", ""),
                "capacity": r.get("capacity"),
                "building": r.get("building"),
                "floor": r.get("floorNumber"),
            }
            for r in rooms
        ]
        return _json_result(result)

    if name == "o365_get_room_availability":
        from office_con.msgraph.places_handler import PlacesHandler
        from office_con.msgraph.mailbox_settings_handler import MailboxSettingsHandler
        from datetime import datetime as _dt, timezone as _tz
        from zoneinfo import ZoneInfo as _ZI

        # Resolve user timezone
        mbs = MailboxSettingsHandler(graph)
        settings = await mbs.get_mailbox_settings_async()
        _WIN_TZ = {
            "W. Europe Standard Time": "Europe/Berlin",
            "Central European Standard Time": "Europe/Berlin",
            "Romance Standard Time": "Europe/Paris",
            "GMT Standard Time": "Europe/London",
            "Eastern Standard Time": "America/New_York",
            "Central Standard Time": "America/Chicago",
            "Pacific Standard Time": "America/Los_Angeles",
            "China Standard Time": "Asia/Shanghai",
            "India Standard Time": "Asia/Kolkata",
        }
        win_tz = settings.get("timeZone", "W. Europe Standard Time")
        iana_tz = _WIN_TZ.get(win_tz, win_tz)
        local_tz = _ZI(iana_tz)

        # Resolve date
        date_str = args.get("date", "")
        if date_str:
            target_date = _dt.strptime(date_str, "%Y-%m-%d").date()
        else:
            target_date = _dt.now(local_tz).date()

        # Find matching rooms by name
        ph = PlacesHandler(graph)
        all_rooms = await ph.get_rooms_async()
        room_queries = args.get("rooms", [])
        matched = []
        for rq in room_queries:
            rq_lower = rq.lower()
            for r in all_rooms:
                if rq_lower in r.get("displayName", "").lower() and r not in matched:
                    matched.append(r)

        if not matched:
            return [TextContent(type="text", text="No matching rooms found.")]

        emails = [r["emailAddress"] for r in matched]

        # Query schedule
        start_str = f"{target_date.isoformat()}T07:00:00"
        end_str = f"{target_date.isoformat()}T20:00:00"
        body = {
            "schedules": emails,
            "startTime": {"dateTime": start_str, "timeZone": win_tz},
            "endTime": {"dateTime": end_str, "timeZone": win_tz},
            "availabilityViewInterval": 30,
        }
        token = await graph.get_access_token_async()
        resp = await graph.run_async(
            url=graph.msg_endpoint + "me/calendar/getSchedule",
            method="POST", json=body, token=token,
        )
        schedules = resp.json().get("value", [])

        email_to_name = {r["emailAddress"].lower(): r["displayName"] for r in matched}
        result = []
        for sched in schedules:
            email = sched.get("scheduleId", "").lower()
            room_name = email_to_name.get(email, email)
            items = sched.get("scheduleItems", [])
            bookings = []
            for item in items:
                s = item.get("start", {}).get("dateTime", "")
                e = item.get("end", {}).get("dateTime", "")
                if s and e:
                    # Convert UTC → local
                    st = _dt.fromisoformat(s.rstrip("Z")).replace(tzinfo=_tz.utc).astimezone(local_tz)
                    en = _dt.fromisoformat(e.rstrip("Z")).replace(tzinfo=_tz.utc).astimezone(local_tz)
                    booking = {
                        "start": st.strftime("%H:%M"),
                        "end": en.strftime("%H:%M"),
                        "status": item.get("status", "busy"),
                    }
                    if show_room_booking_names:
                        booking["subject"] = item.get("subject", "")
                    bookings.append(booking)
            result.append({
                "room": room_name,
                "date": target_date.isoformat(),
                "timezone": iana_tz,
                "bookings": bookings,
                "free_slots": _compute_free_slots(bookings),
            })
        return _json_result(result)

    return [TextContent(type="text", text=f"Unknown tool: {name}")]


def _compute_free_slots(bookings: list[dict]) -> list[dict]:
    """Compute free 30-min slots between 07:00 and 20:00 from a list of bookings."""
    busy = set()
    for b in bookings:
        start_h, start_m = map(int, b["start"].split(":"))
        end_h, end_m = map(int, b["end"].split(":"))
        t = start_h * 60 + start_m
        end = end_h * 60 + end_m
        while t < end:
            busy.add(t)
            t += 30

    free = []
    t = 7 * 60
    while t < 20 * 60:
        if t not in busy:
            h, m = divmod(t, 60)
            free.append({"start": f"{h:02d}:{m:02d}", "end": f"{h:02d}:{m + 30:02d}" if m == 0 else f"{h + 1:02d}:00"})
        t += 30

    # Merge consecutive free slots
    if not free:
        return []
    merged = [free[0].copy()]
    for slot in free[1:]:
        if merged[-1]["end"] == slot["start"]:
            merged[-1]["end"] = slot["end"]
        else:
            merged.append(slot.copy())
    return merged


# ---------------------------------------------------------------------------
# Server setup
# ---------------------------------------------------------------------------


def create_server(keyfile: str) -> tuple[Server, str]:
    """Create the MCP server and return (server, keyfile) for deferred graph init."""
    mcp = Server(
        "office-365-mcp",
        instructions=(
            "Office 365 MCP Server — read-only access to Microsoft 365.\n"
            "Provides tools for reading mail, calendar, teams, chats, files, "
            "directory, and profile data via Microsoft Graph API."
        ),
    )

    # We defer graph creation to first tool call (needs async)
    _state: dict[str, Any] = {"graph": None, "keyfile": keyfile}

    async def _get_graph() -> MsGraphInstance:
        if _state["graph"] is None:
            _state["graph"] = await _create_graph(_state["keyfile"])
        return _state["graph"]

    @mcp.list_tools()
    async def list_tools() -> list[Tool]:
        return TOOLS

    @mcp.call_tool()
    async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
        try:
            graph = await _get_graph()
            return await _handle_tool(graph, name, arguments)
        except Exception:
            logger.exception("Tool %s failed", name)
            return [TextContent(type="text", text=f"Error: tool '{name}' failed. Check server logs for details.")]

    return mcp, keyfile


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


async def main(keyfile: str) -> None:
    """Run the MCP server over stdio."""
    print("Office 365 MCP Server starting...", file=sys.stderr)
    mcp, _ = create_server(keyfile)
    async with stdio_server() as (read_stream, write_stream):
        await mcp.run(read_stream, write_stream, mcp.create_initialization_options())


def cli() -> None:
    """CLI entry point for the MCP server."""
    parser = argparse.ArgumentParser(description="Office 365 MCP Server")
    parser.add_argument("--keyfile", required=True, help="Path to JSON token file")
    parsed = parser.parse_args()
    asyncio.run(main(parsed.keyfile))


if __name__ == "__main__":
    cli()

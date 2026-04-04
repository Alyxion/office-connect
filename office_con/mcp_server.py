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


async def _handle_tool(graph: MsGraphInstance, name: str, args: dict[str, Any]) -> list[TextContent]:
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

    return [TextContent(type="text", text=f"Unknown tool: {name}")]


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

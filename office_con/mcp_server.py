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
import base64
import json
import logging
import os
import sys
from pathlib import Path
from typing import Any

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import TextContent, Tool

from office_con.mcp_permissions import (
    DEFAULT_LEVEL,
    ENV_VAR as PERMISSION_ENV_VAR,
    PermissionLevel,
    level_allows,
    parse_level,
    resolve_level,
)
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
    Tool(
        name="o365_search_mail",
        description=(
            "Search emails efficiently using MS Graph KQL ($search). Prefer this over "
            "paging through o365_list_mail when the user gives ANY criterion "
            "(sender name, subject keyword, date range, etc.). All structured params "
            "are AND-ed. Dates are inclusive. Returns recent matches across all folders."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "from": {"type": "string", "description": "Sender (name or email). Display names work: 'Julia Bader'."},
                "to": {"type": "string", "description": "Recipient (name or email)"},
                "subject": {"type": "string", "description": "Substring in subject"},
                "body": {"type": "string", "description": "Free-text match in body/subject"},
                "since": {"type": "string", "description": "Received on or after (YYYY-MM-DD)"},
                "until": {"type": "string", "description": "Received on or before (YYYY-MM-DD)"},
                "has_attachments": {"type": "boolean", "description": "Only mails with attachments"},
                "query": {"type": "string", "description": "Raw KQL (overrides/augments other params)"},
                "limit": {"type": "integer", "description": "Max results (default 25, Graph caps at 250)", "default": 25},
            },
            "required": [],
        },
    ),
    # ── Mail writes: drafts (DRAFTS tier) ─────────────────────────────
    Tool(
        name="o365_create_mail_draft",
        description=(
            "Create a draft email. NOT sent — the user reviews and sends manually.\n\n"
            "Body: pass EITHER 'body' (inline string) OR 'body_path' (path to a file "
            "on the server's filesystem, read as UTF-8). body_path requires the "
            "server to have an attachment-root configured; disabled by default.\n\n"
            "Attachments: each item provides EITHER 'content_base64' OR 'path'. "
            "path mode has the same safety gating as body_path. Large files go "
            "through MS Graph's upload session automatically."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "to": {"type": "array", "items": {"type": "string"}, "description": "Recipient email addresses"},
                "subject": {"type": "string", "description": "Email subject"},
                "body": {"type": "string", "description": "Inline email body (alternative to body_path)"},
                "body_path": {"type": "string", "description": "Filesystem path to a body file (alternative to body). Disabled unless an attachment root is configured."},
                "is_html": {"type": "boolean", "description": "Body is HTML (default false)", "default": False},
                "cc": {"type": "array", "items": {"type": "string"}, "description": "CC recipients"},
                "bcc": {"type": "array", "items": {"type": "string"}, "description": "BCC recipients"},
                "attachments": {
                    "type": "array",
                    "description": "Files to attach. Each item: {name, content_type?, content_base64} OR {name, content_type?, path}.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string", "description": "Filename, e.g. report.pdf"},
                            "content_type": {"type": "string", "description": "MIME type; default application/octet-stream"},
                            "content_base64": {"type": "string", "description": "Base64 bytes (alternative to path)"},
                            "path": {"type": "string", "description": "Path on the MCP server's filesystem (alternative to content_base64)"},
                        },
                        "required": ["name"],
                    },
                },
            },
            "required": ["to", "subject"],
        },
    ),
    Tool(
        name="o365_update_mail_draft",
        description=(
            "Update an existing draft email. The target message MUST be a draft — "
            "attempts to modify a sent/received message are refused.\n\n"
            "Body: pass EITHER 'body' or 'body_path' (same semantics as "
            "o365_create_mail_draft).\n\n"
            "Attachments semantics: omit to keep existing, [] to clear all, "
            "[items] to REPLACE. Each item: {name, content_base64} OR "
            "{name, path}. Uses MS Graph upload session for files >3 MB."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Draft message id"},
                "to": {"type": "array", "items": {"type": "string"}, "description": "Recipient email addresses"},
                "subject": {"type": "string", "description": "Email subject"},
                "body": {"type": "string", "description": "Inline email body (alternative to body_path)"},
                "body_path": {"type": "string", "description": "Filesystem path to a body file (alternative to body)"},
                "is_html": {"type": "boolean", "description": "Body is HTML (default false)", "default": False},
                "cc": {"type": "array", "items": {"type": "string"}, "description": "CC recipients"},
                "bcc": {"type": "array", "items": {"type": "string"}, "description": "BCC recipients"},
                "attachments": {
                    "type": "array",
                    "description": "Replace draft attachments; omit to keep. Each: {name, content_type?, content_base64 OR path}.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "content_type": {"type": "string"},
                            "content_base64": {"type": "string"},
                            "path": {"type": "string"},
                        },
                        "required": ["name"],
                    },
                },
            },
            "required": ["email_id", "to", "subject"],
        },
    ),
    # ── Mail writes: send / mutate (ALL tier) ─────────────────────────
    Tool(
        name="o365_send_mail",
        description=(
            "Send an email immediately (no draft step). Requires ALL permission level.\n\n"
            "Body: pass EITHER 'body' or 'body_path'.\n\n"
            "Attachments: each item {name, content_type?, content_base64 OR path}. "
            "When attachments are present the server internally drafts → uploads "
            "in parallel (with upload session for >3 MB) → sends, to support large files."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "to": {"type": "array", "items": {"type": "string"}, "description": "Recipient email addresses"},
                "subject": {"type": "string", "description": "Email subject"},
                "body": {"type": "string", "description": "Inline email body (alternative to body_path)"},
                "body_path": {"type": "string", "description": "Filesystem path to a body file (alternative to body)"},
                "is_html": {"type": "boolean", "description": "Body is HTML (default false)", "default": False},
                "cc": {"type": "array", "items": {"type": "string"}, "description": "CC recipients"},
                "bcc": {"type": "array", "items": {"type": "string"}, "description": "BCC recipients"},
                "save_to_sent_items": {"type": "boolean", "description": "Save to Sent Items (default true)", "default": True},
                "attachments": {
                    "type": "array",
                    "description": "Files to attach. Each: {name, content_type?, content_base64 OR path}.",
                    "items": {
                        "type": "object",
                        "properties": {
                            "name": {"type": "string"},
                            "content_type": {"type": "string"},
                            "content_base64": {"type": "string"},
                            "path": {"type": "string"},
                        },
                        "required": ["name"],
                    },
                },
            },
            "required": ["to", "subject"],
        },
    ),
    Tool(
        name="o365_send_mail_draft",
        description="Send an existing draft email by its message id. Requires ALL permission level.",
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Draft message id"},
            },
            "required": ["email_id"],
        },
    ),
    Tool(
        name="o365_delete_mail",
        description="Soft-delete an email (moves to Deleted Items). Requires ALL permission level.",
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Message id to delete"},
            },
            "required": ["email_id"],
        },
    ),
    Tool(
        name="o365_move_mail",
        description=(
            "Move an email to another folder. Destination may be a folder id or a well-known "
            "name (inbox, deleteditems, archive, junkemail, drafts, sentitems). Requires ALL."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Message id to move"},
                "destination": {"type": "string", "description": "Destination folder id or well-known name"},
            },
            "required": ["email_id", "destination"],
        },
    ),
    Tool(
        name="o365_flag_mail_read",
        description="Mark a message as read or unread. Requires ALL permission level.",
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Message id"},
                "is_read": {"type": "boolean", "description": "true = read, false = unread"},
            },
            "required": ["email_id", "is_read"],
        },
    ),
    Tool(
        name="o365_set_mail_categories",
        description="Set the Outlook categories on a message (replaces existing). Requires ALL.",
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Message id"},
                "categories": {"type": "array", "items": {"type": "string"}, "description": "Category names"},
            },
            "required": ["email_id", "categories"],
        },
    ),
    # ── Calendar ──────────────────────────────────────────────────────
    Tool(
        name="o365_list_calendars",
        description="List the user's calendars.",
        inputSchema={"type": "object", "properties": {}, "required": []},
    ),
    Tool(
        name="o365_get_events",
        description=(
            "Get calendar events within a date range. Both bounds are INCLUSIVE of the "
            "whole calendar day when a date-only value (no time) is passed — e.g. "
            "start=2026-04-21, end=2026-04-21 returns all events on 2026-04-21."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "start_date": {
                    "type": "string",
                    "description": (
                        "Start date. ISO 8601 date (2026-04-21) or datetime "
                        "(2026-04-21T09:00:00). Date-only = start of that day."
                    ),
                },
                "end_date": {
                    "type": "string",
                    "description": (
                        "End date. ISO 8601 date (2026-04-21) or datetime. "
                        "Date-only = end of that day (23:59:59), inclusive."
                    ),
                },
                "limit": {"type": "integer", "description": "Max events to return (default 25)", "default": 25},
            },
            "required": ["start_date", "end_date"],
        },
    ),
    Tool(
        name="o365_search_events",
        description=(
            "Search calendar events by subject/attendee/organizer/date using MS Graph "
            "$filter on /me/events. Prefer over o365_get_events when given ANY textual "
            "criterion. Note: body content is not searchable here (Graph limitation); "
            "recurring events are NOT expanded — use o365_get_events with a date range "
            "for pure occurrence listing."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {"type": "string", "description": "Substring in event subject (contains match)"},
                "attendee": {"type": "string", "description": "Attendee email (exact) or name (contains)"},
                "organizer": {"type": "string", "description": "Organizer email (exact) or name (contains)"},
                "since": {"type": "string", "description": "Start on/after (YYYY-MM-DD)"},
                "until": {"type": "string", "description": "Start on/before (YYYY-MM-DD)"},
                "limit": {"type": "integer", "description": "Max results (default 25)", "default": 25},
            },
            "required": [],
        },
    ),
    Tool(
        name="o365_create_event",
        description=(
            "Create a calendar event. Note: if attendees are set, MS Graph sends invites "
            "immediately. Requires ALL permission level."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "subject": {"type": "string", "description": "Event subject"},
                "start": {"type": "string", "description": "Start datetime (ISO 8601)"},
                "end": {"type": "string", "description": "End datetime (ISO 8601)"},
                "body": {"type": "string", "description": "Event body/description"},
                "is_html": {"type": "boolean", "description": "Body is HTML (default false)", "default": False},
                "location": {"type": "string", "description": "Location display name"},
                "attendees": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "email": {"type": "string"},
                            "name": {"type": "string"},
                        },
                        "required": ["email"],
                    },
                    "description": "Attendees; each {email, name?}",
                },
                "is_all_day": {"type": "boolean", "description": "All-day event (default false)", "default": False},
                "calendar_id": {"type": "string", "description": "Calendar id (omit for default)"},
            },
            "required": ["subject", "start", "end"],
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
    Tool(
        name="o365_search_messages",
        description=(
            "Search Teams channel messages AND 1:1/group chat messages in a single call "
            "via Graph's /search/query endpoint with entity type 'chatMessage'. Prefer "
            "this over paging through o365_get_chat_messages / o365_get_channel_messages. "
            "Requires Chat.Read scope for chats; channel hits also need ChannelMessage.Read.All."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "Search text or raw KQL. Required."},
                "from": {"type": "string", "description": "Sender name or email (added to KQL)"},
                "since": {"type": "string", "description": "Created on/after (YYYY-MM-DD)"},
                "until": {"type": "string", "description": "Created on/before (YYYY-MM-DD)"},
                "limit": {"type": "integer", "description": "Max results (default 25, Graph caps at 500)", "default": 25},
            },
            "required": ["query"],
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
        name="o365_peek_drive_file",
        description=(
            "Peek at a OneDrive/SharePoint file (PDF, xlsx, docx) WITHOUT returning "
            "the full bytes. Returns a compact summary: page/sheet/paragraph count, "
            "first-page text, metadata, and (for PDFs) a rendered PNG of page 1. "
            "Much cheaper than o365_get_file_content for assessing relevance. "
            "Unsupported types return a message with byte size so the caller can "
            "decide whether to fetch."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "item_id": {"type": "string", "description": "Drive item ID"},
                "drive_id": {"type": "string", "description": "Drive ID (omit for default OneDrive)"},
                "pages": {"type": "integer", "description": "PDF: number of pages to extract/render (default 1)", "default": 1},
                "render": {"type": "boolean", "description": "PDF: include rendered PNG (default true)", "default": True},
                "max_rows": {"type": "integer", "description": "xlsx: max rows to sample (default 30)", "default": 30},
                "max_paragraphs": {"type": "integer", "description": "docx: max paragraphs to sample (default 30)", "default": 30},
                "all_sheets": {"type": "boolean", "description": "xlsx: include all sheets instead of just active (default false)", "default": False},
            },
            "required": ["item_id"],
        },
    ),
    Tool(
        name="o365_peek_mail_attachment",
        description=(
            "Peek at an email attachment (PDF, xlsx, docx) WITHOUT returning full "
            "bytes. Use IDs from o365_get_mail (which lists attachments with their "
            "IDs). Returns compact summary with metadata, first-page text, and a "
            "rendered PNG for PDFs. Prefer this over fetching the full attachment "
            "when deciding whether an email is relevant."
        ),
        inputSchema={
            "type": "object",
            "properties": {
                "email_id": {"type": "string", "description": "Message id"},
                "attachment_id": {"type": "string", "description": "Attachment id from o365_get_mail"},
                "pages": {"type": "integer", "description": "PDF: pages to extract/render (default 1)", "default": 1},
                "render": {"type": "boolean", "description": "PDF: include PNG (default true)", "default": True},
                "max_rows": {"type": "integer", "description": "xlsx: max rows (default 30)", "default": 30},
                "max_paragraphs": {"type": "integer", "description": "docx: max paragraphs (default 30)", "default": 30},
                "all_sheets": {"type": "boolean", "description": "xlsx: include all sheets (default false)", "default": False},
            },
            "required": ["email_id", "attachment_id"],
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
# Permission classification
# ---------------------------------------------------------------------------
#
# Every tool in ``TOOLS`` MUST appear in this table. Tools without an entry
# are denied by ``_require_allowed`` (fail-closed). A test in
# ``tests/test_mcp_permissions.py`` enforces full coverage.

TOOL_PERMISSIONS: dict[str, PermissionLevel] = {
    # Profile
    "o365_get_profile": PermissionLevel.READ_ONLY,
    # Mail — read
    "o365_list_mail": PermissionLevel.READ_ONLY,
    "o365_get_mail": PermissionLevel.READ_ONLY,
    "o365_get_mail_categories": PermissionLevel.READ_ONLY,
    "o365_search_mail": PermissionLevel.READ_ONLY,
    # Mail — draft lifecycle (target is always a draft)
    "o365_create_mail_draft": PermissionLevel.DRAFTS,
    "o365_update_mail_draft": PermissionLevel.DRAFTS,
    # Mail — sending and mutation of real messages
    "o365_send_mail": PermissionLevel.ALL,
    "o365_send_mail_draft": PermissionLevel.ALL,
    "o365_delete_mail": PermissionLevel.ALL,
    "o365_move_mail": PermissionLevel.ALL,
    "o365_flag_mail_read": PermissionLevel.ALL,
    "o365_set_mail_categories": PermissionLevel.ALL,
    # Calendar
    "o365_list_calendars": PermissionLevel.READ_ONLY,
    "o365_get_events": PermissionLevel.READ_ONLY,
    "o365_get_schedule": PermissionLevel.READ_ONLY,
    "o365_search_events": PermissionLevel.READ_ONLY,
    "o365_create_event": PermissionLevel.ALL,
    # Teams
    "o365_list_teams": PermissionLevel.READ_ONLY,
    "o365_list_channels": PermissionLevel.READ_ONLY,
    "o365_get_channel_messages": PermissionLevel.READ_ONLY,
    "o365_get_team_members": PermissionLevel.READ_ONLY,
    # Chats
    "o365_list_chats": PermissionLevel.READ_ONLY,
    "o365_get_chat_messages": PermissionLevel.READ_ONLY,
    "o365_get_chat_members": PermissionLevel.READ_ONLY,
    "o365_search_messages": PermissionLevel.READ_ONLY,
    # Files / OneDrive / SharePoint
    "o365_get_my_drive": PermissionLevel.READ_ONLY,
    "o365_list_drive_items": PermissionLevel.READ_ONLY,
    "o365_get_file_content": PermissionLevel.READ_ONLY,
    "o365_peek_drive_file": PermissionLevel.READ_ONLY,
    "o365_peek_mail_attachment": PermissionLevel.READ_ONLY,
    "o365_search_files": PermissionLevel.READ_ONLY,
    "o365_search_sites": PermissionLevel.READ_ONLY,
    "o365_get_site_drives": PermissionLevel.READ_ONLY,
    # Directory
    "o365_list_users": PermissionLevel.READ_ONLY,
    "o365_get_user_manager": PermissionLevel.READ_ONLY,
    # Rooms / Places
    "o365_list_rooms": PermissionLevel.READ_ONLY,
    "o365_get_room_availability": PermissionLevel.READ_ONLY,
}


class PermissionDenied(Exception):
    """Raised when a tool call is not permitted at the configured level."""


def filter_tools(level: PermissionLevel) -> list[Tool]:
    """Return the subset of ``TOOLS`` allowed at the given permission level."""
    return [
        t for t in TOOLS
        if t.name in TOOL_PERMISSIONS
        and level_allows(TOOL_PERMISSIONS[t.name], level)
    ]


def _require_allowed(tool_name: str, level: PermissionLevel) -> None:
    """Raise ``PermissionDenied`` unless the tool is permitted at ``level``.

    Fail-closed: an unknown tool name (not in ``TOOL_PERMISSIONS``) is denied.
    """
    required = TOOL_PERMISSIONS.get(tool_name)
    if required is None:
        logger.warning(
            "permission denied: unclassified tool name=%r level=%s", tool_name, level.value,
        )
        raise PermissionDenied(
            f"Tool '{tool_name}' is not classified and is denied by default."
        )
    if not level_allows(required, level):
        logger.warning(
            "permission denied: tool=%s required=%s configured=%s",
            tool_name, required.value, level.value,
        )
        raise PermissionDenied(
            f"Tool '{tool_name}' requires permission level '{required.value}'; "
            f"server is configured for '{level.value}'."
        )


# ---------------------------------------------------------------------------
# Attachment-path safety
# ---------------------------------------------------------------------------
#
# Path-based attachments let the MCP server read local files instead of
# receiving base64 bytes through the tool call. This is a significant
# privilege (the server reads its own filesystem), so the safety model is:
#
#   * FAIL-CLOSED DEFAULT — zero roots configured means ALL path attachments
#     are rejected. A shared/server deployment keeps the defaults and
#     literally cannot read any file.
#   * EXPLICIT OPT-IN via --attachment-root (repeatable) or the env var
#     OFFICE_CONNECT_ATTACHMENT_ROOTS (os.pathsep-separated).
#   * STARTUP VISIBILITY — the server logs the active mode clearly so
#     operators can't silently misconfigure.
#   * PER-REQUEST GUARD — every candidate path is resolved with strict=True
#     (follows symlinks, errors on missing), verified to be a regular file,
#     and checked against the allowed roots via ``Path.relative_to`` which
#     is immune to ``..`` escapes, null bytes are rejected explicitly, and
#     file size is capped.

ATTACHMENT_ROOTS_ENV = "OFFICE_CONNECT_ATTACHMENT_ROOTS"
MAX_ATTACHMENT_BYTES_ENV = "OFFICE_CONNECT_MAX_ATTACHMENT_BYTES"
DEFAULT_MAX_ATTACHMENT_BYTES = 150 * 1024 * 1024  # 150 MB
SIMPLE_ATTACHMENT_LIMIT = 3 * 1024 * 1024  # Graph simple-endpoint cap
UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024  # 5 MiB; Graph requires multiples of 320 KiB


class AttachmentPathError(Exception):
    """A path-based attachment failed the safety check."""


def _parse_attachment_roots(cli_values: list[str] | None) -> list[Path]:
    """Parse attachment-root CLI args + env var into a list of resolved dirs.

    Raises ``ValueError`` on any invalid entry — we want configuration errors
    to fail loudly at startup, not silently drop a root.
    """
    raw: list[str] = []
    if cli_values:
        raw.extend(cli_values)
    env = os.environ.get(ATTACHMENT_ROOTS_ENV, "")
    if env:
        raw.extend(env.split(os.pathsep))

    roots: list[Path] = []
    for v in raw:
        v = v.strip()
        if not v:
            continue
        p = Path(v)
        if not p.is_absolute():
            raise ValueError(f"attachment root {v!r} must be an absolute path")
        try:
            resolved = p.resolve(strict=True)
        except FileNotFoundError:
            raise ValueError(f"attachment root {v!r} does not exist") from None
        if not resolved.is_dir():
            raise ValueError(f"attachment root {v!r} is not a directory")
        roots.append(resolved)
    return roots


def _parse_max_attachment_bytes() -> int:
    raw = os.environ.get(MAX_ATTACHMENT_BYTES_ENV)
    if not raw:
        return DEFAULT_MAX_ATTACHMENT_BYTES
    try:
        v = int(raw)
    except ValueError:
        raise ValueError(
            f"{MAX_ATTACHMENT_BYTES_ENV}={raw!r} is not a valid integer byte count"
        ) from None
    if v <= 0:
        raise ValueError(f"{MAX_ATTACHMENT_BYTES_ENV} must be positive, got {v}")
    return v


def _resolve_safe_attachment_path(
    path_str: str, roots: list[Path], max_bytes: int,
) -> Path:
    """Resolve and validate a user-supplied attachment path. Never returns
    unsafely — raises AttachmentPathError on any violation.

    Catches: disabled mode, empty/null paths, nonexistent files, symlink
    escapes, non-regular files (directories, devices, fifos, sockets), paths
    outside every allowed root, oversized files.
    """
    if not roots:
        raise AttachmentPathError(
            "path-based attachments are DISABLED on this server. "
            "To enable, start with --attachment-root <dir> or set the "
            f"{ATTACHMENT_ROOTS_ENV} environment variable."
        )
    if not path_str or not path_str.strip():
        raise AttachmentPathError("attachment path is empty")
    if "\x00" in path_str:
        raise AttachmentPathError("attachment path contains a null byte")

    candidate = Path(path_str)
    try:
        # strict=True: raises if the file doesn't exist; follows symlinks.
        resolved = candidate.resolve(strict=True)
    except FileNotFoundError:
        raise AttachmentPathError(f"attachment not found: {path_str!r}")
    except (OSError, RuntimeError) as exc:
        raise AttachmentPathError(f"could not resolve {path_str!r}: {exc}")

    # Must be a regular file (not dir, symlink target dir, device, socket).
    try:
        st = resolved.stat()
    except OSError as exc:
        raise AttachmentPathError(f"could not stat {path_str!r}: {exc}")
    import stat as _stat
    if not _stat.S_ISREG(st.st_mode):
        raise AttachmentPathError(
            f"not a regular file: {path_str!r} (type=0o{st.st_mode:o})"
        )

    # Escape-proof root check: relative_to raises ValueError if outside.
    inside_root = False
    for root in roots:
        try:
            resolved.relative_to(root)
            inside_root = True
            break
        except ValueError:
            continue
    if not inside_root:
        raise AttachmentPathError(
            f"attachment path {path_str!r} is outside every configured "
            f"safe root ({', '.join(str(r) for r in roots)})"
        )

    if st.st_size > max_bytes:
        raise AttachmentPathError(
            f"attachment {resolved.name} is {st.st_size} bytes; "
            f"exceeds configured max of {max_bytes}"
        )
    return resolved


# ---------------------------------------------------------------------------
# Tool execution
# ---------------------------------------------------------------------------


def _kql_value(v: str) -> str:
    """Quote a KQL value if it contains whitespace or colons."""
    v = v.strip()
    if any(c.isspace() for c in v) or ":" in v:
        return '"' + v.replace('"', '\\"') + '"'
    return v


def _kql_range(field: str, since: str | None, until: str | None) -> str | None:
    """Build a KQL date-range fragment or None if no bounds."""
    if since and until:
        return f"{field}:{since}..{until}"
    if since:
        return f"{field}>={since}"
    if until:
        return f"{field}<={until}"
    return None


def _build_mail_kql(args: dict) -> str:
    parts: list[str] = []
    if v := args.get("from"):
        parts.append(f"from:{_kql_value(v)}")
    if v := args.get("to"):
        parts.append(f"to:{_kql_value(v)}")
    if v := args.get("subject"):
        parts.append(f"subject:{_kql_value(v)}")
    if v := args.get("body"):
        parts.append(_kql_value(v))  # free text matches body+subject+people
    if r := _kql_range("received", args.get("since"), args.get("until")):
        parts.append(r)
    if args.get("has_attachments") is True:
        parts.append("hasAttachment:true")
    if v := args.get("query"):
        parts.append(v.strip())
    return " ".join(parts).strip()


def _build_messages_kql(args: dict) -> str:
    parts: list[str] = [args["query"].strip()]
    if v := args.get("from"):
        parts.append(f"from:{_kql_value(v)}")
    if r := _kql_range("created", args.get("since"), args.get("until")):
        parts.append(r)
    return " ".join(p for p in parts if p).strip()


def _decode_attachments(
    raw: list[dict] | None,
    *,
    attachment_roots: list[Path],
    max_attachment_bytes: int,
):
    """Turn MCP-input attachment dicts into OfficeMailAttachment objects.

    Each input item accepts ONE of:
      * ``content_base64`` — Base64-encoded bytes (always available).
      * ``path`` — filesystem path; requires a configured attachment root and
        passes through ``_resolve_safe_attachment_path``. Disabled by default.

    ``name`` is required. ``content_type`` defaults to application/octet-stream.

    Returns ``None`` if ``raw`` is ``None`` (caller distinguishes "no change"
    from "empty list"). Raises ``AttachmentPathError`` or ``ValueError`` with
    a user-safe message on validation failure.
    """
    if raw is None:
        return None
    from office_con.msgraph.mail_handler import OfficeMailAttachment
    out = []
    for idx, a in enumerate(raw):
        name = a.get("name")
        if not name:
            raise ValueError(f"attachment[{idx}] missing 'name'")
        content_type = a.get("content_type") or "application/octet-stream"
        has_path = bool(a.get("path"))
        has_b64 = bool(a.get("content_base64"))
        if has_path and has_b64:
            raise ValueError(
                f"attachment[{idx}] {name!r}: provide exactly one of "
                "'path' or 'content_base64', not both"
            )
        if not has_path and not has_b64:
            raise ValueError(
                f"attachment[{idx}] {name!r}: provide either 'path' or 'content_base64'"
            )
        if has_path:
            safe = _resolve_safe_attachment_path(
                a["path"], attachment_roots, max_attachment_bytes,
            )
            content_bytes = safe.read_bytes()
        else:
            try:
                content_bytes = base64.b64decode(a["content_base64"], validate=True)
            except Exception as exc:
                raise ValueError(
                    f"attachment[{idx}] {name!r}: invalid base64: {exc}"
                ) from None
            if len(content_bytes) > max_attachment_bytes:
                raise ValueError(
                    f"attachment[{idx}] {name!r}: {len(content_bytes)} bytes "
                    f"exceeds max {max_attachment_bytes}"
                )
        out.append(OfficeMailAttachment(
            name=name,
            content_type=content_type,
            content_bytes=content_bytes,
        ))
    return out


# ---------------------------------------------------------------------------
# Attachment upload (parallel, upload-session for large files)
# ---------------------------------------------------------------------------


async def _upload_one_attachment_async(
    graph: MsGraphInstance, message_id: str, att, token: str,
) -> bool:
    """Upload a single attachment. Uses the simple endpoint for small files
    (<=3 MB raw) and an upload session with chunked PUT for larger ones."""
    data = att.content_bytes or b""
    size = len(data)
    if size == 0:
        return False

    if size <= SIMPLE_ATTACHMENT_LIMIT:
        url = f"{graph.msg_endpoint}me/messages/{message_id}/attachments"
        payload = {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": att.name,
            "contentType": att.content_type,
            "contentBytes": base64.b64encode(data).decode(),
        }
        resp = await graph.run_async(url=url, method="POST", json=payload, token=token)
        return resp is not None and resp.status_code == 201

    # Large: upload session
    import aiohttp
    session_url = (
        f"{graph.msg_endpoint}me/messages/{message_id}"
        "/attachments/createUploadSession"
    )
    body = {
        "AttachmentItem": {
            "attachmentType": "file",
            "name": att.name,
            "size": size,
            "contentType": att.content_type,
        }
    }
    resp = await graph.run_async(url=session_url, method="POST", json=body, token=token)
    if resp is None or resp.status_code not in (200, 201):
        logger.warning("createUploadSession failed for %s: status=%s",
                       att.name, resp.status_code if resp else None)
        return False
    upload_url = resp.json().get("uploadUrl")
    if not upload_url:
        return False

    # Chunked PUT; upload URL is pre-signed — no bearer token.
    async with aiohttp.ClientSession() as client:
        offset = 0
        while offset < size:
            end = min(offset + UPLOAD_CHUNK_SIZE, size)
            chunk = data[offset:end]
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end - 1}/{size}",
            }
            async with client.put(upload_url, data=chunk, headers=headers) as put_resp:
                if put_resp.status not in (200, 201, 202):
                    logger.warning(
                        "chunk PUT failed for %s at offset=%d status=%s",
                        att.name, offset, put_resp.status,
                    )
                    return False
            offset = end
    return True


async def _upload_attachments_parallel_async(
    graph: MsGraphInstance, message_id: str, attachments: list,
) -> tuple[int, int]:
    """Upload attachments concurrently. Returns (succeeded, total)."""
    if not attachments:
        return 0, 0
    token = await graph.get_access_token_async()
    results = await asyncio.gather(
        *(_upload_one_attachment_async(graph, message_id, a, token)
          for a in attachments),
        return_exceptions=True,
    )
    succeeded = sum(1 for r in results if r is True)
    return succeeded, len(attachments)


def _resolve_body_text(
    args: dict,
    *,
    attachment_roots: list[Path],
    max_attachment_bytes: int,
) -> str:
    """Return the message body. Accepts ``body`` (literal) or ``body_path``
    (filesystem path, gated by the same attachment-root allowlist). Exactly
    one must be provided. Raises ValueError / AttachmentPathError on problems.

    Decoded as UTF-8 with errors='replace' so a near-miss encoding doesn't
    fail the whole draft; the user sees a character or two substituted
    rather than a cryptic error.
    """
    has_body = args.get("body") is not None
    has_path = bool(args.get("body_path"))
    if has_body and has_path:
        raise ValueError("provide exactly one of 'body' or 'body_path', not both")
    if not has_body and not has_path:
        raise ValueError("missing message body: pass 'body' or 'body_path'")
    if has_body:
        return args["body"]
    safe = _resolve_safe_attachment_path(
        args["body_path"], attachment_roots, max_attachment_bytes,
    )
    return safe.read_bytes().decode("utf-8", errors="replace")


async def _clear_attachments_parallel_async(
    graph: MsGraphInstance, message_id: str,
) -> None:
    """Delete all attachments on a message in parallel."""
    token = await graph.get_access_token_async()
    url = f"{graph.msg_endpoint}me/messages/{message_id}/attachments"
    resp = await graph.run_async(url=url, token=token)
    if resp is None or resp.status_code != 200:
        return
    ids = [a.get("id") for a in resp.json().get("value", []) if a.get("id")]
    if not ids:
        return
    await asyncio.gather(
        *(graph.run_async(url=f"{url}/{aid}", method="DELETE", token=token)
          for aid in ids),
        return_exceptions=True,
    )


async def _is_draft(graph: MsGraphInstance, message_id: str) -> bool | None:
    """Return True/False for the message's isDraft flag, or None if not found."""
    token = await graph.get_access_token_async()
    resp = await graph.run_async(
        url=f"{graph.msg_endpoint}me/messages/{message_id}?$select=isDraft",
        token=token,
    )
    if resp is None or resp.status_code != 200:
        return None
    return bool(resp.json().get("isDraft", False))


async def _fetch_drive_item_meta(
    graph: MsGraphInstance, item_id: str, drive_id: str | None,
) -> dict | None:
    """Look up a drive item's metadata (name, mimeType, size) before download."""
    token = await graph.get_access_token_async()
    base = f"drives/{drive_id}/items" if drive_id else "me/drive/items"
    url = f"{graph.msg_endpoint}{base}/{item_id}?$select=name,file,size"
    resp = await graph.run_async(url=url, token=token)
    if resp is None or resp.status_code != 200:
        return None
    return resp.json()


def _peek_result(
    content: bytes,
    args: dict,
    *,
    name: str | None,
    content_type: str | None,
    size_from_meta: int | None = None,
) -> list:
    """Run peek_document, return a list of MCP content items.

    Text peek is a JSON TextContent; PDF renders become ImageContent items
    (stripped from the JSON to keep it compact) so MCP clients can display
    them natively.
    """
    from mcp.types import ImageContent
    from office_con.peek import peek_document

    peek = peek_document(
        content,
        name=name,
        content_type=content_type,
        pages=args.get("pages", 1),
        render=args.get("render", True),
        max_rows=args.get("max_rows", 30),
        max_paragraphs=args.get("max_paragraphs", 30),
        all_sheets=args.get("all_sheets", False),
    )
    peek.setdefault("name", name)
    peek.setdefault("content_type", content_type)
    peek["byte_size"] = len(content)
    if size_from_meta is not None:
        peek["size_from_metadata"] = size_from_meta

    # Split rendered PNGs into ImageContent items, keep a compact JSON.
    images: list = []
    renders = peek.pop("renders", None)
    if renders:
        compact_renders = []
        for r in renders:
            data_b64 = r.get("png_base64")
            if data_b64:
                images.append(ImageContent(
                    type="image", data=data_b64, mimeType="image/png",
                ))
            compact_renders.append({
                "index": r.get("index"),
                "width": r.get("width"),
                "height": r.get("height"),
                "byte_size": r.get("byte_size"),
            })
        peek["renders"] = compact_renders

    return [TextContent(type="text", text=json.dumps(peek, default=str, indent=2))] + images


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


async def _handle_tool(
    graph: MsGraphInstance,
    name: str,
    args: dict[str, Any],
    *,
    show_room_booking_names: bool = False,
    attachment_roots: list[Path] | None = None,
    max_attachment_bytes: int = DEFAULT_MAX_ATTACHMENT_BYTES,
) -> list[TextContent]:
    """Route a tool call to the appropriate handler."""
    if attachment_roots is None:
        attachment_roots = []

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

    if name == "o365_search_mail":
        kql = _build_mail_kql(args)
        if not kql:
            return [TextContent(
                type="text",
                text="Refused: no search criteria given. Pass at least one of "
                     "from/to/subject/body/since/until/has_attachments/query.",
            )]
        mail = graph.get_mail()
        result = await mail.email_index_async(query=kql, limit=args.get("limit", 25))
        return _json_result(result)

    # ── Mail: drafts (DRAFTS tier) ────────────────────────────────────
    if name == "o365_create_mail_draft":
        try:
            body_text = _resolve_body_text(
                args, attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
            decoded = _decode_attachments(
                args.get("attachments"),
                attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
        except (AttachmentPathError, ValueError) as exc:
            return [TextContent(type="text", text=f"Refused: {exc}")]

        mail = graph.get_mail()
        # Create the draft WITHOUT attachments — we'll attach in parallel.
        result = await mail.create_draft_async(
            to_recipients=args["to"],
            subject=args["subject"],
            body=body_text,
            is_html=args.get("is_html", False),
            cc_recipients=args.get("cc"),
            bcc_recipients=args.get("bcc"),
            attachments=None,
        )
        if result is None:
            return [TextContent(type="text", text="Failed to create draft.")]
        if decoded:
            ok, total = await _upload_attachments_parallel_async(
                graph, result["id"], decoded,
            )
            result = {**result, "attachments_uploaded": ok, "attachments_total": total}
        return _json_result(result)

    if name == "o365_update_mail_draft":
        # Bullet-proof check via Graph $select: only patch actual drafts.
        is_draft = await _is_draft(graph, args["email_id"])
        if is_draft is None:
            return [TextContent(type="text", text="Message not found.")]
        if not is_draft:
            return [TextContent(
                type="text",
                text="Refused: target message is not a draft. o365_update_mail_draft only modifies drafts.",
            )]
        try:
            body_text = _resolve_body_text(
                args, attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
            decoded = _decode_attachments(
                args.get("attachments"),
                attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
        except (AttachmentPathError, ValueError) as exc:
            return [TextContent(type="text", text=f"Refused: {exc}")]

        mail = graph.get_mail()
        # PATCH the draft's scalar fields; handle attachments ourselves.
        result = await mail.update_draft_async(
            message_id=args["email_id"],
            to_recipients=args["to"],
            subject=args["subject"],
            body=body_text,
            is_html=args.get("is_html", False),
            cc_recipients=args.get("cc"),
            bcc_recipients=args.get("bcc"),
            attachments=None,
        )
        if result is None:
            return [TextContent(type="text", text="Failed to update draft.")]
        # Attachments semantics: omitted = untouched; [] = clear; [items] = replace.
        if args.get("attachments") is not None:
            await _clear_attachments_parallel_async(graph, args["email_id"])
            if decoded:
                ok, total = await _upload_attachments_parallel_async(
                    graph, args["email_id"], decoded,
                )
                result = {**result, "attachments_uploaded": ok, "attachments_total": total}
            else:
                result = {**result, "attachments_cleared": True}
        return _json_result(result)

    # ── Mail: send / mutate (ALL tier) ────────────────────────────────
    if name == "o365_send_mail":
        try:
            body_text = _resolve_body_text(
                args, attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
            decoded = _decode_attachments(
                args.get("attachments"),
                attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
        except (AttachmentPathError, ValueError) as exc:
            return [TextContent(type="text", text=f"Refused: {exc}")]

        mail = graph.get_mail()
        # Fast path: no attachments -> one shot via /sendMail.
        if not decoded:
            ok = await mail.send_message_async(
                to_recipients=args["to"],
                subject=args["subject"],
                body=body_text,
                is_html=args.get("is_html", False),
                save_to_sent_items=args.get("save_to_sent_items", True),
                cc_recipients=args.get("cc"),
                bcc_recipients=args.get("bcc"),
            )
            return _json_result({"sent": ok})

        # With attachments: draft → upload in parallel → send.
        draft = await mail.create_draft_async(
            to_recipients=args["to"],
            subject=args["subject"],
            body=body_text,
            is_html=args.get("is_html", False),
            cc_recipients=args.get("cc"),
            bcc_recipients=args.get("bcc"),
            attachments=None,
        )
        if draft is None:
            return [TextContent(type="text", text="Failed to stage draft for send.")]
        uploaded, total = await _upload_attachments_parallel_async(
            graph, draft["id"], decoded,
        )
        if uploaded != total:
            return [TextContent(
                type="text",
                text=f"Aborted: only {uploaded}/{total} attachments uploaded; draft {draft['id']} left in Drafts.",
            )]
        ok = await mail.send_draft_async(draft["id"])
        return _json_result({"sent": ok, "attachments_uploaded": uploaded})

    if name == "o365_send_mail_draft":
        mail = graph.get_mail()
        ok = await mail.send_draft_async(args["email_id"])
        return _json_result({"sent": ok})

    if name == "o365_delete_mail":
        mail = graph.get_mail()
        ok = await mail.delete_message_async(args["email_id"])
        return _json_result({"deleted": ok})

    if name == "o365_move_mail":
        mail = graph.get_mail()
        result = await mail.move_message_async(args["email_id"], args["destination"])
        if result is None:
            return [TextContent(type="text", text="Failed to move message.")]
        return _json_result(result)

    if name == "o365_flag_mail_read":
        mail = graph.get_mail()
        url = f"{graph.msg_endpoint}me/messages/{args['email_id']}"
        ok = await mail.flag_read_async(url, args["is_read"])
        return _json_result({"updated": ok})

    if name == "o365_set_mail_categories":
        mail = graph.get_mail()
        url = f"{graph.msg_endpoint}me/messages/{args['email_id']}"
        ok = await mail.set_mail_categories_async(url, args["categories"])
        return _json_result({"updated": ok})

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
        # Make date-only bounds span the whole day. MS Graph's endDateTime is
        # exclusive, so a bare date for end must be pushed to 23:59:59 or the
        # range collapses (a caller passing start=end=YYYY-MM-DD gets zero hits).
        if "T" not in args["start_date"]:
            start = start.replace(hour=0, minute=0, second=0, microsecond=0)
        if "T" not in args["end_date"]:
            end = end.replace(hour=23, minute=59, second=59, microsecond=0)
        result = await cal.get_events_async(
            start_date=start,
            end_date=end,
            limit=args.get("limit", 25),
        )
        return _json_result(result)

    if name == "o365_search_events":
        from urllib.parse import quote
        filters: list[str] = []

        def _esc(s: str) -> str:
            return s.replace("'", "''")

        if v := args.get("subject"):
            filters.append(f"contains(subject, '{_esc(v)}')")
        if v := args.get("organizer"):
            if "@" in v:
                filters.append(f"organizer/emailAddress/address eq '{_esc(v)}'")
            else:
                filters.append(f"contains(organizer/emailAddress/name, '{_esc(v)}')")
        if v := args.get("attendee"):
            if "@" in v:
                filters.append(f"attendees/any(a: a/emailAddress/address eq '{_esc(v)}')")
            else:
                filters.append(f"attendees/any(a: contains(a/emailAddress/name, '{_esc(v)}'))")
        if v := args.get("since"):
            filters.append(f"start/dateTime ge '{v}T00:00:00'")
        if v := args.get("until"):
            filters.append(f"start/dateTime le '{v}T23:59:59'")

        if not filters:
            return [TextContent(
                type="text",
                text="Refused: no search criteria given. Pass at least one of "
                     "subject/attendee/organizer/since/until.",
            )]

        filter_str = " and ".join(filters)
        limit = min(args.get("limit", 25), 250)
        fields = "id,subject,bodyPreview,start,end,location,organizer,attendees,isAllDay,webLink"
        url = (
            f"{graph.msg_endpoint}me/events"
            f"?$filter={quote(filter_str, safe='')}"
            f"&$select={fields}&$top={limit}&$orderby=start/dateTime desc"
        )
        token = await graph.get_access_token_async()
        resp = await graph.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            status = resp.status_code if resp else "no response"
            detail = resp.text[:300] if resp else ""
            return [TextContent(
                type="text",
                text=f"Event search failed: status={status}. {detail}",
            )]
        events = resp.json().get("value", [])
        return _json_result({"filter": filter_str, "count": len(events), "events": events})

    if name == "o365_create_event":
        from datetime import datetime
        cal = graph.get_calendar()
        result = await cal.create_event_async(
            subject=args["subject"],
            start_time=datetime.fromisoformat(args["start"]),
            end_time=datetime.fromisoformat(args["end"]),
            body=args.get("body"),
            is_html=args.get("is_html", False),
            location=args.get("location"),
            attendees=args.get("attendees"),
            is_all_day=args.get("is_all_day", False),
            calendar_id=args.get("calendar_id"),
        )
        if result is None:
            return [TextContent(type="text", text="Failed to create event.")]
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

    if name == "o365_search_messages":
        kql = _build_messages_kql(args)
        limit = min(args.get("limit", 25), 500)
        body = {
            "requests": [{
                "entityTypes": ["chatMessage"],
                "query": {"queryString": kql},
                "from": 0,
                "size": limit,
            }]
        }
        token = await graph.get_access_token_async()
        resp = await graph.run_async(
            url=f"{graph.msg_endpoint}search/query",
            method="POST", json=body, token=token,
        )
        if resp is None or resp.status_code != 200:
            return [TextContent(
                type="text",
                text=f"Message search failed: status={resp.status_code if resp else 'no response'}. "
                     "Missing Chat.Read or ChannelMessage.Read.All scope is a common cause.",
            )]
        hits = []
        for container in resp.json().get("value", []):
            for container_hit in container.get("hitsContainers", []):
                for hit in container_hit.get("hits", []):
                    r = hit.get("resource", {})
                    body_content = r.get("body", {}).get("content", "") or ""
                    hits.append({
                        "id": r.get("id"),
                        "created": r.get("createdDateTime"),
                        "from": r.get("from", {}).get("user", {}).get("displayName")
                                or r.get("from", {}).get("emailAddress", {}).get("name"),
                        "chat_id": r.get("chatId"),
                        "channel_identity": r.get("channelIdentity"),
                        "preview": body_content[:300],
                        "summary": hit.get("summary"),
                        "web_url": r.get("webUrl"),
                    })
        return _json_result({"kql": kql, "count": len(hits), "hits": hits})

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

    if name == "o365_peek_drive_file":
        files = graph.get_files()
        drive_id = args.get("drive_id", "") or None
        item_meta = await _fetch_drive_item_meta(graph, args["item_id"], drive_id)
        content = await files.get_file_content_async(args["item_id"], drive_id=drive_id)
        if content is None:
            return [TextContent(type="text", text="File not found or is a folder.")]
        name_hint = (item_meta or {}).get("name")
        mime_hint = ((item_meta or {}).get("file") or {}).get("mimeType")
        return _peek_result(
            content, args, name=name_hint, content_type=mime_hint,
            size_from_meta=(item_meta or {}).get("size"),
        )

    if name == "o365_peek_mail_attachment":
        token = await graph.get_access_token_async()
        url = (
            f"{graph.msg_endpoint}me/messages/{args['email_id']}"
            f"/attachments/{args['attachment_id']}"
        )
        resp = await graph.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            status = resp.status_code if resp else "no response"
            return [TextContent(type="text", text=f"Attachment fetch failed: status={status}")]
        att = resp.json()
        cb = att.get("contentBytes")
        if not cb:
            return [TextContent(
                type="text",
                text=f"Attachment has no inline content (type={att.get('@odata.type')}). "
                     f"May be an item attachment or reference attachment.",
            )]
        content = base64.b64decode(cb)
        return _peek_result(
            content, args,
            name=att.get("name"),
            content_type=att.get("contentType"),
            size_from_meta=att.get("size"),
        )

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


def create_server(
    keyfile: str,
    *,
    permission_level: PermissionLevel | None = None,
    attachment_roots: list[Path] | None = None,
    max_attachment_bytes: int | None = None,
) -> tuple[Server, str]:
    """Create the MCP server and return (server, keyfile) for deferred graph init.

    ``permission_level`` gates which tools are advertised AND which can be
    invoked. If omitted, it is resolved from ``OFFICE_CONNECT_PERMISSION_LEVEL``
    or falls back to ``DEFAULT_LEVEL`` (``DRAFTS``).

    ``attachment_roots`` is the fail-closed allowlist for path-based attachments
    and body-path reads. Empty list (the default) disables path mode entirely —
    use this for shared/server deployments.
    """
    level = permission_level if permission_level is not None else resolve_level()
    if attachment_roots is None:
        attachment_roots = _parse_attachment_roots(None)
    if max_attachment_bytes is None:
        max_attachment_bytes = _parse_max_attachment_bytes()
    allowed = filter_tools(level)
    logger.info(
        "office-connect MCP server: permission level=%s (%d/%d tools exposed)",
        level.value, len(allowed), len(TOOLS),
    )
    if attachment_roots:
        logger.warning(
            "attachment-path mode: ENABLED for roots=%s (max=%d bytes)",
            [str(r) for r in attachment_roots], max_attachment_bytes,
        )
    else:
        logger.info(
            "attachment-path mode: DISABLED — path-based attachments and "
            "body_path will be rejected. Use --attachment-root to enable.",
        )

    mcp = Server(
        "office-365-mcp",
        instructions=(
            f"Office 365 MCP Server (permission level: {level.value}).\n"
            "Provides tools for mail, calendar, teams, chats, files, directory, "
            "and profile data via Microsoft Graph API. Write operations are "
            "gated by permission level — the server advertises and accepts "
            "only tools permitted at the configured level."
        ),
    )

    _state: dict[str, Any] = {"graph": None, "keyfile": keyfile}

    async def _get_graph() -> MsGraphInstance:
        if _state["graph"] is None:
            _state["graph"] = await _create_graph(_state["keyfile"])
        return _state["graph"]

    @mcp.list_tools()
    async def list_tools() -> list[Tool]:
        return filter_tools(level)

    @mcp.call_tool()
    async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
        try:
            _require_allowed(name, level)
        except PermissionDenied as exc:
            return [TextContent(type="text", text=f"Permission denied: {exc}")]
        try:
            graph = await _get_graph()
            return await _handle_tool(
                graph, name, arguments,
                attachment_roots=attachment_roots,
                max_attachment_bytes=max_attachment_bytes,
            )
        except Exception:
            logger.exception("Tool %s failed", name)
            return [TextContent(type="text", text=f"Error: tool '{name}' failed. Check server logs for details.")]

    return mcp, keyfile


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


async def main(
    keyfile: str,
    level: PermissionLevel,
    attachment_roots: list[Path],
    max_attachment_bytes: int,
) -> None:
    """Run the MCP server over stdio."""
    mode = (
        f"path-mode ENABLED roots={[str(r) for r in attachment_roots]}"
        if attachment_roots else "path-mode DISABLED"
    )
    print(
        f"Office 365 MCP Server starting (permission: {level.value}; {mode})",
        file=sys.stderr,
    )
    mcp, _ = create_server(
        keyfile,
        permission_level=level,
        attachment_roots=attachment_roots,
        max_attachment_bytes=max_attachment_bytes,
    )
    async with stdio_server() as (read_stream, write_stream):
        await mcp.run(read_stream, write_stream, mcp.create_initialization_options())


def cli() -> None:
    """CLI entry point for the MCP server."""
    parser = argparse.ArgumentParser(description="Office 365 MCP Server")
    parser.add_argument("--keyfile", required=True, help="Path to JSON token file")
    parser.add_argument(
        "--permission-level",
        choices=[l.value for l in PermissionLevel],
        default=None,
        help=(
            "Permission tier: read_only, drafts (default), or all. "
            f"May also be set via the {PERMISSION_ENV_VAR} environment variable."
        ),
    )
    parser.add_argument(
        "--attachment-root",
        action="append",
        default=None,
        metavar="DIR",
        help=(
            "Absolute directory from which attachments / body_path files may "
            "be read. Repeat for multiple roots. OMIT for shared/server "
            f"deployments (fail-closed default). Env: {ATTACHMENT_ROOTS_ENV} "
            "(os.pathsep-separated)."
        ),
    )
    parsed = parser.parse_args()
    try:
        level = resolve_level(parsed.permission_level)
        roots = _parse_attachment_roots(parsed.attachment_root)
        max_bytes = _parse_max_attachment_bytes()
    except ValueError as exc:
        parser.error(str(exc))
    asyncio.run(main(parsed.keyfile, level, roots, max_bytes))


if __name__ == "__main__":
    cli()

# office-connect

[![Python 3.14+](https://img.shields.io/badge/python-3.14%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/Alyxion/office-connect/blob/main/LICENSE)
[![PyPI version](https://img.shields.io/pypi/v/office-connect.svg)](https://pypi.org/project/office-connect/)
[![Code style: ruff](https://img.shields.io/badge/code%20style-ruff-purple.svg)](https://github.com/astral-sh/ruff)

**Microsoft 365 superpowers for Python and AI agents.**

office-connect is a Python library and stdio MCP server that gives both human developers and AI agents structured access to Microsoft 365 — mail, calendar, Teams, chats, files, directory, profile — through the Microsoft Graph API, behind a simple token-based authentication flow.

## Features

**Platform**

- **Stdio MCP server** -- 38 tools, plug into Claude Desktop, Cursor, or any MCP client
- **Device-code sign-in** -- `office-connect login` runs the OAuth flow from the terminal; no admin token export needed
- **Self-healing token refresh** -- proactive (15 min before expiry) + reactive (on a 401 from Graph), persisted back to the keyfile; clients never need to restart for a token refresh
- **Permission tiers** -- `read_only` / `drafts` / `all`, enforced both at `list_tools` and at `call_tool`; a global policy file acts as a host-wide ceiling
- **Path-mode attachments** -- explicit allow-list of filesystem roots from which the MCP may read attachments and message bodies

**Microsoft 365 surfaces**

- **Mail** -- List, read with attachments, search, create/update drafts, send, move, delete, flag read, manage categories
- **Calendar** -- List calendars, query and search events, create events, free/busy schedules, room availability
- **Teams** -- Joined teams, channels, channel messages, members
- **Chat** -- Recent 1:1, group, and meeting chats with message history
- **Files / OneDrive** -- Browse drives, list folders, download content, search; on-the-fly previews for PDF / XLSX / DOCX
- **SharePoint** -- Search sites and list document libraries
- **Directory** -- Users, managers, profile photos
- **Profile** -- Authenticated user's profile

**Development**

- **Mock transport** -- Full synthetic data layer for testing without a real O365 account; mock is fail-closed on production environments

## Quick Start

### 1. Install

```bash
poetry add office-connect
```

### 2. Sign in (device-code flow)

```bash
# First time — supply your Azure AD app credentials once
office-connect login --client-id <APP_ID> --tenant-id <TENANT_ID> [--client-secret <SECRET>]

# Every subsequent re-auth — zero arguments
office-connect login
```

`office-connect login` prints a `https://microsoft.com/devicelogin` URL plus a short code, polls Microsoft until you finish signing in, and then **verifies the new token end-to-end** by calling `/me` and `/me/mailFolders/inbox` — you'll see your name, title, and inbox count right after sign-in. Two files are written, both `0600`:

- `~/.config/office-connect/token.json` — access + refresh tokens (the keyfile)
- `~/.config/office-connect/config.json` — Azure AD app credentials, so re-auths need no arguments

> Your Azure AD app registration must have **"Allow public client flows"** enabled in its authentication manifest for device-code flow to work.

Narrow scopes with repeatable `--scope` (default: all eight groups — `profile`, `directory`, `mail`, `calendar`, `chat`, `teams`, `drive`, `tasks`).

> Not sure which command you need? Run `office-connect` with no arguments — it prints a banner listing the `login` and `import-token` subcommands plus the server invocation.

### 3. Wire into your MCP client

For Claude Desktop, add to `~/Library/Application Support/Claude/claude_desktop_config.json`:

```json
"office-connect": {
  "command": "/usr/local/bin/python",
  "args": [
    "-m", "office_con.mcp_server",
    "--keyfile", "/Users/<you>/.config/office-connect/token.json",
    "--permission-level", "drafts",
    "--attachment-root", "/Users/<you>/Downloads"
  ]
}
```

Restart the client once to pick up the new MCP entry. From then on, tokens stay fresh on their own — the MCP refreshes proactively within 15 min of expiry, reactively on a 401, and persists every refresh back to the keyfile.

### Standalone CLI (rare)

```bash
office-connect --keyfile path/to/token.json
```

This launches the stdio MCP server directly and **requires** `--keyfile`. If you just want to (re-)authenticate, you want `office-connect login` instead — running bare `office-connect` prints the subcommand banner to point you there rather than erroring on the missing flag.

## Authentication & permissions

### Tokens stay fresh automatically

- **Proactive refresh** when the cached access token has under 15 min of life left
- **Reactive refresh-and-retry** if Graph rejects a request with 401 even though the JWT `exp` is still in the future (e.g. server-side revoke)
- **Persisted** back to the keyfile on every successful refresh, so a process restart never loads stale credentials
- **Hot-reloaded** — the MCP `mtime`-watches the keyfile, so updates from `import-token` or another `login` are picked up on the next tool call

Manual top-ups only matter if the refresh token itself has been invalidated (typically after ~90 days of inactivity, or by tenant policy).

### When the session is truly dead, the tools say so

If the refresh token *is* invalidated, the server no longer pretends everything is fine. Instead of letting a 401 turn into an empty inbox or a `null` profile, the failing tool returns a clear, actionable message:

> ⚠️ Office 365 authentication is not working, so this request could not be completed. … Reconnect by running: `office-connect login`

The MCP server also advertises this in its `instructions`, so an assistant like Claude Desktop knows to surface the re-auth step rather than reporting "0 mails" as if the mailbox were empty. To check the connection explicitly, call the read-only **`o365_check_connection`** tool — it returns `{connected: true, email, display_name}` when healthy, or the same re-auth guidance when not. (Ask Claude "are you still connected to my mail?" and it will run this tool.)

### Refreshing tokens externally

If a host application exports tokens through an admin endpoint, drop the exported JSON at the canonical keyfile location:

```bash
office-connect import-token ~/Downloads/token_export.json
# override destination with --dest /custom/path/token.json
```

Both `login` and `import-token` write atomically with `0600` permissions.

### Permission tiers

| tier | what it allows |
|---|---|
| `read_only` | list / get / search / peek — no mutation of Microsoft 365 state |
| `drafts` *(default)* | `read_only` + create and update *draft* emails. No sending. |
| `all` | `drafts` + send mail, move/delete mail, flag read, set categories, create calendar events |

Three places can set the tier — **the most restrictive of the ones that are set wins**:

1. **MCP launcher CLI flag:** `--permission-level read_only|drafts|all`
2. **Environment variable:** `OFFICE_CONNECT_PERMISSION_LEVEL=read_only`
3. **Global policy file:** a JSON object at `~/.config/office-connect/policy.json` (override via `--policy-file PATH` or `$OFFICE_CONNECT_POLICY`):
   ```json
   { "permission_level": "drafts" }
   ```

The policy file acts as a host-wide ceiling: regardless of how any individual MCP launcher is configured, the resolver clamps the effective tier down to the most restrictive value found. A launcher can always tighten further on top. Tools above the effective tier are removed from `list_tools` *and* refused by `call_tool` (defense in depth; unknown tool names are fail-closed too).

### Keyfile format

`office-connect login` and `import-token` write this format; you can also create it by hand if you have tokens from another source:

```json
{
  "app": "MyApp",
  "email": "user@example.com",
  "access_token": "eyJ...",
  "refresh_token": "1.AUs...",
  "client_id": "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx",
  "client_secret": "...",
  "tenant_id": "common"
}
```

`client_secret` is only needed for confidential-client refresh flows; device-code tokens (from `office-connect login`) refresh without one and the field can be omitted or empty.

## Mock Transport

A full mock layer for development and testing — no real O365 account needed.

```python
from office_con.msgraph.ms_graph_handler import MsGraphInstance
from office_con.testing.fixtures import default_mock_profile

profile = default_mock_profile()
graph = MsGraphInstance(endpoint="https://graph.microsoft.com/v1.0/")
graph.enable_mock(profile)

# Now use graph exactly like the real thing
mail = graph.get_mail()
inbox = await mail.email_index_async(limit=10)
```

The mock provides:
- 18+ inbox messages (rich HTML with signatures, newsletters, notifications)
- 9 mail folders with subfolders, fake downloadable attachments
- 127 calendar events across 3 months (OOF, Teams calls, tentative, free blocks)
- 25 directory users with org hierarchy, departments, and profile photos
- Teams, chats, categories, OneDrive stubs
- Synthetic JWT tokens

Face photos can be loaded from JPEG files via `set_faces_dir()` or the `FACES_DIR` env var. Falls back to generated SVG initials.

Safety: mock is automatically blocked on Azure App Service and production URLs.

## Project Structure

```
office-connect/
├── office_con/
│   ├── auth/                       # Azure AD OAuth, scopes, background refresh
│   ├── db/                         # Company directory builder and storage
│   ├── msgraph/                    # MS Graph API handlers
│   │   ├── ms_graph_handler.py     #   Central class: MsGraphInstance
│   │   ├── mail_handler.py         #   List, get, search, draft, send, move, delete
│   │   ├── mail_filter.py          #   KQL builder for mail/event search
│   │   ├── calendar_handler.py     #   Events, schedules, timezones
│   │   ├── places_handler.py       #   Meeting rooms, room availability
│   │   ├── directory_handler.py    #   Users, managers, photos
│   │   ├── teams_handler.py        #   Teams, channels, messages
│   │   ├── chat_handler.py         #   1:1, group, meeting chats
│   │   ├── files_handler.py        #   OneDrive, SharePoint
│   │   └── profile_handler.py      #   /me profile
│   ├── testing/                    # Mock transport, fixtures, synthetic tokens
│   ├── web/                        # FastAPI helpers (image cache router)
│   ├── utils/                      # Misc utilities (excel parser, …)
│   ├── peek.py                     # PDF / xlsx / docx preview generation
│   ├── mcp_permissions.py          # Permission tiers + policy file resolution
│   └── mcp_server.py               # MCP server entry point + CLI subcommands
├── tests/
├── scripts/                        # Standalone scripts (room availability, …)
├── docs/
├── pyproject.toml
└── LICENSE
```

## Development

```bash
poetry install
poetry run pytest          # full suite — mocks always run; integration tests auto-discover a token
poetry run ruff check      # lint
```

The mock-based unit tests need no credentials. Integration tests resolve a token from, in order:

1. `tests/msgraph_test_token.json` (legacy, gitignored)
2. `~/.config/office-connect/token.json` (written by `office-connect login`)
3. `token_file` path inside `tests/test_config.json` (gitignored)

Once you've run `office-connect login` once, the integration suite picks the keyfile up automatically and the bulk of the suite runs without any extra config. A handful of tests stay skipped until you supply tenant-specific expectations (room names, team names, presence users) in `tests/test_config.json`.

## MCP Tools (38)

Every tool is gated by the permission tier in the third column. At `read_only` only the `read_only` rows are advertised; `drafts` adds the `drafts` rows; `all` exposes everything.

### Profile & directory

| Tool | Description | Tier |
|------|-------------|------|
| `o365_check_connection` | Verify the session is authenticated (health check) | read_only |
| `o365_get_profile` | Current user's profile | read_only |
| `o365_list_users` | Organization directory | read_only |
| `o365_get_user_manager` | A user's manager | read_only |

### Mail

Mail results carry recipient lists (`to_recipients`/`cc_recipients`), `conversation_id`, and `internet_message_id` on **both** list/search and get — so an agent can thread a conversation and filter senders/recipients in-memory without N extra fetches. List/search return header metadata only (no body); fetch bodies with `o365_get_mail`/`o365_get_mails` (default `body_format="text"`, truncated at 50 000 chars with `body_truncated`).

| Tool | Description | Tier |
|------|-------------|------|
| `o365_list_mail` | List emails in a folder (header metadata + recipients); `folder` / `exclude_folders` | read_only |
| `o365_get_mail` | Single email; `body_format=text\|html\|none`, `max_body_chars`; `event_id` for meeting requests | read_only |
| `o365_get_mails` | Batch-fetch many emails in one `$batch` round trip (e.g. a whole thread) | read_only |
| `o365_search_mail` | KQL search via Graph `$search`; `folder` / `exclude_folders` | read_only |
| `o365_unread_counts` | Per-folder unread/total counts without paging | read_only |
| `o365_get_mail_categories` | Outlook categories | read_only |
| `o365_create_mail_draft` | Create a draft (NOT sent) | drafts |
| `o365_update_mail_draft` | Update an existing draft | drafts |
| `o365_send_mail` | Send an email immediately | all |
| `o365_send_mail_draft` | Send an existing draft | all |
| `o365_reply_to_mail` | Reply / reply-all and send (keeps threading) | all |
| `o365_forward_mail` | Forward to new recipients and send | all |
| `o365_delete_mail` | Soft-delete (Deleted Items) | all |
| `o365_move_mail` | Move to another folder | all |
| `o365_flag_mail_read` | Mark read / unread | all |
| `o365_set_mail_categories` | Set categories on a message | all |

### Calendar & rooms

| Tool | Description | Tier |
|------|-------------|------|
| `o365_list_calendars` | User's calendars | read_only |
| `o365_get_events` | Events in a date range | read_only |
| `o365_search_events` | Search events via `$filter` | read_only |
| `o365_get_schedule` | Free/busy availability | read_only |
| `o365_list_rooms` | Meeting rooms with capacity / building / floor | read_only |
| `o365_get_room_availability` | Today's room availability | read_only |
| `o365_create_event` | Create event (sends invites if attendees) | all |
| `o365_update_event` | Update an existing event (PATCH; notifies attendees) | all |
| `o365_send_event_invite` | Create a meeting + send invites end-to-end | all |

### Teams & chat

| Tool | Description | Tier |
|------|-------------|------|
| `o365_list_teams` | Joined teams | read_only |
| `o365_list_channels` | Channels in a team | read_only |
| `o365_get_channel_messages` | Recent channel messages | read_only |
| `o365_get_team_members` | Team members | read_only |
| `o365_list_chats` | Recent chats (1:1, group, meeting) | read_only |
| `o365_get_chat_messages` | Recent chat messages | read_only |
| `o365_get_chat_members` | Chat members | read_only |
| `o365_search_messages` | Search Teams + chat messages in one call | read_only |

### Files / OneDrive / SharePoint

| Tool | Description | Tier |
|------|-------------|------|
| `o365_get_my_drive` | Default OneDrive info | read_only |
| `o365_list_drive_items` | Files and folders | read_only |
| `o365_get_file_content` | Download content (UTF-8 or base64) | read_only |
| `o365_peek_drive_file` | Preview a PDF / xlsx / docx without full download | read_only |
| `o365_peek_mail_attachment` | Preview an email attachment without full download | read_only |
| `o365_search_files` | Search OneDrive by name or content | read_only |
| `o365_search_sites` | Search SharePoint sites | read_only |
| `o365_get_site_drives` | Site document libraries | read_only |

## License

[MIT](https://github.com/Alyxion/office-connect/blob/main/LICENSE) -- Copyright (c) 2026 Michael Ikemann

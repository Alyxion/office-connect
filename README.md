# office-connect

[![Python 3.14+](https://img.shields.io/badge/python-3.14%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](https://github.com/Alyxion/office-connect/blob/main/LICENSE)
[![PyPI version](https://img.shields.io/pypi/v/office-connect.svg)](https://pypi.org/project/office-connect/)
[![Code style: ruff](https://img.shields.io/badge/code%20style-ruff-purple.svg)](https://github.com/astral-sh/ruff)

**Python library and MCP server for Microsoft 365 via MS Graph.**

office-connect provides structured access to Microsoft 365 data through the MS Graph API. It lets AI assistants and automation tools query mail, calendar, teams, chats, files, directory, and profile information using a simple token-based authentication flow.

## Features

- **MCP Server** -- Stdio-based MCP server for seamless integration with AI assistants
- **Mail** -- List messages, read bodies and attachments, send, draft, reply, manage categories
- **Calendar** -- List calendars, query events, create events, check free/busy schedules
- **Teams** -- List joined teams, channels, channel messages, and team members
- **Chat** -- List recent 1:1, group, and meeting chats with message history
- **Files / OneDrive** -- Browse drives, list folders, download file content, search by name
- **SharePoint** -- Search sites and list document libraries
- **Directory** -- List organization users, resolve managers, fetch profile photos
- **Profile** -- Retrieve the authenticated user's profile details
- **Mock Transport** -- Full synthetic data layer for testing without a real O365 account

## Quick Start

### Installation

```bash
poetry add office-connect
```

### MCP Server

```bash
office-connect --keyfile path/to/token.json
```

### Token File Format

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

Tokens are automatically refreshed on startup and persisted back to the file (`0600` permissions).

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
│   ├── auth/                # Azure AD OAuth, scopes, background refresh
│   ├── db/                  # Company directory builder and storage
│   ├── msgraph/             # MS Graph API handlers
│   │   ├── ms_graph_handler.py   # Central class: MsGraphInstance
│   │   ├── mail_handler.py       # Send, draft, reply, list, categories
│   │   ├── calendar_handler.py   # Events, schedules, timezones
│   │   ├── directory_handler.py  # Users, managers, photos
│   │   ├── teams_handler.py      # Teams, channels, messages
│   │   ├── chat_handler.py       # 1:1, group, meeting chats
│   │   ├── files_handler.py      # OneDrive, SharePoint
│   │   └── profile_handler.py    # /me profile
│   ├── testing/             # Mock transport, fixtures, synthetic tokens
│   ├── utils/               # Excel parser, health check
│   └── mcp_server.py        # MCP server entry point + CLI
├── tests/
├── docs/
├── pyproject.toml
└── LICENSE
```

## Development

```bash
poetry install
poetry run pytest          # tests (skips integration tests without token file)
poetry run ruff check      # lint
```

## MCP Tools

| Tool | Description |
|------|-------------|
| `o365_get_profile` | Current user's profile |
| `o365_list_mail` | List recent inbox emails |
| `o365_get_mail` | Single email with full body and attachments |
| `o365_get_mail_categories` | Outlook mail categories |
| `o365_list_calendars` | User's calendars |
| `o365_get_events` | Calendar events in a date range |
| `o365_get_schedule` | Free/busy availability |
| `o365_list_teams` | Joined Microsoft Teams |
| `o365_list_channels` | Channels in a team |
| `o365_get_channel_messages` | Channel messages |
| `o365_get_team_members` | Team members |
| `o365_list_chats` | Recent chats |
| `o365_get_chat_messages` | Chat messages |
| `o365_get_chat_members` | Chat members |
| `o365_get_my_drive` | Default OneDrive info |
| `o365_list_drive_items` | Files and folders |
| `o365_get_file_content` | Download file content |
| `o365_search_files` | Search OneDrive |
| `o365_search_sites` | Search SharePoint sites |
| `o365_get_site_drives` | Site document libraries |
| `o365_list_users` | Organization directory |
| `o365_get_user_manager` | User's manager |

## License

[MIT](https://github.com/Alyxion/office-connect/blob/main/LICENSE) -- Copyright (c) 2026 Michael Ikemann

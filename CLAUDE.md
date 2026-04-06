# CLAUDE.md

This file provides guidance to Claude Code when working with this repository.

## Project Overview

office-connect is a Python library and MCP server providing access to Microsoft 365 via MS Graph. It is part of the llming ecosystem.

- **Package name:** `office-connect` (PyPI), import as `office_con`
- **Repo:** https://github.com/Alyxion/office-connect
- **License:** MIT
- **Python:** >=3.14

This is a **library only** — no web apps, no frontend, no UI code. Applications that consume this library live in other repos (e.g., `llming-docs/apps/mail/`).

## Architecture

### Core: MsGraphInstance

Central class in `office_con/msgraph/ms_graph_handler.py`. Manages OAuth tokens and HTTP calls to MS Graph. All handler classes consume this instance:

```
MsGraphInstance (inherits WebUserInstance)
  ├── ProfileHandler     — /me
  ├── OfficeMailHandler  — /me/mailFolders, /me/messages, /me/sendMail
  ├── CalendarHandler    — /me/calendars, calendarView, getSchedule
  ├── DirectoryHandler   — /users, manager, photos
  ├── TeamsHandler       — /me/joinedTeams, channels, messages
  ├── ChatHandler        — /me/chats, messages, members
  └── FilesHandler       — /me/drives, items, search, SharePoint sites
```

### MCP Server (`office_con/mcp_server.py`)

Stdio-based, 22 tools. Entry: `office-connect --keyfile path/to/token.json`. Defers graph creation to first tool call. Auto-refreshes tokens and persists back to keyfile.

### Mock System (`office_con/testing/`)

Operates at the HTTP transport layer — `MockGraphTransport` intercepts `run_async()` in `WebUserInstance`. Handlers work unchanged.

- `mock_data.py` — `MockUserProfile`, `generate_avatar_svg()`, `load_face_photo()`, `set_faces_dir()`
- `mock_transport.py` — URL routing to synthetic responses for all endpoints
- `mock_tokens.py` — Synthetic JWT tokens (HS256, "mock-secret")
- `fixtures.py` — `default_mock_profile()` and factory functions

Face photos: `load_face_photo(gender, index)` searches for JPEGs in configurable directories. Set via `set_faces_dir(path)` or `FACES_DIR` env var. Falls back to SVG initials if no photos found. Photo assets are NOT bundled in this repo — consumers provide their own.

Safety: `__init__.py` blocks mock activation on production (detects Azure App Service env vars).

## Key Conventions

### OAuth Scopes

Use constants from `OfficeUserInstance`, never hardcode strings:

```python
from office_con.auth.office_user_instance import OfficeUserInstance
scopes = OfficeUserInstance.PROFILE_SCOPE + OfficeUserInstance.MAIL_SCOPE
```

Available: `PROFILE_SCOPE`, `MAIL_SCOPE`, `CALENDAR_SCOPE`, `CHAT_SCOPE`, `ONE_DRIVE_SCOPE`, `DIRECTORY_SCOPE`.

### Token File Format

```json
{
  "app": "office-connect",
  "access_token": "eyJ...",
  "refresh_token": "...",
  "client_id": "...",
  "client_secret": "...",
  "tenant_id": "...",
  "email": "user@example.com"
}
```

Use `export_keyfile()` from `mcp_server.py` to write with 0600 permissions.

### MS Graph Timestamps

Graph returns `%Y-%m-%dT%H:%M:%SZ` (UTC, Z-suffix). The mail handler's `parse_mail()` expects this exact format. Mock data must use `dt.strftime("%Y-%m-%dT%H:%M:%SZ")`, not `.isoformat()`.

### CID Image Embedding

Emails may contain `cid:` references to inline images. Pattern:
1. Fetch with `get_mail_async(email_id=id)` — attachments included
2. Parse HTML with BeautifulSoup
3. Match `cid:xxx` to attachment `content_id`
4. Replace with `data:{mime};base64,{bytes}`

## Development

```bash
poetry install
poetry run pytest tests/ -v    # integration tests skip without token file
poetry run ruff check
```

## Sensitive Files (gitignored)

- `tests/msgraph_test_token.json` — OAuth tokens for integration tests
- `.env` — if present

## Relationship to Other Projects

- **llming** (`/Users/michael/projects/llming`) — orchestrator for publishing
- **llming-docs** (`/Users/michael/projects/llming-docs`) — contains the mail app at `apps/mail/` that consumes this library
- **SalesBot** (`/Users/michael/projects/SalesBot`) — source of truth for `office_con/` at `dependencies/office-mcp/`. Push script: `scripts/sync/push_office_connect.sh`
- Do NOT modify files in `SalesBot/dependencies/nice-office/`

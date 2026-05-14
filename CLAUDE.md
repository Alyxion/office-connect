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

Stdio-based. Entry: `office-connect --keyfile path/to/token.json`. Defers graph creation to first tool call. Auto-refreshes tokens and persists back to keyfile.

**Hot keyfile reload.** Before each tool call, the server compares `mtime(keyfile)` against the value captured when the cached graph was built; if it changed, the graph is rebuilt from the new file contents. Clients (Claude Desktop, etc.) do not need to restart after a token refresh.

**Self-healing token refresh.** During a session the MCP refreshes the access token automatically: `MsGraphInstance.get_access_token_async` runs a proactive refresh when the cached token has under 15 minutes of life left, and `run_async` does a reactive refresh-and-retry on a 401 from Graph. Every successful in-process refresh is written back to the keyfile (`_create_graph` wraps `refresh_token_async` with a persister), so a process restart never loads stale credentials.

**Fresh sign-in via the CLI.** When you have neither an access nor a refresh token (truly cold start), use the device-code login:

```bash
# First time — pass the Azure AD app credentials once
office-connect login --client-id <APP_ID> --tenant-id <TENANT_ID> [--client-secret <SECRET>]

# Subsequent re-auths — credentials are persisted in the keyfile
office-connect login
```

The flow prints a microsoft.com/devicelogin URL and short code, blocks polling, then writes the keyfile (default `~/.config/office-connect/token.json`, 0600). All eight scope groups are requested by default; narrow with repeatable `--scope` (`profile`, `directory`, `mail`, `calendar`, `chat`, `teams`, `drive`, `tasks`). The Azure AD app must have **"Allow public client flows"** enabled in its manifest.

**Updating the keyfile from an externally-exported token.** Some host applications offer an admin "Export Token" endpoint. To install the exported JSON at the canonical location:

```bash
office-connect import-token ~/Downloads/token_export.json
# or with a custom destination:
office-connect import-token ~/Downloads/token_export.json --dest /etc/office-connect/token.json
```

Both `login` and `import-token` write atomically with 0600 perms. Any running MCP server pointed at the destination picks up the new token on its next tool invocation — no client restart needed (mtime-watched).

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

## Publishing to PyPI

The package is published as [`office-connect`](https://pypi.org/project/office-connect/) on PyPI. The repo's owner has a PyPI API token stored in `~/.pypirc` (used by Poetry transparently).

To cut a release:

1. **Bump version in two places** (must stay in sync):
   - `pyproject.toml` → `version = "X.Y.Z"`
   - `office_con/__init__.py` → `__version__ = "X.Y.Z"`
2. **Commit and push to git first** so the GitHub `main` matches what PyPI will serve:
   ```bash
   git add pyproject.toml office_con/__init__.py
   git commit -m "Bump version to X.Y.Z"
   git push origin main
   ```
3. **Build artifacts and publish**:
   ```bash
   rm -rf dist/
   poetry publish --build
   ```
   `--build` produces both wheel + sdist into `dist/` and uploads them. The `readme`, `classifiers`, `keywords`, and `repository` fields from `pyproject.toml` populate the PyPI listing — keep those current.
4. **Verify**: `curl -sS https://pypi.org/pypi/office-connect/json | jq .info.version` should report the new version within ~30s of upload (the JSON cache lags the simple index slightly).

PyPI versions are immutable — once `X.Y.Z` is published it cannot be reused; bump again. If a parent project vendors this repo, mirror the bump there too via its sync script.

## Testing

### Mock Tests (no credentials needed)

```bash
poetry run pytest tests/test_new_handlers.py -v
```

Runs against synthetic mock data — covers all handlers (Presence, Tasks, People, Places, MailboxSettings, OnlineMeetings, etc.).

### Integration Tests (real MS Graph)

To run tests against a real Microsoft 365 account:

1. Export a token from a host application (e.g. via an admin token-export endpoint)
2. Create `tests/test_config.json`:

```json
{
    "token_file": "~/Downloads/token_export.json",
    "expected_rooms": ["Chicago", "Paris"],
    "expected_teams": ["My Team Name"],
    "expected_presence_users": ["colleague@example.com"]
}
```

3. Run:

```bash
poetry run pytest tests/test_integration_handlers.py -v -s
```

All fields in the config are optional. Tests for unconfigured features are skipped. The config file is gitignored.

## Sensitive Files (gitignored)

- `tests/msgraph_test_token.json` — OAuth tokens for integration tests
- `tests/test_config.json` — integration test configuration (user-specific)
- `.env` — if present

## Relationship to Other Projects

- **llming** — orchestrator for publishing
- **llming-docs** — contains the mail app at `apps/mail/` that consumes this library

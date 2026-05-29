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
  ├── OfficeMailHandler  — /me/mailFolders, /me/messages, /me/sendMail, reply/forward, $batch
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

**Loud auth failures (no silent empties).** When a 401 *cannot* be recovered — no refresh capability, refresh token missing/expired/revoked, or Graph still 401s after a successful refresh — `run_async` raises `AuthExpiredError` instead of returning the 401 for handlers to quietly turn into an empty `UserProfile()`/`[]`. `call_tool` catches it (and pre-flights the no-token-at-all cold start) and returns a clear `⚠️ Office 365 authentication…` message naming the fix: run `office-connect login`. The server's MCP `instructions` tell the assistant to treat that message as a dead session rather than real (empty) data. This is what stops the client from reporting "0 mails / profile null" as if the mailbox were genuinely empty. A dedicated read-only `o365_check_connection` tool probes `/me` and returns `{connected: true, email, display_name}` when healthy, or the same re-auth message when not — call it first when the user asks "are you still connected?" or when results look suspiciously empty. Tests: `tests/test_auth_errors.py`.

**Fresh sign-in via the CLI.** When you have neither an access nor a refresh token (truly cold start), use the device-code login:

```bash
# First time — pass the Azure AD app credentials once
office-connect login --client-id <APP_ID> --tenant-id <TENANT_ID> [--client-secret <SECRET>]

# Subsequent re-auths — zero arguments, credentials come from the saved config
office-connect login
```

The flow prints a microsoft.com/devicelogin URL and short code, blocks polling, writes the keyfile (default `~/.config/office-connect/token.json`, 0600), **and** saves the Azure AD app credentials to a separate app-config file (`~/.config/office-connect/config.json`, 0600). From then on, plain `office-connect login` resolves credentials in this order:

1. CLI args (`--client-id` / `--tenant-id` / `--client-secret`)
2. Values already in the keyfile (back-compat with existing setups)
3. Env vars (`O365_CLIENT_ID` / `O365_TENANT_ID` / `O365_CLIENT_SECRET`)
4. App-config file (overridable with `--app-config PATH` or `$OFFICE_CONNECT_APP_CONFIG`)

You can also pre-create the app-config file by hand (it's just `{"client_id": "...", "tenant_id": "...", "client_secret": "..."}`). All eight scope groups are requested by default; narrow with repeatable `--scope` (`profile`, `directory`, `mail`, `calendar`, `chat`, `teams`, `drive`, `tasks`). The Azure AD app must have **"Allow public client flows"** enabled in its manifest.

**Updating the keyfile from an externally-exported token.** Some host applications offer an admin "Export Token" endpoint. To install the exported JSON at the canonical location:

```bash
office-connect import-token ~/Downloads/token_export.json
# or with a custom destination:
office-connect import-token ~/Downloads/token_export.json --dest /etc/office-connect/token.json
```

Both `login` and `import-token` write atomically with 0600 perms. Any running MCP server pointed at the destination picks up the new token on its next tool invocation — no client restart needed (mtime-watched).

### Permission tiers and the global policy file

Three trust tiers: `read_only` < `drafts` (default) < `all`. Each tool in `mcp_server.py` is tagged with its required tier in `TOOL_PERMISSIONS`; `list_tools` filters by tier *and* `call_tool` re-checks via `_require_allowed` before dispatching (defense in depth, fail-closed for any unknown tool name).

The effective tier is computed from three sources — **the most restrictive level among the sources that are set wins**. If nothing is set anywhere, the default is `drafts`.

| source | how to set |
|---|---|
| MCP CLI flag | `--permission-level read_only\|drafts\|all` in the launcher args |
| environment variable | `OFFICE_CONNECT_PERMISSION_LEVEL=read_only` |
| **global policy file** *(new)* | a JSON object with `permission_level` (or `max_permission_level`) at `~/.config/office-connect/policy.json` (default), overridable via `--policy-file PATH` or `$OFFICE_CONNECT_POLICY` |

Example `~/.config/office-connect/policy.json`:

```json
{ "permission_level": "drafts" }
```

That single file then enforces a host-wide ceiling regardless of how an individual MCP launcher is configured — if any launcher asks for `all`, the resolver clamps it back down. Conversely, a launcher can always *tighten* further by setting its own `--permission-level read_only`.

A malformed or partial policy file logs a `[PERM]` warning and is ignored (no silent loosening). All tool-gating logic is in `office_con/mcp_permissions.py`; tests live in `tests/test_mcp_permissions.py`.

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

### Mail data model (`OfficeMail`)

`parse_mail()` populates recipient/header metadata so agents can thread and filter without extra round-trips: `to_recipients` / `cc_recipients` / `bcc_recipients` / `reply_to` / `sender_*` (lists of `MailAddress`), `conversation_id` (pull a whole thread via search), and `internet_message_id` (RFC-822, for deduping forward chains). These ride on **list/search results too** — `email_index_async` selects them (`_INDEX_FIELDS`); list/search never fetch the full body (use `get_mail`/`get_mails`).

- **URL fields:** `graph_url` (API) and `outlook_url` (human-openable) are the clear names; `email_url` is kept as a deprecated alias of `graph_url`, `web_link` as the original of `outlook_url`.
- **Body hygiene:** `get_mail_async(body_format="text"|"html"|"none", max_body_chars=…)`. `body_text` is always provided (HTML stripped via BeautifulSoup); over-limit bodies are cut with `body_truncated=True`. The MCP tools default to `body_format="text"`, `max_body_chars=50000` to avoid context blowups; the library default is no limit (back-compat).
- **Batch:** `get_mails_async(ids, …)` pulls many messages in one Graph `$batch` (chunked at 20), preserving input order.
- **Folder scoping:** `folder=` (well-known name via `resolve_well_known_folder` — inbox/sent/deleteditems/junk/archive/… — or id) scopes list+search; `exclude_folders=` drops by parentFolderId (client-side, may return fewer than `limit`).
- **Legacy EX-DN:** `_is_legacy_dn` detects X500/Exchange DNs; `resolve_legacy_addresses_async` best-effort resolves them to SMTP via a cached directory lookup by display name (silently keeps the DN if directory scope is absent). Recipient-level DNs are isolated into `MailAddress.legacy_dn`.
- **Meeting requests:** `event_id` links an `eventMessageRequest` to its calendar event (fetched via `$expand=event` on get) — hand it to `o365_get_events`.

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

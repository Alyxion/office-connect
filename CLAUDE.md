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

**Throttling (HTTP 429) + request timeouts.** Graph throttles `/search/query` and other endpoints aggressively under intensive/bursty use (e.g. an agent firing many searches for one person's chats). `run_async` now retries a 429 up to `_MAX_THROTTLE_RETRIES` (3) times, honoring the `Retry-After` header, with bounded exponential-backoff fallback — each wait clamped to `_THROTTLE_WAIT_CAP_S` (10s) so a hostile/huge `Retry-After` degrades to "give up" rather than a long hang. If still throttled, `o365_search_messages` returns a throttling-specific message (not the misleading "missing scope" hint). Every Graph HTTP call is also bounded by `aiohttp.ClientTimeout(total=60, sock_connect=10)` in `WebUserInstance.run_async`; without it aiohttp's 5-minute default let one stuck request hang the whole stdio server. `call_tool` turns a `TimeoutError` into a clean `⏱️` message. Tests: `tests/test_logging_and_throttle.py`.

**Rotated, secret-redacting logs (`office_con/logging_setup.py`) — local MCP only.** `configure_logging()` (called from `cli()`) attaches a `RotatingFileHandler` (1 MB × 5 files) to the `office_con` package logger. Default `~/.config/office-connect/logs/office-connect.log` (dir 0700, file 0600); override with `--log-file PATH` / `$OFFICE_CONNECT_LOG_FILE` (`none` disables) and `--log-level` / `$OFFICE_CONNECT_LOG_LEVEL` (default INFO). A `_RedactingFilter` masks bearer tokens, JWTs, and email addresses as defense-in-depth; URLs are redacted to scheme+host+path before logging (`_redact_url`) so `$search`/`$filter` query content (names, emails, ids) never reaches disk. Idempotent — safe to call more than once.

**File logging is fail-safe disabled on servers.** `is_server_environment()` gates `configure_logging()`: in any hosted/server deployment it writes **no** file at all (privacy — never persist user data on a shared/multi-tenant host), regardless of `--log-file`. This is **not** overridable on purpose. Detection reuses the Azure App Service signals the mock-block trusts (`WEBSITE_INSTANCE_ID`, production `WEBSITE_URL`) plus `FUNCTIONS_WORKER_RUNTIME`, `KUBERNETES_SERVICE_HOST`, and an explicit `OFFICE_CONNECT_SERVER=1` marker a deployment can set defensively. In a server case the package logger keeps `propagate=True`, so library records still flow to the host app's own logging (which owns redaction/retention). A library consumer never calls `configure_logging()` anyway — it's only invoked by the local stdio `cli()` — so this is belt-and-suspenders for someone running the stdio server on a hosted box.

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

**Browser sign-in (`login --browser`) for tenants that block device code.** Some tenants block the device-code flow; the code-entry page then fails with `AADSTS900561` and no sign-in log entry is written. `office-connect login --browser` runs the authorization-code + PKCE flow instead (`office_con/auth/browser_login.py`): MSAL builds the authorize URL, the browser opens it, and a short-lived HTTPS listener bound to loopback only (127.0.0.1 / ::1, non-loopback peers get 403) catches the redirect on `https://localhost:8443/auth/callback` (override with `--redirect-host/--redirect-port/--redirect-path`; `--no-open` just prints the URL). That URI must be registered on the Azure AD app under **Mobile and desktop applications** (public client, so no secret is needed to redeem the code). TLS: `--cert-file/--key-file`, else a cached cert in `~/.config/office-connect/tls/`, else `mkcert` (browser-trusted), else a self-signed cert (one-time browser warning). Port 8443 rather than 443 because macOS only lets unprivileged processes bind <1024 on the wildcard address, and the listener must stay on loopback. Tests: `tests/test_browser_login.py`.

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
- **Attachments:** `get_mail_async` lists attachment **metadata only** (`attachment_id`, `name`, `size`; no bytes) via a separate constant-size `/attachments?$select=…` call — inlining bytes via `$expand=attachments` previously timed out on large PDFs. `$select` must use base-type properties only (`contentId`/`contentBytes` are on the derived `fileAttachment` type; selecting them 400s the collection). MCP tools: `o365_peek_mail_attachment` (preview text+PNG, degrades to `unreadable` on a corrupt file) and `o365_get_mail_attachment` (full bytes as base64, or — with `--file-server` — an unguessable loopback URL; see `docs/file_server.md`).
- **Folder scoping:** `folder=` (well-known name via `resolve_well_known_folder` — inbox/sent/deleteditems/junk/archive/… — or id) scopes list+search; `exclude_folders=` drops by parentFolderId (client-side, may return fewer than `limit`).
- **Legacy EX-DN:** `_is_legacy_dn` detects X500/Exchange DNs; `resolve_legacy_addresses_async` best-effort resolves them to SMTP via a cached directory lookup by display name (silently keeps the DN if directory scope is absent). Recipient-level DNs are isolated into `MailAddress.legacy_dn`.
- **Meeting requests:** `event_id` links an `eventMessageRequest` to its calendar event (fetched via `$expand=event` on get) — hand it to `o365_get_events`.

### Public repository — no internal names (explicit rule)

This repository is public. Organisation names, internal domains and hosts, colleague names,
cost centres / internal orders, tenant ids and customer or card numbers must never appear — not
in code, docs, tests, examples, file names or commit messages. Use neutral examples (`Jane Doe`,
`example.com`). Enforced by `scripts/content_guard.py`: the blocked terms are stored only as
salted SHA-256 hashes (`scripts/content_guard.sha256`, add entries with `--hash TERM
[--substring]`, never plaintext), local extras go in `~/.config/office-connect/content-guard.txt`.

- **Enable the hooks once per clone:** `git config core.hooksPath .githooks` (`pre-commit` checks
  staged blobs + file names, `commit-msg` the message).
- `tests/test_content_guard.py` scans the whole working tree, so CI and hook-less clones fail too.
- The guard never prints the matched term, only `path:line` — its output may land in CI logs.

### One interface for all apps (explicit rule)

office-connect is THE Microsoft 365 interface for every app in the ecosystem. Apps never build
Graph/Outlook/login URLs, never call `run_async` with their own URLs and never parse raw Graph
JSON — they call handler methods and get Pydantic models. When an app needs something Graph can
do but the library can't yet, it is added HERE (in the style of the existing handlers, with a
test), not worked around in the app. Human-openable links come from Graph itself
(`OfficeMail.outlook_url` / `web_link`, `DriveItem.web_url`), never from string templates.

### Stateless — no local file storage (explicit rule)

`MsGraphInstance`, `WebUserInstance` and every handler in `office_con/msgraph/` run on
**stateless servers**: they must never read or write local files — no temp/staging
dirs, no caches on disk, no token files. Tokens live in `cache_dict` (and the
optional Redis session store); content goes in and out as `bytes` / Pydantic models
(`OfficeMailAttachment.content_bytes`, `upload_file_by_path_async(content=…)`). Writing
files is the CALLER's business (an app saving attachments, the CLI's keyfile via
`export_keyfile()` in `mcp_server.py`). Keyfile loading therefore stays out of the
handlers too.

### Complete listings, drafts and drive writes

- **Complete listings raise, lookups degrade.** Interactive lookups (`email_index_async`,
  `search_items_async`, …) return an empty model on a non-200. Operations that promise
  the FULL result — `OfficeMailHandler.list_messages_async` (paged `@odata.nextLink`,
  filters `folder` / `received_since` / `has_attachments` / `subject_prefix`) and the
  conditional drive writes — raise `GraphRequestError` (`office_con.msgraph.errors`)
  instead: a rejected listing once looked exactly like "no mails with attachments".
  Graph requires the `$orderby` property FIRST in `$filter` (else 400 `InefficientFilter`).
- **Drafts:** `save_draft_async(...) -> OfficeDraft` (id, web_link, failed_attachments),
  `patch_draft_async(id, body=/subject=/to_recipients=/cc_recipients=)` changes only the
  given fields; `add_attachment_async` uses an upload session above 3 MB (inline POST 413s),
  `list_attachments_async` / `clear_attachments_async` for re-attaching edited files.
  `create_draft_async` / `update_draft_async` keep their dict return for the MCP tools.
- **Calendar:** `get_events_async` follows `@odata.nextLink` up to `limit`; events carry
  `is_cancelled` and `categories`.
- **Drive by path** (relative to the user's OneDrive root): `get_item_by_path_async`,
  `get_file_content_by_path_async -> (bytes, DriveItem)` (privacy checks applied),
  `upload_file_by_path_async(path, content, if_match=e_tag | create_only=True)` —
  optimistic concurrency, `DriveItemConflictError` on 409/412, upload session above 4 MB —
  `create_folder_path_async`, `move_item_async`. There is deliberately **no delete
  helper**: callers move obsolete items into a trash folder instead.
- **Upload sessions:** `MsGraphInstance.upload_session_async(upload_url, content)` sends the
  chunks (multiples of 320 KiB) WITHOUT the bearer token — the URL is pre-authenticated
  and Graph rejects an Authorization header on it.

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

### End-to-End MCP Tests (no credentials needed)

`office_con.testing.mcp_client` is a small stdlib-only MCP client that spawns the
real stdio server (`python -m office_con.mcp_server`) and drives it over JSON-RPC
against mock data — no Office 365 credentials, no network. This exercises the
full tool layer (registration, permission filtering, `call_tool` dispatch), which
handler unit tests don't. Mock mode is requested by a `{"mock": true}` keyfile and
triple-gated (`LLMING_MOCK_USERS=1` + `is_mock_enabled()` production block).

```bash
poetry run pytest tests/test_mcp_client_e2e.py -v
# Poke tools by hand:
python -m office_con.testing.mcp_client --mock list
python -m office_con.testing.mcp_client --mock call o365_get_mail '{"email_id":"…"}'
```

Full guide: `docs/mcp_testing.md`.

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

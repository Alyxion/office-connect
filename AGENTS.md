# AGENTS.md

Operational rules for AI agents (Claude Code and others) working in this repo.
These are **mandatory** and override any default behavior.

## 🔒 Live token — READ-ONLY, NO EXCEPTIONS

A real Microsoft 365 OAuth token may exist on this machine at
`~/.config/office-connect/token.json` (and possibly `tests/msgraph_test_token.json`
or a path in `tests/test_config.json`). It belongs to a real human mailbox.

When using that token — directly, via `MsGraphInstance`, via the MCP server, or
through any integration test — you may perform **only read-only** Microsoft Graph
calls. You must **NEVER** trigger an operation that sends, mutates, or deletes
anything in the real account. Specifically forbidden against the live token:

- **Sending mail of any kind**: `send_message_async`, `send_draft_async`,
  `reply_async` / `replyAll`, `forward_async`, and the MCP tools
  `o365_send_mail`, `o365_send_mail_draft`, `o365_reply_to_mail`,
  `o365_forward_mail`.
- **Creating / changing calendar items**: `create_event_async`,
  `update_event_async`, and the tools `o365_create_event`, `o365_update_event`,
  `o365_send_event_invite` (Graph emails invitations on create/update — this
  sends real mail to real attendees).
- **Any other write**: creating/updating/deleting drafts, deleting or moving
  messages, flagging read/unread, setting categories — `o365_create_mail_draft`,
  `o365_update_mail_draft`, `o365_delete_mail`, `o365_move_mail`,
  `o365_flag_mail_read`, `o365_set_mail_categories`.

Allowed against the live token: read-only Microsoft Graph calls only. Most are
GETs: profile, list/search/get mail, folder listing, unread counts, calendar
reads, teams/chat/files/directory reads, `o365_check_connection`. Some Graph
read APIs are POST-shaped but still allowed when they do not mutate server
state: `$batch` containing GET subrequests only, `/search/query` for file or
message search, and `/me/calendar/getSchedule` for schedule/free-busy.

**Send/write code paths must be exercised against the mock transport only**
(`office_con/testing/`), never the real token. If you believe a real-world write
test is genuinely necessary, STOP and ask the human first — do not proceed on
your own judgment.

### Enforcement in test code

- Integration tests that hit real Graph live in `tests/test_*integration*.py`
  must contain only read-only Graph calls. POST is allowed only for Graph APIs
  that are explicitly read-only (`$batch` of GETs only, `/search/query`,
  `/me/calendar/getSchedule`).
- Any new real-API test must be read-only. Do not add a test that sends mail or
  creates/updates events against the resolved token file.
- Mock-based tests (`MockGraphTransport`) are the home for all send/write
  coverage — the mock never reaches Microsoft.

## Other repo rules

- This repo stays **anonymized**: no real company / HQ / domain / team names
  anywhere (code, tests, fixtures, docs).
- Ask before `git commit`, `git push`, PyPI upload, or any `gh` action — each
  needs its own approval; one approval does not carry to the next action.
- See `CLAUDE.md` for architecture, conventions, and the publishing checklist.

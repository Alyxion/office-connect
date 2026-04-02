# Mock User System

The mock user system enables headless testing (Playwright, integration tests) without real Microsoft 365 credentials. It operates at the **HTTP transport layer**: `run_async()` intercepts MS Graph API URLs and returns synthetic JSON; `_async_token_request()` returns synthetic tokens. All real handlers (`CalendarHandler`, `MailHandler`, `DirectoryHandler`, `ProfileHandler`) work unchanged.

## Architecture

### Interception points

All MS Graph communication flows through two methods on `WebUserInstance` / `MsGraphInstance`:

1. **`run_async(url, method, json, token)`** — every handler calls this. The mock intercepts here and routes by URL pattern to return synthetic MS Graph JSON responses.

2. **`_async_token_request(**form_data)`** — called by `acquire_token_async()` (code→token) and `refresh_token_async()`. The mock returns a synthetic JWT access token + refresh token.

3. **`build_auth_url(auth_url)`** — returns a local redirect URL with a fake auth code instead of the Microsoft login URL.

### Key classes

| Class | File | Purpose |
|-------|------|---------|
| `MockGraphTransport` | `office_mcp/testing/mock_transport.py` | URL-routing mock for `run_async()` |
| `MockUserProfile` | `office_mcp/testing/mock_data.py` | Pydantic model holding synthetic data |
| `make_mock_token_response()` | `office_mcp/testing/mock_tokens.py` | Synthetic JWT token generation |
| `is_mock_enabled()` | `office_mcp/testing/__init__.py` | Safety guard with production blockers |

### Enabling on an instance

```python
from office_mcp.msgraph.ms_graph_handler import MsGraphInstance
from office_mcp.testing.mock_data import MockUserProfile
from office_mcp.testing.fixtures import default_calendar_events, default_mail_inbox, default_company_directory

profile = MockUserProfile(
    email="test@example.com",
    user_id="mock-user-001",
    given_name="Test", surname="User", full_name="Test User",
    calendar_events=default_calendar_events(),
    mail_messages=default_mail_inbox(),
    directory_users=default_company_directory(),
)

graph = MsGraphInstance(app="my-app", session_id="test-session")
graph.enable_mock(profile)

# All subsequent calls go through MockGraphTransport:
response = await graph.run_async(url="https://graph.microsoft.com/v1.0/me")
# → returns synthetic profile JSON, no HTTP request made
```

### Safety guards

The mock system requires `LLMING_MOCK_USERS=1` **and** blocks automatically on Azure App Service:

- `WEBSITE_INSTANCE_ID` set → blocked (Azure App Service)
- `WEBSITE_URL` is set to a non-localhost URL (e.g. `azurewebsites.net`) → blocked
- Mock tokens contain `iss: "mock-issuer"` → unusable against real Azure AD
- `_mock_transport is None` check is O(1) — zero overhead when disabled

## Mock data fixtures

`office_mcp/testing/fixtures.py` provides default data generators returning MS Graph JSON:

| Function | Returns |
|----------|---------|
| `default_calendar_events()` | ~8 events (standup, 1:1, team sync, all-day) |
| `default_mail_inbox()` | ~5 messages with realistic subjects/senders |
| `default_company_directory()` | ~15 users with manager_id links (org tree) |

All data uses `@example.com` emails and generic names. No customer-specific references.

## Extending the mock transport

To add mock responses for new MS Graph endpoints, add a handler method to `MockGraphTransport`:

```python
# In mock_transport.py

async def handle_request(self, url, method, json_body):
    path = _extract_path(url)
    ...
    if path.startswith("me/drive/"):
        return self._drive_response(path, method, json_body)
    ...

def _drive_response(self, path, method, json_body):
    return _make_response(200, {"value": [...]})
```

The response must match the MS Graph JSON schema that the real handler expects.

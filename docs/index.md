# office-mcp

Microsoft 365 / MS Graph logic layer with MCP server support.

**office-mcp** provides a fully async Python library for accessing Microsoft 365
services via the Microsoft Graph API. It includes OAuth token management,
user authentication, mail, calendar, Teams, chats, files, directory, and
profile operations, plus a standalone MCP (Model Context Protocol) server
for AI tool integration.


## Architecture

The package is organized into four layers:

1. **Session layer** — `WebUserInstance` and `DBUserInstance` manage OAuth
   tokens, Redis caching, and user sessions.
2. **Handler layer** — Typed async handlers for each Graph domain (mail,
   calendar, teams, chats, files, directory, profile).
3. **MCP layer** — In-process MCP servers (`Office365MCPServer`) and a
   standalone stdio MCP server (`mcp_server.py`).
4. **Utilities** — Database helpers, health checks, company directory, and
   file format parsers.

```text
office_mcp/
├── __init__.py               # Public API exports
├── web_user_instance.py      # OAuth session management
├── db_user_instance.py       # DB-based auth (password/secret + JWT)
├── _db_helpers.py            # MongoDB / Redis connection factories
├── _mcp_base.py              # InProcessMCPServer ABC
├── mcp_server.py             # Standalone stdio MCP server (CLI)
├── auth/
│   ├── azure_auth_utils.py   # FastAPI middleware, redirect URL builder
│   ├── background_service_registry.py  # Login/logout/loop callbacks
│   └── office_user_instance.py         # Unified user instance wrapper
├── msgraph/
│   ├── ms_graph_handler.py   # MsGraphInstance (central handler)
│   ├── profile_handler.py    # User profile
│   ├── mail_handler.py       # Outlook mail (read + write)
│   ├── mail_filter.py        # Email filtering rules
│   ├── calendar_handler.py   # Calendar events (read + write)
│   ├── teams_handler.py      # Teams & channels (read-only)
│   ├── chat_handler.py       # 1:1 and group chats (read-only)
│   ├── files_handler.py      # OneDrive & SharePoint (read-only)
│   ├── directory_handler.py  # Azure AD directory (read-only)
│   └── mcp/
│       ├── base.py           # MsGraphMCPServer base
│       └── office_mcp.py     # Office365MCPServer (contact cards)
├── db/
│   └── company/
│       ├── company_dir.py        # CompanyDir + LiveCompanyDirData
│       └── company_dir_builder.py # Build directory from Graph
└── utils/
    ├── health_check.py           # DB health monitoring
    └── file_formats/
        └── excel_parser.py       # .xls / .xlsx parser
```


## Installation

Add to your `pyproject.toml` dependencies:

```
[tool.poetry.dependencies]
office-mcp = {path = "../office-mcp", develop = true}
```

Or install directly:

```
pip install -e dependencies/office-mcp
```


## Configuration

All configuration is via environment variables. No values are hardcoded.

### Core OAuth

| Variable | Description | Required |
| --- | --- | --- |
| `O365_CLIENT_ID` | Azure AD application (client) ID | Yes |
| `O365_CLIENT_SECRET` | Azure AD client secret value | Yes |
| `O365_TENANT_ID` | Azure AD tenant ID (default: `common`) | No |
| `O365_ENDPOINT` | MS Graph endpoint (default: `https://graph.microsoft.com/v1.0/`) | No |

### Security

| Variable | Description | Required |
| --- | --- | --- |
| `O365_SALT` | Encryption salt for Redis token storage. **Must be set in production.** | Yes |
| `O365_TOKEN_SECRET` | HMAC key for `DBUserToken` signing (falls back to `O365_SALT`) | Yes |

### Caching

| Variable | Description | Required |
| --- | --- | --- |
| `O365_REDIS_URL` | Redis connection URL (supports comma-separated cluster nodes) | No |
| `MONGODB_CONNECTION` | MongoDB connection string | No |

### Auth Redirects

| Variable | Description | Required |
| --- | --- | --- |
| `WEBSITE_REDIRECT_URL` | Explicit OAuth redirect URL (highest priority, most secure) | No |
| `ALLOWED_REDIRECT_HOSTS` | Comma-separated allowed hosts for redirect URL construction | No |
| `WEBSITE_HOSTNAME` | Fallback for `ALLOWED_REDIRECT_HOSTS` | No |

### Other

| Variable | Description | Required |
| --- | --- | --- |
| `WEBSITE_URL` | Base URL for user photo URLs (default: `https://localhost:8000`) | No |
| `O365_MSAL_REGION` | MSAL region hint (for sovereign clouds) | No |
| `TRUSTED_PROXY_HOSTS` | Comma-separated proxy IPs for `ProxyHeadersMiddleware` | No |


## Quick Start

### Creating an authenticated Graph instance

```python
from office_mcp.msgraph.ms_graph_handler import MsGraphInstance

graph = MsGraphInstance(
    app="my-app",
    client_id="...",
    client_secret="...",
    tenant_id="...",
    endpoint="https://graph.microsoft.com/v1.0/",
    can_refresh=True,
)
graph.cache_dict["access_token"] = "<token>"
graph.cache_dict["refresh_token"] = "<refresh_token>"
```

### Reading mail

```python
mail = graph.get_mail()
inbox = await mail.email_index_async(limit=10)
for m in inbox.elements:
    print(f"{m.from_name}: {m.subject}")
```

### Reading calendar events

```python
from datetime import datetime, timedelta

cal = graph.get_calendar()
events = await cal.get_events_async(
    start_date=datetime.now(),
    end_date=datetime.now() + timedelta(days=7),
    limit=25,
)
for e in events.events:
    print(f"{e.subject} at {e.start_time}")
```

### Listing Teams and channels

```python
teams = graph.get_teams()
team_list = await teams.get_joined_teams_async()
for t in team_list.teams:
    channels = await teams.get_channels_async(t.id)
    print(f"{t.display_name}: {channels.total_channels} channels")
```

### Browsing OneDrive files

```python
files = graph.get_files()
drive = await files.get_my_drive_async()
root = await files.get_root_items_async(limit=20)
for item in root.items:
    icon = "folder" if item.is_folder else "file"
    print(f"  [{icon}] {item.name}")
```

### Running the standalone MCP server

```bash
# Create a keyfile with your credentials
python -m office_mcp.mcp_server --keyfile token.json
```

The keyfile must contain:

```json
{
  "app": "my-app",
  "email": "user@example.com",
  "access_token": "eyJ...",
  "refresh_token": "1.AUs...",
  "client_id": "your-client-id-here",
  "client_secret": "your-client-secret-here",
  "tenant_id": "your-tenant-id-here"
}
```

Tokens are automatically refreshed and persisted back to the keyfile.


## Detailed Documentation

- [Session Layer](session_layer.md)
- [Handlers](handlers.md)
- [MCP Servers](mcp_servers.md)
- [Company Directory](company_directory.md)
- [Utilities](utilities.md)
- [Azure AD Permissions](azure_ad_permissions.md)
- [Mock Users (Testing)](mock_users.md)
- [Security](security.md)

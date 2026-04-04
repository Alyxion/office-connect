# MCP Servers

The office-mcp package provides a full MCP (Model Context Protocol) integration
layer that exposes Microsoft 365 operations as AI-consumable tools. There are two
deployment modes: **in-process** servers that run inside a host application with
direct access to runtime state, and a **standalone** stdio server that runs as an
independent process communicating over stdin/stdout.

## Table of Contents

- [Overview](#overview)
- [InProcessMCPServer ABC](#inprocessmcpserver-abc)
- [MsGraphMCPServer](#msgraphmcpserver)
- [Office365MCPServer](#office365mcpserver)
- [Standalone MCP Server](#standalone-mcp-server)
- [Complete Tool Index](#complete-tool-index)
- [Integration Patterns](#integration-patterns)

## Overview

The Model Context Protocol (MCP) is a standard for exposing tools to AI models.
Each tool has a name, a description, and a JSON Schema input specification. The
AI model selects tools based on descriptions and constructs the required
arguments. The MCP server executes the tool and returns a text result.

office-mcp implements MCP at three levels:

1. **InProcessMCPServer** (`_mcp_base.py`) -- Abstract base class defining the
   in-process server contract. No external dependencies.

2. **MsGraphMCPServer** (`msgraph/mcp/base.py`) -- Base class that binds an
   authenticated `MsGraphInstance` to the in-process server contract.

3. **Office365MCPServer** (`msgraph/mcp/office_mcp.py`) -- Concrete server
   exposing calendar, people search, free-slot finding, and org chart tools,
   plus contact card rendering for chat UIs.

4. **Standalone MCP server** (`mcp_server.py`) -- A full stdio MCP server with
   22 tools covering mail, calendar, Teams, chats, files, SharePoint, directory,
   and profile operations. Authenticates via a JSON keyfile.

```text
   _mcp_base.py                   mcp_server.py
   +------------------------+     +---------------------------+
   | InProcessMCPServer     |     | Standalone stdio server   |
   |   (ABC)                |     |   (mcp.server.Server)     |
   +----------+-------------+     |   22 Graph tools          |
              |                   |   keyfile auth            |
              v                   +---------------------------+
   msgraph/mcp/base.py
   +------------------------+
   | MsGraphMCPServer       |
   |   holds MsGraphInstance|
   +----------+-------------+
              |
              v
   msgraph/mcp/office_mcp.py
   +------------------------+
   | Office365MCPServer     |
   |   4 high-level tools   |
   |   contact card UI      |
   |   fuzzy name matching  |
   +------------------------+
```

## InProcessMCPServer ABC

*Module:* `office_mcp._mcp_base`

The `InProcessMCPServer` abstract base class defines the contract for MCP
servers that run in the same process as the host application. This gives them
direct access to runtime state such as authenticated user sessions, database
connections, and cached data -- without any serialization overhead.

### Class: `InProcessMCPServer`

Abstract methods (must be implemented by subclasses):

`async list_tools() -> List[Dict[str, Any]]`
:   Return the list of available tools. Each tool is a dictionary with at
    minimum `name`, `description`, and `inputSchema` keys. May also
    include display metadata such as `displayName`, `displayDescription`,
    and `icon`.

`async call_tool(name: str, arguments: Dict[str, Any]) -> str`
:   Execute a tool by name with the given arguments. Returns the tool result
    as a string (plain text or formatted Markdown).

Optional methods (have default implementations):

`async get_prompt_hints() -> List[str]`
:   Return a list of prompt snippets that should be appended to the system
    prompt when this server is active. Used to inject rendering instructions,
    usage guidelines, or other context that helps the AI model use the tools
    correctly. Default implementation returns an empty list.

`async get_client_renderers() -> List[Dict[str, str]]`
:   Return a list of client-side renderer definitions for custom fenced code
    block languages. Each renderer is a dictionary with keys:

    - `lang` -- the fenced code block language identifier (e.g. `contact`)
    - `css` -- CSS rules for the renderer
    - `js` -- JavaScript code for the renderer

    The host application injects these into the chat UI so that tool results
    containing custom fenced blocks are rendered as interactive widgets.
    Default implementation returns an empty list.

### Tool Registration Pattern

To create a custom in-process MCP server, subclass `InProcessMCPServer` and
implement `list_tools()` and `call_tool()`:

```python
from office_mcp._mcp_base import InProcessMCPServer

class MyCustomMCPServer(InProcessMCPServer):

    async def list_tools(self):
        return [
            {
                "name": "my_tool",
                "description": "Does something useful.",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "query": {
                            "type": "string",
                            "description": "Search query",
                        },
                    },
                    "required": ["query"],
                },
            },
        ]

    async def call_tool(self, name, arguments):
        if name == "my_tool":
            return f"Result for: {arguments['query']}"
        return f"Unknown tool: {name}"
```

The host application registers the server, calls `list_tools()` to advertise
available tools to the AI model, and dispatches `call_tool()` when the model
selects a tool.

## MsGraphMCPServer

*Module:* `office_mcp.msgraph.mcp.base`

The `MsGraphMCPServer` extends `InProcessMCPServer` with an authenticated
`MsGraphInstance`. It serves as the base class for any in-process MCP server
that needs to call Microsoft Graph APIs.

### Class: `MsGraphMCPServer(graph: MsGraphInstance)`

**Constructor parameter:**

`graph`
:   An authenticated `MsGraphInstance` with valid tokens.

**Instance attribute:**

`graph` *(MsGraphInstance)*
:   The MS Graph instance used by all tool implementations. Subclasses access
    this attribute to call Graph API handlers (calendar, directory, mail, etc.).

This class does not define any tools itself -- it is a binding layer. Concrete
subclasses like `Office365MCPServer` implement `list_tools()` and
`call_tool()` to expose specific Graph operations.

```python
from office_mcp.msgraph.mcp.base import MsGraphMCPServer

class MyGraphServer(MsGraphMCPServer):
    async def list_tools(self):
        return [...]

    async def call_tool(self, name, arguments):
        # Access self.graph to call MS Graph APIs
        mail = self.graph.get_mail()
        inbox = await mail.email_index_async(limit=10)
        ...
```

## Office365MCPServer

*Module:* `office_mcp.msgraph.mcp.office_mcp`

The `Office365MCPServer` is the primary in-process MCP server for Office 365
integration. It provides four high-level tools designed for conversational AI
use cases: calendar events, people search (with fuzzy matching), free-slot
finding, and org chart navigation.

### Class: `Office365MCPServer(graph, *, photo_url_prefix, company_dir)`

**Constructor parameters:**

`graph`
:   An authenticated `MsGraphInstance`.

`photo_url_prefix` *(str, default "/api/photo")*
:   URL prefix for user photo endpoints. Photos are served at
    `{photo_url_prefix}/{user_id}`.

`company_dir` *(optional, default None)*
:   Optional pre-loaded company directory (`CompanyDir` instance). When
    provided, people search and org chart tools use cached directory data
    instead of live Graph API calls, improving performance and enabling
    enriched fields (company, building, room, street, city, country, join
    date, birthday).

### Tools

The server exposes four tools:

#### o365_calendar_get_events

Search or list calendar events for the logged-in user.

- **Display Name:** Calendar Events
- **Icon:** `event`

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `query` | string | Search term to filter events by subject (case-insensitive substring match). When set, the date range defaults to 6 months and up to 500 events are fetched then filtered locally. | No |
| `start_date` | string | Start date in `YYYY-MM-DD` format. Defaults to today. | No |
| `end_date` | string | End date in `YYYY-MM-DD` format. Defaults to 14 days from start (or 6 months when `query` is set). | No |
| `limit` | integer | Maximum events to return. Default 50. | No |

Returns a formatted list of events with subject, time, location, attendees,
and Teams meeting links.

#### o365_resolve_contact

Look up people in the company directory by name with fuzzy matching.

- **Display Name:** People Search
- **Icon:** `person_search`

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `query` | string | Name(s) to search for. For multiple people, separate with commas (e.g. `Daniel Sagmeister, Joerg Sauter`). | Yes |

**Fuzzy matching** handles:

- Umlaut variants (`Mueller` / `Muller`, `Sauter` / `Sautter`)
- Accent normalization (diacritics are stripped)
- Partial names (single surname queries work correctly)
- Typos (token-level `SequenceMatcher` scoring with length-difference penalty)

The minimum match score is 0.75. Results are capped at 5 per query term and
include match confidence labels (`exact` or `fuzzy NN%`).

Each result includes: display name, email, job title, department, phone, office
location, photo URL (if available), manager name, meeting frequency over the
last 90 days, and (when using `CompanyDir`) company, building, room, street,
city, country, join date, and birthday.

#### o365_calendar_find_free_slots

Find common free time slots between the logged-in user and one or more attendees.

- **Display Name:** Free Slot Finder
- **Icon:** `schedule`

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `attendee_emails` | array[string] | Email addresses of attendees to check availability for. | Yes |
| `duration_minutes` | integer | Desired meeting length in minutes. Default 60. | No |
| `start_date` | string | Start date in `YYYY-MM-DD`. Defaults to tomorrow. | No |
| `end_date` | string | End date in `YYYY-MM-DD`. Defaults to 5 business days from start. | No |
| `start_hour` | integer | Earliest hour (0--23) for suggested slots. Default 8. | No |
| `end_hour` | integer | Latest hour (0--23) for suggested slots. Default 18. | No |

Uses the Graph `getSchedule` API with 15-minute granularity. The logged-in
user is automatically included in the schedule query. Slots where all
participants are free or tentative are returned, filtered by working hours
and minimum duration.

#### o365_org_chart

Show organizational structure for a person: manager chain upward and direct
reports downward.

- **Display Name:** Org Chart
- **Icon:** `account_tree`

| Parameter | Type | Description | Required |
|-----------|------|-------------|----------|
| `email` | string | Email address of the person to show the org chart for. | Yes |
| `depth_up` | integer | How many levels of managers to show upward. Default 3. | No |
| `depth_down` | integer | How many levels of direct reports to show downward. Default 1. | No |

Returns a structured Markdown document with `contact` fenced code blocks
for each person in the chain, suitable for rendering as interactive contact
cards in the chat UI.

### Contact Card Rendering

The `Office365MCPServer` provides a complete client-side rendering system for
displaying people as interactive contact cards in chat UIs.

#### Prompt Hints

The `get_prompt_hints()` method injects two system prompt directives:

1. **Contact Card Rendering** -- Instructs the AI model to always present people
   found via People Search using `contact` fenced code blocks, never as plain
   text bullet points. The format is:

   ````text
   ```contact
   name: Jane Doe
   email: jane.doe@example.com
   department: Engineering
   title: Senior Engineer
   photo: /api/photo/a1b2c3d4-e5f6-7890-abcd-ef1234567890
   phone: +49 123 456 789
   location: Building A
   manager: John Smith
   company: Acme Corp
   city: Stuttgart
   country: Germany
   ```
   ````

   Supported fields: `name`, `email`, `department`, `title`, `photo`,
   `phone`, `location`, `manager`, `company`, `building`, `room`,
   `street`, `zip`, `city`, `country`, `joined`, `birthday`. Only
   `name` and `email` are required.

2. **Org Chart Tool** -- Instructs the AI model when to use `o365_org_chart`
   and how to visualize results (Mermaid diagrams, text trees, or contact cards).

#### Client Renderers

The `get_client_renderers()` method returns JavaScript and CSS for two
renderers:

1. **Contact card renderer** (`lang: "contact"`) -- Renders `contact` fenced
   code blocks as compact chip-style cards with avatar (initials fallback with
   async photo loading), name, title, department, phone, location, email link,
   and configurable action buttons (Outlook, Teams, org chart). The org chart
   action button shows a hover popup with manager and direct reports fetched
   from a contact API endpoint.

2. **Inline email enhancer** (`type: "inline"`) -- Scans rendered messages for
   email addresses matching configured domains and enhances them with hover
   popups showing the person's contact card. Enhances both `mailto:` links and
   plain-text email addresses. Only enabled when `emailEnhancerDomains` is
   configured.

#### Contact Card Configuration

The `_contact_card_config()` method returns a configuration dictionary that
controls the contact card renderer. Override this in subclasses to customize
behavior without modifying the renderer logic.

| Key | Description | Default |
|-----|-------------|---------|
| `emailHref` | URL template for email click. Use `{email}` placeholder. | Outlook compose deeplink |
| `actions` | List of action button definitions (see below). | Outlook + Teams |
| `extraCss` | Additional CSS rules appended to the base styles. | `""` |
| `emailEnhancerDomains` | List of email domains to enhance inline (empty disables the enhancer). | `[]` |
| `contactApiPrefix` | URL prefix for contact lookup API (used by org chart hover). | `/api/contact` |
| `orgSvg` | Inline SVG icon for the org chart action button. | Tree diagram icon |

Each action in the `actions` list is a dictionary:

| Key | Description | Required |
|-----|-------------|----------|
| `url` | URL template with `{email}` and `{name}` placeholders. | Yes |
| `label` | Tooltip text. | Yes |
| `svg` | Inline SVG icon HTML. | No (use `text` instead) |
| `text` | Text label (if no SVG). | No |
| `css` | Extra inline CSS on the `<a>` tag. | No |
| `domain` | Only show for emails ending with this domain. | No |

## Standalone MCP Server

*Module:* `office_mcp.mcp_server`

The standalone MCP server (`mcp_server.py`) provides a complete stdio-based
MCP server with 22 tools covering all Microsoft 365 domains. It uses the
`mcp` Python SDK (`mcp.server.Server`) and communicates over stdin/stdout,
making it compatible with any MCP client (Claude Desktop, custom integrations,
etc.).

Unlike the in-process servers, the standalone server authenticates via a JSON
keyfile containing OAuth tokens. It does not require a web framework or user
session management.

### CLI Usage

```bash
# Run the standalone MCP server
python -m office_mcp.mcp_server --keyfile /path/to/token.json
```

The server prints `Office 365 MCP Server starting...` to stderr and then
enters the stdio event loop. It creates the MS Graph instance lazily on the
first tool call.

For integration with Claude Desktop, add to your MCP configuration:

```json
{
  "mcpServers": {
    "office365": {
      "command": "python",
      "args": ["-m", "office_mcp.mcp_server", "--keyfile", "/path/to/token.json"]
    }
  }
}
```

### Keyfile Format

The keyfile is a JSON file containing OAuth credentials. It is read at startup
and updated in place when tokens are refreshed.

```json
{
  "app": "office-mcp",
  "session_id": "optional-session-id",
  "email": "user@example.com",
  "access_token": "<YOUR_ACCESS_TOKEN>",
  "refresh_token": "<YOUR_REFRESH_TOKEN>",
  "client_id": "<YOUR_CLIENT_ID>",
  "client_secret": "<YOUR_CLIENT_SECRET>",
  "tenant_id": "<YOUR_TENANT_ID>",
  "endpoint": "https://graph.microsoft.com/v1.0/"
}
```

#### JSON Schema

```json
{
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "type": "object",
  "properties": {
    "app":            {"type": "string", "description": "Application name for logging", "default": "office-mcp"},
    "session_id":     {"type": "string", "description": "Optional session identifier"},
    "email":          {"type": "string", "format": "email", "description": "User email address"},
    "access_token":   {"type": "string", "description": "MS Graph OAuth access token (JWT)"},
    "refresh_token":  {"type": "string", "description": "OAuth refresh token for automatic renewal"},
    "client_id":      {"type": "string", "description": "Azure AD application (client) ID"},
    "client_secret":  {"type": "string", "description": "Azure AD client secret value"},
    "tenant_id":      {"type": "string", "description": "Azure AD tenant ID", "default": "common"},
    "endpoint":       {"type": "string", "format": "uri", "description": "MS Graph API endpoint", "default": "https://graph.microsoft.com/v1.0/"}
  },
  "required": ["access_token"]
}
```

**Security:** The keyfile is written with restrictive permissions (`0600`) when
created or updated by the server. It contains secrets and must be protected
accordingly. Do not commit keyfiles to version control.

#### Token Refresh

On startup, the server attempts to refresh the access token using the refresh
token. If successful, both the new access token and refresh token are persisted
back to the keyfile. If the refresh fails (e.g. expired refresh token, network
error), the server falls back to the original access token from the keyfile.

This means a keyfile with a valid refresh token, `client_id`, `client_secret`,
and `tenant_id` will keep itself alive across restarts without manual token
rotation.

#### export_keyfile()

The `export_keyfile()` function creates a keyfile programmatically.

**Signature:**

```python
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
) -> None
```

**Parameters:**

`path`
:   Filesystem path for the keyfile.

`access_token`
:   MS Graph OAuth access token.

`refresh_token`
:   OAuth refresh token.

`client_id`
:   Azure AD application (client) ID.

`client_secret`
:   Azure AD client secret.

`tenant_id` *(default "common")*
:   Azure AD tenant ID.

`app` *(default "office-mcp")*
:   Application name for logging.

`session_id` *(default None)*
:   Optional session identifier.

`email` *(default None)*
:   Optional user email address.

The file is written with `0600` permissions.

**Example:**

```python
from office_mcp.mcp_server import export_keyfile

export_keyfile(
    "/path/to/token.json",
    access_token="<YOUR_ACCESS_TOKEN>",
    refresh_token="<YOUR_REFRESH_TOKEN>",
    client_id="<YOUR_CLIENT_ID>",
    client_secret="<YOUR_CLIENT_SECRET>",
    tenant_id="<YOUR_TENANT_ID>",
    app="my-app",
    email="user@example.com",
)
```

### Standalone Server Tool Reference

The standalone server exposes 22 tools organized by Microsoft 365 domain. All
tools are read-only.

#### Profile

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_get_profile` | Get the current user's profile (name, email, job title, department, phone, location). | None |

#### Mail

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_list_mail` | List recent emails from the user's inbox. | `limit` (int, default 10), `skip` (int, default 0) |
| `o365_get_mail` | Get a single email by ID, including full body and attachments metadata. | `email_id` (string, **required**) |
| `o365_get_mail_categories` | List the user's Outlook mail categories. | None |

#### Calendar

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_list_calendars` | List the user's calendars. | None |
| `o365_get_events` | Get calendar events within a date range. | `start_date` (string, **required**), `end_date` (string, **required**), `limit` (int, default 25) |
| `o365_get_schedule` | Get free/busy availability for one or more users. | `emails` (array[string], **required**), `start` (string, **required**), `end` (string, **required**) |

#### Teams

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_list_teams` | List Microsoft Teams the user has joined. | None |
| `o365_list_channels` | List channels in a team. | `team_id` (string, **required**) |
| `o365_get_channel_messages` | Get recent messages from a team channel. | `team_id` (string, **required**), `channel_id` (string, **required**), `limit` (int, default 20) |
| `o365_get_team_members` | List members of a team. | `team_id` (string, **required**) |

#### Chats

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_list_chats` | List the user's recent chats (1:1, group, meeting). | `limit` (int, default 25) |
| `o365_get_chat_messages` | Get recent messages from a chat. | `chat_id` (string, **required**), `limit` (int, default 20) |
| `o365_get_chat_members` | List members of a chat. | `chat_id` (string, **required**) |

#### Files / OneDrive

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_get_my_drive` | Get the user's default OneDrive info. | None |
| `o365_list_drive_items` | List files and folders in a drive location. | `folder_id` (string), `drive_id` (string), `limit` (int, default 25) |
| `o365_get_file_content` | Download a file's content as text (UTF-8). For binary files, returns base64. | `item_id` (string, **required**), `drive_id` (string) |
| `o365_search_files` | Search for files by name or content in OneDrive. | `query` (string, **required**), `limit` (int, default 10) |

#### SharePoint

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_search_sites` | Search for SharePoint sites. Use `*` for all sites. | `query` (string, **required**) |
| `o365_get_site_drives` | List document libraries in a SharePoint site. | `site_id` (string, **required**) |

#### Directory

| Tool | Description | Parameters |
|------|-------------|------------|
| `o365_list_users` | List users in the organization directory. | `limit` (int, default 25) |
| `o365_get_user_manager` | Get a user's manager. | `user_id` (string, **required**) |

## Complete Tool Index

This table lists all MCP tools across both the in-process and standalone servers
for quick reference.

| Tool Name | Server | Description |
|-----------|--------|-------------|
| `o365_calendar_get_events` | In-process | Search or list calendar events with optional subject filtering |
| `o365_resolve_contact` | In-process | Fuzzy people search in the company directory |
| `o365_calendar_find_free_slots` | In-process | Find common free meeting slots across multiple attendees |
| `o365_org_chart` | In-process | Show manager chain and direct reports for a person |
| `o365_get_profile` | Standalone | Current user's profile |
| `o365_list_mail` | Standalone | List inbox emails with pagination |
| `o365_get_mail` | Standalone | Get a single email by ID |
| `o365_get_mail_categories` | Standalone | List Outlook mail categories |
| `o365_list_calendars` | Standalone | List user's calendars |
| `o365_get_events` | Standalone | Get calendar events in a date range |
| `o365_get_schedule` | Standalone | Get free/busy availability for users |
| `o365_list_teams` | Standalone | List joined Microsoft Teams |
| `o365_list_channels` | Standalone | List channels in a team |
| `o365_get_channel_messages` | Standalone | Get messages from a team channel |
| `o365_get_team_members` | Standalone | List members of a team |
| `o365_list_chats` | Standalone | List recent chats |
| `o365_get_chat_messages` | Standalone | Get messages from a chat |
| `o365_get_chat_members` | Standalone | List members of a chat |
| `o365_get_my_drive` | Standalone | Get default OneDrive info |
| `o365_list_drive_items` | Standalone | List files and folders in a drive location |
| `o365_get_file_content` | Standalone | Download file content (text or base64) |
| `o365_search_files` | Standalone | Search OneDrive by name or content |
| `o365_search_sites` | Standalone | Search SharePoint sites |
| `o365_get_site_drives` | Standalone | List document libraries in a SharePoint site |
| `o365_list_users` | Standalone | List organization directory users |
| `o365_get_user_manager` | Standalone | Get a user's manager |

## Integration Patterns

### In-Process vs Standalone

The two deployment modes serve different use cases:

**In-process** (`Office365MCPServer`)
:   Best for web applications where users are already authenticated via OAuth.
    The server runs inside the application process and shares the authenticated
    `MsGraphInstance` directly -- no credential files, no separate processes.
    Supports advanced features like contact card rendering, prompt hints, and
    company directory caching.

    ```python
    from office_mcp.msgraph.mcp.office_mcp import Office365MCPServer

    # Inside your web app, after user login
    server = Office365MCPServer(
        graph=user_graph_instance,
        photo_url_prefix="/api/photo",
        company_dir=preloaded_company_dir,  # optional
    )

    # Register with your AI framework
    tools = await server.list_tools()
    hints = await server.get_prompt_hints()
    renderers = await server.get_client_renderers()

    # When the AI model selects a tool
    result = await server.call_tool("o365_resolve_contact", {"query": "Jane"})
    ```

**Standalone** (`mcp_server.py`)
:   Best for desktop integrations, CLI tools, or any scenario where the MCP
    client and server run as separate processes. Authenticates via a keyfile
    and communicates over stdio. Provides the broadest tool coverage (22 tools)
    but does not support UI rendering features.

    ```bash
    # Export credentials
    python -c "
    from office_mcp.mcp_server import export_keyfile
    export_keyfile('token.json',
        access_token='<YOUR_ACCESS_TOKEN>',
        refresh_token='<YOUR_REFRESH_TOKEN>',
        client_id='<YOUR_CLIENT_ID>',
        client_secret='<YOUR_CLIENT_SECRET>',
        tenant_id='<YOUR_TENANT_ID>',
        email='user@example.com',
    )
    "

    # Run the server
    python -m office_mcp.mcp_server --keyfile token.json
    ```

### Extending Office365MCPServer

To customize the contact card behavior (e.g. add an intranet profile link or
restrict the email enhancer to your domain), subclass `Office365MCPServer`
and override `_contact_card_config()`:

```python
from office_mcp.msgraph.mcp.office_mcp import Office365MCPServer

class CustomOffice365MCPServer(Office365MCPServer):

    def _contact_card_config(self):
        config = super()._contact_card_config()
        # Add a custom action button
        config["actions"].append({
            "label": "Intranet Profile",
            "text": "Profile",
            "url": "https://intranet.example.com/people?email={email}",
            "domain": "example.com",
        })
        # Enable inline email enhancement for your domain
        config["emailEnhancerDomains"] = ["example.com"]
        return config
```

#### Adding Tools

To add custom tools alongside the built-in ones:

```python
class ExtendedOffice365MCPServer(Office365MCPServer):

    async def list_tools(self):
        tools = await super().list_tools()
        tools.append({
            "name": "my_custom_tool",
            "description": "A custom tool.",
            "inputSchema": {
                "type": "object",
                "properties": {"arg": {"type": "string"}},
                "required": ["arg"],
            },
        })
        return tools

    async def call_tool(self, name, arguments):
        if name == "my_custom_tool":
            return f"Custom result: {arguments['arg']}"
        return await super().call_tool(name, arguments)
```

### Server Lifecycle

Both server types handle the MS Graph connection lifecycle automatically:

- **In-process:** The `MsGraphInstance` is injected at construction time and
  must already have valid tokens. Token refresh is handled by the instance
  itself (or by the host application's session layer).

- **Standalone:** The `MsGraphInstance` is created lazily on the first tool
  call from the keyfile. Token refresh is attempted at creation time and the
  refreshed tokens are persisted to the keyfile. The Graph instance is then
  reused for all subsequent tool calls within the server's lifetime.

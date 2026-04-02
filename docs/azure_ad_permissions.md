# Azure AD Permissions Reference for office-mcp

**Date:** 2026-03-07
**Permission type:** Delegated (on behalf of the signed-in user)

---

- [Overview](#overview)
- [Azure AD App Registration Setup](#azure-ad-app-registration-setup)
  - [Creating the App Registration](#creating-the-app-registration)
  - [Client Credentials](#client-credentials)
  - [Redirect URIs](#redirect-uris)
- [Token Types and Authentication Flow](#token-types-and-authentication-flow)
  - [Delegated vs. Application Permissions](#delegated-vs-application-permissions)
  - [OAuth 2.0 Authorization Code Flow](#oauth-20-authorization-code-flow)
  - [Scope Aggregation](#scope-aggregation)
- [Required Permissions by Handler](#required-permissions-by-handler)
  - [Profile Handler](#profile-handler)
  - [Mail Handler](#mail-handler)
  - [Calendar Handler](#calendar-handler)
  - [Teams Handler](#teams-handler)
  - [Chat Handler](#chat-handler)
  - [Files Handler](#files-handler)
  - [Directory Handler](#directory-handler)
- [Consolidated Permission Reference](#consolidated-permission-reference)
  - [Minimum Permission Set](#minimum-permission-set)
  - [Full Permission Set](#full-permission-set)
  - [Scope Constants in Code](#scope-constants-in-code)
- [Admin Consent Requirements](#admin-consent-requirements)
  - [Permissions Requiring Admin Consent](#permissions-requiring-admin-consent)
  - [User-Consentable Permissions](#user-consentable-permissions)
  - [Granting Tenant-Wide Admin Consent](#granting-tenant-wide-admin-consent)
- [Feature-Based Permission Selection](#feature-based-permission-selection)
- [MCP Server Mode](#mcp-server-mode)
  - [MCP Server Tool-to-Permission Mapping](#mcp-server-tool-to-permission-mapping)
- [Security Considerations](#security-considerations)

---

## Overview

The `office-mcp` package integrates with Microsoft 365 services through the
Microsoft Graph API. All API access is performed using **delegated permissions**,
meaning the application acts on behalf of the signed-in user and can only access
resources the user themselves has access to.

To use `office-mcp`, you must register an application in Azure Active Directory
(now called Microsoft Entra ID) and configure the appropriate API permissions.
This document provides a comprehensive reference of every permission required,
organized by functional handler, along with guidance on app registration setup,
token management, and admin consent.

## Azure AD App Registration Setup

### Creating the App Registration

1. Sign in to the [Azure Portal](https://portal.azure.com) and navigate to
   **Microsoft Entra ID** (formerly Azure Active Directory) > **App registrations**.
2. Click **New registration**.
3. Fill in the registration form:

   - **Name**: A descriptive name, e.g. `office-mcp` or your organization's app name.
   - **Supported account types**: Choose based on your deployment:

     - *Accounts in this organizational directory only* -- single-tenant (recommended
       for enterprise deployments).
     - *Accounts in any organizational directory* -- multi-tenant.

   - **Redirect URI**: See the section below.

4. Click **Register** to create the application.
5. Note the **Application (client) ID** and **Directory (tenant) ID** from the
   overview page.

### Client Credentials

1. In the app registration, navigate to **Certificates & secrets** >
   **Client secrets**.
2. Click **New client secret**, provide a description and expiry period.
3. Copy the secret value immediately (it is only shown once).

The following environment variables must be set for `office-mcp`:

| Environment Variable | Description |
|---|---|
| `O365_CLIENT_ID` | Application (client) ID from the app registration |
| `O365_CLIENT_SECRET` | Client secret value |
| `O365_TENANT_ID` | Directory (tenant) ID, or `common` for multi-tenant |
| `O365_ENDPOINT` | Microsoft Graph API base URL (default: `https://graph.microsoft.com/v1.0/`) |

### Redirect URIs

The OAuth 2.0 authorization code flow requires a redirect URI. Configure this under
**Authentication** > **Platform configurations** > **Web**.

`office-mcp` determines the redirect URL at runtime using the following priority:

1. The `WEBSITE_REDIRECT_URL` environment variable (most secure, recommended for
   production).
2. Reconstruction from the incoming request's `X-Forwarded-Proto` and
   `Disguised-Host` / `X-Forwarded-Host` headers (common in Azure App Service).
3. Fallback to the request's `base_url`.

For security, the `ALLOWED_REDIRECT_HOSTS` environment variable (or
`WEBSITE_HOSTNAME`) restricts which hosts are accepted when building redirect
URLs from headers, preventing open-redirect attacks.

Typical redirect URI patterns:

- Production: `https://your-app.azurewebsites.net/auth/callback`
- Development: `http://localhost:8081/auth/callback`

All configured redirect URIs in your app registration must match exactly what
the application sends during the OAuth flow.

## Token Types and Authentication Flow

### Delegated vs. Application Permissions

`office-mcp` uses **delegated permissions** exclusively. This means:

- The application always acts on behalf of a signed-in user.
- Access is limited to what the user themselves can access.
- No application-level (daemon/service) permissions are used.
- A user must sign in and consent (or have admin consent granted) before the
  application can access their data.

### OAuth 2.0 Authorization Code Flow

`office-mcp` implements the standard OAuth 2.0 authorization code flow with
refresh tokens:

1. **Authorization request**: The application redirects the user to the Microsoft
   identity platform authorization endpoint
   (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize`) with the
   requested scopes.

2. **Token acquisition**: After the user consents, the authorization code is
   exchanged for an access token and refresh token via the token endpoint
   (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`).

3. **Token refresh**: When the access token approaches expiration, the refresh
   token is used to obtain a new access token without user interaction. The
   `MsGraphInstance` handles this automatically, checking expiry before each
   API call.

4. **MCP server mode**: When running as an MCP server (`mcp_server.py`), tokens
   are loaded from a JSON key file. On startup, the server attempts a token refresh
   and persists refreshed tokens back to the key file.

### Scope Aggregation

When the `MsGraphInstance` is initialized, all required scopes are passed as a
single list. During the authorization request and token exchanges, these scopes are
joined with spaces and sent as the `scope` parameter. The scopes defined in
`OfficeUserInstance` (see below) are combined based on which features the
application needs.

## Required Permissions by Handler

This section details every Microsoft Graph API endpoint used by each handler and
the corresponding delegated permission(s) required. The scope constants are defined
in `office_mcp.auth.office_user_instance.OfficeUserInstance`.

### Profile Handler

**Source**: `office_mcp.msgraph.profile_handler.ProfileHandler`

**Scope constant**: `PROFILE_SCOPE = ["User.Read"]`

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /me` | `User.Read` | Retrieve the signed-in user's profile (display name, email, job title, office location, phone numbers, user ID) |

The profile handler is invoked automatically during token acquisition to populate
the user's identity (email, name, user ID) on the `MsGraphInstance`.

### Mail Handler

**Source**: `office_mcp.msgraph.mail_handler.OfficeMailHandler`

**Scope constant**:

```python
MAIL_SCOPE = [
    "Mail.Read",
    "Mail.Read.Shared",
    "Mail.ReadWrite",
    "Mail.ReadWrite.Shared",
    "Mail.Send",
    "Mail.Send.Shared",
    "User.ReadBasic.All",
]
```

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /me` | `User.Read` | Fetch user profile for sender identification |
| `GET /me/mailFolders/inbox/messages` | `Mail.Read` | List inbox emails for the signed-in user |
| `GET /users/{email}/mailFolders/Inbox/messages` | `Mail.Read.Shared` | List inbox emails from a shared or delegated mailbox |
| `GET /me/messages/{id}` | `Mail.Read` | Read a single email with full body and attachments |
| `PATCH /me/messages/{id}` (isRead) | `Mail.ReadWrite` | Mark an email as read or unread |
| `PATCH /me/messages/{id}` (categories) | `Mail.ReadWrite` | Set categories on an email |
| `PATCH /users/{email}/messages/{id}` | `Mail.ReadWrite.Shared` | Modify emails in shared mailboxes |
| `GET /me/outlook/masterCategories` | `Mail.Read` | List the user's Outlook category definitions |
| `GET /users/{email}/outlook/masterCategories` | `Mail.Read.Shared` | List categories for a shared mailbox |
| `POST /me/outlook/masterCategories` | `Mail.ReadWrite` | Create a new Outlook category |
| `POST /me/sendMail` | `Mail.Send` | Send a new email on behalf of the signed-in user |
| `POST /me/messages` (isDraft=true) | `Mail.ReadWrite` | Create a draft email |
| `PATCH /me/messages/{id}` (draft update) | `Mail.ReadWrite` | Update an existing draft |
| `POST /me/messages/{id}/send` | `Mail.Send` | Send an existing draft |
| `POST /me/messages/{id}/attachments` | `Mail.ReadWrite` | Add attachments to a draft message |
| `GET /me/messages/{id}/attachments` | `Mail.Read` | List attachments on a message |
| `DELETE /me/messages/{id}/attachments/{att_id}` | `Mail.ReadWrite` | Remove an attachment from a draft message |

`User.ReadBasic.All` is included in the mail scope to resolve user display names
and email addresses for shared mailbox scenarios.

### Calendar Handler

**Source**: `office_mcp.msgraph.calendar_handler.CalendarHandler`

**Scope constant**: `CALENDAR_SCOPE = ["Calendars.ReadWrite", "Place.Read.All"]`

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /me/calendars` | `Calendars.Read` * | List the user's calendars |
| `GET /me/calendars/{id}/calendarView` | `Calendars.Read` * | Get events within a date range (calendar view) |
| `POST /me/calendars/{id}/events` | `Calendars.ReadWrite` | Create a new calendar event |
| `GET /me/mailboxSettings` | `MailboxSettings.Read` * | Read the user's timezone setting |
| `POST /me/calendar/getSchedule` | `Calendars.Read` * | Query free/busy availability for one or more users |
| `GET /places/microsoft.graph.room` | `Place.Read.All` | List all meeting rooms in the organization |
| `GET /places/microsoft.graph.roomList` | `Place.Read.All` | List room lists (building groups) |
| `GET /places/{id}` | `Place.Read.All` | Get details of a specific room (capacity, building, floor, etc.) |

\* `Calendars.ReadWrite` is a superset that includes `Calendars.Read`
capabilities. Similarly, `MailboxSettings.Read` is implicitly available when the
user has consented to read their own profile. The code uses `Calendars.ReadWrite`
as the single calendar scope to support both read and write operations.

> **Note:** `Place.Read.All` requires **admin consent**.

### Teams Handler

**Source**: `office_mcp.msgraph.teams_handler.TeamsHandler`

Teams-related permissions are included within the `CHAT_SCOPE` constant (see
Chat Handler below). The Teams handler is read-only.

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /me/joinedTeams` | `Team.ReadBasic.All` | List all teams the signed-in user has joined |
| `GET /teams/{id}/channels` | `Channel.ReadBasic.All` | List channels within a team |
| `GET /teams/{id}/channels/{id}/messages` | `ChannelMessage.Read.All` | Read messages in a team channel |
| `GET /teams/{id}/members` | `TeamMember.Read.All` | List members and their roles in a team |

> **Note:** `Team.ReadBasic.All`, `ChannelMessage.Read.All`, and `TeamMember.Read.All`
> require **admin consent**. `Channel.ReadBasic.All` can be user-consented in
> most tenant configurations but is typically granted via admin consent as well.

### Chat Handler

**Source**: `office_mcp.msgraph.chat_handler.ChatHandler`

**Scope constant**: `CHAT_SCOPE = ["Chat.Read", "ChannelMessage.Read.All"]`

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /me/chats` | `Chat.Read` | List the user's 1:1, group, and meeting chats |
| `GET /me/chats/{id}/messages` | `Chat.Read` | Read messages within a specific chat |
| `GET /me/chats/{id}/members` | `Chat.Read` | List members of a chat |

The `CHAT_SCOPE` constant also includes `ChannelMessage.Read.All` because chat
and Teams channel operations are typically enabled together.

### Files Handler

**Source**: `office_mcp.msgraph.files_handler.FilesHandler`

**Scope constant**: `ONE_DRIVE_SCOPE = ["Files.Read.All", "Files.ReadWrite.All"]`

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /me/drives` | `Files.Read.All` | List all drives accessible to the user |
| `GET /me/drive` | `Files.Read.All` | Get the user's default OneDrive |
| `GET /me/drive/root/children` | `Files.Read.All` | List files and folders at the OneDrive root |
| `GET /drives/{id}/root/children` | `Files.Read.All` | List root items in a specific drive |
| `GET /me/drive/items/{id}/children` | `Files.Read.All` | List children of a folder in the user's OneDrive |
| `GET /drives/{id}/items/{id}/children` | `Files.Read.All` | List children of a folder in a specific drive |
| `GET /me/drive/items/{id}` | `Files.Read.All` | Get metadata for a single file or folder |
| `GET /drives/{id}/items/{id}` | `Files.Read.All` | Get metadata for a file in a specific drive |
| `GET /me/drive/items/{id}/content` | `Files.Read.All` | Download file content from the user's OneDrive |
| `GET /drives/{id}/items/{id}/content` | `Files.Read.All` | Download file content from a specific drive |
| `GET /me/drive/root/search(q='...')` | `Files.Read.All` | Search for files by name or content in the user's OneDrive |
| `GET /drives/{id}/root/search(q='...')` | `Files.Read.All` | Search for files in a specific drive |
| `GET /me/followedSites` | `Sites.Read.All` | List SharePoint sites the user follows |
| `GET /sites?search=...` | `Sites.Read.All` | Search for SharePoint sites by keyword |
| `GET /sites/{id}/drives` | `Sites.Read.All` | List document libraries in a SharePoint site |

`Files.ReadWrite.All` is included in the scope constant to support future write
operations. For read-only deployments, `Files.Read.All` is sufficient for all
current file and drive operations. `Sites.Read.All` is required for SharePoint
site discovery and browsing.

### Directory Handler

**Source**: `office_mcp.msgraph.directory_handler.DirectoryHandler`

**Scope constant**: `DIRECTORY_SCOPE = ["Directory.Read.All", "ProfilePhoto.Read.All"]`

| Graph API Endpoint | Permission | Purpose |
|---|---|---|
| `GET /users` | `User.Read.All` or `Directory.Read.All` | List users in the organization directory with rich fields (name, email, job title, department, manager, phone, office location) |
| `GET /users` (paginated, all pages) | `User.Read.All` or `Directory.Read.All` | Fetch all users across multiple pages (`@odata.nextLink`) |
| `GET /users/{id}/manager` | `User.Read.All` or `Directory.Read.All` | Get a user's manager |
| `GET /users/{id}/photo/$value` | `ProfilePhoto.Read.All` | Download a user's profile photo |

The `$expand=manager($select=id)` query parameter is used on user listings to
include manager information in a single request. This requires
`Directory.Read.All` rather than the simpler `User.Read.All`.

> **Note:** Both `Directory.Read.All` and `ProfilePhoto.Read.All` require
> **admin consent**.

## Consolidated Permission Reference

### Minimum Permission Set

For a **read-only** deployment that needs access to all handlers, the following
is the minimum set of delegated permissions:

| Permission | Admin Consent | Handlers |
|---|---|---|
| `User.Read` | No | Profile (sign-in, identity) |
| `Mail.Read` | No | Mail (read inbox, read emails) |
| `Calendars.Read` | No | Calendar (list calendars, view events, check availability) |
| `Place.Read.All` | Yes | Calendar (list meeting rooms, room details) |
| `Chat.Read` | No | Chat (list chats, read messages, list members) |
| `Team.ReadBasic.All` | Yes | Teams (list joined teams) |
| `Channel.ReadBasic.All` | No | Teams (list channels) |
| `ChannelMessage.Read.All` | Yes | Teams (read channel messages) |
| `TeamMember.Read.All` | Yes | Teams (list team members) |
| `Files.Read.All` | No | Files (browse OneDrive, download files, search) |
| `Sites.Read.All` | Yes | Files (SharePoint site discovery and browsing) |
| `Directory.Read.All` | Yes | Directory (list users, org hierarchy, manager chain) |
| `ProfilePhoto.Read.All` | Yes | Directory (user profile photos) |
| `offline_access` | No | Token refresh (persistent sessions) |

### Full Permission Set

The full permission set as defined in `OfficeUserInstance` includes read-write
capabilities for mail, calendar, files, and shared mailbox access:

| Permission | Admin Consent | Purpose |
|---|---|---|
| `User.Read` | No | Read signed-in user's profile |
| `User.ReadBasic.All` | No | Read basic profiles of all users (people search, mail resolution) |
| `User.Read.All` | Yes | Read full profiles of all users (directory) |
| `Directory.Read.All` | Yes | Read directory data (org hierarchy, manager chain) |
| `ProfilePhoto.Read.All` | Yes | Read profile photos of all users |
| `Mail.Read` | No | Read the signed-in user's mail |
| `Mail.Read.Shared` | No | Read mail in shared/delegated mailboxes |
| `Mail.ReadWrite` | No | Read and write the user's mail (drafts, categories, flags) |
| `Mail.ReadWrite.Shared` | No | Read and write shared mailbox emails |
| `Mail.Send` | No | Send mail as the signed-in user |
| `Mail.Send.Shared` | No | Send mail on behalf of shared mailboxes |
| `Calendars.ReadWrite` | No | Read and write calendar events |
| `Place.Read.All` | Yes | List meeting rooms and room details |
| `Chat.Read` | No | Read the user's chats and chat messages |
| `ChannelMessage.Read.All` | Yes | Read messages in team channels |
| `Team.ReadBasic.All` | Yes | List teams the user has joined |
| `Channel.ReadBasic.All` | No | List channels within a team |
| `TeamMember.Read.All` | Yes | List team members and their roles |
| `Files.Read.All` | No | Read all files the user can access (OneDrive and SharePoint) |
| `Files.ReadWrite.All` | No | Read and write all accessible files |
| `Sites.Read.All` | Yes | Read SharePoint sites and document libraries |
| `offline_access` | No | Obtain refresh tokens for persistent sessions |
| `openid` | No | OpenID Connect sign-in |
| `profile` | No | Read basic profile (name, picture) via OpenID Connect |
| `email` | No | Read the user's email address via OpenID Connect |

### Scope Constants in Code

The `OfficeUserInstance` class defines the following scope constants that are
combined when initializing the `MsGraphInstance`:

```python
class OfficeUserInstance:
    PROFILE_SCOPE = ["User.Read"]
    DIRECTORY_SCOPE = ["Directory.Read.All", "ProfilePhoto.Read.All"]
    MAIL_SCOPE = [
        "Mail.Read", "Mail.Read.Shared",
        "Mail.ReadWrite", "Mail.ReadWrite.Shared",
        "Mail.Send", "Mail.Send.Shared",
        "User.ReadBasic.All",
    ]
    CALENDAR_SCOPE = ["Calendars.ReadWrite", "Place.Read.All"]
    CHAT_SCOPE = ["Chat.Read", "ChannelMessage.Read.All"]
    ONE_DRIVE_SCOPE = ["Files.Read.All", "Files.ReadWrite.All"]
```

These lists are typically combined and passed as the `scopes` parameter when
constructing the `MsGraphInstance`. The actual scopes requested during
authorization are sent as a space-separated string in the `scope` parameter
of the OAuth token request.

> **Note:** Some permissions used by the Teams and directory handlers (e.g.,
> `Team.ReadBasic.All`, `TeamMember.Read.All`, `Channel.ReadBasic.All`,
> `Sites.Read.All`) are not explicitly listed in the scope constants but must
> be configured in the Azure AD app registration. They are either implicitly
> covered by broader scopes in the token or need to be added to the combined
> scope list for your deployment.

## Admin Consent Requirements

### Permissions Requiring Admin Consent

The following delegated permissions require an Azure AD administrator (Global
Administrator or Privileged Role Administrator) to grant consent before any user
in the tenant can use them:

| Permission | Reason |
|---|---|
| `Directory.Read.All` | Allows reading all directory data including user profiles, groups, and organizational relationships across the entire tenant |
| `User.Read.All` | Allows reading the full profile of any user in the organization |
| `ProfilePhoto.Read.All` | Allows reading the profile photo of any user in the organization |
| `ChannelMessage.Read.All` | Allows reading messages in any team channel the user is a member of |
| `Team.ReadBasic.All` | Allows reading basic properties of all teams the user has joined |
| `TeamMember.Read.All` | Allows reading team membership information |
| `Sites.Read.All` | Allows reading SharePoint sites across the tenant |
| `Place.Read.All` | Allows reading meeting room information (names, email addresses, capacity, building, floor) across the organization |

### User-Consentable Permissions

The following permissions can be consented to by individual users without
administrator involvement:

- `User.Read`
- `User.ReadBasic.All`
- `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`
- `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`
- `Calendars.Read`, `Calendars.ReadWrite`
- `Chat.Read`
- `Channel.ReadBasic.All`
- `Files.Read.All`, `Files.ReadWrite.All`
- `offline_access`, `openid`, `profile`, `email`

### Granting Tenant-Wide Admin Consent

For enterprise deployments, it is recommended to grant **tenant-wide admin
consent** so that no individual user sees a consent dialog. This covers both
admin-required and user-consentable permissions.

**Via the Azure Portal:**

1. Navigate to **Microsoft Entra ID** > **App registrations**.
2. Open your application registration.
3. Go to **API permissions**.
4. Click **Grant admin consent for Your Organization**.
5. Confirm the prompt.

**Via Azure CLI:**

```bash
az ad app permission admin-consent --id <your-application-client-id>
```

This must be run by an account with Global Administrator or Privileged Role
Administrator privileges.

**Effects of tenant-wide consent:**

- All configured permissions are consented for every user in the tenant.
- No user will see any consent dialog, even for user-consentable permissions.
- Existing users pick up new scopes automatically on their next token refresh
  without needing to re-authenticate.

## Feature-Based Permission Selection

When deploying `office-mcp`, you may not need all handlers. The following table
maps deployment features to the permission sets you should configure:

| Feature | Scope Constant | Permissions |
|---|---|---|
| User sign-in and identity | `PROFILE_SCOPE` | `User.Read`, `openid`, `profile`, `email`, `offline_access` |
| Email access | `MAIL_SCOPE` | `Mail.Read`, `Mail.Read.Shared`, `Mail.ReadWrite`, `Mail.ReadWrite.Shared`, `Mail.Send`, `Mail.Send.Shared`, `User.ReadBasic.All` |
| Calendar access | `CALENDAR_SCOPE` | `Calendars.ReadWrite`, `Place.Read.All` |
| Teams and chat | `CHAT_SCOPE` | `Chat.Read`, `ChannelMessage.Read.All`, `Team.ReadBasic.All`, `Channel.ReadBasic.All`, `TeamMember.Read.All` |
| OneDrive and SharePoint | `ONE_DRIVE_SCOPE` | `Files.Read.All`, `Files.ReadWrite.All`, `Sites.Read.All` |
| Organization directory | `DIRECTORY_SCOPE` | `Directory.Read.All`, `ProfilePhoto.Read.All` |

All features always require `PROFILE_SCOPE` as a baseline, since user identity
is needed for every authenticated API call.

## MCP Server Mode

When running `office-mcp` as a standalone MCP server (via `python -m
office_mcp.mcp_server --keyfile token.json`), scopes are not requested at
authorization time by the server itself. Instead, the JSON key file must contain
tokens that were originally obtained with the appropriate scopes. The MCP server
loads the access token and refresh token from the key file and refreshes tokens as
needed.

Key file format:

```json
{
  "app": "office-mcp",
  "session_id": "optional-session-id",
  "email": "user@example.com",
  "access_token": "eyJ...",
  "refresh_token": "1.AUs...",
  "client_id": "your-client-id",
  "client_secret": "your-client-secret",
  "tenant_id": "your-tenant-id"
}
```

The key file should be created with restrictive file permissions (`0600`) since
it contains sensitive credentials. The `export_keyfile` function in
`office_mcp.mcp_server` handles this automatically.

### MCP Server Tool-to-Permission Mapping

The following table maps each MCP server tool to the permissions it requires:

| MCP Tool | Handler | Required Permissions |
|---|---|---|
| `o365_get_profile` | ProfileHandler | `User.Read` |
| `o365_list_mail` | OfficeMailHandler | `Mail.Read` |
| `o365_get_mail` | OfficeMailHandler | `Mail.Read` |
| `o365_get_mail_categories` | OfficeMailHandler | `Mail.Read` |
| `o365_list_calendars` | CalendarHandler | `Calendars.Read` |
| `o365_get_events` | CalendarHandler | `Calendars.Read` |
| `o365_get_schedule` | CalendarHandler | `Calendars.Read` |
| `o365_list_teams` | TeamsHandler | `Team.ReadBasic.All` |
| `o365_list_channels` | TeamsHandler | `Channel.ReadBasic.All` |
| `o365_get_channel_messages` | TeamsHandler | `ChannelMessage.Read.All` |
| `o365_get_team_members` | TeamsHandler | `TeamMember.Read.All` |
| `o365_list_chats` | ChatHandler | `Chat.Read` |
| `o365_get_chat_messages` | ChatHandler | `Chat.Read` |
| `o365_get_chat_members` | ChatHandler | `Chat.Read` |
| `o365_get_my_drive` | FilesHandler | `Files.Read.All` |
| `o365_list_drive_items` | FilesHandler | `Files.Read.All` |
| `o365_get_file_content` | FilesHandler | `Files.Read.All` |
| `o365_search_files` | FilesHandler | `Files.Read.All` |
| `o365_search_sites` | FilesHandler | `Sites.Read.All` |
| `o365_get_site_drives` | FilesHandler | `Sites.Read.All` |
| `o365_list_users` | DirectoryHandler | `Directory.Read.All` |
| `o365_get_user_manager` | DirectoryHandler | `Directory.Read.All` |

## Security Considerations

- **Principle of least privilege**: Only configure the permissions your deployment
  actually needs. If you do not use the Teams handler, omit `Team.ReadBasic.All`,
  `ChannelMessage.Read.All`, and `TeamMember.Read.All`.

- **Delegated-only access**: The application never accesses data beyond what the
  signed-in user can access. There are no application-level permissions that would
  grant tenant-wide data access independent of a user context.

- **Redirect URI validation**: The `azure_auth_utils` module validates redirect
  hosts against an allowlist (`ALLOWED_REDIRECT_HOSTS`) to prevent open-redirect
  attacks during the OAuth flow.

- **Token storage**: Access tokens and refresh tokens should be stored securely.
  The MCP server key file is written with `0600` permissions. In web deployments,
  tokens are stored in Redis with per-user keys.

- **No write access to SharePoint structure**: While `Sites.Read.All` provides
  read access to SharePoint sites, no permissions are configured for modifying site
  structure, lists, or site-level settings.

- **Client secret rotation**: Rotate client secrets regularly. Azure AD allows
  multiple concurrent secrets to enable zero-downtime rotation.

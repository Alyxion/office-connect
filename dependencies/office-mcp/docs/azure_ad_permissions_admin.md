# Microsoft Graph API — Required Delegated Permissions

**Permission type:** Delegated (on behalf of the signed-in user)
**Date:** 2026-03-10

---

## Summary

This application requires the following Microsoft Graph delegated permissions to provide full Office 365 integration. All access is performed on behalf of the signed-in user — the application does not use application-level permissions.

To ensure **no user ever sees a consent dialog**, a Global Administrator must grant tenant-wide admin consent after all permissions are configured. See "Admin Consent" at the bottom of this document.

---

## 1. User Profile & Directory

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `User.Read` | No | Read signed-in user's profile | Login, display name, email, user ID |
| `User.Read.All` | Yes | Read all users' full profiles | Directory lookups, org chart |
| `User.ReadBasic.All` | No | Read basic profile of all users | People search, user autocomplete |
| `Directory.Read.All` | Yes | Read directory data | Org hierarchy, department info, manager chain |
| `ProfilePhoto.Read.All` | Yes | Read profile photos of all users | Display user avatars in UI |
| `People.Read` | No | Read relevant people | "People you work with" suggestions |
| `Presence.Read.All` | Yes | Read presence of all users | Show online/busy/away status of colleagues |

## 2. Mail

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `Mail.Read` | No | Read user's mail | Inbox display, email processing |
| `Mail.ReadWrite` | No | Read and write user's mail | Move, categorize, flag emails |
| `Mail.Send` | No | Send mail as the user | Send replies and new emails |
| `Mail.Read.Shared` | No | Read shared mailboxes | Access to shared/delegated mailboxes |
| `Mail.ReadWrite.Shared` | No | Read/write shared mailboxes | Manage shared mailbox emails |
| `Mail.Send.Shared` | No | Send from shared mailboxes | Send on behalf of shared mailboxes |
| `Mail.ReadBasic` | No | Read basic mail properties | Efficient mail index (subject, sender, date) |
| `Mail.ReadBasic.Shared` | No | Read basic shared mail properties | Shared mailbox listing |
| `MailboxSettings.ReadWrite` | No | Read/write mailbox settings | Auto-replies (out-of-office), timezone, working hours |

## 3. Calendar & Rooms

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `Calendars.Read` | No | Read user's calendars | View calendar events |
| `Calendars.ReadWrite` | No | Read and write calendars | Create/update calendar events |
| `Calendars.ReadWrite.Shared` | No | Read/write shared calendars | Access shared/delegated calendars |
| `Place.Read.All` | Yes | Read room and workspace info | List meeting rooms, check room availability, room details (capacity, building, floor) |

## 4. Microsoft Teams — Chats & Channels

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `Chat.Read` | No | Read user's chats | View personal and group chat messages |
| `Chat.ReadWrite` | No | Read and write chats | Send chat messages |
| `Chat.Create` | No | Create new chats | Initiate new 1:1 or group chats |
| `Channel.ReadBasic.All` | No | Read team channels | List channels within a team |
| `ChannelMessage.Read.All` | Yes | Read channel messages | View messages in team channels |
| `ChannelMessage.ReadWrite` | No | Read and write channel messages | Post messages to team channels |
| `Team.ReadBasic.All` | Yes | Read joined teams | List teams the user is a member of |
| `TeamMember.Read.All` | Yes | Read team members | List members and roles within a team |

## 5. OneDrive & SharePoint

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `Files.Read.All` | No | Read all files the user can access | Browse OneDrive and SharePoint files |
| `Files.ReadWrite` | No | Read/write user's OneDrive files | Upload and manage personal files |
| `Files.ReadWrite.All` | No | Read/write all accessible files | Manage files across OneDrive and SharePoint |
| `Sites.Read.All` | Yes | Read SharePoint sites | Discover SharePoint sites, list document libraries |

## 6. Contacts

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `Contacts.Read` | No | Read user's contacts | Access Outlook contacts for sales context, customer lookup |

## 7. Online Meetings

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `OnlineMeetings.ReadWrite` | No | Create and read online meetings | Generate Teams meeting links for customer calls |

## 8. Tasks

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `Tasks.ReadWrite` | No | Read and write user's tasks | Create follow-up tasks from emails and chats |

## 9. Standard OpenID Connect

| Permission | Admin Consent Required | Description | Used for |
|---|---|---|---|
| `openid` | No | Sign in | OpenID Connect authentication |
| `profile` | No | Read basic profile | Name, picture |
| `email` | No | Read email address | User email for identification |
| `offline_access` | No | Maintain access | Refresh tokens for persistent sessions |

---

## Complete Permission List (39 total)

```
Calendars.Read
Calendars.ReadWrite
Calendars.ReadWrite.Shared
Channel.ReadBasic.All
ChannelMessage.Read.All
ChannelMessage.ReadWrite
Chat.Create
Chat.Read
Chat.ReadWrite
Contacts.Read
Directory.Read.All
email
Files.Read.All
Files.ReadWrite
Files.ReadWrite.All
Mail.Read
Mail.Read.Shared
Mail.ReadBasic
Mail.ReadBasic.Shared
Mail.ReadWrite
Mail.ReadWrite.Shared
Mail.Send
Mail.Send.Shared
MailboxSettings.ReadWrite
offline_access
OnlineMeetings.ReadWrite
openid
People.Read
Place.Read.All
Presence.Read.All
profile
ProfilePhoto.Read.All
Sites.Read.All
Tasks.ReadWrite
Team.ReadBasic.All
TeamMember.Read.All
User.Read
User.Read.All
User.ReadBasic.All
```

---

## Admin Consent

After all permissions are configured in the Azure AD app registration, a **Global Administrator** or **Privileged Role Administrator** must grant **tenant-wide admin consent**. This pre-approves all permissions for every user in the tenant, so no individual user will ever see a consent dialog.

### Via Azure Portal

1. Navigate to **Microsoft Entra ID** > **App registrations**.
2. Open the application registration.
3. Go to **API permissions**.
4. Click **"Grant admin consent for [Organization]"**.
5. Confirm the prompt.

### Via Azure CLI

```bash
az ad app permission admin-consent --id <application-client-id>
```

### Effect

- Grants consent for **all** configured permissions tenant-wide (both admin-required and user-consentable).
- No user will see any consent dialog — not even for permissions like `Contacts.Read` that normally allow user-level consent.
- Existing users pick up new scopes automatically on their next token refresh. No re-login required.

---

## Notes

- All permissions are **delegated** (not application). The app acts on behalf of the signed-in user and can only access what the user themselves has access to.
- Permissions requiring admin consent: `Directory.Read.All`, `User.Read.All`, `ProfilePhoto.Read.All`, `ChannelMessage.Read.All`, `Team.ReadBasic.All`, `TeamMember.Read.All`, `Presence.Read.All`, `Sites.Read.All`, `Place.Read.All`.
- No permissions grant write access to SharePoint sites or team structure — all SharePoint and Teams access is read-only at the data level.

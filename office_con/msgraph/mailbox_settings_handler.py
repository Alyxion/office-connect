"""Microsoft Graph Mailbox Settings handler (async-only, read-only).

Provides read access to:
- Full mailbox settings
- Automatic replies (out-of-office) configuration
- Working hours configuration
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance

class MailboxSettingsHandler:
    """Handler for Microsoft Graph Mailbox Settings API (read-only, async-only).

    Requires MailboxSettings.Read scope.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    async def get_mailbox_settings_async(self) -> dict:
        """Get the current user's full mailbox settings."""
        token = await self.msg.get_access_token_async()
        if not token:
            return {}
        url = f"{self.msg.msg_endpoint}me/mailboxSettings"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return {}
        return resp.json()

    async def get_automatic_replies_async(self) -> dict:
        """Get the current user's automatic replies (out-of-office) settings."""
        token = await self.msg.get_access_token_async()
        if not token:
            return {}
        url = f"{self.msg.msg_endpoint}me/mailboxSettings/automaticRepliesSetting"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return {}
        return resp.json()

    async def get_working_hours_async(self) -> dict:
        """Get the current user's working hours configuration."""
        token = await self.msg.get_access_token_async()
        if not token:
            return {}
        url = f"{self.msg.msg_endpoint}me/mailboxSettings/workingHours"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return {}
        return resp.json()

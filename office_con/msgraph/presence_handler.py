"""Microsoft Graph Presence handler (async-only, read-only).

Provides read access to:
- Current user's presence status
- Other users' presence status
- Batch presence lookup by user IDs
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class PresenceHandler:
    """Handler for Microsoft Graph Presence API (read-only, async-only).

    Requires Presence.Read and Presence.Read.All scopes.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    async def get_my_presence_async(self) -> dict:
        """Get the current user's presence status."""
        token = await self.msg.get_access_token_async()
        if not token:
            return {}
        url = f"{self.msg.msg_endpoint}me/presence"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return {}
        return resp.json()

    async def get_user_presence_async(self, user_id: str) -> dict:
        """Get presence status for a specific user.

        Args:
            user_id: Azure AD user ID.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return {}
        url = f"{self.msg.msg_endpoint}users/{user_id}/presence"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return {}
        return resp.json()

    async def get_presences_async(self, user_ids: list[str]) -> list[dict]:
        """Get presence status for multiple users in a single batch call.

        Args:
            user_ids: List of Azure AD user IDs (max 650 per request).
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}communications/getPresencesByUserId"
        body = {"ids": user_ids}
        resp = await self.msg.run_async(url=url, method="POST", json=body, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

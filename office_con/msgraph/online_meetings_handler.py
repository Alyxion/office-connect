"""Microsoft Graph Online Meetings handler (async-only, read-only).

Provides read access to:
- Online meetings list
- Individual meeting details
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class OnlineMeetingsHandler:
    """Handler for Microsoft Graph Online Meetings API (read-only, async-only).

    Requires OnlineMeetings.Read scope.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    async def get_meetings_async(self, limit: int = 20) -> list[dict]:
        """List the current user's online meetings.

        Args:
            limit: Maximum number of meetings to return.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}me/onlineMeetings?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

    async def get_meeting_async(self, meeting_id: str) -> dict | None:
        """Get details of a specific online meeting.

        Args:
            meeting_id: The ID of the online meeting.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        url = f"{self.msg.msg_endpoint}me/onlineMeetings/{meeting_id}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return resp.json()

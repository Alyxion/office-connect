"""Microsoft Graph Places handler (async-only, read-only).

Provides read access to:
- Meeting rooms
- Room lists (groupings of rooms)
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class PlacesHandler:
    """Handler for Microsoft Graph Places API (read-only, async-only).

    Requires Place.Read.All scope.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    async def get_rooms_async(self, limit: int = 50) -> list[dict]:
        """List available meeting rooms.

        Args:
            limit: Maximum number of rooms to return.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}places/microsoft.graph.room?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

    async def get_room_lists_async(self) -> list[dict]:
        """List room lists (logical groupings of rooms)."""
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}places/microsoft.graph.roomList"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

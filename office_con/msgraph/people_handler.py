"""Microsoft Graph People handler (async-only, read-only).

Provides read access to:
- Relevant people (based on communication and collaboration patterns)
- People search
- Personal contacts
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class PeopleHandler:
    """Handler for Microsoft Graph People API (read-only, async-only).

    Requires People.Read and Contacts.Read scopes.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    async def get_relevant_people_async(self, limit: int = 20) -> list[dict]:
        """Get people most relevant to the current user.

        Relevance is determined by communication and collaboration patterns,
        business relationships, and organizational proximity.

        Args:
            limit: Maximum number of people to return.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}me/people?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

    async def search_people_async(self, query: str, limit: int = 10) -> list[dict]:
        """Search for people by name or email.

        Args:
            query: Search query string (name, email, etc.).
            limit: Maximum number of results to return.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f'{self.msg.msg_endpoint}me/people?$search="{query}"&$top={limit}'
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

    async def get_contacts_async(self, limit: int = 50) -> list[dict]:
        """List the current user's personal contacts.

        Args:
            limit: Maximum number of contacts to return.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}me/contacts?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

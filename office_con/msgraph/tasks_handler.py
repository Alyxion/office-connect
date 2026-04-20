"""Microsoft Graph To Do Tasks handler (async-only, read-only).

Provides read access to:
- To Do task lists
- Tasks within a list
- Individual task details
"""

from __future__ import annotations

from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class TasksHandler:
    """Handler for Microsoft Graph To Do Tasks API (read-only, async-only).

    Requires Tasks.Read scope.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    async def get_task_lists_async(self) -> list[dict]:
        """List all To Do task lists for the current user."""
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}me/todo/lists"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

    async def get_tasks_async(self, list_id: str, limit: int = 50) -> list[dict]:
        """List tasks in a specific To Do task list.

        Args:
            list_id: The ID of the task list.
            limit: Maximum number of tasks to return.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return []
        url = f"{self.msg.msg_endpoint}me/todo/lists/{list_id}/tasks?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return []
        return resp.json().get("value", [])

    async def get_task_async(self, list_id: str, task_id: str) -> dict | None:
        """Get a specific task from a To Do task list.

        Args:
            list_id: The ID of the task list.
            task_id: The ID of the task.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        url = f"{self.msg.msg_endpoint}me/todo/lists/{list_id}/tasks/{task_id}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return resp.json()

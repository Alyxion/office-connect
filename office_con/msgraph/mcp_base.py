"""Abstract base for in-process MCP servers and MS Graph MCP server base."""

from abc import ABC, abstractmethod
from typing import Any, Dict, List, TYPE_CHECKING

if TYPE_CHECKING:
    from office_con.msgraph.ms_graph_handler import MsGraphInstance


class InProcessMCPServer(ABC):
    """Abstract base class for in-process MCP servers.

    Subclass this to create MCP-compatible tool servers that run in the same
    process as the host application, with direct access to runtime state
    (e.g. authenticated user sessions, database connections).
    """

    @abstractmethod
    async def list_tools(self) -> List[Dict[str, Any]]:
        """List available tools."""
        pass

    @abstractmethod
    async def call_tool(self, name: str, arguments: Dict[str, Any]) -> str:
        """Call a tool by name."""
        pass

    async def get_prompt_hints(self) -> List[str]:
        """Optional prompt snippets appended to the system prompt."""
        return []

    async def get_client_renderers(self) -> List[Dict[str, str]]:
        """Optional client-side renderers for custom fenced code block languages."""
        return []


class MsGraphMCPServer(InProcessMCPServer):
    """Base in-process MCP server with access to an authenticated MS Graph instance.

    Subclasses implement list_tools() and call_tool() to expose
    Graph API operations as MCP-compatible tools.
    """

    def __init__(self, graph: "MsGraphInstance"):
        self.graph = graph
        self.show_room_booking_names: bool = False
        self.room_exclude_patterns: list[str] = []
        self.room_domain_filter: str = ""

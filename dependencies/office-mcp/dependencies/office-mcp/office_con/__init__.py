"""office-mcp -- Microsoft 365 / MS Graph logic layer with MCP server support."""

__version__ = "0.1.0"

from .web_user_instance import WebUserInstance
from .db_user_instance import DBUserInstance
from .msgraph.ms_graph_handler import MsGraphInstance
from .msgraph.mail_handler import OfficeMail, OfficeMailHandler
from .msgraph.calendar_handler import CalendarHandler, CalendarEvent, CalendarEventList
from .msgraph.profile_handler import ProfileHandler, UserProfile
from .msgraph.directory_handler import DirectoryHandler, DirectoryUser
from .msgraph.teams_handler import TeamsHandler
from .msgraph.chat_handler import ChatHandler
from .msgraph.files_handler import FilesHandler
from .msgraph.mcp_base import InProcessMCPServer, MsGraphMCPServer
from .msgraph.office365_server import Office365MCPServer
from .db import CompanyDir, CompanyDirBuilder

__all__ = [
    # Session layer
    "WebUserInstance",
    "DBUserInstance",
    # MS Graph
    "MsGraphInstance",
    "OfficeMail",
    "OfficeMailHandler",
    "CalendarHandler",
    "CalendarEvent",
    "CalendarEventList",
    "ProfileHandler",
    "UserProfile",
    "DirectoryHandler",
    "DirectoryUser",
    "TeamsHandler",
    "ChatHandler",
    "FilesHandler",
    # MCP servers
    "InProcessMCPServer",
    "MsGraphMCPServer",
    "Office365MCPServer",
    # Company directory
    "CompanyDir",
    "CompanyDirBuilder",
]

from threading import RLock
from typing import Any, Optional

from pydantic import BaseModel
from office_con.msgraph import MsGraphInstance
from office_con import WebUserInstance, DBUserInstance


class OfficeUserConfig(BaseModel):
    """Base config class that can be extended by specific implementations"""
    pass


class OfficeUserInstance:
    """High-level Office 365 user session — wraps MsGraphInstance with scoped feature access."""
    PROFILE_SCOPE = ["User.Read", "User.Read.All", "User.ReadBasic.All",
                     "People.Read", "Presence.Read.All"]
    """User profile, people search, and presence status."""
    DIRECTORY_SCOPE = ["Directory.Read.All", "ProfilePhoto.Read.All", "Contacts.Read"]
    """Company directory, profile photos, and contacts."""
    MAIL_SCOPE = ["Mail.Read", "Mail.ReadBasic", "Mail.Read.Shared", "Mail.ReadBasic.Shared",
                  "Mail.ReadWrite", "Mail.ReadWrite.Shared", "Mail.Send", "Mail.Send.Shared",
                  "MailboxSettings.ReadWrite"]
    """Read, write, and send mail (own + shared mailboxes) + mailbox settings."""
    CALENDAR_SCOPE = ["Calendars.Read", "Calendars.ReadWrite", "Calendars.ReadWrite.Shared",
                      "Place.Read.All", "OnlineMeetings.ReadWrite"]
    """Calendars (own + shared), room/place lookup, and Teams meetings."""
    CHAT_SCOPE = ["Chat.Read", "Chat.ReadWrite", "Chat.Create",
                  "ChannelMessage.Read.All", "ChannelMessage.ReadWrite"]
    """Teams chats and channel messages."""
    TEAMS_SCOPE = ["Team.ReadBasic.All", "TeamMember.Read.All", "Channel.ReadBasic.All"]
    """Teams structure — teams, channels, and members."""
    ONE_DRIVE_SCOPE = ["Files.Read.All", "Files.ReadWrite", "Files.ReadWrite.All",
                       "Sites.Read.All"]
    """OneDrive, SharePoint files, and site document libraries."""
    TASKS_SCOPE = ["Tasks.ReadWrite"]
    """Microsoft To Do / Planner tasks."""

    def __init__(self, config: Optional[OfficeUserConfig] = None,
                 user_instance: MsGraphInstance | DBUserInstance | WebUserInstance | None = None):
        self.session_id = user_instance.session_id if user_instance else None
        self.config = config or OfficeUserConfig()
        self.access_lock = RLock()
        self.user_instance: MsGraphInstance | DBUserInstance | WebUserInstance | None = user_instance
        self.app_data: dict[str, Any] = {}

    def get_office_handler(self) -> MsGraphInstance | WebUserInstance | None:
        return self.user_instance

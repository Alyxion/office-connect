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
    PROFILE_SCOPE = ["User.Read"]
    """A scope to read the user profile such as name, email, etc."""
    DIRECTORY_SCOPE = ["Directory.Read.All", "ProfilePhoto.Read.All"]
    """A scope to read the directory such as hierarchy and profile images."""
    MAIL_SCOPE = ["Mail.Read", "Mail.Read.Shared", "Mail.ReadWrite", "Mail.ReadWrite.Shared", "Mail.Send",
                  "Mail.Send.Shared", "User.ReadBasic.All"] #
    """A scope to read and write mails."""
    CALENDAR_SCOPE = ["Calendars.ReadWrite"]
    """A scope to read and write the calendar."""
    CHAT_SCOPE = ["Chat.Read", "ChannelMessage.Read.All"]
    """A scope to read and write the chat."""
    ONE_DRIVE_SCOPE = ["Files.Read.All", "Files.ReadWrite.All"]
    """A scope to read and write from and to OneDrive."""

    def __init__(self, config: Optional[OfficeUserConfig] = None,
                 user_instance: MsGraphInstance | DBUserInstance | WebUserInstance | None = None):
        self.session_id = user_instance.session_id if user_instance else None
        self.config = config or OfficeUserConfig()
        self.access_lock = RLock()
        self.user_instance: MsGraphInstance | DBUserInstance | WebUserInstance | None = user_instance
        self.app_data: dict[str, Any] = {}

    def get_office_handler(self) -> MsGraphInstance | WebUserInstance | None:
        return self.user_instance

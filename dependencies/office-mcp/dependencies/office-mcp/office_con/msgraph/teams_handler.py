"""Microsoft Graph Teams & Channels handler (async-only).

Provides read access to:
- Joined teams
- Team channels
- Channel messages
"""

from __future__ import annotations

from datetime import datetime
from typing import List, Optional, TYPE_CHECKING

from pydantic import BaseModel, Field

if TYPE_CHECKING:
    from office_con import MsGraphInstance


class Team(BaseModel):
    """A Microsoft Teams team."""
    id: str = Field(..., description="Team (group) ID")
    display_name: Optional[str] = Field(default=None, description="Display name of the team")
    description: Optional[str] = Field(default=None, description="Team description")
    visibility: Optional[str] = Field(default=None, description="Visibility: public, private, or hiddenMembership")
    is_archived: bool = Field(default=False, description="Whether the team is archived")
    web_url: Optional[str] = Field(default=None, description="Deep-link URL to the team in Teams client")


class TeamList(BaseModel):
    """Paginated list of teams."""
    teams: List[Team] = Field(default_factory=list, description="Joined teams")
    total_teams: int = Field(default=0, description="Total number of teams")


class Channel(BaseModel):
    """A channel within a team."""
    id: str = Field(..., description="Channel ID")
    display_name: Optional[str] = Field(default=None, description="Display name of the channel")
    description: Optional[str] = Field(default=None, description="Channel description")
    membership_type: Optional[str] = Field(default=None, description="standard, private, or shared")
    web_url: Optional[str] = Field(default=None, description="Deep-link URL to the channel")
    is_favorite_by_default: Optional[bool] = Field(default=None, description="Auto-favorite for new members")


class ChannelList(BaseModel):
    """List of channels in a team."""
    channels: List[Channel] = Field(default_factory=list, description="Team channels")
    total_channels: int = Field(default=0, description="Total number of channels")


class ChannelMessageFrom(BaseModel):
    """Sender of a channel message."""
    display_name: Optional[str] = Field(default=None, description="Display name of the sender")
    email: Optional[str] = Field(default=None, description="Email of the sender")
    user_id: Optional[str] = Field(default=None, description="Azure AD user ID")


class ChannelMessage(BaseModel):
    """A message in a Teams channel."""
    id: str = Field(..., description="Message ID")
    created_at: Optional[datetime] = Field(default=None, description="When the message was created (UTC)")
    subject: Optional[str] = Field(default=None, description="Message subject (thread root)")
    body_content: Optional[str] = Field(default=None, description="Message body text or HTML")
    body_type: Optional[str] = Field(default=None, description="text or html")
    sender: Optional[ChannelMessageFrom] = Field(default=None, description="Who sent the message")
    importance: Optional[str] = Field(default=None, description="normal, high, or urgent")
    web_url: Optional[str] = Field(default=None, description="Deep-link to the message")


class ChannelMessageList(BaseModel):
    """Paginated list of channel messages."""
    messages: List[ChannelMessage] = Field(default_factory=list, description="Channel messages")
    total_messages: int = Field(default=0, description="Total number of messages")


class TeamMember(BaseModel):
    """A member of a team."""
    id: str = Field(..., description="Membership ID")
    display_name: Optional[str] = Field(default=None, description="Display name")
    email: Optional[str] = Field(default=None, description="Email address")
    user_id: Optional[str] = Field(default=None, description="Azure AD user ID")
    roles: List[str] = Field(default_factory=list, description="Roles: owner, member, guest")


class TeamMemberList(BaseModel):
    """List of team members."""
    members: List[TeamMember] = Field(default_factory=list, description="Team members")
    total_members: int = Field(default=0, description="Total number of members")


class TeamsHandler:
    """Handler for Microsoft Graph Teams & Channels API (read-only, async-only)."""

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    # ── parsing helpers (no I/O) ──────────────────────────────────────────

    @staticmethod
    def _parse_team(data: dict) -> Team:
        return Team(
            id=data["id"],
            display_name=data.get("displayName"),
            description=data.get("description"),
            visibility=data.get("visibility"),
            is_archived=data.get("isArchived", False),
            web_url=data.get("webUrl"),
        )

    @staticmethod
    def _parse_channel(data: dict) -> Channel:
        return Channel(
            id=data["id"],
            display_name=data.get("displayName"),
            description=data.get("description"),
            membership_type=data.get("membershipType"),
            web_url=data.get("webUrl"),
            is_favorite_by_default=data.get("isFavoriteByDefault"),
        )

    @staticmethod
    def _parse_channel_message(data: dict) -> ChannelMessage:
        sender_data = (data.get("from") or {}).get("user") or {}
        sender = ChannelMessageFrom(
            display_name=sender_data.get("displayName"),
            email=sender_data.get("userIdentityType"),
            user_id=sender_data.get("id"),
        ) if sender_data else None

        created = data.get("createdDateTime")
        created_dt = None
        if created:
            try:
                created_dt = datetime.fromisoformat(created.replace("Z", "+00:00"))
            except (ValueError, TypeError):
                pass

        body = data.get("body") or {}
        return ChannelMessage(
            id=data["id"],
            created_at=created_dt,
            subject=data.get("subject"),
            body_content=body.get("content"),
            body_type=body.get("contentType"),
            sender=sender,
            importance=data.get("importance"),
            web_url=data.get("webUrl"),
        )

    @staticmethod
    def _parse_member(data: dict) -> TeamMember:
        return TeamMember(
            id=data.get("id", ""),
            display_name=data.get("displayName"),
            email=data.get("email"),
            user_id=data.get("userId"),
            roles=data.get("roles", []),
        )

    # ── async API ─────────────────────────────────────────────────────────

    async def get_joined_teams_async(self) -> TeamList:
        """List all teams the current user has joined."""
        token = await self.msg.get_access_token_async()
        if not token:
            return TeamList()
        url = f"{self.msg.msg_endpoint}me/joinedTeams"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return TeamList()
        data = resp.json()
        teams = [self._parse_team(t) for t in data.get("value", [])]
        return TeamList(teams=teams, total_teams=len(teams))

    async def get_channels_async(self, team_id: str) -> ChannelList:
        """List channels in a team."""
        token = await self.msg.get_access_token_async()
        if not token:
            return ChannelList()
        url = f"{self.msg.msg_endpoint}teams/{team_id}/channels"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return ChannelList()
        data = resp.json()
        channels = [self._parse_channel(c) for c in data.get("value", [])]
        return ChannelList(channels=channels, total_channels=len(channels))

    async def get_channel_messages_async(
        self, team_id: str, channel_id: str, limit: int = 20
    ) -> ChannelMessageList:
        """List recent messages in a channel."""
        token = await self.msg.get_access_token_async()
        if not token:
            return ChannelMessageList()
        url = f"{self.msg.msg_endpoint}teams/{team_id}/channels/{channel_id}/messages?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return ChannelMessageList()
        data = resp.json()
        messages = [self._parse_channel_message(m) for m in data.get("value", [])]
        return ChannelMessageList(messages=messages, total_messages=len(messages))

    async def get_team_members_async(self, team_id: str) -> TeamMemberList:
        """List members of a team."""
        token = await self.msg.get_access_token_async()
        if not token:
            return TeamMemberList()
        url = f"{self.msg.msg_endpoint}teams/{team_id}/members"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return TeamMemberList()
        data = resp.json()
        members = [self._parse_member(m) for m in data.get("value", [])]
        return TeamMemberList(members=members, total_members=len(members))

"""Microsoft Graph Chat handler (async-only).

Provides read access to:
- Personal and group chats
- Chat messages
- Chat members
"""

from __future__ import annotations

from datetime import datetime
from typing import List, Optional, TYPE_CHECKING

from pydantic import BaseModel, Field

if TYPE_CHECKING:
    from office_con import MsGraphInstance


# ── Pydantic models ──────────────────────────────────────────────────────


class ChatMember(BaseModel):
    """A member of a chat."""
    id: str = Field(..., description="Membership ID")
    display_name: Optional[str] = Field(default=None, description="Display name")
    email: Optional[str] = Field(default=None, description="Email address")
    user_id: Optional[str] = Field(default=None, description="Azure AD user ID")
    roles: List[str] = Field(default_factory=list, description="Roles in the chat")


class ChatMemberList(BaseModel):
    """List of chat members."""
    members: List[ChatMember] = Field(default_factory=list, description="Chat members")
    total_members: int = Field(default=0, description="Total number of members")


class ChatMessageFrom(BaseModel):
    """Sender of a chat message."""
    display_name: Optional[str] = Field(default=None, description="Display name of the sender")
    email: Optional[str] = Field(default=None, description="Email of the sender")
    user_id: Optional[str] = Field(default=None, description="Azure AD user ID")


class ChatMessage(BaseModel):
    """A message in a chat."""
    id: str = Field(..., description="Message ID")
    created_at: Optional[datetime] = Field(default=None, description="When the message was created (UTC)")
    body_content: Optional[str] = Field(default=None, description="Message body text or HTML")
    body_type: Optional[str] = Field(default=None, description="text or html")
    sender: Optional[ChatMessageFrom] = Field(default=None, description="Who sent the message")
    importance: Optional[str] = Field(default=None, description="normal, high, or urgent")
    message_type: Optional[str] = Field(default=None, description="message, chatEvent, typing, etc.")
    web_url: Optional[str] = Field(default=None, description="Deep-link to the message")


class ChatMessageList(BaseModel):
    """Paginated list of chat messages."""
    messages: List[ChatMessage] = Field(default_factory=list, description="Chat messages")
    total_messages: int = Field(default=0, description="Total number of messages")


class Chat(BaseModel):
    """A personal or group chat."""
    id: str = Field(..., description="Chat ID")
    topic: Optional[str] = Field(default=None, description="Chat topic (group chats)")
    chat_type: Optional[str] = Field(default=None, description="oneOnOne, group, or meeting")
    created_at: Optional[datetime] = Field(default=None, description="When the chat was created")
    last_updated_at: Optional[datetime] = Field(default=None, description="When the chat was last updated")
    web_url: Optional[str] = Field(default=None, description="Deep-link to the chat in Teams")
    tenant_id: Optional[str] = Field(default=None, description="Tenant ID")


class ChatList(BaseModel):
    """Paginated list of chats."""
    chats: List[Chat] = Field(default_factory=list, description="Chat conversations")
    total_chats: int = Field(default=0, description="Total number of chats")


# ── Handler ──────────────────────────────────────────────────────────────


class ChatHandler:
    """Handler for Microsoft Graph Chat API (read-only, async-only).

    Covers personal (1:1) chats, group chats, and meeting chats.
    Requires Chat.Read or Chat.ReadWrite scope.
    """

    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui

    # ── parsing helpers (no I/O) ──────────────────────────────────────────

    @staticmethod
    def _parse_datetime(value: str | None) -> datetime | None:
        if not value:
            return None
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except (ValueError, TypeError):
            return None

    @staticmethod
    def _parse_chat(data: dict) -> Chat:
        return Chat(
            id=data["id"],
            topic=data.get("topic"),
            chat_type=data.get("chatType"),
            created_at=ChatHandler._parse_datetime(data.get("createdDateTime")),
            last_updated_at=ChatHandler._parse_datetime(data.get("lastUpdatedDateTime")),
            web_url=data.get("webUrl"),
            tenant_id=data.get("tenantId"),
        )

    @staticmethod
    def _parse_chat_message(data: dict) -> ChatMessage:
        sender_data = (data.get("from") or {}).get("user") or {}
        sender = ChatMessageFrom(
            display_name=sender_data.get("displayName"),
            email=sender_data.get("userIdentityType"),
            user_id=sender_data.get("id"),
        ) if sender_data else None

        body = data.get("body") or {}
        return ChatMessage(
            id=data["id"],
            created_at=ChatHandler._parse_datetime(data.get("createdDateTime")),
            body_content=body.get("content"),
            body_type=body.get("contentType"),
            sender=sender,
            importance=data.get("importance"),
            message_type=data.get("messageType"),
            web_url=data.get("webUrl"),
        )

    @staticmethod
    def _parse_member(data: dict) -> ChatMember:
        return ChatMember(
            id=data.get("id", ""),
            display_name=data.get("displayName"),
            email=data.get("email"),
            user_id=data.get("userId"),
            roles=data.get("roles", []),
        )

    # ── async API ─────────────────────────────────────────────────────────

    async def get_chats_async(self, limit: int = 50) -> ChatList:
        """List the current user's chats (1:1, group, and meeting)."""
        token = await self.msg.get_access_token_async()
        if not token:
            return ChatList()
        url = f"{self.msg.msg_endpoint}me/chats?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return ChatList()
        data = resp.json()
        chats = [self._parse_chat(c) for c in data.get("value", [])]
        return ChatList(chats=chats, total_chats=len(chats))

    async def get_chat_messages_async(
        self, chat_id: str, limit: int = 20
    ) -> ChatMessageList:
        """List recent messages in a chat."""
        token = await self.msg.get_access_token_async()
        if not token:
            return ChatMessageList()
        url = f"{self.msg.msg_endpoint}me/chats/{chat_id}/messages?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return ChatMessageList()
        data = resp.json()
        messages = [self._parse_chat_message(m) for m in data.get("value", [])]
        return ChatMessageList(messages=messages, total_messages=len(messages))

    async def get_chat_members_async(self, chat_id: str) -> ChatMemberList:
        """List members of a chat."""
        token = await self.msg.get_access_token_async()
        if not token:
            return ChatMemberList()
        url = f"{self.msg.msg_endpoint}me/chats/{chat_id}/members"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return ChatMemberList()
        data = resp.json()
        members = [self._parse_member(m) for m in data.get("value", [])]
        return ChatMemberList(members=members, total_members=len(members))

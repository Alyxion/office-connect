from __future__ import annotations

import asyncio
import logging
from typing import List, Optional, Dict, TYPE_CHECKING

from pydantic import BaseModel, Field

if TYPE_CHECKING:
    from office_con import MsGraphInstance

_log = logging.getLogger(__name__)


# Pydantic models



class DirectoryUser(BaseModel):
    """
    Represents a user from the Microsoft Graph directory.

    Fields:
        id: User UUID
        display_name: User display name
        email: User email
        job_title: Job title (e.g., Sales Manager)
        department: Department (e.g., Sales)
        manager_id: Manager UUID
        account_enabled: Whether the account is enabled
        surname: Surname (e.g., Doe)
        given_name: Given name (e.g., John)
        office_location: Office location (e.g., Office K)
        mobile_phone: Mobile phone (e.g., +49 123 456 789)
    """
    id: str = Field(..., description="User UUID")
    display_name: Optional[str] = Field(default=None, description="User display name")
    email: Optional[str] = Field(default=None, description="User email")
    job_title: Optional[str] = Field(default=None, description="Job title, e.g. Sales Manager")
    department: Optional[str] = Field(default=None, description="Department, e.g. Sales")
    manager_id: Optional[str] = Field(default=None, description="Manager UUID")
    account_enabled: Optional[bool] = Field(default=None, description="Account enabled")
    surname: Optional[str] = Field(default=None, description="Surname, e.g. Doe")
    given_name: Optional[str] = Field(default=None, description="Given name, e.g. John")
    office_location: Optional[str] = Field(default=None, description="Office location, e.g. Office K")
    mobile_phone: Optional[str] = Field(default=None, description="Mobile phone, e.g. +49 123 456 789")


class DirectoryUserList(BaseModel):
    """
    List of directory users.

    Fields:
        users: List of DirectoryUser objects
        total_users: Total number of users
    """
    users: List[DirectoryUser] = Field(default_factory=list, description="List of directory user objects")
    total_users: int = Field(default=0, description="Total number of users in the directory")




class DirectoryHandler:
    """
    Handles Microsoft Graph directory data with async support.

    Provides methods to fetch and manage directory users and their images from Microsoft Graph.

    :param msgraph: Microsoft Graph instance
    """

    # ---- constants ---------------------------------------------------------
    _USERS_JSON = "users.json"
    _PHOTO_DIR = "photos"

    # Graph query tuning
    _SELECT_FIELDS = (
        "id,displayName,mail,userPrincipalName,jobTitle,department,accountEnabled,surname,givenName,officeLocation,mobilePhone"
    )
    _EXPAND_MANAGER = "manager($select=id)"
    _EXPAND_ALL = _EXPAND_MANAGER  # Only manager is a valid $expand; accountEnabled is in $select
    _MAX_PAGE = 100

    # ---------------------------------------------------------------------
    def __init__(self, wui: "MsGraphInstance") -> None:
        self.msg = wui
        self._alock = asyncio.Lock()

    # ── async API ─────────────────────────────────────────────────────────

    async def get_users_async(self, limit: int = 100) -> DirectoryUserList:
        """Return **first page** of users with rich fields."""
        limit = min(limit, 100)
        token = await self.msg.get_access_token_async()
        if not token:
            return DirectoryUserList()

        url = (
            f"{self.msg.msg_endpoint}users"
            f"?$top={limit}"
            f"&$select={self._SELECT_FIELDS}"
            f"&$expand={self._EXPAND_MANAGER}"
        )

        rsp = await self.msg.run_async(url=url, token=token)
        if rsp is None or rsp.status_code != 200:
            return DirectoryUserList()

        users: List[DirectoryUser] = []
        for u in rsp.json().get("value", []):
            mgr = u.get("manager") or {}
            users.append(
                DirectoryUser(
                    id=u["id"],
                    display_name=u.get("displayName"),
                    email=u.get("mail") or u.get("userPrincipalName"),
                    job_title=u.get("jobTitle"),
                    department=u.get("department"),
                    manager_id=mgr.get("id"),
                    account_enabled=u.get("accountEnabled"),
                    surname=u.get("surname"),
                    given_name=u.get("givenName"),
                    office_location=u.get("officeLocation"),
                    mobile_phone=u.get("mobilePhone"),
                )
            )
        return DirectoryUserList(users=users, total_users=len(users))

    async def get_all_users_async(self) -> DirectoryUserList:
        """Return **all** users (paginated)."""
        async with self._alock:
            return await self._fetch_all_users_async()

    async def get_user_manager_async(self, user_id: str) -> Optional[Dict]:
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        url = f"{self.msg.msg_endpoint}users/{user_id}/manager"
        rsp = await self.msg.run_async(url=url, token=token)
        return None if rsp is None or rsp.status_code != 200 else rsp.json()

    async def get_user_photo_async(self, user_id: str) -> Optional[bytes]:
        return await self._fetch_user_photo_async(user_id)

    # ── internal network helpers ──────────────────────────────────────────

    async def _fetch_all_users_async(self) -> DirectoryUserList:
        token = await self.msg.get_access_token_async()
        if not token:
            _log.warning("[DIR] _fetch_all_users_async: no token available")
            return DirectoryUserList()

        url = (
            f"{self.msg.msg_endpoint}users"
            f"?$top={self._MAX_PAGE}"
            f"&$select={self._SELECT_FIELDS}"
            f"&$expand={self._EXPAND_ALL}"
        )
        _log.info("[DIR] _fetch_all_users_async: starting, endpoint=%s", self.msg.msg_endpoint)
        users: List[DirectoryUser] = []

        while url:
            rsp = await self.msg.run_async(url=url, token=token)
            if rsp is None or rsp.status_code != 200:
                status = rsp.status_code if rsp else 'None'
                body = ''
                try:
                    body = rsp.text[:500] if rsp else ''
                except Exception:
                    pass
                _log.error("[DIR] _fetch_all_users_async: API error status=%s body=%s", status, body)
                break
            payload = rsp.json()
            for u in payload.get("value", []):
                mgr = u.get("manager") or {}
                users.append(
                    DirectoryUser(
                        id=u["id"],
                        display_name=u.get("displayName"),
                        email=u.get("mail") or u.get("userPrincipalName"),
                        job_title=u.get("jobTitle"),
                        department=u.get("department"),
                        manager_id=mgr.get("id"),
                        surname=u.get("surname"),
                        given_name=u.get("givenName"),
                        office_location=u.get("officeLocation"),
                        mobile_phone=u.get("mobilePhone"),
                    )
                )
            url = payload.get("@odata.nextLink")

        return DirectoryUserList(users=users, total_users=len(users))

    async def _fetch_user_photo_async(self, user_id: str) -> Optional[bytes]:
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        url = f"{self.msg.msg_endpoint}users/{user_id}/photo/$value"
        rsp = await self.msg.run_async(url=url, token=token)
        return None if rsp is None or rsp.status_code != 200 else rsp.content

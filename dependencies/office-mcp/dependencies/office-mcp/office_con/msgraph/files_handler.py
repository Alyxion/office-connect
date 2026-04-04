"""Microsoft Graph Files & SharePoint handler (async-only).

Provides read access to:
- OneDrive drives (personal and shared)
- Drive items (files and folders)
- SharePoint sites
- File content download
"""

from __future__ import annotations

from datetime import datetime
from typing import List, Optional, TYPE_CHECKING

from pydantic import BaseModel, Field

if TYPE_CHECKING:
    from office_con import MsGraphInstance


# ── Pydantic models ──────────────────────────────────────────────────────


class DriveItemUser(BaseModel):
    """User reference in a drive item (created/modified by)."""
    display_name: Optional[str] = Field(default=None, description="Display name")
    email: Optional[str] = Field(default=None, description="Email address")
    user_id: Optional[str] = Field(default=None, description="Azure AD user ID")


class DriveItem(BaseModel):
    """A file or folder in OneDrive or SharePoint.

    For folders, ``folder_child_count`` is set and ``size`` refers to total folder size.
    For files, ``mime_type`` is set.
    """
    id: str = Field(..., description="Drive item ID")
    name: Optional[str] = Field(default=None, description="File or folder name")
    size: Optional[int] = Field(default=None, description="Size in bytes")
    web_url: Optional[str] = Field(default=None, description="Browser URL")
    created_at: Optional[datetime] = Field(default=None, description="Creation timestamp (UTC)")
    modified_at: Optional[datetime] = Field(default=None, description="Last modified timestamp (UTC)")
    created_by: Optional[DriveItemUser] = Field(default=None, description="Who created the item")
    modified_by: Optional[DriveItemUser] = Field(default=None, description="Who last modified the item")
    mime_type: Optional[str] = Field(default=None, description="MIME type (files only)")
    is_folder: bool = Field(default=False, description="True if item is a folder")
    folder_child_count: Optional[int] = Field(default=None, description="Number of children (folders only)")
    download_url: Optional[str] = Field(default=None, description="Pre-authenticated download URL (short-lived)")


class DriveItemList(BaseModel):
    """List of drive items."""
    items: List[DriveItem] = Field(default_factory=list, description="Files and folders")
    total_items: int = Field(default=0, description="Total number of items")


class Drive(BaseModel):
    """A OneDrive or SharePoint document library."""
    id: str = Field(..., description="Drive ID")
    name: Optional[str] = Field(default=None, description="Drive name")
    drive_type: Optional[str] = Field(default=None, description="personal, business, or documentLibrary")
    owner_name: Optional[str] = Field(default=None, description="Owner display name")
    quota_total: Optional[int] = Field(default=None, description="Total quota in bytes")
    quota_used: Optional[int] = Field(default=None, description="Used quota in bytes")
    web_url: Optional[str] = Field(default=None, description="Browser URL to the drive")


class DriveList(BaseModel):
    """List of drives."""
    drives: List[Drive] = Field(default_factory=list, description="OneDrive and SharePoint drives")
    total_drives: int = Field(default=0, description="Total number of drives")


class SharePointSite(BaseModel):
    """A SharePoint site."""
    id: str = Field(..., description="Site ID (host,site-collection-id,web-id)")
    display_name: Optional[str] = Field(default=None, description="Site display name")
    name: Optional[str] = Field(default=None, description="Site URL name")
    web_url: Optional[str] = Field(default=None, description="Full URL to the site")
    description: Optional[str] = Field(default=None, description="Site description")
    created_at: Optional[datetime] = Field(default=None, description="When the site was created")


class SharePointSiteList(BaseModel):
    """List of SharePoint sites."""
    sites: List[SharePointSite] = Field(default_factory=list, description="SharePoint sites")
    total_sites: int = Field(default=0, description="Total number of sites")


# ── Handler ──────────────────────────────────────────────────────────────


class FilesHandler:
    """Handler for Microsoft Graph Files & SharePoint API (read-only, async-only).

    Covers OneDrive files, SharePoint document libraries, and site discovery.
    Requires Files.Read.All (or Files.ReadWrite.All) and Sites.Read.All scopes.
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
    def _parse_user_ref(data: dict | None) -> DriveItemUser | None:
        if not data:
            return None
        user = data.get("user") or {}
        if not user:
            return None
        return DriveItemUser(
            display_name=user.get("displayName"),
            email=user.get("email"),
            user_id=user.get("id"),
        )

    @staticmethod
    def _parse_drive(data: dict) -> Drive:
        owner = (data.get("owner") or {}).get("user") or {}
        quota = data.get("quota") or {}
        return Drive(
            id=data["id"],
            name=data.get("name"),
            drive_type=data.get("driveType"),
            owner_name=owner.get("displayName"),
            quota_total=quota.get("total"),
            quota_used=quota.get("used"),
            web_url=data.get("webUrl"),
        )

    @staticmethod
    def _parse_drive_item(data: dict) -> DriveItem:
        folder = data.get("folder")
        file_info = data.get("file") or {}
        return DriveItem(
            id=data["id"],
            name=data.get("name"),
            size=data.get("size"),
            web_url=data.get("webUrl"),
            created_at=FilesHandler._parse_datetime(data.get("createdDateTime")),
            modified_at=FilesHandler._parse_datetime(data.get("lastModifiedDateTime")),
            created_by=FilesHandler._parse_user_ref(data.get("createdBy")),
            modified_by=FilesHandler._parse_user_ref(data.get("lastModifiedBy")),
            mime_type=file_info.get("mimeType"),
            is_folder=folder is not None,
            folder_child_count=folder.get("childCount") if folder else None,
            download_url=data.get("@microsoft.graph.downloadUrl"),
        )

    @staticmethod
    def _parse_site(data: dict) -> SharePointSite:
        return SharePointSite(
            id=data["id"],
            display_name=data.get("displayName"),
            name=data.get("name"),
            web_url=data.get("webUrl"),
            description=data.get("description"),
            created_at=FilesHandler._parse_datetime(data.get("createdDateTime")),
        )

    # ── async API — Drives ────────────────────────────────────────────────

    async def get_my_drives_async(self) -> DriveList:
        """List drives accessible to the current user (OneDrive + shared)."""
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveList()
        url = f"{self.msg.msg_endpoint}me/drives"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return DriveList()
        data = resp.json()
        drives = [self._parse_drive(d) for d in data.get("value", [])]
        return DriveList(drives=drives, total_drives=len(drives))

    async def get_my_drive_async(self) -> Drive | None:
        """Get the current user's default OneDrive."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        url = f"{self.msg.msg_endpoint}me/drive"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return self._parse_drive(resp.json())

    # ── async API — Drive items ───────────────────────────────────────────

    async def get_root_items_async(self, drive_id: str | None = None, limit: int = 50) -> DriveItemList:
        """List items in the root of a drive.

        If ``drive_id`` is None, uses the current user's default OneDrive.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveItemList()
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/root/children?$top={limit}"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/root/children?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return DriveItemList()
        data = resp.json()
        items = [self._parse_drive_item(i) for i in data.get("value", [])]
        return DriveItemList(items=items, total_items=len(items))

    async def get_folder_items_async(
        self, item_id: str, drive_id: str | None = None, limit: int = 50
    ) -> DriveItemList:
        """List children of a folder.

        If ``drive_id`` is None, uses the current user's default OneDrive.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveItemList()
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}/children?$top={limit}"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}/children?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return DriveItemList()
        data = resp.json()
        items = [self._parse_drive_item(i) for i in data.get("value", [])]
        return DriveItemList(items=items, total_items=len(items))

    async def get_item_async(self, item_id: str, drive_id: str | None = None) -> DriveItem | None:
        """Get metadata for a single drive item."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return self._parse_drive_item(resp.json())

    async def get_file_content_async(
        self, item_id: str, drive_id: str | None = None
    ) -> bytes | None:
        """Download the content of a file as bytes.

        If ``drive_id`` is None, uses the current user's default OneDrive.
        Returns None if the item is not found or is a folder.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}/content"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}/content"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return resp.content

    async def search_items_async(
        self, query: str, drive_id: str | None = None, limit: int = 25
    ) -> DriveItemList:
        """Search for files and folders by name or content.

        If ``drive_id`` is None, searches the current user's default OneDrive.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveItemList()
        safe_query = query.replace("'", "''")
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/root/search(q='{safe_query}')?$top={limit}"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/root/search(q='{safe_query}')?$top={limit}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return DriveItemList()
        data = resp.json()
        items = [self._parse_drive_item(i) for i in data.get("value", [])]
        return DriveItemList(items=items, total_items=len(items))

    # ── async API — SharePoint sites ──────────────────────────────────────

    async def get_followed_sites_async(self) -> SharePointSiteList:
        """List SharePoint sites the current user follows."""
        token = await self.msg.get_access_token_async()
        if not token:
            return SharePointSiteList()
        url = f"{self.msg.msg_endpoint}me/followedSites"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return SharePointSiteList()
        data = resp.json()
        sites = [self._parse_site(s) for s in data.get("value", [])]
        return SharePointSiteList(sites=sites, total_sites=len(sites))

    async def search_sites_async(self, query: str) -> SharePointSiteList:
        """Search for SharePoint sites by keyword."""
        token = await self.msg.get_access_token_async()
        if not token:
            return SharePointSiteList()
        from urllib.parse import quote
        url = f"{self.msg.msg_endpoint}sites?search={quote(query)}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return SharePointSiteList()
        data = resp.json()
        sites = [self._parse_site(s) for s in data.get("value", [])]
        return SharePointSiteList(sites=sites, total_sites=len(sites))

    async def get_site_drives_async(self, site_id: str) -> DriveList:
        """List document libraries (drives) in a SharePoint site."""
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveList()
        url = f"{self.msg.msg_endpoint}sites/{site_id}/drives"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return DriveList()
        data = resp.json()
        drives = [self._parse_drive(d) for d in data.get("value", [])]
        return DriveList(drives=drives, total_drives=len(drives))

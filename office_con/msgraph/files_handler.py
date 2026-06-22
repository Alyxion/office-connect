"""Microsoft Graph Files & SharePoint handler (async-only).

Provides file access to:
- OneDrive drives (personal and shared)
- Drive items (files and folders)
- SharePoint sites
- File content download
- Explicit file/folder writes for whitelisted callers; no delete helper
"""

from __future__ import annotations

import logging
from datetime import datetime
from typing import List, Optional, TYPE_CHECKING
from urllib.parse import quote

from pydantic import BaseModel, Field

from office_con.privacy import (
    OfficeContentBlockedError,
    any_text_matches,
    decode_text_for_rules,
    folder_path_for_match,
)

logger = logging.getLogger(__name__)

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
    e_tag: Optional[str] = Field(default=None, description="Graph eTag; changes when the item resource changes")
    c_tag: Optional[str] = Field(default=None, description="Graph cTag; changes when file content changes")
    created_at: Optional[datetime] = Field(default=None, description="Creation timestamp (UTC)")
    modified_at: Optional[datetime] = Field(default=None, description="Last modified timestamp (UTC)")
    created_by: Optional[DriveItemUser] = Field(default=None, description="Who created the item")
    modified_by: Optional[DriveItemUser] = Field(default=None, description="Who last modified the item")
    mime_type: Optional[str] = Field(default=None, description="MIME type (files only)")
    is_folder: bool = Field(default=False, description="True if item is a folder")
    folder_child_count: Optional[int] = Field(default=None, description="Number of children (folders only)")
    download_url: Optional[str] = Field(default=None, description="Pre-authenticated download URL (short-lived)")
    drive_id: Optional[str] = Field(default=None, description="ID of the drive that holds this item (needed to fetch/download items found via tenant-wide search, since they may live in SharePoint libraries rather than the user's OneDrive)")
    parent_path: Optional[str] = Field(default=None, description="Path of the parent folder (from parentReference)")
    parent_id: Optional[str] = Field(default=None, description="ID of the parent folder (from parentReference)")
    access_status: str = Field(default="ok", description="ok, content_blocked, or folder_blocked")
    access_reason: str = Field(default="", description="Human-readable privacy/access reason")

    @property
    def change_signature(self) -> str:
        """Cheap provider-side token for detecting item/content changes."""
        tokens = [self.c_tag or "", self.e_tag or ""]
        token_signature = "|".join(token for token in tokens if token)
        if token_signature:
            return token_signature
        modified = self.modified_at.isoformat() if self.modified_at is not None else ""
        size = self.size if self.size is not None else 0
        return f"{modified}|{size}"


class DriveItemList(BaseModel):
    """List of drive items."""
    items: List[DriveItem] = Field(default_factory=list, description="Files and folders")
    total_items: int = Field(default=0, description="Total number of items")


class DriveItemVersion(BaseModel):
    """A metadata-only file version entry."""
    id: str = Field(..., description="Version ID")
    modified_at: Optional[datetime] = Field(default=None, description="Version modified timestamp (UTC)")
    size: Optional[int] = Field(default=None, description="Version size in bytes")

    @property
    def change_signature(self) -> str:
        modified = self.modified_at.isoformat() if self.modified_at is not None else ""
        size = self.size if self.size is not None else 0
        return f"{self.id}|{modified}|{size}"


class DriveItemVersionList(BaseModel):
    """List of file versions."""
    versions: List[DriveItemVersion] = Field(default_factory=list, description="File versions")
    total_versions: int = Field(default=0, description="Total number of returned versions")


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
    """Handler for Microsoft Graph Files & SharePoint API (async-only).

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
        parent = data.get("parentReference") or {}
        return DriveItem(
            id=data["id"],
            name=data.get("name"),
            size=data.get("size"),
            web_url=data.get("webUrl"),
            e_tag=data.get("eTag"),
            c_tag=data.get("cTag"),
            created_at=FilesHandler._parse_datetime(data.get("createdDateTime")),
            modified_at=FilesHandler._parse_datetime(data.get("lastModifiedDateTime")),
            created_by=FilesHandler._parse_user_ref(data.get("createdBy")),
            modified_by=FilesHandler._parse_user_ref(data.get("lastModifiedBy")),
            mime_type=file_info.get("mimeType"),
            is_folder=folder is not None,
            folder_child_count=folder.get("childCount") if folder else None,
            download_url=data.get("@microsoft.graph.downloadUrl"),
            drive_id=parent.get("driveId"),
            parent_path=parent.get("path"),
            parent_id=parent.get("id"),
        )

    def _privacy_enabled(self) -> bool:
        return self.msg.privacy_settings.files.enabled

    def _folder_block_reason(self, item: DriveItem) -> str:
        if not self._privacy_enabled():
            return ""
        item_drive = item.drive_id or ""
        item_path = folder_path_for_match(item.parent_path, item.name)
        for blocked in self.msg.privacy_settings.files.blocked_folders:
            blocked_drive = blocked.drive_id or ""
            if blocked_drive and item_drive and blocked_drive != item_drive:
                continue
            if item.is_folder and blocked.item_id and item.id == blocked.item_id:
                return "Folder is blocked by your privacy settings."
            if blocked.item_id and item.parent_id == blocked.item_id:
                return "Parent folder is blocked by your privacy settings."
            blocked_path = folder_path_for_match(blocked.parent_path, blocked.name)
            if blocked_path and item_path and item_path.startswith(blocked_path.rstrip("/") + "/"):
                return "Parent folder is blocked by your privacy settings."
        return ""

    def _whitelisted_by_name(self, item: DriveItem) -> bool:
        """True when an allowed term matches the item's name/url/path — the
        keyword hide/block rules are then skipped. Explicit folder/item blocks
        are not affected."""
        cfg = self.msg.privacy_settings.files
        if not cfg.allowed_terms:
            return False
        return any_text_matches([item.name, item.web_url, item.parent_path], cfg.allowed_terms)

    def _hidden_by_metadata(self, item: DriveItem) -> bool:
        if not self._privacy_enabled():
            return False
        if self._whitelisted_by_name(item):
            return False
        cfg = self.msg.privacy_settings.files
        if any_text_matches([item.name, item.web_url, item.parent_path], cfg.hidden_name_terms):
            return True
        return False

    def _item_is_blocked(self, item: DriveItem) -> bool:
        if not item.id:
            return False
        item_drive = item.drive_id or ""
        for blocked in self.msg.privacy_settings.files.blocked_items:
            if blocked.item_id != item.id:
                continue
            blocked_drive = blocked.drive_id or ""
            if blocked_drive and item_drive and blocked_drive != item_drive:
                continue
            return True
        return False

    def _apply_metadata_privacy(self, item: DriveItem, *, enforce_folders: bool = True) -> DriveItem | None:
        if self._hidden_by_metadata(item):
            return None
        if enforce_folders:
            folder_reason = self._folder_block_reason(item)
            if folder_reason:
                item.access_status = "folder_blocked"
                item.access_reason = folder_reason
                return item
        if not self._privacy_enabled():
            return item
        cfg = self.msg.privacy_settings.files
        if self._item_is_blocked(item):
            item.access_status = "content_blocked"
            item.access_reason = "This file is blocked by your privacy settings."
            return item
        if self._whitelisted_by_name(item):
            return item
        if any_text_matches([item.name, item.web_url, item.parent_path], cfg.blocked_content_name_terms):
            item.access_status = "content_blocked"
            item.access_reason = "Content access is blocked by your privacy settings."
        return item

    def _filter_items(self, items: list[DriveItem], *, enforce_folders: bool = True) -> list[DriveItem]:
        filtered: list[DriveItem] = []
        for item in items:
            visible = self._apply_metadata_privacy(item, enforce_folders=enforce_folders)
            if visible is not None:
                filtered.append(visible)
        return filtered

    def _content_block_reason(
        self, item: DriveItem, content: bytes | None = None, *, enforce_folders: bool = True
    ) -> str:
        metadata_item = self._apply_metadata_privacy(item, enforce_folders=enforce_folders)
        if metadata_item is None:
            return "hidden"
        if metadata_item.access_status == "folder_blocked":
            return metadata_item.access_reason or "Folder is blocked by your privacy settings."
        if metadata_item.access_status == "content_blocked":
            return metadata_item.access_reason or "Content access is blocked by your privacy settings."
        if not self._privacy_enabled():
            return ""
        cfg = self.msg.privacy_settings.files
        if content is not None:
            text = decode_text_for_rules(content)
            whitelisted = self._whitelisted_by_name(item) or bool(
                text and any_text_matches([text], cfg.allowed_terms)
            )
            if not whitelisted:
                if text and any_text_matches([text], cfg.hidden_content_terms):
                    return "hidden"
                if text and any_text_matches([text], cfg.blocked_content_terms):
                    return "Content access is blocked by your privacy settings."
        return ""

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

    @staticmethod
    def _parse_drive_item_version(data: dict) -> DriveItemVersion:
        return DriveItemVersion(
            id=str(data.get("id") or ""),
            modified_at=FilesHandler._parse_datetime(data.get("lastModifiedDateTime")),
            size=data.get("size"),
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
        items = self._filter_items([self._parse_drive_item(i) for i in data.get("value", [])])
        return DriveItemList(items=items, total_items=len(items))

    async def get_folder_items_async(
        self, item_id: str, drive_id: str | None = None, limit: int = 50,
        *, enforce_folder_block: bool = True,
    ) -> DriveItemList:
        """List children of a folder.

        If ``drive_id`` is None, uses the current user's default OneDrive.

        ``enforce_folder_block`` is True by default (used by the AI/MCP path);
        a host UI may pass False to let the owner browse *into* a blocked folder
        (the children are not folder-blocked, but hide and content/name rules
        still apply). This is an explicit human-authorized override only.
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
        items = self._filter_items(
            [self._parse_drive_item(i) for i in data.get("value", [])],
            enforce_folders=enforce_folder_block,
        )
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
        item = self._parse_drive_item(resp.json())
        return self._apply_metadata_privacy(item)

    async def get_file_content_async(
        self, item_id: str, drive_id: str | None = None, *, allow_blocked: bool = False
    ) -> bytes | None:
        """Download the content of a file as bytes.

        If ``drive_id`` is None, uses the current user's default OneDrive.
        Returns None if the item is not found or is a folder.

        ``allow_blocked`` is False by default (the AI/MCP path always enforces
        privacy). A host UI may pass True ONLY for a file the owner explicitly
        and deliberately selected from a blocked folder — this bypasses the
        folder block and content/name block for that single file. Hidden files
        are never returned regardless.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        item = await self.get_item_async(item_id, drive_id=drive_id)
        if item is None or item.is_folder:
            return None
        if item.access_status == "folder_blocked" and not allow_blocked:
            return None
        if item.access_status == "content_blocked" and not allow_blocked:
            raise OfficeContentBlockedError(item.access_reason or "Content access is blocked by your privacy settings.")
        if allow_blocked:
            # Normalize so the post-download content check recomputes cleanly.
            item.access_status = "ok"
            item.access_reason = ""
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}/content"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}/content"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        reason = self._content_block_reason(item, resp.content, enforce_folders=not allow_blocked)
        if reason == "hidden":
            return None
        if reason and not allow_blocked:
            raise OfficeContentBlockedError(reason)
        return resp.content

    async def get_file_versions_async(
        self,
        item_id: str,
        drive_id: str | None = None,
        limit: int = 1,
    ) -> DriveItemVersionList:
        """List metadata for recent versions of a file."""
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveItemVersionList()
        top = max(1, min(int(limit), 20))
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}/versions?$top={top}"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}/versions?$top={top}"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return DriveItemVersionList()
        data = resp.json()
        versions = [self._parse_drive_item_version(v) for v in data.get("value", [])]
        return DriveItemVersionList(versions=versions, total_versions=len(versions))

    async def put_file_content_async(
        self,
        item_id: str,
        content: bytes,
        drive_id: str | None = None,
        content_type: str = "application/octet-stream",
    ) -> DriveItem | None:
        """Overwrite one existing drive item with new bytes."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}/content"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}/content"
        resp = await self.msg.run_async(
            url=url,
            method="PUT",
            data=content,
            token=token,
            add_headers={"Content-Type": content_type or "application/octet-stream"},
        )
        if resp is None or resp.status_code not in (200, 201):
            if resp is None:
                logger.warning("[GRAPH_FILES] PUT content failed for item %s: no response", item_id)
            else:
                logger.warning(
                    "[GRAPH_FILES] PUT content failed for item %s: status=%s body=%s",
                    item_id,
                    resp.status_code,
                    (resp.text or "")[:500],
                )
            return None
        return self._parse_drive_item(resp.json())

    async def upload_file_to_folder_async(
        self,
        folder_id: str,
        filename: str,
        content: bytes,
        drive_id: str | None = None,
        content_type: str = "application/octet-stream",
    ) -> DriveItem | None:
        """Create a file directly below a folder; fail if the name exists."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        safe_name = quote(filename, safe="")
        if drive_id:
            url = (
                f"{self.msg.msg_endpoint}drives/{drive_id}/items/{folder_id}:/{safe_name}:/content"
                "?@microsoft.graph.conflictBehavior=fail"
            )
        else:
            url = (
                f"{self.msg.msg_endpoint}me/drive/items/{folder_id}:/{safe_name}:/content"
                "?@microsoft.graph.conflictBehavior=fail"
            )
        resp = await self.msg.run_async(
            url=url,
            method="PUT",
            data=content,
            token=token,
            add_headers={"Content-Type": content_type or "application/octet-stream"},
        )
        if resp is None or resp.status_code not in (200, 201):
            if resp is None:
                logger.warning("[GRAPH_FILES] Upload content failed for folder %s/%s: no response", folder_id, filename)
            else:
                logger.warning(
                    "[GRAPH_FILES] Upload content failed for folder %s/%s: status=%s body=%s",
                    folder_id,
                    filename,
                    resp.status_code,
                    (resp.text or "")[:500],
                )
            return None
        return self._parse_drive_item(resp.json())

    async def create_folder_async(
        self,
        parent_id: str,
        name: str,
        drive_id: str | None = None,
    ) -> DriveItem | None:
        """Create a child folder. Existing folders are not overwritten."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{parent_id}/children"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{parent_id}/children"
        body = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }
        resp = await self.msg.run_async(url=url, method="POST", json=body, token=token)
        if resp is None or resp.status_code not in (200, 201):
            return None
        return self._parse_drive_item(resp.json())

    async def rename_item_async(
        self,
        item_id: str,
        name: str,
        drive_id: str | None = None,
    ) -> DriveItem | None:
        """Rename a drive item. Does not move or delete the item."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        if drive_id:
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/items/{item_id}"
        else:
            url = f"{self.msg.msg_endpoint}me/drive/items/{item_id}"
        resp = await self.msg.run_async(url=url, method="PATCH", json={"name": name}, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return self._parse_drive_item(resp.json())

    async def search_items_async(
        self, query: str, drive_id: str | None = None, limit: int = 25
    ) -> DriveItemList:
        """Search for files and folders by name or content.

        When ``drive_id`` is given, the search is scoped to that single drive
        via the per-drive ``search(q=…)`` endpoint.

        When ``drive_id`` is None (the default), the search is **tenant-wide**:
        it uses the unified Microsoft Search API (``POST /search/query`` with
        ``entityTypes: ["driveItem"]``), which spans the user's OneDrive *and*
        every SharePoint document library they have access to — the same corpus
        the OneDrive/SharePoint web UI searches. The old behaviour scoped this
        to ``/me/drive`` only, so files living in SharePoint sites (and content
        the user could otherwise open in the browser) were never found, and an
        unmatched term surfaced loosely-ranked personal files instead of an
        empty result.
        """
        token = await self.msg.get_access_token_async()
        if not token:
            return DriveItemList()

        if drive_id:
            safe_query = query.replace("'", "''")
            url = f"{self.msg.msg_endpoint}drives/{drive_id}/root/search(q='{safe_query}')?$top={limit}"
            resp = await self.msg.run_async(url=url, token=token)
            if resp is None or resp.status_code != 200:
                return DriveItemList()
            data = resp.json()
            items = self._filter_items([self._parse_drive_item(i) for i in data.get("value", [])])
            return DriveItemList(items=items, total_items=len(items))

        # Tenant-wide unified search (OneDrive + SharePoint libraries).
        body = {
            "requests": [{
                "entityTypes": ["driveItem"],
                "query": {"queryString": query},
                "from": 0,
                "size": limit,
            }]
        }
        resp = await self.msg.run_async(
            url=f"{self.msg.msg_endpoint}search/query",
            method="POST", json=body, token=token,
        )
        if resp is None or resp.status_code != 200:
            # Compatibility fallback: some tenants/apps do not have Microsoft
            # Search enabled or consented. Preserve the old personal-OneDrive
            # search behavior instead of returning a misleading empty result.
            safe_query = query.replace("'", "''")
            url = f"{self.msg.msg_endpoint}me/drive/root/search(q='{safe_query}')?$top={limit}"
            fallback = await self.msg.run_async(url=url, token=token)
            if fallback is None or fallback.status_code != 200:
                return DriveItemList()
            data = fallback.json()
            items = self._filter_items([self._parse_drive_item(i) for i in data.get("value", [])])
            return DriveItemList(items=items, total_items=len(items))
        items: List[DriveItem] = []
        total = 0
        for container in resp.json().get("value", []):
            for hits_container in container.get("hitsContainers", []):
                total = hits_container.get("total", total)
                for hit in hits_container.get("hits", []):
                    resource = hit.get("resource")
                    if resource:
                        item = self._apply_metadata_privacy(self._parse_drive_item(resource))
                        if item is not None:
                            items.append(item)
        # Fall back to the parsed-hit count when the API omits ``total``.
        return DriveItemList(items=items, total_items=total or len(items))

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

    async def get_team_channel_files_folder_async(
        self, team_id: str, channel_id: str
    ) -> DriveItem | None:
        """Get the SharePoint-backed files folder for a Teams channel."""
        token = await self.msg.get_access_token_async()
        if not token:
            return None
        url = f"{self.msg.msg_endpoint}teams/{team_id}/channels/{channel_id}/filesFolder"
        resp = await self.msg.run_async(url=url, token=token)
        if resp is None or resp.status_code != 200:
            return None
        return self._parse_drive_item(resp.json())

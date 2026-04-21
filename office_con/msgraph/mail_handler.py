import base64
import hashlib
import json
import logging
from datetime import datetime, timezone
from typing import List, Optional, Dict

from pydantic import BaseModel, Field

from typing import TYPE_CHECKING, Literal

if TYPE_CHECKING:
    from office_con import MsGraphInstance

_log = logging.getLogger(__name__)


class OfficeMailAttachment(BaseModel):
    """A file or inline image attached to an Outlook email."""
    name: str = Field(description="Attachment filename")
    content_type: str = Field(description="MIME type, e.g. application/pdf")
    content_bytes: bytes | None = Field(default=None, description="Raw attachment bytes")
    content_id: Optional[str] = Field(default=None, description="Content-ID for embedded images (cid: references)")
    is_embedded: bool = Field(default=False, description="Whether the attachment is embedded inline in the email body")


class OfficeMailCategory(BaseModel):
    """An Outlook mail category (colour label)."""
    id: str = Field(description="Category identifier")
    name: str = Field(description="Display name, e.g. 'Red category'")
    preset_color: str = Field(description="MS Graph preset colour name, e.g. 'preset0'")
    color: str = Field(description="Resolved HTML colour, e.g. '#e74856'")


class OfficeMail(BaseModel):
    """A single Outlook email message with metadata and body."""
    email_id: str = Field(description="MS Graph message id")
    email_url: Optional[str] = Field(default=None, description="Full MS Graph URL for this message")
    flag_state: Literal["flagged", "notFlagged", "done"] = Field(default="notFlagged", description="Follow-up flag state")
    importance: str | None = Field(default="normal", description="Importance level: low, normal, high")
    is_read: bool = Field(default=False, description="Whether the message has been read")
    email_type: str = Field(description="Type of email, e.g. 'inbox'")
    local_timestamp: str | None = Field(default=None, description="Received time in local timezone as string")
    from_name: str | None = Field(default=None, description="Sender display name")
    from_email: str | None = Field(default=None, description="Sender email address")
    subject: str | None = Field(default=None, description="Email subject line")
    body_preview: str | None = Field(default=None, description="Short plain-text preview of the body")
    body: str | None = Field(default=None, description="Full email body content")
    body_type: str | None = Field(default=None, description="Body content type: 'html' or 'text'")
    has_attachments: bool = Field(default=False, description="Whether the message has attachments")
    web_link: Optional[str] = Field(default=None, description="Outlook Web App URL to open this message")
    categories: List[str] = Field(default_factory=list, description="Assigned category labels")
    confidential_level: Optional[str] = Field(default=None, description="Sensitivity: normal, personal, private, confidential")
    attachments: List[OfficeMailAttachment] = Field(default_factory=list, description="File and inline attachments")
    zip_data: Optional[bytes] = Field(default=None, description="Compressed attachment bundle for transport")

    @property
    def scanning(self) -> bool:
        """Virus-scan detection: True when Graph indicates attachments exist
        but none are available yet (scan in progress), or when a placeholder
        'virus scan in progress.html' attachment is present.

        Only meaningful after a full message fetch (``$expand=attachments``).
        Index queries never populate ``attachments``, so callers listing
        messages should not use this property — see ``_mail_to_row``.
        """
        if not self.has_attachments:
            return False
        if any(a.name.lower() == "virus scan in progress.html" for a in self.attachments):
            return True
        return len(self.attachments) == 0


class OfficeMailList(BaseModel):
    """Paginated list of Outlook email messages."""
    elements: List[OfficeMail] = Field(default_factory=list, description="Email messages in this page")
    total_mails: int = Field(default=0, description="Total number of mails in the folder")


class FolderInfo(BaseModel):
    """A mail folder with counts."""
    id: str = Field(description="MS Graph folder id")
    name: str = Field(default="", description="Display name")
    unread: int = Field(default=0, description="Unread message count")
    total: int = Field(default=0, description="Total message count")
    parent_id: str | None = Field(default=None, description="Parent folder id for tree rendering")


class MoveResult(BaseModel):
    """Result of moving a message to another folder."""
    id: str = Field(description="New message id in the destination folder")
    web_link: str = Field(default="", description="Outlook Web App URL")


def compute_folder_signature(rows: list[dict]) -> str:
    """Cache-busting signature for a folder's message list.

    Rows are sorted by ``id`` before hashing so that Graph API
    pagination-order jitter does not produce false cache misses.

    Returns 16 hex characters of SHA-256.
    """
    h = hashlib.sha256()
    for row in sorted(rows, key=lambda r: r["id"]):
        h.update(json.dumps(row, sort_keys=True, default=str).encode())
    return h.hexdigest()[:16]


class MailFolderHandler:
    """Reads and manages Outlook mail folders via the MS Graph API."""

    def __init__(self, wui: "MsGraphInstance"):
        self.msg = wui

    def _base_path(self, mail_address: str | None = None) -> str:
        if not mail_address or mail_address == self.msg.email:
            return "me"
        return f"users/{mail_address}"

    async def get_folders_async(
        self, *, include_hidden: bool = True, limit: int = 100,
        recursive: bool = False, max_depth: int = 6,
        mail_address: str | None = None,
    ) -> list[FolderInfo]:
        """Return a flat list of mail folders with counts + ``parent_id``.

        When ``recursive=True`` we walk ``childFolders`` breadth-first so
        nested folders (e.g. ``Inbox/News``) appear in the result.  BFS
        is bounded by ``max_depth`` (6 by default) so a pathological
        structure can't hang the handler.  ``childFolderCount`` in the
        response tells us whether a BFS step is worth taking.
        """
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return []
        base = self._base_path(mail_address)
        root_url = (
            f"{self.msg.msg_endpoint}{base}/mailFolders"
            f"?$select=id,displayName,totalItemCount,unreadItemCount,parentFolderId,childFolderCount"
            f"&$top={limit}"
        )
        if include_hidden:
            root_url += "&$includeHiddenFolders=true"
        resp = await self.msg.run_async(url=root_url, token=access_token)
        if resp is None or resp.status_code != 200:
            return []
        all_rows: list[dict] = list(resp.json().get("value", []))

        if recursive:
            # BFS through ``childFolders`` for every folder that reports
            # at least one child.  Each level costs one request per
            # parent with children — for typical mailboxes (≤50 folders
            # total) this is a handful of calls.
            queue = [r for r in all_rows if (r.get("childFolderCount") or 0) > 0]
            depth = 0
            seen_ids = {r["id"] for r in all_rows}
            while queue and depth < max_depth:
                next_queue: list[dict] = []
                for parent in queue:
                    child_url = (
                        f"{self.msg.msg_endpoint}{base}/mailFolders/{parent['id']}/childFolders"
                        f"?$select=id,displayName,totalItemCount,unreadItemCount,parentFolderId,childFolderCount"
                        f"&$top={limit}"
                    )
                    if include_hidden:
                        child_url += "&$includeHiddenFolders=true"
                    r = await self.msg.run_async(url=child_url, token=access_token)
                    if r is None or r.status_code != 200:
                        continue
                    for child in r.json().get("value", []):
                        if child["id"] in seen_ids:
                            continue
                        seen_ids.add(child["id"])
                        all_rows.append(child)
                        if (child.get("childFolderCount") or 0) > 0:
                            next_queue.append(child)
                queue = next_queue
                depth += 1

        return [
            FolderInfo(
                id=f["id"],
                name=f.get("displayName", ""),
                unread=f.get("unreadItemCount", 0),
                total=f.get("totalItemCount", 0),
                parent_id=f.get("parentFolderId"),
            )
            for f in all_rows
        ]

    async def get_folder_async(
        self, folder_id: str, *, mail_address: str | None = None,
    ) -> FolderInfo | None:
        """Return a single mail folder by ID, or None if not found."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None
        base = self._base_path(mail_address)
        url = (
            f"{self.msg.msg_endpoint}{base}/mailFolders/{folder_id}"
            f"?$select=id,displayName,totalItemCount,unreadItemCount,parentFolderId"
        )
        resp = await self.msg.run_async(url=url, token=access_token)
        if resp is None or resp.status_code != 200:
            return None
        f = resp.json()
        return FolderInfo(
            id=f["id"],
            name=f.get("displayName", ""),
            unread=f.get("unreadItemCount", 0),
            total=f.get("totalItemCount", 0),
            parent_id=f.get("parentFolderId"),
        )


class OfficeMailHandler:
    """Reads, sends, and manages Outlook emails via the MS Graph API."""

    def __init__(self, wui: "MsGraphInstance"):
        self.msg = wui

    # ── parsing helpers (no I/O) ──────────────────────────────────────────

    def parse_mail(self, email: dict[str, object]) -> OfficeMail:
        mail_address = email.get('from', {}).get('emailAddress', {})
        time_stamp = email.get('receivedDateTime', None)
        if time_stamp:
            utc_time = datetime.strptime(time_stamp, "%Y-%m-%dT%H:%M:%SZ").replace(tzinfo=timezone.utc)
            local_time = utc_time.astimezone()
            local_time_str = local_time.strftime("%Y-%m-%d %H:%M:%S")
        else:
            local_time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        attachments = []
        for attachment in email.get('attachments', []):
            attachment_name = attachment['name']
            if 'contentBytes' not in attachment:
                continue
            attachment_content = attachment['contentBytes']
            attachment_type = attachment['contentType']
            content_id = attachment.get('contentId', None)
            content_bytes = base64.b64decode(attachment_content)
            new_attachment = OfficeMailAttachment(
                name=attachment_name,
                content_type=attachment_type,
                content_bytes=content_bytes,
                content_id=content_id
            )
            attachments.append(new_attachment)

        body = email.get('body', {}).get('content', "")
        if body:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(body, 'html.parser')
            img_tags = soup.find_all('img')
            for img in img_tags:
                src = str(img.get('src', ''))
                if not src:
                    continue
                if 'base64' in src:
                    continue
                if src.startswith('cid:'):
                    cid = src[4:]
                    for attachment in attachments:
                        if attachment.content_id == cid or attachment.content_id == cid.split('@')[0]:
                            attachment.is_embedded = True
                            break
                    else:
                        img_name = cid.split('@')[0]
                        for attachment in attachments:
                            if attachment.name == img_name:
                                attachment.is_embedded = True
                                break
                else:
                    img_name = src.split('/')[-1]
                    for attachment in attachments:
                        if attachment.name == img_name:
                            attachment.is_embedded = True
                            break

        new_mail = OfficeMail(
            email_id=email.get('id', None),
            email_type=email.get('@odata.type', "mail"),
            local_timestamp=local_time_str,
            from_name=mail_address.get('name', None),
            from_email=mail_address.get('address', None),
            subject=email.get('subject', None),
            body_preview=email['bodyPreview'],
            body=email.get('body', {}).get('content', ""),
            body_type=email.get('body', {}).get('contentType', None),
            is_read=email.get('isRead', False),
            has_attachments=email.get('hasAttachments', False),
            categories=email.get('categories', []),
            importance=email.get('importance', 'normal').lower(),
            confidential_level=email.get('sensitivity', 'normal').lower(),
            attachments=attachments,
            flag_state='flagged' if email.get('flag', {}).get('flagStatus', 'notFlagged') else 'notFlagged',
            web_link=email.get('webLink', None)
        )
        return new_mail

    def _build_mail_url(self, email_id: Optional[str] = None, url: Optional[str] = None, attachments: bool = True) -> str:
        if not url and not email_id:
            raise ValueError("Either email_id or url must be provided")
        if url is not None:
            if email_id is not None:
                raise ValueError("Only one of email_id or url should be provided")
            if not url.startswith(("http://", "https://")):
                url = f"{self.msg.msg_endpoint}{url}"
            elif not url.startswith("https://graph.microsoft.com/") and (
                not self.msg.msg_endpoint or not url.startswith(self.msg.msg_endpoint)
            ):
                raise ValueError("URL must point to the MS Graph API endpoint")
        else:
            url = f"{self.msg.msg_endpoint}me/messages/{email_id}"
        if attachments:
            if '?' in url:
                url += '&$expand=attachments'
            else:
                url += '?$expand=attachments'
        return url

    def _build_message_payload(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        is_html: bool = False,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
    ) -> dict:
        """Build a Graph API message JSON object (without attachments)."""
        def fmt(addrs):
            return [{"emailAddress": {"address": a}} for a in (addrs or [])]
        message: Dict = {
            "subject": subject,
            "body": {"contentType": "HTML" if is_html else "Text", "content": body},
            "from": {"emailAddress": {"address": self.msg.email}},
            "toRecipients": fmt(to_recipients),
        }
        if cc_recipients:
            message["ccRecipients"] = fmt(cc_recipients)
        if bcc_recipients:
            message["bccRecipients"] = fmt(bcc_recipients)
        return message

    @staticmethod
    def _parse_category(cat: dict) -> OfficeMailCategory:
        preset_color = cat.get('color', 'None')
        color_name = OFFICE_CATEGORY_PRESET_TO_NAME.get(preset_color, 'none')
        html_color = OFFICE_CATEGORY_NAME_TO_HTML.get(color_name, 'white')
        return OfficeMailCategory(
            id=cat.get('id', ''),
            name=cat.get('displayName', ''),
            preset_color=preset_color,
            color=html_color,
        )

    # ── async API ─────────────────────────────────────────────────────────

    async def get_user_profile_async(self):
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None
        response = await self.msg.run_async(url=f"{self.msg.msg_endpoint}me", token=access_token)
        if response is None or response.status_code != 200:
            return None
        return response.json()

    async def email_index_async(
        self, limit: int = 40, skip: int = 0, *,
        mail_address: Optional[str] = None,
        folder_id: str | None = None,
        query: str | None = None,
    ) -> OfficeMailList:
        """List or search messages.

        When *query* is provided, performs a full-text ``$search`` across
        all folders (``folder_id`` is ignored).  Otherwise lists messages
        in the given folder (defaults to Inbox).
        """
        if mail_address is None:
            mail_address = self.msg.email
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return OfficeMailList()

        base = "me" if not mail_address or mail_address == self.msg.email else f"users/{mail_address}"

        if query is not None:
            safe_q = query.replace('"', '\\"')
            fields = "id,from,subject,bodyPreview,receivedDateTime,isRead,hasAttachments,categories,importance"
            url = (
                f'{self.msg.msg_endpoint}{base}/messages'
                f'?$search="{safe_q}"&$select={fields}&$top={limit}'
            )
        else:
            fields = "isRead,id,from,subject,bodyPreview,receivedDateTime,hasAttachments,categories,importance,webLink"
            folder = folder_id or "inbox"
            url = (
                f"{self.msg.msg_endpoint}{base}/mailFolders/{folder}/messages"
                f"?$select={fields}&$top={limit}&$skip={skip}"
                f"&$orderby=receivedDateTime desc&$count=true"
            )

        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return OfficeMailList()
        json_object = response.json()
        emails = json_object.get('value', [])
        total_count = json_object.get('@odata.count', len(emails))
        email_list = []
        end_point = (self.msg.msg_endpoint or "").rstrip('/')
        for email in emails:
            new_mail = self.parse_mail(email)
            new_mail.email_url = f"{end_point}/{base}/messages/{new_mail.email_id}"
            email_list.append(new_mail)

        return OfficeMailList(elements=email_list, total_mails=total_count)

    async def get_mail_async(self, email_id: Optional[str] = None, email_url: Optional[str] = None, attachments=True) -> OfficeMail | None:
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None
        url = self._build_mail_url(email_id, email_url, attachments)
        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return None
        email = response.json()
        return self.parse_mail(email)

    async def set_mail_categories_async(self, email_url: str, categories: list[str]) -> bool:
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return False
        url = f"{email_url}"
        json_payload = {
            "categories": categories
        }
        response = await self.msg.run_async(url=url, method="PATCH", json=json_payload, token=access_token)
        return response is not None and response.status_code == 200

    async def get_categories_async(self, mail_address: str | None = None) -> list[OfficeMailCategory]:
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return []
        if not mail_address or mail_address == self.msg.email:
            url = f"{self.msg.msg_endpoint}me/outlook/masterCategories"
        else:
            url = f"{self.msg.msg_endpoint}users/{mail_address}/outlook/masterCategories"
        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return []
        categories = response.json().get("value", [])
        return [self._parse_category(cat) for cat in categories]

    async def ensure_category_exists_async(self, *, name: str, color: str = "preset0", mail_address: str | None = None) -> bool:
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return False
        if not mail_address or mail_address == self.msg.email:
            url = f"{self.msg.msg_endpoint}me/outlook/masterCategories"
        else:
            url = f"{self.msg.msg_endpoint}users/{mail_address}/outlook/masterCategories"
        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return False
        categories = response.json().get("value", [])
        if any(cat.get("displayName") == name for cat in categories):
            return True
        json_payload = {
            "displayName": name,
            "color": color
        }
        response = await self.msg.run_async(url=url, method="POST", json=json_payload, token=access_token)
        return response is not None and response.status_code == 201

    async def flag_read_async(self, email_url: str, read_state: bool) -> bool:
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return False
        url = f"{email_url}"
        json_payload = {
            "isRead": read_state
        }
        response = await self.msg.run_async(url=url, method="PATCH", json=json_payload, token=access_token)
        return response is not None and response.status_code == 200

    async def send_message_async(self, to_recipients: List[str], subject: str, body: str, is_html: bool = False, save_to_sent_items: bool = True, is_draft: bool = False, attachments: Optional[List[OfficeMailAttachment]] = None, cc_recipients: Optional[List[str]] = None, bcc_recipients: Optional[List[str]] = None) -> bool:
        """Send an email using Microsoft Graph API."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return False

        sender_email = self.msg.email
        fmt = lambda addrs: [{"emailAddress": {"address": e}} for e in addrs]
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if is_html else "Text",
                "content": body
            },
            "from": {
                "emailAddress": {
                    "address": sender_email
                }
            },
            "toRecipients": fmt(to_recipients),
            "isDraft": is_draft
        }
        if cc_recipients:
            message["ccRecipients"] = fmt(cc_recipients)
        if bcc_recipients:
            message["bccRecipients"] = fmt(bcc_recipients)

        if attachments:
            message["attachments"] = [{
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": att.name,
                "contentType": att.content_type,
                "contentBytes": base64.b64encode(att.content_bytes).decode() if att.content_bytes else None
            } for att in attachments]

        if is_draft:
            url = f"{self.msg.msg_endpoint}me/messages"
            response = await self.msg.run_async(url=url, method="POST", json=message, token=access_token)
            return response is not None and response.status_code == 201
        else:
            url = f"{self.msg.msg_endpoint}me/sendMail"
            json_payload = {
                "message": message,
                "saveToSentItems": str(save_to_sent_items).lower()
            }
            response = await self.msg.run_async(url=url, method="POST", json=json_payload, token=access_token)
            return response is not None and response.status_code == 202

    # ── Draft lifecycle (async) ───────────────────────────────────────────

    async def _add_attachments_async(self, message_id: str, attachments: List[OfficeMailAttachment], access_token: str) -> int:
        """Add attachments to an existing message. Returns count of successfully added."""
        added = 0
        url = f"{self.msg.msg_endpoint}me/messages/{message_id}/attachments"
        _log.info("[MailHandler] _add_attachments_async: msg=%s, count=%d",
                  message_id[:20] if message_id else None, len(attachments))
        for att in attachments:
            if not att.content_bytes:
                _log.warning("[MailHandler]   skip att %s: no content_bytes", att.name)
                continue
            _log.info("[MailHandler]   adding att: %s (%s, %d bytes)",
                      att.name, att.content_type, len(att.content_bytes))
            payload = {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": att.name,
                "contentType": att.content_type,
                "contentBytes": base64.b64encode(att.content_bytes).decode(),
            }
            resp = await self.msg.run_async(url=url, method="POST", json=payload, token=access_token)
            _log.info("[MailHandler]   response: status=%s", resp.status_code if resp is not None else None)
            if resp is not None and resp.status_code == 201:
                added += 1
            else:
                body_text = ""
                try:
                    body_text = resp.text[:500] if resp else "None"
                except Exception:
                    pass
                _log.warning("[MailHandler]   FAILED to add attachment %s: status=%s body=%s",
                             att.name, resp.status_code if resp is not None else None, body_text)
        _log.info("[MailHandler] _add_attachments_async: added %d/%d", added, len(attachments))
        return added

    async def _clear_attachments_async(self, message_id: str, access_token: str) -> None:
        """Remove all attachments from a message."""
        url = f"{self.msg.msg_endpoint}me/messages/{message_id}/attachments"
        resp = await self.msg.run_async(url=url, method="GET", token=access_token)
        if resp is None or resp.status_code != 200:
            return
        for att in resp.json().get("value", []):
            att_id = att.get("id")
            if att_id:
                await self.msg.run_async(url=f"{url}/{att_id}", method="DELETE", token=access_token)

    async def create_draft_async(
        self,
        to_recipients: List[str],
        subject: str,
        body: str,
        is_html: bool = False,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        attachments: Optional[List[OfficeMailAttachment]] = None,
    ) -> Optional[Dict]:
        """Create a draft and return ``{"id": "...", "webLink": "..."}`` or *None*."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None
        message = self._build_message_payload(
            to_recipients, subject, body, is_html,
            cc_recipients, bcc_recipients,
        )
        message["isDraft"] = True
        url = f"{self.msg.msg_endpoint}me/messages"
        resp = await self.msg.run_async(url=url, method="POST", json=message, token=access_token)
        if resp is not None and resp.status_code == 201:
            data = resp.json()
            msg_id = data.get("id", "")
            if attachments and msg_id:
                await self._add_attachments_async(msg_id, attachments, access_token)
            return {"id": msg_id, "webLink": data.get("webLink", "")}
        return None

    async def update_draft_async(
        self,
        message_id: str,
        to_recipients: List[str],
        subject: str,
        body: str,
        is_html: bool = False,
        cc_recipients: Optional[List[str]] = None,
        bcc_recipients: Optional[List[str]] = None,
        attachments: Optional[List[OfficeMailAttachment]] = None,
    ) -> Optional[Dict]:
        """Update an existing draft and return ``{"id": "...", "webLink": "..."}`` or *None*."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None
        message = self._build_message_payload(
            to_recipients, subject, body, is_html,
            cc_recipients, bcc_recipients,
        )
        url = f"{self.msg.msg_endpoint}me/messages/{message_id}"
        resp = await self.msg.run_async(url=url, method="PATCH", json=message, token=access_token)
        if resp is not None and resp.status_code == 200:
            data = resp.json()
            if attachments is not None:
                await self._clear_attachments_async(message_id, access_token)
                if attachments:
                    await self._add_attachments_async(message_id, attachments, access_token)
            return {"id": data.get("id", ""), "webLink": data.get("webLink", "")}
        return None

    async def send_draft_async(self, message_id: str) -> bool:
        """Send an existing draft by its message ID."""
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return False
        url = f"{self.msg.msg_endpoint}me/messages/{message_id}/send"
        resp = await self.msg.run_async(url=url, method="POST", token=access_token)
        return resp is not None and resp.status_code == 202

    # ── Delete, move ───────────────────────────────────────────────────

    async def delete_message_async(self, message: "str | OfficeMail") -> bool:
        """Soft-delete a message (moves to Deleted Items)."""
        msg_id = message.email_id if isinstance(message, OfficeMail) else message
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return False
        resp = await self.msg.run_async(
            url=f"{self.msg.msg_endpoint}me/messages/{msg_id}",
            method="DELETE", token=access_token,
        )
        return resp is not None and resp.status_code < 300

    async def move_message_async(
        self, message: "str | OfficeMail", destination: "str | FolderInfo",
    ) -> MoveResult | None:
        """Move a message to another folder."""
        msg_id = message.email_id if isinstance(message, OfficeMail) else message
        folder_id = destination.id if isinstance(destination, FolderInfo) else destination
        access_token = await self.msg.get_access_token_async()
        if not access_token:
            return None
        resp = await self.msg.run_async(
            url=f"{self.msg.msg_endpoint}me/messages/{msg_id}/move",
            method="POST",
            json={"destinationId": folder_id},
            token=access_token,
        )
        if resp is not None and resp.status_code in (200, 201):
            data = resp.json()
            return MoveResult(id=data.get("id", ""), web_link=data.get("webLink", ""))
        return None


class OfficeCategoryColor:
    """Outlook category colors"""
    NONE = "None"
    RED = "preset0"
    ORANGE = "preset1"
    BROWN = "preset2"
    YELLOW = "preset3"
    GREEN = "preset4"
    TEAL = "preset5"
    OLIVE = "preset6"
    BLUE = "preset7"
    PURPLE = "preset8"
    CRANBERRY = "preset9"
    STEEL = "preset10"
    DARK_STEEL = "preset11"
    GRAY = "preset12"
    DARK_GRAY = "preset13"
    BLACK = "preset14"
    DARK_RED = "preset15"
    DARK_ORANGE = "preset16"
    DARK_BROWN = "preset17"
    DARK_YELLOW = "preset18"
    DARK_GREEN = "preset19"
    DARK_TEAL = "preset20"
    DARK_OLIVE = "preset21"
    DARK_BLUE = "preset22"
    DARK_PURPLE = "preset23"
    DARK_CRANBERRY = "preset24"

# Reverse mapping: preset -> color name
OFFICE_CATEGORY_PRESET_TO_NAME = {v: k.lower() for k, v in OfficeCategoryColor.__dict__.items() if not k.startswith('__') and not callable(v)}
# HTML color mapping
OFFICE_CATEGORY_NAME_TO_HTML = {
    'red': 'red',
    'orange': 'orange',
    'brown': 'brown',
    'yellow': 'yellow',
    'green': 'rgb(76, 187, 23)',
    'teal': 'teal',
    'olive': 'olive',
    'blue': 'blue',
    'purple': 'purple',
    'cranberry': 'crimson',
    'steel': 'slateblue',
    'dark_steel': 'slategray',
    'gray': 'gray',
    'dark_gray': 'dimgray',
    'black': 'black',
    'dark_red': 'darkred',
    'dark_orange': 'darkorange',
    'dark_brown': 'saddlebrown',
    'dark_yellow': 'goldenrod',
    'dark_green': 'darkgreen',
    'dark_teal': 'teal',
    'dark_olive': 'olivedrab',
    'dark_blue': 'darkblue',
    'dark_purple': 'indigo',
    'dark_cranberry': 'firebrick',
    'none': 'white',
}

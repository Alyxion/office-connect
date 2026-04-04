import base64
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


class OfficeMailList(BaseModel):
    """Paginated list of Outlook email messages."""
    elements: List[OfficeMail] = Field(default_factory=list, description="Email messages in this page")
    total_mails: int = Field(default=0, description="Total number of mails in the folder")


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

    async def email_index_async(self, limit: int = 40, skip: int = 0, mail_address: Optional[str] = None) -> OfficeMailList:
        if mail_address is None:
            mail_address = self.msg.email
        access_token = await self.msg.get_access_token_async()

        if not access_token:
            return OfficeMailList()

        fields = "isRead,id,from,subject,bodyPreview,receivedDateTime,hasAttachments,categories,importance,webLink"

        if not mail_address or mail_address == self.msg.email:
            url = f"{self.msg.msg_endpoint}me/mailFolders/inbox/messages?$select={fields}&top={limit}&skip={skip}&$count=true"
        else:
            url = f"{self.msg.msg_endpoint}users/{mail_address}/mailFolders/Inbox/messages?$select={fields}&top={limit}&skip={skip}&$count=true"
        response = await self.msg.run_async(url=url, token=access_token)
        if response is None or response.status_code != 200:
            return OfficeMailList()
        json_object = response.json()
        emails = json_object.get('value', [])
        total_count = json_object.get('@odata.count', 0)
        email_list = []
        end_point = (self.msg.msg_endpoint or "").rstrip('/')
        for email in emails:
            new_mail = self.parse_mail(email)
            new_mail.email_url = f"{end_point}/users/{mail_address}/messages/{new_mail.email_id}"
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
            status = getattr(resp, 'status_code', None)
            _log.info("[MailHandler]   response: status=%s", status)
            if resp is not None and status == 201:
                added += 1
            else:
                body_text = ""
                try:
                    body_text = resp.text[:500] if resp else "None"
                except Exception:
                    pass
                _log.warning("[MailHandler]   FAILED to add attachment %s: status=%s body=%s",
                             att.name, status, body_text)
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

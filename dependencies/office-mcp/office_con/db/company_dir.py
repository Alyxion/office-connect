from __future__ import annotations

import fnmatch
import io
import logging

from typing import List, Callable, Optional, Literal, ClassVar, Set, TYPE_CHECKING
import json
from pydantic import BaseModel, Field
from office_con.msgraph.directory_handler import DirectoryUser, DirectoryHandler
import threading

if TYPE_CHECKING:
    from office_con.db.company_dir_builder import UserReplacements
    from office_con.msgraph.ms_graph_handler import MsGraphInstance


class CompanyUser(DirectoryUser):
    """
    Extends :class:`DirectoryUser` with company-specific fields.
    """
    # Object type constants
    OT_USER: ClassVar[str] = "user"
    OT_SERVICE_ACCOUNT: ClassVar[str] = "serviceAccount"
    OT_ROOM: ClassVar[str] = "room"
    OT_DEVICE: ClassVar[str] = "device"
    OT_GROUP: ClassVar[str] = "group"
    OT_APPLICATION: ClassVar[str] = "application"
    OT_CALENDAR: ClassVar[str] = "calendar"
    OT_SERVICE_PRINCIPAL: ClassVar[str] = "servicePrincipal"

    company: str | None = Field(default=None, description="Company of the user")
    external: bool = Field(default=False, description="Is the user external?")
    gender: Literal["undefined", "male", "female", "other"] = Field(default="undefined", description="Gender of the user")
    object_type: str = Field(default=OT_USER, description="Type of the directory object")
    building: str | None = Field(default=None, description="Building of the user")
    street: str | None = Field(default=None, description="Street of the user")
    manager_email: Optional[str] = Field(default=None, description="Manager email of the user")
    zip: str | None = Field(default=None, description="Zip code of the user")    
    city: str | None = Field(default=None, description="City of the user")
    country: str | None = Field(default=None, description="Country of the user")
    room_name: str | None = Field(default=None, description="Room name of the user")
    guessed_fields: dict[str, bool] = Field(default_factory=dict, description="Fields that have been guessed")
    has_image: bool = Field(default=False, description="Has an image")

    join_date: str | None = Field(default=None, description="Join date of the user")
    birth_date: str | None = Field(default=None, description="Birth date of the user")
    termination_date: str | None = Field(default=None, description="Termination date of the user")
    

    @property
    def main_type(self) -> str:
        """Returns the main type part of object_type (before colon, if present)."""
        return self.object_type.split(".", 1)[0] if self.object_type else ""

    @property
    def sub_type(self) -> str | None:
        """Returns the sub type part of object_type (after colon, if present), else None."""
        parts = self.object_type.split(".", 1)
        return parts[1] if len(parts) == 2 else None

    # Hardcoded list of fields for rule matching (string, bool, numeric). Update this list if fields change.
    # All fields from DirectoryUser and CompanyUser. Keep in sync with both models.
    _match_fields: ClassVar[Set[str]] = {
        # DirectoryUser fields
        "id",
        "display_name",
        "email",
        "job_title",
        "department",
        "account_enabled",
        "surname",
        "given_name",
        "office_location",
        "mobile_phone",
        # CompanyUser fields
        "company",
        "external",
        "gender",
        "object_type",
        "building",
        "street",
        "zip",
        "city",
        "country",
        "room_name",
        "has_image",
        "join_date",
        "termination_date",
        "birth_date",
        "manager_email",
        # Add more fields here if needed
    }

    def matches_rule(self, rule: dict, use_guessed: bool = True) -> bool:
        """
        Returns True if this user matches the given rule.
        Supports _, _|, _& for catch-all (OR/AND), and per-field filters.
        Matches on both string and numeric/bool fields (see _match_fields).
        """
        mask_key = None
        mask_value = None
        for key in rule.keys():
            if key.startswith("_"):
                mask_key = key
                mask_value = rule[key].lower() if isinstance(rule[key], str) else rule[key]
        mask_values: list[str] = [m.strip() for m in mask_value.split('|') if m.strip()] if isinstance(mask_value, str) else [mask_value]  # type: ignore[list-item]
        if mask_key is None:
            return False
        mask_key = mask_key[1:]
        matching = False
        if mask_key == "":
            for field in self._match_fields:
                # if we are not guessing we are also not allowed to use any guessed fields
                if not use_guessed and field in self.guessed_fields:
                    continue
                attr = self.__getattribute__(field)
                if isinstance(attr, str):
                    for mask in mask_values:
                        if fnmatch.fnmatch(attr.lower(), mask):
                            matching = True
                            break
        else:
            if mask_key not in self._match_fields:
                return False
            if not use_guessed and mask_key in self.guessed_fields:
                return False
            attr = self.__getattribute__(mask_key)
            if isinstance(attr, str):
                for mask in mask_values:
                    if fnmatch.fnmatch(attr.lower(), mask):
                        matching = True
                        break
            elif isinstance(attr, bool):
                if attr == mask_value:
                    matching = True
        return matching

class CompanyUserList(BaseModel):
    """Versioned container for a list of company directory users."""
    CURRENT_VERSION: ClassVar[str] = "1.0"
    version: str = Field(default="1.0", description="Schema version for forward compatibility")
    users: List[CompanyUser] = Field(default_factory=list, description="Company directory users")


class CompanyDirData:
    """In-memory company directory data (users and images).

    Used as a data container for CompanyDir. No disk I/O — all data
    lives in memory and is populated by LiveCompanyDirData or programmatically.
    """

    def __init__(self, src: object = None, in_memory: bool = False) -> None:
        self.users: CompanyUserList | None = CompanyUserList()

    def get_image_bytes(self, user_id: str) -> bytes | None:
        return None

    async def get_image_bytes_async(self, user_id: str) -> bytes | None:
        return None

_log = logging.getLogger(__name__)


def generate_initials_avatar(display_name: str, user_id: str, width: int = 256, height: int = 256) -> bytes:
    """Generate a pastel-colored avatar with the user's initials.

    Uses a deterministic color derived from user_id so the same user always
    gets the same background color.
    """
    import hashlib
    from PIL import Image, ImageDraw, ImageFont

    # Extract initials (first letter of first + last name)
    parts = (display_name or "?").split()
    if len(parts) >= 2:
        initials = (parts[0][0] + parts[-1][0]).upper()
    elif parts:
        initials = parts[0][0].upper()
    else:
        initials = "?"

    # Deterministic pastel color from user_id hash
    h = int(hashlib.md5((user_id or display_name or "").encode()).hexdigest(), 16)
    hue = h % 360
    # HSL → RGB with high lightness (pastel)
    from colorsys import hls_to_rgb
    r, g, b = hls_to_rgb(hue / 360.0, 0.75, 0.5)
    bg_color = (int(r * 255), int(g * 255), int(b * 255))

    img = Image.new("RGB", (width, height), bg_color)
    draw = ImageDraw.Draw(img)

    # Use a large font size relative to image dimensions
    font_size = min(width, height) // 2
    font = None
    for font_path in [
        "/System/Library/Fonts/Helvetica.ttc",          # macOS
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",  # Debian/Ubuntu
        "/usr/share/fonts/dejavu-sans-fonts/DejaVuSans-Bold.ttf",  # RHEL/Fedora
    ]:
        try:
            font = ImageFont.truetype(font_path, font_size)
            break
        except (OSError, IOError):
            continue
    if font is None:
        font = ImageFont.load_default()

    # Center the text
    bbox = draw.textbbox((0, 0), initials, font=font)
    text_w = bbox[2] - bbox[0]
    text_h = bbox[3] - bbox[1]
    x = (width - text_w) / 2 - bbox[0]
    y = (height - text_h) / 2 - bbox[1]
    draw.text((x, y), initials, fill="white", font=font)

    out = io.BytesIO()
    img.save(out, "JPEG", quality=90)
    return out.getvalue()


class LiveCompanyDirData:
    """CompanyDirData-compatible backend backed by a live DirectoryHandler.

    Users are populated in a background thread (started once).
    Photos are fetched on demand, never in bulk.
    All access is thread-safe.
    """

    def __init__(self, user_replacements: UserReplacements | None = None, add_genders: bool = False) -> None:
        self._lock = threading.RLock()
        self._users: CompanyUserList = CompanyUserList()
        self._user_by_id: dict[str, CompanyUser] = {}
        self._user_by_email: dict[str, CompanyUser] = {}
        self._handler: DirectoryHandler | None = None
        self._replacements = user_replacements
        self._add_genders = add_genders
        self._populated = False
        self._populating = False
        self._on_populated: Callable[[], None] | None = None
        self._photo_known: dict[str, bool] = {}  # survives population swaps
        self._photo_blobs: dict[str, bytes] = {}  # cached raw photo bytes

    @property
    def is_populated(self) -> bool:
        return self._populated

    @property
    def users(self) -> CompanyUserList:
        with self._lock:
            return self._users

    def start_population(self, graph_instance: MsGraphInstance, *, on_done: Callable[[], None] | None = None) -> None:
        """Start async task to fetch all users. Called once on first auth.

        on_done is called after population completes so CompanyDir can rebuild its indices.
        """
        import asyncio
        with self._lock:
            if self._populating:
                if self._handler:
                    self._handler.msg = graph_instance
                return
            self._populating = True
            self._handler = DirectoryHandler(graph_instance)
            self._on_populated = on_done

        asyncio.get_running_loop().create_task(self._populate_async())

    def refresh_graph(self, graph_instance: MsGraphInstance) -> None:
        """Keep graph token fresh (called on subsequent logins)."""
        with self._lock:
            if self._handler:
                self._handler.msg = graph_instance

    async def _populate_async(self) -> None:
        """Runs as async task. Fetches users, enriches, atomic swap."""
        try:
            dir_users = await self._handler.get_all_users_async()  # type: ignore[union-attr]

            company_users = CompanyUserList(
                users=[CompanyUser(**u.model_dump()) for u in dir_users.users]
            )

            if self._replacements:
                from office_con.db.company_dir_builder import apply_user_replacements
                apply_user_replacements(company_users, self._replacements, add_genders=self._add_genders)

            self._scan_photo_status(company_users)

            by_id = {u.id: u for u in company_users.users}
            by_email = {u.email.lower(): u for u in company_users.users if u.email}

            with self._lock:
                self._users = company_users
                self._user_by_id = by_id
                self._user_by_email = by_email
                # Re-apply photo status from in-memory cache (survives swap)
                for uid, has_img in self._photo_known.items():
                    if uid in by_id:
                        by_id[uid].has_image = has_img
                self._populated = True

            _log.info("LiveCompanyDirData: populated %d users", len(company_users.users))

            if self._on_populated:
                self._on_populated()
        except Exception:
            _log.exception("LiveCompanyDirData: background population failed")
            with self._lock:
                self._populating = False

    def _scan_photo_status(self, users: CompanyUserList) -> None:
        """Set has_image from in-memory photo cache."""
        for user in users.users:
            if user.id in self._photo_known:
                user.has_image = self._photo_known[user.id]

    def get_user_by_id(self, user_id: str) -> CompanyUser | None:
        with self._lock:
            return self._user_by_id.get(user_id)

    def get_user_by_email(self, email: str) -> CompanyUser | None:
        with self._lock:
            return self._user_by_email.get(email.lower())

    def get_image_bytes(self, user_id: str) -> bytes | None:
        """Return cached photo blob or None."""
        return self._photo_blobs.get(user_id)

    async def get_image_bytes_async(self, user_id: str) -> bytes | None:
        """Fetch photo on demand from MS Graph. Results cached in memory."""
        # Return cached blob if available
        cached = self._photo_blobs.get(user_id)
        if cached is not None:
            return cached
        # Skip fetch if known to have no photo
        with self._lock:
            if user_id in self._photo_known and not self._photo_known[user_id]:
                return None

        handler = self._handler
        if not handler:
            return None

        _log.info("[LIVE-DIR] Fetching photo on demand from MS Graph for %s", user_id)
        blob = await handler.get_user_photo_async(user_id)

        has_img = blob is not None
        with self._lock:
            self._photo_known[user_id] = has_img
            if blob:
                self._photo_blobs[user_id] = blob
            user = self._user_by_id.get(user_id)
            if user:
                user.has_image = has_img
        return blob

    def check_photo_cache(self, user_id: str) -> bool:
        """Check in-memory cache for a user's photo status and set has_image.

        Returns True if the photo status is known, False if a network fetch is needed.
        """
        with self._lock:
            if user_id in self._photo_known:
                user = self._user_by_id.get(user_id)
                if user:
                    user.has_image = self._photo_known[user_id]
                return True
        return False

    async def prefetch_photo_async(self, user_id: str) -> None:
        """Async prefetch a single user's photo and set has_image.

        Call this during login so that has_image is accurate by the time
        the page renders.
        """
        handler = self._handler
        if not handler:
            return

        # Skip if already known
        with self._lock:
            if user_id in self._photo_known:
                return

        _log.info("[LIVE-DIR] Async prefetch photo for %s", user_id)
        blob = await handler.get_user_photo_async(user_id)

        has_img = blob is not None
        with self._lock:
            self._photo_known[user_id] = has_img
            if blob:
                self._photo_blobs[user_id] = blob
            user = self._user_by_id.get(user_id)
            if user:
                user.has_image = has_img



class CompanyDir:
    """
    Loads users and images from a local directory, a CompanyDirData bundle, or a LiveCompanyDirData backend.
    Provides get_user_image_bytes(user_id) for lazy image access from folder, zip, or on-demand Graph fetch.
    """

    def __init__(self, source: str | None | CompanyDirData | LiveCompanyDirData = None,
                 user_image_url_callback: Optional[Callable] = None, in_memory: bool = False):
        """
        :param source: Path, CompanyDirData, or LiveCompanyDirData instance.
        :param user_image_url_callback: Optional callback (user_id, ext, width, height) -> url
        :param in_memory: If True, the CompanyDirData is loaded into memory
        """
        self._live = isinstance(source, LiveCompanyDirData)
        if self._live:
            self.data = source
        elif isinstance(source, CompanyDirData):
            self.data = source
        elif source is not None:
            self.data = CompanyDirData(source, in_memory=in_memory)
        else:
            self.data = CompanyDirData(in_memory=in_memory)
        self.user_image_url_callback = user_image_url_callback
        self._user_by_email: dict[str, CompanyUser] = {}
        self._user_by_id: dict[str, CompanyUser] = {}
        self._users: CompanyUserList | None = None
        self._image_cache: dict[str, bytes] = {}
        self._image_lock = threading.RLock()
        if not self._live:
            self._load_users()

    def refresh(self) -> None:
        """Rebuild local indices from the data source.

        Called by LiveCompanyDirData._on_populated callback after background
        population completes. Also usable for CompanyDirData sources.
        """
        if self._live:
            live: LiveCompanyDirData = self.data  # type: ignore[assignment]
            with live._lock:
                self._users = live._users
                self._user_by_id = dict(live._user_by_id)
                self._user_by_email = dict(live._user_by_email)
        else:
            self._load_users()
        # Invalidate resized-image cache so new photos are picked up
        with self._image_lock:
            self._image_cache.clear()

    def get_user_image_bytes(self, user_id: str) -> bytes | None:
        """Returns the raw image bytes for a user."""
        return self.data.get_image_bytes(user_id)

    def get_user_image_url(self, user_id: str, ext: str = '.jpg', *, width: int = 256, height: int = 256) -> str:
        if self.user_image_url_callback:
            return self.user_image_url_callback(user_id, ext, width=width, height=height)  # type: ignore[misc]
        import os
        base_url = os.environ.get('WEBSITE_URL', 'https://localhost:8000').rstrip('/')
        return f"{base_url}/profiles/{user_id}{ext}"

    def _load_users(self) -> None:
        self._users = self.data.users
        if not self._users:
            return
        try:
            self._user_by_email = {u.email.lower(): u for u in self._users.users if u.email}
            self._user_by_id = {u.id: u for u in self._users.users if u.id}
        except (json.JSONDecodeError, TypeError, ValueError, KeyError):
            pass

    @property
    def users(self) -> List[CompanyUser]:
        if self._live:
            ul = self.data.users  # type: ignore[union-attr]
            return ul.users if ul else []
        return self._users.users if self._users else []

    def get_user_by_id(self, user_id: str) -> CompanyUser | None:
        """Fast lookup by user id, returns user or None."""
        if self._live:
            return self.data.get_user_by_id(user_id)  # type: ignore[union-attr]
        return self._user_by_id.get(user_id)

    def get_user_by_email(self, email: str) -> CompanyUser | None:
        """Fast lookup by email (case-insensitive), returns user or None."""
        if not email:
            return None
        if self._live:
            return self.data.get_user_by_email(email)  # type: ignore[union-attr]
        return self._user_by_email.get(email.lower())

    def get_user_image(self, user_id: str, *, width: int = 256, height: int = 256) -> bytes:
        def resize_image(image_bytes: bytes, width: int, height: int) -> bytes:
            from PIL import Image
            image = Image.open(io.BytesIO(image_bytes))
            # ensure RGB
            if image.mode != 'RGB':
                image = image.convert('RGB')
            # if image not quadratic, crop on shorter side
            if image.width != image.height:
                if image.width > image.height:
                    image = image.crop((0, 0, image.height, image.height))
                else:
                    image = image.crop((0, 0, image.width, image.width))
            image = image.resize((width, height), Image.LANCZOS)
            # encode to jpg
            out = io.BytesIO()
            image.save(out, 'JPEG')
            return out.getvalue()

        with self._image_lock:
            valid_resolutions = {32, 48, 64, 96, 128, 256, 512}
            if width not in valid_resolutions or height not in valid_resolutions:
                raise ValueError(f"Invalid image resolution: {width}x{height}")
            cache_id = f"{user_id}_{width}_{height}"
            if cache_id in self._image_cache:
                return self._image_cache[cache_id]
            image_bytes = self.data.get_image_bytes(user_id)  # type: ignore[union-attr]
            if image_bytes is None:
                # Generate initials avatar as fallback — do NOT cache so that
                # subsequent requests can pick up the real photo once it
                # becomes available (e.g. after directory population completes
                # or MS Graph fetch succeeds).
                user = self.get_user_by_id(user_id)
                display_name = user.display_name if user else user_id
                return generate_initials_avatar(display_name, user_id, width, height)
            resized_image = resize_image(image_bytes, width, height)
            self._image_cache[cache_id] = resized_image
            return resized_image

    def find_users(self, *, display_name: str | None = None, email: str | None = None, first_name: str | None = None, last_name: str | None = None, exact: bool = False) -> list[CompanyUser]:
        """
        Search users by any combination of display_name, email, first_name (given_name), and last_name (surname).
        - Name fields: case-insensitive substring match (unless exact=True).
        - Email: case-insensitive exact match.
        Returns a list of matching users.
        """
        results = []
        for user in self.users:
            match = True
            if display_name:
                val = user.display_name.lower() if user.display_name else ""
                if (exact and val != display_name.lower()) or (not exact and display_name.lower() not in val):
                    match = False
            if email:
                val = user.email.lower() if user.email else ""
                if val != email.lower():
                    match = False
            if first_name:
                val = user.given_name.lower() if user.given_name else ""
                if (exact and val != first_name.lower()) or (not exact and first_name.lower() not in val):
                    match = False
            if last_name:
                val = user.surname.lower() if user.surname else ""
                if (exact and val != last_name.lower()) or (not exact and last_name.lower() not in val):
                    match = False
            if match:
                results.append(user)
        return results

    def get_user(self, **kwargs: str | bool | None) -> CompanyUser | None:
        """
        Return the first user matching the given criteria (see find_users).
        """
        found = self.find_users(**kwargs)  # type: ignore[arg-type]
        return found[0] if found else None

    def users_to_dataframe(self, rules: list[dict[str, str]] | dict[str, str], use_guessed: bool = True) -> object:
        """
        Return a pandas DataFrame for users matching any of the given rules (mask(s)).
        Flattens all nested structures to strings.
        In the output, replaces manager_id with manager_email (looked up from user list).
        
        :param rules: A rule dict or list of rule dicts
        :param use_guessed: Whether to allow guessed fields in matching
        :return: pandas.DataFrame
        """
        try:
            import pandas as pd
        except ImportError:
            raise ImportError("pandas is required for users_to_dataframe")

        if isinstance(rules, dict):
            rules = [rules]
        filtered = []
        for user in self.users:
            for rule in rules:
                if hasattr(user, 'matches_rule') and user.matches_rule(rule, use_guessed=use_guessed):
                    filtered.append(user)
                    break

        def flatten(val):
            if isinstance(val, dict):
                return json.dumps(val, ensure_ascii=False)
            if isinstance(val, list):
                return ", ".join(str(flatten(v)) for v in val)
            return str(val) if val is not None else ""

        # Build id->email mapping for manager lookup
        id_to_email = {u.id: u.email for u in self.users}

        records = []
        for user in filtered:
            data = user.model_dump()
            flat = {k: flatten(v) for k, v in data.items()}
            # Replace manager_id with manager_email in export
            if 'manager_id' in flat:
                manager_email = id_to_email.get(flat['manager_id'], "")
                flat['manager_email'] = manager_email
                del flat['manager_id']
            records.append(flat)
        return pd.DataFrame(records)


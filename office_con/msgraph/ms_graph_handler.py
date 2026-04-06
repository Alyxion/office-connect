import os
import logging
import time as _time
from typing import Optional, TYPE_CHECKING

import aiohttp
from fastapi.responses import HTMLResponse
from msal import ConfidentialClientApplication

from office_con.msgraph.mail_handler import OfficeMailHandler
from office_con.msgraph.profile_handler import ProfileHandler, UserProfile
from office_con.msgraph.calendar_handler import CalendarHandler
from ..web_user_instance import WebUserInstance

if TYPE_CHECKING:
    from office_con.testing.mock_data import MockUserProfile

logger = logging.getLogger(__name__)


class MsGraphInstance(WebUserInstance):

    def __init__(self, scopes: list[str] | None = None,
                *,
                 cache_dict: Optional[dict] = None,
                 redis_url: str | None = None,
                 mongodb_url: str | None = None,
                 auth_url: str | None = None,
                 app="office",
                 session_id: str | None = None,
                 can_refresh: bool = True,
                 client_id: str | None = None,
                 client_secret: str | None = None,
                 endpoint: str | None = None,
                 tenant_id: str | None = None,
                 select_account: bool = False):

        super().__init__(app, cache_dict=cache_dict, redis_url=redis_url, session_id=session_id, mongodb_url=mongodb_url)
        self.scopes = scopes

        self.client_id = client_id or os.environ.get('O365_CLIENT_ID')
        self.client_secret = client_secret or os.environ.get('O365_CLIENT_SECRET')
        self.msg_endpoint = endpoint or os.environ.get("O365_ENDPOINT")
        self.tenant_id = tenant_id or os.environ.get("O365_TENANT_ID", "common")
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        self.ensure_unique_user_id()
        self.auth_url = auth_url
        self._msg_handler: Optional[OfficeMailHandler] = None
        self._profile_handler: Optional[ProfileHandler] = None
        self._calendar_handler: Optional[CalendarHandler] = None
        self.me: Optional[UserProfile] = None
        self.can_refresh = can_refresh
        self.features = {WebUserInstance.FEATURE_MAIL, WebUserInstance.FEATURE_CALENDAR, WebUserInstance.FEATURE_PROFILE}

        self.auth_kwargs = {}
        if select_account:
            self.auth_kwargs["prompt"] = "select_account"

    async def refresh_async(self) -> str | None:
        if not self.can_refresh:
            return None
        if not (token := await self.get_access_token_async()):
            logger.warning("[OP] refresh_async — no access token in cache/Redis")
            return None
        expiration_time = self.time_until_token_expiration(token)
        if expiration_time <= self.min_expiry:
            logger.info("[OP] refresh_async — token expiring in %ds, refreshing", expiration_time)
            return await self.refresh_token_async()
        else:
            return None

    async def refresh_token_async(self) -> str | None:
        if not self.can_refresh:
            return None

        refresh_token = await self.get_refresh_token_async()
        if not refresh_token:
            logger.warning("[OP] refresh_token_async — no refresh token available (memory cache empty, Redis returned nothing)")
            return None
        _pid = os.getpid()
        logger.info("[OP] token_refresh — starting (pid=%d)", _pid)
        _t0 = _time.monotonic()
        result = await self._async_token_request(
            grant_type="refresh_token",
            refresh_token=refresh_token,
        )
        _elapsed = (_time.monotonic() - _t0) * 1000
        if "error" in result:
            logger.warning("[OP] token_refresh — failed (%.0fms, pid=%d): %s", _elapsed, _pid, result.get("error_description", result["error"]))
            # Keep the expired access token in Redis — deleting it would prevent
            # the next page load from even attempting a refresh (the factory checks
            # for access_token presence before trying refresh_token_async).
            return None
        logger.info("[OP] token_refresh — done (%.0fms, pid=%d)", _elapsed, _pid)
        await self.set_access_token_async(result["access_token"])
        await self.set_refresh_token_async(result.get("refresh_token"))
        return result["access_token"]

    async def acquire_token_async(self, code, redirect_url: str):
        _pid = os.getpid()
        logger.info("[OP] acquire_token — starting (pid=%d)", _pid)
        _t0 = _time.monotonic()
        result = await self._async_token_request(
            grant_type="authorization_code",
            code=code,
            redirect_uri=redirect_url,
        )
        _elapsed = (_time.monotonic() - _t0) * 1000
        if "error" in result:
            logger.warning("[OP] acquire_token — failed (%.0fms, pid=%d): %s", _elapsed, _pid, result.get("error_description", result["error"]))
            return HTMLResponse("Login failed. Please try again.", status_code=400)
        logger.info("[OP] acquire_token — done (%.0fms, pid=%d)", _elapsed, _pid)
        await self.set_access_token_async(result["access_token"])
        await self.set_refresh_token_async(result.get("refresh_token"))

        # Fetch profile to populate email/user_id
        logger.info("[OP] acquire_token_profile — starting (pid=%d)", _pid)
        _t0 = _time.monotonic()
        profile = await self.get_profile_async()
        logger.info("[OP] acquire_token_profile — done (%.0fms, pid=%d)", (_time.monotonic() - _t0) * 1000, _pid)
        if not profile.me or not profile.me.id:
            return HTMLResponse("Failed to verify user profile", status_code=400)

        # Cache profile to Redis so subsequent page loads skip the MS Graph HTTP call
        await self.cache_profile_to_redis_async()
        return HTMLResponse("Token acquired successfully")

    def build_msal_app(self, authority=None):
        return ConfidentialClientApplication(
            self.client_id, authority=authority or self.authority,
            client_credential=self.client_secret,
            azure_region=os.environ.get("O365_MSAL_REGION") or None)

    def build_auth_url(self, auth_url, authority=None):
        # Mock OAuth: redirect back immediately with fake auth code
        if self._mock_transport is not None:
            from urllib.parse import urlencode
            return f"{auth_url}?{urlencode({'code': 'mock-auth-code'})}"
        return self.build_msal_app(authority).get_authorization_request_url(
            **self.auth_kwargs,
            scopes=self.scopes or [],
            redirect_uri=auth_url)

    @property
    def _token_endpoint(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

    async def _async_token_request(self, **form_data) -> dict:
        """POST to the Azure AD token endpoint using aiohttp (non-blocking)."""
        # Mock token intercept
        if self._mock_transport is not None:
            from office_con.testing.mock_tokens import make_mock_token_response
            return make_mock_token_response(self._mock_profile.email, self._mock_profile.user_id)
        grant_type = form_data.get("grant_type", "unknown")
        form_data.setdefault("client_id", self.client_id)
        form_data.setdefault("client_secret", self.client_secret)
        if self.scopes:
            form_data.setdefault("scope", " ".join(self.scopes))
        timeout = aiohttp.ClientTimeout(total=30)
        _pid = os.getpid()
        logger.info("[OP] azure_ad_token_request(%s) — starting (pid=%d)", grant_type, _pid)
        _t0 = _time.monotonic()
        try:
            async with aiohttp.ClientSession(timeout=timeout) as session:
                async with session.post(self._token_endpoint, data=form_data) as resp:
                    result = await resp.json()
                    logger.info("[OP] azure_ad_token_request(%s) — done (%.0fms, pid=%d, status=%d)", grant_type, (_time.monotonic() - _t0) * 1000, _pid, resp.status)
                    return result
        except Exception as exc:
            logger.error("[OP] azure_ad_token_request(%s) — failed (%.0fms, pid=%d): %s", grant_type, (_time.monotonic() - _t0) * 1000, _pid, exc)
            return {"error": "async_request_failed", "error_description": str(exc)}

    def get_mail(self):
        return OfficeMailHandler(self)

    def get_calendar(self):
        return CalendarHandler(self)

    def get_directory(self):
        from .directory_handler import DirectoryHandler
        return DirectoryHandler(self)

    def get_teams(self):
        from .teams_handler import TeamsHandler
        return TeamsHandler(self)

    def get_chat(self):
        from .chat_handler import ChatHandler
        return ChatHandler(self)

    def get_files(self):
        from .files_handler import FilesHandler
        return FilesHandler(self)

    def enable_mock(self, profile: "MockUserProfile"):
        """Enable mock transport — all HTTP calls return synthetic data.

        Args:
            profile: A ``MockUserProfile`` instance with synthetic data.
        """
        from office_con.testing.mock_transport import MockGraphTransport
        self._mock_transport = MockGraphTransport(profile)
        self._mock_profile = profile

    async def get_profile_async(self):
        """Get the profile handler, fetching from MS Graph if needed."""
        has_valid_profile = self.me is not None and self.me.id
        profile_handler = ProfileHandler(self, me=self.me if has_valid_profile else None)

        if not has_valid_profile:
            me = await profile_handler.me_async()
            with self.access_lock:
                if self._profile_handler is None:
                    self._profile_handler = profile_handler
                if me and me.id:
                    self.user_id = me.id
                    self.given_name = me.given_name or me.display_name or "Anonymous"
                    self.surname = me.surname
                    self.full_name = me.display_name
                    self.email = (me.mail or me.user_principal_name).lower()
                    self.me = me

        return profile_handler

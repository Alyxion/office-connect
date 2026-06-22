import asyncio
import os
import logging
import time as _time
from typing import Optional, TYPE_CHECKING

import aiohttp
from fastapi.responses import HTMLResponse
from msal import ConfidentialClientApplication

from office_con.msgraph.mail_handler import OfficeMailHandler, MailFolderHandler
from office_con.msgraph.profile_handler import ProfileHandler, UserProfile
from office_con.msgraph.calendar_handler import CalendarHandler
from office_con.privacy import OfficePrivacyConfig
from ..web_user_instance import WebUserInstance

if TYPE_CHECKING:
    from office_con.testing.mock_data import MockUserProfile

logger = logging.getLogger(__name__)

# Throttling (HTTP 429) handling for run_async. Graph throttles /search/query
# and other endpoints aggressively under bursty/intensive use; without honoring
# Retry-After the call surfaces a misleading "search failed" to the user.
_MAX_THROTTLE_RETRIES = 3
_THROTTLE_WAIT_CAP_S = 10.0  # never sleep longer than this per retry — a huge
# Retry-After is treated as "give up" rather than re-introducing a long hang.


def _redact_url(url: str) -> str:
    """Strip the query string (and fragment) from a Graph URL before logging.

    Query strings carry user content — ``$search="<name>"``, ``$filter`` with
    email addresses, item ids — so only the scheme+host+path is log-safe.
    """
    if not url:
        return url
    for sep in ("?", "#"):
        idx = url.find(sep)
        if idx != -1:
            url = url[:idx]
    return url


def _parse_retry_after(headers, attempt: int) -> float:
    """Seconds to wait before retrying a 429. Honors the ``Retry-After`` header
    (delta-seconds form) when present and sane; otherwise falls back to bounded
    exponential backoff. Always clamped to ``_THROTTLE_WAIT_CAP_S``."""
    raw = None
    try:
        raw = headers.get("Retry-After") if headers else None
    except Exception:
        raw = None
    if raw is not None:
        try:
            return min(max(float(raw), 0.0), _THROTTLE_WAIT_CAP_S)
        except (ValueError, TypeError):
            pass
    return min(0.5 * (2 ** attempt), _THROTTLE_WAIT_CAP_S)


class AuthExpiredError(RuntimeError):
    """Raised when Microsoft Graph rejects the session and it cannot be
    recovered by refreshing — i.e. the user must re-authenticate.

    Surfacing this as an exception (instead of letting handlers silently
    return empty results) lets the MCP layer tell the client exactly what is
    wrong and how to fix it: run ``office-connect login``.
    """


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
        self.privacy_settings = OfficePrivacyConfig()
        # Serializes refresh attempts so concurrent tool calls don't burn
        # multiple refresh_tokens at once (Microsoft rotates them).
        self._refresh_lock = asyncio.Lock()
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
            await self.set_access_token_async(None)
            return None
        logger.info("[OP] token_refresh — done (%.0fms, pid=%d)", _elapsed, _pid)
        await self.set_access_token_async(result["access_token"])
        await self.set_refresh_token_async(result.get("refresh_token"))
        return result["access_token"]

    async def get_access_token_async(self):
        """Return a valid access token, refreshing proactively if it's about
        to expire. Falls back to the parent implementation on a non-refreshable
        instance or when no refresh token is available."""
        token = await super().get_access_token_async()
        if (
            not token
            or not self.can_refresh
            or self.time_until_token_expiration(token) > self.min_expiry
        ):
            return token

        async with self._refresh_lock:
            # Re-check inside the lock — a concurrent caller may already have
            # refreshed.
            current = self.cache_dict.get("access_token")
            if current and self.time_until_token_expiration(current) > self.min_expiry:
                return current
            refreshed = await self.refresh_token_async()
            return refreshed or token

    async def run_async(self, *, url, method="GET", json=None, data=None, token=None, add_headers=None):
        """Send a Graph request. Retries once after a refresh if the server
        rejects the cached token with 401.

        Raises :class:`AuthExpiredError` when a 401 cannot be recovered — no
        refresh capability, the refresh token is gone/invalid, or Graph still
        rejects the freshly-refreshed token. Raising (instead of returning the
        401 for handlers to quietly turn into empty results) is what lets the
        MCP layer report a clear "re-authenticate" message to the client.
        """
        if token is None:
            token = await self.get_access_token_async()

        async def _send(current_token):
            kwargs = {
                "url": url,
                "method": method,
                "json": json,
                "token": current_token,
                "add_headers": add_headers,
            }
            if data is not None:
                kwargs["data"] = data
            # Log every physical Graph request with status + latency (URL
            # redacted to scheme+host+path). Without this, a failed search left
            # NO trace — non-200s are swallowed into empty results upstream and a
            # timeout only surfaced as a generic error, so every diagnosis was a
            # guess ("must be the known mail timeout"). Now the log states facts.
            t0 = _time.monotonic()
            try:
                resp = await super(MsGraphInstance, self).run_async(**kwargs)
            except (asyncio.TimeoutError, TimeoutError):
                logger.warning(
                    "[OP] %s %s -> TIMEOUT after %dms",
                    method, _redact_url(url), int((_time.monotonic() - t0) * 1000),
                )
                raise
            logger.info(
                "[OP] %s %s -> %s (%dms)",
                method, _redact_url(url), getattr(resp, "status_code", None),
                int((_time.monotonic() - t0) * 1000),
            )
            return resp

        response = await _send(token)

        # Absorb transient hiccups (a dropped connection during a token-refresh
        # race, or a brief Graph 5xx) with one short retry, so the first call
        # after a cold start doesn't surface a spurious empty result.
        if response is None or getattr(response, "status_code", 0) in (502, 503, 504):
            await asyncio.sleep(0.1)
            if token is None:
                token = await self.get_access_token_async()
            response = await _send(token)

        # Throttling: Graph returns 429 under intensive/bursty use (notably
        # /search/query). Honor Retry-After with a few bounded retries so an
        # intensive search recovers instead of surfacing a misleading failure.
        attempt = 0
        while (
            response is not None
            and getattr(response, "status_code", 0) == 429
            and attempt < _MAX_THROTTLE_RETRIES
        ):
            wait = _parse_retry_after(getattr(response, "headers", None), attempt)
            logger.warning(
                "[OP] run_async — throttled (429) on %s, retry %d/%d after %.1fs",
                _redact_url(url), attempt + 1, _MAX_THROTTLE_RETRIES, wait,
            )
            await asyncio.sleep(wait)
            response = await _send(token)
            attempt += 1

        if response is None or getattr(response, "status_code", 0) != 401:
            return response

        # Graph rejected the token. Attempt one refresh-and-retry if we can;
        # otherwise the session is unrecoverable and the user must re-auth.
        if not self.can_refresh:
            raise AuthExpiredError(
                "Microsoft Graph rejected the access token (HTTP 401) and this "
                "session is configured without refresh capability."
            )

        async with self._refresh_lock:
            new_token = await self.refresh_token_async()
        if not new_token or new_token == token:
            raise AuthExpiredError(
                "Microsoft Graph rejected the access token (HTTP 401) and it "
                "could not be refreshed — the refresh token is missing, "
                "expired, or has been revoked."
            )

        logger.info("[OP] run_async — refreshed after 401, retrying %s", _redact_url(url))
        retry = await _send(new_token)
        if retry is not None and getattr(retry, "status_code", 0) == 401:
            raise AuthExpiredError(
                "Microsoft Graph returned HTTP 401 even after a successful "
                "token refresh — the account may have revoked access or the "
                "app's permissions/consent changed."
            )
        return retry

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
        # Only attach a secret when one is actually set — public-client flows
        # (e.g. tokens minted via the device-code login) refresh without one,
        # and Azure AD rejects requests that include an empty secret.
        if self.client_secret:
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

    def get_mail_folders(self):
        return MailFolderHandler(self)

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

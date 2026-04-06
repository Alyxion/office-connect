from __future__ import annotations

import asyncio
import os
import hashlib
import time
import uuid
import base64
from threading import RLock
from typing import Optional, TYPE_CHECKING

import jwt
import aiohttp
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from office_con.db_helpers import get_async_redis_client

if TYPE_CHECKING:
    from redis.asyncio import Redis
    from office_con.msgraph.profile_handler import UserProfile

REFRESH_TOKEN_ID = "refresh_token"
ACCESS_TOKEN_ID = "access_token"

class WebUserInstance:
    """Base user session — holds tokens, Redis/MongoDB handles, and user identity."""
    FEATURE_MAIL = "mail"
    FEATURE_CALENDAR = "calendar"
    FEATURE_PROFILE = "profile"


    def __init__(self, application: str, cache_dict: Optional[dict] = None,
                 redis_url: str | None = None, mongodb_url: str | None = None, session_id: str | None = None):
        self.app = application
        self.access_lock = RLock()
        self._redis_client_async = None
        self._redis_url = redis_url
        self._mongodb_url: str | None = mongodb_url
        self._mongodb_client_async = None
        self.redis_path = f"{self.app}:token_cache_v2:"
        self.user_path: str | None = None
        self.redis_lifetime = 7 * 24 * 60 * 60  # 7 days
        self.session_id = session_id
        self.cache_dict = cache_dict if cache_dict is not None else {}
        self.min_expiry = 15 * 60  # 15 minutes
        self.user_id: str | None = None
        self.given_name: str  = "Anonymous"
        self.surname: str | None = None
        self.full_name: str | None = None
        self.email: str | None = None
        self._encryption_key: Fernet | None = None
        self._redis_key_hash = None
        self._salt = os.environ.get("O365_SALT", "")
        if not self._salt:
            import logging as _logging
            _logging.getLogger(__name__).warning("[AUTH] O365_SALT not set — using fallback salt. Set O365_SALT for production security.")
            self._salt = "nice_office_salt"
        self.me: UserProfile | None = None
        self.features: set[str] = set()
        """Set of features this user instance has access to"""
        self._bg_task: "asyncio.Task | None" = None
        self._mock_transport = None  # Set by enable_mock() for testing
        self._mock_profile = None

    async def register_with_background_service(self):
        """Register this user instance with the background service registry and set up notifications."""
        import asyncio
        from office_con.auth.background_service_registry import BackgroundServiceRegistry
        registry = BackgroundServiceRegistry.instance()
        await registry.notify_login(self)

        async def _loop():
            try:
                while True:
                    await asyncio.sleep(5.0)
                    await registry.notify_loop(self)
            except asyncio.CancelledError:
                return

        # Cancel any previous task to avoid pile-up
        if self._bg_task and not self._bg_task.done():
            self._bg_task.cancel()
        self._bg_task = asyncio.get_running_loop().create_task(_loop())

    async def redis_client_async(self) -> Optional["Redis"]:
        """Lazy initialization of Redis client"""
        if self._redis_client_async is None and self._redis_url:
            try:
                client = await get_async_redis_client(self._redis_url)
                self._redis_client_async = client
                if client is not None:
                    await client.ping()  # Test connection
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning("[AUTH] redis_client_async — connection failed: %s", e)
                self._redis_client_async = None
        return self._redis_client_async

    async def mongo_client_async(self) -> object | None:
        """Return the shared async MongoDB client (from mongo_helpers cache)."""
        if self._mongodb_client_async is None and self._mongodb_url:
            try:
                from office_con.db_helpers import get_async_mongo_client
                self._mongodb_client_async = get_async_mongo_client(self._mongodb_url)
            except Exception:
                self._mongodb_client_async = None
        return self._mongodb_client_async

    def _derive_encryption_key(self, client_id: str) -> bytes:
        """Derive an encryption key from the client_id."""
        # Use PBKDF2 to derive a key from the client_id
        salt = self._salt.encode()  # A fixed salt
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=salt,
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(client_id.encode()))
        return key

    def _get_encryption_key(self) -> Fernet:
        """Get or create the encryption key for this user instance."""
        if self._encryption_key is None:
            # Use session_id as the client_id, or generate a new one
            client_id = self.session_id or str(uuid.uuid4())
            key = self._derive_encryption_key(client_id)
            self._encryption_key = Fernet(key)
        return self._encryption_key

    def _encrypt(self, data: str) -> bytes:
        """Encrypt data using the user's encryption key."""
        if not data:
            return b''
        return self._get_encryption_key().encrypt(data.encode())

    def _decrypt(self, data: bytes) -> str:
        """Decrypt data using the user's encryption key."""
        if not data:
            return ''
        try:
            return self._get_encryption_key().decrypt(data).decode()
        except Exception as e:
            import logging
            logging.getLogger(__name__).warning("[AUTH] _decrypt failed (session_id=%s): %s", self.session_id[:8] if self.session_id else "?", e)
            return ''

    def _hash_for_redis_key(self, user_id: str) -> str:
        """Create a hash of the user_id for use in Redis keys."""
        # Create a hash that's safe for Redis keys
        return hashlib.sha256(user_id.encode()).hexdigest()

    def ensure_unique_user_id(self):
        with self.access_lock:
            if "sc_user" not in self.cache_dict:
                if self.session_id is not None:
                    self.cache_dict["sc_user"] = self.session_id
                else:
                    self.cache_dict["sc_user"] = str(uuid.uuid4()) + "-" + str(uuid.uuid4())
            
            # Store the original user ID
            original_user_id = self.cache_dict["sc_user"]
            
            # Create a hash for Redis key
            self._redis_key_hash = self._hash_for_redis_key(original_user_id)
            
            # Set the user path with the hashed ID
            self.user_path = self.redis_path + self._redis_key_hash + ":"
            
            return original_user_id

    PROFILE_CACHE_ID = "profile_cache"

    async def cache_profile_to_redis_async(self):
        """Store profile fields (email, user_id, names, phones) in Redis for fast restore."""
        if not self.email or not (redis_client := await self.redis_client_async()) or not self.user_path:
            return
        import json as _json
        data = {
            "email": self.email,
            "user_id": self.user_id,
            "given_name": self.given_name,
            "surname": self.surname,
            "full_name": self.full_name,
        }
        me = getattr(self, "me", None)
        if me and hasattr(me, "business_phones"):
            data["business_phones"] = me.business_phones or []
            data["mail"] = me.mail
            data["user_principal_name"] = getattr(me, "user_principal_name", "")
        try:
            await redis_client.set(self.user_path + self.PROFILE_CACHE_ID, _json.dumps(data), ex=self.redis_lifetime)
        except Exception:
            pass

    async def restore_profile_from_redis_async(self) -> bool:
        """Restore cached profile fields from Redis. Returns True if restored."""
        import logging as _logging
        _log = _logging.getLogger(__name__)
        if self.email:
            return True
        if not (redis_client := await self.redis_client_async()) or not self.user_path:
            _log.info("[AUTH] restore_profile — no redis or no user_path")
            return False
        try:
            import json as _json
            data = await redis_client.get(self.user_path + self.PROFILE_CACHE_ID)
            if data:
                profile = _json.loads(data)
                self.email = profile.get("email")
                self.user_id = profile.get("user_id")
                self.given_name = profile.get("given_name", "Anonymous")
                self.surname = profile.get("surname")
                self.full_name = profile.get("full_name")
                if hasattr(self, "me") and self.user_id:
                    from office_con.msgraph.profile_handler import UserProfile
                    self.me = UserProfile(
                        id=self.user_id,
                        mail=profile.get("mail", self.email),
                        givenName=self.given_name,
                        surname=self.surname or "",
                        displayName=self.full_name or self.given_name or "",
                        businessPhones=profile.get("business_phones", []),
                        userPrincipalName=profile.get("user_principal_name", self.email or ""),
                    )
                return bool(self.email)
        except Exception:
            pass
        return False

    async def get_access_token_async(self):
        import logging as _logging
        _log = _logging.getLogger(__name__)
        # First try memory cache
        with self.access_lock:
            if self.cache_dict.get("access_token", None) is not None:
                return self.cache_dict["access_token"]

        # Then try Redis if available
        redis_client = await self.redis_client_async()
        if not redis_client or not self.user_path:
            _log.warning("[AUTH] get_access_token_async — no redis (%s) or no user_path (%s)", bool(redis_client), bool(self.user_path))
            return None
        try:
            redis_key = self.user_path + ACCESS_TOKEN_ID
            encrypted_token = await redis_client.get(redis_key)
            if encrypted_token is not None:
                # Decrypt the token
                decrypted_token = self._decrypt(encrypted_token)
                if not decrypted_token:
                    _log.warning("[AUTH] get_access_token_async — decryption returned empty (key=%s, session=%s, data_len=%d)",
                                 redis_key, self.session_id[:8] if self.session_id else "?", len(encrypted_token))
                    return None
                with self.access_lock:
                    self.cache_dict["access_token"] = decrypted_token
                return decrypted_token
            else:
                _log.warning("[AUTH] get_access_token_async — no token in Redis (key=%s)", redis_key)
        except Exception as e:
            _log.warning("[AUTH] get_access_token_async — Redis error: %s", e)
            return None
        return None

    async def set_access_token_async(self, access_token: str | None):
        # Update memory cache first (matches sync version)
        with self.access_lock:
            self.cache_dict["access_token"] = access_token
        if not (redis_client := await self.redis_client_async()) or not self.user_path:
            return
        try:
            if access_token is None:
                await redis_client.delete(self.user_path + ACCESS_TOKEN_ID)
            else:
                # Encrypt the token before storing
                encrypted_token = self._encrypt(access_token)
                await redis_client.set(self.user_path + ACCESS_TOKEN_ID, encrypted_token, ex=self.redis_lifetime)
        except Exception:
            pass

    async def get_refresh_token_async(self):
        import logging as _logging
        _log = _logging.getLogger(__name__)
        # First try memory cache
        with self.access_lock:
            if self.cache_dict.get("refresh_token", None) is not None:
                _log.info("[AUTH] get_refresh_token_async — found in memory cache")
                return self.cache_dict["refresh_token"]
        # Then try Redis if available
        redis_client = await self.redis_client_async()
        if not redis_client:
            _log.warning("[AUTH] get_refresh_token_async — no redis client (url=%s)", bool(self._redis_url))
            return None
        if not self.user_path:
            _log.warning("[AUTH] get_refresh_token_async — no user_path (session_id=%s)", self.session_id[:8] if self.session_id else "None")
            return None
        redis_key = self.user_path + REFRESH_TOKEN_ID
        try:
            encrypted_token = await redis_client.get(redis_key)
            if encrypted_token is not None:
                decrypted = self._decrypt(encrypted_token)
                if not decrypted:
                    _log.warning("[AUTH] get_refresh_token_async — decryption failed (key=%s, session=%s, data_len=%d)",
                                 redis_key, self.session_id[:8] if self.session_id else "?", len(encrypted_token))
                    return None
                _log.info("[AUTH] get_refresh_token_async — found in Redis (key=%s)", redis_key)
                return decrypted
            else:
                _log.warning("[AUTH] get_refresh_token_async — key not found in Redis (key=%s)", redis_key)
        except Exception as e:
            _log.warning("[AUTH] get_refresh_token_async — Redis error: %s", e)
            return None
        return None

    async def set_refresh_token_async(self, refresh_token: str | None):
        # Update memory cache first (matches sync version)
        with self.access_lock:
            self.cache_dict["refresh_token"] = refresh_token
        if not (redis_client := await self.redis_client_async()) or not self.user_path:
            return
        try:
            if refresh_token is None:
                await redis_client.delete(self.user_path + REFRESH_TOKEN_ID)
            else:
                # Encrypt the token before storing
                encrypted_token = self._encrypt(refresh_token)
                await redis_client.set(self.user_path + REFRESH_TOKEN_ID, encrypted_token, ex=self.redis_lifetime)
        except Exception as e:
            import logging
            logging.getLogger(__name__).warning("[AUTH] set_refresh_token_async — error: %s", e)

    async def logout_async(self):
        """Removes all tokens and user data from memory and Redis."""
        await self.set_access_token_async(None)
        await self.set_refresh_token_async(None)
        # Also clear cached profile
        if (redis_client := await self.redis_client_async()) and self.user_path:
            try:
                await redis_client.delete(self.user_path + self.PROFILE_CACHE_ID)
            except Exception:
                pass

    async def run_async(self, *, url: str, method: str = "GET", json: dict[str, object] | None = None, token: Optional[str] = None, add_headers: dict[str, str] | None = None) -> object | None:
        """Async HTTP helper using aiohttp.

        Returns a response object that mimics the requests library response
        to maintain compatibility with code expecting requests.Response objects.
        """
        # Mock transport intercept — returns synthetic responses without HTTP
        if self._mock_transport is not None:
            return await self._mock_transport.handle_request(url, method, json)

        if token is None:
            token = await self.get_access_token_async()
        if token is None:
            return None
        headers = {
            "Authorization": f"Bearer {token}"
        }
        if add_headers:
            headers.update(add_headers)
        
        class AsyncResponseWrapper:
            """Wrapper class to make aiohttp response compatible with requests response"""
            def __init__(self, status, content, text, headers, url):
                self.status_code = status
                self.content = content
                self.text = text
                self.headers = headers
                self.url = url
                
            def json(self):
                import json
                return json.loads(self.text)
                
            def raise_for_status(self):
                if self.status_code >= 400:
                    raise Exception(f"HTTP Error: {self.status_code}")
        
        async with aiohttp.ClientSession() as session:
            if method == "POST":
                async with session.post(url, headers=headers, json=json) as response:
                    content = await response.read()
                    text = await response.text()
                    return AsyncResponseWrapper(
                        response.status, content, text, response.headers, response.url
                    )
            elif method == "PATCH":
                async with session.patch(url, headers=headers, json=json) as response:
                    content = await response.read()
                    text = await response.text()
                    return AsyncResponseWrapper(
                        response.status, content, text, response.headers, response.url
                    )
            elif method == "DELETE":
                async with session.delete(url, headers=headers, json=json) as response:
                    content = await response.read()
                    text = await response.text()
                    return AsyncResponseWrapper(
                        response.status, content, text, response.headers, response.url
                    )
            else:
                async with session.get(url, headers=headers, json=json) as response:
                    content = await response.read()
                    try:
                        text = await response.text()
                    except Exception:
                        text = None
                    return AsyncResponseWrapper(
                        response.status, content, text, response.headers, response.url
                    )

    async def set_user_str_async(self, key: str, value: str, timeout: int = 60 * 30):
        if (redis_client := await self.redis_client_async()) and self.user_path:
            try:
                encrypted_value = self._encrypt(value)
                await redis_client.set(self.user_path + key, encrypted_value, ex=timeout)
            except Exception as e:
                import logging
                logging.getLogger(__name__).warning("[AUTH] set_user_str_async(%s) — Redis error: %s", key, e)

    async def get_user_str_async(self, key: str, default: str | None = None) -> str | None:
        import logging as _logging
        _log = _logging.getLogger(__name__)
        redis_client = await self.redis_client_async()
        if not redis_client or not self.user_path:
            _log.warning("[AUTH] get_user_str_async(%s) — no redis client or no user_path", key)
            return default
        redis_key = self.user_path + key
        try:
            encrypted_value = await redis_client.get(redis_key)
            if encrypted_value is not None:
                result = self._decrypt(encrypted_value)
                if not result:
                    _log.warning("[AUTH] get_user_str_async(%s) — found in Redis but decrypt failed (key=%s)", key, redis_key)
                return result
            _log.info("[AUTH] get_user_str_async(%s) — not found in Redis (key=%s)", key, redis_key)
        except Exception as e:
            _log.warning("[AUTH] get_user_str_async(%s) — Redis error: %s", key, e)
        return default

    @staticmethod
    def get_token_expiration(token: str) -> int:
        # Decode the token to extract the expiration time (exp)
        try:
            decoded_token = jwt.decode(token, options={"verify_signature": False})
        except Exception:
            return 0
        expiration_time: int = decoded_token.get("exp", 0)
        return expiration_time

    def is_token_still_valid(self, token: str | None) -> bool:
        if token is None:
            return False
        expiration_time = self.get_token_expiration(token)
        current_time = int(time.time())
        return expiration_time > current_time

    def time_until_token_expiration(self, token: str | None) -> int:
        if token is None:
            return 0
        expiration_time = self.get_token_expiration(token)
        current_time = int(time.time())
        return expiration_time - current_time

    @property
    def identifier(self):
        return self.get_identifier()

    def get_identifier(self):
        if not self.email:
            raise ValueError("User email not available - profile not loaded or authentication incomplete")
        return self.email.lower()

    @property
    def mail(self):
        return self.get_mail()

    def get_mail(self):
        return None

    @property
    def calendar(self):
        return self.get_calendar()

    def get_calendar(self):
        return None

    @property
    def directory(self):
        return self.get_directory()

    def get_directory(self):
        return None

    def close(self):
        import asyncio
        try:
            if self._redis_client_async or self._mongodb_client_async:
                loop = asyncio.get_running_loop()
                if not loop.is_running():
                    raise RuntimeError("No valid event loop found. Shutdown async")
                loop.create_task(self.close_async())
                return
        except RuntimeError:
            pass

    async def close_async(self):
        if self._bg_task and not self._bg_task.done():
            self._bg_task.cancel()
            self._bg_task = None
        if self._redis_client_async:
            await self._redis_client_async.aclose()
            self._redis_client_async = None
        self._mongodb_client_async = None

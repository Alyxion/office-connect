import time
import bcrypt
import asyncio
from office_con.web_user_instance import WebUserInstance

from pydantic import BaseModel, Field
import base64

DB_USER_TOKEN_VERSION = 3

class DBUserToken(BaseModel):
    """JWT-like token payload for database-authenticated users."""
    version: int = Field(default=DB_USER_TOKEN_VERSION, description="Token schema version")
    user_id: str = Field(default="", description="User UUID")
    given_name: str = Field(default="", description="User's given / first name")
    surname: str = Field(default="", description="User's surname / last name")
    exp: float = Field(default=0.0, description="Expiry timestamp (epoch seconds)")
    issued_at: float = Field(default=0.0, description="Issued-at timestamp (epoch seconds)")
    secret_hash: str = Field(default="", description="SHA-256 hash of the password secret")
    max_duration: float | None = Field(default=None, description="Maximum token lifetime in seconds, None for default")

def _hash_secret(secret: str) -> str:
    """One-way hash of password_secret for embedding in tokens (never store the raw secret)."""
    import hashlib
    return hashlib.sha256(secret.encode("utf-8")).hexdigest()


class DBUserInstance(WebUserInstance):
    """WebUserInstance backed by MongoDB credentials instead of OAuth."""
    TOKEN_DURATION_SECONDS = 60 * 60  # 60 minutes

    PASSWORD_HASH_FIELD = "pw_hash"

    def __init__(self, app: str, cache_dict: dict | None = None, redis_url: str | None = None, 
                       mongodb_url: str | None = None, session_id: str | None = None, 
                       user_db: str | None = None,
                       user_collection: str | None = None,
                       user_field: str | None = None):
        super().__init__(application=app, cache_dict=cache_dict, redis_url=redis_url, 
                         mongodb_url=mongodb_url, session_id=session_id)
        self._db_identifier: str | None = None
        self._user_db = user_db
        self._user_collection = user_collection
        self._user_field = user_field or "email"
        self.ensure_unique_user_id()

    async def restore_from_token_async(self) -> bool:
        token = await self.get_access_token_async()
        if not token:
            return False
        token_data = await self.get_token_data_async(token)
        if not token_data:
            return False
        # check if valid
        self._db_identifier = token_data.user_id
        self.given_name = token_data.given_name
        self.surname = token_data.surname
        self.full_name = f"{self.given_name} {self.surname}" if self.given_name and self.surname else self._db_identifier
        if self.is_token_still_valid(token):
            return True
        return False

    async def _get_password_hash_and_secret_async(self) -> tuple[bytes | None, str]:
        client = await self.mongo_client_async()
        if not client:
            return None, ""
        db = client[self._user_db]
        collection = db[self._user_collection]
        user = await collection.find_one(
            {self._user_field: self._db_identifier},
            {self.PASSWORD_HASH_FIELD: 1, 'password_secret': 1, '_id': 0}
        )
        if not user:
            return None, ""
        return user.get(self.PASSWORD_HASH_FIELD, None), user.get("password_secret", "")

    def _estimate_name(self) -> tuple[str, str]:
        if self.surname:
            return self.given_name, self.surname
        if self._db_identifier and "@" in self._db_identifier:
            name_part = self._db_identifier.split("@", 1)[0]
            if "." in name_part:
                first_name, last_name = name_part.split(".", 1)
                first_name = first_name.capitalize()
                last_name = last_name.capitalize()
                return first_name, last_name
        return self._db_identifier or "", ""

    async def login_async(self, *, password: str | None = None, secret: str | None = None, max_duration: float | None = None, user_id: str | None = None) -> bool:
        # Auth by password or secret (async)
        if user_id is not None:
            self._db_identifier = user_id
        if max_duration is None:
            max_duration = self.TOKEN_DURATION_SECONDS
        password_hash, password_secret = await self._get_password_hash_and_secret_async()
        if password_hash is None or password_secret is None:
            return False
        authenticated = False
        if password and password_hash:
            if await asyncio.to_thread(bcrypt.checkpw, password.encode('utf-8'), password_hash):
                authenticated = True
        elif secret and password_secret:
            import hmac as _hmac
            if _hmac.compare_digest(secret, password_secret):
                authenticated = True
        if not authenticated:
            return False
        # Enforce max_duration
        now = time.time()
        exp = now + max_duration
        if max_duration is not None:
            self.max_duration = max_duration
            max_exp = now + max_duration
            if exp > max_exp:
                exp = max_exp
        self.given_name, self.surname = self._estimate_name()
        self.full_name = f"{self.given_name} {self.surname}" if self.given_name and self.surname else self._db_identifier
        token_model = DBUserToken(
            user_id=self._db_identifier or "",
            given_name=self.given_name,
            surname=self.surname,
            exp=exp,
            issued_at=now,
            secret_hash=_hash_secret(password_secret),
            max_duration=max_duration,
        )
        token_str = self.encode_token(token_model)
        await self.set_access_token_async(token_str)
        await self.set_refresh_token_async(None)
        return True

    async def set_password_async(self, *, password: str | None = None, user_id: str | None = None) -> bool:
        import uuid
        client = await self.mongo_client_async()
        if not client:
            return False
        db = client[self._user_db]
        collection = db[self._user_collection]
        if user_id is not None:
            self._db_identifier = user_id
        username = self._db_identifier
        salt = bcrypt.gensalt()
        hashed_password = await asyncio.to_thread(bcrypt.hashpw, (password or "").encode('utf-8'), salt)
        password_secret = uuid.uuid4().hex
        # remove password if empty
        if not password:
            update_doc: dict[str, str | bytes] = {"password_secret": password_secret}
            # delete old password field if set
            await collection.update_one({self._user_field: username}, {"$unset": {self.PASSWORD_HASH_FIELD: 1}})
            await collection.update_one({self._user_field: username}, {"$set": update_doc}, upsert=True)
            return True
        else:
            update_doc = {
                self.PASSWORD_HASH_FIELD: hashed_password,
                "password_secret": password_secret
            }
            await collection.update_one({self._user_field: username}, {"$set": update_doc}, upsert=True)

        return True

    def is_token_still_valid(self, token: str | None) -> bool:
        if not token:
            return False
        token_model = self.decode_token(token)
        if not token_model or token_model.user_id != self._db_identifier:
            return False
        return time.time() < token_model.exp and token_model.version == DB_USER_TOKEN_VERSION

    def time_until_token_expiration(self, token: str | None) -> int:
        token_model = self.decode_token(token) if token else None
        if not token_model or token_model.user_id != self._db_identifier:
            return 0
        return int(token_model.exp - time.time())

    async def refresh_async(self) -> str | None:
        token = await self.get_access_token_async()
        if self.is_token_still_valid(token):
            time_left = self.time_until_token_expiration(token)
            if time_left > self.min_expiry:
                return token
        new_token = await self.refresh_token_async()
        if new_token and self.is_token_still_valid(new_token):
            return new_token
        return None

    async def refresh_token_async(self) -> str | None:
        success = await self.login_async()
        if success:
            return await self.get_access_token_async()
        return None

    def get_identifier(self) -> str:
        return self._db_identifier or ""

    def build_auth_url(self, auth_url, authority=None):
        raise ValueError("DBUserInstance does not dynamic login")

    async def get_token_data_async(self, token: str | None = None) -> DBUserToken | None:
        """
        Returns the current access token as a DBUserToken instance, or None if invalid or not present.
        """
        if not token:
            token = await self.get_access_token_async()
        if not token:
            return None
        try:
            return self.decode_token(token)
        except Exception:
            return None

    @staticmethod
    def _token_hmac_key() -> bytes:
        import os
        key = os.environ.get("O365_TOKEN_SECRET", os.environ.get("O365_SALT", ""))
        if not key:
            raise ValueError("O365_TOKEN_SECRET or O365_SALT must be set for token signing")
        return key.encode("utf-8")

    @staticmethod
    def encode_token(token_model: DBUserToken) -> str:
        import hmac
        import hashlib
        json_bytes = token_model.model_dump_json().encode("utf-8")
        sig = hmac.new(DBUserInstance._token_hmac_key(), json_bytes, hashlib.sha256).digest()
        payload = sig + json_bytes  # 32-byte signature prefix
        return base64.urlsafe_b64encode(payload).decode("utf-8")

    @staticmethod
    def decode_token(token: str) -> DBUserToken | None:
        import hmac
        import hashlib
        try:
            payload = base64.urlsafe_b64decode(token.encode("utf-8"))
            if len(payload) < 33:
                return None
            sig, json_bytes = payload[:32], payload[32:]
            expected = hmac.new(DBUserInstance._token_hmac_key(), json_bytes, hashlib.sha256).digest()
            if not hmac.compare_digest(sig, expected):
                return None
            return DBUserToken.model_validate_json(json_bytes)
        except Exception:
            return None


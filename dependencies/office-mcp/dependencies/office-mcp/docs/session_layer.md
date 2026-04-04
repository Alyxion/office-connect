# Session Layer Reference

The session layer is the foundation of `office-mcp`. It manages user
authentication state, OAuth token storage and retrieval, encrypted Redis
caching, and database-backed credential verification. Every higher-level
handler (mail, calendar, teams, etc.) ultimately delegates to a session
instance for its access tokens and user identity.

The layer comprises three modules:

- [WebUserInstance](#webuserinstance) -- OAuth-based session with Redis-backed token
  cache and Fernet encryption.
- [DBUserInstance](#dbuserinstance) -- Database-authenticated session using
  password/secret credentials, HMAC-signed custom tokens, and bcrypt
  password hashing.
- [\_db\_helpers](#_db_helpers----connection-factories) -- Thread-safe connection factories for MongoDB and
  Redis (both sync and async).

---

## `WebUserInstance`

**Module:** `office_mcp.web_user_instance`

`WebUserInstance` is the primary session class. It holds a single user's
OAuth tokens (access and refresh), manages encrypted persistence to Redis,
and exposes helper methods for profile caching, arbitrary key-value storage,
and authenticated HTTP requests against Microsoft Graph or other APIs.

### Class Constants

| Constant | Description |
| --- | --- |
| `FEATURE_MAIL` | Feature flag string `"mail"` -- indicates the user has mail access. |
| `FEATURE_CALENDAR` | Feature flag string `"calendar"` -- indicates calendar access. |
| `FEATURE_PROFILE` | Feature flag string `"profile"` -- indicates profile access. |
| `REFRESH_TOKEN_ID` | Module-level constant `"refresh_token"` -- Redis sub-key for the encrypted refresh token. |
| `ACCESS_TOKEN_ID` | Module-level constant `"access_token"` -- Redis sub-key for the encrypted access token. |
| `PROFILE_CACHE_ID` | Class constant `"profile_cache"` -- Redis sub-key for cached user profile JSON. |

### Constructor

```python
WebUserInstance(
    application: str,
    cache_dict: dict | None = None,
    redis_url: str | None = None,
    mongodb_url: str | None = None,
    session_id: str | None = None,
)
```

| Parameter | Type | Description |
| --- | --- | --- |
| `application` | `str` | Application identifier used as the Redis key namespace prefix (`{app}:token_cache_v2:`). |
| `cache_dict` | `dict \| None` | In-memory token cache. If `None`, an empty dict is created internally. Tokens are always written here first; Redis is a secondary persistent store. |
| `redis_url` | `str \| None` | Redis connection URL. Supports standalone (`redis://host:6379`) and cluster mode (comma-separated URLs). Pass `None` to disable Redis persistence entirely. |
| `mongodb_url` | `str \| None` | MongoDB connection string. Used by `DBUserInstance` for credential lookups and optionally by subclasses for user data. |
| `session_id` | `str \| None` | Unique session identifier. Used as the encryption key derivation input and as the default `sc_user` value. If `None`, a random UUID pair is generated on first call to `ensure_unique_user_id()`. |

### Instance Attributes

| Attribute | Type | Description |
| --- | --- | --- |
| `app` | `str` | Application name (from `application` parameter). |
| `access_lock` | `RLock` | Reentrant lock protecting concurrent access to `cache_dict`. |
| `redis_path` | `str` | Base Redis key prefix: `"{app}:token_cache_v2:"`. |
| `user_path` | `str \| None` | Full Redis key prefix for this user, set by `ensure_unique_user_id()`. Format: `"{app}:token_cache_v2:{sha256_hash}:"`. |
| `redis_lifetime` | `int` | TTL in seconds for Redis keys. Default: **604800** (7 days). |
| `session_id` | `str \| None` | The session identifier passed at construction. |
| `cache_dict` | `dict` | In-memory cache holding `"access_token"`, `"refresh_token"`, and `"sc_user"` keys. |
| `min_expiry` | `int` | Minimum remaining token lifetime (in seconds) before a refresh is triggered. Default: **900** (15 minutes). |
| `user_id` | `str \| None` | Microsoft Graph user ID (set after profile load). |
| `given_name` | `str` | User's first name. Default: `"Anonymous"`. |
| `surname` | `str \| None` | User's last name. |
| `full_name` | `str \| None` | Display name (typically `"{given_name} {surname}"`). |
| `email` | `str \| None` | User's primary email address. |
| `me` | `UserProfile \| None` | Pydantic model with the user's full Graph profile (lazily populated). |
| `features` | `set[str]` | Set of feature flags this instance has access to (e.g., `{"mail", "calendar"}`). |

### Encryption Subsystem

All tokens and user strings stored in Redis are encrypted at rest using
**Fernet** symmetric encryption. The encryption key is derived from the
`session_id` using PBKDF2-HMAC-SHA256 with 100,000 iterations.

The salt is read from the `O365_SALT` environment variable. If unset, a
hardcoded fallback (`"nice_office_salt"`) is used with a warning logged.
**Always set** `O365_SALT` **in production.**

```text
session_id  --+--> PBKDF2HMAC(SHA256, 100k iterations, O365_SALT) --> Fernet key
              |
Redis value --+--> Fernet.encrypt(plaintext) --> ciphertext (stored)
              +--> Fernet.decrypt(ciphertext) --> plaintext (retrieved)
```

Redis keys themselves use a SHA-256 hash of the user ID so that the
original user identifier is never stored in Redis key names.

### Public Methods

#### `ensure_unique_user_id()`

```python
def ensure_unique_user_id(self) -> str
```

Assigns a stable user identity for Redis key construction. If
`cache_dict["sc_user"]` is already set, it is reused; otherwise the
`session_id` is used, or a random UUID pair is generated.

Sets `self.user_path` to the full Redis key prefix including a SHA-256
hash of the user ID. Returns the original (unhashed) user ID string.

This method is **thread-safe** (guarded by `access_lock`).

#### `redis_client_async()`

```python
async def redis_client_async(self) -> Redis | None
```

Returns a lazily initialized async Redis client. On first call, connects to
`self._redis_url` using `get_async_redis_client()` and verifies with
a `PING`. Returns `None` if no URL was configured or if the connection
fails.

#### `mongo_client_async()`

```python
async def mongo_client_async(self) -> AsyncMongoClient | None
```

Returns a lazily initialized async MongoDB client via
`get_async_mongo_client()`. Returns `None` if no URL was configured
or if the connection fails.

#### `get_access_token_async()`

```python
async def get_access_token_async(self) -> str | None
```

Retrieves the current access token. Checks the in-memory `cache_dict`
first; falls back to Redis (decrypting the stored value). Returns `None`
if no token is available.

#### `set_access_token_async(access_token)`

```python
async def set_access_token_async(self, access_token: str | None) -> None
```

Stores or clears the access token. Updates `cache_dict` immediately,
then persists to Redis (encrypted) with the configured `redis_lifetime`
TTL. Passing `None` deletes the Redis key.

#### `get_refresh_token_async()`

```python
async def get_refresh_token_async(self) -> str | None
```

Retrieves the refresh token with the same memory-then-Redis strategy as
`get_access_token_async()`.

#### `set_refresh_token_async(refresh_token)`

```python
async def set_refresh_token_async(self, refresh_token: str | None) -> None
```

Stores or clears the refresh token, following the same pattern as
`set_access_token_async()`.

#### `cache_profile_to_redis_async()`

```python
async def cache_profile_to_redis_async(self) -> None
```

Serializes the user's profile fields (`email`, `user_id`,
`given_name`, `surname`, `full_name`, and optional `UserProfile`
fields such as `business_phones` and `user_principal_name`) to JSON
and stores them in Redis under the `profile_cache` sub-key. Requires
that `self.email` and `self.user_path` are set.

#### `restore_profile_from_redis_async()`

```python
async def restore_profile_from_redis_async(self) -> bool
```

Restores the user's profile fields from the Redis cache. If the profile
already exists in memory (`self.email` is set), returns `True`
immediately. On successful restore, also reconstructs a `UserProfile`
Pydantic model on `self.me`. Returns `True` if a profile was restored,
`False` otherwise.

#### `set_user_str_async(key, value, timeout)`

```python
async def set_user_str_async(
    self, key: str, value: str, timeout: int = 1800
) -> None
```

Stores an arbitrary encrypted string in Redis under the user's namespace.
The key is appended to `user_path`. Default TTL is 30 minutes.

#### `get_user_str_async(key, default)`

```python
async def get_user_str_async(
    self, key: str, default: str | None = None
) -> str | None
```

Retrieves and decrypts an arbitrary user string from Redis. Returns
`default` if the key is not found or if decryption fails.

#### `get_token_expiration(token)` *(staticmethod)*

```python
@staticmethod
def get_token_expiration(token: str) -> int
```

Decodes a JWT token **without verifying the signature** and returns the
`exp` claim as a Unix timestamp. Returns `0` if decoding fails.

#### `is_token_still_valid(token)`

```python
def is_token_still_valid(self, token: str | None) -> bool
```

Returns `True` if the token's `exp` claim is in the future.
Returns `False` for `None` or expired tokens.

#### `time_until_token_expiration(token)`

```python
def time_until_token_expiration(self, token: str | None) -> int
```

Returns the number of seconds remaining until the token expires. Returns
`0` for `None` tokens. May return a negative value for already-expired
tokens.

#### `run_async(url, method, json, token, add_headers)`

```python
async def run_async(
    self,
    *,
    url: str,
    method: str = "GET",
    json: dict | None = None,
    token: str | None = None,
    add_headers: dict | None = None,
) -> AsyncResponseWrapper | None
```

General-purpose async HTTP client. Automatically attaches the user's
`Bearer` access token. Returns an `AsyncResponseWrapper` object that
mimics the `requests.Response` interface (`status_code`, `text`,
`content`, `json()`, `raise_for_status()`).

Supported methods: `GET`, `POST`, `PATCH`, `DELETE`.

Returns `None` if no token is available.

#### `register_with_background_service()`

```python
async def register_with_background_service(self) -> None
```

Registers this user instance with the singleton
`BackgroundServiceRegistry`, fires the `notify_login` event, and starts
a background `asyncio.Task` that calls `notify_loop` every 5 seconds.
If a previous background task exists, it is cancelled first to prevent
pile-up.

#### `logout_async()`

```python
async def logout_async(self) -> None
```

Removes all tokens and the cached profile from both memory and Redis.
Calls `set_access_token_async(None)`, `set_refresh_token_async(None)`,
and deletes the `profile_cache` Redis key.

#### `close()` / `close_async()`

```python
def close(self) -> None
async def close_async(self) -> None
```

Cleans up resources. `close()` is a sync wrapper that schedules
`close_async()` on the running event loop. `close_async()` cancels the
background task, closes the async Redis connection, and releases the
MongoDB client reference.

### Properties

| Property | Type | Description |
| --- | --- | --- |
| `identifier` | `str` | Lowercase email address. Raises `ValueError` if `email` is not set. |
| `mail` | `None` | Placeholder; returns `None`. Overridden by subclasses (e.g., `MsGraphInstance`) to return a mail handler. |
| `calendar` | `None` | Placeholder; returns `None`. Overridden by subclasses to return a calendar handler. |
| `directory` | `None` | Placeholder; returns `None`. Overridden by subclasses to return a directory handler. |

---

## `DBUserInstance`

**Module:** `office_mcp.db_user_instance`

`DBUserInstance` extends `WebUserInstance` to provide database-backed
authentication using MongoDB-stored bcrypt password hashes and/or plain
secrets. Instead of OAuth JWTs from Azure AD, it issues its own
HMAC-SHA256-signed tokens encoded as URL-safe Base64 strings.

This class is designed for scenarios where users authenticate with a
username/password or a pre-shared secret rather than through an OAuth
redirect flow.

### Class Constants

| Constant | Description |
| --- | --- |
| `TOKEN_DURATION_SECONDS` | Default token lifetime: **3600** seconds (60 minutes). |
| `PASSWORD_HASH_FIELD` | MongoDB document field name for the bcrypt hash: `"pw_hash"`. |
| `DB_USER_TOKEN_VERSION` | Module-level constant: **3**. Tokens with a different version are rejected by `is_token_still_valid()`. |

### Constructor

```python
DBUserInstance(
    app: str,
    cache_dict: dict | None = None,
    redis_url: str | None = None,
    mongodb_url: str | None = None,
    session_id: str | None = None,
    user_db: str | None = None,
    user_collection: str | None = None,
    user_field: str | None = None,
)
```

| Parameter | Type | Description |
| --- | --- | --- |
| `app` | `str` | Application name (passed to `WebUserInstance` as `application`). |
| `cache_dict` | `dict \| None` | In-memory token cache (inherited). |
| `redis_url` | `str \| None` | Redis connection URL (inherited). |
| `mongodb_url` | `str \| None` | MongoDB connection string for credential storage. |
| `session_id` | `str \| None` | Session identifier (inherited). |
| `user_db` | `str \| None` | MongoDB database name containing the user collection. |
| `user_collection` | `str \| None` | MongoDB collection name containing user documents. |
| `user_field` | `str \| None` | Field name used to look up the user in MongoDB. Default: `"email"`. |

The constructor automatically calls `ensure_unique_user_id()` to
initialize the Redis key path.

### `DBUserToken` Dataclass

```python
class DBUserToken(BaseModel):
    version: int = DB_USER_TOKEN_VERSION
    user_id: str
    given_name: str
    surname: str
    exp: float
    issued_at: float
    secret_hash: str = ""
    max_duration: float | None = None
```

A Pydantic model representing the payload of a `DBUserInstance` token.

| Field | Type | Description |
| --- | --- | --- |
| `version` | `int` | Token schema version. Must match `DB_USER_TOKEN_VERSION` (3) for the token to be considered valid. |
| `user_id` | `str` | The user's identifier (typically an email address). |
| `given_name` | `str` | User's first name (estimated from email if not explicitly set). |
| `surname` | `str` | User's last name (estimated from email if not explicitly set). |
| `exp` | `float` | Expiration time as a Unix timestamp. |
| `issued_at` | `float` | Issuance time as a Unix timestamp. |
| `secret_hash` | `str` | SHA-256 hash of the user's `password_secret` at the time of issuance. Used for token revocation when the secret is rotated. |
| `max_duration` | `float \| None` | Maximum allowed token duration in seconds (optional cap). |

### HMAC Token Signing

`DBUserInstance` does **not** use standard JWTs. Instead it uses a custom
binary format:

```text
token = base64url( HMAC-SHA256(key, json_bytes) || json_bytes )
```

- **Signing key:** Read from the `O365_TOKEN_SECRET` environment variable.
  Falls back to `O365_SALT`. Raises `ValueError` if neither is set.
- **Payload:** The Pydantic `model_dump_json()` serialization of
  `DBUserToken`.
- **Signature:** 32-byte HMAC-SHA256 digest prepended to the JSON payload.
- **Encoding:** URL-safe Base64.

Verification recomputes the HMAC over the extracted JSON bytes and uses
`hmac.compare_digest()` for constant-time comparison.

### Public Methods

#### `login_async(password, secret, max_duration, user_id)`

```python
async def login_async(
    self,
    *,
    password: str | None = None,
    secret: str | None = None,
    max_duration: float | None = None,
    user_id: str | None = None,
) -> bool
```

Authenticates the user against MongoDB credentials.

**Authentication modes:**

1. **Password mode** -- If `password` is provided, the stored bcrypt hash
   (`pw_hash` field) is verified using `bcrypt.checkpw()` (offloaded to
   a thread via `asyncio.to_thread`).
2. **Secret mode** -- If `secret` is provided, it is compared to the
   stored `password_secret` field using `hmac.compare_digest()` for
   constant-time comparison.

On success:

- Estimates the user's name from the email address (splits on `@` and
  `.` in the local part).
- Creates a `DBUserToken` with the configured expiration.
- Encodes and signs the token with HMAC-SHA256.
- Stores the token via `set_access_token_async()`.
- Clears any refresh token (DB auth does not use refresh tokens).

Returns `True` on successful authentication, `False` otherwise.

| Parameter | Type | Description |
| --- | --- | --- |
| `password` | `str \| None` | Plaintext password for bcrypt verification. |
| `secret` | `str \| None` | Pre-shared secret for direct comparison. |
| `max_duration` | `float \| None` | Maximum token lifetime in seconds. Defaults to `TOKEN_DURATION_SECONDS` (3600). |
| `user_id` | `str \| None` | Override the user identifier for this login attempt. |

#### `set_password_async(password, user_id)`

```python
async def set_password_async(
    self,
    *,
    password: str | None = None,
    user_id: str | None = None,
) -> bool
```

Sets or removes a user's password in MongoDB. Always generates a new
random `password_secret` (UUID hex). If `password` is falsy, the
`pw_hash` field is removed (`$unset`) while still updating the secret.
Uses `upsert=True` so the user document is created if it does not exist.

Returns `True` on success, `False` if no MongoDB client is available.

#### `restore_from_token_async()`

```python
async def restore_from_token_async(self) -> bool
```

Attempts to restore the session from a previously stored access token.
Retrieves the token from Redis, decodes it, populates identity fields
(`_db_identifier`, `given_name`, `surname`, `full_name`), and
validates expiration and version. Returns `True` only if the token is
still valid.

#### `refresh_async()`

```python
async def refresh_async(self) -> str | None
```

Checks the current token's validity. If the token has more than
`min_expiry` seconds remaining, returns it directly. Otherwise, calls
`refresh_token_async()` to re-authenticate and issue a new token.
Returns the valid token string or `None`.

#### `refresh_token_async()`

```python
async def refresh_token_async(self) -> str | None
```

Re-authenticates by calling `login_async()` with no explicit credentials
(relies on stored MongoDB credentials). Returns the new access token on
success, `None` on failure.

#### `is_token_still_valid(token)` *(override)*

```python
def is_token_still_valid(self, token: str | None) -> bool
```

Overrides the parent method. Decodes the HMAC-signed token, verifies that
the `user_id` matches the current `_db_identifier`, checks that
`time.time() < exp`, and confirms the token `version` matches
`DB_USER_TOKEN_VERSION`.

#### `time_until_token_expiration(token)` *(override)*

```python
def time_until_token_expiration(self, token: str | None) -> int
```

Overrides the parent method. Returns the number of seconds remaining from
the decoded token's `exp` field. Returns `0` if the token cannot be
decoded or the user ID does not match.

#### `get_token_data_async(token)`

```python
async def get_token_data_async(
    self, token: str | None = None
) -> DBUserToken | None
```

Returns the decoded `DBUserToken` for the given token string. If
`token` is `None`, retrieves it from Redis first. Returns `None` on
any decoding error.

#### `encode_token(token_model)` *(staticmethod)*

```python
@staticmethod
def encode_token(token_model: DBUserToken) -> str
```

Serializes a `DBUserToken` to JSON, prepends a 32-byte HMAC-SHA256
signature, and returns the result as a URL-safe Base64 string.

#### `decode_token(token)` *(staticmethod)*

```python
@staticmethod
def decode_token(token: str) -> DBUserToken | None
```

Decodes a URL-safe Base64 token string, verifies the HMAC-SHA256
signature, and returns the deserialized `DBUserToken`. Returns `None`
if the payload is too short (< 33 bytes), the signature is invalid, or
deserialization fails.

#### `get_identifier()` *(override)*

```python
def get_identifier(self) -> str
```

Returns `self._db_identifier` (the user's database identifier) or an
empty string. Unlike the parent class, this does **not** raise
`ValueError` when the identifier is unset.

#### `build_auth_url(auth_url, authority)`

```python
def build_auth_url(self, auth_url, authority=None)
```

Always raises `ValueError`. `DBUserInstance` does not support OAuth
redirect flows.

---

## `_db_helpers` -- Connection Factories

**Module:** `office_mcp._db_helpers`

Provides thread-safe, cached connection factories for MongoDB and Redis.
Each function returns a **shared singleton client** per URL, preventing
connection pool exhaustion in multi-instance applications.

### MongoDB Functions

#### `get_async_mongo_client(url)`

```python
def get_async_mongo_client(url: str | None = None) -> AsyncMongoClient
```

Returns a shared `pymongo.AsyncMongoClient` for the given URL. If
`url` is `None`, falls back to the `MONGODB_CONNECTION` or
`O365_MONGODB_URL` environment variables. Raises `ValueError` if no
URL is available.

**Connection parameters:**

| Parameter | Value |
| --- | --- |
| `serverSelectionTimeoutMS` | 15000 |
| `connectTimeoutMS` | 10000 |
| `maxPoolSize` | 10 |

#### `get_mongo_client(url)`

```python
def get_mongo_client(url: str | None = None) -> MongoClient
```

Returns a shared sync `pymongo.MongoClient`. Same URL resolution logic.

**Connection parameters:**

| Parameter | Value |
| --- | --- |
| `serverSelectionTimeoutMS` | 5000 |
| `connectTimeoutMS` | 5000 |
| `maxPoolSize` | 10 |

### Redis Functions

#### `get_async_redis_client(url)`

```python
async def get_async_redis_client(url: str) -> Redis | RedisCluster
```

Returns a shared async Redis client. Clients are cached per
`(url, event_loop_id)` tuple so that each event loop gets its own
connection.

**Cluster detection:** If the URL contains a comma (`,`), it is treated
as a Redis Cluster with multiple startup nodes. The scheme of the first
node determines whether SSL is enabled (`rediss://` = SSL). Passwords
and ports are extracted from the parsed URLs.

**Standalone mode:** Uses `redis.asyncio.from_url()` with
`decode_responses=False`.

#### `get_redis_client(url)`

```python
def get_redis_client(url: str) -> Redis | RedisCluster
```

Returns a shared sync Redis client with the same cluster detection logic
as the async variant. Thread-safe via an internal lock.

---

## Environment Variables

The following environment variables are relevant to the session layer:

| Variable | Description | Required |
| --- | --- | --- |
| `O365_SALT` | Salt for PBKDF2 key derivation (Fernet encryption of Redis values). **Must be set in production.** A fallback value is used in development with a logged warning. | Yes (prod) |
| `O365_TOKEN_SECRET` | HMAC signing key for `DBUserToken` tokens. Falls back to `O365_SALT` if not set. At least one of the two must be set for `DBUserInstance` to function. | Yes (prod) |
| `MONGODB_CONNECTION` | MongoDB connection string used by `_db_helpers` when no explicit URL is passed. | Conditional |
| `O365_MONGODB_URL` | Fallback MongoDB connection string (checked after `MONGODB_CONNECTION`). | No |
| `O365_REDIS_URL` | Not read directly by the session layer (callers pass `redis_url` to constructors), but conventionally used by application code to supply the URL. | No |

---

## Usage Examples

### Creating a `WebUserInstance` and Storing Tokens

```python
import asyncio
from office_mcp.web_user_instance import WebUserInstance

async def main():
    user = WebUserInstance(
        application="my-app",
        redis_url="redis://localhost:6379",
        session_id="user-session-abc123",
    )

    # Establish the user's Redis key namespace
    user.ensure_unique_user_id()

    # Store an OAuth access token (encrypted in Redis)
    await user.set_access_token_async("eyJhbGciOiJSUzI1NiIs...")

    # Retrieve it later (from memory or Redis)
    token = await user.get_access_token_async()
    print(f"Token valid: {user.is_token_still_valid(token)}")
    print(f"Expires in: {user.time_until_token_expiration(token)}s")

    # Store and retrieve arbitrary user data
    await user.set_user_str_async("preferred_language", "de", timeout=3600)
    lang = await user.get_user_str_async("preferred_language")

    # Cache the user profile for fast session restore
    user.email = "user@example.com"
    user.given_name = "Max"
    user.surname = "Mustermann"
    user.full_name = "Max Mustermann"
    await user.cache_profile_to_redis_async()

    # On a subsequent request, restore without hitting Graph API
    user2 = WebUserInstance(
        application="my-app",
        redis_url="redis://localhost:6379",
        session_id="user-session-abc123",
    )
    user2.ensure_unique_user_id()
    restored = await user2.restore_profile_from_redis_async()
    if restored:
        print(f"Welcome back, {user2.full_name}")

    # Make an authenticated API call
    response = await user.run_async(
        url="https://graph.microsoft.com/v1.0/me",
        method="GET",
    )
    if response and response.status_code == 200:
        profile = response.json()
        print(profile["displayName"])

    # Clean up
    await user.close_async()

asyncio.run(main())
```

### Authenticating with `DBUserInstance`

```python
import asyncio
from office_mcp.db_user_instance import DBUserInstance

async def main():
    user = DBUserInstance(
        app="my-app",
        redis_url="redis://localhost:6379",
        mongodb_url="mongodb://localhost:27017",
        session_id="session-xyz-789",
        user_db="app_database",
        user_collection="users",
        user_field="email",
    )

    # --- Registration: set a password for a new user ---
    await user.set_password_async(
        password="s3cureP@ssw0rd",
        user_id="alice@example.com",
    )

    # --- Login with password ---
    success = await user.login_async(
        user_id="alice@example.com",
        password="s3cureP@ssw0rd",
        max_duration=7200,  # 2 hour token
    )
    if success:
        print(f"Logged in as {user.get_identifier()}")
        print(f"Name: {user.full_name}")

    # --- Inspect the token ---
    token = await user.get_access_token_async()
    token_data = await user.get_token_data_async(token)
    if token_data:
        print(f"Token expires at: {token_data.exp}")
        print(f"Token version: {token_data.version}")

    # --- Restore session on reconnect ---
    user2 = DBUserInstance(
        app="my-app",
        redis_url="redis://localhost:6379",
        mongodb_url="mongodb://localhost:27017",
        session_id="session-xyz-789",
        user_db="app_database",
        user_collection="users",
        user_field="email",
    )
    user2._db_identifier = "alice@example.com"
    restored = await user2.restore_from_token_async()
    print(f"Session restored: {restored}")

    # --- Refresh expiring token ---
    refreshed_token = await user2.refresh_async()
    if refreshed_token:
        print("Token refreshed successfully")

    # --- Login with secret (for service-to-service auth) ---
    # The secret is stored alongside the password hash in MongoDB
    success = await user.login_async(
        user_id="alice@example.com",
        secret="<the-password_secret-from-mongodb>",
    )

    await user.close_async()

asyncio.run(main())
```

### Using `_db_helpers` Directly

```python
import asyncio
from office_mcp._db_helpers import (
    get_async_mongo_client,
    get_mongo_client,
    get_async_redis_client,
    get_redis_client,
)

# --- Async MongoDB ---
async def query_users():
    client = get_async_mongo_client("mongodb://localhost:27017")
    db = client["app_database"]
    user = await db["users"].find_one({"email": "alice@example.com"})
    print(user)

# --- Sync MongoDB ---
client = get_mongo_client()  # uses MONGODB_CONNECTION env var
db = client["app_database"]
user = db["users"].find_one({"email": "alice@example.com"})

# --- Async Redis ---
async def cache_value():
    redis = await get_async_redis_client("redis://localhost:6379")
    await redis.set("key", "value", ex=300)
    val = await redis.get("key")
    print(val)

# --- Async Redis Cluster ---
async def cluster_example():
    redis = await get_async_redis_client(
        "rediss://pwd@node1:6380,rediss://pwd@node2:6380"
    )
    await redis.set("cluster-key", b"cluster-value")

# --- Sync Redis ---
redis = get_redis_client("redis://localhost:6379")
redis.set("sync-key", "sync-value", ex=60)
```

### Session Lifecycle Diagram

The following diagram shows the typical lifecycle of a `WebUserInstance`
session with Redis persistence:

```text
1. Construction
   WebUserInstance(app, redis_url, session_id)
        |
2. Identity Setup
   ensure_unique_user_id()
     --> cache_dict["sc_user"] = session_id
     --> user_path = "{app}:token_cache_v2:{sha256(session_id)}:"
        |
3. Token Storage (after OAuth callback or DB login)
   set_access_token_async(token)
     --> cache_dict["access_token"] = token
     --> Redis SET {user_path}access_token = Fernet.encrypt(token)
   set_refresh_token_async(refresh)
     --> cache_dict["refresh_token"] = refresh
     --> Redis SET {user_path}refresh_token = Fernet.encrypt(refresh)
        |
4. Profile Caching
   cache_profile_to_redis_async()
     --> Redis SET {user_path}profile_cache = JSON(profile fields)
        |
5. Subsequent Requests
   get_access_token_async()
     --> Check cache_dict first
     --> Fallback: Redis GET + Fernet.decrypt()
        |
6. Session Restore (new process / reconnect)
   WebUserInstance(same app, same redis_url, same session_id)
   ensure_unique_user_id()  --> same user_path
   restore_profile_from_redis_async()  --> restore identity fields
   get_access_token_async()  --> retrieve token from Redis
        |
7. Logout / Cleanup
   logout_async()
     --> Deletes access_token, refresh_token, profile_cache from Redis
   close_async()
     --> Cancels background task
     --> Closes Redis connection
```

---

## Security Considerations

- **Encryption at rest:** All tokens and user strings in Redis are
  encrypted with Fernet (AES-128-CBC + HMAC-SHA256). The key is derived
  via PBKDF2 with 100,000 iterations.

- **Key material:** The `session_id` serves as the encryption key input.
  If an attacker obtains both the Redis data and the `session_id`, they
  can derive the key. Ensure session IDs are treated as secrets.

- **Redis key privacy:** User IDs are SHA-256 hashed before being used in
  Redis key names, preventing user enumeration via key scanning.

- **Token signing:** `DBUserInstance` tokens use HMAC-SHA256 with a
  server-side secret (`O365_TOKEN_SECRET`). The secret must be kept
  confidential and should be rotated periodically.

- **Password storage:** Passwords are hashed with bcrypt (random salt) and
  verified via `bcrypt.checkpw()` in a separate thread to avoid blocking
  the event loop.

- **Secret comparison:** The `password_secret` comparison in
  `login_async()` uses `hmac.compare_digest()` to prevent timing
  attacks.

- **Token versioning:** The `DB_USER_TOKEN_VERSION` constant (currently
  3) allows old tokens to be invalidated by incrementing the version
  number in a code update.

- **Fallback salt warning:** If `O365_SALT` is not set, a hardcoded
  fallback is used and a warning is logged. This is acceptable for local
  development but **must not be used in production**.

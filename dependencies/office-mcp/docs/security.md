# Security Model

This document describes the security architecture of the **office-mcp** package,
covering authentication, token management, input validation, and deployment
hardening. It is intended for developers integrating or deploying office-mcp and
for security reviewers auditing the codebase.

## Threat Model

office-mcp operates as a bridge between end users and the Microsoft Graph API.
The primary assets it protects are:

1. **OAuth2 tokens** -- access tokens and refresh tokens that grant access to a
   user's Microsoft 365 data (mail, calendar, files, Teams, directory).
2. **User credentials** -- passwords and API secrets stored in MongoDB for
   DB-based authentication.
3. **Session integrity** -- ensuring that tokens and sessions cannot be forged,
   replayed, or hijacked.

The threat model assumes:

- **Untrusted network** -- All HTTP traffic between the browser and the
  server may be intercepted. HTTPS is mandatory in production.
- **Untrusted user input** -- Query parameters, search terms, URLs, and
  identifiers supplied by users or API callers may contain injection payloads.
- **Shared infrastructure** -- Redis and MongoDB instances may be shared
  across services. Tokens stored in Redis must be encrypted so that a
  compromise of the cache layer does not directly yield usable tokens.
- **Compromised key files** -- Standalone MCP server key files contain
  sensitive credentials and must have restrictive filesystem permissions.

The following sections describe each security mechanism in detail.

## OAuth2 Token Management

office-mcp implements a multi-layer token management strategy built on
`WebUserInstance` (`web_user_instance.py`) and `MsGraphInstance`
(`msgraph/ms_graph_handler.py`).

### Token Encryption at Rest (Fernet / AES)

All OAuth tokens stored in Redis are encrypted using [Fernet symmetric
encryption](https://cryptography.io/en/latest/fernet/), which provides
AES-128-CBC encryption with HMAC-SHA256 authentication.

**Key derivation:**

The encryption key is derived using PBKDF2-HMAC-SHA256 with the following
parameters:

- **Input keying material**: the user's `session_id`
- **Salt**: the value of the `O365_SALT` environment variable
- **Iterations**: 100,000
- **Output length**: 32 bytes (256 bits), base64url-encoded for Fernet

```python
# From web_user_instance.py
kdf = PBKDF2HMAC(
    algorithm=hashes.SHA256(),
    length=32,
    salt=self._salt.encode(),
    iterations=100000,
)
key = base64.urlsafe_b64encode(kdf.derive(client_id.encode()))
```

This design ensures that:

- Each user session has a **unique encryption key** derived from its
  `session_id`. Compromising one session's key does not expose other
  sessions' tokens.
- The `O365_SALT` acts as a **server-side secret** that must be known to
  derive any key. Changing the salt invalidates all previously stored tokens.
- The high iteration count (100,000) makes brute-force attacks against the
  key derivation computationally expensive.

**Encryption and decryption:**

Tokens are encrypted before writing to Redis and decrypted after reading:

```python
# Encrypt before storing
encrypted_token = self._encrypt(access_token)
await redis_client.set(key, encrypted_token, ex=self.redis_lifetime)

# Decrypt after reading
decrypted_token = self._decrypt(encrypted_token)
```

The `_encrypt` and `_decrypt` methods use the Fernet instance, which
provides authenticated encryption -- any tampering with the ciphertext is
detected and causes decryption to fail gracefully (returning an empty string
and logging a warning).

**Fallback behavior:**

If `O365_SALT` is not set, a warning is logged and a hardcoded fallback
salt (`"nice_office_salt"`) is used. This fallback is **not secure for
production** and exists only to allow development without environment
configuration. Production deployments **must** set `O365_SALT`.

### Redis Token Storage Security

Tokens are stored in Redis under keys that incorporate a **SHA-256 hash** of the
user's identifier rather than the raw user ID or email:

```python
# From web_user_instance.py
def _hash_for_redis_key(self, user_id: str) -> str:
    return hashlib.sha256(user_id.encode()).hexdigest()

# Redis key structure: "{app}:token_cache_v2:{sha256_hash}:{token_type}"
```

This prevents user identifiers (which may be email addresses) from appearing
in Redis key listings.

Additional Redis security measures:

- **TTL enforcement**: All token entries have a 7-day TTL
  (`redis_lifetime = 7 * 24 * 60 * 60`), ensuring automatic cleanup of
  stale tokens.
- **Lazy connection**: The Redis client is lazily initialized and tested with
  a `ping()` before use. Connection failures are logged and handled
  gracefully.
- **Encrypted user data**: The `set_user_str_async` and
  `get_user_str_async` methods also encrypt arbitrary user data stored in
  Redis, not just tokens.

### Token Refresh Flow

Token refresh follows the standard OAuth2 refresh token grant:

1. `MsGraphInstance.refresh_async()` checks whether the current access token
   is within the `min_expiry` window (15 minutes before expiration).
2. If refresh is needed, `refresh_token_async()` retrieves the encrypted
   refresh token from Redis, decrypts it, and sends a `POST` to the Azure
   AD token endpoint (`https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`).
3. On success, both the new access token and the rotated refresh token are
   encrypted and stored back in Redis.
4. On failure, the access token is cleared from the cache (set to `None`)
   to force re-authentication.

For the standalone MCP server (`mcp_server.py`), refreshed tokens are
persisted back to the keyfile on disk using `_write_secure_json` with
restrictive file permissions (see [Keyfile Security](#keyfile-security)).

**Token expiration checking:**

Access token expiration is determined by decoding the JWT `exp` claim
without signature verification (since the token will be validated by
Microsoft Graph on use):

```python
decoded_token = jwt.decode(token, options={"verify_signature": False})
expiration_time = decoded_token.get("exp", 0)
```

## HMAC Token Signing (DBUserToken)

The `DBUserInstance` (`db_user_instance.py`) implements a self-contained
token system for database-backed authentication, independent of Azure AD.

### Token Structure

Tokens are Pydantic models (`DBUserToken`) serialized to JSON, then signed
with HMAC-SHA256:

```python
class DBUserToken(BaseModel):
    version: int = DB_USER_TOKEN_VERSION  # currently 3
    user_id: str
    given_name: str
    surname: str
    exp: float          # expiration timestamp
    issued_at: float    # issuance timestamp
    secret_hash: str    # SHA-256 hash of the user's password_secret
    max_duration: float | None = None
```

### Signing with O365_TOKEN_SECRET

Tokens are signed using HMAC-SHA256 with a server-side secret:

```python
@staticmethod
def _token_hmac_key() -> bytes:
    key = os.environ.get("O365_TOKEN_SECRET", os.environ.get("O365_SALT", ""))
    if not key:
        raise ValueError("O365_TOKEN_SECRET or O365_SALT must be set for token signing")
    return key.encode("utf-8")
```

The signing key is resolved in order of preference:

1. `O365_TOKEN_SECRET` (dedicated token signing key)
2. `O365_SALT` (fallback to the encryption salt)
3. If neither is set, a `ValueError` is raised -- the system **refuses to
   operate** without a signing key.

**Encoding (signing):**

```python
json_bytes = token_model.model_dump_json().encode("utf-8")
sig = hmac.new(key, json_bytes, hashlib.sha256).digest()
payload = sig + json_bytes  # 32-byte HMAC prefix + JSON
token_str = base64.urlsafe_b64encode(payload).decode("utf-8")
```

**Decoding (verification):**

```python
payload = base64.urlsafe_b64decode(token_str)
sig, json_bytes = payload[:32], payload[32:]
expected = hmac.new(key, json_bytes, hashlib.sha256).digest()
if not hmac.compare_digest(sig, expected):
    return None  # signature mismatch
return DBUserToken.model_validate_json(json_bytes)
```

The use of `hmac.compare_digest` prevents timing side-channel attacks during
signature comparison.

### Version Field for Token Invalidation

The `version` field (currently `DB_USER_TOKEN_VERSION = 3`) provides a
mechanism for **mass token invalidation**. When the version constant is
incremented, all previously issued tokens fail validation:

```python
def is_token_still_valid(self, token: str | None) -> bool:
    token_model = self.decode_token(token)
    if not token_model or token_model.user_id != self._db_identifier:
        return False
    return time.time() < token_model.exp and token_model.version == DB_USER_TOKEN_VERSION
```

This allows operators to invalidate all outstanding tokens by deploying a
version bump, without needing to rotate the signing secret or clear any
external state.

### Secret Hashing in Tokens

When a DB user logs in, the token embeds a **SHA-256 hash** of the user's
`password_secret` rather than the cleartext secret:

```python
def _hash_secret(secret: str) -> str:
    return hashlib.sha256(secret.encode("utf-8")).hexdigest()

# In login_async:
token_model = DBUserToken(
    ...
    secret_hash=_hash_secret(password_secret),
)
```

This ensures that even if a token is intercepted and decoded (the JSON
payload is not encrypted, only signed), the raw `password_secret` is not
exposed. The hash can be used for validation purposes (e.g., detecting
whether the secret has been rotated since the token was issued) without
revealing the secret itself.

## Authentication Flows

office-mcp supports three authentication flows, each suited to different
deployment scenarios.

### Web OAuth Flow (MsGraphInstance)

The web OAuth flow is used for browser-based applications where users log in
with their Microsoft 365 account.

**Flow:**

1. The application calls `build_auth_url()` on the `MsGraphInstance`, which
   uses MSAL (`ConfidentialClientApplication`) to generate an authorization
   URL pointing to `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize`.

2. The user is redirected to Microsoft's login page. On successful
   authentication, Microsoft redirects back to the application's registered
   redirect URI with an authorization code.

3. The application calls `acquire_token_async(code, redirect_url)` which
   exchanges the authorization code for access and refresh tokens via a
   `POST` to the Azure AD token endpoint.

4. Tokens are encrypted and stored in Redis. The user's profile is fetched
   from Microsoft Graph and cached.

**Redirect URL security:**

The redirect URL construction (`azure_auth_utils.py`) implements multiple
safeguards against open redirect attacks:

```python
# 1. Explicit URL takes highest priority (most secure)
explicit_url = os.environ.get("WEBSITE_REDIRECT_URL")
if explicit_url:
    return (explicit_url.rstrip("/") + "/" + base_path.lstrip("/")).rstrip("/")

# 2. Host allowlist validation
_allowed_hosts_str = os.environ.get(
    "ALLOWED_REDIRECT_HOSTS",
    os.environ.get("WEBSITE_HOSTNAME", ""),
)
_allowed_hosts = {h.strip().lower() for h in _allowed_hosts_str.split(",") if h.strip()}

# 3. Validate forwarded host against allowlist
host_name = host.split(":")[0].lower()
if _allowed_hosts and host_name not in _allowed_hosts:
    logger.warning("[AUTH] Redirect host '%s' not in allowed hosts", host_name)
    return base_url  # fall back to request's base URL
```

The security layers, in order of priority:

1. **WEBSITE_REDIRECT_URL** -- If set, the redirect URL is entirely
   determined by this environment variable. No request headers are consulted.
   This is the most secure option for production.
2. **ALLOWED_REDIRECT_HOSTS** -- A comma-separated allowlist of hostnames.
   If a forwarded host header (`x-forwarded-host`, `disguised-host`) is
   present, it is validated against this list. Unrecognized hosts are rejected
   and the system falls back to the request's base URL.
3. **WEBSITE_HOSTNAME** -- Falls back to this Azure App Service variable if
   `ALLOWED_REDIRECT_HOSTS` is not set.

**Cache prevention:**

The `NoCacheMiddleware` ensures that authentication-related responses are
never cached by browsers or proxies:

```python
response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
response.headers["Pragma"] = "no-cache"
response.headers["Expires"] = "0"
```

### DB-Based Authentication (DBUserInstance)

DB-based authentication is used for headless or API-driven scenarios where
users authenticate with a password or API secret rather than through an OAuth
browser flow.

**Password authentication:**

Passwords are hashed using **bcrypt** with a random salt:

```python
salt = bcrypt.gensalt()
hashed_password = bcrypt.hashpw(password.encode('utf-8'), salt)
```

Password verification uses `bcrypt.checkpw`, which is run in a separate
thread (`asyncio.to_thread`) to avoid blocking the event loop:

```python
if await asyncio.to_thread(bcrypt.checkpw, password.encode('utf-8'), password_hash):
    authenticated = True
```

**Secret-based authentication:**

API secrets are compared using `hmac.compare_digest` to prevent timing
attacks:

```python
import hmac as _hmac
if _hmac.compare_digest(secret, password_secret):
    authenticated = True
```

**Password rotation:**

When a password is set via `set_password_async`, a new random
`password_secret` (UUID hex) is also generated. This ensures that any
tokens containing the old `secret_hash` will no longer match the user's
current secret, providing an implicit mechanism for token invalidation on
password change.

**Token duration limits:**

Login accepts an optional `max_duration` parameter (default: 60 minutes)
that caps the token's `exp` claim. This prevents indefinitely-valid tokens.

### Session Management

Sessions are managed through the `WebUserInstance` base class with the
following properties:

- **Session identity**: Each session is identified by a `session_id`
  (typically a UUID). If no session ID is provided, one is generated.
- **Redis-backed persistence**: Tokens and profile data are persisted in
  Redis with a 7-day TTL, allowing sessions to survive server restarts.
- **Logout**: The `logout_async` method clears access tokens, refresh
  tokens, and cached profile data from both memory and Redis.
- **Thread safety**: In-memory token caches are protected by a
  `threading.RLock` to prevent race conditions in concurrent access.

## The `auth/` Subpackage

The `office_mcp/auth/` subpackage provides authentication utilities,
lifecycle callbacks, and a unified user abstraction.

### `azure_auth_utils.py` -- Redirect URL Builder and Middleware

This module provides two components:

**NoCacheMiddleware:**

A FastAPI/Starlette middleware that sets `Cache-Control`, `Pragma`, and
`Expires` headers on every response to prevent caching of sensitive
authentication data. This is applied as `BaseHTTPMiddleware`.

**get_redirect_url():**

Constructs the OAuth redirect URL with security hardening against open
redirect attacks. The function handles Azure App Service proxy headers
(`x-forwarded-proto`, `disguised-host`, `x-forwarded-host`) and
validates them against an allowlist. See the [Web OAuth Flow](#web-oauth-flow-msgraphinstance) section above
for a detailed description of the security layers.

### `background_service_registry.py` -- Login/Logout Lifecycle

The `BackgroundServiceRegistry` is a **singleton** that manages lifecycle
callbacks for user authentication events:

- **on_login(callback)** -- Register a callback invoked when a user logs in.
  The callback receives the `WebUserInstance`.
- **on_logout(callback)** -- Register a callback invoked when a user logs
  out. The callback receives the `user_id`.
- **on_loop(callback)** -- Register a periodic callback invoked every 5
  seconds for each active user. Used for background tasks such as token
  refresh monitoring.

**Security relevance:**

- The registry tracks `active_users` by user ID. The `notify_logout`
  method removes users from this map, ensuring that background services stop
  operating on behalf of logged-out users.
- Callbacks are wrapped in try/except blocks to prevent one faulty callback
  from disrupting the authentication lifecycle.
- The registry supports both synchronous and asynchronous callbacks,
  detecting coroutines at runtime and awaiting them as needed.

### `office_user_instance.py` -- Unified User Wrapper

The `OfficeUserInstance` class provides a unified wrapper around the
different authentication backends:

```python
class OfficeUserInstance:
    def __init__(self, config=None,
                 user_instance: MsGraphInstance | DBUserInstance | WebUserInstance | None = None):
        ...
```

It encapsulates:

- A `user_instance` that may be any of the three authentication types.
- An `OfficeUserConfig` (extensible Pydantic model) for implementation-specific
  configuration.
- Microsoft Graph **scope constants** defining the permissions required for
  each feature area (mail, calendar, profile, directory, chat, OneDrive).

**Scope constants** define the principle of least privilege for each feature:

| Scope Set | Permissions |
| --- | --- |
| `PROFILE_SCOPE` | `User.Read` |
| `DIRECTORY_SCOPE` | `Directory.Read.All`, `ProfilePhoto.Read.All` |
| `MAIL_SCOPE` | `Mail.Read`, `Mail.ReadWrite`, `Mail.Send`, `Mail.Read.Shared`, `Mail.ReadWrite.Shared`, `Mail.Send.Shared`, `User.ReadBasic.All` |
| `CALENDAR_SCOPE` | `Calendars.ReadWrite` |
| `CHAT_SCOPE` | `Chat.Read`, `ChannelMessage.Read.All` |
| `ONE_DRIVE_SCOPE` | `Files.Read.All`, `Files.ReadWrite.All` |

All permissions are **delegated** (not application-level), meaning the
application can only access data the authenticated user has access to.

## Input Validation and Injection Prevention

### OData Injection Prevention

The `FilesHandler` (`msgraph/files_handler.py`) constructs OData queries
for file search operations. User-supplied search queries are sanitized by
escaping single quotes to prevent OData injection:

```python
# From files_handler.py, search_items_async()
safe_query = query.replace("'", "''")
url = f".../root/search(q='{safe_query}')?$top={limit}"
```

Without this escaping, an attacker could inject arbitrary OData filter
expressions by including a single quote in the search query (e.g.,
`') or 1 eq 1 or ('`).

The `search_sites_async` method uses `urllib.parse.quote` for URL-encoding
the query parameter, which provides equivalent protection for the sites
search endpoint.

### URL Validation for Mail Handler

The `OfficeMailHandler` (`msgraph/mail_handler.py`) validates URLs passed
to `_build_mail_url` to prevent **Server-Side Request Forgery (SSRF)**:

```python
if url is not None:
    if not url.startswith(("http://", "https://")):
        # Relative path -- prepend the Graph endpoint
        url = f"{self.msg.msg_endpoint}{url}"
    elif not url.startswith("https://graph.microsoft.com/") and (
        not self.msg.msg_endpoint or not url.startswith(self.msg.msg_endpoint)
    ):
        raise ValueError("URL must point to the MS Graph API endpoint")
```

This validation ensures that:

1. **Relative URLs** are resolved against the configured MS Graph endpoint.
2. **Absolute URLs** must point to `https://graph.microsoft.com/` or the
   configured endpoint. Any other absolute URL is rejected with a
   `ValueError`, preventing an attacker from directing the server to make
   requests to arbitrary hosts while carrying the user's bearer token.

### Path Traversal Prevention

Redis keys are constructed using SHA-256 hashes of user identifiers rather
than raw identifiers. This eliminates the possibility of path traversal
attacks through crafted user IDs that might contain directory separators or
Redis key delimiters (`:`, `/`).

Additionally, the standalone MCP server (`mcp_server.py`) reads key files
from a path specified on the command line, but the tool definitions use only
opaque IDs (email IDs, drive item IDs, team IDs) that are passed to the
Graph API as URL segments. User-supplied IDs are never used to construct
local filesystem paths.

## Keyfile Security

The standalone MCP server (`mcp_server.py`) stores OAuth credentials in a
JSON key file on disk. This file contains highly sensitive material including
the `client_secret`, `access_token`, and `refresh_token`.

**Restrictive file permissions:**

Key files are written using `os.open` with explicit mode `0o600`
(owner read/write only), preventing other users on the system from reading
the credentials:

```python
def _write_secure_json(path: str, data: dict) -> None:
    fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w") as f:
        json.dump(data, f, indent=2)
```

This function is used both when exporting a new key file
(`export_keyfile`) and when persisting refreshed tokens back to the key
file after a successful token refresh.

**Key file contents:**

```json
{
  "app": "office-mcp",
  "session_id": "<SESSION_ID>",
  "email": "user@example.com",
  "access_token": "<ACCESS_TOKEN>",
  "refresh_token": "<REFRESH_TOKEN>",
  "client_id": "<CLIENT_ID>",
  "client_secret": "<CLIENT_SECRET>",
  "tenant_id": "<TENANT_ID>"
}
```

**Recommendations:**

- Store key files in a directory with restrictive permissions
  (e.g., `chmod 700 ~/.office-mcp/`).
- Do not commit key files to version control. Add `*.json` or the
  specific key file path to `.gitignore`.
- Rotate `client_secret` periodically in the Azure AD app registration.
- On systems with full-disk encryption, the key file is protected at rest
  by the OS. On unencrypted filesystems, consider additional encryption.

## Required Environment Variables for Security

The following environment variables are critical for the security of an
office-mcp deployment:

| Variable | Purpose | Required |
| --- | --- | --- |
| `O365_SALT` | Salt for PBKDF2 key derivation used to encrypt tokens in Redis. Changing this value invalidates all stored tokens. **Must be set in production** -- a fallback value is used in development but is not secure. | Yes |
| `O365_TOKEN_SECRET` | HMAC signing key for `DBUserToken` tokens. Falls back to `O365_SALT` if not set. Should be a high-entropy random string. If neither this nor `O365_SALT` is set, token signing raises a `ValueError` at runtime. | Yes (or `O365_SALT`) |
| `O365_CLIENT_ID` | Azure AD application (client) ID for OAuth2. | Yes |
| `O365_CLIENT_SECRET` | Azure AD client secret for the confidential client. | Yes |
| `WEBSITE_REDIRECT_URL` | Explicit OAuth redirect URL. When set, no request headers are consulted for redirect URL construction, eliminating open redirect risk entirely. | Recommended |
| `ALLOWED_REDIRECT_HOSTS` | Comma-separated allowlist of hostnames for redirect URL validation. Used when `WEBSITE_REDIRECT_URL` is not set. | Recommended |
| `O365_TENANT_ID` | Azure AD tenant ID. Defaults to `common` if not set, which allows any Azure AD tenant to authenticate. Set this to your organization's tenant ID to restrict authentication to your tenant only. | Recommended |

**Upstream application variables:**

The consuming application may use additional
security-related variables (e.g. a NiceGUI storage secret). These are
not part of the office-mcp package itself but may interact with it at the
integration layer.

## Security Best Practices for Deployment

### Environment and Secrets

- **Set O365_SALT to a cryptographically random value** (at least 32
  characters). Generate with: `python -c "import secrets; print(secrets.token_urlsafe(32))"`
- **Set O365_TOKEN_SECRET separately from O365_SALT** to provide independent
  key material for token encryption and token signing.
- **Never commit secrets to version control.** Use environment variables,
  Azure Key Vault, or a secrets manager.
- **Rotate O365_CLIENT_SECRET** periodically in the Azure AD app registration.
  After rotation, update the environment variable and restart the service.

### Network Security

- **Enforce HTTPS** in production. The OAuth redirect URL validation assumes
  HTTPS when `x-forwarded-proto` is `https`.
- **Set WEBSITE_REDIRECT_URL** to an explicit HTTPS URL to eliminate any
  possibility of open redirect attacks via header manipulation.
- **Restrict O365_TENANT_ID** to your organization's tenant to prevent
  authentication from external Azure AD tenants.

### Redis Security

- **Use TLS for Redis connections** (`rediss://` scheme) in production.
- **Require authentication** on the Redis instance (`requirepass` or ACLs).
- **Network-isolate Redis** -- it should not be accessible from the public
  internet.
- Even with these measures, tokens in Redis are encrypted, providing
  defense-in-depth against cache-layer compromise.

### MongoDB Security

- **Use TLS for MongoDB connections.**
- **Enable authentication** and use a dedicated database user with minimal
  permissions.
- Password hashes (bcrypt) and `password_secret` values are stored in
  MongoDB. Ensure the database is not publicly accessible.

### Keyfile Handling

- Key files are written with `0600` permissions. Do not relax these
  permissions.
- Do not store key files in shared directories or world-readable locations.
- If the MCP server is run as a service, ensure the service user has
  exclusive access to the key file directory.

### Token Lifecycle

- Access tokens have a short lifetime (typically 60 minutes, set by Azure AD
  or the `max_duration` parameter for DB tokens).
- The `DB_USER_TOKEN_VERSION` constant can be incremented to perform an
  emergency mass invalidation of all DB tokens.
- Changing `O365_SALT` invalidates all Redis-stored OAuth tokens,
  forcing all users to re-authenticate.
- Changing a user's password automatically generates a new
  `password_secret`, implicitly invalidating any tokens containing the old
  `secret_hash`.

### Logging and Monitoring

- Authentication events are logged with the `[AUTH]` prefix. Monitor these
  logs for:

  - Decryption failures (`_decrypt failed`) -- may indicate salt mismatch
    or token tampering.
  - Redirect host rejections (`Redirect host ... not in allowed hosts`) --
    may indicate open redirect attempts.
  - Token refresh failures (`token_refresh -- failed`) -- may indicate
    expired refresh tokens or revoked access.

- Token operations are logged with the `[OP]` prefix, including timing
  information for Azure AD token requests.

## Summary of Cryptographic Primitives

| Purpose | Algorithm | Details |
| --- | --- | --- |
| Token encryption at rest | Fernet (AES-128-CBC + HMAC-SHA256) | Key derived via PBKDF2-HMAC-SHA256 (100k iterations) |
| Redis key obfuscation | SHA-256 | One-way hash of user identifiers for key names |
| DB token signing | HMAC-SHA256 | 32-byte signature prepended to JSON payload |
| Password hashing | bcrypt | Random salt via `bcrypt.gensalt()` |
| Secret hashing in tokens | SHA-256 | One-way hash of `password_secret` embedded in token |
| Secret comparison | `hmac.compare_digest` | Constant-time comparison to prevent timing attacks |

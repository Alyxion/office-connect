import logging
from fastapi import Request
from starlette.middleware.base import BaseHTTPMiddleware

logger = logging.getLogger(__name__)

class NoCacheMiddleware(BaseHTTPMiddleware):
    """Middleware to prevent caching of responses"""
    async def dispatch(self, request: Request, call_next: object) -> object:
        response = await call_next(request)
        response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

def get_redirect_url(request: Request, base_path: str = "") -> str:
    """Get the authentication URL with proper protocol and host handling.

    Args:
        request: The FastAPI request object
        base_path: Optional base path to append to the URL

    Returns:
        str: The properly formatted authentication URL
    """
    import os

    # Prefer explicit redirect URL from environment (most secure)
    explicit_url = os.environ.get("WEBSITE_REDIRECT_URL")
    if explicit_url:
        return (explicit_url.rstrip("/") + "/" + base_path.lstrip("/")).rstrip("/")

    base_url = (str(request.base_url) + base_path).rstrip('/')

    # Allowed hosts for redirect URL construction (from env or derived from WEBSITE_HOSTNAME)
    _allowed_hosts_str = os.environ.get("ALLOWED_REDIRECT_HOSTS", os.environ.get("WEBSITE_HOSTNAME", ""))
    _allowed_hosts = {h.strip().lower() for h in _allowed_hosts_str.split(",") if h.strip()} if _allowed_hosts_str else set()

    # Try Azure App Service headers first, then fall back to standard forwarded headers
    if "x-forwarded-proto" in request.headers:
        proto = request.headers["x-forwarded-proto"]
        # Try Azure's disguised-host first, then fall back to x-forwarded-host
        if "disguised-host" in request.headers:
            host = request.headers["disguised-host"]
        elif "x-forwarded-host" in request.headers:
            host = request.headers["x-forwarded-host"].split(":")[0]
            port = request.headers.get("x-forwarded-port", "443" if proto == "https" else "80")
            host = f"{host}:{port}"
        else:
            host = request.headers.get("host", "localhost")

        # Validate host against allowlist to prevent open redirect
        host_name = host.split(":")[0].lower()
        if _allowed_hosts and host_name not in _allowed_hosts:
            logger.warning("[AUTH] Redirect host '%s' not in allowed hosts, using request base_url", host_name)
            return base_url

        base_url = str(base_url)
        # Replace the protocol and host parts of the URL
        elements = base_url.split('/', 3)
        path = elements[-1] if len(elements) > 3 else ""
        if not path.startswith('/'):
            path = '/' + path
        base_url = f"{proto}://{host}{path}"
    return base_url


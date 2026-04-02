"""Minimal FastAPI sample for office-connect OAuth + Mail + Calendar.

    cd /path/to/office-connect
    cp samples/.env.template samples/.env   # fill in credentials
    python samples/web_app.py

Accessible via https://localhost:8443/ (behind the HTTPS proxy on port 8080).
"""

from __future__ import annotations

import logging
import os
import secrets
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

from dotenv import load_dotenv

import uvicorn
from fastapi import FastAPI, Request, Response, HTTPException, Query
from fastapi.middleware.trustedhost import TrustedHostMiddleware
from fastapi.responses import HTMLResponse, RedirectResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

SAMPLES_DIR = Path(__file__).resolve().parent
PROJECT_DIR = SAMPLES_DIR.parent
TOKEN_FILE = PROJECT_DIR / "tests" / "msgraph_test_token.json"
STATIC_DIR = SAMPLES_DIR / "static"

# Load .env from samples directory
load_dotenv(SAMPLES_DIR / ".env")

# Add project root so office_con is importable without install
sys.path.insert(0, str(PROJECT_DIR))

from office_con.msgraph.ms_graph_handler import MsGraphInstance       # noqa: E402
from office_con.mcp_server import export_keyfile                     # noqa: E402
from office_con.auth.office_user_instance import OfficeUserInstance   # noqa: E402

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

PORT = int(os.environ.get("SAMPLE_PORT", "8080"))
# Base URL as seen by the browser (behind the HTTPS proxy)
BASE_URL = os.environ.get("SAMPLE_BASE_URL", "https://localhost:8443")
REDIRECT_URI = f"{BASE_URL}/auth"

SCOPES = list(set(
    OfficeUserInstance.PROFILE_SCOPE
    + OfficeUserInstance.MAIL_SCOPE
    + OfficeUserInstance.CALENDAR_SCOPE
))

logging.basicConfig(level=logging.INFO)
log = logging.getLogger("sample")

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

app = FastAPI(docs_url=None, redoc_url=None)

# CSRF token per process (single-user dev tool)
CSRF_TOKEN = secrets.token_hex(16)

# In-memory session — single-user sample app
_graph: Optional[MsGraphInstance] = None


def _get_graph() -> MsGraphInstance:
    if _graph is None:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return _graph


# ---------------------------------------------------------------------------
# Security middleware
# ---------------------------------------------------------------------------

@app.middleware("http")
async def security_headers(request: Request, call_next):
    response: Response = await call_next(request)
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "strict-origin-when-cross-origin"
    # CSP: allow Quasar inline styles, font loading, sandboxed mail body iframe
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; "
        "script-src 'self' 'unsafe-eval'; "
        "style-src 'self' 'unsafe-inline'; "
        "font-src 'self'; "
        "img-src 'self' data:; "
        "connect-src 'self'; "
        "frame-src 'self' blob:; "
        "frame-ancestors 'none'"
    )
    return response


@app.middleware("http")
async def csrf_check(request: Request, call_next):
    """Require X-CSRF-Token header on state-changing requests."""
    if request.method in ("POST", "PUT", "PATCH", "DELETE"):
        token = request.headers.get("X-CSRF-Token", "")
        if token != CSRF_TOKEN:
            return JSONResponse({"error": "CSRF token mismatch"}, status_code=403)
    return await call_next(request)


# ---------------------------------------------------------------------------
# Auth endpoints
# ---------------------------------------------------------------------------

@app.get("/")
async def index():
    return HTMLResponse((STATIC_DIR / "index.html").read_text())


@app.get("/auth")
async def oauth_callback(code: str = Query(...)):
    """OAuth callback — Microsoft redirects here with ?code=."""
    global _graph
    graph = MsGraphInstance(
        scopes=SCOPES,
        client_id=os.environ.get("O365_CLIENT_ID"),
        client_secret=os.environ.get("O365_CLIENT_SECRET"),
        tenant_id=os.environ.get("O365_TENANT_ID", "common"),
        endpoint=os.environ.get("O365_ENDPOINT", "https://graph.microsoft.com/v1.0/"),
    )
    result = await graph.acquire_token_async(code, REDIRECT_URI)
    if isinstance(result, HTMLResponse) and result.status_code >= 400:
        return result
    _graph = graph
    # Persist token for automated tests
    TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
    export_keyfile(
        str(TOKEN_FILE),
        access_token=graph.cache_dict.get("access_token", ""),
        refresh_token=graph.cache_dict.get("refresh_token", ""),
        client_id=graph.client_id or "",
        client_secret=graph.client_secret or "",
        tenant_id=graph.tenant_id or "common",
        app="office-connect-sample",
        email=graph.email,
    )
    log.info("Token saved to %s", TOKEN_FILE)
    return RedirectResponse(f"{BASE_URL}/")


# Serve static assets (vendor/, app.js, app.css)
app.mount("/vendor", StaticFiles(directory=str(STATIC_DIR / "vendor")), name="vendor")


@app.get("/app.js")
async def serve_js():
    return Response(
        (STATIC_DIR / "app.js").read_text(),
        media_type="application/javascript",
    )


@app.get("/app.css")
async def serve_css():
    return Response(
        (STATIC_DIR / "app.css").read_text(),
        media_type="text/css",
    )


@app.get("/csrf-token")
async def get_csrf():
    return {"token": CSRF_TOKEN}


@app.get("/login")
async def login():
    """Start the OAuth flow — redirect to Microsoft."""
    graph = MsGraphInstance(
        scopes=SCOPES,
        client_id=os.environ.get("O365_CLIENT_ID"),
        client_secret=os.environ.get("O365_CLIENT_SECRET"),
        tenant_id=os.environ.get("O365_TENANT_ID", "common"),
        endpoint=os.environ.get("O365_ENDPOINT", "https://graph.microsoft.com/v1.0/"),
        select_account=True,
    )
    auth_url = graph.build_auth_url(REDIRECT_URI)
    return RedirectResponse(auth_url)


@app.get("/auth-status")
async def auth_status():
    if _graph is None:
        return {"authenticated": False}
    return {"authenticated": True, "email": _graph.email}


# ---------------------------------------------------------------------------
# Mail API
# ---------------------------------------------------------------------------

@app.get("/api/mail/folders")
async def list_mail_folders():
    """List mail folders via Graph API."""
    graph = _get_graph()
    token = await graph.get_access_token_async()
    if not token:
        raise HTTPException(401, "Token expired")
    resp = await graph.run_async(
        url=f"{graph.msg_endpoint}me/mailFolders?$top=50",
        token=token,
    )
    if resp is None or resp.status_code != 200:
        raise HTTPException(502, "Failed to fetch folders")
    folders = resp.json().get("value", [])
    return [
        {
            "id": f["id"],
            "name": f.get("displayName", ""),
            "unread": f.get("unreadItemCount", 0),
            "total": f.get("totalItemCount", 0),
        }
        for f in folders
    ]


@app.get("/api/mail/messages")
async def list_messages(
    folder_id: str = Query("inbox"),
    limit: int = Query(20, ge=1, le=100),
    skip: int = Query(0, ge=0),
):
    """List messages in a folder."""
    graph = _get_graph()
    token = await graph.get_access_token_async()
    if not token:
        raise HTTPException(401, "Token expired")
    fields = "id,from,subject,bodyPreview,receivedDateTime,isRead,hasAttachments,importance"
    url = (
        f"{graph.msg_endpoint}me/mailFolders/{folder_id}/messages"
        f"?$select={fields}&$top={limit}&$skip={skip}"
        f"&$orderby=receivedDateTime desc&$count=true"
    )
    resp = await graph.run_async(url=url, token=token)
    if resp is None or resp.status_code != 200:
        raise HTTPException(502, "Failed to fetch messages")
    data = resp.json()
    messages = []
    for m in data.get("value", []):
        from_addr = m.get("from", {}).get("emailAddress", {})
        messages.append({
            "id": m["id"],
            "from_name": from_addr.get("name", ""),
            "from_email": from_addr.get("address", ""),
            "subject": m.get("subject", "(no subject)"),
            "preview": m.get("bodyPreview", ""),
            "received": m.get("receivedDateTime", ""),
            "is_read": m.get("isRead", False),
            "has_attachments": m.get("hasAttachments", False),
            "importance": m.get("importance", "normal"),
        })
    return {"messages": messages, "total": data.get("@odata.count", len(messages))}


@app.get("/api/mail/messages/{message_id}")
async def get_message(message_id: str):
    """Get a single message with body."""
    graph = _get_graph()
    mail = graph.get_mail()
    result = await mail.get_mail_async(email_id=message_id)
    if result is None:
        raise HTTPException(404, "Message not found")
    return result.model_dump(exclude={"attachments", "zip_data"})


@app.post("/api/mail/draft")
async def create_draft(request: Request):
    """Create a draft message."""
    graph = _get_graph()
    body = await request.json()
    to = body.get("to", [])
    subject = body.get("subject", "")
    content = body.get("body", "")
    is_html = body.get("is_html", False)
    if not to or not subject:
        raise HTTPException(400, "to and subject are required")
    mail = graph.get_mail()
    result = await mail.create_draft_async(
        to_recipients=to, subject=subject, body=content, is_html=is_html,
    )
    if result is None:
        raise HTTPException(502, "Failed to create draft")
    return result


@app.post("/api/mail/send")
async def send_mail(request: Request):
    """Send a new message directly."""
    graph = _get_graph()
    body = await request.json()
    to = body.get("to", [])
    subject = body.get("subject", "")
    content = body.get("body", "")
    is_html = body.get("is_html", False)
    if not to or not subject:
        raise HTTPException(400, "to and subject are required")
    mail = graph.get_mail()
    ok = await mail.send_message_async(
        to_recipients=to, subject=subject, body=content, is_html=is_html,
    )
    if not ok:
        raise HTTPException(502, "Failed to send")
    return {"ok": True}


@app.post("/api/mail/draft/{message_id}/send")
async def send_draft(message_id: str):
    """Send an existing draft."""
    graph = _get_graph()
    mail = graph.get_mail()
    ok = await mail.send_draft_async(message_id)
    if not ok:
        raise HTTPException(502, "Failed to send draft")
    return {"ok": True}


# ---------------------------------------------------------------------------
# Mail reply
# ---------------------------------------------------------------------------

@app.post("/api/mail/reply")
async def reply_to_message(request: Request):
    """Reply (or reply-all) to a message."""
    graph = _get_graph()
    body = await request.json()
    message_id = body.get("message_id")
    comment = body.get("body", "")
    reply_all = body.get("reply_all", False)
    if not message_id:
        raise HTTPException(400, "message_id is required")
    token = await graph.get_access_token_async()
    if not token:
        raise HTTPException(401, "Token expired")
    action = "replyAll" if reply_all else "reply"
    url = f"{graph.msg_endpoint}me/messages/{message_id}/{action}"
    resp = await graph.run_async(
        url=url, method="POST", json={"comment": comment}, token=token,
    )
    if resp is None or resp.status_code >= 300:
        raise HTTPException(502, "Failed to send reply")
    return {"ok": True}


# ---------------------------------------------------------------------------
# People / Directory API
# ---------------------------------------------------------------------------

@app.get("/api/people/search")
async def search_people(q: str = Query("", min_length=0)):
    """Search the directory for people matching *q*."""
    graph = _get_graph()
    directory = graph.get_directory()
    result = await directory.get_users_async()
    query = q.strip().lower()
    people = []
    for u in result.users:
        if query and query not in (u.display_name or "").lower() \
                and query not in (u.email or "").lower() \
                and query not in (u.job_title or "").lower() \
                and query not in (u.department or "").lower():
            continue
        first = (u.given_name or u.display_name or "?")[0].upper()
        last = (u.surname or "")[0].upper() if u.surname else ""
        people.append({
            "id": u.id,
            "display_name": u.display_name,
            "email": u.email,
            "job_title": u.job_title,
            "department": u.department,
            "initials": first + last,
        })
    return people


@app.get("/api/people/{user_id}/photo")
async def get_person_photo(user_id: str):
    """Return a user's profile photo as image/jpeg, or 404."""
    graph = _get_graph()
    directory = graph.get_directory()
    photo_bytes = await directory.get_user_photo_async(user_id)
    if photo_bytes is None:
        raise HTTPException(404, "Photo not found")
    return Response(content=photo_bytes, media_type="image/jpeg")


# ---------------------------------------------------------------------------
# Calendar API
# ---------------------------------------------------------------------------

@app.get("/api/calendar/list")
async def list_calendars():
    graph = _get_graph()
    cal = graph.get_calendar()
    calendars = await cal.get_calendars_async()
    return [
        {
            "id": c.get("id"),
            "name": c.get("name", ""),
            "color": c.get("hexColor", ""),
            "is_default": c.get("isDefaultCalendar", False),
        }
        for c in calendars
    ]


@app.get("/api/calendar/events")
async def list_events(
    start: str = Query(None),
    end: str = Query(None),
    limit: int = Query(50, ge=1, le=200),
):
    graph = _get_graph()
    cal = graph.get_calendar()
    now = datetime.now()
    start_dt = datetime.fromisoformat(start) if start else now - timedelta(days=7)
    end_dt = datetime.fromisoformat(end) if end else now + timedelta(days=30)
    result = await cal.get_events_async(start_date=start_dt, end_date=end_dt, limit=limit)
    return result.model_dump()


# ---------------------------------------------------------------------------
# Profile
# ---------------------------------------------------------------------------

@app.get("/api/profile")
async def get_profile():
    graph = _get_graph()
    handler = await graph.get_profile_async()
    if handler.me is None:
        raise HTTPException(502, "Could not load profile")
    return handler.me.model_dump()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    uvicorn.run(
        "web_app:app",
        host="0.0.0.0",
        port=PORT,
        log_level="info",
        app_dir=str(SAMPLES_DIR),
    )

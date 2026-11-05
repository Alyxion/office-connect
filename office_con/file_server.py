"""A loopback-only HTTP server for handing bytes (e.g. a mail attachment) to a
*local* fetcher — Claude Code / cowork on the same host — via an **unguessable
URL**, instead of inlining base64 into the model context.

Singleton across processes
--------------------------
Several MCP clients (Claude Desktop, Claude Code, cowork) each spawn their *own*
``office_con.mcp_server`` process, and each may be launched with ``--file-server``.
We must not wind up one HTTP listener per process fighting over ports. Instead all
file-server-enabled processes share **one** listener and **one** on-disk spool:

* **The bound TCP port is the mutex.** On ``start()`` a process tries to bind the
  shared port. Exactly one can hold an active listener at a time (the OS rejects a
  second bind with ``EADDRINUSE``); that process is the *listener*. The others bind
  nothing and become *publishers only*.
* **The spool is shared on disk.** Payloads live as ``spool/<token>.bin`` +
  ``spool/<token>.json`` under a state dir (default
  ``~/.config/office-connect/file-server/``, dir ``0700``, files ``0600``). Any
  process — listener or publisher — publishes by writing those two files; the
  listener serves whatever is in the spool. So a publisher's URL works even though
  it never opened a socket.
* **It lives as long as any process needs it.** A heartbeat in every process keeps
  re-trying the bind; the instant the listener dies and frees the port, a surviving
  process takes it over on the next tick — same port, same spool, so already-issued
  URLs keep working. When the *last* file-server process exits, nothing re-binds and
  the listener is simply gone. The port is the singleton; no daemon, no lockfile.

Security model
--------------
* **Loopback only.** Binds ``127.0.0.1`` — never reachable from off-host. (A
  sandboxed/network-isolated client on the same machine also cannot reach it; such
  clients must use base64 or a shared filesystem path, not the URL.)
* **Unguessable URLs.** Each file is addressed by a ``secrets.token_urlsafe(32)``
  token (~256 bits). The token is the only credential and is URL-safe, so it doubles
  as the spool filename without path-traversal risk. An unknown/expired token gets a
  bare ``404``.
* **Expiry + one-shot.** Entries expire after a TTL (default 1 h); one-shot entries
  are unlinked on first successful fetch. The listener sweeps expired files.
* **Opt-in + never on a server.** Disabled by default; the caller gates on
  ``is_server_environment()`` so mailbox bytes are never spooled to disk on a
  hosted/multi-tenant box.
"""

from __future__ import annotations

import asyncio
import json
import logging
import os
import secrets
import time
from pathlib import Path
from urllib.parse import quote

from aiohttp import web

logger = logging.getLogger(__name__)

_DEFAULT_STATE_DIR = Path.home() / ".config" / "office-connect" / "file-server"
_HEARTBEAT_SECONDS = 3.0


# Strip control chars / quotes from a filename before putting it in a
# Content-Disposition header (defense against header injection).
def _safe_header_filename(name: str | None) -> str:
    base = (name or "file").replace("\r", "").replace("\n", "").replace('"', "")
    return base[:200] or "file"


def _looks_like_token(token: str) -> bool:
    """A spool token is exactly what ``secrets.token_urlsafe`` emits: URL-safe
    base64 (``A-Z a-z 0-9 - _``). Reject anything else so a crafted path component
    can never escape the spool dir."""
    return bool(token) and len(token) <= 128 and all(
        c.isalnum() or c in "-_" for c in token
    )


class LocalFileServer:
    """Serves published byte payloads over loopback HTTP at unguessable URLs,
    sharing one listener + one disk spool across every file-server process."""

    def __init__(
        self, *, host: str = "127.0.0.1", port: int = 0,
        ttl_seconds: int = 3600, max_entries: int = 256,
        state_dir: str | os.PathLike[str] | None = None,
    ):
        self._host = host
        self._configured_port = port  # 0 = let the first process pick one
        self._ttl = max(1, ttl_seconds)
        self._max_entries = max(1, max_entries)
        self._state_dir = Path(state_dir) if state_dir is not None else _DEFAULT_STATE_DIR
        self._spool = self._state_dir / "spool"
        self._port_file = self._state_dir / "port"
        self._runner: web.AppRunner | None = None
        self._is_listener = False
        self._port: int | None = None  # the shared port (listener or recorded)
        self._heartbeat: asyncio.Task | None = None
        self._started = False

    # -- lifecycle ---------------------------------------------------------

    @property
    def running(self) -> bool:
        return self._started and self._port is not None

    @property
    def base_url(self) -> str:
        if self._port is None:
            raise RuntimeError("file server not started")
        return f"http://{self._host}:{self._port}"

    def _ensure_dirs(self) -> None:
        self._state_dir.mkdir(parents=True, exist_ok=True)
        self._spool.mkdir(parents=True, exist_ok=True)
        try:
            os.chmod(self._state_dir, 0o700)
            os.chmod(self._spool, 0o700)
        except OSError:
            pass

    def _read_recorded_port(self) -> int | None:
        try:
            return int(self._port_file.read_text().strip())
        except (OSError, ValueError):
            return None

    def _write_recorded_port(self, port: int) -> None:
        tmp = self._port_file.with_suffix(".tmp")
        tmp.write_text(str(port))
        os.replace(tmp, self._port_file)

    async def _try_listen(self, port: int) -> bool:
        """Attempt to become THE listener on ``port``. Returns True on success,
        False if the port is already held by a sibling process (EADDRINUSE)."""
        app = web.Application()
        app.router.add_get("/f/{token}/{name}", self._handle)
        app.router.add_get("/f/{token}", self._handle)
        runner = web.AppRunner(app, access_log=None)
        await runner.setup()
        site = web.TCPSite(runner, self._host, port)
        try:
            await site.start()
        except OSError:
            await runner.cleanup()
            return False
        self._runner = runner
        bound = port
        server = getattr(site, "_server", None)
        socks = getattr(server, "sockets", None) if server else None
        if socks:
            bound = socks[0].getsockname()[1]
        self._port = bound
        self._is_listener = True
        self._write_recorded_port(bound)
        logger.warning(
            "[FILES] file server listener bound on %s (ttl=%ds, loopback only)",
            self.base_url, self._ttl,
        )
        return True

    async def start(self) -> "LocalFileServer":
        self._ensure_dirs()
        # Prefer an explicitly configured port; otherwise reuse the port a sibling
        # already recorded so issued URLs survive a listener handover.
        desired = self._configured_port or self._read_recorded_port() or 0
        if not await self._try_listen(desired):
            # A sibling owns the listener — run as publisher only. Its recorded
            # port is what our URLs must point at.
            for _ in range(20):  # brief race window: wait for the port file
                self._port = self._read_recorded_port()
                if self._port is not None:
                    break
                await asyncio.sleep(0.05)
            if self._port is None:
                # No sibling actually recorded a port (lost the race after all) —
                # try once more to bind ourselves.
                await self._try_listen(self._configured_port or 0)
            logger.warning(
                "[FILES] file server attached as publisher (listener on port %s)",
                self._port,
            )
        self._started = True
        self._heartbeat = asyncio.ensure_future(self._heartbeat_loop())
        return self

    async def stop(self) -> None:
        self._started = False
        if self._heartbeat is not None:
            self._heartbeat.cancel()
            try:
                await self._heartbeat
            except (asyncio.CancelledError, Exception):
                pass
            self._heartbeat = None
        if self._runner is not None:
            await self._runner.cleanup()
            self._runner = None
        self._is_listener = False

    async def _heartbeat_loop(self) -> None:
        while True:
            await asyncio.sleep(_HEARTBEAT_SECONDS)
            if not self._is_listener:
                # Re-election: the listener may have died and freed the port.
                target = self._port or self._configured_port or self._read_recorded_port() or 0
                if target and await self._try_listen(target):
                    logger.warning("[FILES] took over file server listener on port %s", self._port)
            if self._is_listener:
                self._sweep_expired()

    # -- publish / serve ---------------------------------------------------

    def publish(
        self, data: bytes, *, name: str,
        content_type: str = "application/octet-stream", one_shot: bool = False,
    ) -> dict:
        """Write ``data`` to the shared spool and return ``{url, expires_in, bytes}``.
        Works whether this process is the listener or a publisher — the bytes go to
        disk and whichever process holds the port serves them. The URL embeds a
        fresh high-entropy token unrelated to the content."""
        if not self.running:
            raise RuntimeError("file server not started")
        self._sweep_expired()
        self._enforce_max_entries()
        token = secrets.token_urlsafe(32)
        now = time.time()
        meta = {
            "name": name,
            "content_type": content_type or "application/octet-stream",
            "created": now,
            "expires": now + self._ttl,
            "one_shot": bool(one_shot),
        }
        bin_path = self._spool / f"{token}.bin"
        meta_path = self._spool / f"{token}.json"
        # Write payload first, then the sidecar (its presence marks the entry ready).
        self._atomic_write(bin_path, data)
        self._atomic_write(meta_path, json.dumps(meta).encode("utf-8"))
        url = f"{self.base_url}/f/{token}/{quote(name or 'file', safe='')}"
        return {"url": url, "expires_in": self._ttl, "bytes": len(data)}

    def _atomic_write(self, path: Path, data: bytes) -> None:
        tmp = path.with_suffix(path.suffix + ".tmp")
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        try:
            os.write(fd, data)
        finally:
            os.close(fd)
        os.replace(tmp, path)

    def _entry_paths(self):
        for meta_path in self._spool.glob("*.json"):
            yield meta_path

    def _sweep_expired(self) -> None:
        now = time.time()
        for meta_path in list(self._entry_paths()):
            try:
                meta = json.loads(meta_path.read_text())
            except (OSError, ValueError):
                continue
            if meta.get("expires", 0) < now:
                self._unlink_entry(meta_path.stem)

    def _enforce_max_entries(self) -> None:
        metas = list(self._entry_paths())
        if len(metas) < self._max_entries:
            return
        # Evict oldest by creation time until under the cap.
        scored = []
        for m in metas:
            try:
                scored.append((json.loads(m.read_text()).get("created", 0), m.stem))
            except (OSError, ValueError):
                scored.append((0, m.stem))
        scored.sort()
        for _, token in scored[: len(scored) - self._max_entries + 1]:
            self._unlink_entry(token)

    def _unlink_entry(self, token: str) -> None:
        for suffix in (".bin", ".json"):
            try:
                (self._spool / f"{token}{suffix}").unlink()
            except OSError:
                pass

    async def _handle(self, request: web.Request) -> web.Response:
        token = request.match_info.get("token", "")
        if not _looks_like_token(token):
            raise web.HTTPNotFound()
        meta_path = self._spool / f"{token}.json"
        bin_path = self._spool / f"{token}.bin"
        try:
            meta = json.loads(meta_path.read_text())
            data = bin_path.read_bytes()
        except (OSError, ValueError):
            raise web.HTTPNotFound()
        if meta.get("expires", 0) < time.time():
            self._unlink_entry(token)
            raise web.HTTPNotFound()
        if meta.get("one_shot"):
            self._unlink_entry(token)
        return web.Response(
            body=data,
            content_type=meta.get("content_type") or "application/octet-stream",
            headers={
                "Content-Disposition": f'attachment; filename="{_safe_header_filename(meta.get("name"))}"',
                "Cache-Control": "no-store",
                "X-Content-Type-Options": "nosniff",
            },
        )

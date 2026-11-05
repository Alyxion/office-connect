"""A minimal MCP **client** for exercising the office-connect stdio server.

This speaks the MCP stdio transport directly — newline-delimited JSON-RPC 2.0
over a subprocess's stdin/stdout — using only the Python standard library (no
MCP SDK). It is deliberately small so it *documents* the wire protocol and works
as a lightweight, scriptable test harness for the very tools the server exposes.

Two entry points:

* :class:`MCPClient` — spawn any MCP stdio server and call initialize /
  ``tools/list`` / ``tools/call``.
* :func:`mock_server_client` — spawn *this project's* server in offline MOCK
  mode (synthetic data, no Office 365 credentials, no network) and yield a
  connected client. This is what the e2e tests and the CLI use.

CLI::

    python -m office_con.testing.mcp_client --mock list
    python -m office_con.testing.mcp_client --mock call o365_search_mail '{"body":"hotel","since":"2025-01-01","until":"2025-12-31"}'
    python -m office_con.testing.mcp_client --keyfile ~/.config/office-connect/token.json list

See ``docs/mcp_testing.md`` for the full guide.
"""

from __future__ import annotations

import asyncio
import json
import os
import sys
import tempfile
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Any, AsyncIterator

# Negotiated with the server during initialize; the server echoes back the
# version it actually supports — we don't hard-fail on a mismatch.
_PROTOCOL_VERSION = "2025-06-18"
# Tool results (e.g. a batch of mails) can be large; raise the per-line cap well
# above asyncio's 64 KiB default so readline() never trips on a big JSON-RPC msg.
_STDOUT_LIMIT = 16 * 1024 * 1024


class MCPError(RuntimeError):
    """An MCP request failed at the protocol level (JSON-RPC error, timeout, or
    a dead server). Tool-level failures are NOT raised — the server returns them
    as normal result content, so inspect the returned text instead."""


class MCPClient:
    """JSON-RPC-over-stdio client for an MCP server subprocess."""

    def __init__(self, command: list[str], *, env: dict | None = None, cwd: str | None = None):
        self._command = command
        self._env = env
        self._cwd = cwd
        self._proc: asyncio.subprocess.Process | None = None
        self._next_id = 0
        self._pending: dict[int, asyncio.Future] = {}
        self._reader_task: asyncio.Task | None = None
        self._stderr_task: asyncio.Task | None = None
        self._stderr_buf: list[str] = []
        self.server_info: dict[str, Any] = {}

    async def __aenter__(self) -> "MCPClient":
        await self.start()
        return self

    async def __aexit__(self, *exc) -> None:
        await self.stop()

    # ── lifecycle ────────────────────────────────────────────
    async def start(self) -> None:
        self._proc = await asyncio.create_subprocess_exec(
            *self._command,
            stdin=asyncio.subprocess.PIPE,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
            env=self._env,
            cwd=self._cwd,
            limit=_STDOUT_LIMIT,
        )
        self._reader_task = asyncio.create_task(self._read_loop())
        self._stderr_task = asyncio.create_task(self._drain_stderr())
        await self.initialize()

    async def stop(self) -> None:
        proc = self._proc
        if proc is None:
            return
        try:
            if proc.returncode is None:
                proc.terminate()
                try:
                    await asyncio.wait_for(proc.wait(), timeout=5)
                except asyncio.TimeoutError:
                    proc.kill()
                    await proc.wait()
        finally:
            for task in (self._reader_task, self._stderr_task):
                if task is not None:
                    task.cancel()
            for fut in self._pending.values():
                if not fut.done():
                    fut.set_exception(MCPError("client stopped before a response arrived"))
            self._pending.clear()

    # ── wire protocol ────────────────────────────────────────
    async def _read_loop(self) -> None:
        assert self._proc is not None and self._proc.stdout is not None
        try:
            while True:
                raw = await self._proc.stdout.readline()
                if not raw:
                    break  # EOF — server exited
                line = raw.strip()
                if not line:
                    continue
                try:
                    msg = json.loads(line)
                except json.JSONDecodeError:
                    continue  # stray non-JSON line on stdout — ignore
                mid = msg.get("id")
                if mid is not None and mid in self._pending:
                    fut = self._pending.pop(mid)
                    if not fut.done():
                        fut.set_result(msg)
                # server-initiated requests / notifications: not needed here
        except asyncio.CancelledError:
            pass

    async def _drain_stderr(self) -> None:
        assert self._proc is not None and self._proc.stderr is not None
        try:
            while True:
                raw = await self._proc.stderr.readline()
                if not raw:
                    break
                self._stderr_buf.append(raw.decode(errors="replace").rstrip())
                if len(self._stderr_buf) > 2000:  # keep memory bounded
                    del self._stderr_buf[:1000]
        except asyncio.CancelledError:
            pass

    def stderr_tail(self, n: int = 20) -> str:
        return "\n".join(self._stderr_buf[-n:])

    def _write(self, obj: dict) -> None:
        if self._proc is None or self._proc.stdin is None:
            raise MCPError("server process is not running")
        self._proc.stdin.write((json.dumps(obj) + "\n").encode())

    async def _notify(self, method: str, params: dict | None = None) -> None:
        self._write({"jsonrpc": "2.0", "method": method, "params": params or {}})
        await self._proc.stdin.drain()  # type: ignore[union-attr]

    async def _request(self, method: str, params: dict | None = None, *, timeout: float = 60.0) -> dict:
        if self._proc is not None and self._proc.returncode is not None:
            raise MCPError(
                f"server already exited (code {self._proc.returncode}).\n"
                f"stderr tail:\n{self.stderr_tail()}"
            )
        self._next_id += 1
        mid = self._next_id
        fut: asyncio.Future = asyncio.get_running_loop().create_future()
        self._pending[mid] = fut
        self._write({"jsonrpc": "2.0", "id": mid, "method": method, "params": params or {}})
        await self._proc.stdin.drain()  # type: ignore[union-attr]
        try:
            msg = await asyncio.wait_for(fut, timeout=timeout)
        except asyncio.TimeoutError:
            self._pending.pop(mid, None)
            raise MCPError(
                f"timed out after {timeout}s waiting for '{method}'.\n"
                f"stderr tail:\n{self.stderr_tail()}"
            )
        if "error" in msg:
            raise MCPError(f"'{method}' returned a JSON-RPC error: {msg['error']}")
        return msg.get("result", {})

    # ── high-level API ───────────────────────────────────────
    async def initialize(self) -> dict:
        result = await self._request("initialize", {
            "protocolVersion": _PROTOCOL_VERSION,
            "capabilities": {},
            "clientInfo": {"name": "office-connect-test-client", "version": "1.0"},
        })
        self.server_info = result.get("serverInfo", {})
        await self._notify("notifications/initialized")
        return result

    async def list_tools(self) -> list[dict]:
        return (await self._request("tools/list")).get("tools", [])

    async def call_tool(self, name: str, arguments: dict | None = None, *, timeout: float = 60.0) -> dict:
        return await self._request(
            "tools/call", {"name": name, "arguments": arguments or {}}, timeout=timeout,
        )

    async def call_tool_text(self, name: str, arguments: dict | None = None, **kw) -> str:
        """Call a tool and return its concatenated text content (the common case
        — every office-connect tool replies with TextContent blocks)."""
        return self.text_content(await self.call_tool(name, arguments, **kw))

    @staticmethod
    def text_content(result: dict) -> str:
        return "\n".join(
            b.get("text", "") for b in result.get("content", []) if b.get("type") == "text"
        )


def _clean_env(extra: dict | None = None) -> dict:
    """Base environment for a spawned server, scrubbed of production markers so
    the mock safety gate can't be tripped by an inherited variable."""
    env = {**os.environ}
    for key in ("WEBSITE_INSTANCE_ID", "WEBSITE_URL", "OFFICE_CONNECT_SERVER", "FUNCTIONS_WORKER_RUNTIME"):
        env.pop(key, None)
    if extra:
        env.update(extra)
    return env


@asynccontextmanager
async def mock_server_client(
    *, permission_level: str = "all", python: str | None = None,
    extra_args: list[str] | None = None,
) -> AsyncIterator[MCPClient]:
    """Spawn the office-connect server in offline MOCK mode and yield a connected
    :class:`MCPClient`. Writes a throwaway ``"mock": true`` keyfile and sets
    ``LLMING_MOCK_USERS=1``; everything is served by the mock transport, so no
    credentials and no network are involved. The temp keyfile is removed on exit.
    """
    python = python or sys.executable
    with tempfile.TemporaryDirectory(prefix="oc-mcp-mock-") as tmp:
        keyfile = Path(tmp) / "mock_token.json"
        keyfile.write_text(json.dumps({
            "app": "office-connect",
            "mock": True,
            "access_token": "mock",
            "endpoint": "https://graph.microsoft.com/v1.0/",
        }))
        cmd = [
            python, "-m", "office_con.mcp_server",
            "--keyfile", str(keyfile),
            "--permission-level", permission_level,
            "--log-file", "none",  # don't write a log file during tests
            *(extra_args or []),
        ]
        # Isolate the file-server spool/port under the temp dir so a test never
        # touches (or contends with) a real ~/.config file server.
        client = MCPClient(cmd, env=_clean_env({
            "LLMING_MOCK_USERS": "1",
            "OFFICE_CONNECT_FILE_SERVER_DIR": str(Path(tmp) / "fileserver"),
        }))
        try:
            await client.start()
            yield client
        finally:
            await client.stop()


# ── CLI ──────────────────────────────────────────────────────
async def _run(client: MCPClient, args) -> int:
    if args.command == "list":
        tools = await client.list_tools()
        print(f"# {len(tools)} tools (server: {client.server_info.get('name', '?')})")
        for t in tools:
            desc = (t.get("description") or "").split("\n", 1)[0]
            print(f"  {t['name']:<32} {desc[:90]}")
        return 0
    if args.command == "call":
        arguments = json.loads(args.json_args) if args.json_args else {}
        result = await client.call_tool(args.tool, arguments)
        text = client.text_content(result)
        print(text if text else json.dumps(result, indent=2))
        return 0
    print(f"unknown command: {args.command}", file=sys.stderr)
    return 2


async def _cli_async(args) -> int:
    if args.mock:
        async with mock_server_client(permission_level=args.permission_level) as client:
            return await _run(client, args)
    if not args.keyfile:
        print("error: provide --mock or --keyfile PATH", file=sys.stderr)
        return 2
    cmd = [
        sys.executable, "-m", "office_con.mcp_server",
        "--keyfile", args.keyfile,
        "--permission-level", args.permission_level,
        "--log-file", "none",
    ]
    async with MCPClient(cmd, env=_clean_env()) as client:
        return await _run(client, args)


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(
        prog="office_con.testing.mcp_client",
        description="Drive the office-connect MCP server over stdio (test client).",
    )
    parser.add_argument("--mock", action="store_true", help="Spawn the server in offline MOCK mode (synthetic data)")
    parser.add_argument("--keyfile", help="Token keyfile for a real server (ignored with --mock)")
    parser.add_argument("--permission-level", default="all", choices=["read_only", "drafts", "all"])
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("list", help="List the tools the server exposes")
    call = sub.add_parser("call", help="Call a tool")
    call.add_argument("tool", help="Tool name, e.g. o365_get_mail")
    call.add_argument("json_args", nargs="?", default="", help="Arguments as a JSON object")

    args = parser.parse_args(argv)
    return asyncio.run(_cli_async(args))


if __name__ == "__main__":
    raise SystemExit(main())

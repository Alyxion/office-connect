# Testing the MCP server end-to-end

office-connect ships its own tiny MCP **client** so you can drive the stdio
server exactly the way Claude Desktop would — over JSON-RPC on a subprocess's
stdin/stdout — but against **synthetic mock data**, with no Office 365
credentials and no network. This is what lets us catch tool-layer bugs (e.g. an
attachment fetch that times out, or a `$select` that Graph rejects) that
handler-level unit tests can miss.

Two pieces make this work:

1. **Offline mock mode in the server** — a keyfile flag that makes the server
   serve `MockGraphTransport` data instead of calling Graph.
2. **`office_con.testing.mcp_client`** — a dependency-light client (stdlib
   `asyncio` only) that spawns the server and speaks the MCP protocol.

---

## 1. Offline mock mode

`_create_graph` activates mock mode when the keyfile contains `"mock": true`.
It is **triple-gated and cannot run in production**:

| gate | requirement |
|---|---|
| keyfile | `{"mock": true}` |
| environment | `LLMING_MOCK_USERS=1` |
| not production | no Azure App Service / `OFFICE_CONNECT_SERVER` / Functions markers (`is_mock_enabled()`) |

If the keyfile asks for mock mode but the env gate is not satisfied, the server
raises rather than silently falling back to a real (unauthenticated) graph.

A mock keyfile is just:

```json
{ "app": "office-connect", "mock": true, "access_token": "mock",
  "endpoint": "https://graph.microsoft.com/v1.0/" }
```

Launch the server in mock mode manually if you like:

```bash
LLMING_MOCK_USERS=1 office-connect --keyfile mock_token.json \
  --permission-level all --log-file none
```

…but you normally don't — the client's `mock_server_client()` helper writes a
throwaway keyfile and sets the env flag for you.

---

## 2. The client

### CLI

```bash
# List the tools the server exposes (mock mode)
python -m office_con.testing.mcp_client --mock list

# Call a tool — arguments are a JSON object
python -m office_con.testing.mcp_client --mock call o365_check_connection
python -m office_con.testing.mcp_client --mock call o365_search_mail \
  '{"body":"hotel","since":"2025-01-01","until":"2025-12-31","limit":5}'

# Narrow the permission tier the server runs at
python -m office_con.testing.mcp_client --mock --permission-level read_only list

# Point at a REAL server / keyfile instead of mock (drives your live mailbox)
python -m office_con.testing.mcp_client \
  --keyfile ~/.config/office-connect/token.json call o365_get_profile
```

The `call` output is the tool's text content (the common case); if a tool
returns non-text blocks the raw JSON-RPC result is printed instead.

### Library / pytest

```python
import json
from office_con.testing.mcp_client import mock_server_client

async def example():
    async with mock_server_client(permission_level="all") as client:
        tools = await client.list_tools()

        # call_tool returns the raw result; text_content extracts text blocks
        mail = json.loads(
            client.text_content(await client.call_tool("o365_get_mail",
                                                        {"email_id": some_id})))

        # or the shortcut for the text-only common case
        text = await client.call_tool_text("o365_check_connection")
```

Drive a **real** server (e.g. against an exported token) by constructing the
client directly:

```python
import sys
from office_con.testing.mcp_client import MCPClient

cmd = [sys.executable, "-m", "office_con.mcp_server",
       "--keyfile", "/path/to/token.json", "--permission-level", "read_only",
       "--log-file", "none"]
async with MCPClient(cmd) as client:
    print(await client.list_tools())
```

### API surface

| method | purpose |
|---|---|
| `await client.list_tools()` | the server's `tools/list` (already permission-filtered) |
| `await client.call_tool(name, args, timeout=60)` | `tools/call`; returns the raw result dict |
| `await client.call_tool_text(name, args)` | call + extract text content |
| `MCPClient.text_content(result)` | join the text blocks of a result |
| `client.server_info` | `serverInfo` from the initialize handshake |
| `client.stderr_tail(n)` | last *n* server stderr lines (handy on failures) |

`MCPError` is raised only for **protocol-level** problems (JSON-RPC error,
timeout, dead server). Tool-level failures come back as ordinary result content
— inspect the returned text, don't expect an exception.

---

## 3. The automated e2e suite

`tests/test_mcp_client_e2e.py` spawns the real server in mock mode and asserts
the tool layer behaves — including the regressions this client was built to
guard:

```bash
poetry run pytest tests/test_mcp_client_e2e.py -v
```

It covers the initialize handshake, permission-tier filtering, `o365_get_mail`
returning attachment **metadata** (id/size) without inlining bytes,
`o365_peek_mail_attachment` fetching bytes by id (and degrading gracefully on an
unparseable file), and `o365_search_mail` with a date range routing through
`/search/query`.

Each test spawns a fresh ~1-2s subprocess; no credentials or network required.

---

## Safety

Mock mode never activates without `LLMING_MOCK_USERS=1`, and `is_mock_enabled()`
hard-blocks it on any Azure App Service / production environment. The
`mock_server_client` helper additionally scrubs production markers from the
spawned server's environment so an inherited variable can't trip the gate. None
of this can cause a real, credentialed server to serve synthetic data — the
keyfile flag alone is inert without the env gate.

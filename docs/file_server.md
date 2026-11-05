# Downloading mail attachments (base64 or local URL)

`o365_get_mail` lists a message's attachments as **metadata only** (id, name,
size) — it never inlines the bytes, because doing so used to time out on large
PDFs. To get the actual bytes (e.g. to re-attach a mailed receipt to a new
draft) use **`o365_get_mail_attachment`**.

It has two delivery modes:

| `delivery` | result | use when |
|---|---|---|
| `base64` | `content_base64` in the tool result | small files; you want the bytes in context to feed straight into a draft — **or your fetcher is network-isolated and cannot reach the loopback URL** (see below) |
| `url` | an unguessable loopback `http://127.0.0.1:…` URL | larger files, or to avoid putting the payload in the model context — a local fetcher (Claude Code / cowork) GETs the URL |
| `auto` *(default)* | `url` if the file server is enabled, else `base64` | — |

> **`base64` is always available — the `max_bytes` cap is the guardrail, not a
> disable.** A multi-MB blob in context is what bloats the model, so over
> `max_bytes` (default 10 MB) the base64 path *refuses* and points you at the URL.
> But a small file (e.g. a 30 KB receipt) is fine inline, and base64 is the **only**
> way to reach a fetcher that can't open the loopback URL — a sandboxed/Docker
> client (e.g. cowork) sees `127.0.0.1` as *itself* and `host.docker.internal` may
> be blocked by its network allowlist. For such a client: use `delivery:"base64"`,
> or land the bytes on a filesystem path shared with the sandbox (a configured
> `--attachment-root` that is also mounted into it) and attach by `path`.

```jsonc
// o365_get_mail_attachment {"email_id": "...", "attachment_id": "..."}
// → with --file-server:
{ "name": "Receipt-2543.pdf", "content_type": "application/pdf", "size": 29571,
  "delivery": "url",
  "url": "http://127.0.0.1:54123/f/J8x…43-char-token…/Receipt-2543.pdf",
  "expires_in_seconds": 3600 }
```

The base64 path is capped by `max_bytes` (default 10 MB) to protect context;
over it the tool refuses and points you at `--file-server` or
`o365_peek_mail_attachment`.

### Re-attach a receipt to a new draft

* **base64:** pass the returned `content_base64` (+ `name`, `content_type`)
  straight into `o365_create_mail_draft` / `o365_update_mail_draft` `attachments`.
* **url:** fetch the URL to obtain the bytes, then attach them (or, if an
  attachment root is configured, save and attach by `path`).

---

## The local file server

When enabled, the MCP server runs a tiny HTTP server **bound to `127.0.0.1`**
that hands out published byte payloads at unguessable URLs. It is **off by
default**.

### One shared server across every client (singleton)

Each MCP client (Claude Desktop, Claude Code, cowork) launches its **own**
`office_con.mcp_server` process, and each may pass `--file-server`. They do **not**
each spin up a separate listener fighting over ports. Instead every file-server
process that points at the same state dir (default
`~/.config/office-connect/file-server/`) shares **one** listener and **one**
on-disk spool:

* The **bound TCP port is the mutex** — exactly one process can hold the listener
  at a time (the OS rejects a second bind). The others run *publisher-only*.
* The **spool is shared on disk** (`spool/<token>.bin` + `.json`, dir `0700`,
  files `0600`). Any process publishes by writing there; the listener serves it —
  so a publisher's URL works even though it never opened a socket.
* It **lives as long as any file-server process is alive.** A heartbeat re-binds
  the port the instant the listener exits, so a survivor takes over on the *same*
  port (issued URLs keep working). When the last one exits, the listener is gone.

Override the shared dir with `OFFICE_CONNECT_FILE_SERVER_DIR` (mainly for tests).

> **A sandboxed client still can't reach it.** The singleton is about not running
> N servers — it does **not** make the loopback URL reachable from a
> network-isolated client. cowork-in-Docker must use `delivery:"base64"` or a
> mounted attachment-root path; the URL is for fetchers that share the host's
> network namespace (Claude Code on the host).

### Enable it

```bash
office-connect --keyfile token.json --file-server
# choose a fixed port / TTL (defaults: OS-assigned port, 3600s):
office-connect --keyfile token.json --file-server --file-server-port 54123 --file-server-ttl 1800
```

Environment equivalents:

| flag | env var |
|---|---|
| `--file-server` | `OFFICE_CONNECT_FILE_SERVER=1` |
| `--file-server-port` | `OFFICE_CONNECT_FILE_SERVER_PORT` |
| `--file-server-ttl` | `OFFICE_CONNECT_FILE_SERVER_TTL` |
| *(shared spool dir)* | `OFFICE_CONNECT_FILE_SERVER_DIR` |

### Security model

* **Loopback only.** Binds `127.0.0.1` — never reachable from another host.
* **Unguessable URLs.** Each file is addressed by a `secrets.token_urlsafe(32)`
  token (~256 bits). The path token is the only credential; an unknown or expired
  token returns a bare `404` with no hint. Brute force is infeasible.
* **Expiry + one-shot.** Published URLs expire after the TTL (default 1 h) and the
  spool is bounded (oldest evicted past `max_entries`); the listener sweeps expired
  files. Spool bytes live on local disk at `0600` and are unlinked on expiry/fetch.
* **Never on a server.** `is_server_environment()` hard-blocks the file server in
  any hosted/Azure environment regardless of the flag — mailbox bytes are never
  exposed over HTTP **or spooled to disk** on a shared host. Not overridable.
* **Privacy filters still apply.** The same hidden/blocked name and content rules
  that gate `o365_peek_mail_attachment` gate this tool before any bytes are
  published.

### Caveat: who fetches the URL?

The URL is only reachable from the **same machine** as the MCP server. This works
when the fetcher runs locally (Claude Code / cowork on your machine). A
server-side fetcher (e.g. a cloud web client) cannot reach `127.0.0.1` on your
host — use `delivery='base64'` there.

See [Testing the MCP server](mcp_testing.md) to exercise all of this offline via
the mock client.

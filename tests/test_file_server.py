"""Unit tests for the loopback file server used to hand attachment bytes to a
local fetcher via an unguessable URL. Each test gets an isolated state dir so the
shared on-disk spool never leaks between tests or parallel runs."""

import asyncio
import json
import urllib.error
import urllib.request

import pytest

from office_con.file_server import LocalFileServer, _safe_header_filename


def _get(url: str):
    r = urllib.request.urlopen(url)
    return r.read(), dict(r.headers)


def _token_of(url: str) -> str:
    return url.split("/f/")[1].split("/")[0]


@pytest.mark.asyncio
async def test_publish_and_fetch_roundtrip(tmp_path):
    fs = await LocalFileServer(ttl_seconds=60, state_dir=tmp_path).start()
    try:
        pub = fs.publish(b"receipt-bytes", name="r.pdf", content_type="application/pdf")
        assert pub["bytes"] == len(b"receipt-bytes")
        assert pub["url"].startswith("http://127.0.0.1:")
        body, headers = await asyncio.to_thread(_get, pub["url"])
        assert body == b"receipt-bytes"
        assert headers["Content-Type"] == "application/pdf"
        assert 'filename="r.pdf"' in headers["Content-Disposition"]
    finally:
        await fs.stop()


@pytest.mark.asyncio
async def test_unknown_token_404(tmp_path):
    fs = await LocalFileServer(state_dir=tmp_path).start()
    try:
        with pytest.raises(urllib.error.HTTPError) as exc:
            await asyncio.to_thread(_get, f"{fs.base_url}/f/notarealtoken/x")
        assert exc.value.code == 404
    finally:
        await fs.stop()


@pytest.mark.asyncio
async def test_one_shot_is_deleted_after_first_fetch(tmp_path):
    fs = await LocalFileServer(state_dir=tmp_path).start()
    try:
        pub = fs.publish(b"x", name="x", one_shot=True)
        await asyncio.to_thread(_get, pub["url"])  # first ok
        with pytest.raises(urllib.error.HTTPError) as exc:
            await asyncio.to_thread(_get, pub["url"])  # second gone
        assert exc.value.code == 404
    finally:
        await fs.stop()


@pytest.mark.asyncio
async def test_expired_entry_404(tmp_path):
    fs = await LocalFileServer(ttl_seconds=1, state_dir=tmp_path).start()
    try:
        pub = fs.publish(b"x", name="x")
        # Force expiry without sleeping: rewrite the on-disk sidecar's deadline.
        token = _token_of(pub["url"])
        meta_path = fs._spool / f"{token}.json"
        meta = json.loads(meta_path.read_text())
        meta["expires"] = 0.0
        meta_path.write_text(json.dumps(meta))
        with pytest.raises(urllib.error.HTTPError) as exc:
            await asyncio.to_thread(_get, pub["url"])
        assert exc.value.code == 404
    finally:
        await fs.stop()


@pytest.mark.asyncio
async def test_max_entries_evicts_oldest(tmp_path):
    fs = await LocalFileServer(max_entries=2, state_dir=tmp_path).start()
    try:
        a = fs.publish(b"a", name="a")
        fs.publish(b"b", name="b")
        fs.publish(b"c", name="c")  # evicts the oldest (a)
        assert len(list(fs._spool.glob("*.json"))) == 2
        with pytest.raises(urllib.error.HTTPError):
            await asyncio.to_thread(_get, a["url"])
    finally:
        await fs.stop()


@pytest.mark.asyncio
async def test_tokens_are_unguessable_and_unique(tmp_path):
    fs = await LocalFileServer(state_dir=tmp_path).start()
    try:
        tokens = set()
        for _ in range(50):
            token = _token_of(fs.publish(b"x", name="x")["url"])
            assert len(token) >= 32  # ~256 bits of entropy
            tokens.add(token)
        assert len(tokens) == 50  # no collisions
    finally:
        await fs.stop()


@pytest.mark.asyncio
async def test_second_process_shares_port_and_spool(tmp_path):
    """A second file server on the same state dir cannot bind the port (the OS
    mutex) — it attaches as a publisher, reuses the listener's port, and what it
    publishes is served by the listener. This is the cross-process singleton."""
    listener = await LocalFileServer(state_dir=tmp_path).start()
    publisher = await LocalFileServer(state_dir=tmp_path).start()
    try:
        # Same shared port; only the first actually owns the socket.
        assert publisher.base_url == listener.base_url
        assert listener._is_listener is True
        assert publisher._is_listener is False
        # A publish from the non-listener is fetchable via the shared listener.
        pub = publisher.publish(b"shared", name="s.pdf", content_type="application/pdf")
        body, _ = await asyncio.to_thread(_get, pub["url"])
        assert body == b"shared"
    finally:
        await publisher.stop()
        await listener.stop()


@pytest.mark.asyncio
async def test_publisher_takes_over_when_listener_dies(tmp_path):
    """When the listener exits, a surviving process re-binds the SAME port so
    already-issued URLs keep working — the server lives as long as any process
    needs it."""
    listener = await LocalFileServer(state_dir=tmp_path).start()
    publisher = await LocalFileServer(state_dir=tmp_path).start()
    port = listener._port
    pub = publisher.publish(b"survives", name="s.pdf")
    try:
        await listener.stop()  # listener gone; port freed
        # Re-election (normally driven by the heartbeat) — invoke directly so the
        # test doesn't wait a full heartbeat interval.
        assert await publisher._try_listen(port) is True
        assert publisher._is_listener is True
        assert publisher._port == port  # same port → the old URL still resolves
        body, _ = await asyncio.to_thread(_get, pub["url"])
        assert body == b"survives"
    finally:
        await publisher.stop()


def test_header_filename_is_sanitized():
    assert _safe_header_filename('a"b\r\nc') == "abc"
    assert _safe_header_filename(None) == "file"
    assert _safe_header_filename("") == "file"

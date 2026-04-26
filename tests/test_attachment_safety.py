"""Safety tests for path-based attachments / body_path.

Covers the full escape-catching surface: disabled by default, null bytes,
nonexistent paths, directory/device paths, symlink escapes, and size caps.
"""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from office_con.mcp_server import (
    ATTACHMENT_ROOTS_ENV,
    AttachmentPathError,
    _decode_attachments,
    _parse_attachment_roots,
    _resolve_body_text,
    _resolve_safe_attachment_path,
)


# ---------------------------------------------------------------------------
# _parse_attachment_roots
# ---------------------------------------------------------------------------


class TestParseAttachmentRoots:

    def test_empty_by_default(self, monkeypatch):
        monkeypatch.delenv(ATTACHMENT_ROOTS_ENV, raising=False)
        assert _parse_attachment_roots(None) == []
        assert _parse_attachment_roots([]) == []

    def test_cli_roots(self, tmp_path, monkeypatch):
        monkeypatch.delenv(ATTACHMENT_ROOTS_ENV, raising=False)
        roots = _parse_attachment_roots([str(tmp_path)])
        assert len(roots) == 1
        assert roots[0] == tmp_path.resolve()

    def test_env_roots(self, tmp_path, monkeypatch):
        second = tmp_path / "second"
        second.mkdir()
        monkeypatch.setenv(
            ATTACHMENT_ROOTS_ENV,
            os.pathsep.join([str(tmp_path), str(second)]),
        )
        roots = _parse_attachment_roots(None)
        assert len(roots) == 2

    def test_rejects_relative_root(self, monkeypatch):
        monkeypatch.delenv(ATTACHMENT_ROOTS_ENV, raising=False)
        with pytest.raises(ValueError, match="absolute"):
            _parse_attachment_roots(["relative/path"])

    def test_rejects_nonexistent_root(self, tmp_path, monkeypatch):
        monkeypatch.delenv(ATTACHMENT_ROOTS_ENV, raising=False)
        with pytest.raises(ValueError, match="does not exist"):
            _parse_attachment_roots([str(tmp_path / "nope")])

    def test_rejects_file_as_root(self, tmp_path, monkeypatch):
        monkeypatch.delenv(ATTACHMENT_ROOTS_ENV, raising=False)
        f = tmp_path / "file.txt"
        f.write_text("x")
        with pytest.raises(ValueError, match="not a directory"):
            _parse_attachment_roots([str(f)])


# ---------------------------------------------------------------------------
# _resolve_safe_attachment_path
# ---------------------------------------------------------------------------


class TestResolveSafeAttachmentPath:

    MAX = 10 * 1024 * 1024

    def test_disabled_when_no_roots(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_text("x")
        with pytest.raises(AttachmentPathError, match="DISABLED"):
            _resolve_safe_attachment_path(str(f), [], self.MAX)

    def test_rejects_empty_path(self, tmp_path):
        with pytest.raises(AttachmentPathError, match="empty"):
            _resolve_safe_attachment_path("", [tmp_path], self.MAX)
        with pytest.raises(AttachmentPathError, match="empty"):
            _resolve_safe_attachment_path("   ", [tmp_path], self.MAX)

    def test_rejects_null_byte(self, tmp_path):
        with pytest.raises(AttachmentPathError, match="null byte"):
            _resolve_safe_attachment_path(
                str(tmp_path / "a\x00b.txt"), [tmp_path], self.MAX,
            )

    def test_rejects_nonexistent_file(self, tmp_path):
        with pytest.raises(AttachmentPathError, match="not found"):
            _resolve_safe_attachment_path(
                str(tmp_path / "ghost.txt"), [tmp_path], self.MAX,
            )

    def test_rejects_directory(self, tmp_path):
        sub = tmp_path / "subdir"
        sub.mkdir()
        with pytest.raises(AttachmentPathError, match="not a regular file"):
            _resolve_safe_attachment_path(str(sub), [tmp_path], self.MAX)

    def test_rejects_path_outside_root(self, tmp_path):
        # Create two separate directories; root covers only one.
        root = tmp_path / "root"
        other = tmp_path / "other"
        root.mkdir()
        other.mkdir()
        outside = other / "leak.txt"
        outside.write_text("secret")
        with pytest.raises(AttachmentPathError, match="outside every configured"):
            _resolve_safe_attachment_path(str(outside), [root], self.MAX)

    def test_accepts_file_inside_root(self, tmp_path):
        f = tmp_path / "ok.txt"
        f.write_text("hello")
        resolved = _resolve_safe_attachment_path(str(f), [tmp_path], self.MAX)
        assert resolved == f.resolve()

    def test_rejects_dotdot_escape(self, tmp_path):
        # Build a path that LOOKS inside the root but resolves outside.
        root = tmp_path / "root"
        root.mkdir()
        secret = tmp_path / "secret.txt"
        secret.write_text("shh")
        sneaky = root / ".." / "secret.txt"
        with pytest.raises(AttachmentPathError, match="outside every configured"):
            _resolve_safe_attachment_path(str(sneaky), [root], self.MAX)

    def test_rejects_symlink_pointing_outside_root(self, tmp_path):
        root = tmp_path / "root"
        root.mkdir()
        secret = tmp_path / "secret.txt"
        secret.write_text("shh")
        link = root / "link.txt"
        link.symlink_to(secret)
        with pytest.raises(AttachmentPathError, match="outside every configured"):
            _resolve_safe_attachment_path(str(link), [root], self.MAX)

    def test_accepts_symlink_inside_root(self, tmp_path):
        target = tmp_path / "real.txt"
        target.write_text("x")
        link = tmp_path / "ln.txt"
        link.symlink_to(target)
        resolved = _resolve_safe_attachment_path(str(link), [tmp_path], self.MAX)
        assert resolved == target.resolve()

    def test_size_cap(self, tmp_path):
        f = tmp_path / "big.bin"
        f.write_bytes(b"x" * 2048)
        with pytest.raises(AttachmentPathError, match="exceeds"):
            _resolve_safe_attachment_path(str(f), [tmp_path], max_bytes=1024)


# ---------------------------------------------------------------------------
# _decode_attachments — branch coverage for path vs base64
# ---------------------------------------------------------------------------


class TestDecodeAttachments:

    def test_none_passthrough(self, tmp_path):
        assert _decode_attachments(
            None, attachment_roots=[tmp_path], max_attachment_bytes=1024,
        ) is None

    def test_base64_mode_works_without_roots(self):
        import base64
        data = b"hello"
        out = _decode_attachments(
            [{"name": "a.txt", "content_base64": base64.b64encode(data).decode()}],
            attachment_roots=[],  # disabled for path mode, but base64 still works
            max_attachment_bytes=1024,
        )
        assert len(out) == 1
        assert out[0].content_bytes == data

    def test_path_mode_requires_roots(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_text("x")
        with pytest.raises(AttachmentPathError, match="DISABLED"):
            _decode_attachments(
                [{"name": "a.txt", "path": str(f)}],
                attachment_roots=[],
                max_attachment_bytes=1024,
            )

    def test_path_mode_reads_file(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_bytes(b"hello")
        out = _decode_attachments(
            [{"name": "a.txt", "path": str(f)}],
            attachment_roots=[tmp_path],
            max_attachment_bytes=1024,
        )
        assert out[0].content_bytes == b"hello"

    def test_rejects_both_path_and_base64(self, tmp_path):
        f = tmp_path / "a.txt"
        f.write_text("x")
        with pytest.raises(ValueError, match="exactly one"):
            _decode_attachments(
                [{"name": "a.txt", "path": str(f), "content_base64": "eA=="}],
                attachment_roots=[tmp_path],
                max_attachment_bytes=1024,
            )

    def test_rejects_neither(self, tmp_path):
        with pytest.raises(ValueError, match="either"):
            _decode_attachments(
                [{"name": "a.txt"}],
                attachment_roots=[tmp_path],
                max_attachment_bytes=1024,
            )

    def test_rejects_missing_name(self, tmp_path):
        with pytest.raises(ValueError, match="name"):
            _decode_attachments(
                [{"content_base64": "eA=="}],
                attachment_roots=[],
                max_attachment_bytes=1024,
            )

    def test_rejects_invalid_base64(self):
        with pytest.raises(ValueError, match="invalid base64"):
            _decode_attachments(
                [{"name": "a.txt", "content_base64": "!!!not base64!!!"}],
                attachment_roots=[],
                max_attachment_bytes=1024,
            )


# ---------------------------------------------------------------------------
# _resolve_body_text
# ---------------------------------------------------------------------------


class TestResolveBodyText:

    def test_inline_body(self, tmp_path):
        out = _resolve_body_text(
            {"body": "hello"},
            attachment_roots=[tmp_path], max_attachment_bytes=1024,
        )
        assert out == "hello"

    def test_body_path_reads_file(self, tmp_path):
        f = tmp_path / "body.html"
        f.write_text("<p>hello</p>", encoding="utf-8")
        out = _resolve_body_text(
            {"body_path": str(f)},
            attachment_roots=[tmp_path], max_attachment_bytes=1024,
        )
        assert "<p>hello</p>" in out

    def test_body_path_rejected_without_roots(self, tmp_path):
        f = tmp_path / "body.html"
        f.write_text("x")
        with pytest.raises(AttachmentPathError, match="DISABLED"):
            _resolve_body_text(
                {"body_path": str(f)},
                attachment_roots=[], max_attachment_bytes=1024,
            )

    def test_rejects_both(self, tmp_path):
        f = tmp_path / "b.txt"
        f.write_text("x")
        with pytest.raises(ValueError, match="exactly one"):
            _resolve_body_text(
                {"body": "x", "body_path": str(f)},
                attachment_roots=[tmp_path], max_attachment_bytes=1024,
            )

    def test_rejects_neither(self):
        with pytest.raises(ValueError, match="missing"):
            _resolve_body_text(
                {}, attachment_roots=[], max_attachment_bytes=1024,
            )

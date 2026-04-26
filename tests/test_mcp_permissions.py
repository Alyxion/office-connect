"""Unit tests for the MCP server's permission model.

These tests run without a Graph token — they verify only the classification
logic, the fail-closed default, the list/call gating, and config resolution.
"""

from __future__ import annotations

import os
from unittest.mock import patch

import pytest
from mcp.types import Tool

from office_con.mcp_permissions import (
    DEFAULT_LEVEL,
    ENV_VAR,
    PermissionLevel,
    level_allows,
    parse_level,
    resolve_level,
)
from office_con.mcp_server import (
    TOOL_PERMISSIONS,
    TOOLS,
    PermissionDenied,
    _require_allowed,
    create_server,
    filter_tools,
)


# ---------------------------------------------------------------------------
# PermissionLevel helpers
# ---------------------------------------------------------------------------


class TestPermissionLevel:

    def test_default_is_drafts(self):
        assert DEFAULT_LEVEL is PermissionLevel.DRAFTS

    @pytest.mark.parametrize(
        "configured,required,expected",
        [
            (PermissionLevel.READ_ONLY, PermissionLevel.READ_ONLY, True),
            (PermissionLevel.READ_ONLY, PermissionLevel.DRAFTS, False),
            (PermissionLevel.READ_ONLY, PermissionLevel.ALL, False),
            (PermissionLevel.DRAFTS, PermissionLevel.READ_ONLY, True),
            (PermissionLevel.DRAFTS, PermissionLevel.DRAFTS, True),
            (PermissionLevel.DRAFTS, PermissionLevel.ALL, False),
            (PermissionLevel.ALL, PermissionLevel.READ_ONLY, True),
            (PermissionLevel.ALL, PermissionLevel.DRAFTS, True),
            (PermissionLevel.ALL, PermissionLevel.ALL, True),
        ],
    )
    def test_level_allows(self, configured, required, expected):
        assert level_allows(required, configured) is expected

    def test_parse_level_known_values(self):
        assert parse_level("read_only") is PermissionLevel.READ_ONLY
        assert parse_level("drafts") is PermissionLevel.DRAFTS
        assert parse_level("all") is PermissionLevel.ALL

    def test_parse_level_trims_and_lowercases(self):
        assert parse_level(" READ_ONLY ") is PermissionLevel.READ_ONLY

    def test_parse_level_empty_returns_default(self):
        assert parse_level(None) is DEFAULT_LEVEL
        assert parse_level("") is DEFAULT_LEVEL

    def test_parse_level_invalid_raises(self):
        with pytest.raises(ValueError, match="Invalid permission level"):
            parse_level("god_mode")


class TestResolveLevel:

    def test_cli_beats_env(self):
        with patch.dict(os.environ, {ENV_VAR: "all"}, clear=False):
            assert resolve_level("read_only") is PermissionLevel.READ_ONLY

    def test_env_used_when_cli_missing(self):
        with patch.dict(os.environ, {ENV_VAR: "all"}, clear=False):
            assert resolve_level(None) is PermissionLevel.ALL

    def test_default_when_neither_set(self):
        env = {k: v for k, v in os.environ.items() if k != ENV_VAR}
        with patch.dict(os.environ, env, clear=True):
            assert resolve_level(None) is DEFAULT_LEVEL


# ---------------------------------------------------------------------------
# Classification registry
# ---------------------------------------------------------------------------


class TestToolClassification:

    def test_every_advertised_tool_is_classified(self):
        """Every tool in TOOLS must have an entry in TOOL_PERMISSIONS."""
        unclassified = {t.name for t in TOOLS} - set(TOOL_PERMISSIONS)
        assert not unclassified, (
            f"Tools missing a permission classification: {sorted(unclassified)}. "
            "Add them to TOOL_PERMISSIONS in office_con/mcp_server.py — "
            "unclassified tools are denied by default."
        )

    def test_no_orphan_classifications(self):
        """TOOL_PERMISSIONS entries must correspond to real tools."""
        orphans = set(TOOL_PERMISSIONS) - {t.name for t in TOOLS}
        assert not orphans, f"Classification for unknown tools: {sorted(orphans)}"

    def test_all_draft_mail_tools_are_drafts_tier(self):
        assert TOOL_PERMISSIONS["o365_create_mail_draft"] is PermissionLevel.DRAFTS
        assert TOOL_PERMISSIONS["o365_update_mail_draft"] is PermissionLevel.DRAFTS

    def test_sending_is_all_tier(self):
        for name in ("o365_send_mail", "o365_send_mail_draft"):
            assert TOOL_PERMISSIONS[name] is PermissionLevel.ALL

    def test_destructive_mail_ops_are_all_tier(self):
        for name in ("o365_delete_mail", "o365_move_mail",
                     "o365_flag_mail_read", "o365_set_mail_categories"):
            assert TOOL_PERMISSIONS[name] is PermissionLevel.ALL

    def test_create_event_is_all_tier(self):
        assert TOOL_PERMISSIONS["o365_create_event"] is PermissionLevel.ALL


# ---------------------------------------------------------------------------
# filter_tools / _require_allowed
# ---------------------------------------------------------------------------


class TestFilterTools:

    def test_read_only_hides_drafts_and_all(self):
        names = {t.name for t in filter_tools(PermissionLevel.READ_ONLY)}
        assert "o365_list_mail" in names
        assert "o365_create_mail_draft" not in names
        assert "o365_send_mail" not in names
        assert "o365_create_event" not in names

    def test_drafts_shows_read_and_drafts_but_hides_all(self):
        names = {t.name for t in filter_tools(PermissionLevel.DRAFTS)}
        assert "o365_list_mail" in names
        assert "o365_create_mail_draft" in names
        assert "o365_update_mail_draft" in names
        assert "o365_send_mail" not in names
        assert "o365_send_mail_draft" not in names
        assert "o365_delete_mail" not in names
        assert "o365_create_event" not in names

    def test_all_exposes_everything_classified(self):
        names = {t.name for t in filter_tools(PermissionLevel.ALL)}
        assert names == set(TOOL_PERMISSIONS)

    def test_filter_returns_tool_instances(self):
        for tool in filter_tools(PermissionLevel.READ_ONLY):
            assert isinstance(tool, Tool)


class TestRequireAllowed:

    def test_read_only_denies_draft_tool(self):
        with pytest.raises(PermissionDenied):
            _require_allowed("o365_create_mail_draft", PermissionLevel.READ_ONLY)

    def test_drafts_allows_create_draft(self):
        _require_allowed("o365_create_mail_draft", PermissionLevel.DRAFTS)

    def test_drafts_denies_send(self):
        with pytest.raises(PermissionDenied):
            _require_allowed("o365_send_mail", PermissionLevel.DRAFTS)

    def test_drafts_denies_delete(self):
        with pytest.raises(PermissionDenied):
            _require_allowed("o365_delete_mail", PermissionLevel.DRAFTS)

    def test_drafts_denies_create_event(self):
        with pytest.raises(PermissionDenied):
            _require_allowed("o365_create_event", PermissionLevel.DRAFTS)

    def test_all_allows_send(self):
        _require_allowed("o365_send_mail", PermissionLevel.ALL)

    def test_unknown_tool_denied_at_every_level(self):
        for lvl in PermissionLevel:
            with pytest.raises(PermissionDenied, match="not classified"):
                _require_allowed("o365_bogus_tool", lvl)


# ---------------------------------------------------------------------------
# create_server wiring
# ---------------------------------------------------------------------------


class TestCreateServer:

    def test_respects_explicit_level(self, tmp_path):
        # Keyfile is never read until first tool call, so a dummy path is fine.
        keyfile = str(tmp_path / "k.json")
        server, _ = create_server(keyfile, permission_level=PermissionLevel.READ_ONLY)
        assert server is not None

    def test_env_var_respected(self, tmp_path):
        keyfile = str(tmp_path / "k.json")
        with patch.dict(os.environ, {ENV_VAR: "read_only"}, clear=False):
            server, _ = create_server(keyfile)
            assert server is not None

    def test_invalid_env_raises(self, tmp_path):
        keyfile = str(tmp_path / "k.json")
        with patch.dict(os.environ, {ENV_VAR: "superuser"}, clear=False):
            with pytest.raises(ValueError, match="Invalid permission level"):
                create_server(keyfile)

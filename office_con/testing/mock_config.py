"""Optional configuration for mock data and integration tests.

Loads a JSON config file to customize mock data (room names, people, etc.)
and provide credentials for integration tests against real MS Graph.

Test-config file locations (first found wins):
1. Path in ``OFFICE_CONNECT_TEST_CONFIG`` env var
2. ``tests/test_config.json`` relative to the package root
3. ``~/.office-connect/test_config.json``

If no test-config file is found, defaults are used and config-dependent
integration tests are skipped.

Token resolution for integration tests (first found wins):
1. ``token_file`` path inside the test-config, when present and readable
2. ``~/.config/office-connect/token.json`` — the canonical MCP keyfile written
   by ``office-connect login``. Picked up automatically so a local developer
   does not need to maintain a parallel ``token_file`` entry.

Test-config file format::

    {
        "token_file": "/path/to/token_export.json",   # optional — falls back to canonical keyfile
        "rooms": ["Room A", "Room B"],
        "room_lists": ["Building 1"],
        "expected_teams": ["Team Alpha"],
        "expected_presence_users": ["user@example.com"]
    }
"""

from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

_config: dict[str, Any] | None = None
_loaded = False

# Canonical MCP keyfile, written by ``office-connect login``. Used as a
# zero-config fallback when integration tests are run without a test_config.json
# pointing at a separate token file.
DEFAULT_MCP_KEYFILE = Path.home() / ".config" / "office-connect" / "token.json"


def _find_config_path() -> Path | None:
    """Find the config file in known locations."""
    env_path = os.environ.get("OFFICE_CONNECT_TEST_CONFIG")
    if env_path:
        p = Path(env_path)
        if p.is_file():
            return p

    # Relative to package root (office-connect/)
    pkg_root = Path(__file__).resolve().parent.parent.parent
    local = pkg_root / "tests" / "test_config.json"
    if local.is_file():
        return local

    # Home directory
    home = Path.home() / ".office-connect" / "test_config.json"
    if home.is_file():
        return home

    return None


def get_test_config() -> dict[str, Any]:
    """Load and return the test config (cached)."""
    global _config, _loaded
    if _loaded:
        return _config or {}
    _loaded = True

    path = _find_config_path()
    if path is None:
        _config = {}
        return _config

    try:
        _config = json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        _config = {}
    return _config


def _read_token_file(path: Path) -> dict[str, Any] | None:
    """Read a token JSON file. Returns None if missing, malformed, or lacks
    an access_token."""
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return None
    if not isinstance(data, dict) or not data.get("access_token"):
        return None
    return data


def get_token_data() -> dict[str, Any] | None:
    """Load MS Graph token data for integration tests.

    Resolution order:
      1. ``token_file`` path inside the test_config (if set and readable)
      2. Canonical MCP keyfile at ``~/.config/office-connect/token.json``
         (written by ``office-connect login``)

    Returns the parsed JSON dict, or None if no usable token file is found.
    """
    cfg = get_test_config()
    token_file = cfg.get("token_file")
    if token_file:
        data = _read_token_file(Path(token_file).expanduser())
        if data is not None:
            return data
    # Fallback so a local developer who has run `office-connect login` doesn't
    # have to also maintain a separate test_config.json.
    return _read_token_file(DEFAULT_MCP_KEYFILE)

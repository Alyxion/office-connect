"""Optional configuration for mock data and integration tests.

Loads a JSON config file to customize mock data (room names, people, etc.)
and provide credentials for integration tests against real MS Graph.

Config file locations (first found wins):
1. Path in ``OFFICE_CONNECT_TEST_CONFIG`` env var
2. ``tests/test_config.json`` relative to the package root
3. ``~/.office-connect/test_config.json``

If no config file is found, defaults are used and integration tests are skipped.

Config file format::

    {
        "token_file": "/path/to/token_export.json",
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


def get_token_data() -> dict[str, Any] | None:
    """Load MS Graph token from the config-specified token file.

    Returns the parsed JSON dict, or None if not configured.
    """
    cfg = get_test_config()
    token_file = cfg.get("token_file")
    if not token_file:
        return None
    p = Path(token_file).expanduser()
    if not p.is_file():
        return None
    try:
        return json.loads(p.read_text())
    except (json.JSONDecodeError, OSError):
        return None

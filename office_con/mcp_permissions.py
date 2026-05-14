"""Permission tiers for the Office 365 MCP server.

Three tiers (increasing trust):

* ``READ_ONLY`` — list/get tools only; no mutation of Microsoft 365 state.
* ``DRAFTS`` — READ_ONLY plus creating and updating *draft* emails.
  No sending, no modification of non-draft messages, no calendar writes.
* ``ALL`` — DRAFTS plus sending mail, deleting/moving mail, modifying any
  message, and creating calendar events.

Three configuration sources, evaluated together — **the most restrictive
level among the sources that are set wins**. If no source is set the level
falls back to ``DRAFTS``.

1. ``--permission-level`` CLI flag (per-MCP, set in the launcher config)
2. ``$OFFICE_CONNECT_PERMISSION_LEVEL`` environment variable
3. Global policy file (default ``~/.config/office-connect/policy.json``,
   overridable via ``$OFFICE_CONNECT_POLICY`` or ``--policy-file``).

The global file is a JSON object with a ``permission_level`` (or
``max_permission_level``) field. It acts as a host-wide ceiling: any MCP
launcher requesting a *less* restrictive level is silently clamped down.

Enforcement is defense-in-depth: the MCP server filters the advertised tool
list AND re-checks on every call. Any tool name not present in the server's
classification table is denied (fail-closed).
"""

from __future__ import annotations

import json
import logging
import os
from enum import Enum
from pathlib import Path


logger = logging.getLogger(__name__)


class PermissionLevel(str, Enum):
    READ_ONLY = "read_only"
    DRAFTS = "drafts"
    ALL = "all"


_RANK: dict["PermissionLevel", int] = {
    PermissionLevel.READ_ONLY: 0,
    PermissionLevel.DRAFTS: 1,
    PermissionLevel.ALL: 2,
}

DEFAULT_LEVEL = PermissionLevel.DRAFTS
ENV_VAR = "OFFICE_CONNECT_PERMISSION_LEVEL"
POLICY_ENV_VAR = "OFFICE_CONNECT_POLICY"
DEFAULT_POLICY_FILE = "~/.config/office-connect/policy.json"


def parse_level(value: str | None) -> PermissionLevel:
    """Parse a permission-level string; raise ValueError if unknown."""
    if value is None or value == "":
        return DEFAULT_LEVEL
    try:
        return PermissionLevel(value.strip().lower())
    except ValueError:
        valid = ", ".join(l.value for l in PermissionLevel)
        raise ValueError(
            f"Invalid permission level {value!r}; must be one of: {valid}"
        ) from None


def _read_policy_file(path: str | os.PathLike[str]) -> PermissionLevel | None:
    """Read the global policy file. Returns the configured level or ``None``
    if the file is absent, unreadable, malformed, or omits the field."""
    p = Path(path).expanduser()
    if not p.is_file():
        return None
    try:
        data = json.loads(p.read_text())
    except (json.JSONDecodeError, OSError) as exc:
        logger.warning("[PERM] policy file %s unreadable (%s); ignoring", p, exc)
        return None
    if not isinstance(data, dict):
        logger.warning("[PERM] policy file %s is not a JSON object; ignoring", p)
        return None
    raw = data.get("permission_level") or data.get("max_permission_level")
    if not raw:
        return None
    try:
        return parse_level(raw)
    except ValueError as exc:
        logger.warning("[PERM] policy file %s: %s; ignoring", p, exc)
        return None


def resolve_level(cli_value: str | None = None,
                  policy_file: str | None = None) -> PermissionLevel:
    """Resolve the effective permission level from the three configuration
    sources. The most restrictive level among the sources that are actually
    set wins. If nothing is set, returns ``DEFAULT_LEVEL`` (``DRAFTS``)."""
    candidates: list[PermissionLevel] = []
    if cli_value:
        candidates.append(parse_level(cli_value))
    env = os.environ.get(ENV_VAR)
    if env:
        candidates.append(parse_level(env))
    policy_path = policy_file or os.environ.get(POLICY_ENV_VAR) or DEFAULT_POLICY_FILE
    policy_level = _read_policy_file(policy_path)
    if policy_level is not None:
        candidates.append(policy_level)
    if not candidates:
        return DEFAULT_LEVEL
    return min(candidates, key=lambda l: _RANK[l])


def level_allows(required: PermissionLevel, configured: PermissionLevel) -> bool:
    """Return True iff a tool needing ``required`` is permitted at ``configured``."""
    return _RANK[configured] >= _RANK[required]

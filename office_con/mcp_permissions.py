"""Permission tiers for the Office 365 MCP server.

Three tiers (increasing trust):

* ``READ_ONLY`` — list/get tools only; no mutation of Microsoft 365 state.
* ``DRAFTS`` — READ_ONLY plus creating and updating *draft* emails.
  No sending, no modification of non-draft messages, no calendar writes.
* ``ALL`` — DRAFTS plus sending mail, deleting/moving mail, modifying any
  message, and creating calendar events.

Configure via CLI flag ``--permission-level`` or environment variable
``OFFICE_CONNECT_PERMISSION_LEVEL``. Default is ``DRAFTS``.

Enforcement is defense-in-depth: the MCP server filters the advertised tool
list AND re-checks on every call. Any tool name not present in the server's
classification table is denied (fail-closed).
"""

from __future__ import annotations

import os
from enum import Enum


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


def resolve_level(cli_value: str | None = None) -> PermissionLevel:
    """Resolve the effective level. Priority: CLI arg > env var > default."""
    if cli_value:
        return parse_level(cli_value)
    env = os.environ.get(ENV_VAR)
    if env:
        return parse_level(env)
    return DEFAULT_LEVEL


def level_allows(required: PermissionLevel, configured: PermissionLevel) -> bool:
    """Return True iff a tool needing ``required`` is permitted at ``configured``."""
    return _RANK[configured] >= _RANK[required]

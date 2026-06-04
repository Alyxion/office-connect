"""Rotated, PII-conscious logging for the office-connect MCP server.

The server runs as a long-lived stdio process under a host (Claude Desktop,
etc.) where stderr is usually invisible. To investigate errors after the fact
we attach a *rotating file handler* to the ``office_con`` package logger.

Two design constraints:

* **Bounded disk.** ``RotatingFileHandler`` caps total size (default ~5 MB
  across 5 files) so logs never grow unbounded on a user's machine.
* **No secret/PII leakage.** Call sites already avoid logging tokens and
  redact URLs (see ``ms_graph_handler._redact_url``). As defense-in-depth a
  :class:`_RedactingFilter` masks anything that still looks like a bearer
  token, JWT, or email address before it reaches disk.

Configuration (all optional):

* ``--log-file PATH`` / ``$OFFICE_CONNECT_LOG_FILE`` — log destination.
  Default ``~/.config/office-connect/logs/office-connect.log``. ``"none"``
  (case-insensitive) or empty disables file logging entirely.
* ``--log-level LEVEL`` / ``$OFFICE_CONNECT_LOG_LEVEL`` — default ``INFO``.
"""

from __future__ import annotations

import logging
import os
import re
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

LOG_FILE_ENV = "OFFICE_CONNECT_LOG_FILE"
LOG_LEVEL_ENV = "OFFICE_CONNECT_LOG_LEVEL"
DEFAULT_LOG_FILE = "~/.config/office-connect/logs/office-connect.log"
_MAX_BYTES = 1_000_000
_BACKUP_COUNT = 5


def is_server_environment() -> bool:
    """True when running in a hosted/server deployment rather than on a user's
    own machine via the local stdio MCP CLI.

    File logging must NEVER be set up server-side: writing (even redacted) user
    data to disk on a shared/multi-tenant host is a data-privacy problem, and a
    server app already owns its own logging. Detection reuses the same Azure
    App Service signals the mock system trusts (see ``testing/__init__.py``),
    plus common container/orchestrator markers and an explicit opt-in marker a
    deployment can set defensively.

    A server deployment never calls :func:`configure_logging` (it consumes the
    library, it doesn't run ``cli()``); this is belt-and-suspenders for the case
    where someone runs the ``office-connect`` stdio server on a hosted box.
    """
    env = os.environ
    if env.get("OFFICE_CONNECT_SERVER") == "1":
        return True
    # Azure App Service / Functions.
    if env.get("WEBSITE_INSTANCE_ID") or env.get("FUNCTIONS_WORKER_RUNTIME"):
        return True
    website_url = env.get("WEBSITE_URL", "")
    if "azurewebsites.net" in website_url or (
        website_url
        and "localhost" not in website_url
        and "127.0.0.1" not in website_url
        and "0.0.0.0" not in website_url
    ):
        return True
    # Kubernetes pod.
    if env.get("KUBERNETES_SERVICE_HOST"):
        return True
    return False

# Patterns scrubbed from every record as a last line of defense. These are
# intentionally broad — a false-positive mask is harmless, a leaked token is not.
_REDACTIONS = (
    # Authorization: Bearer <token>
    (re.compile(r"(?i)\bBearer\s+[A-Za-z0-9._\-]+"), "Bearer <redacted>"),
    # Bare JWTs (header.payload.signature) — access/refresh tokens.
    (re.compile(r"\beyJ[A-Za-z0-9._\-]{10,}"), "<redacted-jwt>"),
    # Email addresses.
    (re.compile(r"\b[\w.+\-]+@[\w\-]+\.[\w.\-]+\b"), "<redacted-email>"),
)

# Module-level guard so repeated configure_logging() calls (e.g. tests, or a
# CLI subcommand that also imports the server) don't stack duplicate handlers.
_HANDLER_NAME = "office-connect-rotating-file"


class _RedactingFilter(logging.Filter):
    """Mask tokens / emails in the formatted message and args."""

    def filter(self, record: logging.LogRecord) -> bool:
        try:
            msg = record.getMessage()
        except Exception:
            return True  # never drop a record because redaction failed
        for pattern, repl in _REDACTIONS:
            msg = pattern.sub(repl, msg)
        record.msg = msg
        record.args = ()  # already interpolated into msg above
        return True


def _resolve_log_file(explicit: Optional[str]) -> Optional[Path]:
    raw = explicit if explicit is not None else os.environ.get(LOG_FILE_ENV)
    if raw is None:
        raw = DEFAULT_LOG_FILE
    if raw == "" or raw.strip().lower() == "none":
        return None
    return Path(raw).expanduser()


def _resolve_level(explicit: Optional[str]) -> int:
    raw = explicit or os.environ.get(LOG_LEVEL_ENV) or "INFO"
    return getattr(logging, raw.strip().upper(), logging.INFO)


def configure_logging(
    log_file: Optional[str] = None, level: Optional[str] = None
) -> Optional[Path]:
    """Attach a rotating, redacting file handler to the ``office_con`` logger.

    Idempotent: a second call replaces the existing office-connect handler
    rather than stacking another. Returns the resolved log path, or ``None``
    when file logging is disabled or could not be set up (never raises —
    logging must not take down the server).
    """
    pkg_logger = logging.getLogger("office_con")
    pkg_logger.setLevel(_resolve_level(level))

    # Drop any handler we previously installed so re-config is clean.
    for h in list(pkg_logger.handlers):
        if getattr(h, "name", None) == _HANDLER_NAME:
            pkg_logger.removeHandler(h)

    # Privacy fail-safe: never write a log file in a hosted/server deployment,
    # regardless of flags. Library records still propagate to the host app's
    # own logging. This is NOT overridable by --log-file on purpose.
    if is_server_environment():
        pkg_logger.propagate = True
        import sys
        print("office-connect: server environment detected — file logging "
              "disabled (privacy).", file=sys.stderr)
        return None

    path = _resolve_log_file(log_file)
    if path is None:
        # File logging disabled — restore default propagation so records still
        # reach any root handler (and don't suppress test log capture).
        pkg_logger.propagate = True
        return None

    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            os.chmod(path.parent, 0o700)
        except OSError:
            pass
        handler = RotatingFileHandler(
            path, maxBytes=_MAX_BYTES, backupCount=_BACKUP_COUNT, encoding="utf-8",
        )
        handler.name = _HANDLER_NAME
        handler.setFormatter(logging.Formatter(
            "%(asctime)s %(levelname)s %(name)s %(message)s"
        ))
        handler.addFilter(_RedactingFilter())
        pkg_logger.addHandler(handler)
        # The package logger owns its output; don't double-log via root.
        pkg_logger.propagate = False
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass
        pkg_logger.info("[LOG] file logging started -> %s (level=%s)",
                        path, logging.getLevelName(pkg_logger.level))
        return path
    except Exception as exc:  # pragma: no cover - disk/permission edge cases
        # Fall back silently to no file logging; surface once on stderr.
        import sys
        print(f"office-connect: could not set up log file {path}: {exc}",
              file=sys.stderr)
        return None

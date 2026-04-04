"""Mock user system for headless testing.

Enables automated testing (Playwright, integration tests) without real
Office 365 credentials.  The mock operates at the HTTP transport layer:
``run_async()`` intercepts MS Graph API URLs and returns synthetic JSON;
``_async_token_request()`` returns synthetic tokens.

Safety: requires ``LLMING_MOCK_USERS=1`` env flag **and** blocks
automatically on any Azure App Service environment.
"""

from __future__ import annotations

import logging
import os

logger = logging.getLogger(__name__)


def is_mock_enabled() -> bool:
    """Check if the mock user system is enabled AND safe to use."""
    if os.environ.get("LLMING_MOCK_USERS") != "1":
        return False

    # ── Production blockers ──────────────────────────────────
    if os.environ.get("WEBSITE_INSTANCE_ID"):
        logger.critical("[MOCK] BLOCKED — WEBSITE_INSTANCE_ID detected (Azure App Service)")
        return False

    website_url = os.environ.get("WEBSITE_URL", "")
    if "azurewebsites.net" in website_url or (website_url and "localhost" not in website_url and "127.0.0.1" not in website_url and "0.0.0.0" not in website_url):
        logger.critical("[MOCK] BLOCKED — production WEBSITE_URL detected: %s", website_url)
        return False

    return True

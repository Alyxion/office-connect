"""Synthetic JWT token generation for mock users."""

from __future__ import annotations

import time

import jwt


def make_mock_access_token(email: str, user_id: str, expires_in: int = 3600) -> str:
    """Create a JWT token with a valid ``exp`` claim.

    Not cryptographically signed against any real authority — contains
    ``iss: "mock-issuer"`` so it is unusable against real Azure AD.
    """
    payload = {
        "sub": user_id,
        "upn": email,
        "exp": int(time.time()) + expires_in,
        "iss": "mock-issuer",
    }
    return jwt.encode(payload, "mock-secret", algorithm="HS256")


def make_mock_token_response(email: str, user_id: str) -> dict:
    """Return a dict matching the Azure AD token endpoint response shape."""
    return {
        "access_token": make_mock_access_token(email, user_id),
        "refresh_token": f"mock-refresh-{email}",
        "token_type": "Bearer",
        "expires_in": 3600,
    }

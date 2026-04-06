"""Data models for mock users."""

from __future__ import annotations

from pydantic import BaseModel, Field


class MockUserProfile(BaseModel):
    """Generic mock user profile — no customer-specific references."""

    email: str
    user_id: str
    given_name: str
    surname: str
    full_name: str
    job_title: str = ""
    department: str = ""
    office_location: str = ""

    # Synthetic MS Graph data (dicts in Graph JSON shape)
    calendar_events: list[dict] = Field(default_factory=list)
    mail_messages: list[dict] = Field(default_factory=list)
    directory_users: list[dict] = Field(default_factory=list)

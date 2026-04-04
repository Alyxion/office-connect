"""Data models for mock users."""

from __future__ import annotations

import os
from pathlib import Path

from pydantic import BaseModel, Field

# Search paths for face photo assets (first match wins)
_FACES_SEARCH_PATHS = [
    Path(__file__).resolve().parent.parent.parent / "assets" / "faces",       # office-connect/assets/faces
    Path(os.environ.get("FACES_DIR", "")) if os.environ.get("FACES_DIR") else None,
]

# Pre-sorted face files by gender
_male_faces: list[Path] | None = None
_female_faces: list[Path] | None = None
_faces_dir: Path | None = None


def set_faces_dir(path: Path | str) -> None:
    """Override the face photo directory at runtime."""
    global _faces_dir, _male_faces, _female_faces
    _faces_dir = Path(path)
    _male_faces = None  # force reload
    _female_faces = None


def _load_face_lists() -> None:
    global _male_faces, _female_faces, _faces_dir
    if _male_faces is not None:
        return
    # Find the first existing faces directory
    if _faces_dir is None:
        for p in _FACES_SEARCH_PATHS:
            if p and p.is_dir() and any(p.glob("*.jpg")):
                _faces_dir = p
                break
    if _faces_dir and _faces_dir.is_dir():
        _male_faces = sorted(_faces_dir.glob("male_*.jpg"))
        _female_faces = sorted(_faces_dir.glob("female_*.jpg"))
    else:
        _male_faces = []
        _female_faces = []


def load_face_photo(gender: str, index: int) -> bytes | None:
    """Load a face JPEG by gender and index. Returns None if not available."""
    _load_face_lists()
    faces = _male_faces if gender == "male" else _female_faces
    if not faces:
        return None
    path = faces[index % len(faces)]
    return path.read_bytes()


def generate_avatar_svg(initials: str, color: str, size: int = 128) -> bytes:
    """Generate a simple SVG avatar with colored circle background and white initials.

    Returns the SVG as raw bytes (UTF-8 encoded).
    """
    font_size = size * 0.4
    svg = (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 {size} {size}">'
        f'<circle cx="{size // 2}" cy="{size // 2}" r="{size // 2}" fill="{color}"/>'
        f'<text x="50%" y="50%" text-anchor="middle" dy=".35em"'
        f' font-family="Arial, sans-serif" font-size="{font_size}" fill="white"'
        f' font-weight="bold">{initials}</text>'
        f'</svg>'
    )
    return svg.encode("utf-8")


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
    mail_folders: list[dict] = Field(default_factory=list)
    teams: list[dict] = Field(default_factory=list)
    chats: list[dict] = Field(default_factory=list)
    drives: list[dict] = Field(default_factory=list)
    user_photos: dict[str, bytes] = Field(default_factory=dict)

"""Privacy filtering for Microsoft 365 metadata and content."""
from __future__ import annotations

import fnmatch
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field
from pydantic.alias_generators import to_camel


PrivacyAccessStatus = Literal["ok", "content_blocked", "folder_blocked"]


class OfficeContentBlockedError(PermissionError):
    """Raised when metadata may be shown but content access is denied."""

    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


class OfficeFolderBlock(BaseModel):
    """One folder that must not be browsed or read below."""
    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True)

    item_id: str
    name: str = ""
    drive_id: str | None = None
    parent_path: str = ""
    web_url: str = ""
    source: str = ""


class OfficeFilterRules(BaseModel):
    """Hide/block term rules for a single surface (mail or files).

    Mail and files are configured independently because their behaviour
    differs: a term that makes sense for filenames (``.env``) is noise for
    mail subjects, and vice versa.
    """
    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True)

    enabled: bool = True
    hidden_name_terms: list[str] = Field(default_factory=list)
    hidden_content_terms: list[str] = Field(default_factory=list)
    blocked_content_name_terms: list[str] = Field(default_factory=list)
    blocked_content_terms: list[str] = Field(default_factory=list)
    # Whitelist: when any allowed term matches a name or content value, the
    # term-based hide/block rules above are skipped for that item. Explicit
    # folder/item blocks are NOT overridden.
    allowed_terms: list[str] = Field(default_factory=list)


class OfficeFileFilterRules(OfficeFilterRules):
    """File-surface rules, extended with explicit folder and file blocks.

    ``blocked_folders`` blocks a folder and everything beneath it;
    ``blocked_items`` blocks individual files (content access only — the item
    stays visible so the user can unblock it again).
    """

    blocked_folders: list[OfficeFolderBlock] = Field(default_factory=list)
    blocked_items: list[OfficeFolderBlock] = Field(default_factory=list)


class OfficePrivacyConfig(BaseModel):
    """User-configurable visibility and content-access rules.

    Mail and online-file rules are kept in separate groups so each surface can
    be enabled and tuned on its own.
    """
    model_config = ConfigDict(alias_generator=to_camel, populate_by_name=True)

    mail: OfficeFilterRules = Field(default_factory=OfficeFilterRules)
    files: OfficeFileFilterRules = Field(default_factory=OfficeFileFilterRules)


def normalized_terms(terms: list[str]) -> list[str]:
    """Clean user-entered term lists while preserving order."""
    result: list[str] = []
    seen: set[str] = set()
    for raw in terms:
        term = raw.strip()
        if not term:
            continue
        key = term.casefold()
        if key in seen:
            continue
        seen.add(key)
        result.append(term)
    return result


def _normalize_rules(rules: OfficeFilterRules) -> dict:
    return {
        "hidden_name_terms": normalized_terms(rules.hidden_name_terms),
        "hidden_content_terms": normalized_terms(rules.hidden_content_terms),
        "blocked_content_name_terms": normalized_terms(rules.blocked_content_name_terms),
        "blocked_content_terms": normalized_terms(rules.blocked_content_terms),
        "allowed_terms": normalized_terms(rules.allowed_terms),
    }


def normalize_privacy_config(config: OfficePrivacyConfig) -> OfficePrivacyConfig:
    """Return a config with trimmed, de-duplicated term lists per surface."""
    return config.model_copy(update={
        "mail": config.mail.model_copy(update=_normalize_rules(config.mail)),
        "files": config.files.model_copy(update=_normalize_rules(config.files)),
    })


def term_matches(value: str | None, terms: list[str]) -> bool:
    """Case-insensitive contains/glob matcher used by privacy rules."""
    if not value:
        return False
    folded = value.casefold()
    for raw in terms:
        term = raw.strip()
        if not term:
            continue
        folded_term = term.casefold()
        if "*" in folded_term or "?" in folded_term:
            if fnmatch.fnmatch(folded, folded_term):
                return True
        elif folded_term in folded:
            return True
    return False


def any_text_matches(values: list[str | None], terms: list[str]) -> bool:
    """True when any supplied text matches any term."""
    for value in values:
        if term_matches(value, terms):
            return True
    return False


def decode_text_for_rules(content: bytes) -> str:
    """Best-effort text decode for content privacy checks."""
    for encoding in ("utf-8", "utf-16", "latin-1"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    return ""


def folder_path_for_match(parent_path: str | None, name: str | None) -> str:
    """Build a best-effort Graph path for matching descendants."""
    base = (parent_path or "").rstrip("/")
    folder_name = (name or "").strip("/")
    if not base:
        return folder_name
    if not folder_name:
        return base
    return f"{base}/{folder_name}"

"""
Email filtering classes for Office Mail using Pydantic v2.
"""
import json
import fnmatch
import re
from typing import List, Optional, Dict, Any, Literal, ClassVar, Tuple
from pathlib import Path
from pydantic import BaseModel, Field, ConfigDict

from .mail_handler import OfficeMail

PRIORITY_PATTERN = re.compile(r"^(?:P(\d+):)?(.+)$")

def parse_priority_string(s: str) -> Tuple[int, str]:
    """
    Parse a string with optional priority prefix, e.g. "P2:foo@bar.com" -> (2, "foo@bar.com")
    If no priority is given, returns (100, s) (default low priority).
    """
    m = PRIORITY_PATTERN.match(s)
    if m:
        prio = int(m.group(1)) if m.group(1) else 0
        value = m.group(2)
        return prio, value
    return 0, s

class OfficeMailFilterReason(BaseModel):
    """Reason for a filter hit."""
    SENDER: ClassVar[str] = "sender"
    SUBJECT: ClassVar[str] = "subject"
    BODY: ClassVar[str] = "body"

    filter_name: str = Field(..., description="Name of the filter")
    reason_type: Literal["sender", "subject", "body"] = Field(..., description="Type of reason")
    value: str = Field(..., description="Value of the reason")
    inclusive: bool = Field(False, description="Whether the filter is inclusive")
    priority: int = Field(100, description="Priority of the filter")

class OfficeMailFilterResults(BaseModel):
    """Results of applying email filters to an OfficeMail."""
    model_config = ConfigDict(arbitrary_types_allowed=True)
    
    email: OfficeMail = Field(..., description="The email that was filtered")
    any_filter_hit: bool = Field(False, description="Whether any filter was hit")
    matched_filters: List[str] = Field(default_factory=list, description="List of matched filters")
    matched_reasons: Dict[str, List[OfficeMailFilterReason]] = Field(default_factory=dict, description="Dictionary of matched reasons")
    excluded: bool = Field(False, description="Whether the email was excluded")

    def get_reason_text(self, lang: str = "en") -> List[str]:
        """
        Return short human-readable explanations for all reasons, grouped by filter name.
        Example:
            Sender Filter: Sender blocked: 'no-reply@example.com'
            Body Filter: Body blocked: '...'
        """
        explanations = []
        for filter_name, reasons in self.matched_reasons.items():
            reason_lines = []
            for reason in reasons:
                incl = " (inclusive)" if reason.inclusive else ""
                if reason.reason_type == OfficeMailFilterReason.SENDER:
                    if lang == "de":
                        reason_lines.append(f"Absender blockiert{incl}: '{reason.value}'")
                    else:
                        reason_lines.append(f"Sender blocked{incl}: '{reason.value}'")
                elif reason.reason_type == OfficeMailFilterReason.SUBJECT:
                    if lang == "de":
                        reason_lines.append(f"Betreff blockiert{incl}: '{reason.value}'")
                    else:
                        reason_lines.append(f"Subject blocked{incl}: '{reason.value}'")
                elif reason.reason_type == OfficeMailFilterReason.BODY:
                    if lang == "de":
                        reason_lines.append(f"Inhalt blockiert{incl}: '{reason.value}'")
                    else:
                        reason_lines.append(f"Body blocked{incl}: '{reason.value}'")
            if reason_lines:
                explanations.append(f"{filter_name}: " + "; ".join(reason_lines))
        return explanations

class OfficeMailFilter(BaseModel):
    """Defines a filter for emails.
    
    Each filter can have multiple rules for senders, subjects, and body content.
    
    Each element in the list can have a priority prefix (e.g. "P1:sender@domain.com").

    The priority is used to determine which filter to apply in case of conflicts.
    
    The priority is a number from 0 to 100, where 0 is the highest priority.
    
    Example:
    
    ```
    senders_excluded: ["P1:sender@domain.com"]
    body_content_included: ["P4:Super urgent"]
    ```
    
    In this example the sender sender@domain.com is excluded but if the body contains "Super urgent" it will be included.
    
    The priority is used to determine which filter to apply in case of conflicts.
    """
    model_config = ConfigDict(arbitrary_types_allowed=True)
    
    name: str = Field(..., description="Name of the filter")
    senders_excluded: List[str] = Field(default_factory=list, description="List of sender addresses to exclude")
    subjects_excluded: List[str] = Field(default_factory=list, description="List of subjects to exclude")
    body_content_excluded: List[str] = Field(default_factory=list, description="List of body content to exclude")
    senders_included: Optional[List[str]] = Field(None, description="List of sender addresses to include")
    subjects_included: Optional[List[str]] = Field(None, description="List of subjects to include")
    body_content_included: Optional[List[str]] = Field(None, description="List of body content to include")
    
    @classmethod
    def from_json_file(cls, file_path: str | Path) -> "OfficeMailFilter":
        """Load filter from JSON file."""
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return cls(**data)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "OfficeMailFilter":
        """Load filter from dictionary."""
        return cls(**data)
    
    def _match_mask(self, value: Optional[str], mask: str) -> bool:
        """Case-insensitive fnmatch with * wildcards."""
        if value is None:
            return False
        return fnmatch.fnmatchcase(value.lower(), mask.lower())

    def _contains_substring(self, value: Optional[str], substring: str) -> bool:
        """Case-insensitive substring search."""
        if value is None:
            return False
        return substring.lower() in value.lower()

    def apply(self, mail: OfficeMail) -> OfficeMailFilterResults:
        """
        Test this filter against an OfficeMail and return results.
        Inclusive filters with higher priority win, with lower lose.
        Returns machine-readable reasons.
        """
        results = OfficeMailFilterResults(email=mail, any_filter_hit=False, excluded=False)
        matched_reasons: List[OfficeMailFilterReason] = []

        # SENDER
        sender_matches = []
        if mail.from_email:
            # Inclusive
            if self.senders_included:
                for s in self.senders_included:
                    prio, mask = parse_priority_string(s)
                    if self._match_mask(mail.from_email, mask):
                        sender_matches.append((prio, True, mask))
            # Exclusive
            if self.senders_excluded:
                for s in self.senders_excluded:
                    prio, mask = parse_priority_string(s)
                    if self._match_mask(mail.from_email, mask):
                        sender_matches.append((prio, False, mask))
            if sender_matches:
                sender_matches.sort(key=lambda x: (x[0], not x[1]))  # Inclusive wins on tie
                best = sender_matches[0]
                matched_reasons.append(OfficeMailFilterReason(
                    filter_name=self.name,
                    reason_type="sender",
                    value=mail.from_email,
                    inclusive=best[1],
                    priority=best[0]
                ))
        # SUBJECT
        subject_matches = []
        if mail.subject:
            # Inclusive
            if self.subjects_included:
                for s in self.subjects_included:
                    prio, substr = parse_priority_string(s)
                    if self._contains_substring(mail.subject, substr):
                        subject_matches.append((prio, True, substr))
            # Exclusive
            if self.subjects_excluded:
                for s in self.subjects_excluded:
                    prio, substr = parse_priority_string(s)
                    if self._contains_substring(mail.subject, substr):
                        subject_matches.append((prio, False, substr))
            if subject_matches:
                subject_matches.sort(key=lambda x: (x[0], not x[1]))
                best = subject_matches[0]
                matched_reasons.append(OfficeMailFilterReason(
                    filter_name=self.name,
                    reason_type="subject",
                    value=mail.subject,
                    inclusive=best[1],
                    priority=best[0]
                ))
        # BODY (body or body_preview)
        body_matches = []
        body_text = mail.body if mail.body else mail.body_preview
        if body_text:
            # Inclusive
            if self.body_content_included:
                for s in self.body_content_included:
                    prio, substr = parse_priority_string(s)
                    if self._contains_substring(body_text, substr):
                        body_matches.append((prio, True, substr))
            # Exclusive
            if self.body_content_excluded:
                for s in self.body_content_excluded:
                    prio, substr = parse_priority_string(s)
                    if self._contains_substring(body_text, substr):
                        body_matches.append((prio, False, substr))
            if body_matches:
                body_matches.sort(key=lambda x: (x[0], not x[1]))
                best = body_matches[0]
                matched_reasons.append(OfficeMailFilterReason(
                    filter_name=self.name,
                    reason_type="body",
                    value=body_text,
                    inclusive=best[1],
                    priority=best[0]
                ))
        if matched_reasons:
            results.any_filter_hit = True
            results.matched_filters.append(self.name)
            results.matched_reasons[self.name] = matched_reasons
            # Set excluded if any reason is exclusive and has highest or equal priority
            for reason in matched_reasons:
                if not reason.inclusive:
                    results.excluded = True
                    break
        
        return results

class OfficeMailFilterList(BaseModel):
    """A list of email filters that can be applied to emails."""
    model_config = ConfigDict(arbitrary_types_allowed=True)
    
    filters: List[OfficeMailFilter] = Field(default_factory=list, description="Ordered list of email filters to apply")

    def combine(self, other):
        """
        Combine this filter list with another OfficeMailFilterList or OfficeMailFilter.
        Returns a new OfficeMailFilterList.
        """
        if isinstance(other, OfficeMailFilterList):
            return OfficeMailFilterList(filters=self.filters + other.filters)
        elif isinstance(other, OfficeMailFilter):
            return OfficeMailFilterList(filters=self.filters + [other])
        else:
            raise TypeError(f"Cannot combine OfficeMailFilterList with {type(other)}")

    def __add__(self, other):
        """
        Allow using the + operator to merge filter lists or add a filter.
        """
        return self.combine(other)

    def __radd__(self, other):
        """
        Support sum() and right-hand add with OfficeMailFilter.
        """
        if isinstance(other, OfficeMailFilter):
            return OfficeMailFilterList(filters=[other] + self.filters)
        elif isinstance(other, OfficeMailFilterList):
            return OfficeMailFilterList(filters=other.filters + self.filters)
        else:
            raise TypeError(f"Cannot add {type(other)} to OfficeMailFilterList")
    
    @classmethod
    def from_json_files(cls, file_paths: List[str | Path]) -> "OfficeMailFilterList":
        """Load filters from multiple JSON files."""
        filters = []
        for file_path in file_paths:
            filter_obj = OfficeMailFilter.from_json_file(file_path)
            filters.append(filter_obj)
        return cls(filters=filters)
    
    @classmethod
    def from_dicts(cls, filter_dicts: List[Dict[str, Any]]) -> "OfficeMailFilterList":
        """Load filters from list of dictionaries."""
        filters = []
        for filter_dict in filter_dicts:
            filter_obj = OfficeMailFilter.from_dict(filter_dict)
            filters.append(filter_obj)
        return cls(filters=filters)
    
    def add_filter(self, filter_obj: OfficeMailFilter) -> None:
        """Add a filter to the list."""
        self.filters.append(filter_obj)
    
    def apply(self, mail: OfficeMail) -> OfficeMailFilterResults:
        """
        Test all filters against an OfficeMail and return combined results.
        Returns machine-readable reasons.
        """
        combined_results = OfficeMailFilterResults(email=mail, any_filter_hit=False, excluded=False)
        excluded = False
        
        for filter_obj in self.filters:
            filter_results = filter_obj.apply(mail)
            if filter_results.any_filter_hit:
                combined_results.any_filter_hit = True
                combined_results.matched_filters.extend(filter_results.matched_filters)
                combined_results.matched_reasons.update(filter_results.matched_reasons)
                if filter_results.excluded:
                    excluded = True
        combined_results.excluded = excluded
        return combined_results

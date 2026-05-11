"""Shared types for rule findings."""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable


@dataclass
class Finding:
    """A single rule hit at a specific location."""

    category: str  # "te_point" | "variant" | "spelling" | "bias" | "compound" | "article" | "casing"
    rule_id: str  # stable identifier, e.g. "percent_style"
    rule_label: str  # human-readable
    surface: str  # the actual text matched (e.g. "decisionmaking")
    canonical: str  # normalized form used to group variants (e.g. "decision making")
    chapter_index: int
    chapter_name: str
    source: str
    page: int
    para_index: int
    context: str
    severity: str = "info"  # "info" | "warn" | "error"
    replacement: str | None = None
    search_pattern: str | None = None
    region: str = "body"
    match_start: int = 0
    match_end: int = 0
    replacement_options: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "category": self.category,
            "rule_id": self.rule_id,
            "rule_label": self.rule_label,
            "surface": self.surface,
            "canonical": self.canonical,
            "chapter_index": self.chapter_index,
            "chapter_name": self.chapter_name,
            "source": self.source,
            "page": self.page,
            "para_index": self.para_index,
            "context": self.context,
            "severity": self.severity,
            "replacement": self.replacement,
            "search_pattern": self.search_pattern,
            "region": self.region,
            "match_start": self.match_start,
            "match_end": self.match_end,
            "replacement_options": self.replacement_options,
        }


def context_snippet(text: str, start: int, end: int, window: int = 60) -> str:
    """Return `... text before [match] text after ...` with fixed window."""
    s = max(0, start - window)
    e = min(len(text), end + window)
    prefix = "… " if s > 0 else ""
    suffix = " …" if e < len(text) else ""
    return prefix + text[s:start] + "⟪" + text[start:end] + "⟫" + text[end:e] + suffix


def iter_unmasked_matches(
    pattern: re.Pattern, text: str, mask: list[bool]
) -> Iterable[re.Match]:
    """Yield regex matches that do NOT overlap an excluded (masked) region.

    A match is skipped if any character of the match sits in the mask.
    """
    for m in pattern.finditer(text):
        span = range(m.start(), m.end())
        # If mask shorter than text for any reason, treat missing as False.
        if any(pos < len(mask) and mask[pos] for pos in span):
            continue
        yield m


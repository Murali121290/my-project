"""US vs UK spelling detection.

Strategy: scan the text for every US-form OR UK-form from the pair list. Each
hit becomes a Finding keyed by the canonical (US) form. The aggregator then
tells the editor the dominant convention per chapter/manuscript and flags
outliers.
"""
from __future__ import annotations

import re
from typing import Iterable

from manuscript_core.data.uk_us_pairs import UK_US_PAIRS
from manuscript_core.extractor import Segment
from manuscript_core.rules.base import Finding, context_snippet, iter_unmasked_matches


def _build_patterns() -> list[tuple[re.Pattern, str, str]]:
    """Return (pattern, canonical_us_form, variant_type) for every word.

    We compile one regex per side of each pair with a word boundary.
    """
    out: list[tuple[re.Pattern, str, str]] = []
    for uk, us in UK_US_PAIRS:
        out.append((re.compile(r"\b" + re.escape(uk) + r"\b", re.IGNORECASE), us.lower(), "UK"))
        out.append((re.compile(r"\b" + re.escape(us) + r"\b", re.IGNORECASE), us.lower(), "US"))
    return out


_PATTERNS = _build_patterns()


def run_spelling_rules(seg: Segment) -> Iterable[Finding]:
    for pat, canonical, variant in _PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="spelling",
                rule_id=f"spelling_{canonical}",
                rule_label=f"US/UK: {canonical}",
                surface=m.group(0),
                canonical=canonical,
                chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name,
                source=seg.source,
                page=seg.page,
                para_index=seg.para_index,
                context=context_snippet(seg.text, m.start(), m.end()),
                severity="info" if variant == "US" else "warn",
                replacement=canonical,
                search_pattern=pat.pattern,
                region=seg.region,
            )


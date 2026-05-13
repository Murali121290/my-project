"""Compound variant detection â€” hyphenated / closed-up / spaced.
We maintain a curated list of compound "bases" (like "decision making",
"health care", "e-mail"). For each, we scan for all three surface forms:
hyphenated, spaced, and closed-up. Every occurrence becomes a Finding;
the aggregator later groups them by `canonical` so the dashboard can
report "decision-making (34) Â· decision making (12) Â· decisionmaking (1)".
Casing variants (Likert vs likert) are surfaced by the aggregator when the
same canonical form appears in multiple casings.
"""
from __future__ import annotations
import re
from typing import Iterable
from manuscript_core.extractor import Segment
from manuscript_core.rules.base import Finding, context_snippet, iter_unmasked_matches
# Curated list from the brief. Each entry is the "spaced" canonical form;
# we auto-generate hyphenated and closed-up patterns from it.
COMPOUND_BASES: tuple[str, ...] = (
    "co identity",
    "co dosed",
    "semi independent",
    "hull less",
    "ultra atomic",
    "de emphasize",
    "intra abdominal",
    "bell like",
    "anti inflammatory",
    "under resourced",
    "internet",
    "amino acid",
    "birth control",
    "bone marrow",
    "primary care",
    "public health",
    "soft tissue",
    "tertiary care",
    "foreign body",
    "small cell",
    "natural killer",
    "open access",
    "deep vein",
    "fresh frozen",
    "peer review",
    "Student t test",
    "Cronbach alpha",
    # From the brief â€” core editorial triplets
    "decision making",
    "health care",
    "three dimensional",
    "two dimensional",
    "two thirds",
    "three fourths",
    "beta blocker",
    "beta blockers",
    "half life",
    "half lives",
    "false positive test",
    "tonic clonic",
    "double blind",
    "blue gray",
    "blue black",
    "bluish gray",
    "vice chancellor",
    "vice consul",
    "micro organism",
    "co operation",
    "re enter",
    "anti microbial",
    "auto immune",
    "co existence",
    "co exist",
    "counter measure",
    "co worker",
    "meta analysis",
    "meta analyses",
    "multi institutional",
    "mild to moderate",
    "waist to hip ratio",
    "cup to disc ratio",
    "African American",
    "Mexican American",
    "Latin American",
    "a priori",
    "prima facie",
    "ex officio",
    "in vivo",
    "post hoc",
    "B cell",
    "T cell",
    "brow lift",
    "face lift",
    "white coat hypertension",
    "C reactive protein",
    "T wave",
    "Mann Whitney test",
    "T shirt",
    "prostate specific antigen",
    "amino acid levels",
    "birth control methods",
    "bone marrow biopsy",
    "health care system",
    "inner ear disorder",
    "lower extremity amputation",
    "medical school students",
    "multiple organ disease",
    "natural killer cell",
    "open access journal",
    "open heart surgery",
    "parallel furrow pattern",
    "deep vein thrombosis",
    "foreign body aspiration",
    "fresh frozen plasma",
    "patch test series",
    "peer review process",
    "primary care physician",
    "public health official",
    "small cell carcinoma",
    "soft tissue sarcoma",
    "tertiary care center",
    "school age",
    "data set",
    "end point",
    "flow chart",
    "flow diagram",
    "gall bladder",
    "heart beat",
    "needle stick",
    "radio frequency",
    "radio guided",
    "skin fold",
    "slit lamp",
    "wave form",
    "web site",
    "brain stem",
    "blood stream",
    "care giver",
    "care taker",
    "case load",
    "fiber optic",
    "e mail",
    "e book",
    "e cigarette",
    "web page",
    "web cast",
    "web cam",
)
def _build_compound_patterns() -> list[tuple[re.Pattern, str, str, str]]:
    """Return list of (pattern, canonical_key, rule_id, form_label).
    form_label is "spaced" | "hyphenated" | "closed".
    canonical_key is the normalized form (lowercased, single-spaced) used to
    group all variants together.
    """
    out: list[tuple[re.Pattern, str, str, str]] = []
    for base in COMPOUND_BASES:
        canonical = base.lower()
        words = base.split()
        if len(words) < 2:
            continue
        # Build patterns â€” case-insensitive word-boundary-anchored
        spaced = r"\b" + r"\s+".join(re.escape(w) for w in words) + r"\b"
        hyphenated = r"\b" + r"-".join(re.escape(w) for w in words) + r"\b"
        closed = r"\b" + "".join(re.escape(w) for w in words) + r"\b"
        rule_id = canonical.replace(" ", "_")
        out.append((re.compile(spaced, re.IGNORECASE), canonical, rule_id, "spaced"))
        out.append((re.compile(hyphenated, re.IGNORECASE), canonical, rule_id, "hyphenated"))
        # Only scan closed-up form if it wouldn't trigger on common words â€”
        # skip 2-letter-first-word cases like "e mail"â†’"email" via a safe check.
        # "email" is fine; to avoid false positives for short joins we still scan
        # but the aggregator will only flag if multiple forms co-occur.
        out.append((re.compile(closed, re.IGNORECASE), canonical, rule_id, "closed"))
    return out
_COMPOUND_PATTERNS = _build_compound_patterns()
_CUSTOM_COMPOUND_PATTERNS = [
    # (pattern, canonical, rule_id, form_label)
    (re.compile(r"\bthree dimensional\b", re.IGNORECASE), "three dimensional", "three_dimensional", "spaced"),
    (re.compile(r"\bthree-dimensional\b", re.IGNORECASE), "three dimensional", "three_dimensional", "hyphenated"),
    (re.compile(r"\b3D\b", re.IGNORECASE), "three dimensional", "three_dimensional", "3D"),
    (re.compile(r"\b3-D\b", re.IGNORECASE), "three dimensional", "three_dimensional", "3-D"),
    (re.compile(r"\btwo dimensional\b", re.IGNORECASE), "two dimensional", "two_dimensional", "spaced"),
    (re.compile(r"\btwo-dimensional\b", re.IGNORECASE), "two dimensional", "two_dimensional", "hyphenated"),
    (re.compile(r"\b2D\b", re.IGNORECASE), "two dimensional", "two_dimensional", "2D"),
    (re.compile(r"\b2-D\b", re.IGNORECASE), "two dimensional", "two_dimensional", "2-D"),
    (re.compile(r"\bβ blocker\b", re.IGNORECASE), "beta blocker", "beta_blocker_greek", "spaced"),
    (re.compile(r"\bβ-blocker\b", re.IGNORECASE), "beta blocker", "beta_blocker_greek", "hyphenated"),
    (re.compile(r"\bβblocker\b", re.IGNORECASE), "beta blocker", "beta_blocker_greek", "closed"),
]
_LY_HYPHEN_PATTERN = re.compile(r"\b\w+ly-\w+\b")
def run_compound_rules(seg: Segment) -> Iterable[Finding]:
    # Run the standard compound bases
    for pat, canonical, rule_id, form in _COMPOUND_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="compound",
                rule_id=rule_id,
                rule_label=f"Compound: {canonical}",
                surface=m.group(0),
                canonical=canonical,
                chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name,
                source=seg.source,
                page=seg.page,
                para_index=seg.para_index,
                context=context_snippet(seg.text, m.start(), m.end()),
            )
    # Run custom patterns (3D, etc.)
    for pat, canonical, rule_id, form in _CUSTOM_COMPOUND_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="compound",
                rule_id=rule_id,
                rule_label=f"Compound: {canonical}",
                surface=m.group(0),
                canonical=canonical,
                chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name,
                source=seg.source,
                page=seg.page,
                para_index=seg.para_index,
                context=context_snippet(seg.text, m.start(), m.end()),
            )
    # Run -ly hyphen pattern
    for m in iter_unmasked_matches(_LY_HYPHEN_PATTERN, seg.text, seg.mask):
        yield Finding(
            category="te_point",
            rule_id="ly_hyphen_style",
            rule_label="Hyphenated -ly compound",
            surface=m.group(0),
            canonical="ly_hyphen_style",
            chapter_index=seg.chapter_index,
            chapter_name=seg.chapter_name,
            source=seg.source,
            page=seg.page,
            para_index=seg.para_index,
            context=context_snippet(seg.text, m.start(), m.end()),
        )

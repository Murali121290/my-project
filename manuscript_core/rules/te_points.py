"""TE (technical editing) style-point detectors.
Each detector returns Findings for every hit. We don't try to pick the
"correct" form — we surface all forms so the editor can see which dominates
and enforce consistency.
"""
from __future__ import annotations
from manuscript_core.rules.base import Finding, context_snippet, iter_unmasked_matches
import re
from typing import Iterable
from manuscript_core.extractor import Segment
# Shortcut: build a Finding from a segment + match with common fields.
def _f(seg: Segment, m: re.Match, category: str, rule_id: str,
       rule_label: str, canonical: str, severity: str = "info",
       pat: re.Pattern | None = None) -> Finding:
    from manuscript_core.rules.te_point_replacements import get_te_replacement, get_te_replacement_options
    rep = get_te_replacement(rule_id, m)
    rep_opts = get_te_replacement_options(rule_id, m)
    # For pattern matching in fixer: use surface-specific pattern instead of shared regex
    # This ensures "one" matches only "one", not any number when multiple number rules apply
    surface = m.group(0)
    surface_pattern = r'\b' + re.escape(surface) + r'\b'
    return Finding(
        category=category, rule_id=rule_id, rule_label=rule_label,
        surface=m.group(0), canonical=canonical,
        chapter_index=seg.chapter_index, chapter_name=seg.chapter_name,
        source=seg.source, page=seg.page, para_index=seg.para_index,
        context=context_snippet(seg.text, m.start(), m.end()),
        severity=severity,
        replacement=rep,
        search_pattern=surface_pattern,
        region=seg.region,
        match_start=m.start(),
        match_end=m.end(),
        replacement_options=rep_opts,
    )
# ---------------------------------------------------------------------------
# Rule 1: percent style — %, percent, per cent
# ---------------------------------------------------------------------------
PERCENT_PATTERNS = [
    (re.compile(r"\b\d+(?:\.\d+)?\s*%"), "percent_symbol"),
    (re.compile(r"\bpercent\b", re.IGNORECASE), "percent_word"),
    (re.compile(r"\bper\s+cent\b", re.IGNORECASE), "per_cent_word"),
    (re.compile(r"\bpercentage\b", re.IGNORECASE), "percent_percentage"),
]
def detect_percent(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in PERCENT_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "% vs percent vs per cent", "percent_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 2: ellipsis form
# ---------------------------------------------------------------------------
ELLIPSIS_PATTERNS = [
    (re.compile(r"\u2026"), "ellipsis_symbol"),
    (re.compile(r"\.\.\."), "ellipsis_3dots"),
    (re.compile(r"\.\s\.\s\."), "ellipsis_spaced"),
]
def detect_ellipsis(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in ELLIPSIS_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Ellipsis style", "ellipsis_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 3: AM/PM style — split by AM vs PM and dot variant
# ---------------------------------------------------------------------------
AMPM_PATTERNS = [
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*AM\b"), "ampm_am_upper"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*PM\b"), "ampm_pm_upper"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*A\.M\b"), "ampm_am_upper_nodot"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*A\.M\."), "ampm_am_upper_dots"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*P\.M\b"), "ampm_pm_upper_nodot"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*P\.M\."), "ampm_pm_upper_dots"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*a\.m\b"), "ampm_am_lower_nodot"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*a\.m\."), "ampm_am_lower_dots"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*p\.m\b"), "ampm_pm_lower_nodot"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*p\.m\."), "ampm_pm_lower_dots"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*am\b"), "ampm_am_lower"),
    (re.compile(r"\b\d{1,2}(?::\d{2})?\s*pm\b"), "ampm_pm_lower"),
]
def detect_ampm(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in AMPM_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "AM/PM style", "ampm_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 4: era markers — AD/BC with and without dots, split by marker
# ---------------------------------------------------------------------------
ERA_PATTERNS = [
    (re.compile(r"\bA\.D\."), "era_ad_dots"),
    (re.compile(r"\bB\.C\."), "era_bc_dots"),
    (re.compile(r"\bA\.D\b"), "era_ad_nodot"),
    (re.compile(r"\bB\.C\b"), "era_bc_nodot"),
    (re.compile(r"\bAD\b"), "era_ad"),
    (re.compile(r"\bBC\b"), "era_bc"),
]
def detect_era(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in ERA_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "AD/BC style", "era_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 5: number ranges — "1 to 6", "1–6", "1-6"
# ---------------------------------------------------------------------------
RANGE_PATTERNS = [
    (re.compile(r"\b\d+\s+to\s+\d+\b"), "range_to"),
    (re.compile(r"\b\d+\u2013\d+\b"), "range_endash"),
    (re.compile(r"\b\d+-\d+\b"), "range_hyphen"),
]
def detect_number_ranges(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in RANGE_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Number range style", "range_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 6: leading zero — .5 vs 0.5
# ---------------------------------------------------------------------------
LEADING_ZERO_PATTERNS = [
    (re.compile(r"(?<![\w.])\.\d+"), "leading_zero_missing"),
    (re.compile(r"\b0\.\d+"), "leading_zero_present"),
]
def detect_leading_zero(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in LEADING_ZERO_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Leading zero on decimals", "leading_zero_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 7: multiplication sign — × vs x
# ---------------------------------------------------------------------------
TIMES_SIGN_PATTERNS = [
    (re.compile(r"\b\d+\s*\u00d7\s*\d+\b"), "times_symbol"),
    (re.compile(r"\b\d+\s*[xX]\s*\d+\b"), "times_letter"),
]
def detect_times(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in TIMES_SIGN_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Multiplication sign (× vs x)", "times_sign_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 8: versus abbreviation
# ---------------------------------------------------------------------------
VERSUS_PATTERNS = [
    (re.compile(r"\bversus\b", re.IGNORECASE), "versus_full"),
    (re.compile(r"\bvs\."), "versus_vs_dot"),
    (re.compile(r"\bvs\b(?!\.)"), "versus_vs"),
    (re.compile(r"\bv\.(?=\s)"), "versus_v_dot"),
]
def detect_versus(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in VERSUS_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "versus / vs. / vs", "versus_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 9: Latin abbreviations
# ---------------------------------------------------------------------------
LATIN_PATTERNS = [
    (re.compile(r"\be\.g\.,?"), "latin_eg_dots"),
    (re.compile(r"\beg\b(?!\.)", re.IGNORECASE), "latin_eg_nodots"),
    (re.compile(r"\bi\.e\.,?"), "latin_ie_dots"),
    (re.compile(r"\bie\b(?!\.)", re.IGNORECASE), "latin_ie_nodots"),
    (re.compile(r"\betc\."), "latin_etc_dot"),
    (re.compile(r"\betc\b(?!\.)"), "latin_etc_nodot"),
]
def detect_latin_abbr(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in LATIN_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Latin abbreviation style", "latin_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 10: century style
# ---------------------------------------------------------------------------
CENTURY_PATTERNS = [
    (re.compile(r"\b\d{1,2}(?:st|nd|rd|th)\s+century\b", re.IGNORECASE), "century_num_sg"),
    (re.compile(r"\b\d{1,2}(?:st|nd|rd|th)\s+centuries\b", re.IGNORECASE), "century_num_pl"),
    (re.compile(r"\b(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth|seventeenth|eighteenth|nineteenth|twentieth|twenty-first|twenty-second|twenty-third|twenty-fourth|twenty-fifth|twenty-sixth|twenty-seventh|twenty-eighth|twenty-ninth|thirtieth)\s+century\b", re.IGNORECASE), "century_spelled_sg"),
    (re.compile(r"\b(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth|seventeenth|eighteenth|nineteenth|twentieth|twenty-first|twenty-second|twenty-third|twenty-fourth|twenty-fifth|twenty-sixth|twenty-seventh|twenty-eighth|twenty-ninth|thirtieth)\s+centuries\b", re.IGNORECASE), "century_spelled_pl"),
    (re.compile(r"\b\d{1,2}st\s+century\b", re.IGNORECASE), "century_st_sg"),
    (re.compile(r"\b\d{1,2}nd\s+century\b", re.IGNORECASE), "century_nd_sg"),
    (re.compile(r"\b\d{1,2}rd\s+century\b", re.IGNORECASE), "century_rd_sg"),
    (re.compile(r"\b\d{1,2}th\s+century\b", re.IGNORECASE), "century_th_sg"),
    (re.compile(r"\b\d{1,2}st\s+centuries\b", re.IGNORECASE), "century_st_pl"),
    (re.compile(r"\b\d{1,2}nd\s+centuries\b", re.IGNORECASE), "century_nd_pl"),
    (re.compile(r"\b\d{1,2}rd\s+centuries\b", re.IGNORECASE), "century_rd_pl"),
    (re.compile(r"\b\d{1,2}th\s+centuries\b", re.IGNORECASE), "century_th_pl"),
]
def detect_century(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in CENTURY_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Century style (numeric vs spelled)", "century_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 11: thousands separator
# ---------------------------------------------------------------------------
THOUSANDS_PATTERNS = [
    (re.compile(r"\b\d{1,3}(?:,\d{3})+\b"), "thousands_comma"),
    (re.compile(r"\b\d{1,3}(?:\u00a0\d{3})+\b"), "thousands_nbsp"),
    (re.compile(r"(?<!\d)\d{4,}(?!\d)"), "thousands_nosep"),
]
_YEAR_RE = re.compile(r"^(?:1\d{3}|20\d{2})$")
def _looks_like_year(text: str, start: int, end: int) -> bool:
    surface = text[start:end]
    if _YEAR_RE.match(surface):
        return True
    if end < len(text) and text[end] == "s" and _YEAR_RE.match(surface):
        return True
    if _YEAR_RE.match(surface):
        left = text[max(0, start - 8):start]
        right = text[end:end + 8]
        if re.search(r"\b(?:1\d{3}|20\d{2})\s*[\u2013\-]\s*$", left):
            return True
        if re.search(r"\b(?:1\d{3}|20\d{2})\s+to\s+$", left):
            return True
        if re.match(r"\s*[\u2013\-]\s*(?:1\d{3}|20\d{2})\b", right):
            return True
        if re.match(r"\s+to\s+(?:1\d{3}|20\d{2})\b", right):
            return True
    return False
def detect_thousands(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in THOUSANDS_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            if _looks_like_year(seg.text, m.start(), m.end()):
                continue
            yield _f(seg, m, "te_point", rule_id, "Thousands separator style", "thousands_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 12: Figure / Table / Box / Chapter references (numbered)
#
# Pattern hierarchy (most specific first to avoid double-counting):
#   Figures/Tables/Boxes (plural) with range connectors
#   Figure/Table/Box (singular) with dotted or dashed number
#   Figure/Table/Box (singular) plain
#   Fig/Figs, Tab/Tabs abbreviations (same hierarchy)
#   Lowercase variants
#   Chapter / Chapters
# ---------------------------------------------------------------------------
_N = r"\d+"          # integer
_NN = r"\d+\.\d+"    # dotted  e.g. 1.1
_SEP = r"(?:\.\d+)?" # optional dot part
# Connector alternation (order: longest/most specific first)
_CONN_DOTTED = (
    r"(?P<range>"
    r"(?:{nn})\s+through\s+(?:{nn})"
    r"|(?:{nn})\s+to\s+(?:{nn})"
    r"|(?:{nn})\s+and\s+(?:{nn})"
    r"|(?:{nn})\s*&\s*(?:{nn})"
    r"|(?:{nn})\u2013(?:{nn})"
    r"|(?:{nn})-(?:{nn})"
    r")"
).format(nn=_NN)
def _ref_patterns(singular: str, plural: str, abbr_s: str, abbr_p: str) -> list[tuple[re.Pattern, str, str]]:
    """Return (pattern, rule_id, label) triples for one reference type."""
    # Prefix for rule ids  e.g. "figure" -> "ref_figure_*"
    pfx = singular.lower()
    rows: list[tuple[re.Pattern, str, str]] = []
    # ---- Plural abbreviated (Figs/Tabs/Boxes) with connectors ----
    for rid, conn in [
        (f"ref_{abbr_p.lower()}_through",  r"\s+through\s+"),
        (f"ref_{abbr_p.lower()}_to",       r"\s+to\s+"),
        (f"ref_{abbr_p.lower()}_and",      r"\s+and\s+"),
        (f"ref_{abbr_p.lower()}_amp",      r"\s*&\s*"),
        (f"ref_{abbr_p.lower()}_endash",   r"\u2013"),
        (f"ref_{abbr_p.lower()}_dash",     r"-"),
    ]:
        pat = re.compile(
            rf"\b{re.escape(abbr_p)}\.?\s+{_NN}{conn}{_NN}\b",
            re.IGNORECASE,
        )
        rows.append((pat, rid, f"{abbr_p} range"))
    # ---- Plural full (Figures/Tables/Boxes) with connectors ----
    for rid, conn in [
        (f"ref_{plural.lower()}_through", r"\s+through\s+"),
        (f"ref_{plural.lower()}_to",      r"\s+to\s+"),
        (f"ref_{plural.lower()}_and",     r"\s+and\s+"),
        (f"ref_{plural.lower()}_amp",     r"\s*&\s*"),
        (f"ref_{plural.lower()}_endash",  r"\u2013"),
        (f"ref_{plural.lower()}_dash",    r"-"),
    ]:
        pat = re.compile(
            rf"\b{re.escape(plural)}\s+{_NN}{conn}{_NN}\b",
            re.IGNORECASE,
        )
        rows.append((pat, rid, f"{plural} range"))
    # ---- Abbreviated singular dotted / dashed / plain (Fig 1.1, Fig 1-1, Fig 1) ----
    rows.append((re.compile(rf"\b{re.escape(abbr_s)}\.?\s+{_NN}\b", re.IGNORECASE),
                 f"ref_{abbr_s.lower()}_dotted", f"{abbr_s} dotted"))
    rows.append((re.compile(rf"\b{re.escape(abbr_s)}\.?\s+{_N}-{_N}\b", re.IGNORECASE),
                 f"ref_{abbr_s.lower()}_dash", f"{abbr_s} dash"))
    rows.append((re.compile(rf"\b{re.escape(abbr_s)}\.?\s+{_N}\b(?![\.\-\d])", re.IGNORECASE),
                 f"ref_{abbr_s.lower()}_single", f"{abbr_s} single"))
    # ---- Abbreviated PLURAL dotted / single (Figs 1.1, Tabs 3.2, Tabs 3) ----
    rows.append((re.compile(rf"\b{re.escape(abbr_p)}\.?\s+{_NN}\b", re.IGNORECASE),
                 f"ref_{abbr_p.lower()}_dotted", f"{abbr_p} dotted"))
    rows.append((re.compile(rf"\b{re.escape(abbr_p)}\.?\s+{_N}\b(?![\.\-\d])", re.IGNORECASE),
                 f"ref_{abbr_p.lower()}_single_num", f"{abbr_p} single"))
    # ---- Full singular lowercase dotted ----
    rows.append((re.compile(rf"\b{singular.lower()}\s+{_NN}\b"),
                 f"ref_{pfx}_lc_dotted", f"{singular.lower()} dotted (lc)"))
    # ---- Abbreviated singular lowercase dotted ----
    rows.append((re.compile(rf"\b{abbr_s.lower()}\s+{_NN}\b"),
                 f"ref_{abbr_s.lower()}_lc_dotted", f"{abbr_s.lower()} dotted (lc)"))
    # ---- Full singular dotted / dashed / plain (Figure 1.1, Figure 1-1, Figure 1) ----
    rows.append((re.compile(rf"\b{re.escape(singular)}\s+{_NN}\b", re.IGNORECASE),
                 f"ref_{pfx}_dotted", f"{singular} dotted"))
    rows.append((re.compile(rf"\b{re.escape(singular)}\s+{_N}-{_N}\b", re.IGNORECASE),
                 f"ref_{pfx}_dash", f"{singular} dash"))
    rows.append((re.compile(rf"\b{re.escape(singular)}\s+{_N}\b(?![\.\-\d])", re.IGNORECASE),
                 f"ref_{pfx}_single", f"{singular} single"))
    # ---- Lowercase variants ----
    # The generic lowercase variant should only capture plain numbers if more specific dotted/dashed patterns are added.
    # rows.append((re.compile(rf"\b{singular.lower()}s?\s+{_N}(?:[\.\-]{_N})?\b"),
    #              f"ref_{pfx}_lc", f"{singular.lower()} lowercase"))
    return rows
_FIGURE_PATTERNS = _ref_patterns("Figure", "Figures", "Fig", "Figs")
_TABLE_PATTERNS  = _ref_patterns("Table",  "Tables",  "Tab", "Tabs")
_BOX_PATTERNS    = _ref_patterns("Box",    "Boxes",   "Box", "Boxes")
# Chapter references (simpler — no abbreviation, no dotted numbers)
_CHAPTER_PATTERNS = [
    (re.compile(r"\bChapters?\s+\d+\s+through\s+\d+\b", re.IGNORECASE), "ref_chapters_through"),
    (re.compile(r"\bChapters?\s+\d+\s+to\s+\d+\b",      re.IGNORECASE), "ref_chapters_to"),
    (re.compile(r"\bChapters?\s+\d+\s+and\s+\d+\b",     re.IGNORECASE), "ref_chapters_and"),
    (re.compile(r"\bChapters?\s+\d+\u2013\d+\b",         re.IGNORECASE), "ref_chapters_endash"),
    (re.compile(r"\bChapters?\s+\d+-\d+\b",              re.IGNORECASE), "ref_chapters_dash"),
    (re.compile(r"\bChapter\s+\d+\b",                    re.IGNORECASE), "ref_chapter_single"),
    (re.compile(r"\bchapters?\s+\d+\b"),                                  "ref_chapter_lc"),
]
def detect_caption_labels(seg: Segment) -> Iterable[Finding]:
    """Extract figure/table/box labels from segments identified as captions."""
    if seg.exclude_reason != "caption":
        return
    for group, label in [
        (_FIGURE_PATTERNS,  "Figure reference style"),
        (_TABLE_PATTERNS,   "Table reference style"),
        (_BOX_PATTERNS,     "Box reference style"),
    ]:
        for pat, rule_id, _lbl in group:
            # We only want to capture the label itself, which usually appears 
            # at the beginning of the caption paragraph.
            m = pat.match(seg.text.lstrip())
            if m:
                # Adjust match start/end if we lstripped
                offset = len(seg.text) - len(seg.text.lstrip())
                # Re-match against the original text to get correct indices
                m_orig = pat.search(seg.text, pos=offset)
                if m_orig and m_orig.start() == offset:
                    cap_rule_id = rule_id.replace('ref_', 'cap_')
                    yield _f(seg, m_orig, "te_point", cap_rule_id, label, "reference_style", pat=pat)
                break # Only capture the leading label to avoid internal references
def detect_references(seg: Segment) -> Iterable[Finding]:
    for group, label in [
        (_FIGURE_PATTERNS,  "Figure reference style"),
        (_TABLE_PATTERNS,   "Table reference style"),
        (_BOX_PATTERNS,     "Box reference style"),
    ]:
        for pat, rule_id, _lbl in group:
            for m in iter_unmasked_matches(pat, seg.text, seg.mask):
                yield _f(seg, m, "te_point", rule_id, label, "reference_style", pat=pat)
    for pat, rule_id in _CHAPTER_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Chapter reference style", "reference_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 13: number style (0–9 numeral vs spelled out and 0–99 numeral vs spelled out)
# ---------------------------------------------------------------------------
DIGIT_SPELL_PATTERNS = [
    (re.compile(r"(?<!\d)\b[0-9]\b(?!\d)"), "num_single_numeral"),
    (re.compile(r"\b(?:zero|one|two|three|four|five|six|seven|eight|nine)\b", re.IGNORECASE), "num_single_spelled"),
    (re.compile(r"(?<!\d)\b(?:[1-9]?[0-9])\b(?!\d)"), "num_double_numeral"),
    (re.compile(r"\b(?:ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|(?:twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)(?:[- \u2013](?:one|two|three|four|five|six|seven|eight|nine))?)\b", re.IGNORECASE), "num_double_spelled"),
    (re.compile(r"(?<!\d)\b0\b(?!\.)"), "num_zero_numeral"),
    (re.compile(r"\bzero\b", re.IGNORECASE), "num_zero_spelled"),
]
_CITATION_LABEL_PREFIX = re.compile(
    r'(?:'
    r'\b(?:Fig(?:ure)?s?|Figs?|Tables?|Tab(?:le)?s?|Box(?:es)?|'
    r'Chapters?|CHAPTER|Sections?|Eq(?:uation)?s?|'
    r'App(?:endix)?(?:endices)?|Ref(?:erence)?s?)\s*\.?\s*|'
    r'\d+\.\s*'  # matches "3." in "Fig. 3.[1]"
    r')$',
    re.IGNORECASE,
)
def detect_digit_spell(seg: Segment) -> Iterable[Finding]:
    # First, collect all range matches so we can exclude individual numbers within ranges
    range_matches = set()
    for pat, _ in RANGE_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            # Mark the start and end positions of range matches
            range_matches.add((m.start(), m.end()))
    # Track which (position, rule_id) pairs we've already yielded to avoid duplicates
    seen = set()
    for pat, rule_id in DIGIT_SPELL_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            # Skip if this number is part of a range (e.g., the "0" or "5" in "0-5")
            is_in_range = False
            for range_start, range_end in range_matches:
                if range_start <= m.start() < range_end:
                    is_in_range = True
                    break
            if is_in_range:
                continue
            prefix = seg.text[max(0, m.start() - 30) : m.start()]
            if _CITATION_LABEL_PREFIX.search(prefix):
                continue
            # Skip numbers followed by degree symbol (measurements: 30°, 45°, etc.)
            if m.end() < len(seg.text) and seg.text[m.end()] in '°º':
                continue
            # Skip digits adjacent to word characters (e.g., "four3" should not flag the "3")
            if rule_id.startswith("num_") and rule_id.endswith("_numeral"):
                # Check if preceded by word char (letter/digit)
                if m.start() > 0 and seg.text[m.start()-1].isalpha():
                    continue
                # Check if followed by word char (letter/digit)
                if m.end() < len(seg.text) and seg.text[m.end()].isalpha():
                    continue
            # Skip duplicate: if we've already reported this position with a different rule_id, skip it
            pos_key = (m.start(), m.end())
            if pos_key in seen:
                continue
            seen.add(pos_key)
            yield _f(seg, m, "te_point", rule_id, "Single-digit number style", "digit_spell_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 14: trademark / registered / copyright symbols
# ---------------------------------------------------------------------------
TRADEMARK_PATTERNS = [
    (re.compile(r"\u2122"), "sym_trademark_char"),
    (re.compile(r"\(TM\)", re.IGNORECASE), "sym_trademark_text"),
    (re.compile(r"\u00ae"), "sym_registered_char"),
    (re.compile(r"\(R\)", re.IGNORECASE), "sym_registered_text"),
    (re.compile(r"\u00a9"), "sym_copyright_char"),
    (re.compile(r"\(C\)", re.IGNORECASE), "sym_copyright_text"),
]
def detect_trademark(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in TRADEMARK_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Trademark/Register/Copyright symbol", "trademark_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 15: quotation mark style
# ---------------------------------------------------------------------------
QUOTE_PATTERNS = [
    (re.compile(r"\u201c[^\u201d]+\u201d"), "quote_double_curly"),
    (re.compile(r"\u2018[^\u2019]+\u2019"), "quote_single_curly"),
    (re.compile(r'"[^"]+"'), "quote_double_straight"),
    (re.compile(r"'[^']{2,}'"), "quote_single_straight"),
]
def detect_quote_style(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in QUOTE_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Quotation mark style", "quote_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 16: fold usage (Xfold, X-fold, X fold)
# ---------------------------------------------------------------------------
_NUMERAL = r"\d+"
_SPELLED = r"(?:one|two|three|four|five|six|seven|eight|nine|ten|hundred)"
FOLD_PATTERNS = [
    (re.compile(rf"\b{_NUMERAL}-fold\b", re.IGNORECASE), "fold_numeral_hyphen"),
    (re.compile(rf"\b{_NUMERAL}\s+fold\b", re.IGNORECASE), "fold_numeral_open"),
    (re.compile(rf"\b{_NUMERAL}fold\b", re.IGNORECASE), "fold_numeral_closed"),
    (re.compile(rf"\b{_SPELLED}-fold\b", re.IGNORECASE), "fold_word_hyphen"),
    (re.compile(rf"\b{_SPELLED}\s+fold\b", re.IGNORECASE), "fold_word_open"),
    (re.compile(rf"\b{_SPELLED}fold\b", re.IGNORECASE), "fold_word_closed"),
]
def detect_fold_style(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in FOLD_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "fold usage", "fold_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 17: times (numeral/word "X times")
# ---------------------------------------------------------------------------
_SPELLED_NUM = r"(?:once|twice|one|two|three|four|five|six|seven|eight|nine|ten)"
TIMES_WORD_PATTERNS = [
    (re.compile(r"\b\d+-times\b", re.IGNORECASE), "times_numeral_hyphen"),
    (re.compile(r"\b\d+\s+times\b", re.IGNORECASE), "times_numeral"),
    (re.compile(rf"\b{_SPELLED_NUM}\s+times\b", re.IGNORECASE), "times_word"),
    (re.compile(r"\btwice\b", re.IGNORECASE), "times_twice"),
]
def detect_times_style(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in TIMES_WORD_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Times (numeral/word)", "times_word_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 18: date formats — UK and US long-form + numeric
# ---------------------------------------------------------------------------
_MON_ABBR = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*"
_MON_FULL = r"(?:January|February|March|April|May|June|July|August|September|October|November|December)"
DATE_PATTERNS = [
    # UK long form: 1 January 2020 / 1 Jan 2020
    (re.compile(rf"\b\d{{1,2}}\s+{_MON_ABBR}\s+\d{{4}}\b"), "date_uk_long"),
    # US long form: January 1, 2020
    (re.compile(rf"\b{_MON_FULL}\s+\d{{1,2}},\s+\d{{4}}\b"), "date_us_long"),
    # Numeric slash D/M/YY or M/D/YY or D/M/YYYY
    (re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b"), "date_numeric_slash"),
    # Numeric dot D.M.YY
    (re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b"), "date_numeric_dot"),
    # Numeric dash D-M-YY
    (re.compile(r"\b\d{1,2}-\d{1,2}-\d{2,4}\b"), "date_numeric_dash"),
]
def detect_date_format(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in DATE_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Date format", "date_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 19: inline list marker style
# ---------------------------------------------------------------------------
LIST_MARKER_PATTERNS = [
    (re.compile(r"\([A-Z]\)"), "list_paren_alpha_uc"),
    (re.compile(r"\([a-z]\)"), "list_paren_alpha_lc"),
    (re.compile(r"\([IVX]+\)"), "list_paren_roman_uc"),
    (re.compile(r"\([ivx]+\)"), "list_paren_roman_lc"),
    (re.compile(r"\(\d+\)"), "list_paren_num"),
    (re.compile(r"\b[A-Z]\)"), "list_alpha_close_uc"),
    (re.compile(r"\b[a-z]\)"), "list_alpha_close_lc"),
    (re.compile(r"\b[IVX]+\)"), "list_roman_close_uc"),
    (re.compile(r"\b[ivx]+\)"), "list_roman_close_lc"),
    (re.compile(r"(?<![.,\-\s\d])\b\d+\)"), "list_num_close"),
]
def detect_list_marker(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in LIST_MARKER_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Inline list marker style", "list_marker_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 20: positional cross-references
# ---------------------------------------------------------------------------
CROSS_REF_PATTERNS = [
    (re.compile(r"\bsee\s+above\b", re.IGNORECASE), "pos_see_above"),
    (re.compile(r"\bsee\s+below\b", re.IGNORECASE), "pos_see_below"),
    (re.compile(r"\bdiscussed\s+above\b", re.IGNORECASE), "pos_discussed_above"),
    (re.compile(r"\bdiscussed\s+below\b", re.IGNORECASE), "pos_discussed_below"),
    (re.compile(r"\binfra\b", re.IGNORECASE), "pos_infra"),
    (re.compile(r"\bsupra\b", re.IGNORECASE), "pos_supra"),
]
def detect_cross_ref(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in CROSS_REF_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Positional reference style", "cross_ref_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 21: ordinal number style
# ---------------------------------------------------------------------------
ORDINAL_PATTERNS = [
    (re.compile(r"\b\d+st\b"), "ordinal_st"),
    (re.compile(r"\b\d+nd\b"), "ordinal_nd"),
    (re.compile(r"\b\d+rd\b"), "ordinal_rd"),
    (re.compile(r"\b\d+th\b"), "ordinal_th"),
    (re.compile(r"\b(?:first|second|third|fourth|fifth|sixth|seventh|eighth|ninth|tenth|"
                r"eleventh|twelfth|thirteenth|fourteenth|fifteenth|sixteenth|seventeenth|"
                r"eighteenth|nineteenth|twentieth|twenty-first|twenty-second|"
                r"thirtieth|fortieth|fiftieth|sixtieth|seventieth|eightieth|ninetieth|"
                r"hundredth|thousandth)\b", re.IGNORECASE), "ordinal_spelled"),
]
def detect_ordinal(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in ORDINAL_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Ordinal number style", "ordinal_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 22: fraction style
# ---------------------------------------------------------------------------
FRACTION_PATTERNS = [
    # Unicode fraction symbols: ¼ ½ ¾ ⅓ ⅔
    (re.compile(r"[\u00bc\u00bd\u00be\u2153\u2154]"), "fraction_symbol"),
    # Numeric slash fraction
    (re.compile(r"\b\d+/\d+\b"), "fraction_slash"),
    # Spelled fractions
    (re.compile(r"\b(?:one|two|three)-(?:half|third|quarter|fourth|fifth|sixth|"
                r"seventh|eighth|ninth|tenth)s?\b", re.IGNORECASE), "fraction_spelled"),
]
def detect_fraction(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in FRACTION_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Fraction style", "fraction_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 23: temperature / degree style
# ---------------------------------------------------------------------------
TEMP_PATTERNS = [
    # °C and °F (most specific first)
    (re.compile(r"\u00b0C\b"), "degree_celsius"),
    (re.compile(r"\u00b0F\b"), "degree_fahrenheit"),
    # degree spelled out with C/F
    (re.compile(r"\bdeg(?:rees?)?\s+C\b", re.IGNORECASE), "degree_deg_c"),
    (re.compile(r"\bdeg(?:rees?)?\s+F\b", re.IGNORECASE), "degree_deg_f"),
    (re.compile(r"\bCelsius\b", re.IGNORECASE), "degree_celsius_spelled"),
    (re.compile(r"\bFahrenheit\b", re.IGNORECASE), "degree_fahrenheit_spelled"),
    # bare degree symbol (not followed by C/F)
    (re.compile(r"\u00b0(?![CF])"), "degree_symbol_alone"),
    # degrees spelled without unit
    (re.compile(r"\bdeg(?:rees?)\b(?!\s+[CF])", re.IGNORECASE), "degree_spelled"),
]
def detect_temp(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in TEMP_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Temperature/degree style", "temp_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 24: comparison operators
# ---------------------------------------------------------------------------
COMPARISON_PATTERNS = [
    (re.compile(r"\u2265|>="), "cmp_gte_sym"),
    (re.compile(r"\u2264|<="), "cmp_lte_sym"),
    (re.compile(r"\bgreater\s+than\s+or\s+equal\b", re.IGNORECASE), "cmp_gte_words"),
    (re.compile(r"\bless\s+than\s+or\s+equal\b", re.IGNORECASE), "cmp_lte_words"),
    (re.compile(r"\bgreater\s+than\b", re.IGNORECASE), "cmp_gt_words"),
    (re.compile(r"\bless\s+than\b", re.IGNORECASE), "cmp_lt_words"),
    (re.compile(r"\bmore\s+than\b", re.IGNORECASE), "cmp_more_words"),
    (re.compile(r"\bapproximately\b", re.IGNORECASE), "cmp_approx_words"),
    (re.compile(r"(?<![=<>!])>(?!=)"), "cmp_gt_sym"),
    (re.compile(r"(?<![=<>!])<(?!=)"), "cmp_lt_sym"),
    (re.compile(r"~"), "cmp_tilde"),
]
def detect_comparison(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in COMPARISON_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Comparison operator style", "comparison_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 25: virgule/per usage
# ---------------------------------------------------------------------------
VIRGULE_PATTERNS = [
    (re.compile(r"\bper\b"), "per_word"),
    (re.compile(r"(?<!\w)/(?!\w)"), "virgule_slash"),
]
def detect_virgule(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in VIRGULE_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Virgule/per usage", "virgule_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 26: gas abbreviations
# ---------------------------------------------------------------------------
GAS_PATTERNS = [
    (re.compile(r"\bPaCO2\b"), "gas_paco2_inline"),
    (re.compile(r"\bPaCO\u2082\b"), "gas_paco2_sub"),
    (re.compile(r"\bPCO2\b"), "gas_pco2_inline"),
    (re.compile(r"\bPCO\u2082\b"), "gas_pco2_sub"),
    (re.compile(r"\bPaO2\b"), "gas_pao2_inline"),
    (re.compile(r"\bPaO\u2082\b"), "gas_pao2_sub"),
    (re.compile(r"\bFIO2\b|\bFiO2\b"), "gas_fio2_inline"),
    (re.compile(r"\bFIO\u2082\b|\bFiO\u2082\b"), "gas_fio2_sub"),
    (re.compile(r"\bSaO2\b"), "gas_sao2_inline"),
    (re.compile(r"\bSaO\u2082\b"), "gas_sao2_sub"),
    (re.compile(r"\bSpO2\b"), "gas_spo2_inline"),
    (re.compile(r"\bSpO\u2082\b"), "gas_spo2_sub"),
]
def detect_gas_abbrev(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in GAS_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Gas abbreviations", "gas_abbrev_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 27: SI / common units (spelled vs abbreviated)
# ---------------------------------------------------------------------------
SI_PATTERNS = [
    (re.compile(r"\bmg\b"), "unit_mg_abbr"),
    (re.compile(r"\bmilligrams?\b", re.IGNORECASE), "unit_mg_spelled"),
    (re.compile(r"\bkg\b"), "unit_kg_abbr"),
    (re.compile(r"\bkilograms?\b", re.IGNORECASE), "unit_kg_spelled"),
    (re.compile(r"\bmL\b|\bml\b"), "unit_ml_abbr"),
    (re.compile(r"\bmillilitr(?:es?|ers?)\b", re.IGNORECASE), "unit_ml_spelled"),
    (re.compile(r"\b\u03bcg\b|\bmcg\b"), "unit_mcg_abbr"),
    (re.compile(r"\bmicrograms?\b", re.IGNORECASE), "unit_mcg_spelled"),
    (re.compile(r"\bmm\b"), "unit_mm_abbr"),
    (re.compile(r"\bmillimetr(?:es?|ers?)\b", re.IGNORECASE), "unit_mm_spelled"),
    (re.compile(r"\bcm\b"), "unit_cm_abbr"),
    (re.compile(r"\bcentimetr(?:es?|ers?)\b", re.IGNORECASE), "unit_cm_spelled"),
    (re.compile(r"\bkm\b"), "unit_km_abbr"),
    (re.compile(r"\bkilometr(?:es?|ers?)\b", re.IGNORECASE), "unit_km_spelled"),
    (re.compile(r"\bh\b(?=\s)"), "unit_hour_h"),
    (re.compile(r"\bhr\b"), "unit_hour_hr"),
    (re.compile(r"\bhours?\b", re.IGNORECASE), "unit_hour_spelled"),
    (re.compile(r"\bmin\b"), "unit_min_abbr"),
    (re.compile(r"\bminutes?\b", re.IGNORECASE), "unit_min_spelled"),
    (re.compile(r"(?<!['’\u2019`´])\bs\b(?=\s)"), "unit_sec_s"),
    (re.compile(r"\bsec\b"), "unit_sec_abbr"),
    (re.compile(r"\bseconds?\b", re.IGNORECASE), "unit_sec_spelled"),
]
def detect_si_unit(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in SI_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "SI/unit style", "si_unit_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 28: numerical citations within square brackets in caption segments
# ---------------------------------------------------------------------------
CITATION_PATTERNS = [
    (re.compile(r"(?P<citation>\[\d+(?:[\u2013\-]\d+)?(?:,\s*\d+(?:[\u2013\-]\d+)?)*\])"), "citation_bracket_numeric"),
    # Add more patterns here for other citation styles if needed (e.g., author-year)
]
def detect_citations(seg: Segment) -> Iterable[Finding]:
    # Only run citation detection within segments identified as captions
    if seg.exclude_reason == "caption":
        for pat, rule_id in CITATION_PATTERNS:
            for m in pat.finditer(seg.text):
                yield _f(seg, m, "te_point", rule_id, "Citation style", "citation_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 13: US / UK date format
# ---------------------------------------------------------------------------
US_UK_DATE_PATTERNS = [
    (re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b"), "date_numeric_slash"),
    (re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b"), "date_numeric_dot"),
    (re.compile(r"\b\d{1,2}-\d{1,2}-\d{2,4}\b"), "date_numeric_dash"),
    (re.compile(r"\b\d{1,2}\s+Jan\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\bJan\s+\d{1,2}\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\b\d{1,2}\s+Jan\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\bJan\s+\d{1,2}\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b"), "date_numeric_slash"),
    (re.compile(r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b"), "date_numeric_dot"),
    (re.compile(r"\b\d{1,2}-\d{1,2}-\d{2,4}\b"), "date_numeric_dash"),
    (re.compile(r"\b\d{1,2}\s+Jan\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\bJan\s+\d{1,2}\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\b\d{1,2}\s+Jan\s+\d{4}\b"), "date_us_long"),
    (re.compile(r"\bJan\s+\d{1,2}\s+\d{4}\b"), "date_us_long"),
]
def detect_date_format(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in US_UK_DATE_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Date format", "date_style", pat=pat)
# ---------------------------------------------------------------------------
# Main dispatch
# ---------------------------------------------------------------------------
TE_DETECTORS: list = [
    detect_percent,
    detect_ellipsis,
    detect_ampm,
    detect_era,
    detect_number_ranges,
    detect_leading_zero,
    detect_times,
    detect_versus,
    detect_latin_abbr,
    detect_century,
    detect_thousands,
    detect_references,
    detect_digit_spell,
    detect_trademark,
    detect_quote_style,
    detect_fold_style,
    detect_times_style,
    detect_date_format,
    detect_list_marker,
    detect_cross_ref,
    detect_ordinal,
    detect_fraction,
    detect_temp,
    detect_comparison,
    detect_virgule,
    detect_gas_abbrev,
    detect_si_unit,
    detect_citations,
]
def run_te_rules(seg: Segment) -> list[Finding]:
    # Skip TE rules for front matter and references sections
    if seg.region in ("front", "references"):
        return []
    out: list[Finding] = []
    for fn in TE_DETECTORS:
        out.extend(fn(seg))
    return out
# ---------------------------------------------------------------------------
# Rule 29: Caps/Lowercase after colon
# ---------------------------------------------------------------------------
COLON_PATTERNS = [
    (re.compile(r":\s+[A-Z]"), "caps_after_colon"),
    (re.compile(r":\s+[a-z]"), "lowercase_after_colon"),
]
def detect_colon_case(seg):
    for pat, rule_id in COLON_PATTERNS:
        label = "Caps after colon (:)" if "caps" in rule_id else "Lowercase after colon (:)"
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, label, "colon_case_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 30: Spaced hyphens
# ---------------------------------------------------------------------------
SPACED_HYPHEN_PATTERNS = [
    (re.compile(r"\s+-\s+"), "spaced_hyphens"),
]
def detect_spaced_hyphens(seg):
    for pat, rule_id in SPACED_HYPHEN_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Spaced hyphens", "spaced_hyphen_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 31: Currency
# ---------------------------------------------------------------------------
CURRENCY_PATTERNS = [
    (re.compile(r"[$£€¥]"), "currency_symbols"),
    (re.compile(r"\b(?:dollars?|pounds?|euros?|yen)\b", re.IGNORECASE), "currency_spelled"),
]
def detect_currency(seg):
    for pat, rule_id in CURRENCY_PATTERNS:
        label = "Currency symbols" if "symbols" in rule_id else "Currency spelled out"
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, label, "currency_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 32: (-al) endings - Medical/Scientific terms only
# ---------------------------------------------------------------------------
AL_ENDING_WORDS = {
    "biological", "chronological", "cytological", "dermatological",
    "ecological", "embryological", "epidemiological", "etiological",
    "gynecological", "hematological", "histological", "immunological",
    "morphological", "oncological", "ophthalmological", "pathological",
    "pharmacological", "radiological", "rheumatological", "sociological",
    "symptomatological", "toxicological", "traumatological", "urological",
    "anatomical", "physiological", "neurological"
}
AL_ENDING_PATTERNS = [
    (re.compile(r"\b(" + "|".join(AL_ENDING_WORDS) + r")\b", re.IGNORECASE), "al_endings"),
]
def detect_al_endings(seg):
    for pat, rule_id in AL_ENDING_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "(-al) endings", "al_ending_style")
# ---------------------------------------------------------------------------
# Rule 33: Chart / Diagram / Image / Illustration reference detection
# ---------------------------------------------------------------------------
_CHART_PATTERNS = _ref_patterns("Chart", "Charts", "Chart", "Charts")
_LOWERCASE_MEDIA_PATTERNS = [
    (re.compile(r"\bdiagram\s+\d+\b"),   "ref_diagram_lc"),
    (re.compile(r"\bdiagrams\s+\d+\b"),  "ref_diagrams_lc"),
    (re.compile(r"\bimage\s+\d+\b"),     "ref_image_lc"),
    (re.compile(r"\bimages\s+\d+\b"),    "ref_images_lc"),
    (re.compile(r"\billustration\s+\d+\b"),  "ref_illustration_lc"),
    (re.compile(r"\billustrations\s+\d+\b"), "ref_illustrations_lc"),
]
def detect_chart_refs(seg: Segment) -> Iterable[Finding]:
    for group, label in [(_CHART_PATTERNS, "Chart reference style")]:
        for pat, rule_id, _lbl in group:
            for m in iter_unmasked_matches(pat, seg.text, seg.mask):
                yield _f(seg, m, "te_point", rule_id, label, "reference_style", pat=pat)
def detect_chart_caption_labels(seg: Segment) -> Iterable[Finding]:
    if seg.exclude_reason != "caption":
        return
    for group, label in [(_CHART_PATTERNS, "Chart reference style")]:
        for pat, rule_id, _lbl in group:
            m = pat.match(seg.text.lstrip())
            if m:
                offset = len(seg.text) - len(seg.text.lstrip())
                m_orig = pat.search(seg.text, pos=offset)
                if m_orig and m_orig.start() == offset:
                    cap_rule_id = rule_id.replace("ref_", "cap_")
                    yield _f(seg, m_orig, "te_point", cap_rule_id, label, "reference_style", pat=pat)
                break
def detect_lowercase_media_refs(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in _LOWERCASE_MEDIA_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            label = rule_id.split("_")[1].capitalize() + " reference (lowercase)"
            yield _f(seg, m, "te_point", rule_id, label, "reference_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 34: Section reference detection (full pattern set via _ref_patterns)
# ---------------------------------------------------------------------------
_SECTION_PATTERNS = _ref_patterns("Section", "Sections", "Section", "Sections")
def detect_section_refs(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id, _lbl in _SECTION_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Section reference style", "reference_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 35: Currency expansion (full symbol + spelled-out list)
# ---------------------------------------------------------------------------
CURRENCY_PATTERNS = [
    (re.compile(
        r'[$£€¥₹₽₩₦₴₫฿₪₨₱₭₮₸₺₾₿¢]'
        r'|(?<!\w)(?:A\$|NZ\$|HK\$|CAD?\$|SGD?\$|Mex\$|R\$|RD\$|Col\$|Ch\$|Bd?\$|NT\$)'
        r'|(?<!\w)(?:CHF|kr|Rs|SR|ft|ZK|MK|JD|LBP)(?!\w)'
    ), "currency_symbols"),
    (re.compile(
        r'\b(?:dollar|euro|pound|yuan|rupee|ruble|rubric|won|peso|boliviano|franc|'
        r'dinar|rial|riyal|sheqel|kwacha|naira|zloty|rand|krona|krone|baht|lira|'
        r'shilling|hryvnia|sterling|dong|forint|koruna|birr|cedi|soles?|'
        r'renminbi|rouble)s?\b', re.IGNORECASE
    ), "currency_spelled"),
]
def detect_currency(seg: Segment) -> Iterable[Finding]:  # type: ignore[no-redef]
    for pat, rule_id in CURRENCY_PATTERNS:
        label = "Currency symbols" if rule_id == "currency_symbols" else "Currency spelled out"
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, label, "currency_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 36: Abbreviations (2+ consecutive uppercase letters)
# ---------------------------------------------------------------------------
_ABBREV_PAT = re.compile(r'\b[A-Z]{2,}\b')
def detect_abbreviations(seg: Segment) -> Iterable[Finding]:
    style_lower = (seg.style or "").lower()
    if style_lower.startswith("heading") or "title" in style_lower:
        rule_id, label = "abbrev_heading", "Abbreviations in heading"
    elif seg.source == "table":
        rule_id, label = "abbrev_table", "Abbreviations in CT/TC"
    elif seg.exclude_reason == "caption":
        rule_id, label = "abbrev_caption", "Abbreviations in FC/TC"
    else:
        rule_id, label = "abbrev_body", "Abbreviations"
    for m in iter_unmasked_matches(_ABBREV_PAT, seg.text, seg.mask):
        yield _f(seg, m, "te_point", rule_id, label, "abbreviation_style", pat=_ABBREV_PAT)
# ---------------------------------------------------------------------------
# Rule 37: Latin terms (editorial flag — italics check is manual)
# ---------------------------------------------------------------------------
_LATIN_TERMS = [
    "a priori", "a posteriori", "ad hoc", "ad lib", "alma mater",
    "bona fide", "carpe diem", "caveat", "de facto", "ex officio",
    "ex post", "ibid", "in situ", "in vitro", "in vivo",
    "inter alia", "ipso facto", "magnum opus", "per capita",
    "per diem", "per se", "postmortem", "prima facie", "status quo",
    "sui generis", "vice versa", "viz", r"vis-à-vis",
    "post hoc", "infra", "supra", "veto",
]
_LATIN_TERM_PATTERNS = [
    (re.compile(rf'\b{re.escape(t)}\b', re.IGNORECASE), "latin_term")
    for t in _LATIN_TERMS
]
def detect_latin_terms(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in _LATIN_TERM_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Latin terms (check italics)", "latin_term_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 38: Paras ending without full stop
# ---------------------------------------------------------------------------
_HEADING_TAG_RE = re.compile(r'^</?[A-Za-z][A-Za-z0-9._-]*>', re.IGNORECASE)
def detect_para_ending(seg: Segment) -> Iterable[Finding]:
    text = seg.text.rstrip()
    if not text or seg.exclude_reason == "caption":
        return
    if _HEADING_TAG_RE.match(text):
        return
    last = text[-1]
    if last not in '.!?……':
        # Fake a match at the last character so we have position info
        m = re.search(r'.{1}$', text)
        if m:
            yield _f(seg, m, "te_point", "para_no_full_stop",
                     "Para ending without full stop", "para_ending_style")
# ---------------------------------------------------------------------------
# Rule 39: Genus species patterns
# ---------------------------------------------------------------------------
_COMMON_GENERA = [
    "Acinetobacter", "Actinomyces", "Adenovirus", "Aspergillus", "Bacillus",
    "Bacteroides", "Bifidobacterium", "Bordetella", "Borrelia", "Brucella",
    "Campylobacter", "Candida", "Chlamydia", "Chlamydophila", "Clostridium",
    "Corynebacterium", "Cryptococcus", "Cytomegalovirus", "Enterobacter",
    "Enterococcus", "Escherichia", "Fusobacterium", "Haemophilus",
    "Helicobacter", "Hepatitis", "Herpes", "Histoplasma", "Klebsiella",
    "Lactobacillus", "Legionella", "Leishmania", "Leptospira", "Listeria",
    "Mycobacterium", "Mycoplasma", "Neisseria", "Nocardia", "Papillomavirus",
    "Pasteurella", "Plasmodium", "Pneumocystis", "Proteus", "Pseudomonas",
    "Rickettsia", "Rotavirus", "Rubella", "Salmonella", "Schistosoma",
    "Shigella", "Staphylococcus", "Streptococcus", "Toxoplasma", "Treponema",
    "Trichomonas", "Trypanosoma", "Vibrio", "Yersinia"
]
_GENERA_PAT = "|".join(_COMMON_GENERA)
_GENUS_SPECIES_PATTERNS = [
    (re.compile(rf'\b(?:{_GENERA_PAT})\s+[a-z]{{3,}}\b'), "genus_species_full"),   # Escherichia coli
    (re.compile(r'\b[A-Z]\.\s+[a-z]{3,}\b'),          "genus_species_abbr"),  # E. coli
    (re.compile(r'\b[A-Z]\s+[a-z]{3,}\b'),             "genus_species_nodot"), # E coli
]
def detect_genus_species(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in _GENUS_SPECIES_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "Genus species (check italic)", "genus_species_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 40: Greek letters
# ---------------------------------------------------------------------------
_GREEK_PATTERNS = [
    (None, "greek_lowercase"),
    (None, "greek_uppercase"),
]
def detect_greek_letters(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in _GREEK_PATTERNS:
        if pat is None:
            continue
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            label = "Greek letters (lowercase)" if "lower" in rule_id else "Greek letters (uppercase)"
            yield _f(seg, m, "te_point", rule_id, label, "greek_letter_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 41: AUX verbs in headings and captions
# ---------------------------------------------------------------------------
_AUX_VERB_WORDS = ["be", "am", "is", "do", "are", "was", "has", "had", "did", "may", "can", "will"]
_AUX_VERB_PATTERNS = [
    (re.compile(rf'\b{w}\b', re.IGNORECASE), f"aux_verb_{w}")
    for w in _AUX_VERB_WORDS
]
def detect_aux_verbs(seg: Segment) -> Iterable[Finding]:
    style_lower = (seg.style or "").lower()
    is_heading = style_lower.startswith("heading") or "title" in style_lower
    is_caption = seg.source == "table" or seg.exclude_reason == "caption"
    if not (is_heading or is_caption):
        return
    for pat, rule_id in _AUX_VERB_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, f"AUX verb in heading/caption", "aux_verb_style", pat=pat)
# ---------------------------------------------------------------------------
# Rule 42: WK / publisher-specific flagged terms
# ---------------------------------------------------------------------------
_WK_TERM_PATTERNS = [
    (re.compile(r'\btranssexual\b', re.IGNORECASE),          "wk_transsexual"),
    (re.compile(r'\bConcepts\s+in\s+Action\b', re.IGNORECASE), "wk_concepts_action"),
    (re.compile(r'\bWatch\s*&\s*Learn\b', re.IGNORECASE),    "wk_watch_learn"),
    (re.compile(r'\bPractice\s*&\s*Learn\b', re.IGNORECASE), "wk_practice_learn"),
    (re.compile(r'\bInteractive\s+Tutorial\b', re.IGNORECASE), "wk_interactive_tutorial"),
    (re.compile(r'\bMongolian\s+spots?\b', re.IGNORECASE),   "wk_mongolian_spot"),
]
def detect_wk_terms(seg: Segment) -> Iterable[Finding]:
    for pat, rule_id in _WK_TERM_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield _f(seg, m, "te_point", rule_id, "WK term (query/validate)", "wk_term_style", pat=pat)
# ---------------------------------------------------------------------------
# Wire all late-defined detectors into TE_DETECTORS
# ---------------------------------------------------------------------------
TE_DETECTORS.extend([
    detect_colon_case,
    detect_spaced_hyphens,
    detect_currency,
    detect_al_endings,
    detect_chart_refs,
    detect_chart_caption_labels,
    detect_lowercase_media_refs,
    detect_section_refs,
    detect_abbreviations,
    detect_latin_terms,
    detect_para_ending,
    detect_genus_species,
    detect_greek_letters,
    detect_aux_verbs,
    detect_wk_terms,
])

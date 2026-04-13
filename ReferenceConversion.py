import os
import json
import logging
import re
from typing import Optional, Dict, List, Tuple
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

from gemini_ref_converter import convert_reference, CitationStyle, BIB_FIELDS

# ─────────────────────────────────────────────
# LOGGING
# ─────────────────────────────────────────────
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s: %(message)s")
logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# REFERENCE TYPE DETECTION (fallback if Gemini fails)
# ─────────────────────────────────────────────

def detect_source_style(raw_text: str) -> CitationStyle:
    stripped = raw_text.strip()
    if re.match(r'^\[?\d+\]?\.?\s+', stripped):
        return CitationStyle.AMA
    if re.search(r'\.\s+\d{4};', stripped):
        return CitationStyle.AMA
    if re.search(r'\bpp?\.\s*\d+[-–]\d+', stripped):
        return CitationStyle.APA
    if re.search(r'\(\d{4}[a-z]?\)', stripped):
        return CitationStyle.APA
    if re.search(r'https://doi\.org/', stripped):
        return CitationStyle.APA
    if re.search(r',\s+\d+\(\d+\),\s+\d+', stripped):
        return CitationStyle.APA
    return CitationStyle.APA


def detect_ref_type_from_metadata(metadata: Dict) -> str:
    return metadata.get("bib_reftype") or "unknown"


# ─────────────────────────────────────────────
# TITLE CASE / SENTENCE CASE HELPERS  (#34, #38)
# ─────────────────────────────────────────────

_PROPER_NOUNS = re.compile(
    r'\b('
    r'COVID(?:-19)?|HIV(?:/AIDS)?|AIDS|SARS(?:-CoV(?:-2)?)?|MERS|'
    r'COPD|PCR|DNA|RNA|mRNA|CT|MRI|ICU|ECG|EKG|'
    r'USA|UK|UN|EU|US|WHO|FDA|CDC|NIH|NHS|'
    r'English|French|German|Spanish|Italian|Chinese|Japanese|'
    r'Korean|Arabic|Russian|Portuguese|Dutch|Swedish|'
    r'American|European|Asian|African|Australian|Canadian|'
    r'British|Indian|Brazilian|Mexican|'
    r'Cochrane|PubMed|CrossRef|MEDLINE'
    r')\b',
    re.IGNORECASE,
)

def _to_sentence_case(text: str) -> str:
    """
    Sentence case: capitalise only the first word, first word after colon/em-dash,
    and known proper nouns. Safety net — Gemini usually does this, but applied
    locally to catch any misses.
    """
    if not text:
        return text
        
    alpha_count = sum(1 for c in text if c.isalpha())
    upper_count = sum(1 for c in text if c.isupper())
    
    # If more than 50% of alphabetical characters are uppercase, it's likely ALL CAPS.
    # In that case, we don't preserve acronyms to allow proper sentence casing.
    preserve_acronyms = True
    if alpha_count > 0 and (upper_count / alpha_count) > 0.5:
        preserve_acronyms = False

    acronym_spans = []
    if preserve_acronyms:
        # Match words with at least 2 uppercase letters, or specific lower-upper patterns like mRNA
        acronym_pattern = r'\b(?:[A-Za-z0-9]*[A-Z][A-Za-z0-9]*[A-Z][A-Za-z0-9]*|[a-z][A-Z][A-Za-z0-9]*)\b'
        acronym_spans = [(m.start(), m.end(), m.group()) for m in re.finditer(acronym_pattern, text)]

    proper_spans = [(m.start(), m.end(), m.group()) for m in _PROPER_NOUNS.finditer(text)]
    
    result = text[0].upper() + text[1:].lower() if len(text) > 1 else text.upper()
    
    for start, end, word in acronym_spans:
        result = result[:start] + word + result[end:]
        
    for start, end, word in proper_spans:
        result = result[:start] + word + result[end:]
        
    result = re.sub(r'([:;—]\s+)([a-z])', lambda m: m.group(1) + m.group(2).upper(), result)
    return result


_TITLE_CASE_SMALL = frozenset({
    "a","an","the","and","but","or","for","nor","on","at",
    "to","by","in","of","up","as","is","it","its","via","per","vs","et",
})

def _to_title_case(text: str) -> str:
    """Title case for journal names (APA rule)."""
    if not text:
        return text
    words = text.split()
    result = []
    for i, w in enumerate(words):
        if len(w) >= 2 and w.isupper():
            result.append(w)
        elif i == 0 or w.lower() not in _TITLE_CASE_SMALL:
            result.append(w[0].upper() + w[1:].lower() if len(w) > 1 else w.upper())
        else:
            result.append(w.lower())
    return " ".join(result)


# ─────────────────────────────────────────────
# PUBLISHER SUFFIX STRIPPER  (#58)
# ─────────────────────────────────────────────

_PUB_SUFFIX_RE = re.compile(
    r',?\s+(?:Co\.|Ltd\.?|Limited|Inc\.?|LLC|L\.L\.C\.|Corp\.?|'
    r'GmbH|S\.A\.|Pvt\.?|Pty\.?|(?:Pty|Pvt)\.?\s+Ltd\.?)\s*$',
    re.IGNORECASE,
)

def _strip_publisher_suffixes(pub: str) -> str:
    """Strip corporate-form suffixes from publisher names per APA 7th / AMA 11th."""
    if not pub:
        return pub
    cleaned = _PUB_SUFFIX_RE.sub("", pub).strip().rstrip(",").strip()
    return cleaned or pub


# ─────────────────────────────────────────────
# QUOTE NORMALISER  (#35)
# ─────────────────────────────────────────────

def _normalise_quotes(text: str) -> str:
    """
    Convert curly quotes back to straight quotes.
    Word's AutoCorrect manages smart quotes; inserting pre-curled quotes via
    python-docx creates track-changes noise against straight-quote originals.
    """
    return (
        text
        .replace('\u2018', "'").replace('\u2019', "'")
        .replace('\u201c', '"').replace('\u201d', '"')
    )


# ─────────────────────────────────────────────
# REF-TYPE HEURISTIC CORRECTOR  (#46, #49)
# ─────────────────────────────────────────────

def _fix_ref_type(meta: Dict, raw_text: str) -> Dict:
    """
    Post-correct Gemini's bib_reftype using simple heuristics.
    Returns a shallow copy of meta with bib_reftype corrected where needed.
    """
    rt = (meta.get("bib_reftype") or "").lower()
    fixed = dict(meta)

    # book → journal
    if rt == "book" and fixed.get("bib_journal") and fixed.get("bib_volume"):
        fixed["bib_reftype"] = "journal"
        logger.info("  [TypeFix] 'book' → 'journal'  (has journal+volume)")
    elif rt == "book" and re.search(r'\d{4}\s*;\s*\d+[\s(:]', raw_text):
        fixed["bib_reftype"] = "journal"
        logger.info("  [TypeFix] 'book' → 'journal'  (year;volume pattern in raw text)")
    elif rt == "book" and re.search(r',\s*\*?\d+\*?\s*\(\d+\)\s*,\s*\d+', raw_text):
        fixed["bib_reftype"] = "journal"
        logger.info("  [TypeFix] 'book' → 'journal'  (APA volume(issue),page pattern)")

    # book → book_chapter
    rt2 = (fixed.get("bib_reftype") or "").lower()
    if (rt2 in ("book", "journal") and
            fixed.get("bib_chaptertitle") and
            (fixed.get("bib_ed_surname") or re.search(r'\bIn[:\s]', raw_text))):
        fixed["bib_reftype"] = "book_chapter"
        logger.info("  [TypeFix] → 'book_chapter'  (chapter title + editor/In:)")

    # book → edited_book  (#49)
    rt3 = (fixed.get("bib_reftype") or "").lower()
    if (rt3 == "book" and
            fixed.get("bib_ed_surname") and
            not fixed.get("bib_surname") and
            not fixed.get("bib_chaptertitle")):
        fixed["bib_reftype"] = "edited_book"
        logger.info("  [TypeFix] → 'edited_book'  (editors, no authors, no chapter)")
    if (fixed.get("bib_reftype", "book") == "book" and
            re.search(r'\b(?:eds?)\.\s+\w', raw_text, re.IGNORECASE) and
            not fixed.get("bib_surname")):
        fixed["bib_reftype"] = "edited_book"
        logger.info("  [TypeFix] → 'edited_book'  (ed./eds. marker in raw text)")

    # book → conference
    rt4 = (fixed.get("bib_reftype") or "").lower()
    if (rt4 == "book" and
            (fixed.get("bib_conference") or
             re.search(r'\b(?:presented\s+at|proceedings\s+of|annual\s+(?:meeting|conference))\b',
                       raw_text, re.IGNORECASE))):
        fixed["bib_reftype"] = "conference"
        logger.info("  [TypeFix] → 'conference'  (conference keywords in raw text)")

    return fixed


# ─────────────────────────────────────────────
# INLINE CITATION GUARD  (#41)
# ─────────────────────────────────────────────

_INLINE_CITATION_RE = re.compile(
    r'^\s*\([^)]{1,80}\)\s*[.,]?\s*$'
)

def _looks_like_inline_citation(text: str) -> bool:
    """True if the paragraph looks like a standalone in-text citation to skip."""
    if _INLINE_CITATION_RE.match(text):
        return True
    if re.match(r'^\s*[\[\(]?\d[\d,\s\-–]+[\]\)]?\s*\.?\s*$', text):
        return True
    return False


# ─────────────────────────────────────────────
# FORMATTING FROM METADATA  (fallback)
# ─────────────────────────────────────────────

def format_apa_from_metadata(meta: Dict) -> str:
    ref_type = meta.get("bib_reftype", "journal")
    parts = []
    surnames = [s.strip() for s in (meta.get("bib_surname") or "").split("|") if s.strip()]
    fnames   = [f.strip() for f in (meta.get("bib_fname")   or "").split("|") if f.strip()]
    authors  = []
    for i, surname in enumerate(surnames):
        initial = fnames[i] if i < len(fnames) else ""
        initials_fmt = " ".join(f"{p[0]}." for p in initial.split() if p) if initial else ""
        authors.append(f"{surname}, {initials_fmt}".strip(", "))
    if authors:
        if len(authors) > 20:
            author_str = ", ".join(authors[:19]) + ", ... " + authors[-1]
        elif len(authors) > 1:
            author_str = ", ".join(authors[:-1]) + ", & " + authors[-1]
        else:
            author_str = authors[0]
        parts.append(author_str + ".")
    year = meta.get("bib_year", "n.d.")
    parts.append(f"({year}).")
    if ref_type == "journal":
        title   = meta.get("bib_article", "")
        journal = _to_title_case(meta.get("bib_journal", ""))
        volume  = meta.get("bib_volume", "")
        issue   = meta.get("bib_issue", "")
        fpage   = meta.get("bib_fpage", "")
        lpage   = meta.get("bib_lpage", "")
        doi     = meta.get("bib_doi", "")
        if title:   parts.append(f"{_to_sentence_case(title)}.")
        vol_issue = f"*{journal}*" if journal else ""
        if volume:  vol_issue += f", *{volume}*"
        if issue:   vol_issue += f"({issue})"
        pages = f"{fpage}–{lpage}" if fpage and lpage else fpage or lpage
        if pages:   vol_issue += f", {pages}"
        if vol_issue: parts.append(vol_issue + ".")
        if doi:     parts.append(f"https://doi.org/{doi}")
    elif ref_type in ("book", "edited_book"):
        book_title = meta.get("bib_book") or ""
        edition    = meta.get("bib_editionno", "")
        publisher  = _strip_publisher_suffixes(meta.get("bib_publisher", ""))
        doi        = meta.get("bib_doi", "")
        url        = meta.get("bib_url", "")
        title_str  = f"*{_to_sentence_case(book_title)}*" if book_title else ""
        if edition and _ordinal(edition) not in ("1st", "1", "first"):
            title_str += f" ({_ordinal(edition)} ed.)"
        if title_str: parts.append(title_str + ".")
        if publisher: parts.append(publisher + ".")
        if doi:       parts.append(f"https://doi.org/{doi}")
        elif url:     parts.append(url)
    elif ref_type == "book_chapter":
        chapter   = meta.get("bib_chaptertitle") or ""
        book      = meta.get("bib_book", "")
        edition   = meta.get("bib_editionno", "")
        volume    = meta.get("bib_volume", "")
        fpage     = meta.get("bib_fpage", "")
        lpage     = meta.get("bib_lpage", "")
        publisher = _strip_publisher_suffixes(meta.get("bib_publisher", ""))
        doi       = meta.get("bib_doi", "")
        ed_surnames = [s.strip() for s in (meta.get("bib_ed_surname") or "").split("|") if s.strip()]
        ed_fnames   = [f.strip() for f in (meta.get("bib_ed_fname")   or "").split("|") if f.strip()]
        if chapter: parts.append(f"{_to_sentence_case(chapter)}.")
        editors = []
        for i, s in enumerate(ed_surnames):
            ini = ed_fnames[i] if i < len(ed_fnames) else ""
            ini_fmt = " ".join(f"{p[0]}." for p in ini.split() if p) if ini else ""
            editors.append(f"{ini_fmt} {s}".strip())
        ed_label = "Ed." if len(editors) == 1 else "Eds."
        in_str = "In " + ", ".join(editors) + f" ({ed_label}.), " if editors else "In "
        book_str = f"*{_to_sentence_case(book)}*" if book else ""
        inner = []
        if edition and _ordinal(edition) not in ("1st", "1", "first"):
            inner.append(_ordinal(edition) + " ed.")
        if volume:
            inner.append("Vol. " + volume)
        if fpage:
            inner.append("pp. " + fpage + ("–" + lpage if lpage else ""))
        paren = " (" + ", ".join(inner) + ")" if inner else ""
        parts.append(in_str + book_str + paren + ".")
        if publisher: parts.append(publisher + ".")
        if doi:       parts.append(f"https://doi.org/{doi}")
    elif ref_type == "thesis":
        title  = meta.get("bib_title", "")
        deg    = meta.get("bib_deg", "Doctoral dissertation")
        school = meta.get("bib_school", "")
        url    = meta.get("bib_url", "")
        if title:
            bracket = f" [{deg}, {school}]" if school else f" [{deg}]"
            parts.append(f"*{_to_sentence_case(title)}*{bracket}.")
        if url: parts.append(url)
    elif ref_type == "conference":
        title    = meta.get("bib_title", "")
        conf     = meta.get("bib_conference", "")
        confloc  = meta.get("bib_conflocation", "")
        confdate = meta.get("bib_confdate", "")
        doi      = meta.get("bib_doi", "")
        if title: parts.append(f"*{_to_sentence_case(title)}* [Conference session].")
        conf_str = conf
        if confdate: conf_str += f", {confdate}"
        if confloc:  conf_str += f", {confloc}"
        if conf_str: parts.append(conf_str + ".")
        if doi: parts.append(f"https://doi.org/{doi}")
    elif ref_type in ("website", "ereference"):
        title    = meta.get("bib_title", "")
        site     = meta.get("bib_journal") or meta.get("bib_book", "")
        accessed = meta.get("bib_accessed", "")
        url      = meta.get("bib_url", "")
        if title:    parts.append(f"{_to_sentence_case(title)}.")
        if site:     parts.append(f"*{_to_title_case(site)}*.")
        if accessed: parts.append(f"Retrieved {accessed}, from")
        if url:      parts.append(url)
    elif ref_type == "report":
        title  = meta.get("bib_title", "")
        repnum = meta.get("bib_reportnum", "")
        inst   = _strip_publisher_suffixes(meta.get("bib_institution", ""))
        doi    = meta.get("bib_doi", "")
        url    = meta.get("bib_url", "")
        title_str = f"*{_to_sentence_case(title)}*" if title else ""
        if repnum: title_str += f" (Report No. {repnum})"
        if title_str: parts.append(title_str + ".")
        if inst:      parts.append(inst + ".")
        if doi:       parts.append(f"https://doi.org/{doi}")
        elif url:     parts.append(url)
    return " ".join(parts)


def format_ama_from_metadata(meta: Dict) -> str:
    ref_type = meta.get("bib_reftype", "journal")
    parts = []
    surnames = [s.strip() for s in (meta.get("bib_surname") or "").split("|") if s.strip()]
    fnames   = [f.strip() for f in (meta.get("bib_fname")   or "").split("|") if f.strip()]
    authors  = []
    for i, surname in enumerate(surnames):
        initial = fnames[i] if i < len(fnames) else ""
        initials_fmt = "".join(p[0] for p in initial.split() if p) if initial else ""
        authors.append(f"{surname} {initials_fmt}".strip())
    if authors:
        if len(authors) > 6:
            author_str = ", ".join(authors[:6]) + ", et al"
        else:
            author_str = ", ".join(authors)
        parts.append(author_str + ".")
    if ref_type == "journal":
        title   = meta.get("bib_title", "")
        journal = meta.get("bib_journal", "")
        year    = meta.get("bib_year", "")
        volume  = meta.get("bib_volume", "")
        issue   = meta.get("bib_issue", "")
        fpage   = meta.get("bib_fpage", "")
        lpage   = meta.get("bib_lpage", "")
        doi     = meta.get("bib_doi", "")
        if title:   parts.append(f"{_to_sentence_case(title)}.")
        vol_str = journal or ""
        if year:    vol_str += f". {year}"
        if volume:  vol_str += f";{volume}"
        if issue:   vol_str += f"({issue})"
        pages = f"{fpage}-{lpage}" if fpage and lpage else fpage or lpage
        if pages:   vol_str += f":{pages}"
        if vol_str: parts.append(vol_str + ".")
        if doi:     parts.append(f"doi:{doi}")
    elif ref_type in ("book", "edited_book"):
        book_title = meta.get("bib_book") or meta.get("bib_title", "")
        edition    = meta.get("bib_editionno", "")
        publisher  = _strip_publisher_suffixes(meta.get("bib_publisher", ""))
        year       = meta.get("bib_year", "")
        doi        = meta.get("bib_doi", "")
        url        = meta.get("bib_url", "")
        title_str  = _to_sentence_case(book_title) if book_title else ""
        if edition and _ordinal(edition) not in ("1st", "1", "first"):
            title_str += f". {_ordinal(edition)} ed."
        if title_str: parts.append(title_str + ".")
        if publisher: parts.append(publisher + ";")
        if year:      parts.append(year + ".")
        if doi:       parts.append(f"doi:{doi}")
        elif url:     parts.append(url)
    elif ref_type == "book_chapter":
        chapter   = meta.get("bib_chaptertitle") or ""
        book      = meta.get("bib_book", "")
        edition   = meta.get("bib_editionno", "")
        fpage     = meta.get("bib_fpage", "")
        lpage     = meta.get("bib_lpage", "")
        publisher = _strip_publisher_suffixes(meta.get("bib_publisher", ""))
        year      = meta.get("bib_year", "")
        doi       = meta.get("bib_doi", "")
        ed_surnames = [s.strip() for s in (meta.get("bib_ed_surname") or "").split("|") if s.strip()]
        ed_fnames   = [f.strip() for f in (meta.get("bib_ed_fname")   or "").split("|") if f.strip()]
        if chapter: parts.append(f"{_to_sentence_case(chapter)}.")
        editors = []
        for i, s in enumerate(ed_surnames):
            ini = ed_fnames[i] if i < len(ed_fnames) else ""
            initials_fmt = "".join(p[0] for p in ini.split() if p) if ini else ""
            editors.append(f"{s} {initials_fmt}".strip())
        ed_label = "ed." if len(editors) == 1 else "eds."
        in_str = "In: " + ", ".join(editors) + f", {ed_label}. " if editors else "In: "
        book_str = _to_sentence_case(book) if book else ""
        if edition and _ordinal(edition) not in ("1st", "1", "first"):
            book_str += f". {_ordinal(edition)} ed."
        parts.append(in_str + book_str + ".")
        if publisher: parts.append(publisher + ";")
        if year:      parts.append(year + ".")
        pages = f"{fpage}-{lpage}" if fpage and lpage else fpage or lpage
        if pages:     parts[-1] = parts[-1].rstrip(".") + f":{pages}."
        if doi:       parts.append(f"doi:{doi}")
    elif ref_type == "thesis":
        title  = meta.get("bib_title", "")
        deg    = meta.get("bib_deg", "doctoral dissertation")
        school = meta.get("bib_school", "")
        year   = meta.get("bib_year", "")
        url    = meta.get("bib_url", "")
        if title:  parts.append(f"{_to_sentence_case(title)} [{deg}].")
        if school: parts.append(school + ";")
        if year:   parts.append(year + ".")
        if url:    parts.append(url)
    elif ref_type == "conference":
        title    = meta.get("bib_title", "")
        conf     = meta.get("bib_conference", "")
        confloc  = meta.get("bib_conflocation", "")
        confdate = meta.get("bib_confdate", "")
        doi      = meta.get("bib_doi", "")
        if title: parts.append(f"{_to_sentence_case(title)}.")
        conf_str = f"Paper presented at: {conf}" if conf else ""
        if confdate: conf_str += f"; {confdate}"
        if confloc:  conf_str += f"; {confloc}"
        if conf_str: parts.append(conf_str + ".")
        if doi:      parts.append(f"doi:{doi}")
    elif ref_type in ("website", "ereference"):
        title    = meta.get("bib_title", "")
        site     = meta.get("bib_journal") or meta.get("bib_book", "")
        year     = meta.get("bib_year", "")
        accessed = meta.get("bib_accessed", "")
        url      = meta.get("bib_url", "")
        if title:    parts.append(f"{_to_sentence_case(title)}.")
        if site:     parts.append(f"{site}.")
        if year:     parts.append(f"Published {year}.")
        if accessed: parts.append(f"Accessed {accessed}.")
        if url:      parts.append(url)
    elif ref_type == "report":
        title  = meta.get("bib_title", "")
        repnum = meta.get("bib_reportnum", "")
        inst   = _strip_publisher_suffixes(meta.get("bib_institution", ""))
        year   = meta.get("bib_year", "")
        doi    = meta.get("bib_doi", "")
        url    = meta.get("bib_url", "")
        if title:  parts.append(f"{_to_sentence_case(title)}.")
        if inst:   parts.append(inst + ";")
        if year:   parts.append(year + ".")
        if repnum: parts.append(f"Report No. {repnum}.")
        if doi:    parts.append(f"doi:{doi}")
        elif url:  parts.append(url)
    return " ".join(parts)


def _ordinal(n: str) -> str:
    try:
        m = re.search(r'^(\d+)', str(n).strip())
        if m:
            n_int = int(m.group(1))
            suffix = {1:"st",2:"nd",3:"rd"}.get(
                n_int % 10 if n_int % 100 not in (11,12,13) else 0, "th")
            return f"{n_int}{suffix}"
    except Exception:
        pass
    clean = str(n).lower().replace("edition","").replace("ed.","").replace("ed","").strip()
    return clean


# ─────────────────────────────────────────────
# PARAGRAPH FORMATTING HELPERS
# ─────────────────────────────────────────────

def _clear_paragraph_text(para) -> None:
    """
    Remove all run-text from a paragraph.

    Hyperlink XML elements (<w:hyperlink>) are preserved — only their run-text
    is blanked in-place, so the hyperlink relationship (rId) is not orphaned
    and linked text is not accidentally deleted. (#44)
    """
    p_elem = para._p
    # Remove top-level runs
    for r in list(p_elem.findall(qn("w:r"))):
        p_elem.remove(r)
    # For hyperlinks: blank text but keep the element structure
    for hyperlink in p_elem.findall(qn("w:hyperlink")):
        for r in hyperlink.findall(qn("w:r")):
            for t in r.findall(qn("w:t")):
                t.text = ""


def _ensure_style(doc, styles, style_name):
    if style_name and styles is not None:
        try:
            from docx.enum.style import WD_STYLE_TYPE
            if style_name not in styles:
                styles.add_style(style_name, WD_STYLE_TYPE.CHARACTER)
            return styles[style_name]
        except Exception:
            return style_name
    return style_name


# Styles that receive italic formatting in addition to the character style tag.
# bib_volume added for APA rule: volume number is italic in journal references. (#32)
_ITALIC_STYLES = {
    "bib_journal",
    "bib_book",
    "bib_title",
    "bib_volume",   # APA: volume italic  (#32)
}


def _write_styled_runs(para, segments: List[Tuple[str, Optional[str]]], doc=None, original_text: str = None, is_conversion: bool = False) -> None:
    if original_text is None:
        original_text = para.text
    _clear_paragraph_text(para)
    styles = doc.styles if doc is not None else None

    import re
    match = re.match(r'^(\d+)(\.?)([\t\s]*)', original_text)
    prefix_num = ""
    prefix_sep = ""
    if match:
        prefix_num     = match.group(1)
        prefix_sep     = match.group(2) + match.group(3)
        original_text  = original_text[len(prefix_num) + len(prefix_sep):]

    if prefix_num:
        run = para.add_run(prefix_num)
        style_val = _ensure_style(doc, styles, "bib_number")
        try:
            run.style = style_val
        except Exception:
            pass
    if prefix_sep:
        para.add_run(prefix_sep)

    try:
        from utils.track_changes import add_tracked_deletion, add_tracked_text
        use_track_changes = True
    except ImportError:
        use_track_changes = False

    if not use_track_changes:
        for text, style_name in segments:
            if not text:
                continue
            run = para.add_run(text)
            if style_name:
                style_val = _ensure_style(doc, styles, style_name)
                try:
                    run.style = style_val
                except Exception:
                    pass
                if style_name in _ITALIC_STYLES:
                    run.italic = True
        return

    import difflib

    new_full_text = ""
    style_map = []
    for text, style_name in segments:
        if not text: continue
        new_full_text += text
        style_map.extend([style_name] * len(text))

    matcher = difflib.SequenceMatcher(None, original_text, new_full_text)

    for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
        if opcode == 'equal':
            segment_text   = new_full_text[j1:j2]
            segment_styles = style_map[j1:j2]
            chunk_start = 0
            for k in range(len(segment_text) + 1):
                is_end        = (k == len(segment_text))
                style_changed = (k > 0 and k < len(segment_text) and segment_styles[k] != segment_styles[k-1])
                if is_end or style_changed:
                    chunk = segment_text[chunk_start:k]
                    if chunk:
                        style = segment_styles[chunk_start]
                        run = para.add_run(chunk)
                        if style:
                            style_val = _ensure_style(doc, styles, style)
                            try:
                                run.style = style_val
                            except Exception:
                                pass
                            if style in _ITALIC_STYLES:
                                run.italic = True
                    chunk_start = k

        elif opcode == 'delete':
            deleted_chunk = original_text[i1:i2]
            add_tracked_deletion(para, deleted_chunk, doc=doc, author="S4C Reference Converter")

        elif opcode in ('insert', 'replace'):
            if opcode == 'replace':
                deleted_chunk = original_text[i1:i2]
                add_tracked_deletion(para, deleted_chunk, doc=doc, author="S4C Reference Converter")

            segment_text   = new_full_text[j1:j2]
            segment_styles = style_map[j1:j2]
            chunk_start = 0
            for k in range(len(segment_text) + 1):
                is_end        = (k == len(segment_text))
                style_changed = (k > 0 and k < len(segment_text) and segment_styles[k] != segment_styles[k-1])
                if is_end or style_changed:
                    chunk = segment_text[chunk_start:k]
                    if chunk:
                        style = segment_styles[chunk_start]
                        if style:
                            _ensure_style(doc, styles, style)
                        try:
                            add_tracked_text(para, chunk, style=style, author="S4C Reference Converter", doc=doc)
                        except Exception:
                            para.add_run(chunk)
                    chunk_start = k


def _set_paragraph_text(para, text: str, doc=None, original_text: str = None, is_conversion: bool = False) -> None:
    if original_text is None:
        original_text = para.text
    _clear_paragraph_text(para)
    styles = doc.styles if doc is not None else None

    import re
    match = re.match(r'^(\d+\.[\t\s]*)', original_text)
    prefix_text = ""
    if match:
        prefix_text   = match.group(1)
        original_text = original_text[len(prefix_text):]

    if prefix_text:
        run = para.add_run(prefix_text)
        style_val = _ensure_style(doc, styles, "bib_number")
        try:
            run.style = style_val
        except Exception:
            pass

    try:
        from utils.track_changes import add_tracked_deletion, add_tracked_text
        import difflib

        matcher = difflib.SequenceMatcher(None, original_text, text)
        for opcode, i1, i2, j1, j2 in matcher.get_opcodes():
            if opcode == 'equal':
                para.add_run(text[j1:j2])
            elif opcode == 'delete':
                add_tracked_deletion(para, original_text[i1:i2], author="S4C Reference Converter", doc=doc)
            elif opcode in ('insert', 'replace'):
                if opcode == 'replace':
                    add_tracked_deletion(para, original_text[i1:i2], author="S4C Reference Converter", doc=doc)
                add_tracked_text(para, text[j1:j2], author="S4C Reference Converter", doc=doc)
    except ImportError:
        para.add_run(text)


# ─────────────────────────────────────────────
# DB JOURNAL NAME QUALIFIER STRIPPER
# ─────────────────────────────────────────────

_DB_QUALIFIER_PATTERN = re.compile(
    r'\s+\([A-Z][\w\s]+,\s+[A-Z][\w\s]+(:\s*\d{4})?\)'
)

def _strip_db_journal_qualifiers(raw_source: str, metadata: dict, final_text: str) -> tuple:
    journal = (metadata.get("bib_journal") or "").strip()
    if not journal:
        return metadata, final_text
    m = _DB_QUALIFIER_PATTERN.search(journal)
    if m and m.group(0).strip() not in raw_source:
        clean_journal = journal[:m.start()].strip()
        logger.info(f"  [JournalFix] Stripped DB qualifier: '{journal}' → '{clean_journal}'")
        metadata = dict(metadata)
        metadata["bib_journal"] = clean_journal
        bad_suffix = m.group(0)
        final_text = final_text.replace(journal, clean_journal)
        escaped = re.escape(bad_suffix.strip())
        final_text = re.sub(r'\s*' + escaped, '', final_text)
    return metadata, final_text


# ─────────────────────────────────────────────
# GEMINI OUTPUT PARSER
# ─────────────────────────────────────────────

def _parse_gemini_output_to_segments(text: str) -> List[Tuple[str, Optional[str]]]:
    raw_segs: List[Tuple[str, Optional[str]]] = []
    pattern = re.compile(r'\*\*(.+?)\*\*|\*(.+?)\*')
    last = 0
    for m in pattern.finditer(text):
        start, end = m.start(), m.end()
        if start > last:
            raw_segs.append((text[last:start], None))
        if m.group(1) is not None:
            raw_segs.append((m.group(1), "bib_bold"))
        else:
            raw_segs.append((m.group(2), "bib_journal"))
        last = end
    if last < len(text):
        raw_segs.append((text[last:], None))

    PAGE_RANGE = re.compile(
        r'([A-Za-z]?\d+[A-Za-z0-9]*)\s*[\u2013\u2014-]\s*([A-Za-z]?\d+[A-Za-z0-9]*)'
    )
    segs: List[Tuple[str, Optional[str]]] = []
    for seg_text, seg_style in raw_segs:
        if seg_style is not None or not seg_text:
            segs.append((seg_text, seg_style))
            continue
        last_pos = 0
        for pm in PAGE_RANGE.finditer(seg_text):
            before = seg_text[last_pos:pm.start()]
            if before:
                segs.append((before, None))
            fpage = pm.group(1)
            lpage = pm.group(2)
            dash_start = pm.start() + len(fpage)
            dash_end   = pm.end() - len(lpage)
            dash = seg_text[dash_start:dash_end].strip() or '\u2013'
            segs.append((fpage, "bib_fpage"))
            segs.append((dash, None))
            segs.append((lpage, "bib_lpage"))
            last_pos = pm.end()
        remainder = seg_text[last_pos:]
        if remainder:
            segs.append((remainder, None))
    return segs


# ─────────────────────────────────────────────
# SEGMENT BUILDERS
# ─────────────────────────────────────────────

def _is_organization(name: str) -> bool:
    """Heuristic to check if an author name is an organization."""
    if not name: return False
    keywords = {
        "committee", "group", "task force", "section", "association", 
        "society", "department", "national", "center", "institute", 
        "world health", "collaborative", "network", "council", 
        "board", "organization", "agency", "university", "college"
    }
    lower_name = name.lower()
    return any(kw in lower_name for kw in keywords) or len(name.split()) > 3


def _split_pipe(value: Optional[str]) -> List[str]:
    if not value:
        return []
    return [v.strip() for v in str(value).split("|") if v.strip()]


def _format_initials_ama(initial: str) -> str:
    if not initial: return ""
    if any(len(p) > 1 and any(c.islower() for c in p) for p in initial.split()):
        return "".join(p[0].upper() for p in initial.split() if p)
    else:
        return "".join(c.upper() for c in initial if c.isalpha())


# Generation suffixes preserved as-is in APA formatting  (#40)
_NAME_SUFFIXES = frozenset({"jr","sr","ii","iii","iv","2nd","3rd","4th"})

def _format_initials_apa(initial: str) -> str:
    """
    Format first-name/initials to APA style "F. M.", preserving generation
    suffixes such as Jr., Sr., II, III that appear after a comma. (#40)
    """
    if not initial:
        return ""
    # Split on comma to isolate suffix(es)
    comma_parts  = [p.strip() for p in initial.split(",")]
    suffix_parts: list = []
    name_section = comma_parts[0]
    for part in comma_parts[1:]:
        cleaned = part.rstrip(".").lower()
        if cleaned in _NAME_SUFFIXES:
            suffix_parts.append(cleaned.capitalize() + "." if cleaned in {"jr","sr"} else part.strip())
        else:
            name_section += " " + part

    if any(c.islower() for c in name_section):
        clean     = re.sub(r"[^a-zA-Z\s]", " ", name_section)
        formatted = " ".join(w[0].upper() + "." for w in clean.split() if w)
    else:
        letters   = [c.upper() for c in name_section if c.isalpha()]
        formatted = " ".join(c + "." for c in letters)

    if suffix_parts:
        return formatted + ", " + " ".join(suffix_parts)
    return formatted


def build_segments_ama(meta: Dict, gemini_text: str = "") -> List[Tuple[str, Optional[str]]]:
    segs: List[Tuple[str, Optional[str]]] = []
    ref_type = (meta.get("bib_reftype") or "journal").lower()

    # Local heuristic reclassification  (#47, #49)
    if ref_type == "book" and meta.get("bib_chaptertitle") and meta.get("bib_fpage"):
        ref_type = "book_chapter"
        logger.debug("AMA segs: 'book' → 'book_chapter' (chapter+pages)")
    if ref_type == "book" and meta.get("bib_ed_surname") and not meta.get("bib_surname"):
        ref_type = "edited_book"
        logger.debug("AMA segs: 'book' → 'edited_book' (editors, no authors)")

    surnames    = _split_pipe(meta.get("bib_surname"))
    fnames      = _split_pipe(meta.get("bib_fname"))
    n_auth      = len(surnames)
    ed_surnames = _split_pipe(meta.get("bib_ed_surname") or meta.get("bib_ed-surname"))
    ed_fnames   = _split_pipe(meta.get("bib_ed_fname")   or meta.get("bib_ed-fname"))

    if n_auth == 0:
        # Check bib_surname for org-as-author (bib_fname blank)  (#33)
        org = (meta.get("bib_organization") or
               meta.get("bib_institution") or
               (meta.get("bib_surname") if not meta.get("bib_fname") else "") or
               "")
        if org:
            segs.append((org.rstrip("."), "bib_organization"))
            segs.append((".", None))
        elif ed_surnames and ref_type != "book_chapter":
            for i, es in enumerate(ed_surnames):
                if i > 0: segs.append((", ", None))
                segs.append((es, "bib_ed-surname"))
                ei     = ed_fnames[i] if i < len(ed_fnames) else ""
                ei_str = _format_initials_ama(ei)
                if ei_str:
                    segs.append((" ", None))
                    segs.append((ei_str, "bib_ed-fname"))
            ed_label = "ed." if len(ed_surnames) == 1 else "eds."
            segs.append((f", {ed_label}", None))
    else:
        subset = surnames if n_auth <= 6 else surnames[:6]
        for i, surname in enumerate(subset):
            if i > 0: segs.append((", ", None))
            initial     = fnames[i] if i < len(fnames) else ""
            initials_str = _format_initials_ama(initial)
            if not initials_str and _is_organization(surname):
                disp_name = surname[0].upper() + surname[1:] if i == 0 and surname else surname
                segs.append((disp_name, "bib_organization"))
            else:
                disp_name = surname[0].upper() + surname[1:] if i == 0 and surname else surname
                segs.append((disp_name, "bib_surname"))
                if initials_str:
                    segs.append((" ", None))
                    segs.append((initials_str, "bib_fname"))
        if n_auth > 6:
            segs.append((", ", None))
            segs.append(("et al", "bib_etal"))
    segs.append((". ", None))

    chapter_title = meta.get("bib_chaptertitle") or ""
    main_title    = meta.get("bib_title") or ""
    book_title    = meta.get("bib_book") or ""

    if ref_type == "book_chapter" and chapter_title:
        clean_title = _to_sentence_case(chapter_title.rstrip("."))
        segs.append((clean_title, "bib_chaptertitle"))
        segs.append((" " if re.search(r'[?!]$', clean_title) else ". ", None))
    elif main_title:
        clean_title = _to_sentence_case(main_title.rstrip("."))
        segs.append((clean_title, "bib_article" if ref_type == "journal" else "bib_title"))
        segs.append((" " if re.search(r'[?!]$', clean_title) else ". ", None))

    if ref_type == "book_chapter":
        segs.append(("In: ", None))
        if ed_surnames:
            for i, es in enumerate(ed_surnames):
                if i > 0: segs.append((", ", None))
                segs.append((es, "bib_ed-surname"))
                ei     = ed_fnames[i] if i < len(ed_fnames) else ""
                ei_str = _format_initials_ama(ei)
                if ei_str:
                    segs.append((" ", None))
                    segs.append((ei_str, "bib_ed-fname"))
            ed_label = "ed." if len(ed_surnames) == 1 else "eds."
            segs.append((", " + ed_label + " ", None))
        if book_title:
            segs.append((_to_sentence_case(book_title), "bib_book"))
            segs.append((". ", None))

    if ref_type == "journal":
        journal = meta.get("bib_journal") or ""
        year    = meta.get("bib_year") or ""
        volume  = meta.get("bib_volume") or ""
        issue   = meta.get("bib_issue") or ""
        fpage   = meta.get("bib_fpage") or ""
        lpage   = meta.get("bib_lpage") or ""
        if journal:
            segs.append((journal, "bib_journal"))
            segs.append((".", None))
        if year:
            segs.append((" ", None))
            segs.append((year, "bib_year"))
        if volume:
            segs.append((";", None))
            segs.append((volume, "bib_volume"))
        if issue:
            segs.append(("(", None))
            segs.append((issue, "bib_issue"))
            segs.append((")", None))
        if fpage:
            segs.append((":", None))
            segs.append((fpage, "bib_fpage"))
            if lpage:
                segs.append(("-", None))
                segs.append((lpage, "bib_lpage"))
        elif not volume and not issue and "Published online" in gemini_text:
            segs.append((". Published online", None))
        segs.append((".", None))

    elif ref_type in ("book", "edited_book", "book_chapter"):
        edition   = meta.get("bib_editionno") or ""
        publisher = _strip_publisher_suffixes(meta.get("bib_publisher") or "")
        year      = meta.get("bib_year") or ""
        if ref_type != "book_chapter" and book_title:
            segs.append((_to_sentence_case(book_title), "bib_book"))
            segs.append((". ", None))
        if edition and _ordinal(edition) not in ("1st", "1"):
            segs.append((_ordinal(edition) + " ed. ", "bib_editionno"))
        if publisher:
            segs.append((publisher, "bib_publisher"))
            segs.append(("; ", None))
        if year:
            segs.append((year, "bib_year"))
        if ref_type == "book_chapter":
            fpage = meta.get("bib_fpage") or ""
            lpage = meta.get("bib_lpage") or ""
            if fpage:
                segs.append((":", None))
                segs.append((fpage, "bib_fpage"))
                if lpage:
                    segs.append(("-", None))
                    segs.append((lpage, "bib_lpage"))
        segs.append((".", None))

    elif ref_type == "conference":
        conf     = meta.get("bib_conference") or ""
        confloc  = meta.get("bib_conflocation") or ""
        confdate = meta.get("bib_confdate") or meta.get("bib_year") or ""
        if title := meta.get("bib_title") or "":
            clean_title = _to_sentence_case(title.rstrip("."))
            segs.append((clean_title, "bib_confpaper"))
            segs.append((" " if re.search(r'[?!]$', clean_title) else ". ", None))
        segs.append(("Paper presented at: ", None))
        if conf:
            segs.append((conf, "bib_conference"))
        if confdate:
            segs.append(("; ", None))
            segs.append((confdate, "bib_confdate"))
        if confloc:
            segs.append(("; ", None))
            segs.append((confloc, "bib_conflocation"))
        segs.append((".", None))

    elif ref_type == "thesis":
        title  = meta.get("bib_title") or ""
        deg    = meta.get("bib_deg") or "doctoral dissertation"
        school = meta.get("bib_school") or ""
        year   = meta.get("bib_year") or ""
        url    = meta.get("bib_url") or ""
        if title:
            segs.append((_to_sentence_case(title.rstrip(".")), "bib_title"))
        bracket = f" [{deg}]."
        segs.append((bracket, None))
        if school:
            segs.append((" " + school + ";", None))
        if year:
            segs.append((" " + year + ".", None))
        if url:
            segs.append((" ", None))
            segs.append((url, "bib_url"))
        doi = (meta.get("bib_doi") or "").strip().lstrip("doi:").lstrip()
        if doi:
            segs.append((" doi:", "bib_doi"))
            segs.append((doi, "bib_doi"))
        return segs

    elif ref_type in ("website", "ereference"):
        title    = meta.get("bib_title") or ""
        year     = meta.get("bib_year") or ""
        accessed = meta.get("bib_accessed") or ""
        url      = meta.get("bib_url") or ""
        site     = meta.get("bib_journal") or meta.get("bib_book") or ""
        pub      = _strip_publisher_suffixes(meta.get("bib_publisher") or "")
        if title:
            clean_title = _to_sentence_case(title.rstrip("."))
            segs.append((clean_title, "bib_title"))
            segs.append((" " if re.search(r'[?!]$', clean_title) else ". ", None))
        if site:
            segs.append((site, "bib_journal"))
            segs.append((". ", None))
        if pub and pub.lower() != site.lower():
            segs.append((pub, "bib_publisher"))
            segs.append((". ", None))
        if year:
            segs.append(("Published ", None))
            segs.append((year, "bib_year"))
            segs.append((". ", None))
        if accessed:
            segs.append(("Accessed ", None))
            segs.append((accessed, "bib_accessed"))
            segs.append((". ", None))
        if url:
            segs.append((url, "bib_url"))

    elif ref_type == "report":
        repnum = meta.get("bib_reportnum") or ""
        inst   = _strip_publisher_suffixes(meta.get("bib_institution") or "")
        year   = meta.get("bib_year") or ""
        if inst:
            segs.append((inst, "bib_publisher"))
            segs.append(("; ", None))
        if year:
            segs.append((year, "bib_year"))
            segs.append((".", None))
        if repnum:
            segs.append((" Report No. " + repnum + ".", None))

    doi = (meta.get("bib_doi") or "").strip().lstrip("doi:").lstrip()
    if doi and ref_type not in ("website", "ereference", "thesis"):
        segs.append((" doi:", "bib_doi"))
        segs.append((doi, "bib_doi"))

    return segs


def build_segments_apa(meta: Dict, gemini_text: str = "") -> List[Tuple[str, Optional[str]]]:
    segs: List[Tuple[str, Optional[str]]] = []
    ref_type = (meta.get("bib_reftype") or "journal").lower()

    surnames = _split_pipe(meta.get("bib_surname"))
    fnames   = _split_pipe(meta.get("bib_fname"))
    n_auth   = len(surnames)

    is_edited_book_primary = False
    if ref_type == "edited_book" and n_auth == 0:
        surnames = _split_pipe(meta.get("bib_ed_surname") or meta.get("bib_ed-surname"))
        fnames   = _split_pipe(meta.get("bib_ed_fname")   or meta.get("bib_ed-fname"))
        n_auth   = len(surnames)
        is_edited_book_primary = True

    if n_auth == 0:
        # Org-as-author: Gemini sometimes stores org name in bib_surname with blank bib_fname  (#33)
        org = (meta.get("bib_organization") or
               meta.get("bib_institution") or
               (meta.get("bib_surname") if not meta.get("bib_fname") else "") or
               "")
        if org:
            segs.append((org.rstrip("."), "bib_organization"))
            segs.append((".", None))
    else:
        subset = surnames if n_auth <= 20 else surnames[:19]
        for i, surname in enumerate(subset):
            if i > 0:
                segs.append((", ", None))
                if i == n_auth - 1 and n_auth <= 20:
                    segs.append(("& ", None))
            initial      = fnames[i] if i < len(fnames) else ""
            initials_str = _format_initials_apa(initial)  # suffix-aware  (#40)
            if not initials_str and _is_organization(surname):
                disp_name = surname[0].upper() + surname[1:] if i == 0 and surname else surname
                segs.append((disp_name, "bib_organization"))
            else:
                disp_name = surname[0].upper() + surname[1:] if i == 0 and surname else surname
                segs.append((disp_name, "bib_surname"))
                if initials_str:
                    segs.append((", ", None))
                    segs.append((initials_str, "bib_fname"))
        if n_auth > 20:
            segs.append((", … ", None))
            segs.append((surnames[-1], "bib_surname"))
            last_initial = fnames[-1] if len(fnames) >= n_auth else ""
            initials_str = _format_initials_apa(last_initial)
            if initials_str:
                segs.append((", ", None))
                segs.append((initials_str, "bib_fname"))

    if is_edited_book_primary and n_auth > 0:
        ed_label = " (Ed.)" if n_auth == 1 else " (Eds.)"
        segs.append((ed_label, None))

    # Add a period after the author block if it doesn't already end with one 
    # (e.g., if the last element was an organization without initials)
    if n_auth > 0 and not is_edited_book_primary:
        last_seg_text = segs[-1][0] if segs else ""
        if not last_seg_text.endswith("."):
            segs.append((".", None))

    segs.append((" (", None))
    segs.append((meta.get("bib_year") or "n.d.", "bib_year"))
    segs.append(("). ", None))

    chapter_title = meta.get("bib_chaptertitle") or ""
    main_title    = meta.get("bib_title") or ""
    book_title    = meta.get("bib_book") or ""

    # ── Title block ───────────────────────────────────────────────
    if ref_type == "thesis":
        # Dedicated thesis block  (#57)
        title  = main_title or book_title or ""
        deg    = meta.get("bib_deg") or "Doctoral dissertation"
        school = meta.get("bib_school") or ""
        url    = meta.get("bib_url") or ""
        if title:
            segs.append((_to_sentence_case(title.rstrip(".")), "bib_title"))
        bracket = f" [{deg}"
        if school:
            bracket += f", {school}"
        bracket += "]."
        segs.append((bracket, None))
        if url:
            segs.append((" ", None))
            segs.append((url, "bib_url"))
        doi = (meta.get("bib_doi") or "").strip().lstrip("doi:").lstrip()
        if doi:
            segs.append((" https://doi.org/", "bib_doi"))
            segs.append((doi, "bib_doi"))
        return segs

    elif ref_type == "book_chapter" and chapter_title:
        clean_title = _to_sentence_case(chapter_title.rstrip("."))
        segs.append((clean_title, "bib_chaptertitle"))
        segs.append((" " if re.search(r'[?!]$', clean_title) else ". ", None))

    elif ref_type in ("book", "edited_book"):
        # Always emit book title, including 1st-edition books  (#37)
        display_title = book_title or main_title or ""
        if display_title:
            clean_title = _to_sentence_case(display_title.rstrip("."))
            segs.append((clean_title, "bib_book"))
            segs.append((".", None))

    elif main_title:
        clean_title = _to_sentence_case(main_title.rstrip("."))
        style = "bib_article" if ref_type == "journal" else "bib_title"
        segs.append((clean_title, style))
        segs.append((" " if re.search(r'[?!]$', clean_title) else ". ", None))

    # ── In: editors (book chapter) ────────────────────────────────
    if ref_type == "book_chapter":
        ed_surnames = _split_pipe(meta.get("bib_ed_surname") or meta.get("bib_ed-surname"))
        ed_fnames   = _split_pipe(meta.get("bib_ed_fname")   or meta.get("bib_ed-fname"))
        segs.append(("In ", None))
        if ed_surnames:
            for i, es in enumerate(ed_surnames):
                if i > 0:
                    segs.append((" & ", None) if i == len(ed_surnames) - 1 else (", ", None))
                ei           = ed_fnames[i] if i < len(ed_fnames) else ""
                initials_str = _format_initials_apa(ei)
                if initials_str:
                    segs.append((initials_str + " ", "bib_ed-fname"))
                if not initials_str and _is_organization(es):
                    segs.append((es, "bib_organization"))
                else:
                    segs.append((es, "bib_ed-surname"))
            ed_label = "(Ed.)," if len(ed_surnames) == 1 else "(Eds.),"
            segs.append((" " + ed_label + " ", None))

        # Always emit book title  (#37)
        display_book = book_title or main_title or ""
        if display_book:
            segs.append((_to_sentence_case(display_book.rstrip(".")), "bib_book"))

        edition   = meta.get("bib_editionno") or ""
        volume    = meta.get("bib_volume") or ""   # Vol. support  (#55)
        fpage     = meta.get("bib_fpage") or ""
        lpage     = meta.get("bib_lpage") or ""
        clean_ord = _ordinal(edition)

        inner_segs: list = []
        if edition and clean_ord not in ("1st", "1", "first"):
            inner_segs.append((clean_ord + " ed.", "bib_editionno"))
        if volume:
            if inner_segs: inner_segs.append((", ", None))
            inner_segs.append(("Vol. ", None))
            inner_segs.append((volume, "bib_volume"))
        if fpage:
            if inner_segs: inner_segs.append((", ", None))
            inner_segs.append(("pp. ", None))
            inner_segs.append((fpage, "bib_fpage"))
            if lpage:
                inner_segs.append(("–", None))
                inner_segs.append((lpage, "bib_lpage"))

        if inner_segs:
            segs.append((" (", None))
            segs.extend(inner_segs)
            segs.append((").", None))
        else:
            segs.append((".", None))

        publisher = _strip_publisher_suffixes(meta.get("bib_publisher") or "")  # (#58)
        if publisher:
            segs.append((" ", None))
            segs.append((publisher, "bib_publisher"))
            segs.append((".", None))

    # ── Journal section ───────────────────────────────────────────
    elif ref_type == "journal":
        journal = _to_title_case(meta.get("bib_journal") or "")  # title case  (#31)
        volume  = meta.get("bib_volume") or ""
        issue   = meta.get("bib_issue") or ""
        fpage   = meta.get("bib_fpage") or ""
        lpage   = meta.get("bib_lpage") or ""
        if journal:
            segs.append((journal, "bib_journal"))  # italic via _ITALIC_STYLES
        if volume:
            segs.append((", ", None))
            segs.append((volume, "bib_volume"))    # italic via _ITALIC_STYLES  (#32)
        if issue:
            segs.append(("(", None))
            segs.append((issue, "bib_issue"))
            segs.append((")", None))
        if fpage:
            segs.append((", ", None))
            segs.append((fpage, "bib_fpage"))
            if lpage:
                segs.append(("–", None))
                segs.append((lpage, "bib_lpage"))
        elif not volume and not issue and "Advance online publication" in gemini_text:
            segs.append((". Advance online publication", None))
        segs.append((".", None))

    elif ref_type in ("book", "edited_book"):
        edition   = meta.get("bib_editionno") or ""
        publisher = _strip_publisher_suffixes(meta.get("bib_publisher") or "")  # (#58)
        # Split edition so ordinal gets bib_editionno style  (#36)
        if edition and _ordinal(edition) not in ("1st", "1", "first"):
            segs.append((" (", None))
            segs.append((_ordinal(edition) + " ed.", "bib_editionno"))
            segs.append((").", None))
        if publisher:
            segs.append((" ", None))
            segs.append((publisher, "bib_publisher"))
            segs.append((".", None))

    elif ref_type in ("website", "ereference"):
        site     = meta.get("bib_journal") or meta.get("bib_book") or ""
        pub      = _strip_publisher_suffixes(meta.get("bib_publisher") or "")
        accessed = meta.get("bib_accessed") or ""
        url      = meta.get("bib_url") or ""
        if site:
            segs.append((_to_title_case(site), "bib_journal"))
            segs.append((".", None))
        if pub and pub.lower() != site.lower():
            segs.append((" ", None))
            segs.append((pub, "bib_publisher"))
            segs.append((".", None))
        if accessed:
            segs.append((" Retrieved " + accessed + ", from ", None))
        if url:
            segs.append((url, "bib_url"))

    elif ref_type == "conference":
        conf     = meta.get("bib_conference") or ""
        confloc  = meta.get("bib_conflocation") or ""
        confdate = meta.get("bib_confdate") or ""
        segs.append(("[Conference session]. ", None))
        if conf:
            segs.append((conf, "bib_conference"))
        if confdate:
            segs.append((", " + confdate, None))
        if confloc:
            segs.append((", " + confloc, None))
        segs.append((".", None))

    elif ref_type == "report":
        repnum = meta.get("bib_reportnum") or ""
        inst   = _strip_publisher_suffixes(meta.get("bib_institution") or "")
        if repnum:
            segs.append((" (Report No. " + repnum + ").", None))
        if inst:
            segs.append((" ", None))
            segs.append((inst, "bib_publisher"))
            segs.append((".", None))

    # ── DOI / URL ─────────────────────────────────────────────────
    doi = (meta.get("bib_doi") or "").strip().lstrip("doi:").lstrip()
    url = meta.get("bib_url") or ""
    if doi:
        segs.append((" https://doi.org/", "bib_doi"))
        segs.append((doi, "bib_doi"))
    elif url and ref_type not in ("website", "ereference", "thesis"):
        segs.append((" ", None))
        segs.append((url, "bib_url"))

    return segs


# ─────────────────────────────────────────────
# CONVERSION LOG ENTRY
# ─────────────────────────────────────────────

class ConversionLogEntry:
    def __init__(self, original: str, converted: str, ref_type: str,
                 source_style: str, target_style: str, notes: Optional[str] = None,
                 error: Optional[str] = None):
        self.original     = original
        self.converted    = converted
        self.ref_type     = ref_type
        self.source_style = source_style
        self.target_style = target_style
        self.notes        = notes
        self.error        = error

    def to_log_line(self) -> str:
        lines = [
            f"  TYPE:    {self.ref_type}",
            f"  FROM:    [{self.source_style}] {self.original}",
            f"  TO:      [{self.target_style}] {self.converted}",
        ]
        if self.notes:  lines.append(f"  NOTES:   {self.notes}")
        if self.error:  lines.append(f"  ERROR:   {self.error}")
        return "\n".join(lines)


# ─────────────────────────────────────────────
# MAIN PROCESSOR
# ─────────────────────────────────────────────

def process_conversion(
    input_docx: Path,
    output_dir: Optional[Path] = None,
    source_style: str = "Auto",
    target_style: str = "APA",
    model_name: str = "gemini-2.0-flash",
    prefer_gemini_output: bool = True,
) -> Dict[str, Path]:
    input_docx = Path(input_docx)
    if not input_docx.exists():
        raise FileNotFoundError(f"Input file not found: {input_docx}")

    target_style = target_style.strip().upper() if target_style.upper() != "AUTO" else "AUTO"
    if target_style not in ("AMA", "APA", "AUTO"):
        raise ValueError(f"target_style must be 'AMA', 'APA', or 'AUTO', got: {target_style}")

    if output_dir is None:
        output_dir = input_docx.parent
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    target_enum = CitationStyle.APA if target_style == "APA" else CitationStyle.AMA

    stem             = input_docx.stem
    output_docx_path = output_dir / f"{stem}_Converted.docx"
    log_file_path    = output_dir / f"{stem}_conversion_log.txt"
    json_dump_path   = output_dir / f"{stem}_metadata_dump.json"

    doc = Document(input_docx)

    log_entries: List[ConversionLogEntry] = []
    json_records: List[Dict] = []
    log_header: List[str] = [
        f"Reference Conversion Log",
        f"Input:         {input_docx.name}",
        f"Source Style:  {source_style}",
        f"Target Style:  {target_style}",
        f"Model:         {model_name}",
        "=" * 60, ""
    ]

    total_count     = 0
    converted_count = 0
    error_count     = 0
    in_ref_section  = False

    from concurrent.futures import ThreadPoolExecutor, as_completed
    from ReferencesStructing import find_best_metadata_for_reference, detect_reference_style

    tasks = []

    for idx, para in enumerate(doc.paragraphs):
        raw_text = para.text.strip()
        if not raw_text:
            continue

        raw_lower = raw_text.lower()

        if "<ref-open>" in raw_lower:
            in_ref_section = True
            logger.info("Entering reference section.")
            continue
        if "<ref-close>" in raw_lower:
            in_ref_section = False
            logger.info("Exiting reference section.")
            continue
        if not in_ref_section:
            continue

        if len(raw_text) < 15:
            continue

        # Skip standalone in-text citations  (#41)
        if _looks_like_inline_citation(raw_text):
            logger.debug(f"Skipping inline citation para: {raw_text[:60]}")
            continue

        total_count += 1
        try:
            para_style_name = para.style.name or ''
        except Exception:
            para_style_name = ''
        tasks.append({
            'doc_index':  idx,
            'para_obj':   para,
            'raw_text':   raw_text,
            'count':      total_count,
            'para_style': para_style_name,
        })

    def process_task(task: dict):
        raw_text = task['raw_text']
        count    = task['count']
        logger.info(f"[{count}] Conversion API Call: {raw_text[:80]}...")

        if source_style.upper() == "AUTO":
            para_style = task.get('para_style', '')
            if para_style == 'REF-N':
                detected_source = CitationStyle.AMA
            elif para_style in ('REF-U', 'REF'):
                detected_source = CitationStyle.APA
            else:
                detected_source = detect_source_style(raw_text)
        else:
            detected_source = CitationStyle.AMA if source_style.upper() == "AMA" else CitationStyle.APA

        task['detected_source'] = detected_source

        if target_style.upper() == "AUTO":
            t_enum = detected_source
            logger.info(f"  [{count}] Auto: strict formatting validation for {t_enum.value}")
        else:
            t_enum = CitationStyle.APA if target_style.upper() == "APA" else CitationStyle.AMA

        if detected_source == t_enum:
            logger.info(f"  [{count}] [Formatting Validation] Already in {t_enum.value}")

        cr_item = None
        try:
            temp_cr, source_db, score = find_best_metadata_for_reference(raw_text, detected_source.value)
            is_journal = False
            if temp_cr:
                if 'pubmed' in source_db.lower():
                    is_journal = True
                elif 'crossref' in source_db.lower() and temp_cr.get('type', '').lower() in ('journal-article', 'journal'):
                    is_journal = True
                elif 'crossref' in source_db.lower() and not temp_cr.get('type') and temp_cr.get('container-title'):
                    is_journal = True

            # Lowered PubMed threshold to 0.65 for more reliable matching  (#42)
            if is_journal and score >= 0.65:
                cr_item = temp_cr
                logger.info(f"  [{count}] [DB Match] Journal via {source_db} (Score: {score:.2f})")
            elif temp_cr and score >= 0.75:
                cr_item = temp_cr
                logger.info(f"  [{count}] [DB Match] General via {source_db} (Score: {score:.2f})")
            elif temp_cr:
                logger.info(f"  [{count}] [DB Match] Ignored {source_db} (Score: {score:.2f}) — below threshold")
                cr_item = None
        except Exception as e:
            logger.warning(f"  [{count}] Failed to query CrossRef/PubMed: {e}")

        result = convert_reference(
            raw_text=raw_text,
            source_style=detected_source,
            target_style=t_enum,
            model_name=model_name,
            cr_item=cr_item,
        )
        task['target_enum'] = t_enum
        task['result']      = result
        task['cr_item']     = cr_item
        task['skip']        = False
        return task

    if tasks:
        logger.info(f"Starting parallel conversions for {len(tasks)} references...")
        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = [executor.submit(process_task, t) for t in tasks]
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    logger.error(f"Error in parallel conversion task: {e}")

    for task in sorted(tasks, key=lambda x: x['doc_index']):
        count           = task['count']
        raw_text        = task['raw_text']
        para            = task['para_obj']
        result          = task['result']
        detected_source = task['detected_source']

        if task.get('skip'):
            logger.info(f"  [{count}] Skipping reference: kept original formatting.")
            continue

        if not result:
            error_count += 1
            entry = ConversionLogEntry(
                original=raw_text, converted="[FAILED]",
                ref_type="unknown", source_style=detected_source.value,
                target_style=target_style, error="Gemini returned no result",
            )
            log_entries.append(entry)
            logger.warning(f"  Gemini failed for reference {count}")
            continue

        metadata   = result.get("metadata", {})
        ref_type   = detect_ref_type_from_metadata(metadata)
        gemini_out = result.get("formatted_output", "").strip()
        notes      = result.get("conversion_notes")

        # Heuristic ref-type correction  (#46, #49)
        metadata = _fix_ref_type(metadata, raw_text)
        ref_type = detect_ref_type_from_metadata(metadata)

        resolved_target = task['target_enum'].value

        cr_it = task.get('cr_item')
        if cr_it:
            # DOI: DB is always authoritative — always overwrite  (#42)
            if cr_it.get("DOI"):
                db_doi = str(cr_it["DOI"]).replace("https://doi.org/","").replace("doi:","").strip()
                if db_doi:
                    metadata["bib_doi"] = db_doi

            # Fill missing title from DB  (#42)
            if cr_it.get("title") and not metadata.get("bib_title"):
                raw_t = cr_it["title"]
                metadata["bib_title"] = raw_t[0] if isinstance(raw_t, list) else str(raw_t)

            if cr_it.get("URL") and not metadata.get("bib_url"):
                metadata["bib_url"] = str(cr_it["URL"]).strip()
            if cr_it.get("volume") and not metadata.get("bib_volume"):
                metadata["bib_volume"] = str(cr_it["volume"]).strip()
            if cr_it.get("issue") and not metadata.get("bib_issue"):
                metadata["bib_issue"] = str(cr_it["issue"]).strip()
            if cr_it.get("page") and not metadata.get("bib_fpage"):
                raw_page = str(cr_it["page"]).strip()
                if "-" in raw_page:
                    parts = raw_page.split("-", 1)
                    metadata["bib_fpage"] = parts[0].strip()
                    if not metadata.get("bib_lpage"):
                        metadata["bib_lpage"] = parts[1].strip()
                else:
                    metadata["bib_fpage"] = raw_page

            db_year = None
            for date_key in ("published-print", "published-online", "issued"):
                dp = cr_it.get(date_key, {}).get("date-parts")
                if dp and dp[0] and dp[0][0]:
                    db_year = str(dp[0][0])
                    break
            if not db_year and cr_it.get("year"):
                db_year = str(cr_it["year"])
            if db_year:
                if not metadata.get("bib_year"):
                    metadata["bib_year"] = db_year
                elif metadata.get("bib_year","").rstrip("abcdefghijklmnopqrstuvwxyz") != db_year:
                    suffix = metadata["bib_year"][len(metadata["bib_year"].rstrip("abcdefghijklmnopqrstuvwxyz")):]
                    metadata["bib_year"] = db_year + suffix
                    logger.info(f"  [{count}] [DB Correction] Year → {db_year}{suffix}")

            if not metadata.get("bib_journal"):
                if resolved_target == "AMA":
                    abbr = (cr_it.get("short-container-title") or [""])[0].strip()
                    full = (cr_it.get("container-title") or [""])[0].strip()
                    metadata["bib_journal"] = abbr or full
                else:
                    full = (cr_it.get("container-title") or [""])[0].strip()
                    metadata["bib_journal"] = full

        if prefer_gemini_output and gemini_out:
            final_text = gemini_out
            if metadata.get("bib_doi") and "doi:" not in final_text.lower() and "doi.org" not in final_text.lower():
                if resolved_target == "AMA":
                    final_text = final_text.rstrip(".") + f". doi:{metadata['bib_doi']}"
                else:
                    final_text = final_text.rstrip(".") + f". https://doi.org/{metadata['bib_doi']}"
        else:
            if resolved_target == "AMA":
                final_text = format_ama_from_metadata(metadata)
            else:
                final_text = format_apa_from_metadata(metadata)

        metadata, final_text = _strip_db_journal_qualifiers(raw_text, metadata, final_text)

        # Strip curly quotes — Word manages smart quotes; pre-curled quotes create TC noise  (#35)
        final_text = _normalise_quotes(final_text)

        if not final_text.strip():
            error_count += 1
            entry = ConversionLogEntry(
                original=raw_text, converted="[EMPTY OUTPUT]",
                ref_type=ref_type, source_style=detected_source.value,
                target_style=target_style,
                error="Both Gemini output and metadata fallback produced empty string",
            )
            log_entries.append(entry)
            continue

        try:
            segs = []
            if metadata and metadata.get("bib_reftype"):
                try:
                    if resolved_target == "AMA":
                        segs = build_segments_ama(metadata, gemini_out)
                    else:
                        segs = build_segments_apa(metadata, gemini_out)
                    segs_text = "".join(t for t, _ in segs)
                    if len(segs_text.strip()) < 10:
                        segs = []
                        logger.debug(f"  [{count}] Metadata segments too short; using Gemini text path.")
                except Exception as _meta_err:
                    segs = []
                    logger.warning(f"  [{count}] Metadata segment build failed ({_meta_err}); falling back.")

            if not segs and final_text:
                segs = _parse_gemini_output_to_segments(final_text)
                logger.debug(f"  [{count}] Using Gemini text parse (fallback) for styling.")

            if segs:
                _write_styled_runs(para, segs, doc=doc, is_conversion=(detected_source != task['target_enum']))
            else:
                _set_paragraph_text(para, final_text, doc=doc)
        except Exception as _seg_err:
            logger.warning(f"  Segment build failed ({_seg_err}); falling back to plain text.")
            _set_paragraph_text(para, final_text, doc=doc)

        converted_count += 1

        entry = ConversionLogEntry(
            original=raw_text, converted=final_text,
            ref_type=ref_type, source_style=detected_source.value,
            target_style=target_style, notes=notes,
        )
        log_entries.append(entry)
        json_records.append({
            "index": count, "ref_type": ref_type,
            "source_style": detected_source.value, "target_style": target_style,
            "original": raw_text, "converted": final_text,
            "notes": notes, "metadata": metadata,
        })
        logger.info(f"  ✓ [{ref_type}] → {final_text[:80]}...")

    doc.save(output_docx_path)
    logger.info(f"Saved converted document: {output_docx_path}")

    summary = [
        "", "=" * 60,
        f"SUMMARY",
        f"  Total references found:  {total_count}",
        f"  Successfully converted:  {converted_count}",
        f"  Errors:                  {error_count}",
        f"  Skipped (same style):    {total_count - converted_count - error_count}",
    ]

    with open(log_file_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log_header) + "\n")
        for i, entry in enumerate(log_entries, 1):
            f.write(f"[{i}]\n{entry.to_log_line()}\n\n")
        f.write("\n".join(summary) + "\n")

    logger.info(f"Log written: {log_file_path}")

    with open(json_dump_path, "w", encoding="utf-8") as f:
        json.dump(json_records, f, indent=2, ensure_ascii=False)

    logger.info(f"Metadata dump: {json_dump_path}")

    return {
        "output_docx": output_docx_path,
        "log_file":    log_file_path,
        "json_dump":   json_dump_path,
    }


# ─────────────────────────────────────────────
# CLI ENTRY POINT
# ─────────────────────────────────────────────

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Convert references in a Word document between AMA and APA styles.")
    parser.add_argument("input",              type=str,                                          help="Path to input .docx file")
    parser.add_argument("--output-dir",       type=str,                                          help="Output directory (default: same as input)")
    parser.add_argument("--source-style",     type=str, default="Auto", choices=["AMA","APA","Auto"], help="Source citation style")
    parser.add_argument("--target-style",     type=str, default="APA",  choices=["AMA","APA"],       help="Target citation style")
    parser.add_argument("--model",            type=str, default="gemini-2.0-flash",               help="Gemini model name")
    parser.add_argument("--no-gemini-output", action="store_true",                                help="Rebuild from metadata instead of Gemini formatted output")
    args = parser.parse_args()

    paths = process_conversion(
        input_docx=Path(args.input),
        output_dir=Path(args.output_dir) if args.output_dir else None,
        source_style=args.source_style,
        target_style=args.target_style,
        model_name=args.model,
        prefer_gemini_output=not args.no_gemini_output,
    )

    print("\nConversion complete:")
    for k, v in paths.items():
        print(f"  {k}: {v}")
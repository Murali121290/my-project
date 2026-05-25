"""
gemini_ref_converter.py
-----------------------
Converts and validates bibliographic references between AMA 11th and APA 7th
edition using the Google Gemini API.
"""

from __future__ import annotations

import collections
import json
import logging
import os
import re
import threading
import time
import functools
import requests
from enum import Enum
from typing import Any, Dict, List, Optional, Tuple

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# ---------------------------------------------------------------------------
# Logging – use NullHandler so we don't force output on library consumers.
# ---------------------------------------------------------------------------
logger = logging.getLogger(__name__)
logger.addHandler(logging.NullHandler())

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
DEFAULT_MODEL = os.environ.get("REFERENCE_CONVERTER_GEMINI_MODEL", "gemini-2.5-flash")
_MAX_RETRIES = int(os.environ.get("REFERENCE_CONVERTER_MAX_RETRIES", "5"))
_RETRY_BASE_DELAY = float(os.environ.get("REFERENCE_CONVERTER_RETRY_BASE_DELAY", "5.0"))

# Global sliding-window rate limiter — shared across all worker threads.
# Keeps calls under 12 RPM to stay safely within Gemini free-tier quota (15 RPM).
_rl_lock = threading.Lock()
_rl_timestamps: collections.deque = collections.deque()
_RL_MAX_CALLS = int(os.environ.get("REFERENCE_CONVERTER_RPM", "12"))
_RL_WINDOW = 60.0


def _acquire_api_slot() -> None:
    """Block until a Gemini API call slot is available (12 RPM sliding window)."""
    while True:
        with _rl_lock:
            now = time.time()
            while _rl_timestamps and now - _rl_timestamps[0] > _RL_WINDOW:
                _rl_timestamps.popleft()
            if len(_rl_timestamps) < _RL_MAX_CALLS:
                _rl_timestamps.append(now)
                return
            wait = _RL_WINDOW - (now - _rl_timestamps[0]) + 0.05
        time.sleep(wait)


# ─────────────────────────────────────────────
# CrossRef DOI LOOKUP
# ─────────────────────────────────────────────

def _extract_authors_from_pubmed_article(article: "Dict[str, Any]") -> Optional[List[str]]:
    """Extract pipe-separated author list from PubMed article data."""
    try:
        authors = []
        # PubMed article structure typically has authlist
        if "authorlist" in article:
            for author in article["authorlist"]:
                if isinstance(author, dict):
                    lastname = author.get("lastname", "")
                    initials = author.get("initials", "")
                    if lastname:
                        authors.append(f"{lastname} {initials}".strip())
        return authors if authors else None
    except Exception as e:
        logger.debug(f"  [PubMed] Author extraction error: {e}")
        return None


def _extract_authors_from_crossref_item(item: "Dict[str, Any]") -> Optional[List[str]]:
    """Extract pipe-separated author list from CrossRef item data."""
    try:
        authors = []
        # CrossRef author structure
        if "author" in item:
            for author in item["author"]:
                if isinstance(author, dict):
                    given = author.get("given", "")
                    family = author.get("family", "")
                    if family:
                        # Convert to "Lastname I" format
                        initials = "".join(c for c in given.split() if c and c[0].isupper())[:3]
                        authors.append(f"{family} {initials}".strip())
        return authors if authors else None
    except Exception as e:
        logger.debug(f"  [CrossRef] Author extraction error: {e}")
        return None


def _lookup_doi_from_pubmed(
    title: str,
    authors: List[str] = None,
    year: str = None,
    journal: str = None,
) -> Optional[Dict[str, Optional[List[str]]]]:
    """
    Look up a DOI and full author list from PubMed API using article metadata.
    Returns dict with 'doi' and 'authors' keys if found, else None.

    CRITICAL: Never fabricate DOIs. Only return DOIs from PubMed.
    """
    if not title or not title.strip():
        return None

    try:
        # Build PubMed search query
        title_clean = re.sub(r'[^\w\s]', ' ', title.strip())  # Remove special chars
        query_parts = [title_clean[:100]]  # First 100 chars of title

        if journal:
            journal_clean = re.sub(r'[^\w\s]', ' ', journal.strip())
            query_parts.append(f'"{journal_clean[:50]}"[Journal]')

        if year:
            query_parts.append(f'{year.split()[0]}[PDAT]')

        query = " AND ".join(query_parts)

        # PubMed API endpoint
        headers = {"User-Agent": "PPH-ReferenceConverter/1.0"}

        # Search for article
        search_params = {
            "db": "pubmed",
            "term": query,
            "rettype": "json",
            "retmax": 1,
        }

        search_response = requests.get(
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi",
            params=search_params,
            headers=headers,
            timeout=5
        )

        if search_response.status_code != 200:
            logger.debug(f"  [PubMed] Search returned {search_response.status_code}")
            return None

        search_data = search_response.json()
        pmids = search_data.get("esearchresult", {}).get("idlist", [])

        if not pmids:
            logger.debug(f"  [PubMed] No PMID found for: {title[:50]}")
            return None

        pmid = pmids[0]

        # Fetch full article info
        fetch_params = {
            "db": "pubmed",
            "id": pmid,
            "rettype": "json",
        }

        fetch_response = requests.get(
            "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi",
            params=fetch_params,
            headers=headers,
            timeout=5
        )

        if fetch_response.status_code != 200:
            logger.debug(f"  [PubMed] Fetch returned {fetch_response.status_code}")
            return None

        fetch_data = fetch_response.json()
        articles = fetch_data.get("result", {}).get("uids", [])

        if not articles or articles[0] not in fetch_data.get("result", {}):
            return None

        article = fetch_data["result"][articles[0]]

        # Try to find DOI in article
        doi = None

        # Check uid_list
        for uid_entry in article.get("uid_list", []):
            if uid_entry.get("name") == "DOI":
                doi = uid_entry.get("value")
                break

        # Check article attributes
        if not doi:
            for attr in article.get("article_ids", []):
                if attr.get("idtype") == "doi":
                    doi = attr.get("value")
                    break

        if doi:
            logger.info(f"  [PubMed] Found DOI: {doi}")
            # Extract authors from the article
            author_list = _extract_authors_from_pubmed_article(article)
            if author_list:
                logger.info(f"  [PubMed] Found {len(author_list)} authors")
            return {"doi": doi, "authors": author_list}
        else:
            logger.debug(f"  [PubMed] PMID found ({pmid}) but no DOI available")
            return None

    except requests.RequestException as e:
        logger.debug(f"  [PubMed] Request failed: {e}")
        return None
    except (KeyError, IndexError) as e:
        logger.debug(f"  [PubMed] Parsing error: {e}")
        return None
    except Exception as e:
        logger.debug(f"  [PubMed] Unexpected error: {e}")
        return None


def _lookup_doi_from_crossref(
    title: str,
    authors: List[str] = None,
    year: str = None,
    journal: str = None,
    volume: str = None,
    issue: str = None,
    fpage: str = None,
) -> Optional[Dict[str, Optional[List[str]]]]:
    """
    Look up a DOI and full author list from CrossRef API using article metadata.
    Returns dict with 'doi' and 'authors' keys if found, else None.

    CRITICAL: Never fabricate DOIs. Only return DOIs from CrossRef.
    """
    if not title or not title.strip():
        return None

    try:
        # Build CrossRef query
        query_parts = [title.strip()]
        if authors and len(authors) > 0:
            query_parts.append(authors[0].split("|")[0] if "|" in authors[0] else authors[0])
        if journal:
            query_parts.append(journal.strip())
        if year:
            query_parts.append(year.strip())

        query = " ".join(query_parts)

        # CrossRef API endpoint
        headers = {"User-Agent": "PPH-ReferenceConverter/1.0"}
        params = {
            "query": query,
            "rows": 1,
            "select": "DOI,title,author,published-print,published-online"
        }

        response = requests.get(
            "https://api.crossref.org/works",
            params=params,
            headers=headers,
            timeout=5
        )

        if response.status_code != 200:
            logger.debug(f"  [CrossRef] API returned {response.status_code}")
            return None

        data = response.json()
        if not data.get("message", {}).get("items"):
            logger.debug(f"  [CrossRef] No matches found for: {query}")
            return None

        # Get first result
        item = data["message"]["items"][0]
        doi = item.get("DOI")

        if not doi:
            logger.debug(f"  [CrossRef] Match found but no DOI: {item.get('title')}")
            return None

        # Verify title similarity (basic check)
        matched_title = item.get("title", [""])[0] if isinstance(item.get("title"), list) else item.get("title", "")
        if not matched_title:
            return None

        # Simple title matching: check if key words match
        source_words = set(re.findall(r'\w+', title.lower()))
        matched_words = set(re.findall(r'\w+', matched_title.lower()))

        # If less than 50% of source words match, consider it a mismatch
        if source_words and matched_words:
            overlap = len(source_words & matched_words) / len(source_words)
            if overlap < 0.5:
                logger.debug(f"  [CrossRef] Title match too weak ({overlap:.1%})")
                return None

        logger.info(f"  [CrossRef] Found DOI: {doi}")
        # Extract authors from the item
        author_list = _extract_authors_from_crossref_item(item)
        if author_list:
            logger.info(f"  [CrossRef] Found {len(author_list)} authors")
        return {"doi": doi, "authors": author_list}

    except requests.RequestException as e:
        logger.debug(f"  [CrossRef] Request failed: {e}")
        return None
    except Exception as e:
        logger.debug(f"  [CrossRef] Unexpected error: {e}")
        return None


def _validate_metadata_not_fabricated(metadata: Dict, raw_text: str) -> bool:
    """
    Validate that extracted metadata wasn't fabricated by the model.
    Check for obvious red flags like:
    - bib_surname is numeric or clearly wrong
    - bib_fname is empty when authors exist in raw_text
    - Author names completely absent from raw_text
    
    Returns True if metadata looks reasonable, False if likely fabricated.
    """
    # Check 1: bib_surname should not be purely numeric (like "6")
    surname = (metadata.get("bib_surname") or "").strip()
    if surname and surname.isdigit():
        logger.warning(f"  [Validation] Fabricated metadata detected: bib_surname='{surname}' (numeric)")
        return False
    
    # Check 2: If there's a surname, should contain reasonable names (letters)
    if surname:
        # Check if surname contains at least some letters
        if not re.search(r'[a-zA-Z]', surname):
            logger.warning(f"  [Validation] Fabricated metadata detected: bib_surname='{surname}' (no letters)")
            return False
    
    # Check 3: For journals, bib_journal should not be a single letter like "S"
    journal = (metadata.get("bib_journal") or "").strip()
    if journal and len(journal) == 1 and journal.isalpha():
        logger.warning(f"  [Validation] Fabricated metadata detected: bib_journal='{journal}' (single char)")
        return False
    
    # Check 4: If source has multiple authors (et al.), metadata should reflect it
    if " et al" in raw_text.lower() or "et al." in raw_text.lower():
        surnames = [s.strip() for s in (surname or "").split("|") if s.strip()]
        # Should have at least 2 authors before et al (fewer than 2 indicates extraction failure)
        if len(surnames) < 2:
            logger.warning(f"  [Validation] Source has 'et al' but only {len(surnames)} surnames extracted")
    
    return True


# ─────────────────────────────────────────────
# ENUMS
# ─────────────────────────────────────────────

class CitationStyle(str, Enum):
    AMA  = "AMA"
    APA  = "APA"
    CGRN = "CGRN"


class ReferenceType(str, Enum):
    JOURNAL      = "journal"
    BOOK         = "book"
    EDITED_BOOK  = "edited_book"
    BOOK_CHAPTER = "book_chapter"
    WEBSITE      = "website"
    EREFERENCE   = "ereference"
    CONFERENCE   = "conference"
    THESIS       = "thesis"
    REPORT       = "report"
    PATENT       = "patent"
    LEGAL        = "legal"
    UNKNOWN      = "unknown"


# ─────────────────────────────────────────────
# BIB FIELDS
# ─────────────────────────────────────────────

BIB_FIELDS: List[str] = [
    "bib_reftype",
    "bib_surname",
    "bib_fname",
    "bib_title",
    "bib_journal",
    "bib_year",
    "bib_volume",
    "bib_issue",
    "bib_fpage",
    "bib_lpage",
    "bib_doi",
    "bib_url",
    "bib_accessed",
    "bib_book",
    "bib_chaptertitle",
    "bib_editionno",
    "bib_ed_surname",
    "bib_ed_fname",
    "bib_publisher",
    "bib_location",
    "bib_institution",
    "bib_school",
    "bib_conference",
    "bib_confacronym",
    "bib_conflocation",
    "bib_confdate",
    "bib_deg",
    "bib_reportnum",
    "bib_series",
    "bib_isbn",
    "bib_issn",
    "bib_pmid",
    "bib_assignee",
    "bib_comment",
]


# ─────────────────────────────────────────────
# RESPONSE SCHEMAS  (lazy – avoids import-time crash when SDK absent)
# ─────────────────────────────────────────────

@functools.lru_cache(maxsize=1)
def _get_response_schema():
    from google.genai import types  # deferred import
    return types.Schema(
        type=types.Type.OBJECT,
        required=["formatted_output", "metadata", "conversion_notes"],
        properties={
            "formatted_output": types.Schema(
                type=types.Type.STRING,
                description="The fully formatted reference in the target citation style.",
            ),
            "metadata": types.Schema(
                type=types.Type.OBJECT,
                properties={
                    f: types.Schema(type=types.Type.STRING, nullable=True)
                    for f in BIB_FIELDS
                },
                required=BIB_FIELDS,
            ),
            "conversion_notes": types.Schema(
                type=types.Type.STRING,
                nullable=True,
                description="Warnings, assumptions, or missing data noted during conversion.",
            ),
        },
    )


@functools.lru_cache(maxsize=1)
def _get_validation_schema():
    from google.genai import types  # deferred import
    return types.Schema(
        type=types.Type.OBJECT,
        required=["is_valid", "validation_errors", "corrected_reference", "metadata"],
        properties={
            "is_valid": types.Schema(
                type=types.Type.BOOLEAN,
                description="True if the reference perfectly matches the target citation style rules.",
            ),
            "validation_errors": types.Schema(
                type=types.Type.ARRAY,
                items=types.Schema(type=types.Type.STRING),
                description="Specific formatting errors or missing required data. Empty if valid.",
            ),
            "corrected_reference": types.Schema(
                type=types.Type.STRING,
                description="The reference corrected to perfectly match the target citation style.",
            ),
            "metadata": types.Schema(
                type=types.Type.OBJECT,
                properties={
                    f: types.Schema(type=types.Type.STRING, nullable=True)
                    for f in BIB_FIELDS
                },
                required=BIB_FIELDS,
            ),
        },
    )


# ─────────────────────────────────────────────
# PER-TYPE STYLE RULES
# ─────────────────────────────────────────────

AMA_RULES: Dict[str, str] = {
    ReferenceType.JOURNAL: """
FORMAT (standard):         Surname FM, Surname FM. Title of article. Journal Abbrev. Year;Volume(Issue):fpage-lpage. doi:XXXXX
FORMAT (article number):   Surname FM. Title of article. Journal Abbrev. Year;Volume(Issue):e13284. doi:XXXXX
FORMAT (online ahead):     Surname FM. Title of article. Journal Abbrev. Published online Month Day, Year. doi:XXXXX
FORMAT (no DOI, URL only): Surname FM. Title of article. Journal Abbrev. Year;Volume(Issue):fpage-lpage. https://URL
FORMAT (supplement issue): Surname FM. Title of article. Journal Abbrev. Year;Volume(Suppl N):fpage-lpage. doi:XXXXX
FORMAT (letter/editorial): Surname FM. Title [letter/editorial/commentary]. Journal Abbrev. Year;Volume(Issue):fpage-lpage.
RULES:
- AUTHORS: Last name followed by SPACE (never comma) then initials with NO periods or spaces between initials.
  Format: "Smith JA" NOT "Smith, JA". Comma-space ONLY between author pairs: "Smith JA, Jones BC".
  Up to 6 authors list all; if 7 or more, list first 3 then ", et al." (WITH period after "al").
  CRITICAL: NO COMMA between surname and initials. CRITICAL: Retain EVERY initial exactly as in the source. Never collapse "JA" to "J" or drop
  any initial. bib_fname MUST be populated.
  CRITICAL: Do NOT remove study groups or steering committees (e.g., "International Steering Committee for...") attached to the author list. Treat them as part of the author group.
  No author → start with article title directly.
  NAME PREFIXES: Retain lowercase particles (van, von, de, la, du, etc.) exactly as in source.
  Do NOT capitalise them. Store as part of bib_surname. Example: "van der Berg JA" (NOT "Van Der Berg JA").
- ARTICLE TITLE: sentence case — only the first word, first word after a colon or em-dash,
  and proper nouns capitalised. All other words lowercase. No italics. No quotation marks.
  Do NOT overwrite with PubMed/CrossRef title unless the source title is completely absent.
- JOURNAL NAME: Use NLM/MEDLINE standard abbreviation. Italicise. No period after abbreviation.
- YEAR;VOLUME(ISSUE):PAGES — strict punctuation:
  Semicolon immediately after year (no space): Year;Volume
  Issue in parentheses immediately after volume (no space): Volume(Issue)
  Colon immediately after closing parenthesis (no space): (Issue):pages
  If no issue available, omit parentheses entirely: Year;Volume:pages
  If online ahead of print with no volume/issue/pages, use:
    "Published online Month Day, Year. doi:XXXXX" — omit volume/issue/page block entirely.
- PAGE RANGE: hyphen (-) between fpage and lpage: 100-110.
  Article numbers (e.g. e13284) used as-is in place of page range.
  If only one page number exists in source, output that single page only —
  NEVER repeat fpage as lpage. Set bib_lpage = null in metadata when absent.
  Each page element appears exactly ONCE — never duplicate.
- NO DATE: If publication year is unavailable, omit bib_year entirely (store as null).
  Do NOT use "n.d." in AMA — that is APA-only. Simply omit the year element.
- DOI vs URL: If DOI exists, output DOI only (doi:XXXXX). If no DOI but URL exists, output URL only.
  NEVER output both. If neither exists, end reference with a period.
- RETRACTED PAPERS: If the source indicates the paper is retracted, append "[Retracted]" after the
  article title (before the period). Example: "Smith JA. Title [Retracted]. Journal. 2020;12(3):45."
- SUPPLEMENT ISSUE: Replace the normal issue parenthetical with "(Suppl N)" — e.g. Year;110(Suppl 1):S45-S50.
  Store the supplement designator (e.g. "Suppl 1") verbatim in bib_issue.
- LETTER/EDITORIAL/COMMENTARY: Article-type label in square brackets immediately after the title
  (before the period), all lowercase. Allowed labels: [letter], [editorial], [commentary].
  Example: "Davis L. Response to new guidelines [letter]. BMJ. 2021;372:n123."
- Strip any non-breaking spaces (U+00A0) silently — never output a dot in their place.
- COMPLETENESS: Output the FULL reference. Never truncate, omit, or paraphrase any element.
""",
    ReferenceType.BOOK: """
FORMAT (with edition):    Surname FM, Surname FM. Title of book. Xth ed. Publisher; Year.
FORMAT (1st/only edition): Surname FM, Surname FM. Title of book. Publisher; Year.
FORMAT (with DOI/URL):    Surname FM. Title of book. Publisher; Year. doi:XXXXX
FORMAT (org author):      Organisation Name. Title of book. Publisher; Year.
FORMAT (translated):      Surname FM. Title of book. Translator FM, trans. Publisher; Year.
FORMAT (with volume):     Surname FM. Title of book. Vol N. Publisher; Year.
RULES:
- AUTHORS: Retain every initial. Same format rules as journal (no periods in initials, NO COMMA between surname and initials).
  Up to 6 authors; 7+ → first 6 + ", et al."
  If no personal author and an organisation is the author, use organisation name directly.
- BOOK TITLE: sentence case AND italicise. Sentence case = only first word, first word after
  colon or em-dash, and proper nouns capitalised; all other words lowercase.
- EDITION: Include only if >1st edition. Format: "Xth ed." placed after title, before publisher.
  Do NOT include edition for 1st editions.
- PUBLISHER: Retain publisher name. City NOT required in AMA 11th ed.
  Strip corporate suffixes: "Co.", "Ltd.", "Limited", "Inc.", "LLC", "Corp.", "GmbH", "S.A."
  e.g. "Springer Co., Ltd." → "Springer"
- End with period after year, UNLESS a DOI/URL follows (no period before doi: or URL).
- DOI: include if available. Prefix "doi:". No period after.
  URL: include if no DOI. No period after URL.
- TRANSLATED BOOK: Translator name(s) in same initials format as authors, followed by ", trans."
  placed after the book title and before the publisher.
  Example: "Freud S. The interpretation of dreams. Strachey J, trans. Basic Books; 2010."
- VOLUME NUMBER: "Vol N." placed after book title, before publisher.
  Example: "Smith J. Encyclopedia of biology. Vol 2. Academic Press; 2015."
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EDITED_BOOK: """
FORMAT (single editor):    Surname FM, ed. Title of book. Publisher; Year.
FORMAT (multiple editors): Surname FM, Surname FM, eds. Title of book. Publisher; Year.
FORMAT (with edition):     Surname FM, ed. Title of book. Xth ed. Publisher; Year.
FORMAT (with DOI/URL):     Surname FM, ed. Title of book. Publisher; Year. doi:XXXXX
RULES:
- EDITOR LABEL: "ed." for single editor, "eds." for two or more editors — placed immediately
  after the last editor name, before the title. CRITICAL: Never use "ed." for multiple editors.
- IDENTIFICATION: classify as edited_book ONLY when editor markers ("ed.", "eds.", "Ed.", "Eds.",
  "edited by") are present AND no separate chapter author appears before the title.
  If both chapter author AND editor are present → classify as book_chapter instead.
  In metadata: bib_ed_surname / bib_ed_fname MUST be populated; bib_surname / bib_fname = null.
- EDITOR NAMES: same initials format as author names (no periods between initials, NO COMMA between surname and initials).
- BOOK TITLE: sentence case AND italicise.
- EDITION: Include only if >1st edition. Format: "Xth ed." placed after title, before publisher.
- PUBLISHER: Retain name; strip corporate suffixes; city not required.
- End with period after year, UNLESS DOI/URL follows.
- DOI/URL: include if available. No period after.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.BOOK_CHAPTER: """
FORMAT (single editor):    Author FM. Chapter title. In: Editor FM, ed. Book title. Publisher; Year:fpage-lpage.
FORMAT (multiple editors): Author FM. Chapter title. In: Editor FM, Editor FM, eds. Book title. Publisher; Year:fpage-lpage.
FORMAT (with edition):     Author FM. Chapter title. In: Editor FM, ed. Book title. Xth ed. Publisher; Year:fpage-lpage.
FORMAT (no editor listed): Author FM. Chapter title. In: Book title. Publisher; Year:fpage-lpage.
FORMAT (with DOI):         Author FM. Chapter title. In: Editor FM, ed. Book title. Publisher; Year:fpage-lpage. doi:XXXXX
RULES:
- CRITICAL FIELD MAPPING:
  bib_chaptertitle = the title of the chapter. NEVER put the chapter title in bib_title.
  bib_book = the title of the book.
- CHAPTER AUTHOR: comes first. Retain every initial. Same format as journal authors (NO COMMA between surname and initials).
  Up to 6 authors; 7+ → first 6 + ", et al."
- CHAPTER TITLE: sentence case (same rules as article title). No italics. No quotation marks.
- "In:" (capital I, colon, space) introduces the editor/book block.
- EDITOR LABEL: "ed." for single editor, "eds." for two or more editors — after last editor name.
  CRITICAL: Never use "ed." when there are multiple editors — it MUST be "eds."
  If no editor is listed, omit the editor block entirely and go directly to book title.
- BOOK TITLE: sentence case AND italicise.
- EDITION: "Xth ed." placed after book title, before publisher. Omit for 1st edition.
- PAGES: placed after year, preceded by colon (no space before colon): Year:fpage-lpage.
  Use hyphen (-) between pages. No "pp." prefix in AMA.
  If only one page, output that single page only — never repeat fpage.
- PUBLISHER: Retain name; strip corporate suffixes; city not required.
- DOI/URL: include if available, after the page range. No period after.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.WEBSITE: """
FORMAT (personal author):  Author FM. Title of page. Website Name. Published Month Day, Year. Accessed Month Day, Year. URL
FORMAT (org author):       Organisation Name. Title of page. Website Name. Published Month Day, Year. Accessed Month Day, Year. URL
FORMAT (updated):          Author FM. Title of page. Website Name. Updated Month Day, Year. Accessed Month Day, Year. URL
FORMAT (no author):        Title of page. Website Name. Published Month Day, Year. Accessed Month Day, Year. URL
FORMAT (no pub date):      Title of page. Website Name. Accessed Month Day, Year. URL
RULES:
- PAGE/DOCUMENT TITLE: sentence case (first word and proper nouns only). No italics. No quotation marks.
  Store ONLY in bib_title — NEVER duplicate in bib_journal.
- WEBSITE NAME: title case. The domain or website name (e.g. "CDC", "Merriam-Webster", "NIH").
  MUST be stored in bib_journal field. NOT the page title. Placed after the page title.
  Example: If source is "Merriam-Webster.com - Professional" then:
    bib_title = "Professional" (the page title)
    bib_journal = "Merriam-Webster" (the website name)
- AUTHOR — personal: Surname FM format. NO COMMA between surname and initials. Space only.
  If no personal author, check for an organisation author.
  ORGANIZATION AUTHOR: If a government agency, institution, or organisation authored the page (e.g.
  "Centers for Disease Control and Prevention", "Committee on Accreditation..."), use the organisation
  name in the author position before the title. Store in bib_surname with bib_fname = "". DO NOT omit it.
  Example: "Centers for Disease Control and Prevention. Physical activity. CDC. Published 2022. Accessed..."
  NO AUTHOR: If there is no personal author AND no organisation author, start directly with the page title.
- PUBLICATION DATE: "Published Month Day, Year" or "Updated Month Day, Year" as appropriate.
  Use "Updated" when source says updated/modified; use "Published" for initial publication.
  If only a year is available: "Published Year." If no publication date, omit entirely.
- ACCESS DATE: Always include "Accessed Month Day, Year." before the URL.
- URL: Full URL on same line after access date. No period after URL. NEVER omit the URL.
- bib_url MUST contain the full URL extracted from "Available from:", "URL:", or any URL in the source.
  Strip "Available from:" and "Retrieved from:" prefixes — store only the bare URL.
- No DOI for websites.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference including the URL. Never truncate or omit any element.
""",
    ReferenceType.EREFERENCE: """
FORMAT (author, single editor):    Author FM. Entry title. In: Editor FM, ed. Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (author, multiple editors): Author FM. Entry title. In: Editor FM, Editor FM, eds. Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (author, no editor):        Author FM. Entry title. In: Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (org author):               Organisation Name. Entry title. In: Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (no year):                  Author FM. Entry title. In: Editor FM, ed. Reference Title. Publisher. Accessed Month Day, Year. URL
RULES:
- ENTRY TITLE: sentence case. No italics. No quotation marks. Store ONLY in bib_title — NEVER duplicate.
- REFERENCE BOOK/DATABASE TITLE: title case. Store in bib_book. NOT in bib_journal.
- AUTHOR: Always include the author in the author position before the entry title.
  Personal author: Surname FM format (same as journal). NO COMMA between surname and initials.
  ORGANIZATION AUTHOR: If an organisation, agency, or group authored the entry, store in bib_surname
  with bib_fname = "". Place organisation name before the entry title. DO NOT omit it.
- EDITOR LABEL: "ed." for single editor, "eds." for two or more editors.
  CRITICAL: Never use "ed." when there are multiple editors.
  If no editor listed, omit the editor block entirely — go directly to Reference Title.
- PUBLISHER: Retain name; strip corporate suffixes. Store in bib_publisher.
- YEAR: Include after publisher with semicolon: "Publisher; Year." If no year, omit year entirely.
- ACCESS DATE: Always include "Accessed Month Day, Year." before URL.
- URL: Full URL. No period after URL. NEVER omit the URL.
  bib_url MUST contain the full URL. Strip "Available from:" and "Retrieved from:" prefixes.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference including the URL. Never truncate or omit any element.
""",
    ReferenceType.CONFERENCE: """
FORMAT (paper presentation):   Author FM. Title of paper. Paper presented at: Full Conference Name; Month Day–Day, Year; City, Country.
FORMAT (poster presentation):  Author FM. Title of poster. Poster presented at: Full Conference Name; Month Day–Day, Year; City, Country.
FORMAT (abstract/poster):      Author FM. Title [abstract/poster]. Presented at: Full Conference Name; Month Day–Day, Year; City, Country.
FORMAT (published proceedings, single ed):   Author FM. Chapter/paper title. In: Editor FM, ed. Proceedings Title. Publisher; Year:fpage-lpage.
FORMAT (published proceedings, multiple eds): Author FM. Chapter/paper title. In: Editor FM, Editor FM, eds. Proceedings Title. Publisher; Year:fpage-lpage.
FORMAT (proceedings online URL): Author FM. Title of paper. In: Editor FM, ed. Proceedings Title. Publisher; Year. Accessed Month Day, Year. URL
RULES:
- PRESENTATION TYPE: Use "Paper presented at:" for oral/paper presentations;
  "Poster presented at:" for poster presentations;
  "Presented at:" for abstracts or when type is unspecified (use with [abstract] or [poster] label).
- ABSTRACT/POSTER LABEL: When an abstract or poster is identified, append [abstract] or [poster]
  in square brackets immediately after the title, before the period.
  Example: "Iyer S. COVID-19 vaccine response [poster]. Presented at: IMA Conference; January 2022; Chennai, India."
- DATE: Month Day–Day, Year (en dash between days). Place after conference name, separated by semicolon.
- LOCATION: City, State/Country. Place after date, separated by semicolon.
- End with period after location.
- PUBLISHED PROCEEDINGS: treat exactly like a book chapter — same rules apply.
  EDITOR LABEL: "ed." for single editor, "eds." for two or more editors.
- AUTHORS: same format as journal.
- TITLE: sentence case. No italics for presented paper/poster title.
- CRITICAL FIELD MAPPING:
  bib_title      = title of the PAPER or POSTER being presented (NOT the conference name).
  bib_conference = the FULL conference/symposium name (e.g. "Annual Meeting of the American
                   College of Sports Medicine"). NEVER put this in bib_title.
  bib_confdate   = the DATE RANGE of the conference (e.g. "May 28–31, 2024").
                   NEVER put the standalone year here — year goes in bib_year.
  bib_conflocation = city and country/state of the conference venue.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.THESIS: """
FORMAT (doctoral, with URL):     Author FM. Title of thesis [doctoral dissertation]. University Name; Year. Accessed Month Day, Year. URL
FORMAT (master's, with URL):     Author FM. Title of thesis [master's thesis]. University Name; Year. Accessed Month Day, Year. URL
FORMAT (doctoral, no URL):       Author FM. Title of thesis [doctoral dissertation]. University Name; Year.
FORMAT (master's, no URL):       Author FM. Title of thesis [master's thesis]. University Name; Year.
FORMAT (with DOI):               Author FM. Title of thesis [doctoral dissertation]. University Name; Year. doi:XXXXX
FORMAT (from database):          Author FM. Title of thesis [doctoral dissertation]. University Name; Year. Database Name.
FORMAT (unpublished):            Author FM. Title of thesis [master's thesis]. University Name; Year. Unpublished.
FORMAT (undergraduate):          Author FM. Title of thesis [undergraduate thesis]. University Name; Year.
RULES:
- TITLE: sentence case, italicise if possible.
- DEGREE TYPE in brackets immediately after title (no comma before bracket):
  Use [doctoral dissertation] for PhD/doctoral work.
  Use [master's thesis] for master's-level work.
  Use [undergraduate thesis] for undergraduate-level work.
  Use the exact degree label if specified in the source.
- UNIVERSITY: treat as publisher (replaces publisher field).
- YEAR: placed after university, preceded by semicolon: University; Year.
- If available online: add "Accessed Month Day, Year. URL" after year.  No period after URL.
- DATABASE: If retrieved from a database (e.g. ProQuest Dissertations & Theses Global),
  add the database name after the year. Example: "University; Year. ProQuest Dissertations & Theses Global."
- UNPUBLISHED: If not publicly accessible, append "Unpublished." after year.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.REPORT: """
FORMAT (with report number, DOI):    Author FM. Title of report. Institution Name; Year. Report No. XXXX. doi:XXXXX
FORMAT (with report number, URL):    Author FM. Title of report. Institution Name; Year. Report No. XXXX. URL
FORMAT (internet, with access date): Author FM. Title of report. Institution Name; Year. Accessed Month Day, Year. URL
FORMAT (no report number, DOI):      Author FM. Title of report. Institution Name; Year. doi:XXXXX
FORMAT (no report number, URL only): Author FM. Title of report. Institution Name; Year. URL
FORMAT (org author):                 Organisation Name. Title of report. Institution Name; Year.
FORMAT (government, no URL):         Government Agency. Title of report. Publisher; Year.
FORMAT (unpublished/internal):       Author FM. Title of report. Institution Name; Year. Unpublished.
RULES:
- TITLE: sentence case. No italics.
- AUTHORS: same format as journal (Surname FM). NO COMMA between surname and initials — space only.
  If organisation is the author, use organisation name directly in the author position (before the title).
  Store in bib_surname with bib_fname = "".
  CRITICAL: Do NOT replace the stated org author with personal authors from a related journal article.
  The org author is determined solely by the source text.
- INSTITUTION: replaces publisher. Retain full institution name. Do NOT strip meaningful words.
- YEAR: placed after institution, preceded by semicolon: Institution; Year.
- REPORT NUMBER: include if present. Format: "Report No. XXXX." placed after year. Omit if absent.
- ACCESS DATE (internet reports): If source has a cited/access date and a URL, include
  "Accessed Month Day, Year." immediately before the URL. If no access date, omit it.
- DOI: prefix "doi:". Place after year (or after report number if present). No period after.
- URL: include if no DOI and source has a URL. No period after URL. NEVER omit the URL.
  bib_url MUST contain the full URL extracted from "Available from:" or any URL in the source.
  Strip "Available from:" and "Retrieved from:" prefixes — store only the bare URL.
  bib_accessed MUST contain the access/cited date extracted from "[cited YYYY Month DD]" etc.
- UNPUBLISHED/INTERNAL: append "Unpublished." after year if not publicly accessible.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference including the URL. Never truncate or omit any element.
""",
    ReferenceType.PATENT: """
FORMAT (issued patent):      Inventor FM. Title of invention. Country patent XXXXXXX. Issued Month Day, Year.
FORMAT (patent application): Inventor FM. Title of invention. Country patent application XXXXXXX. Published Month Day, Year.
FORMAT (WIPO/PCT):           Inventor FM. Title of invention. WIPO patent WOXXXX/XXXXXX. Published Month Day, Year.
FORMAT (European):           Inventor FM. Title of invention. European patent EPXXXXXXX. Issued Month Day, Year.
FORMAT (with assignee):      Inventor FM; Assignee Name. Title of invention. Country patent XXXXXXX. Issued Month Day, Year.
FORMAT (online, URL):        Inventor FM. Title of invention. Country patent XXXXXXX. Issued Month Day, Year. Accessed Month Day, Year. URL
FORMAT (with DOI):           Inventor FM. Title of invention. Patent number. Published Month Day, Year. doi:XXXXX
RULES:
- INVENTORS: Last name followed by initials with NO periods between initials (same as journal authors).
  Comma-space between inventors. Semicolon separates inventor(s) from assignee company.
  bib_surname/bib_fname → inventors; bib_assignee → assignee company name.
- TITLE: sentence case. No italics.
- PATENT DESIGNATION (bib_reportnum): The full patent identifier, including country/office prefix and number.
  Examples: "US patent 10,123,456" | "US patent application 2021/0123456" |
            "WIPO patent WO2022/098765" | "European patent EP3456789" | "Indian patent 2020112345"
  For applications: use "Country patent application XXXXXXX".
  For WIPO: use "WIPO patent WOXXXX/XXXXXX".
  For European: use "European patent EPXXXXXXX".
- ISSUED/PUBLISHED DATE: Full date written out — "Issued Month Day, Year." or "Published Month Day, Year."
  Not semicolon-year style. bib_year = year only (e.g. "2020").
- ASSIGNEE (company): appears after inventors, separated by semicolon. bib_assignee = company name.
  Example: "Rao V; Infosys. Blockchain-based healthcare system. Indian patent 2020112345. Issued August 12, 2023."
- ACCESS DATE + URL: if accessed online, add "Accessed Month Day, Year. URL" after issued/published date. No period after URL.
- DOI: if present, prefix "doi:". No period after.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
}

APA_RULES: Dict[str, str] = {
    ReferenceType.JOURNAL: """
FORMAT (standard):          Surname, F. M., & Surname, F. M. (Year). Title of article. *Journal Name*, *Volume*(Issue), fpage–lpage. https://doi.org/XXXXX
FORMAT (article number):    Surname, F. M. (Year). Title of article. *Journal Name*, *Volume*(Issue), Article e13284. https://doi.org/XXXXX
FORMAT (online ahead):      Surname, F. M. (Year). Title of article. *Journal Name*. Advance online publication. https://doi.org/XXXXX
FORMAT (suppl. issue):      Surname, F. M. (Year). Title of article. *Journal Name*, *Volume*(Suppl. X), fpage–lpage. https://doi.org/XXXXX
FORMAT (no DOI, URL):       Surname, F. M. (Year). Title of article. *Journal Name*, *Volume*(Issue), fpage–lpage. https://URL
RULES:
- AUTHORS: Surname, F. M. format. Initials each followed by a period and space (e.g. Smith, J. A.).
  Comma-space between authors. Use "&" before the last author (never "and").
  Up to 20 authors list all; if >20, list first 19, then "…", then last author.
  CRITICAL ELLIPSIS RULE: If the source text ends with "et al." or "..." and provides only a partial list of authors, you MUST respect this truncation. When formatting, use the format `, et al.` (with a comma) before the final author element if the true author list is truncated. Do not invent authors or drop the truncation marker. Do not use an ellipsis for `et al.`.
  CRITICAL: Retain EVERY initial verbatim from source. Never collapse "JA" → "J" or drop any initial.
  CRITICAL: Every initial MUST have a period: "Smith, J. A., Jones, B. C."
  CRITICAL ORGANISATION RULE: Do NOT remove study groups or steering committees. If an organisation, study group, advisory team, or working group is mixed in with personal authors, MOVE it to the very end of the author list. When an organisation is the last author in the group, it must be preceded by "&" (e.g., "Smith, J. A., & Sense Advisory Team Working Group."). DO NOT add any extra authors or labels after the organisation.
  CRITICAL: Retain name generation suffixes exactly (Jr., Sr., II, III, IV). Format as:
    "Collins, Jr., J. W." — suffix comes after surname, before initials, separated by commas.
  No author → start with article title directly.
  NAME PREFIXES: For names with lowercase particles (van, von, de, la, du, der, etc.), retain the
  prefix exactly as in the source. Do NOT capitalise particles. Store as part of bib_surname.
  Example: bib_surname = "van der Berg" → formatted as "van der Berg, A. J." (NOT "Van Der Berg").
- YEAR: (Year). followed by period. CRITICAL: Retain any lowercase suffix (2022a, 2024b) — NEVER remove.
  If full date available, format as (Year, Month Day), e.g. (2025, May 28).
- ARTICLE TITLE: Sentence case ONLY — capitalise only: first word, first word after colon/em-dash,
  proper nouns. All other words lowercase. No italics. No quotation marks.
  CRITICAL: When converting to sentence case, preserve the exact capitalization of country acronyms and abbreviations (e.g., U.S., U.K., USA, UK). Do not convert them to lowercase.
  CRITICAL: Do NOT overwrite with database title. Source title is authoritative.
- JOURNAL NAME: Title case AND italicised (*Journal Name*). Use EXACTLY the journal name from the source.
  CRITICAL: NEVER append parenthetical location qualifiers like "(Basel, Switzerland)" unless verbatim in source.
- STATUS: Retain the EXACT status label from the source:
  "Advance online publication" = article ahead of print, not yet assigned to an issue/volume.
  "Published online" or "Online published" = published but no print issue yet.
  CRITICAL: Use verbatim wording from source. Do NOT convert between these terms.
  Format: *Journal Name*. Advance online publication. https://doi.org/XXXXX
- Volume: italicised (*Volume*). Issue in parentheses immediately after volume, NOT italicised.
  Supplemental issue: *Volume*(Suppl. X) or *Volume*(Suppl. N).
  CRITICAL: NO space between closing parenthesis of issue and comma: *12*(3), fpage
- PAGE RANGE: en dash (–) between fpage and lpage. Use EXACTLY as in source.
  Article numbers: write as "Article XXXXX" or the number as-is.
  If only fpage exists, output only that page — NEVER repeat as lpage. bib_lpage = null when absent.
- DOI vs URL: If DOI exists, output DOI only as https://doi.org/XXXXX — no period after.
  If no DOI but URL exists, output URL only — no period after. NEVER output both.
- RETRACTED PAPERS: If the source indicates the paper is retracted, append "[Retracted]" after
  the article title (before the period). Example: "Smith, J. A. (2020). Title [Retracted]. *Journal*, *12*(3), 45."
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate, omit, or remove any element.
""",
    ReferenceType.BOOK: """
FORMAT (standard):       Surname, F. M., & Surname, F. M. (Year). *Title of book* (Xth ed.). Publisher. https://doi.org/XXXXX
FORMAT (1st/only ed.):   Surname, F. M. (Year). *Title of book*. Publisher. https://doi.org/XXXXX
FORMAT (URL, no DOI):    Surname, F. M. (Year). *Title of book*. Publisher. https://URL
FORMAT (no URL/DOI):     Surname, F. M. (Year). *Title of book* (Xth ed.). Publisher.
FORMAT (org author):     Organisation Name. (Year). *Title of book*. Publisher.
FORMAT (with trans.):    Surname, F. M. (Year). *Title of book* (F. M. Translator, Trans.; Xth ed.). Publisher. (Original work published YYYY)
RULES:
- AUTHORS: Retain every initial with periods. Same format rules as journal.
  Up to 20 authors list all; 21+ → first 19, "…", last author.
  CRITICAL: If the source uses an ellipsis in the author list, apply the same rule as journal: only use "…" if total authors >20; otherwise list all. Never store "…" as an entry in bib_surname/bib_fname.
  No personal author → organisation name in author position.
- YEAR: Retain any letter suffix. Same rules as journal.
- BOOK TITLE: Sentence case AND italicised (*Title*). Sentence case = only first word, first word after
  colon/em-dash, and proper nouns capitalised; all other words lowercase.
  CRITICAL: When converting to sentence case, preserve the exact capitalization of country acronyms and abbreviations (e.g., U.S., U.K., USA, UK). Do not convert them to lowercase.
- EDITION: in parentheses after title if >1st: (2nd ed.). Omit for 1st editions.
  If translator AND edition, combine: (F. M. Translator, Trans.; 2nd ed.)
- TRANSLATOR: "(F. M. Translator, Trans.)" in parentheses after title/edition, before publisher.
  For translated works, add "(Original work published YYYY)" at end, no period after.
- PUBLISHER (APA 7th): CRITICAL — ALWAYS retain the publisher name. NEVER delete it.
  Omit only the city/location prefix (e.g. "New York:", "London:", "Thousand Oaks, CA:").
  Strip ONLY these corporate suffixes: "Co.", "Ltd.", "Limited", "Inc.", "LLC", "Corp.", "GmbH", "S.A.", "Lda.", "Pty."
  EXCEPTION: If publisher name EXACTLY equals the author/org name, omit publisher per APA 7th.
  CRITICAL: If the source shows the publisher as the word "Author" (meaning the authoring org is
  also the publisher), store bib_publisher = "Author" exactly — do NOT expand to the org name.
  Multiple publishers: separate with semicolons.
- DOI as https://doi.org/XXXXX or full URL. No period after.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EDITED_BOOK: """
FORMAT (single editor):     Surname, F. M. (Ed.). (Year). *Title of book*. Publisher. https://doi.org/XXXXX
FORMAT (two editors):       Surname, F. M., & Surname, F. M. (Eds.). (Year). *Title of book*. Publisher.
FORMAT (3+ editors):        Surname, F. M., Surname, F. M., & Surname, F. M. (Eds.). (Year). *Title of book*. Publisher.
FORMAT (with edition):      Surname, F. M. (Ed.). (Year). *Title of book* (Xth ed.). Publisher.
FORMAT (with trans.):       Surname, F. M. (Ed.). (Year). *Title of book* (F. M. Translator, Trans.). Publisher.
RULES:
- EDITOR LABEL: "(Ed.)" for single editor; "(Eds.)" for two or more editors.
  Placed immediately after the last editor name/initial, BEFORE the year parenthesis.
  CRITICAL: NEVER use "(Ed.)" when there are multiple editors — it MUST be "(Eds.)".
  CRITICAL: formatted_output MUST include "(Ed.)" or "(Eds.)".
- IDENTIFICATION: classify as edited_book ONLY when editor markers are present AND no separate
  chapter author appears before the title. Both author + editor → book_chapter. No markers → book.
  bib_ed_surname / bib_ed_fname MUST be populated; bib_surname / bib_fname = null.
- EDITOR NAMES: same initial format as authors (periods after each initial, comma before).
  Use "&" before last editor when 2+ editors.
- YEAR: Retain any letter suffix. Placed in parentheses after "(Ed.)"/"(Eds.)".
- BOOK TITLE: Sentence case AND italicised.
- EDITION: in parentheses after title if >1st.
- PUBLISHER: CRITICAL — ALWAYS retain. Omit city only. Strip corporate suffixes only.
  If publisher = editor's organisation, omit publisher per APA 7th.
- DOI as https://doi.org/XXXXX or URL. No period after.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.BOOK_CHAPTER: """
FORMAT (single editor):      Author, F. M. (Year). Chapter title. In F. M. Editor (Ed.), *Book title* (Xth ed., pp. fpage–lpage). Publisher. https://doi.org/XXXXX
FORMAT (two editors):        Author, F. M. (Year). Chapter title. In F. M. Editor & F. M. Editor (Eds.), *Book title* (pp. fpage–lpage). Publisher.
FORMAT (3+ editors):         Author, F. M. (Year). Chapter title. In F. M. Editor, F. M. Editor, & F. M. Editor (Eds.), *Book title* (pp. fpage–lpage). Publisher.
FORMAT (no editor):          Author, F. M. (Year). Chapter title. In *Book title* (pp. fpage–lpage). Publisher.
FORMAT (1st ed., no DOI):    Author, F. M. (Year). Chapter title. In F. M. Editor (Ed.), *Book title* (pp. fpage–lpage). Publisher.
FORMAT (with volume):        Author, F. M. (Year). Chapter title. In F. M. Editor (Ed.), *Book title* (Vol. 1, pp. fpage–lpage). Publisher.
RULES:
- CRITICAL FIELD MAPPING:
  bib_chaptertitle = the title of the chapter. NEVER put the chapter title in bib_title.
  bib_book = the title of the book.
- CHAPTER AUTHOR: first. Retain every initial with periods. Same format as journal.
  Year + letter suffix rules apply.
- CHAPTER TITLE: Sentence case. No italics. No quotation marks.
  CRITICAL: When converting to sentence case, preserve the exact capitalization of country acronyms and abbreviations (e.g., U.S., U.K., USA, UK). Do not convert them to lowercase.
  CRITICAL: Do NOT overwrite with database title unless source title completely absent.
- "In" (no colon in APA) introduces the editor block. Editor initials appear BEFORE surname.
  e.g. "In R. L. Smith (Ed.)," — NOT "In Smith, R. L. (Ed.),"
- EDITOR LABEL: "(Ed.)" for one editor; "(Eds.)" for two or more editors.
  CRITICAL: Never use "(Ed.)" when there are multiple editors — it MUST be "(Eds.)".
  If no editor listed, use "In *Book title*" with no editor block at all.
- MULTIPLE EDITORS punctuation:
  Two editors:  In A. B. Smith & C. D. Jones (Eds.),
  Three+:       In A. B. Smith, C. D. Jones, & E. F. Brown (Eds.),
- BOOK TITLE: Sentence case AND italicised. Appears inside the parenthetical block.
- EDITION, VOLUME, and PAGES in same parentheses after book title:
  With edition + vol + pages:  *Book title* (2nd ed., Vol. 1, pp. 10–20).
  With edition + pages:        *Book title* (2nd ed., pp. 10–20).
  With volume + pages:         *Book title* (Vol. 1, pp. 10–20).
  Without edition:             *Book title* (pp. 10–20).
  Single page only:            *Book title* (pp. 10).
  NEVER repeat fpage as lpage. "pp." prefix REQUIRED. En dash (–) between pages.
- PUBLISHER: Retain name; strip corporate suffixes; omit city/location.
  If publisher = chapter author's organisation, omit per APA 7th.
- DOI as https://doi.org/XXXXX or URL. No period after.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.WEBSITE: """
FORMAT (with author, date):      Author, F. M. (Year, Month Day). Title of page. *Website Name*. URL
FORMAT (with author, year only): Author, F. M. (Year). Title of page. *Website Name*. URL
FORMAT (no author, with date):   Title of page. (Year, Month Day). *Website Name*. URL
FORMAT (no author, no date):     Title of page. (n.d.). *Website Name*. URL
FORMAT (changing content):       Author, F. M. (Year, Month Day). Title of page. *Website Name*. Retrieved Month Day, Year, from URL
FORMAT (org author):             Organisation Name. (Year, Month Day). Title of page. *Website Name*. URL
RULES:
- PAGE/ARTICLE TITLE: Sentence case (first word and proper nouns only). No italics. No quotation marks.
  bib_title = the PAGE or ENTRY title ONLY — NOT the website name, database name, or URL.
  The website/database name goes in bib_journal or bib_book.
- WEBSITE/ORGANISATION NAME: Title case AND italicised (*Website Name*). Store in bib_journal.
  CRITICAL: bib_journal must be the HUMAN-READABLE website name ONLY (e.g., "Cancer Care Ontario",
  "World Health Organization"). NEVER store a URL, domain, or any "http" string in bib_journal.
  If you cannot identify a distinct website name, leave bib_journal empty.
  If the author IS the website/organisation, omit the site name to avoid repetition.
- AUTHOR: No personal author → title moves to author position.
  Organisation as author → use org name in author position.
- YEAR: CRITICAL — bib_year MUST be stored as "YYYY" or "YYYY, Month Day" (e.g., "2024, June 26").
  NEVER store as "Month Day, YYYY" (e.g., never "June 26, 2024"). Full date if available; year only if not.
  No date available → store as "n.d.".
  bib_year = publication year only. bib_accessed = retrieval date. ALWAYS separate fields. Never confuse.
- RETRIEVAL DATE: Omit for stable content. Include for content that may change (wikis, live data):
  "Retrieved Month Day, Year, from URL"
  CRITICAL: bib_accessed = date ONLY (e.g., "December 15, 2023"). NEVER include "from", "Retrieved",
  or any URL in bib_accessed. The URL goes in bib_url; bib_accessed is the date string only.
- URL: bib_url must contain ONLY the bare URL (starting with https:// or http://).
  Strip any leading text like "from ", "Retrieved from ", "From ", etc.
  Strip any trailing text such as ", on DATE", ", accessed DATE". No period after URL.
  NEVER store a URL in bib_journal or bib_accessed.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EREFERENCE: """
FORMAT (with editor, single):    Author, F. M. (Year). Entry title. In F. M. Editor (Ed.), *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
FORMAT (with editors, multiple): Author, F. M. (Year). Entry title. In F. M. Editor & F. M. Editor (Eds.), *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
FORMAT (no editor):              Author, F. M. (Year). Entry title. In *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
FORMAT (no date):                Author, F. M. (n.d.). Entry title. In F. M. Editor (Ed.), *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
RULES:
- TYPE DETECTION: Classify as ereference when the source is an entry in an online encyclopaedia,
  dictionary, or database (e.g. StatPearls, UpToDate, Encyclopaedia Britannica, MedlinePlus,
  Cochrane Library entry). Do NOT classify these as journal or website.
- ENTRY TITLE: Sentence case. No italics. No quotation marks. Store in bib_title.
- REFERENCE/DATABASE TITLE: Title case AND italicised (*Reference Title*). Store in bib_book.
- EDITOR LABEL: "(Ed.)" for one editor; "(Eds.)" for two or more editors.
  CRITICAL: Never use "(Ed.)" for multiple editors.
  If no editor, use "In *Reference Title*" with no editor block.
- Editor initials appear BEFORE surname, same as book_chapter rule.
- YEAR: Store as "YYYY" if only a year is available, OR as "YYYY, Month Day" if a full publication
  date is given in the source (e.g., "2023, April 24"). bib_accessed = retrieval date — always
  separate; NEVER put retrieval date in bib_year. Use "n.d." only if no date is available at all.
- PUBLISHER: Retain name; strip corporate suffixes; omit city/location.
- RETRIEVAL DATE: Always include "Retrieved Month Day, Year, from URL" — required for e-references
  since content may be updated.
  CRITICAL: bib_accessed = date ONLY (e.g., "December 15, 2023"). NEVER include "from", "Retrieved",
  or any URL in bib_accessed. The URL goes in bib_url; bib_accessed is the date string only.
- URL: bib_url must contain ONLY the bare URL (starting with https:// or http://).
  Strip any leading text like "from ", "Retrieved from ", "From ", etc. No period after URL.
  NEVER store a URL in bib_book or bib_accessed.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.CONFERENCE: """
FORMAT (paper, with date):         Author, F. M. (Year, Month Day–Day). Title of paper [Paper presentation]. Conference Name, City, Country. https://doi.org/XXXXX
FORMAT (paper, year only):         Author, F. M. (Year). Title of paper [Paper presentation]. Conference Name, City, Country.
FORMAT (poster):                   Author, F. M. (Year, Month Day–Day). Title of poster [Poster session]. Conference Name, City, Country.
FORMAT (symposium contribution):   Author, F. M. (Year, Month Day–Day). Title [Symposium contribution]. In F. M. Chair (Chair), *Symposium Title*. Conference Name, City, Country.
FORMAT (proceedings, single ed.):  Author, F. M. (Year). Title. In F. M. Editor (Ed.), *Proceedings Title* (pp. X–X). Publisher. https://doi.org/XXXXX
FORMAT (proceedings, multiple ed.): Author, F. M. (Year). Title. In F. M. Editor & F. M. Editor (Eds.), *Proceedings Title* (pp. X–X). Publisher.
RULES:
- CRITICAL FIELD MAPPING:
  bib_title      = title of the PAPER or POSTER being presented (NOT the conference name).
  bib_conference = the FULL conference/symposium name.
                   NEVER put this in bib_title. NEVER put the paper title in bib_conference.
  bib_confdate   = the DATE RANGE of the conference (e.g. "May 28–31, 2024").
                   NEVER put the standalone conference year here — year goes in bib_year.
  bib_conflocation = city and country/state of the venue.
- PRESENTATION TYPE in brackets after title:
  [Paper presentation] for oral presentations.
  [Poster session] for poster presentations.
  [Symposium contribution] for symposium contributions.
  Use the most accurate descriptor from the source.
- YEAR: Full date in parentheses if available: (Year, Month Day–Day). Year only if no full date.
  CRITICAL: Retain any lowercase letter suffix (2024a). NEVER remove. En dash (–) between days.
- LOCATION: City, Country (or City, State, Country) placed after conference name, separated by comma.
- CONFERENCE NAME: Retain full name as given.
- PUBLISHED PROCEEDINGS: treat exactly like book_chapter.
  EDITOR LABEL: "(Ed.)" for one editor; "(Eds.)" for two or more editors — same rules as book_chapter.
  MULTIPLE EDITORS: same punctuation rules as book_chapter (comma + "&" before last).
- TITLE of paper/poster: Sentence case. No italics.
- AUTHORS: same format as journal (initials with periods, "&" before last).
- DOI as https://doi.org/XXXXX or URL. No period after.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.THESIS: """
FORMAT (doctoral, database):      Author, F. M. (Year). *Title of thesis* [Doctoral dissertation, University Name]. Database Name. URL
FORMAT (doctoral, institutional):  Author, F. M. (Year). *Title of thesis* [Doctoral dissertation, University Name]. Institutional Repository Name. URL
FORMAT (doctoral, unpublished):    Author, F. M. (Year). *Title of thesis* [Unpublished doctoral dissertation]. University Name.
FORMAT (master's, database):       Author, F. M. (Year). *Title of thesis* [Master's thesis, University Name]. Database Name. URL
FORMAT (master's, unpublished):    Author, F. M. (Year). *Title of thesis* [Unpublished master's thesis]. University Name.
FORMAT (no date):                  Author, F. M. (n.d.). *Title of thesis* [Doctoral dissertation, University Name]. Database Name.
RULES:
- TITLE: Sentence case AND italicised (*Title*).
- DEGREE TYPE and INSTITUTION in square brackets immediately after title:
  Published:   [Doctoral dissertation, University Name] or [Master's thesis, University Name]
  Unpublished: [Unpublished doctoral dissertation] — University Name then becomes the publisher
               placed after the bracket: [Unpublished doctoral dissertation]. University Name.
- YEAR: (n.d.) if year not available.
- DATABASE / REPOSITORY NAME: if retrieved from ProQuest, institutional repository, or similar,
  include database/repository name after the bracket, before URL. Omit if not applicable.
- No period after URL.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.REPORT: """
FORMAT (with report no., DOI):   Author, F. M. (Year). *Title of report* (Report No. XXXX). Institution. https://doi.org/XXXXX
FORMAT (with report no., URL):   Author, F. M. (Year). *Title of report* (Report No. XXXX). Institution. https://URL
FORMAT (no report no., DOI):     Author, F. M. (Year). *Title of report*. Institution. https://doi.org/XXXXX
FORMAT (no report no., URL):     Author, F. M. (Year). *Title of report*. Institution. https://URL
FORMAT (no report no., no URL):  Author, F. M. (Year). *Title of report*. Institution.
FORMAT (org author):             Organisation Name. (Year). *Title of report* (Report No. XXXX). Institution. https://doi.org/XXXXX
RULES:
- TITLE: Sentence case AND italicised (*Title*).
- REPORT NUMBER: in parentheses after title if available: (Report No. XXXX). Omit if absent.
- AUTHORS: same format as journal. If organisation is author, use org name in author position.
- INSTITUTION: treated as publisher. Retain full institution name.
  PUBLISHER rule: ALWAYS retain. Omit city only. Strip corporate suffixes only.
  CRITICAL: If the source shows the publisher as the word "Author" (the org is its own publisher),
  store bib_institution = "Author" exactly — do NOT expand to the org name.
- DOI as https://doi.org/XXXXX or full URL. No period after.
  If no DOI and no URL, end with period after institution.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.LEGAL: """
FORMAT (court case):              Case Name, Volume Reporter Page (Court Year).
FORMAT (Indian court case):       Case Name, (Year) Volume Reporter Page.
FORMAT (statute/act, US):         Name of Act, Source § section (Year).
FORMAT (statute/act, India):      Name of Act, Year.
FORMAT (constitution, US):        U.S. Const. art. X, § X.
FORMAT (constitution, India):     India Const. art. X.
FORMAT (bill):                    Title of Bill, H.R./S. Number, Congress (Year).
FORMAT (regulation):              Title or number, Source § section (Year).
FORMAT (legal report/govt doc):   Agency Name. (Year). *Title of report*. URL
FORMAT (court decision + URL):    Case Name, Volume Reporter Page (Court Year). URL
RULES:
- CRITICAL: Legal references do NOT follow the standard Author (Year) format. They use
  jurisdiction-specific citation conventions as shown in the formats above.
- CASE NAME: italicised in APA (e.g. *Brown v. Board of Education*). bib_title = case/statute/bill name.
- REPORTER: Standard abbreviation for the law reporter (e.g. "U.S.", "SCC", "F.2d"). bib_journal = reporter.
- VOLUME: numeric volume or title number. bib_volume = volume.
- PAGE: starting page of the decision/statute. bib_fpage = page or section number.
- COURT: abbreviation or full name of the court. bib_institution = court or legislative body.
- YEAR: year in parentheses at end of citation (court year). bib_year = year.
- SECTION (§): retain the section symbol (§) and number verbatim. Store in bib_fpage.
- URL: for online decisions or government legal documents, append URL after the citation. bib_url = URL.
- GOVERNMENT LEGAL REPORTS: use standard APA report format — Agency. (Year). *Title*. URL.
- No period after URL.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
}


# ─────────────────────────────────────────────
# CGRN RULES  (Chicago Author-Date, 17th ed.)
# ─────────────────────────────────────────────

CGRN_RULES: Dict[str, str] = {
    ReferenceType.JOURNAL: """
FORMAT (standard):         Surname, Firstname M., and Firstname M. Surname. Year. "Article Title." *Journal Name* Vol (Issue): fpage–lpage. https://doi.org/XXXXX
FORMAT (supplement issue): Surname, Firstname M. Year. "Article Title." *Journal Name* Vol (suppl. N): fpage–lpage. https://doi.org/XXXXX
FORMAT (abbreviated name): Surname, Firstname M. Year. "Article Title." ABBREV—*Full Journal Name* Vol (Issue): fpage–lpage. https://doi.org/XXXXX
FORMAT (no DOI, URL):      Surname, Firstname M. Year. "Article Title." *Journal Name* Vol (Issue): fpage–lpage. URL
FORMAT (no DOI, no URL):   Surname, Firstname M. Year. "Article Title." *Journal Name* Vol (Issue): fpage–lpage.
RULES:
- AUTHORS: First author inverted — Lastname, Firstname M. Subsequent authors in normal order
  (Firstname M. Lastname). Joined with ", and" before the last author if two; commas between
  multiple, with "and" before the final one.
  Suffixes (Jr., Sr., II) appended directly after name: e.g. "Thomas A. Callister Jr."
  Abbreviated org author: "MOE (Ministry of Education)." — abbreviation first, full form in parentheses.
  For many authors (≥ 4): first author + ", et al." is acceptable per Chicago style.
  CRITICAL: bib_fname MUST be populated verbatim. Never collapse initials.
- YEAR: Placed after the author block, before the title. Ends with period: "2000."
- ARTICLE TITLE: Title Case AND enclosed in double quotation marks. Ends with period inside the
  closing quotation mark: "Article Title."
- JOURNAL NAME: Title Case AND italicised. Use the name exactly as in the source (abbreviated
  form or full name). Do NOT normalise to NLM abbreviations.
  If the source gives both an abbreviation and full name separated by "—" or "–", retain both:
  "JIPP—*Jurnal Ilmiah Profesi Pendidikan*"
- VOLUME + ISSUE: Vol (Issue) — space before opening parenthesis.
  Supplement: (suppl. N) — lowercase "suppl." with period.
  No issue → omit parentheses.
- PAGE RANGE: colon and space after closing parenthesis: Vol (Issue): fpage–lpage.
  En dash (–) between fpage and lpage. No "pp." prefix.
  Ends with period.
- DOI as full URL: https://doi.org/XXXXX. Ends with period.
  If no DOI but URL exists, use URL. If neither, end after page range.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.BOOK: """
FORMAT (standard):      Surname, Firstname. Year. *Book Title*. Publisher.
FORMAT (with edition):  Surname, Firstname. Year. *Book Title*. Nth ed. Publisher.
FORMAT (translated):    Surname, Firstname. Year. *Book Title*. Translated by Firstname Surname. Publisher.
FORMAT (org author):    Organisation Name. Year. *Book Title*. Publisher.
FORMAT (online):        Organisation Name. Year. *Book Title*. Publisher. URL.
RULES:
- AUTHORS: Same inverted format as journal. "and" before last author.
  Org author used as-is.
- YEAR: After author block, before title. Ends with period.
- BOOK TITLE: Title Case AND italicised. Subtitle separated by colon: *Title: Subtitle*.
- EDITION: "Nth ed." placed after title, before publisher. e.g. "1st ed." "2nd ed."
  Include edition for ALL editions (including 1st) when stated in source.
- TRANSLATED: "Translated by Firstname Surname." placed after title/edition, before publisher.
  If both translated and edited: "Translated and edited by Firstname Surname."
- PUBLISHER: Retain publisher name. Strip corporate suffixes: "Inc.", "Ltd.", "Co.", "Corp.",
  "GmbH", "LLC". City NOT required.
  CRITICAL: No DOI and no web link for pure book references (no online URL unless the source
  is explicitly an online-only publication). Exception: online reports/institutional documents
  may include a URL.
- End with period after publisher (or after URL if included).
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EDITED_BOOK: """
FORMAT (single editor):   Surname, Firstname, ed. Year. *Book Title*. Publisher.
FORMAT (multiple editors): Surname, Firstname, and Firstname Surname, eds. Year. *Book Title*. Publisher.
FORMAT (with edition):    Surname, Firstname, ed. Year. *Book Title*. Nth ed. Publisher.
RULES:
- EDITOR LABEL: ", ed." for single editor; ", eds." for multiple editors — placed after the last
  editor name, before the year. CRITICAL: Never use "ed." for multiple editors.
- EDITOR NAMES: first editor inverted (Lastname, Firstname), additional editors normal order.
  "and" before last editor when 2+. Comma between multiple editors.
- YEAR: After editor block and label, before title. Ends with period.
- BOOK TITLE: Title Case AND italicised.
- EDITION: "Nth ed." after title, before publisher.
- PUBLISHER: same rules as book (strip Inc./Ltd./Co.; no city; no DOI/URL for edited books).
- End with period after publisher.
- CRITICAL: In metadata bib_ed_surname / bib_ed_fname MUST be populated; bib_surname / bib_fname = null.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.BOOK_CHAPTER: """
FORMAT (single editor):    Surname, Firstname. Year. "Chapter Title." In *Book Title*, edited by Firstname Surname, fpage–lpage. Publisher.
FORMAT (multiple editors): Surname, Firstname. Year. "Chapter Title." In *Book Title*, edited by Firstname Surname, Firstname Surname, and Firstname Surname, fpage–lpage. Publisher.
FORMAT (with edition):     Surname, Firstname. Year. "Chapter Title." In *Book Title*, Nth ed., edited by Firstname Surname, fpage–lpage. Publisher.
FORMAT (no page range):    Surname, Firstname. Year. "Chapter Title." In *Book Title*, edited by Firstname Surname. Publisher.
FORMAT (with DOI):         Surname, Firstname. Year. "Chapter Title." In *Book Title*, edited by Firstname Surname, fpage–lpage. Publisher. https://doi.org/XXXXX
RULES:
- CHAPTER AUTHOR: inverted first author + normal subsequent authors.
- YEAR: after author block, before title.
- CHAPTER TITLE: Title Case in double quotation marks. Ends with period inside closing quote.
- "In" (capital I, no colon) introduces the book block.
- BOOK TITLE: Title Case AND italicised.
- EDITION: "Nth ed.," placed after book title, before "edited by" clause.
- "edited by" (lowercase): comma after book title/edition, then "edited by Firstname Surname".
  Two editors: "edited by F. N. Editor and F. N. Editor"
  Three+: "edited by F. N. Editor, F. N. Editor, and F. N. Editor"
  Editor initials in Chicago style appear in normal order (Firstname Surname, NOT inverted).
- PAGE RANGE: en dash (–). Placed after editor(s), before publisher: ", fpage–lpage."
  If no page range, omit. No "pp." prefix in CGRN.
- PUBLISHER: strip Inc./Ltd./Co.; no city.
- DOI or URL: include if available, after publisher. Ends with period.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.WEBSITE: """
FORMAT (with author, date):  Surname, Firstname. Year. "Page Title." Website Name. Accessed Month Day, Year. URL.
FORMAT (org author):         Organisation Name. Year. "Page Title." Website Name. Accessed Month Day, Year. URL.
FORMAT (no author, no date): "Page Title." n.d. Organisation Name. Accessed Month Day, Year. URL.
FORMAT (no author, date):    "Page Title." Year. Website Name. Accessed Month Day, Year. URL.
RULES:
- PAGE TITLE: Title Case in double quotation marks if it has a distinct title.
  If no discrete title, use the page/document description.
- No author → title moves to first position.
- No date → "n.d." in place of year.
- ORGANISATION/WEBSITE NAME: Title Case, not italicised. Placed after title.
- ACCESS DATE: "Accessed Month Day, Year." before the URL.
- URL: full URL. Ends with period.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EREFERENCE: """
FORMAT (with editor):  Surname, Firstname. Year. "Entry Title." In *Reference Title*, edited by Firstname Surname. Publisher. Accessed Month Day, Year. URL.
FORMAT (no editor):    Surname, Firstname. Year. "Entry Title." In *Reference Title*. Publisher. Accessed Month Day, Year. URL.
FORMAT (no date):      Surname, Firstname. n.d. "Entry Title." In *Reference Title*. Publisher. Accessed Month Day, Year. URL.
RULES:
- ENTRY TITLE: Title Case in double quotation marks.
- REFERENCE/DATABASE TITLE: Title Case AND italicised.
- "edited by" (lowercase) with editor in normal order (not inverted).
- ACCESS DATE: Always include "Accessed Month Day, Year." before URL.
- URL: ends with period.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.CONFERENCE: """
FORMAT (paper, presented):   Surname, Firstname. Year. "Paper Title." Presented at Conference Name, Location, Month Day–Day, Year. Publisher. https://doi.org/XXXXX
FORMAT (published proceedings, single ed):  Surname, Firstname. Year. "Paper Title." In *Proceedings Title*, edited by Firstname Surname, fpage–lpage. Publisher. https://doi.org/XXXXX
FORMAT (published proceedings, multiple eds): Surname, Firstname. Year. "Paper Title." In *Proceedings Title*, edited by F. N. Editor, F. N. Editor, and F. N. Editor. Publisher. https://doi.org/XXXXX
FORMAT (poster/abstract):    Surname, Firstname. Year. "Title." Presented at Conference Name, Location, Month Day–Day, Year.
RULES:
- PAPER TITLE: Title Case in double quotation marks.
- "Presented at" introduces conference name, location, and date.
  Format: "Presented at the [Full Conference Name], [City, Country], [Month Day–Day, Year]."
  Retain full conference name including acronym if given in source.
- PUBLISHER (proceedings): listed after date, before DOI.
- PUBLISHED PROCEEDINGS: use book-chapter rules — "In *Proceedings Title*, edited by ..."
  Editor names in normal order (not inverted).
- DOI as full URL. Ends with period.
- CRITICAL FIELD MAPPING:
  bib_title = paper/poster title; bib_conference = full conference name;
  bib_confdate = date range; bib_conflocation = city and country.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.THESIS: """
FORMAT (PhD dissertation):       Surname, Firstname. Year. "Title." PhD diss., University Name.
FORMAT (master's thesis):        Surname, Firstname. Year. "Title." Master's thesis, University Name.
FORMAT (PhD, online):            Surname, Firstname. Year. "Title." PhD diss., University Name. URL.
FORMAT (undergraduate thesis):   Surname, Firstname. Year. "Title." Undergraduate thesis, University Name.
RULES:
- TITLE: Title Case in double quotation marks.
- DEGREE LABEL: "PhD diss.," or "Master's thesis," or "Undergraduate thesis," — comma after label.
- UNIVERSITY: full name. End with period.
- URL: if available online, append URL after period. Ends with period.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.REPORT: """
FORMAT (org author, URL):      Organisation Name. Year. *Report Title*. Publisher. URL.
FORMAT (org author, no URL):   Organisation Name. Year. *Report Title*. Publisher.
FORMAT (govt, with number):    Government Agency. Year. *Report Title*. Report No. XXXX. Publisher.
FORMAT (individual author):    Surname, Firstname. Year. *Report Title*. Publisher. URL.
RULES:
- REPORT TITLE: Title Case AND italicised.
- AUTHORS / ORG: same name rules as book.
- PUBLISHER: retain full name; strip Inc./Ltd./Co.
- REPORT NUMBER: "Report No. XXXX." placed after title, before publisher/URL.
- URL: full URL, ends with period.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.PATENT: """
FORMAT (standard): Inventor, Firstname. Year. "Title of Invention." Country Patent XXXXXXX. Issued Month Day, Year.
FORMAT (with assignee): Inventor, Firstname; Assignee Name. Year. "Title." Country Patent XXXXXXX. Issued Month Day, Year.
FORMAT (online): Inventor, Firstname. Year. "Title." Country Patent XXXXXXX. Issued Month Day, Year. Accessed Month Day, Year. URL.
RULES:
- Inventor name inverted (first inventor), normal order for subsequent inventors.
- Patent number written as: "Country Patent XXXXXXX" (Title Case).
- TITLE: Title Case in quotation marks.
- Issued date written in full.
- Assignee: separated from inventor by semicolon.
- bib_reportnum = full patent designation; bib_assignee = assignee company.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
}


# ─────────────────────────────────────────────
# PROMPT BUILDER
# ─────────────────────────────────────────────

CONVERSION_MAP: Dict[Tuple[CitationStyle, CitationStyle], Tuple[str, str]] = {
    (CitationStyle.AMA,  CitationStyle.APA):  ("AMA 11th Edition",        "APA 7th Edition"),
    (CitationStyle.APA,  CitationStyle.AMA):  ("APA 7th Edition",         "AMA 11th Edition"),
    (CitationStyle.APA,  CitationStyle.APA):  ("APA 7th Edition",         "APA 7th Edition (Strict Formatting Validation)"),
    (CitationStyle.AMA,  CitationStyle.AMA):  ("AMA 11th Edition",        "AMA 11th Edition (Strict Formatting Validation)"),
    # CGRN (Chicago Author-Date) conversions
    (CitationStyle.AMA,  CitationStyle.CGRN): ("AMA 11th Edition",        "CGRN Chicago Author-Date"),
    (CitationStyle.APA,  CitationStyle.CGRN): ("APA 7th Edition",         "CGRN Chicago Author-Date"),
    (CitationStyle.CGRN, CitationStyle.AMA):  ("CGRN Chicago Author-Date","AMA 11th Edition"),
    (CitationStyle.CGRN, CitationStyle.APA):  ("CGRN Chicago Author-Date","APA 7th Edition"),
    (CitationStyle.CGRN, CitationStyle.CGRN): ("CGRN Chicago Author-Date","CGRN Chicago Author-Date (Strict Formatting Validation)"),
}


def _build_prompt(
    raw_text: str,
    source_style: CitationStyle,
    target_style: CitationStyle,
    cr_item: Optional[Dict[str, Any]] = None,
) -> str:
    source_label, target_label = CONVERSION_MAP[(source_style, target_style)]
    if target_style == CitationStyle.APA:
        rules_map = APA_RULES
    elif target_style == CitationStyle.CGRN:
        rules_map = CGRN_RULES
    else:
        rules_map = AMA_RULES
    rules_block = "\n".join(
        f"### {ref_type.upper()}\n{rules}"
        for ref_type, rules in rules_map.items()
    )

    cr_block = ""
    if cr_item:
        cr_block = f"""
## VERIFIED DATABASE MATCH (CrossRef/PubMed)
The following metadata was securely matched for this reference in scientific databases.
You MUST prioritize this database content for updating missing elements, ensuring exact
journal names (abbreviation/capitalisation), and correcting DOIs.
CRITICAL AUTHOR RULE: DO NOT replace the original author list from the input reference
with the database authors if the database authors are generic (e.g. "NA", "&NA;",
"Anonymous"), heavily abbreviated, or missing. HOWEVER, if the original input reference
uses "et al." or "and others" AND the database provides the full author list, you MUST
expand the author list using the database authors to properly format it according to
the target style's rules (e.g. up to 20 authors for APA). Otherwise, prioritise keeping
the specific authors provided in the input string and use the database authors only to
correct minor spelling mistakes.
CRITICAL TITLE RULE: DO NOT overwrite the article/chapter/book title from the input
with the database title unless the input title is completely absent.
{json.dumps(cr_item, indent=2)}
"""

    return f"""You are a professional bibliographic reference conversion expert specialising in {source_label} to {target_label} conversion.

## YOUR TASK
1. Detect the reference type (journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report, patent, legal).
2. Extract ALL available metadata into the bib_ fields. If database match metadata is provided below, incorporate it completely to fix/complete the reference (add DOIs, fix capitals, expand abbreviations).
3. Reformat the reference strictly according to {target_label} rules for the detected type.
4. Note any missing data, assumptions made, or issues in "conversion_notes".

## STRICT EXTRACTION RULES
- bib_reftype: one of: journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report, patent, legal, unknown
  BOOK vs EDITED_BOOK: classify as edited_book ONLY when editor markers ("Ed.", "Eds.", "ed.",
  "eds.", "edited by") are present AND no separate author appears before the title.
  If both author + editor present → book_chapter. Authors only, no editor markers → book.
  WEBSITE vs BOOK vs REPORT — CRITICAL decision rules:
    BOOK: Has a title + author(s)/org + edition number (e.g. "2nd ed.", "3rd ed.") OR a named
      publisher. A URL does not change this — a book hosted online is still a book.
      CRITICAL: If the source contains an edition number → classify as book, NEVER as website.
    REPORT: Published by a government body, institution, or organisation; often has a document/
      report number; may have a URL. Classify as report when there is no edition number and the
      source reads like an institutional report, white paper, or policy document.
    WEBSITE: A page or article on a website where the website itself is the primary identifier
      (e.g. Wikipedia entry, WHO web page, blog post, news article). Classify as website ONLY
      when the source is clearly a web page, NOT a book or formal institutional document.
      CRITICAL: "Publisher: Author" or publisher name = org author name → book or report, NOT website.
    PATENT: Source describes an invention with a patent or patent application number issued by a
      national patent office (USPTO, EPO, WIPO, IPO, etc.). Classify as patent regardless of URL.
      bib_reportnum = full patent designation; bib_assignee = assignee company (if any);
      bib_surname/bib_fname = inventors; bib_title = title of invention.
    LEGAL: Source is a court case citation, statute/act, regulation, constitutional provision, or
      bill/proposed law. Does NOT follow Author–Year format.
      bib_title = case/statute/bill name; bib_journal = reporter abbreviation;
      bib_volume = volume/title number; bib_fpage = page or section number;
      bib_institution = court or legislative body; bib_year = year.
- bib_surname / bib_fname: ALL authors in order, pipe-separated (|) if multiple.
  e.g. "Smith|Jones|Lee" / "John A|Mary B|Chris"
  CRITICAL: For metadata, you MUST extract ALL authors present in the source text. If the source text ends with "et al." or "...", you MUST include "et al" as the final author in `bib_surname` (with a blank `bib_fname`). Never truncate the author list yourself.
  CRITICAL: bib_fname MUST be populated verbatim from source. Never delete, suppress, or collapse initials.
  If source has "JA", store "JA". If source has "J.A.", store "J.A.". If source has "John A", store "John A".
  Never truncate multiple initials into one (e.g. "JA" must stay "JA", never become "J").
  CRITICAL: Retain name generation suffixes (Jr., Sr., II, III, IV) in bib_fname, comma-separated
  after the initials. e.g. bib_fname = "J. W., Jr." for "Collins, Jr., J. W."
  CRITICAL: If the author list includes a named collaborative group, writing committee, study group,
  or task force (e.g. "ACCORD Study Group"), include it as a pipe-separated entry in bib_surname
  with bib_fname = "" for that entry.
  e.g. bib_surname = "Smith|Jones|ACCORD Study Group", bib_fname = "J. A.|B. C.|"
- bib_ed_surname / bib_ed_fname: same pipe-separated format for editors.
  For edited_book: these MUST be populated; bib_surname / bib_fname = null.
- bib_year: Extract the FULL year string verbatim including any letter suffix (2022a, 2024b) and/or
  full date. For APA, normalise order to "Year, Month Day" (e.g. "2002a, April 30").
  CRITICAL: NEVER strip or add a letter suffix — it is required APA disambiguation. If the source
  has "2022a", bib_year MUST be "2022a". If source has "2022" with no suffix, do NOT add one.
  For APA website/conference refs with a full date in the source, store the COMPLETE date as
  "YYYY, Month Day". Never truncate to year-only when a full date is available.
- bib_accessed: the ACCESS/RETRIEVAL date only (format "Month DD, YYYY"). NEVER put the
  publication year in bib_accessed, and NEVER put the access date in bib_year. These are always
  separate fields.
- bib_title: For website and ereference types — this is the PAGE or ENTRY title ONLY.
  NOT the website name, database name, platform name, or URL. The website/database name goes
  in bib_journal or bib_book.
- bib_conference: For conference refs — the FULL conference/symposium name ONLY.
  NEVER put the paper/poster title in bib_conference.
- bib_confdate: For conference refs — the DATE RANGE of the conference ONLY (e.g. "May 28–31, 2024").
  NEVER put a standalone year here; year goes in bib_year.
- bib_volume / bib_issue: numeric string only, no labels
- bib_fpage / bib_lpage: digits only (or article ID like e13284), no "pp.", "p.", or labels.
  CRITICAL: If source has only one page number, set bib_lpage = null — NEVER copy bib_fpage into bib_lpage.
  NEVER fabricate a last page.
- bib_doi: raw DOI string only — strip "https://doi.org/" prefix
- bib_url: full URL only, no trailing period
- bib_editionno: number only (e.g., "2", "3")
- bib_deg: full degree name (e.g., "Doctoral dissertation", "Master's thesis")
- bib_assignee: for patent refs only — the company/organisation holding the patent. Null for all other types.
- bib_comment: always return null — this field is populated externally by the validation layer.
- All other string fields: extract verbatim from source
- Return null for any field not present in the source — NEVER fabricate data
- COMPLETENESS: The formatted_output must be the FULL reference — never truncate, omit, abbreviate,
  or paraphrase any element.
  CRITICAL: Do NOT overwrite source titles with database titles unless source title is completely absent.

## CRITICAL ANTI-FABRICATION RULES
NEVER, under any circumstance, fabricate, hallucinate, or invent any metadata elements:
- NEVER invent author names that do NOT appear in the source.
  If the source says "et al.", list only the authors explicitly shown in the source text,
  then add "et al" as the final author entry (bib_fname = "").
- NEVER invent DOIs. If a DOI is not present in the source text, leave bib_doi = null.
  Do NOT guess, look up, or fabricate a DOI. Only include DOIs explicitly provided in the source.
- NEVER invent page numbers, volumes, issues, or journal names. If they are missing from the source, leave the field null.
- NEVER change a reference type without explicit evidence in the source text.
  If the source says "book", do not reclassify it as "report" or "website" just because it might have a URL.
- NEVER use fields populated in the database block to replace the source if the source is clear
  and unambiguous.  The database block is ONLY for filling null fields and fixing errors.
- NEVER overwrite the explicit author list with a generic database author list (like "Anonymous" or "NA").
- If you are unsure about a field, leave it null — this is better than fabricating data.

## {target_label} FORMATTING RULES BY REFERENCE TYPE
{rules_block}
{cr_block}
## INPUT REFERENCE ({source_label})
{raw_text}

## OUTPUT
Return valid JSON matching the required schema exactly.
"""


def _build_validation_prompt(raw_text: str, target_style: CitationStyle) -> str:
    if target_style == CitationStyle.APA:
        target_label = "APA 7th Edition"
        rules_map = APA_RULES
    elif target_style == CitationStyle.CGRN:
        target_label = "CGRN Chicago Author-Date"
        rules_map = CGRN_RULES
    else:
        target_label = "AMA 11th Edition"
        rules_map = AMA_RULES
    rules_block = "\n".join(
        f"### {ref_type.upper()}\n{rules}"
        for ref_type, rules in rules_map.items()
    )

    return f"""You are a professional bibliographic reference validation expert specialising in {target_label}.

## YOUR TASK
Given an input reference, validate whether it perfectly adheres to {target_label} guidelines.
1. Detect the reference type.
2. Check if the reference perfectly matches the {target_label} rules for the detected type.
3. If there are any formatting issues, punctuation errors, missing required elements, or style
   deviations, set 'is_valid' to false and list the issues in 'validation_errors'.
4. Provide the fully corrected reference in 'corrected_reference'.
5. Extract ALL available metadata into the bib_ fields.

## STRICT EXTRACTION RULES
- bib_reftype: one of: journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report, patent, legal, unknown
  PATENT: Has a patent/application number from a national patent office (USPTO, EPO, WIPO, etc.).
    bib_reportnum = full patent designation; bib_assignee = assignee company;
    bib_surname/bib_fname = inventors; bib_title = title of invention.
  LEGAL: Court case, statute/act, regulation, constitutional provision, or bill. No Author–Year format.
    bib_title = case/statute/bill name; bib_journal = reporter abbreviation;
    bib_volume = volume/title number; bib_fpage = page or section number;
    bib_institution = court or legislative body; bib_year = year.
- bib_surname / bib_fname: ALL authors/inventors in order, pipe-separated (|) if multiple.
  CRITICAL: bib_fname MUST be populated verbatim from source. Never collapse initials.
  Retain generation suffixes (Jr., Sr., II, III) in bib_fname comma-separated after initials.
- bib_ed_surname / bib_ed_fname: same pipe-separated format for editors.
- bib_year: Extract FULL year string including any letter suffix. NEVER strip it.
- bib_volume / bib_issue: numeric string only.
- bib_fpage / bib_lpage: digits only. If only one page, bib_lpage = null.
- bib_doi: raw DOI string only — strip "https://doi.org/" prefix.
- bib_url: full URL only, no trailing period.
- bib_accessed: date in "Month DD, YYYY" format. NEVER confuse with bib_year.
- bib_editionno: number only.
- bib_deg: full degree name.
- bib_title: for website/ereference = page/entry title only, NOT the site/database name.
- bib_assignee: patent assignee company name only (patent refs). Null for all other types.
- bib_comment: leave null — populated externally by the validation layer, not by this prompt.
- Return null for any field not present — NEVER fabricate data.
- JOURNAL NAME: NEVER append parenthetical location qualifiers unless verbatim in source.
- TITLES: NEVER overwrite source titles with database titles unless completely absent.
- PUBLISHER: ALWAYS retain. Strip only corporate suffixes. Omit city/location only.
- COMPLETENESS: corrected_reference must be the FULL reference — never truncate.

## {target_label} FORMATTING RULES BY REFERENCE TYPE
{rules_block}

## INPUT REFERENCE
{raw_text}

## OUTPUT
Return valid JSON matching the required validation schema exactly.
"""


# ─────────────────────────────────────────────
# TEXT CLEANUP HELPERS
# ─────────────────────────────────────────────

_DQUOTE_OPEN  = re.compile(r'(^|[\s(\[{])"')
_SQUOTE_OPEN  = re.compile(r"(^|[\s(\[{])'")

def _to_smart_quotes(text: str) -> str:
    """Convert straight quotes to typographic (curly) quotes."""
    if not text:
        return text
    text = _DQUOTE_OPEN.sub(r'\1\u201c', text)
    text = text.replace('"', '\u201d')
    text = _SQUOTE_OPEN.sub(r'\1\u2018', text)
    text = text.replace("'", '\u2019')
    return text


_NBSP            = re.compile(r'\u00a0')
_MULTI_SPACE     = re.compile(r' {2,}')
_DOT_SPACE_DOT   = re.compile(r'(?<![A-Z])\.\s+\.')
_SPACE_BEFORE_PUNCT = re.compile(r'(?<![A-Z](?:\.\s))\s+([.,;:])')
_DOUBLE_DOT      = re.compile(r'(?<![A-Z])\.\.+')


def _clean_formatted_output(text: str, smart_quotes: bool = False) -> str:
    """
    Deterministic cleanup of common LLM punctuation artefacts.

    smart_quotes=False (default): leave quotes as-is. Word AutoCorrect manages
    smart quotes itself; inserting pre-curled quotes via python-docx creates
    unnecessary track-changes noise against straight-quote originals.
    """
    if not text:
        return text
    text = _NBSP.sub(' ', text)
    text = _MULTI_SPACE.sub(' ', text)
    text = _DOT_SPACE_DOT.sub('.', text)
    text = _SPACE_BEFORE_PUNCT.sub(r'\1', text)
    text = _DOUBLE_DOT.sub('.', text)
    if smart_quotes:
        text = _to_smart_quotes(text)
    return text.strip()


# ─────────────────────────────────────────────
# INTERNAL API CALL HELPER  (with retry)
# ─────────────────────────────────────────────

def _call_gemini(
    prompt: str,
    schema,
    model_name: str,
    api_key: str,
    max_retries: int = _MAX_RETRIES,
) -> Optional[str]:
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    config = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=schema,
        temperature=0.0,
        top_p=1.0,
        max_output_tokens=8192,
    )

    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
            _acquire_api_slot()
            response = client.models.generate_content(
                model=model_name,
                contents=prompt,
                config=config,
            )

            if not response or not response.candidates:
                logger.warning(f"[Attempt {attempt}] No candidates in Gemini response.")
                last_exc = ValueError("No candidates returned")
                _backoff(attempt, max_retries)
                continue

            candidate = response.candidates[0]
            finish_reason = candidate.finish_reason

            if finish_reason not in (
                types.FinishReason.STOP,
                types.FinishReason.MAX_TOKENS,
            ):
                logger.warning(f"[Attempt {attempt}] Unexpected finish reason: {finish_reason}")
                last_exc = ValueError(f"Finish reason: {finish_reason}")
                _backoff(attempt, max_retries)
                continue

            if finish_reason == types.FinishReason.MAX_TOKENS:
                logger.warning(
                    f"[Attempt {attempt}] Gemini hit MAX_TOKENS — response may be truncated. "
                    "Consider increasing max_output_tokens or shortening the input."
                )

            raw_json = response.text
            if not raw_json or not raw_json.strip():
                raw_preview = (response.text or "")[:300]
                logger.warning(f"[Attempt {attempt}] Empty response text. Raw: {raw_preview!r}")
                last_exc = ValueError("Empty response text")
                _backoff(attempt, max_retries)
                continue

            try:
                from utils.gemini_cost_tracker import log_usage as _log_gemini
                _um = response.usage_metadata
                if _um:
                    _log_gemini(
                        "reference_conversion", model_name,
                        getattr(_um, "prompt_token_count", 0) or 0,
                        getattr(_um, "candidates_token_count", 0) or 0,
                    )
            except Exception:
                pass

            return raw_json

        except Exception as exc:
            last_exc = exc
            err_str = str(exc).lower()
            is_rate_limit = any(kw in err_str for kw in ("429", "quota", "resource_exhausted"))
            if is_rate_limit or any(kw in err_str for kw in ("rate", "503", "timeout", "unavailable")):
                logger.warning(f"[Attempt {attempt}] Transient error: {exc}. Retrying…")
                _backoff(attempt, max_retries, is_rate_limit=is_rate_limit)
            else:
                logger.error(f"[Attempt {attempt}] Non-retriable error: {exc}")
                break

    error_type = "Unknown"
    if last_exc:
        err_str = str(last_exc).lower()
        if any(k in err_str for k in ("429", "quota", "resource_exhausted", "rate")):
            error_type = "Rate Limit / Quota Exceeded"
        elif any(k in err_str for k in ("timeout", "timed out")):
            error_type = "Timeout"
        elif any(k in err_str for k in ("503", "unavailable", "service_unavailable")):
            error_type = "Service Unavailable"
        elif any(k in err_str for k in ("401", "unauthenticated", "invalid_api_key")):
            error_type = "API Key / Authentication Error"
        elif any(k in err_str for k in ("400", "badrequest", "invalid")):
            error_type = "Invalid Request"
        elif "no candidates" in err_str or "finish reason" in err_str:
            error_type = "Gemini Response Error"
        elif "empty response" in err_str:
            error_type = "Empty Response"
        else:
            error_type = f"Error: {str(last_exc)[:100]}"
    logger.error(f"All {max_retries} Gemini attempts failed. Type: {error_type} | Details: {last_exc}")
    return None


def _backoff(attempt: int, max_retries: int, is_rate_limit: bool = False) -> None:
    import random
    if attempt < max_retries:
        delay = _RETRY_BASE_DELAY * (2 ** (attempt - 1)) + random.uniform(0, 2.0)
        if is_rate_limit:
            # 429 quota resets on a per-minute basis — floor at 30 s
            delay = max(delay, 30.0)
        logger.warning(f"Back-off: sleeping {delay:.1f}s before retry {attempt + 1}/{max_retries}.")
        time.sleep(delay)


def _resolve_api_key() -> Optional[str]:
    return (
        os.environ.get("REFERENCE_CONVERTER_GEMINI_API_KEY")
        or os.environ.get("GEMINI_API_KEY")
        or os.environ.get("GOOGLE_API_KEY")
    )


# ─────────────────────────────────────────────
# EARLY METADATA EXTRACTION (for DOI lookup before Gemini)
# ─────────────────────────────────────────────

def _extract_metadata_from_raw_text(raw_text: str) -> "Dict[str, Any]":
    """
    Quick extraction of title, authors, year, journal from raw reference text.
    Parses AMA, APA, and formatted references without Gemini.

    Returns dict with keys: title, authors (list), year, journal
    """
    metadata: Dict[str, Any] = {
        "title": "",
        "authors": [],
        "year": "",
        "journal": ""
    }

    try:
        # Remove reference number prefix if present (e.g., "1. Author..." → "Author...")
        text = re.sub(r'^\s*\d+[\.\)]\s+', '', raw_text.strip())

        # Extract year (4 digits, often in formats like "2022;", "(2022)", "2022.")
        year_match = re.search(r'\b(19|20)\d{2}\b', text)
        if year_match:
            metadata["year"] = year_match.group(0)

        # Extract authors from start (before title or first sentence)
        # AMA format: "Author1 AB, Author2 CD, et al. Title..."
        # APA format: "Author, A. B., & Author, C. D. (2022). Title..."
        author_pattern = r'^([A-Z][a-z\'-]*(?:\s+[a-z]{1,4})*\s+[A-Z\.]+(?:,\s*(?:[A-Z][a-z\'-]*(?:\s+[a-z]{1,4})*\s+[A-Z\.]+ *(?:&|\,)?)*)*)'
        author_match = re.match(author_pattern, text)
        if author_match:
            author_section = author_match.group(1)
            # Split by comma or & to get individual authors
            author_parts = re.split(r',\s*|\s+&\s+', author_section)
            metadata["authors"] = [a.strip() for a in author_parts if a.strip() and "et al" not in a.lower()]

        # Extract title: text between authors and journal/year
        # Look for capitalized text before journal name or year
        title_pattern = r'(?:et al\.?\s+)?([A-Z][^\.]*?[a-z\?\!]+[\.\?\!]?)\s+(?:[A-Z][a-z]+(?:\s+[A-Z][a-z]+)?\.?\s+\d{4}|[A-Z][a-z]+\s+\d{4}|[A-Z][a-z]+\.?\s*\d{4})'
        title_match = re.search(title_pattern, text)
        if title_match:
            metadata["title"] = title_match.group(1).strip()
        else:
            # Fallback: take text up to first 4-digit year or known journal patterns
            text_before_year = text
            if year_match:
                text_before_year = text[:year_match.start()]
            # Remove author section to isolate title
            if author_match:
                text_before_year = text_before_year[len(author_match.group(0)):].strip()
            # Clean up and extract meaningful part
            metadata["title"] = re.sub(r'^[\.\,\s]+', '', text_before_year)[:200].strip()

        # Extract journal name: usually capitalized, before volume/pages
        # Common patterns: "Journal Name. Year" or "Journal Name Year;Volume"
        journal_pattern = r'([A-Z][A-Za-z\s&\-]*?)\.?\s*(?:\d{4}|[Vv]ol\.?|;|\()'
        journal_match = re.search(journal_pattern, text)
        if journal_match:
            metadata["journal"] = journal_match.group(1).strip()

        logger.debug(f"  [Metadata Extract] Found: {len(metadata['authors'])} authors, title={metadata['title'][:50] if metadata['title'] else 'N/A'}")

    except Exception as e:
        logger.debug(f"  [Metadata Extract] Error: {e}")

    return metadata


# ─────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────

def convert_reference(
    raw_text: str,
    source_style: CitationStyle,
    target_style: CitationStyle,
    model_name: str = DEFAULT_MODEL,
    cr_item: Optional[Dict[str, Any]] = None,
    validate: bool = False,
) -> Optional[Dict[str, Any]]:
    if not raw_text or not raw_text.strip():
        logger.error("raw_text is empty.")
        return None

    key = (source_style, target_style)
    if key not in CONVERSION_MAP:
        logger.error(f"Unsupported conversion: {source_style} → {target_style}")
        return None

    api_key = _resolve_api_key()
    if not api_key:
        logger.error(
            "API key not found (checked REFERENCE_CONVERTER_GEMINI_API_KEY, "
            "GEMINI_API_KEY, and GOOGLE_API_KEY)."
        )
        return None

    # ─────────────────────────────────────────────
    # [EARLY] Try DOI lookup BEFORE Gemini call
    # This allows us to enrich authors in a single pass
    # ─────────────────────────────────────────────
    early_enriched_authors = None
    if not cr_item:  # Only do early lookup if no cr_item already provided
        try:
            early_metadata = _extract_metadata_from_raw_text(raw_text)
            if early_metadata.get("title") and early_metadata.get("year"):
                logger.info(f"  [Early DOI Lookup] Extracted metadata from raw text")

                # Try PubMed first
                logger.info(f"  [Early DOI Lookup] Trying PubMed...")
                pubmed_result = _lookup_doi_from_pubmed(
                    title=early_metadata.get("title"),
                    authors=early_metadata.get("authors"),
                    year=early_metadata.get("year"),
                    journal=early_metadata.get("journal")
                )
                if pubmed_result:
                    doi = pubmed_result.get("doi")
                    early_enriched_authors = pubmed_result.get("authors")
                    logger.info(f"  [Early DOI Lookup] Found {len(early_enriched_authors) if early_enriched_authors else 0} authors via PubMed")

                # If PubMed failed, try CrossRef
                if not early_enriched_authors:
                    logger.info(f"  [Early DOI Lookup] Trying CrossRef...")
                    crossref_result = _lookup_doi_from_crossref(
                        title=early_metadata.get("title"),
                        authors=early_metadata.get("authors"),
                        year=early_metadata.get("year"),
                        journal=early_metadata.get("journal")
                    )
                    if crossref_result:
                        doi = crossref_result.get("doi")
                        early_enriched_authors = crossref_result.get("authors")
                        logger.info(f"  [Early DOI Lookup] Found {len(early_enriched_authors) if early_enriched_authors else 0} authors via CrossRef")

                # If we found enriched authors, pass them to Gemini
                if early_enriched_authors and not cr_item:
                    cr_item = {
                        "author": early_enriched_authors,
                        "DOI": doi if doi else None,
                        "_source": "early_doi_lookup"
                    }
                    logger.info(f"  [Early DOI Lookup] Passing enriched authors to Gemini prompt")
        except Exception as e:
            logger.debug(f"  [Early DOI Lookup] Exception (non-blocking): {e}")
            # If early lookup fails, continue without enrichment - Gemini will process normally

    schema   = _get_response_schema()
    prompt   = _build_prompt(raw_text, source_style, target_style, cr_item)
    raw_json = _call_gemini(prompt, schema, model_name, api_key)

    if raw_json is None:
        return None

    try:
        parsed: Dict[str, Any] = json.loads(raw_json)
    except json.JSONDecodeError as exc:
        logger.error(f"JSON decode error: {exc}")
        return None

    if "formatted_output" not in parsed or "metadata" not in parsed:
        logger.error(f"Missing top-level keys in response: {list(parsed.keys())}")
        return None

    if not isinstance(parsed["formatted_output"], str) or not parsed["formatted_output"].strip():
        logger.error("formatted_output is empty or not a string.")
        return None

    # smart_quotes=False: Word manages its own smart quotes; pre-curling causes
    # track-changes noise against straight-quote originals (#35)
    parsed["formatted_output"] = _clean_formatted_output(
        parsed["formatted_output"], smart_quotes=False
    )

    meta = parsed.get("metadata", {})
    for field in BIB_FIELDS:
        meta.setdefault(field, None)
    parsed["metadata"] = meta

    # Validate that metadata wasn't fabricated
    if not _validate_metadata_not_fabricated(meta, raw_text):
        logger.warning(f"  [WARNING] Metadata validation failed — possible fabrication detected")
        # Don't reject outright, but log the concern
    
    # For journal references, try DOI lookup if missing (fallback to early lookup or if early failed)
    ref_type = meta.get("bib_reftype", "unknown")
    if cr_item and cr_item.get("_source") == "early_doi_lookup" and cr_item.get("DOI"):
        meta["bib_doi"] = str(cr_item["DOI"]).strip()
    if ref_type == "journal" and not meta.get("bib_doi"):
        # Skip fallback DOI lookup if we already did early enrichment
        if cr_item and cr_item.get("_source") == "early_doi_lookup":
            logger.info(f"  [DOI Lookup] Skipping fallback - early DOI lookup already enriched authors")
        else:
            logger.info(f"  [DOI Lookup] Fallback DOI lookup for journal reference (early lookup didn't find it)...")
            surnames = (meta.get("bib_surname") or "").split("|") if meta.get("bib_surname") else []

            doi = None

            # Try PubMed first (better for medical/biomedical literature)
            logger.info(f"  [DOI Lookup] Trying PubMed...")
            pubmed_result = _lookup_doi_from_pubmed(
                title=meta.get("bib_title"),
                authors=surnames,
                year=meta.get("bib_year"),
                journal=meta.get("bib_journal")
            )
            if pubmed_result:
                doi = pubmed_result.get("doi")

            # If PubMed failed, try CrossRef
            if not doi:
                logger.info(f"  [DOI Lookup] Trying CrossRef...")
                crossref_result = _lookup_doi_from_crossref(
                    title=meta.get("bib_title"),
                    authors=surnames,
                    year=meta.get("bib_year"),
                    journal=meta.get("bib_journal"),
                    volume=meta.get("bib_volume"),
                    issue=meta.get("bib_issue"),
                    fpage=meta.get("bib_fpage")
                )
                if crossref_result:
                    doi = crossref_result.get("doi")

            if doi:
                logger.info(f"  [DOI Lookup] Found DOI: {doi}")
                meta["bib_doi"] = doi
                # Inject DOI into formatted output if missing
                if ref_type == "journal":
                    if "doi:" not in parsed["formatted_output"].lower() and "https://doi.org" not in parsed["formatted_output"]:
                        # Append DOI to end before any period
                        fmt_out = parsed["formatted_output"].rstrip(". ")
                        if fmt_out.endswith("."):
                            fmt_out = fmt_out[:-1]
                        # Use correct DOI format based on target style
                        doi_str = f"https://doi.org/{doi}" if target_style == CitationStyle.APA else f"doi:{doi}"
                        parsed["formatted_output"] = f"{fmt_out}. {doi_str}"
                        logger.info(f"  [DOI Injection] Injected DOI into formatted output")
            else:
                logger.debug(f"  [DOI Lookup] No DOI found via PubMed or CrossRef")
    
    source_lbl, target_lbl = CONVERSION_MAP[key]
    logger.info(f"Converted [{ref_type}] {source_lbl} → {target_lbl}")
    if parsed.get("conversion_notes"):
        logger.warning(f"Conversion notes: {parsed['conversion_notes']}")

    if validate:
        if target_style == CitationStyle.AMA:
            style_label = "AMA"
        elif target_style == CitationStyle.CGRN:
            style_label = "CGRN"
        else:
            style_label = "APA"
        val = validate_reference(parsed["formatted_output"], target_style, model_name)
        if val is not None:
            parsed["is_valid"]            = val.get("is_valid", False)
            parsed["validation_errors"]   = val.get("validation_errors", [])
            parsed["corrected_reference"] = val.get("corrected_reference", "")
            errors = val.get("validation_errors") or []
            if errors:
                comment_text = (
                    f"[{style_label} validation: "
                    + "; ".join(errors)
                    + "]"
                )
                parsed["metadata"]["bib_comment"] = comment_text
                logger.warning(f"Validation issues [{ref_type}]: {comment_text}")
            else:
                parsed["metadata"]["bib_comment"] = None
        else:
            parsed["is_valid"]            = None
            parsed["validation_errors"]   = []
            parsed["corrected_reference"] = ""
            logger.warning("Post-conversion validation call failed.")

    return parsed


def validate_reference(
    raw_text: str,
    target_style: CitationStyle,
    model_name: str = DEFAULT_MODEL,
) -> Optional[Dict[str, Any]]:
    if not raw_text or not raw_text.strip():
        logger.error("raw_text is empty.")
        return None

    api_key = _resolve_api_key()
    if not api_key:
        logger.error(
            "API key not found (checked REFERENCE_CONVERTER_GEMINI_API_KEY, "
            "GEMINI_API_KEY, and GOOGLE_API_KEY)."
        )
        return None

    schema   = _get_validation_schema()
    prompt   = _build_validation_prompt(raw_text, target_style)
    raw_json = _call_gemini(prompt, schema, model_name, api_key)

    if raw_json is None:
        return None

    try:
        parsed: Dict[str, Any] = json.loads(raw_json)
    except json.JSONDecodeError as exc:
        logger.error(f"JSON decode error during validation: {exc}")
        return None

    required = {"is_valid", "validation_errors", "corrected_reference"}
    if not required.issubset(parsed.keys()):
        logger.error(f"Missing keys in validation response: {list(parsed.keys())}")
        return None

    parsed["corrected_reference"] = _clean_formatted_output(
        parsed["corrected_reference"], smart_quotes=False
    )

    meta = parsed.get("metadata", {})
    for field in BIB_FIELDS:
        meta.setdefault(field, None)
    parsed["metadata"] = meta

    ref_type   = meta.get("bib_reftype", "unknown")
    if target_style == CitationStyle.APA:
        target_lbl = "APA 7th"
    elif target_style == CitationStyle.CGRN:
        target_lbl = "CGRN Chicago"
    else:
        target_lbl = "AMA 11th"
    status     = "VALID" if parsed["is_valid"] else "INVALID"
    logger.info(f"Validated [{ref_type}] against {target_lbl}: {status}")

    return parsed


# ─────────────────────────────────────────────
# BATCH CONVERTER
# ─────────────────────────────────────────────

def convert_references_batch(
    references: List[str],
    source_style: CitationStyle,
    target_style: CitationStyle,
    model_name: str = DEFAULT_MODEL,
    cr_items: Optional[List[Optional[Dict[str, Any]]]] = None,
    validate: bool = False,
) -> List[Optional[Dict[str, Any]]]:
    if cr_items and len(cr_items) != len(references):
        raise ValueError(
            f"cr_items length ({len(cr_items)}) must match references length ({len(references)})."
        )

    results: List[Optional[Dict[str, Any]]] = []
    for i, ref in enumerate(references):
        logger.info(f"Processing reference {i + 1}/{len(references)}")
        cr = cr_items[i] if cr_items else None
        result = convert_reference(ref, source_style, target_style, model_name, cr_item=cr, validate=validate)
        results.append(result)
    return results

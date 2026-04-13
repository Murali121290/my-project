"""
gemini_ref_converter.py
-----------------------
Converts and validates bibliographic references between AMA 11th and APA 7th
edition using the Google Gemini API.
"""

from __future__ import annotations

import json
import logging
import os
import re
import time
import functools
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
DEFAULT_MODEL = "gemini-2.5-pro"
_MAX_RETRIES = 3
_RETRY_BASE_DELAY = 2.0  # seconds


# ─────────────────────────────────────────────
# ENUMS
# ─────────────────────────────────────────────

class CitationStyle(str, Enum):
    AMA = "AMA"
    APA = "APA"


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
RULES:
- AUTHORS: Last name followed by space then initials with NO periods or spaces between initials
  (e.g. Smith JA, Jones BC). Comma-space between authors.
  Up to 6 authors list all; if 7 or more, list first 6 then ", et al." (WITH period after "al").
  CRITICAL: Retain EVERY initial exactly as in the source. Never collapse "JA" to "J" or drop
  any initial. bib_fname MUST be populated.
  CRITICAL: Do NOT remove study groups or steering committees (e.g., "International Steering Committee for...") attached to the author list. Treat them as part of the author group.
  No author → start with article title directly.
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
- DOI: prefix with "doi:" (lowercase, no space). No period after doi string.
  If no DOI available, provide full URL. No period after URL.
- Strip any non-breaking spaces (U+00A0) silently — never output a dot in their place.
- COMPLETENESS: Output the FULL reference. Never truncate, omit, or paraphrase any element.
""",
    ReferenceType.BOOK: """
FORMAT (with edition):    Surname FM, Surname FM. Title of book. Xth ed. Publisher; Year.
FORMAT (1st/only edition): Surname FM, Surname FM. Title of book. Publisher; Year.
FORMAT (with DOI/URL):    Surname FM. Title of book. Publisher; Year. doi:XXXXX
FORMAT (org author):      Organisation Name. Title of book. Publisher; Year.
RULES:
- AUTHORS: Retain every initial. Same format rules as journal (no periods in initials).
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
- EDITOR NAMES: same initials format as author names (no periods between initials).
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
- CHAPTER AUTHOR: comes first. Retain every initial. Same format as journal authors.
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
FORMAT (with author):    Author FM. Title of page. Website Name. Published Month Day, Year. Accessed Month Day, Year. URL
FORMAT (updated):        Author FM. Title of page. Website Name. Updated Month Day, Year. Accessed Month Day, Year. URL
FORMAT (no author):      Title of page. Website Name. Published Month Day, Year. Accessed Month Day, Year. URL
FORMAT (no pub date):    Author FM. Title of page. Website Name. Accessed Month Day, Year. URL
RULES:
- PAGE/DOCUMENT TITLE: sentence case (first word and proper nouns only). No italics.
- WEBSITE NAME: title case. Retain as-is from source.
- AUTHOR: If no personal author, start directly with the page title.
- PUBLICATION DATE: Use "Published Month Day, Year" or "Updated Month Day, Year" as appropriate.
  If the source specifies "Updated", use "Updated"; if "Published", use "Published".
  If no publication date is available, omit the date line entirely.
- ACCESS DATE: Always include "Accessed Month Day, Year." before the URL.
- URL: on same line after access date. No period after URL.
- No DOI for websites.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EREFERENCE: """
FORMAT (with editor, single):   Author FM. Entry title. In: Editor FM, ed. Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (with editors, multiple): Author FM. Entry title. In: Editor FM, Editor FM, eds. Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (no editor):              Author FM. Entry title. In: Reference Title. Publisher; Year. Accessed Month Day, Year. URL
FORMAT (no year):                Author FM. Entry title. In: Editor FM, ed. Reference Title. Publisher. Accessed Month Day, Year. URL
RULES:
- ENTRY TITLE: sentence case. No italics.
- REFERENCE BOOK/DATABASE TITLE: title case, italicise if possible.
- EDITOR LABEL: "ed." for single editor, "eds." for two or more editors.
  CRITICAL: Never use "ed." when there are multiple editors.
  If no editor listed, omit editor block entirely.
- PUBLISHER: Retain name; strip corporate suffixes.
- YEAR: Include if available. If no year, omit the year entirely.
- ACCESS DATE: Always include "Accessed Month Day, Year." before URL.
- URL: No period after URL.
- Platform/database name may be included as the Reference Title or after publisher.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.CONFERENCE: """
FORMAT (paper presentation):  Author FM. Title of paper. Paper presented at: Full Conference Name; Month Day–Day, Year; City, Country.
FORMAT (poster presentation):  Author FM. Title of poster. Poster presented at: Full Conference Name; Month Day–Day, Year; City, Country.
FORMAT (published proceedings, single ed):  Author FM. Chapter/paper title. In: Editor FM, ed. Proceedings Title. Publisher; Year:fpage-lpage.
FORMAT (published proceedings, multiple eds): Author FM. Chapter/paper title. In: Editor FM, Editor FM, eds. Proceedings Title. Publisher; Year:fpage-lpage.
RULES:
- PRESENTATION TYPE: Use "Paper presented at:" for oral presentations;
  "Poster presented at:" for poster presentations.
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
RULES:
- TITLE: sentence case, italicise if possible.
- DEGREE TYPE in brackets immediately after title (no comma before bracket):
  Use [doctoral dissertation] for PhD/doctoral work.
  Use [master's thesis] for master's-level work.
  Use the exact degree label if specified in the source.
- UNIVERSITY: treat as publisher (replaces publisher field).
- YEAR: placed after university, preceded by semicolon: University; Year.
- If available online: add "Accessed Month Day, Year. URL" after year.
  No period after URL.
- Strip any non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.REPORT: """
FORMAT (with report number, DOI):  Author FM. Title of report. Institution Name; Year. Report No. XXXX. doi:XXXXX
FORMAT (with report number, URL):  Author FM. Title of report. Institution Name; Year. Report No. XXXX. URL
FORMAT (no report number):         Author FM. Title of report. Institution Name; Year. doi:XXXXX
FORMAT (org author):               Organisation Name. Title of report. Institution Name; Year. Report No. XXXX.
RULES:
- TITLE: sentence case, italicise if possible.
- AUTHORS: same format as journal. If organisation is the author, use organisation name directly.
- INSTITUTION: replaces publisher. Retain full institution name.
- YEAR: placed after institution, preceded by semicolon: Institution; Year.
- REPORT NUMBER: include if present. Format: "Report No. XXXX." placed after year.
  If no report number, omit entirely.
- DOI: prefix "doi:" after report number (or after year if no report number). No period after.
- URL: include if no DOI. No period after URL.
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
  CRITICAL: Retain EVERY initial verbatim from source. Never collapse "JA" → "J" or drop any initial.
  CRITICAL: Every initial MUST have a period: "Smith, J. A., Jones, B. C."
  CRITICAL: Do NOT remove study groups or steering committees (e.g., "International Steering Committee for...") attached to the author list. Treat them as part of the author group.
  CRITICAL: Retain name generation suffixes exactly (Jr., Sr., II, III, IV). Format as:
    "Collins, Jr., J. W." — suffix comes after surname, before initials, separated by commas.
  No author → start with article title directly.
- YEAR: (Year). followed by period. CRITICAL: Retain any lowercase suffix (2022a, 2024b) — NEVER remove.
  If full date available, format as (Year, Month Day), e.g. (2025, May 28).
- ARTICLE TITLE: Sentence case ONLY — capitalise only: first word, first word after colon/em-dash,
  proper nouns. All other words lowercase. No italics. No quotation marks.
  CRITICAL: Do NOT overwrite with database title. Source title is authoritative.
- JOURNAL NAME: Title case AND italicised (*Journal Name*). Use EXACTLY the journal name from the source.
  CRITICAL: NEVER append parenthetical location qualifiers like "(Basel, Switzerland)" unless verbatim in source.
- STATUS: Retain "Advance online publication", "Published online" etc. after journal name when present.
  Format: *Journal Name*. Advance online publication. https://doi.org/XXXXX
- Volume: italicised (*Volume*). Issue in parentheses immediately after volume, NOT italicised.
  Supplemental issue: *Volume*(Suppl. X) or *Volume*(Suppl. N).
  CRITICAL: NO space between closing parenthesis of issue and comma: *12*(3), fpage
- PAGE RANGE: en dash (–) between fpage and lpage. Use EXACTLY as in source.
  Article numbers: write as "Article XXXXX" or the number as-is.
  If only fpage exists, output only that page — NEVER repeat as lpage. bib_lpage = null when absent.
- DOI as full URL: https://doi.org/XXXXX — no period after. If no DOI, use full URL, no period.
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
  Up to 20 authors; 21+ → first 19, "…", last author.
  No personal author → organisation name in author position.
- YEAR: Retain any letter suffix. Same rules as journal.
- BOOK TITLE: Sentence case AND italicised (*Title*). Sentence case = only first word, first word after
  colon/em-dash, and proper nouns capitalised; all other words lowercase.
- EDITION: in parentheses after title if >1st: (2nd ed.). Omit for 1st editions.
  If translator AND edition, combine: (F. M. Translator, Trans.; 2nd ed.)
- TRANSLATOR: "(F. M. Translator, Trans.)" in parentheses after title/edition, before publisher.
  For translated works, add "(Original work published YYYY)" at end, no period after.
- PUBLISHER (APA 7th): CRITICAL — ALWAYS retain the publisher name. NEVER delete it.
  Omit only the city/location prefix (e.g. "New York:", "London:", "Thousand Oaks, CA:").
  Strip ONLY these corporate suffixes: "Co.", "Ltd.", "Limited", "Inc.", "LLC", "Corp.", "GmbH", "S.A.", "Lda.", "Pty."
  EXCEPTION: If publisher name EXACTLY equals the author/org name, omit publisher per APA 7th.
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
- CHAPTER AUTHOR: first. Retain every initial with periods. Same format as journal.
  Year + letter suffix rules apply.
- CHAPTER TITLE: Sentence case. No italics. No quotation marks.
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
- WEBSITE/ORGANISATION NAME: Title case AND italicised (*Website Name*).
  If the author IS the website/organisation, omit the site name to avoid repetition.
- AUTHOR: No personal author → title moves to author position.
  Organisation as author → use org name in author position.
- YEAR: Full date in parentheses if available: (Year, Month Day). Year only if no full date.
  No date available → (n.d.).
  CRITICAL: bib_year = publication year (or n.d.). bib_accessed = retrieval date. These are
  ALWAYS separate fields. Never confuse them.
- RETRIEVAL DATE: Omit for stable content. Include for content that may change (wikis, live data):
  "Retrieved Month Day, Year, from URL"
- No period after URL.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
    ReferenceType.EREFERENCE: """
FORMAT (with editor, single):    Author, F. M. (Year). Entry title. In F. M. Editor (Ed.), *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
FORMAT (with editors, multiple): Author, F. M. (Year). Entry title. In F. M. Editor & F. M. Editor (Eds.), *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
FORMAT (no editor):              Author, F. M. (Year). Entry title. In *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
FORMAT (no date):                Author, F. M. (n.d.). Entry title. In F. M. Editor (Ed.), *Reference Title*. Publisher. Retrieved Month Day, Year, from URL
RULES:
- ENTRY TITLE: Sentence case. No italics. No quotation marks.
- REFERENCE/DATABASE TITLE: Title case AND italicised (*Reference Title*).
- EDITOR LABEL: "(Ed.)" for one editor; "(Eds.)" for two or more editors.
  CRITICAL: Never use "(Ed.)" for multiple editors.
  If no editor, use "In *Reference Title*" with no editor block.
- Editor initials appear BEFORE surname, same as book_chapter rule.
- YEAR: Use (n.d.) if no date available.
- PUBLISHER: Retain name; strip corporate suffixes; omit city/location.
- RETRIEVAL DATE: Always include "Retrieved Month Day, Year, from URL" — required for e-references
  since content may be updated.
- No period after URL.
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
- DOI as https://doi.org/XXXXX or full URL. No period after.
  If no DOI and no URL, end with period after institution.
- Strip non-breaking spaces (U+00A0) silently.
- COMPLETENESS: Output the FULL reference. Never truncate or omit any element.
""",
}


# ─────────────────────────────────────────────
# PROMPT BUILDER
# ─────────────────────────────────────────────

CONVERSION_MAP: Dict[Tuple[CitationStyle, CitationStyle], Tuple[str, str]] = {
    (CitationStyle.AMA, CitationStyle.APA): ("AMA 11th Edition", "APA 7th Edition"),
    (CitationStyle.APA, CitationStyle.AMA): ("APA 7th Edition", "AMA 11th Edition"),
    (CitationStyle.APA, CitationStyle.APA): ("APA 7th Edition", "APA 7th Edition (Strict Formatting Validation)"),
    (CitationStyle.AMA, CitationStyle.AMA): ("AMA 11th Edition", "AMA 11th Edition (Strict Formatting Validation)"),
}


def _build_prompt(
    raw_text: str,
    source_style: CitationStyle,
    target_style: CitationStyle,
    cr_item: Optional[Dict[str, Any]] = None,
) -> str:
    source_label, target_label = CONVERSION_MAP[(source_style, target_style)]
    rules_map = APA_RULES if target_style == CitationStyle.APA else AMA_RULES
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
"Anonymous"), heavily abbreviated, or missing. Always prioritise keeping the specific
authors provided in the input string; only use the database authors to correct minor
spelling mistakes.
CRITICAL TITLE RULE: DO NOT overwrite the article/chapter/book title from the input
with the database title unless the input title is completely absent.
{json.dumps(cr_item, indent=2)}
"""

    return f"""You are a professional bibliographic reference conversion expert specialising in {source_label} to {target_label} conversion.

## YOUR TASK
1. Detect the reference type (journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report).
2. Extract ALL available metadata into the bib_ fields. If database match metadata is provided below, incorporate it completely to fix/complete the reference (add DOIs, fix capitals, expand abbreviations).
3. Reformat the reference strictly according to {target_label} rules for the detected type.
4. Note any missing data, assumptions made, or issues in "conversion_notes".

## STRICT EXTRACTION RULES
- bib_reftype: one of: journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report, unknown
  BOOK vs EDITED_BOOK: classify as edited_book ONLY when editor markers ("Ed.", "Eds.", "ed.",
  "eds.", "edited by") are present AND no separate author appears before the title.
  If both author + editor present → book_chapter. Authors only, no editor markers → book.
- bib_surname / bib_fname: ALL authors in order, pipe-separated (|) if multiple.
  e.g. "Smith|Jones|Lee" / "John A|Mary B|Chris"
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
- All other string fields: extract verbatim from source
- Return null for any field not present in the source — NEVER fabricate data
- COMPLETENESS: The formatted_output must be the FULL reference — never truncate, omit, abbreviate,
  or paraphrase any element.
  CRITICAL: Do NOT overwrite source titles with database titles unless source title is completely absent.

## {target_label} FORMATTING RULES BY REFERENCE TYPE
{rules_block}
{cr_block}
## INPUT REFERENCE ({source_label})
{raw_text}

## OUTPUT
Return valid JSON matching the required schema exactly.
"""


def _build_validation_prompt(raw_text: str, target_style: CitationStyle) -> str:
    target_label = "APA 7th Edition" if target_style == CitationStyle.APA else "AMA 11th Edition"
    rules_map = APA_RULES if target_style == CitationStyle.APA else AMA_RULES
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
- bib_reftype: one of: journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report, unknown
- bib_surname / bib_fname: ALL authors in order, pipe-separated (|) if multiple.
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
    )

    last_exc: Optional[Exception] = None
    for attempt in range(1, max_retries + 1):
        try:
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

            raw_json = response.text
            if not raw_json or not raw_json.strip():
                logger.warning(f"[Attempt {attempt}] Empty response text.")
                last_exc = ValueError("Empty response text")
                _backoff(attempt, max_retries)
                continue

            return raw_json

        except Exception as exc:
            last_exc = exc
            err_str = str(exc).lower()
            if any(kw in err_str for kw in ("rate", "quota", "503", "429", "timeout", "unavailable")):
                logger.warning(f"[Attempt {attempt}] Transient error: {exc}. Retrying…")
                _backoff(attempt, max_retries)
            else:
                logger.error(f"[Attempt {attempt}] Non-retriable error: {exc}")
                break

    logger.error(f"All {max_retries} Gemini attempts failed. Last error: {last_exc}")
    return None


def _backoff(attempt: int, max_retries: int) -> None:
    if attempt < max_retries:
        delay = _RETRY_BASE_DELAY * (2 ** (attempt - 1))
        logger.debug(f"Back-off: sleeping {delay:.1f}s before retry {attempt + 1}.")
        time.sleep(delay)


def _resolve_api_key() -> Optional[str]:
    return os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")


# ─────────────────────────────────────────────
# PUBLIC API
# ─────────────────────────────────────────────

def convert_reference(
    raw_text: str,
    source_style: CitationStyle,
    target_style: CitationStyle,
    model_name: str = DEFAULT_MODEL,
    cr_item: Optional[Dict[str, Any]] = None,
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
        logger.error("API key not found (checked GEMINI_API_KEY and GOOGLE_API_KEY).")
        return None

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

    source_lbl, target_lbl = CONVERSION_MAP[key]
    ref_type = meta.get("bib_reftype", "unknown")
    logger.info(f"Converted [{ref_type}] {source_lbl} → {target_lbl}")
    if parsed.get("conversion_notes"):
        logger.warning(f"Conversion notes: {parsed['conversion_notes']}")

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
        logger.error("API key not found (checked GEMINI_API_KEY and GOOGLE_API_KEY).")
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
    target_lbl = "APA 7th" if target_style == CitationStyle.APA else "AMA 11th"
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
) -> List[Optional[Dict[str, Any]]]:
    if cr_items and len(cr_items) != len(references):
        raise ValueError(
            f"cr_items length ({len(cr_items)}) must match references length ({len(references)})."
        )

    results: List[Optional[Dict[str, Any]]] = []
    for i, ref in enumerate(references):
        logger.info(f"Processing reference {i + 1}/{len(references)}")
        cr = cr_items[i] if cr_items else None
        result = convert_reference(ref, source_style, target_style, model_name, cr_item=cr)
        results.append(result)
    return results
import os
import json
import logging
from typing import Optional, Dict, Any
from enum import Enum

from google import genai
from google.genai import types

logger = logging.getLogger(__name__)


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

BIB_FIELDS = [
    "bib_reftype",        # detected reference type
    "bib_surname",        # all author surnames, pipe-separated (|) if multiple
    "bib_fname",          # all author first names/initials, pipe-separated
    "bib_title",          # article/chapter title
    "bib_journal",        # journal name
    "bib_year",           # publication year (4-digit)
    "bib_volume",         # volume number
    "bib_issue",          # issue number
    "bib_fpage",          # first page
    "bib_lpage",          # last page
    "bib_doi",            # DOI (raw, no URL prefix)
    "bib_url",            # URL
    "bib_accessed",       # access date for web/eref
    "bib_book",           # book title
    "bib_chaptertitle",   # chapter title (for book chapters)
    "bib_editionno",      # edition number
    "bib_ed_surname",     # editor surnames, pipe-separated
    "bib_ed_fname",       # editor first names/initials, pipe-separated
    "bib_publisher",      # publisher name
    "bib_location",       # publisher location (city, state/country)
    "bib_institution",    # institution (reports)
    "bib_school",         # university/school (thesis)
    "bib_conference",     # full conference name
    "bib_confacronym",    # conference acronym (e.g., IEEE CVPR)
    "bib_conflocation",   # conference location
    "bib_confdate",       # conference dates
    "bib_deg",            # degree type (PhD, Master's, etc.)
    "bib_reportnum",      # report number
    "bib_series",         # book series name
    "bib_isbn",           # ISBN
    "bib_issn",           # ISSN
    "bib_pmid",           # PubMed ID
]


# ─────────────────────────────────────────────
# RESPONSE SCHEMA (new google.genai SDK)
# ─────────────────────────────────────────────

def _build_response_schema() -> types.Schema:
    return types.Schema(
        type=types.Type.OBJECT,
        required=["formatted_output", "metadata", "conversion_notes"],
        properties={
            "formatted_output": types.Schema(
                type=types.Type.STRING,
                description="The fully formatted reference in the target citation style."
            ),
            "metadata": types.Schema(
                type=types.Type.OBJECT,
                properties={
                    field: types.Schema(type=types.Type.STRING, nullable=True)
                    for field in BIB_FIELDS
                },
                required=BIB_FIELDS
            ),
            "conversion_notes": types.Schema(
                type=types.Type.STRING,
                nullable=True,
                description="Any warnings, assumptions, or missing data noted during conversion."
            ),
        }
    )


RESPONSE_SCHEMA = _build_response_schema()


# ─────────────────────────────────────────────
# PER-TYPE STYLE RULES
# ─────────────────────────────────────────────

AMA_RULES: Dict[str, str] = {
    ReferenceType.JOURNAL: """
FORMAT:  Surname FM, Surname FM. Title of article. Journal Abbrev. Year;Volume(Issue):fpage-lpage. doi:XXXXX
RULES:
- Authors: Last name followed by initials, no periods, comma-separated. Up to 6 authors; if more use "et al" after 6th.
- Journal name: Abbreviated per NLM/MEDLINE. Italicize if possible.
- Year;Volume(Issue):pages — no spaces around semicolon or colon.
- DOI: prefix with "doi:" (lowercase).
- If no DOI, provide URL.
""",
    ReferenceType.BOOK: """
FORMAT:  Surname FM, Surname FM. Title of Book. Edition ed. Publisher; Year.
RULES:
- Authors: same as journal.
- Book title: sentence case, italicize if possible.
- Edition: numeric only if >1st (e.g., "2nd ed.").
- Publisher city NOT required in AMA 11th ed.
- End with period.
""",
    ReferenceType.EDITED_BOOK: """
FORMAT:  Surname FM, ed. OR Surname FM, Surname FM, eds. Title of Book. Publisher; Year.
RULES:
- Use "ed." for single editor, "eds." for multiple.
- Otherwise same as book.
""",
    ReferenceType.BOOK_CHAPTER: """
FORMAT:  Author FM. Chapter title. In: Editor FM, ed. Book Title. Publisher; Year:fpage-lpage.
RULES:
- Chapter author comes first.
- "In:" introduces the book.
- Editor(s) with "ed."/"eds." designation.
- Pages after Year colon (no "pp.").
""",
    ReferenceType.WEBSITE: """
FORMAT:  Author FM. Title of page/document. Website Name. Published/Updated Month Day, Year. Accessed Month Day, Year. URL
RULES:
- If no author, start with title.
- Include both publication date and access date.
- URL on same line, no period after URL.
- No DOI for websites.
""",
    ReferenceType.EREFERENCE: """
FORMAT:  Author FM. Entry title. In: Editor FM, ed. Reference Book Title. Publisher; Year. Accessed Month Day, Year. URL
RULES:
- Treat like a book chapter but include access date and URL.
- Platform/database name may be included.
""",
    ReferenceType.CONFERENCE: """
FORMAT:  Author FM. Title of paper. Paper presented at: Conference Name; Date; Location.
RULES:
- "Paper presented at:" is literal text.
- Date: Month Day-Day, Year format.
- Location: City, State/Country.
""",
    ReferenceType.THESIS: """
FORMAT:  Author FM. Title of thesis [type of thesis]. University Name; Year.
RULES:
- Degree type in brackets: [doctoral dissertation] or [master's thesis].
- Publisher = University name.
- If available online: add "Accessed Month Day, Year. URL"
""",
    ReferenceType.REPORT: """
FORMAT:  Author FM. Title of Report. Institution Name; Year. Report No. XXXX.
RULES:
- Include report number if present.
- Institution replaces publisher.
""",
}

APA_RULES: Dict[str, str] = {
    ReferenceType.JOURNAL: """
FORMAT:  Surname, F. M., & Surname, F. M. (Year). Title of article. *Journal Name*, *Volume*(Issue), fpage–lpage. https://doi.org/XXXXX
RULES:
- Authors: Surname, Initials. Use "&" before last author. Up to 20 authors; if >20 use "..." then last author.
- Year in parentheses followed by period.
- Article title: sentence case, no italics.
- Journal name and volume: italicize.
- Issue in parentheses, not italicized.
- En dash (–) between pages, not hyphen.
- DOI as full URL: https://doi.org/...
""",
    ReferenceType.BOOK: """
FORMAT:  Surname, F. M., & Surname, F. M. (Year). *Title of book* (Xth ed.). Publisher. https://doi.org/XXXXX
RULES:
- Book title italicized, sentence case.
- Edition in parentheses after title if >1st.
- Publisher name only (no location in APA 7th).
- DOI or URL if available.
""",
    ReferenceType.EDITED_BOOK: """
FORMAT:  Surname, F. M., & Surname, F. M. (Eds.). (Year). *Title of book*. Publisher.
RULES:
- "(Ed.)" or "(Eds.)" after editor name(s).
- Otherwise same as book.
""",
    ReferenceType.BOOK_CHAPTER: """
FORMAT:  Author, F. M. (Year). Chapter title. In F. M. Editor (Ed.), *Book Title* (pp. fpage–lpage). Publisher. https://doi.org/XXXXX
RULES:
- Chapter author first.
- "In" introduces editor.
- Editor initials BEFORE surname here.
- Book title italicized.
- Pages with "pp." in parentheses before publisher.
- En dash between pages.
""",
    ReferenceType.WEBSITE: """
FORMAT:  Author, F. M. (Year, Month Day). Title of page. *Website Name*. Retrieved Month Day, Year, from URL
RULES:
- If content may change over time, include retrieval date.
- If static content, retrieval date optional.
- Website name italicized.
- No period after URL.
""",
    ReferenceType.EREFERENCE: """
FORMAT:  Author, F. M. (Year). Entry title. In F. M. Editor (Ed.), *Reference title*. Publisher. Retrieved Month Day, Year, from URL
RULES:
- Treat like book chapter but with retrieval date and URL.
- No page numbers if content is not paginated.
""",
    ReferenceType.CONFERENCE: """
FORMAT (paper):       Author, F. M. (Year, Month Day–Day). *Title of paper* [Conference session]. Conference Name, Location. https://doi.org/XXXXX
FORMAT (proceedings): Author, F. M. (Year). Title. In F. M. Editor (Ed.), *Proceedings Title* (pp. X–X). Publisher.
RULES:
- Use "Conference session", "Paper presentation", or "Poster session" descriptor in brackets.
- Dates in Month Day–Day, Year format.
""",
    ReferenceType.THESIS: """
FORMAT:  Author, F. M. (Year). *Title of thesis* [Doctoral dissertation/Master's thesis, University Name]. Database/Repository Name. URL
RULES:
- Degree type and institution in brackets.
- Database name if retrieved from ProQuest or institutional repository.
- No period after URL.
""",
    ReferenceType.REPORT: """
FORMAT:  Author, F. M. (Year). *Title of report* (Report No. XXXX). Institution. https://doi.org/XXXXX
RULES:
- Report number in parentheses after title if available.
- Institution as publisher.
""",
}


# ─────────────────────────────────────────────
# PROMPT BUILDER
# ─────────────────────────────────────────────

CONVERSION_MAP = {
    (CitationStyle.AMA, CitationStyle.APA): ("AMA 11th Edition", "APA 7th Edition"),
    (CitationStyle.APA, CitationStyle.AMA): ("APA 7th Edition", "AMA 11th Edition"),
}


def _build_prompt(raw_text: str, source_style: CitationStyle, target_style: CitationStyle) -> str:
    source_label, target_label = CONVERSION_MAP[(source_style, target_style)]
    rules_map = APA_RULES if target_style == CitationStyle.APA else AMA_RULES
    rules_block = "\n".join([
        f"### {ref_type.upper()}\n{rules}"
        for ref_type, rules in rules_map.items()
    ])

    return f"""You are a professional bibliographic reference conversion expert specializing in {source_label} to {target_label} conversion.

## YOUR TASK
1. Detect the reference type (journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report).
2. Extract ALL available metadata into the bib_ fields.
3. Reformat the reference strictly according to {target_label} rules for the detected type.
4. Note any missing data, assumptions made, or issues in "conversion_notes".

## STRICT EXTRACTION RULES
- bib_reftype: one of: journal, book, edited_book, book_chapter, website, ereference, conference, thesis, report, unknown
- bib_surname / bib_fname: ALL authors in order, pipe-separated (|) if multiple. e.g. "Smith|Jones|Lee" / "John A|Mary B|Chris"
- bib_ed_surname / bib_ed_fname: same format for editors
- bib_year: 4-digit year ONLY
- bib_volume / bib_issue: numeric string only, no labels
- bib_fpage / bib_lpage: digits only, no "pp.", "p.", or labels
- bib_doi: raw DOI string only — strip "https://doi.org/" prefix
- bib_url: full URL only, no trailing period
- bib_accessed: date in "Month DD, YYYY" format
- bib_confdate: full date range as written
- bib_editionno: number only (e.g., "2", "3")
- bib_deg: full degree name (e.g., "Doctoral dissertation", "Master's thesis")
- All other string fields: extract verbatim from source
- Return null for any field not present in the source — NEVER fabricate data

## {target_label} FORMATTING RULES BY REFERENCE TYPE
{rules_block}

## INPUT REFERENCE ({source_label})
{raw_text}

## OUTPUT
Return valid JSON matching the required schema exactly.
"""


# ─────────────────────────────────────────────
# MAIN CONVERTER
# ─────────────────────────────────────────────

def convert_reference(
    raw_text: str,
    source_style: CitationStyle,
    target_style: CitationStyle,
    model_name: str = "gemini-2.0-flash",
) -> Optional[Dict[str, Any]]:
    """
    Convert a bibliographic reference between AMA 11th and APA 7th edition.

    Args:
        raw_text:      Raw reference string to convert.
        source_style:  Source citation style (CitationStyle.AMA or CitationStyle.APA).
        target_style:  Target citation style (CitationStyle.AMA or CitationStyle.APA).
        model_name:    Gemini model to use.

    Returns:
        Dict with:
          - formatted_output (str): converted reference string
          - metadata (dict): extracted bib_ fields
          - conversion_notes (str|None): warnings or assumptions
        Returns None on failure.
    """
    # ── Input validation ──────────────────────────────────────────
    if not raw_text or not raw_text.strip():
        logger.error("raw_text is empty")
        return None

    if source_style == target_style:
        logger.error("source_style and target_style must be different")
        return None

    if (source_style, target_style) not in CONVERSION_MAP:
        logger.error(f"Unsupported conversion: {source_style} → {target_style}")
        return None

    api_key = os.environ.get("GOOGLE_API_KEY")
    if not api_key:
        logger.error("GOOGLE_API_KEY not found in environment")
        return None

    # ── API call (new google.genai SDK) ───────────────────────────
    try:
        client = genai.Client(api_key=api_key)

        generation_config = types.GenerateContentConfig(
            response_mime_type="application/json",
            response_schema=RESPONSE_SCHEMA,
            temperature=0.0,
            top_p=1.0,
            top_k=1,
        )

        prompt = _build_prompt(raw_text, source_style, target_style)

        response = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config=generation_config,
        )

        # ── Response validation ───────────────────────────────────
        if not response or not response.candidates:
            logger.error("No candidates in Gemini response")
            return None

        candidate = response.candidates[0]
        finish_reason = candidate.finish_reason

        # STOP = 1, MAX_TOKENS = 2 in the new SDK
        if finish_reason not in (
            types.FinishReason.STOP,
            types.FinishReason.MAX_TOKENS,
        ):
            logger.error(f"Unexpected finish reason: {finish_reason}")
            return None

        raw_json = response.text
        if not raw_json or not raw_json.strip():
            logger.error("Empty response text from Gemini")
            return None

        parsed: Dict[str, Any] = json.loads(raw_json)

        # ── Schema enforcement ────────────────────────────────────
        if "formatted_output" not in parsed or "metadata" not in parsed:
            logger.error(f"Missing top-level keys in response: {list(parsed.keys())}")
            return None

        if not isinstance(parsed["formatted_output"], str) or not parsed["formatted_output"].strip():
            logger.error("formatted_output is empty or not a string")
            return None

        # Normalize metadata — ensure all bib_ fields exist
        meta = parsed.get("metadata", {})
        for field in BIB_FIELDS:
            meta.setdefault(field, None)
        parsed["metadata"] = meta

        # Log conversion summary
        ref_type = meta.get("bib_reftype", "unknown")
        source_lbl, target_lbl = CONVERSION_MAP[(source_style, target_style)]
        logger.info(f"Converted [{ref_type}] {source_lbl} → {target_lbl}")
        if parsed.get("conversion_notes"):
            logger.warning(f"Conversion notes: {parsed['conversion_notes']}")

        return parsed

    except json.JSONDecodeError as e:
        logger.error(f"JSON decode error: {e}")
        return None
    except Exception as e:
        logger.exception(f"Unexpected error during conversion: {e}")
        return None


# ─────────────────────────────────────────────
# BATCH CONVERTER
# ─────────────────────────────────────────────

def convert_references_batch(
    references: list,
    source_style: CitationStyle,
    target_style: CitationStyle,
    model_name: str = "gemini-2.0-flash",
) -> list:
    """
    Convert a list of references. Returns a list of results in the same order.
    Failed conversions return None at their index position.
    """
    results = []
    for i, ref in enumerate(references):
        logger.info(f"Processing reference {i + 1}/{len(references)}")
        result = convert_reference(ref, source_style, target_style, model_name)
        results.append(result)
    return results
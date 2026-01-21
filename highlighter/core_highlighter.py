# word_highlighter.py
import re
import win32com.client
import pythoncom
from pathlib import Path
import os
import sys
import traceback

# ---------------------------
# Word color index map (adjust if you prefer)
# Values chosen to match common Word WdColorIndex constants
# ---------------------------
COLOR_MAP = {
    "nohighlight": 0,
    "black": 1,
    "blue": 2,
    "turquoise": 3,
    "brightgreen": 4,
    "pink": 5,
    "red": 6,
    "yellow": 7,
    "white": 8,
    "darkBlue": 9,
    "teal": 10,
    "darkGreen": 11,
    "darkMagenta": 12,
    "darkRed": 13,
    "darkYellow": 14,
    "lightGray": 15,
    "darkGray": 16,
    # friendly synonyms
    "cyan": 3,
    "green": 4,
    "magenta": 12,
    "darkCyan": 3,
    "darkMagenta": 12,
    "darkRed": 13,
    "darkYellow": 14,
    "darkGreen": 11,
    "darkBlue": 9,
    "darkCyan": 3,
    "lightGray": 15,
}

# ---------------------------
# Utilities
# ---------------------------
def _color_index_from_name(name):
    if not name:
        return COLOR_MAP["nohighlight"]
    key = str(name).strip().lower()
    return COLOR_MAP.get(key, COLOR_MAP["yellow"])  # default yellow if unknown

def highlight_range(range_obj, color_index):
    """Apply highlight color to a Word Range object (safe wrapper)."""
    try:
        range_obj.HighlightColorIndex = int(color_index)
    except Exception:
        # best-effort: ignore failures (range may be invalid)
        pass

def highlight_words_in_styles(doc):
    """
    Python equivalent of the given VB Macro:
    - Find paragraphs whose Style.NameLocal matches style list
    - Highlight whole paragraph yellow
    - Then highlight all words <= 4 chars in bright green
    """

    target_styles = {"T1", "CT", "H1", "H2", "H2A", "H3", "H3A", "NBX1-TTL"}

    wdTurquoise = 3
    wdBrightGreen = 4
    wdFindStop = 0  # no wrap

    for para in doc.Paragraphs:
        try:
            style_name = str(para.Range.Style.NameLocal)
        except:
            continue

        if style_name not in target_styles:
            continue

        # Highlight entire paragraph yellow
        try:
            para.Range.HighlightColorIndex = wdTurquoise
        except:
            pass

        # Extract words
        text = para.Range.Text.strip()
        if not text:
            continue

        words = text.split()

        # Highlight each short word (<= 4 chars)
        for w in words:
            if len(w) > 4:
                continue

            # COM Find inside paragraph
            rng = para.Range.Duplicate
            find = rng.Find

            find.Text = w
            find.MatchWholeWord = True
            find.MatchCase = False
            find.ClearFormatting()
            find.Highlight = True
            find.Wrap = wdFindStop

            # Execute find loop
            while find.Execute():
                try:
                    rng.HighlightColorIndex = wdBrightGreen
                except:
                    pass

                # Collapse to end to continue searching
                rng.Collapse(0)  # wdCollapseEnd
    return

def highlight_multilingual_chars(doc):
    """
    Python equivalent of VBA Highlight_MultilingualChars macro.
    Scans entire document for multilingual Unicode ranges and highlights them.
    """

    # Word highlight color index mapping
    wc = {
        "Chinese": 3,       # wdTurquoise
        "Greek": 3,         # wdTurquoise
        "Cyrillic": 5,      # wdPink
        "Hebrew": 4,        # wdGreen
        "Arabic": 2,        # wdBlue
        "Devanagari": 6,    # wdRed
        "Japanese": 12,     # wdViolet
        "CJK Symbols": 16,  # wdGray25
        "Korean": 12,       # wdViolet
        "Thai": 13,         # wdDarkRed
        "Currency": 9       # wdDarkBlue
    }

    # Unicode ranges to check
    unicode_blocks = [
        ("Chinese",       19968, 40959),
        ("Greek",         0x370, 0x3FF),
        ("Cyrillic",      0x400, 0x4FF),
        ("Hebrew",        0x590, 0x5FF),
        ("Arabic",        0x600, 0x6FF),
        ("Arabic",        0x750, 0x77F),
        ("Devanagari",    0x900, 0x97F),
        ("Japanese",      0x3040, 0x309F),
        ("Japanese",      0x30A0, 0x30FF),
        ("CJK Symbols",   0x3000, 0x303F),
        ("CJK Symbols",   0x3200, 0x32FF),
        ("CJK Symbols",   0x3300, 0x33FF),
        ("CJK Symbols",   0xFE30, 0xFE4F),
        ("Korean",        0xAC00, 0xD7AF),
        ("Korean",        0x1100, 0x11FF),
        ("Korean",        0x3130, 0x318F),
        ("Thai",          0x0E00, 0x0E7F),
        ("Currency",      0x20A0, 0x20CF),
    ]

    def classify_char(code_point):
        for name, start, end in unicode_blocks:
            if start <= code_point <= end:
                return name
        return None

    # Iterate through every story range (main text, headers, footers, footnotes, tables, etc.)
    story = doc.StoryRanges

    for storyRange in story:
        r = storyRange
        while r is not None:
            text = r.Text
            base = r.Start

            for idx, ch in enumerate(text):
                cp = ord(ch)
                char_type = classify_char(cp)
                if not char_type:
                    continue

                color = wc.get(char_type)
                if not color:
                    continue

                try:
                    rr = doc.Range(base + idx, base + idx + 1)
                    rr.HighlightColorIndex = color
                except:
                    pass

            r = r.NextStoryRange

# ---------------------------
# Regex highlight engine (robust)
# ---------------------------
def regex_highlight(doc, patterns, color_name, flags=re.IGNORECASE):
    """
    Highlight all matches for each regex in `patterns` within the document `doc`.
    This function searches:
      - document paragraphs
      - tables (cell-by-cell)
      - headers & footers (all sections)
      - footnotes/endnotes (if present)
    patterns: list of regex strings
    color_name: string (maps to COLOR_MAP)
    """
    color_index = _color_index_from_name(color_name)
    compiled_list = []
    for p in patterns:
        try:
            compiled_list.append(re.compile(p, flags))
        except re.error:
            # skip invalid regex but report
            print(f"[WARNING] Invalid regex skipped: {p}")
    if not compiled_list:
        return

    # Helper to highlight matches inside a text piece with known base_start
    def _highlight_in_text(base_start, text, compiled_regex):
        # text here should be the exact .Range.Text content (including Word marks)
        # but we will strip the trailing cell/row marks before mapping offsets.
        # For paragraphs: trailing '\r' ; for cells: '\r\x07' typical
        stripped = text
        # We will not blindly remove mid-document control chars; just normalize trailing marks
        stripped = stripped.rstrip("\r\x07")
        for m in compiled_regex.finditer(stripped):
            start = base_start + m.start()
            end = base_start + m.end()
            try:
                r = doc.Range(start, end)
                highlight_range(r, color_index)
            except Exception:
                # fallback: ignore highlighting for this match if mapping failed
                pass

    # 1) Paragraphs
    try:
        for para in doc.Paragraphs:
            if is_skip_style(para):
                continue
            try:
                prange = para.Range
                text = prange.Text
                base = prange.Start
                for cre in compiled_list:
                    _highlight_in_text(base, text, cre)
            except Exception:
                # skip bad paragraphs
                continue
    except Exception:
        # if doc.Paragraphs access fails, continue gracefully
        pass

    # 2) Tables - iterate cell-by-cell for correct indices
    try:
        for table in doc.Tables:
            try:
                for row in table.Rows:
                    for cell in row.Cells:
                        try:
                            crange = cell.Range
                            text = crange.Text
                            base = crange.Start
                            for cre in compiled_list:
                                _highlight_in_text(base, text, cre)
                        except Exception:
                            continue
            except Exception:
                continue
    except Exception:
        pass

    # 3) Headers & Footers for each section and each header/footer type
    # Word indexes: 1..3 (Primary, FirstPage, EvenPages)
    try:
        for section in doc.Sections:
            for idx in (1, 2, 3):
                try:
                    header = section.Headers(idx)
                    if header.Exists:
                        hr = header.Range
                        text = hr.Text
                        base = hr.Start
                        for cre in compiled_list:
                            _highlight_in_text(base, text, cre)
                except Exception:
                    pass
            for idx in (1, 2, 3):
                try:
                    footer = section.Footers(idx)
                    if footer.Exists:
                        fr = footer.Range
                        text = fr.Text
                        base = fr.Start
                        for cre in compiled_list:
                            _highlight_in_text(base, text, cre)
                except Exception:
                    pass
    except Exception:
        pass

    # 4) Footnotes and endnotes (if present)
    try:
        if hasattr(doc, "Footnotes"):
            for fn in doc.Footnotes:
                try:
                    fr = fn.Range
                    text = fr.Text
                    base = fr.Start
                    for cre in compiled_list:
                        _highlight_in_text(base, text, cre)
                except Exception:
                    pass
        if hasattr(doc, "Endnotes"):
            for en in doc.Endnotes:
                try:
                    er = en.Range
                    text = er.Text
                    base = er.Start
                    for cre in compiled_list:
                        _highlight_in_text(base, text, cre)
                except Exception:
                    pass
    except Exception:
        pass

# ---------------------------
# Heading hierarchy check
# ---------------------------
def check_heading_hierarchy(doc):
    """
    Simple heading hierarchy checker.
    Assumes headings are implemented with Word OutlineLevels or styles like Heading 1..6.
    Adds comments where a level jumps by more than 1.
    """
    prev_level = 0
    try:
        for para in doc.Paragraphs:
            try:
                style = para.Range.Style
                # Word heading styles usually are "Heading 1", "Heading 2", ...
                style_name = str(style)
                m = re.search(r"H\s+([1-9])", style_name, re.IGNORECASE)
                if m:
                    level = int(m.group(1))
                    if level > prev_level + 1 and prev_level != 0:
                        try:
                            doc.Comments.Add(para.Range, f"Heading hierarchy issue: jumped from {prev_level} to {level}")
                        except Exception:
                            pass
                    prev_level = level
            except Exception:
                continue
    except Exception:
        pass

# ---------------------------
# Unpaired punctuation and quotes check (paragraph-safe)
# ---------------------------
def check_unpaired_punctuation_and_quotes(doc):
    """
    Count unmatched punctuation pairs and unmatched typographic quotes.
    For robustness, operate paragraph-by-paragraph, mapping errors to end-of-doc comment.
    """
    # pairs to check (left, right)
    pairs = [
        ("(", ")"),
        ("[", "]"),
        ("{", "}"),
        ("\"", "\""),
        ("'", "'"),
        ("\u201c", "\u201d"),
        ("\u2018", "\u2019"),
    ]
    # Build full counts from doc.Content.Text but only for reporting — locating each mismatch precisely is complex;
    # We will place a summary comment at end-of-document if mismatches found.
    try:
        full_text = doc.Content.Text
        messages = []
        for left, right in pairs:
            left_count = full_text.count(left)
            right_count = full_text.count(right)
            if left_count != right_count:
                messages.append(f"Unbalanced punctuation: {left}/{right} counts ({left_count}/{right_count})")
        # Typographic quotes check (opening vs closing)
        left_quote = "\u201c"
        right_quote = "\u201d"
        lq = full_text.count(left_quote)
        rq = full_text.count(right_quote)
        if lq != rq:
            messages.append(f"Unbalanced typographic quotes: {lq} opening vs {rq} closing")
        if messages:
            try:
                end_range = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
                doc.Comments.Add(end_range, " ; ".join(messages))
            except Exception:
                pass
    except Exception:
        pass

# ---------------------------
# Main processing function
# ---------------------------
def process_docx(input_file, output_file, skip_validation=False, verbose=True):
    """
    Process a single DOCX:
      - open Word (COM)
      - run many regex_highlight calls
      - run heading/punctuation/quote checks
      - save to output_file
    """
    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        word.ScreenUpdating = False
        word.DisplayAlerts = 0  # wdAlertsNone

        input_path = str(Path(input_file).resolve())
        output_path = str(Path(output_file).resolve())

        doc = word.Documents.Open(input_path, ReadOnly=False)

        if verbose:
            print(f"Processing: {input_path}")

        # ------------------------------------
        # Progress step counter
        # ------------------------------------
        step = 1
        def step_print(msg):
            nonlocal step
            print(f"[{step}] {msg}")
            step += 1

        # ==== Highlight groups (auto-numbered) ====

        step_print("Highlighting words in specific paragraph styles...")
        highlight_words_in_styles(doc)

        step_print("Highlighting numbers + per + unit...")
        regex_highlight(doc, [r"\b[0-9]{1,}\sper\s(dL|µ|µL|mL|ml|L|g|mg|kg|min|h|hr|hour|day|week|month)\b"], "turquoise")

        step_print("Highlighting metric units...")
        regex_highlight(doc, [r"\b(micro|micron|liter|liters|mcg|mcm|mcl)\b"], "turquoise")

        step_print("Highlighting time abbreviations...")
        regex_highlight(doc, [r"\b(a\.m\.|p\.m\.|AM|PM|AD|BC|CE|BCE|A\.D|B\.C|C\.E|B\.C\.E|a\.d|b\.c|c\.e|b\.c\.e\.)\b"], "turquoise")

        step_print("Highlighting pressure comparisons...")
        pressure_patterns = [
            r"([pP])\s([=><-]+)\s([0-9]+(\.[0-9]+)?)",
            r"[pP]\s≤",
            r"[pP][=><]|[=><][pP]",
            r"[pP]\s[=><]",
            r"[pP]\s≤\s\?\s\d+(\.\d+)?",
            r"[pP]≥\s\?\s\d+(\.\d+)?",
            r"[pP]\s≥\s\?\s\d+(\.\d+)?",
            r"[pP]\s[-=><]\s\?\s\d+(\.\d+)?",
            r"[pP]\s[-=><]\s\?\s\d+",
        ]
        regex_highlight(doc, pressure_patterns, "turquoise")

        step_print("Highlighting medical terms...")
        regex_highlight(doc, [r"\b(DSM|COVID-19|ventilation|perfusion|V/Q|VQ|V-Q)\b"], "turquoise")

        step_print("Highlighting percent variations...")
        regex_highlight(doc, [r"(percent|per cent|percentage|%)"], "turquoise")

        step_print("Highlighting numbers ≥ 1000...")
        regex_highlight(doc, [r"\b\d{4,}\b"], "turquoise")

        step_print("Highlighting number words...")
        numwords = r"\b(Zero|one|two|three|four|five|six|seven|eight|nine|Ten|Eleven|Twelve|Thirteen|Fourteen|Fifteen|Sixteen|Seventeen|Eighteen|Nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety|twenty-|thirty-|forty-|fifty-|sixty-|seventy-|eighty-|ninety-|-one|-two|-three|-four|-five|-six|-seven|-eight|-nine)\b"
        regex_highlight(doc, [numwords], "turquoise")

        step_print("Highlighting durations...")
        duration_patterns = [
            r"\b\d+[- ]?(year|years|month|months|week|weeks|day|days|hour|hours|minute|minutes|second|seconds)\b",
            r"\b\d+[- ]?(y|mo|wk|wks|d|h|hr|hrs|min|s|sec)\b",
            r"\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)[- ]?(year|years|month|months|week|weeks|day|days|hour|hours|minute|minutes|second|seconds)\b",
            r"\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|thirty|forty|fifty|sixty|seventy|eighty|ninety)[- ]?(y|mo|wk|wks|d|h|hr|hrs|min|s|sec)\b"
        ]
        regex_highlight(doc, duration_patterns, "turquoise")

        step_print("Highlighting chapter/section markers...")
        chapter_patterns = [
            r"\bchapter\b", r"\bChapter\b",
            r"\bchapters\b", r"\bChapters\b",
            r"\bChap\b",
            r"\bsection\b", r"\bSection\b",
            r"\bSect\b",
            r"\b(Chap\.|Ch\.|Sec\.)"
        ]
        regex_highlight(doc, chapter_patterns, "turquoise")

        step_print("Highlighting common abbreviations...")
        common_abbr = [
            r"\b(e\.g|eg|i\.e|ie|vs|etc|et al)\b",
            r"\b(Dr|Drs|Mr|Mrs|Ms|Prof)\b",
            r"\b(M\.D|MD|M\.A|MA|M\.S|MS|Bsc|MSc)\b",
            r"\b(Blvd|St|Ste)\b"
        ]
        regex_highlight(doc, common_abbr, "turquoise")

        step_print("Highlighting Latin phrases...")
        regex_highlight(doc, [r"\b(in vitro|in vivo|in situ|ex situ|per se|ad hoc|de novo|a priori|de facto|status quo|a posteriori|ad libitum|ad lib|supra|verbatim|cf\.|ibid|id\.)\b"], "turquoise")

        step_print("Highlighting medical superscript terms...")
        regex_highlight(doc, [r"\b(Paco2|Pco2|co2|Sao2|o2max|Cao2|Spo2|Pvo2|o2|Fio2|PIO2|Vo2|Pao2|Pio2|PAo2|Po2)\b"], "turquoise")
        greek_patterns = [
            r"\balpha\b", r"α",
            r"\bbeta\b", r"β",
            r"\bgamma\b", r"γ",
            r"\bdelta\b", r"δ",
            r"\bepsilon\b", r"ε",
            r"\bzeta\b", r"ζ",
            r"\btheta\b", r"θ",
            r"\beta\b", r"η",
            r"\biota\b", r"ι",
            r"\bkappa\b", r"κ",
            r"\blambda\b", r"λ",
            r"\bmu\b", r"μ",
            r"\bnu\b", r"ν",
            r"\bxi\b", r"ξ",
            r"\bomicron\b", r"ο",
            r"\bpi\b", r"π",
            r"\brho\b", r"ρ",
            r"\bsigma\b", r"σ",
            r"\btau\b", r"τ",
            r"\bupsilon\b", r"υ",
            r"\bphi\b", r"φ",
            r"\bchi\b", r"χ",
            r"\bpsi\b", r"ψ",
            r"\bomega\b", r"ω",
            # Uppercase/capitalized
            r"\bAlpha\b", r"Α",
            r"\bBeta\b", r"Β",
            r"\bGamma\b", r"Γ",
            r"\bDelta\b", r"Δ",
            r"\bEpsilon\b", r"Ε",
            r"\bZeta\b", r"Ζ",
            r"\bTheta\b", r"Θ",
            r"\bEta\b", r"Η",
            r"\bIota\b", r"Ι",
            r"\bKappa\b", r"Κ",
            r"\bLambda\b", r"Λ",
            r"\bMu\b", r"Μ",
            r"\bNu\b", r"Ν",
            r"\bXi\b", r"Ξ",
            r"\bOmicron\b", r"Ο",
            r"\bPi\b", r"Π",
            r"\bRho\b", r"Ρ",
            r"\bSigma\b", r"Σ",
            r"\bTau\b", r"Τ",
            r"\bUpsilon\b", r"Υ",
            r"\bPhi\b", r"Φ",
            r"\bChi\b", r"Χ",
            r"\bPsi\b", r"Ψ",
            r"\bOmega\b", r"Ω",
        ]
        step_print("Highlighting Greek letters...")
        regex_highlight(doc, greek_patterns, "turquoise")
        ranges = [
            r"[0-9]{1,}\.[0-9]{1,}-[0-9]{1,}\.[0-9]{1,}",
            r"[0-9]{1,}-[0-9]{1,}",
            r"[0-9]{1,}\s*-\s*[0-9]{1,}",
            r"[0-9]{1,}--[0-9]{1,}",
            r"[0-9]{1,}\s–\s[0-9]{1,}",   # en dash with spaces
            r"[0-9]{1,}–[0-9]{1,}",     # en dash without spaces
            r"[0-9]{1,}\s—\s[0-9]{1,}",   # em dash with spaces
            r"[0-9]{1,}—[0-9]{1,}",     # em dash without spaces
            r"[0-9]{1,}\s+to\s+[0-9]{1,}",
        ]
        step_print("Highlighting numeric ranges...")
        regex_highlight(doc, ranges, "turquoise")

        step_print("Highlighting degree angles...")
        regex_highlight(doc, [r"\b\d{1,3}-degree angle\b", r"\b\d{1,3}\sdegree angle\b"], "turquoise")

        step_print("Highlighting ° angle values...")
        regex_highlight(doc, [r"\b\d{1,3}\s?°\s?angle\b", r"\b\d{1,3}\s?°\b"], "turquoise")

        step_print("Highlighting x-ray variations...")
        regex_highlight(doc, [r"\b[xX]-?ray\b"], "turquoise")

        step_print("Highlighting PaCO2 etc...")
        regex_highlight(doc, [r"\b(Paco2|Pco2|Sao2|Spo2|Fio2|Pao2|Po2)\b"], "turquoise")

        step_print("Highlighting trademark symbols...")
        regex_highlight(doc, [r"[™®©]"], "turquoise")

        step_print("Highlighting ‘versus’ terms...")
        regex_highlight(doc, [r"\b(vs\.?|versus|v\.?)\b"], "turquoise")

        step_print("Highlighting special characters...")
        regex_highlight(doc, [r"[§¶†‡]"], "turquoise")

        step_print("Highlighting Figure/Table references...")
        regex_highlight(doc, [r"\b(Figure|Table)\s*\d+"], "turquoise")

        step_print("Checking heading hierarchy...")
        check_heading_hierarchy(doc)

        step_print("Checking punctuation and quotes...")
        check_unpaired_punctuation_and_quotes(doc)

        step_print("Highlighting multilingual Unicode characters...")
        highlight_multilingual_chars(doc)

        step_print("Highlighting italic punctuation...")
        highlight_italic_punctuation(doc)

        step_print("Highlighting comparison symbols and spellings...")
        highlight_comparison_symbols(doc)

        step_print("Highlighting math operators...")
        highlight_math_symbols(doc)

        step_print("Highlighting times/century/decade/fold...")
        highlight_time_period_terms(doc)

        # Save
        if verbose:
            print("Saving document...")
        doc.SaveAs2(output_path)
        doc.Close()
        if verbose:
            print(f"[OK] Saved processed file: {output_path}")

    except Exception as exc:
        # try to provide some helpful debugging info
        print("[ERROR] Processing failed:", str(exc))
        traceback.print_exc()
        # ensure doc closed if possible
        try:
            if 'doc' in locals() and doc is not None:
                doc.Close(False)
        except Exception:
            pass
        raise
    finally:
        try:
            if word is not None:
                word.ScreenUpdating = True
                word.DisplayAlerts = -1
                word.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

def is_skip_style(paragraph):
    """Skip highlighting in REF-N and REF-U styles."""
    try:
        name = str(paragraph.Range.Style.NameLocal).strip()
        return name in {"REF-N", "REF-U"}
    except:
        return False

def highlight_comparison_symbols(doc):
    symbols = ["<", ">", "=", "≥", "≤", "≈"]
    spellings = ["less than", "greater than", "equal to", "approximately"]

    wdTurquoise = 3

    for para in doc.Paragraphs:
        if is_skip_style(para):
            continue

        text = para.Range.Text
        base = para.Range.Start

        # Direct symbols
        for idx, ch in enumerate(text):
            if ch in symbols:
                r = doc.Range(base + idx, base + idx + 1)
                try:
                    r.HighlightColorIndex = wdTurquoise
                except:
                    pass

        # Spelled-out versions
        for sp in spellings:
            for m in re.finditer(r"\b" + re.escape(sp) + r"\b", text, re.IGNORECASE):
                r = doc.Range(base + m.start(), base + m.end())
                try:
                    r.HighlightColorIndex = wdTurquoise
                except:
                    pass
def highlight_math_symbols(doc):
    symbols = {"-": 3, "+": 3, "×": 3, "÷": 3}  # color per symbol if needed

    for para in doc.Paragraphs:
        if is_skip_style(para):
            continue

        text = para.Range.Text
        base = para.Range.Start

        for idx, ch in enumerate(text):
            if ch in symbols:
                try:
                    r = doc.Range(base + idx, base + idx + 1)
                    r.HighlightColorIndex = symbols[ch]
                except:
                    pass
def highlight_time_period_terms(doc):
    patterns = [
        r"\btimes\b",
        r"\bcentury\b",
        r"\bcenturies\b",
        r"\bdecade\b",
        r"\b\d+-fold\b",
        r"\bfold\b",
    ]

    wdTurquoise = 3

    for para in doc.Paragraphs:
        if is_skip_style(para):
            continue

        text = para.Range.Text
        base = para.Range.Start

        for p in patterns:
            for m in re.finditer(p, text, re.IGNORECASE):
                try:
                    r = doc.Range(base + m.start(), base + m.end())
                    r.HighlightColorIndex = wdTurquoise
                except:
                    pass
def highlight_italic_punctuation(doc):
    """
    Highlight . , : ; only when they are in italic formatting.
    """

    punct = {".", ",", ":", ";"}
    wdturquoise = 3

    for para in doc.Paragraphs:
        if is_skip_style(para):
            continue

        rng = para.Range
        text = rng.Text
        base = rng.Start

        for idx, ch in enumerate(text):
            if ch in punct:
                r = doc.Range(base + idx, base + idx + 1)
                try:
                    if r.Italic:  # Word formatting check
                        r.HighlightColorIndex = wdturquoise
                except:
                    pass

# ---------------------------
# CLI convenience
# ---------------------------
if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Highlight regex patterns in a Word .docx using COM.")
    p.add_argument("input", help="Input .docx file")
    p.add_argument("output", help="Output .docx file")
    p.add_argument("--skip-validation", action="store_true", help="Skip punctuation/heading checks for speed")
    p.add_argument("--quiet", action="store_true", help="Less logging")
    args = p.parse_args()

    if not os.path.exists(args.input):
        print("Input file not found:", args.input)
        sys.exit(2)
    process_docx(args.input, args.output, skip_validation=args.skip_validation, verbose=not args.quiet)

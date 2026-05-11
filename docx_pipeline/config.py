"""
config.py — Central configuration for the DOCX processing pipeline.
Edit this file to match your publisher's style template.
"""

import re
import os

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.environ.get("DOCX_INPUT_DIR",  os.path.join(BASE_DIR, "input"))
OUTPUT_DIR = os.environ.get("DOCX_OUTPUT_DIR", os.path.join(BASE_DIR, "output"))
SECRET_KEY = os.environ.get("DOCX_SECRET_KEY", "")

# ── Pipeline behavior ────────────────────────────────────────────────────────
PIPELINE_HALT_ON_ERROR = True  # If True, skip remaining steps after a step fails

# ── Step 3: Styles from which bold must be stripped ───────────────────────────
SEMANTIC_BOLD_STYLES = [
    "H1", "H2", "H3", "H4", "H5", 
    "CN", "CT", "CAU", "T1", "T2", 
    "REFH1", "REFH2", "BX1-TTL", "BX2-TTL", "BX1-H1"
]

# ── Step 5: Style that must follow every heading ──────────────────────────────
TXT_FLUSH_STYLE = ["TXT-FLUSH", "TXL"]

# Heading styles in level order (index 0 = H1)
HEADING_STYLES = [
    "H1", "H2", "H3",
    "H4", "H5", "H6",
]

# ── Step 7: Character style mapping ──────────────────────────────────────────
CHAR_STYLE_MAP_COMPREHENSIVE = {
    frozenset(('bold',)): r'bold',
    frozenset(('italic',)): r'italic',
    frozenset(('singleunderline',)): r'singleunderline',
    frozenset(('doubleunderline',)): r'doubleunderline',
    frozenset(('superscript',)): r'superscript',
    frozenset(('subscript',)): r'subscript',
    frozenset(('allcaps',)): r'allcaps',
    frozenset(('allcaps', 'smallcaps')): r'smallcaps',
    frozenset(('bold', 'italic')): r'bolditalics',
    frozenset(('bold', 'singleunderline')): r'boldsingleunderline',
    frozenset(('bold', 'doubleunderline')): r'bolddoubleunderline',
    frozenset(('bold', 'superscript')): r'boldsuperscript',
    frozenset(('bold', 'subscript')): r'boldsubscript',
    frozenset(('allcaps', 'bold')): r'boldallcaps',
    frozenset(('allcaps', 'bold', 'smallcaps')): r'boldsmallcaps',
    frozenset(('italic', 'singleunderline')): r'italicsingleunderline',
    frozenset(('doubleunderline', 'italic')): r'italicdoubleunderline',
    frozenset(('italic', 'superscript')): r'italicsuperscript',
    frozenset(('italic', 'subscript')): r'italicsubscript',
    frozenset(('allcaps', 'italic')): r'italicallcaps',
    frozenset(('allcaps', 'italic', 'smallcaps')): r'italicsmallcaps',
    frozenset(('singleunderline', 'superscript')): r'superscriptsingleunderline',
    frozenset(('singleunderline', 'subscript')): r'subscriptsingleunderline',
    frozenset(('allcaps', 'singleunderline')): r'singleunderlineallcaps',
    frozenset(('allcaps', 'singleunderline', 'smallcaps')): r'singleunderlinesmallcaps',
    frozenset(('doubleunderline', 'superscript')): r'superscriptdoubleunderline',
    frozenset(('doubleunderline', 'subscript')): r'subscriptdoubleunderline',
    frozenset(('allcaps', 'doubleunderline')): r'doubleunderlineallcaps',
    frozenset(('allcaps', 'doubleunderline', 'smallcaps')): r'doubleunderlinesmallcaps',
    frozenset(('allcaps', 'superscript')): r'superscriptallcaps',
    frozenset(('allcaps', 'smallcaps', 'superscript')): r'superscriptsmallcaps',
    frozenset(('allcaps', 'subscript')): r'subscriptallcaps',
    frozenset(('allcaps', 'smallcaps', 'subscript')): r'subscriptsmallcaps',
    frozenset(('bold', 'italic', 'singleunderline')): r'bolditalicsingleunderline',
    frozenset(('bold', 'doubleunderline', 'italic')): r'bolditalicdoubleunderline',
    frozenset(('bold', 'italic', 'superscript')): r'bolditalicsuperscipt',
    frozenset(('bold', 'italic', 'subscript')): r'bolditalicsubscript',
    frozenset(('allcaps', 'bold', 'italic')): r'bolditalicallcaps',
    frozenset(('allcaps', 'bold', 'italic', 'smallcaps')): r'bolditalicsmallcaps',
    frozenset(('italic', 'singleunderline', 'superscript')): r'italicsingleunderlinesuperscript',
    frozenset(('italic', 'singleunderline', 'subscript')): r'italicsingleunderlinesubscript',
    frozenset(('allcaps', 'italic', 'singleunderline')): r'italicsingleunderlineallcaps',
    frozenset(('allcaps', 'italic', 'singleunderline', 'smallcaps')): r'italicsingleunderlinesmallcaps',
    frozenset(('bold', 'doubleunderline', 'superscript')): r'doubleunderlinesuperscriptbold',
    frozenset(('doubleunderline', 'italic', 'superscript')): r'italicdoubleunderlinesuperscript',
    frozenset(('allcaps', 'doubleunderline', 'superscript')): r'superscriptallcapsdoubleunderline',
    frozenset(('allcaps', 'doubleunderline', 'smallcaps', 'superscript')): r'doubleunderlinesuperscriptsmallcaps',
    frozenset(('bold', 'singleunderline', 'superscript')): r'boldsingleunderlinesuperscript',
    frozenset(('allcaps', 'singleunderline', 'superscript')): r'superscriptallcapssingleunderline',
    frozenset(('allcaps', 'singleunderline', 'smallcaps', 'superscript')): r'singleunderlinesuperscriptsmallcaps',
    frozenset(('allcaps', 'singleunderline', 'smallcaps', 'subscript')): r'singleunderlinesubscriptsmallcaps',
    frozenset(('allcaps', 'bold', 'subscript')): r'subscriptallcapsbold',
    frozenset(('allcaps', 'italic', 'subscript')): r'italicsubscriptallcaps',
    frozenset(('allcaps', 'singleunderline', 'subscript')): r'subscriptallcapssingleunderline',
    frozenset(('allcaps', 'doubleunderline', 'subscript')): r'subscriptallcapsdoubleunderline',
    frozenset(('allcaps', 'bold', 'superscript')): r'boldsuperscriptallcaps',
    frozenset(('allcaps', 'italic', 'superscript')): r'italicsuperscriptallcaps',
    frozenset(('bold', 'singleunderline', 'subscript')): r'boldsingleunderlinesubscript',
    frozenset(('allcaps', 'bold', 'singleunderline')): r'boldsingleunderlineallcaps',
    frozenset(('allcaps', 'bold', 'singleunderline', 'smallcaps')): r'boldsingleunderlinesmallcaps',
    frozenset(('bold', 'doubleunderline', 'subscript')): r'bolddoubleunderlinesubscript',
    frozenset(('allcaps', 'bold', 'doubleunderline')): r'bolddoubleunderlineallcaps',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'smallcaps')): r'bolddoubleunderlinesmallcaps',
    frozenset(('allcaps', 'bold', 'smallcaps', 'superscript')): r'boldsuperscriptsmallcaps',
    frozenset(('allcaps', 'bold', 'smallcaps', 'subscript')): r'boldsubscriptsmallcaps',
    frozenset(('doubleunderline', 'italic', 'subscript')): r'italicdoubleunderlinesubscript',
    frozenset(('allcaps', 'doubleunderline', 'italic')): r'italicdoubleunderlineallcaps',
    frozenset(('allcaps', 'doubleunderline', 'italic', 'smallcaps')): r'italicdoubleunderlinesmallcaps',
    frozenset(('allcaps', 'italic', 'smallcaps', 'superscript')): r'italicsuperscriptsmallcaps',
    frozenset(('allcaps', 'italic', 'smallcaps', 'subscript')): r'italicsubscriptsmallcaps',
    frozenset(('bold', 'italic', 'singleunderline', 'superscript')): r'bolditalicsingleunderlinesuperscript',
    frozenset(('bold', 'italic', 'singleunderline', 'subscript')): r'bolditalicsingleunderlinesubscript',
    frozenset(('allcaps', 'bold', 'italic', 'singleunderline')): r'bolditalicsingleunderlineallcaps',
    frozenset(('allcaps', 'bold', 'italic', 'singleunderline', 'smallcaps')): r'bolditalicsingleunderlinesmallcaps',
    frozenset(('bold', 'doubleunderline', 'italic', 'superscript')): r'bolditalicdoubleunderlinesuperscript',
    frozenset(('bold', 'doubleunderline', 'italic', 'subscript')): r'bolditalicdoubleunderlinesubscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'italic')): r'bolditalicdoubleunderlineallcaps',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'italic', 'smallcaps')): r'bolditalicdoubleunderlinesmallcaps',
    frozenset(('allcaps', 'bold', 'italic', 'superscript')): r'bolditalicsuperscriptallcaps',
    frozenset(('allcaps', 'bold', 'italic', 'smallcaps', 'superscript')): r'bolditalicsuperscriptsmallcaps',
    frozenset(('allcaps', 'bold', 'italic', 'subscript')): r'bolditalicsubscriptallcaps',
    frozenset(('allcaps', 'bold', 'italic', 'smallcaps', 'subscript')): r'bolditalicsubscriptsmallcaps',
    frozenset(('allcaps', 'bold', 'singleunderline', 'superscript')): r'boldsingleunderlineallcapssuperscript',
    frozenset(('allcaps', 'bold', 'singleunderline', 'subscript')): r'boldsingleunderlineallcapssubscript',
    frozenset(('allcaps', 'bold', 'singleunderline', 'smallcaps', 'superscript')): r'boldsingleunderlinesmallcapssuperscript',
    frozenset(('allcaps', 'bold', 'singleunderline', 'smallcaps', 'subscript')): r'boldsingleunderlinesmallcapssubscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'superscript')): r'bolddoubleunderlineallcapssuperscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'subscript')): r'bolddoubleunderlineallcapssubscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'smallcaps', 'superscript')): r'bolddoubleunderlinesmallcapssuperscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'smallcaps', 'subscript')): r'bolddoubleunderlinesmallcapssubscript',
    frozenset(('allcaps', 'italic', 'singleunderline', 'superscript')): r'italicsingleunderlineallcapssuperscript',
    frozenset(('allcaps', 'italic', 'singleunderline', 'subscript')): r'italicsingleunderlineallcapssubscript',
    frozenset(('allcaps', 'italic', 'singleunderline', 'smallcaps', 'superscript')): r'italicsingleunderlinesmallcapssuperscript',
    frozenset(('allcaps', 'italic', 'singleunderline', 'smallcaps', 'subscript')): r'italicsingleunderlinesmallcapssubscript',
    frozenset(('allcaps', 'doubleunderline', 'italic', 'superscript')): r'italicdoubleunderlineallcapssuperscript',
    frozenset(('allcaps', 'doubleunderline', 'italic', 'subscript')): r'italicdoubleunderlineallcapssubscript',
    frozenset(('allcaps', 'doubleunderline', 'italic', 'smallcaps', 'superscript')): r'italicdoubleunderlinesmallcapssuperscript',
    frozenset(('allcaps', 'doubleunderline', 'italic', 'smallcaps', 'subscript')): r'italicdoubleunderlinesmallcapssubscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'italic', 'superscript')): r'bolditalicdoubleunderlineallcapssuperscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'italic', 'subscript')): r'bolditalicdoubleunderlineallcapssubscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'italic', 'smallcaps', 'superscript')): r'bolditalicdoubleunderlinesmallcapssuperscript',
    frozenset(('allcaps', 'bold', 'doubleunderline', 'italic', 'smallcaps', 'subscript')): r'bolditalicdoubleunderlinesmallcapssubscript',
    frozenset(('allcaps', 'bold', 'italic', 'singleunderline', 'superscript')): r'bolditalicsingleunderlineallcapssuperscript',
    frozenset(('allcaps', 'bold', 'italic', 'singleunderline', 'subscript')): r'bolditalicsingleunderlineallcapssubscript',
    frozenset(('allcaps', 'bold', 'italic', 'singleunderline', 'smallcaps', 'superscript')): r'bolditalicsingleunderlinesmallcapssuperscript',
    frozenset(('allcaps', 'bold', 'italic', 'singleunderline', 'smallcaps', 'subscript')): r'bolditalicsingleunderlinesmallcapssubscript',
}

# Reverse lookup: style name → frozenset of properties (used by step7 to auto-create styles)
STYLE_PROPERTY_MAP: dict[str, frozenset] = {
    v: k for k, v in CHAR_STYLE_MAP_COMPREHENSIVE.items()
}


# ── Step 8: Caption styles ────────────────────────────────────────────────────
FIGURE_CAPTION_STYLE = "FigureLegend", "FGC", "FIG-LEG"
TABLE_CAPTION_STYLE  = "TableCaption", "TT", "T1"

CAPTION_BOUNDARY_STYLES = {FIGURE_CAPTION_STYLE, TABLE_CAPTION_STYLE}

# Styles consumed into a caption group (source, footnote, abbrev lines)
CAPTION_GROUP_STYLES = {
    "FGS", "FIG-SRC", "FigureSource", "TableSource",
    "TableFootnote", "Abbreviation", "TableAbbreviation",
    "TSN", "TFN",
}

# Regex: matches "Figure 1" or "Figure 12.1"
FIGURE_CAPTION_RE = re.compile(
    r'Figure\s+(\d+)(?:\.(\d+))?', re.IGNORECASE)
TABLE_CAPTION_RE  = re.compile(
    r'Table\s+(\d+)(?:\.(\d+))?',  re.IGNORECASE)

# SDT alias/tag values
SDT_TAG_FIGURE   = "FigureCaption"
SDT_TAG_TABLE    = "TableCaption"
SDT_TAG_FIGREF   = "FigureRef"
SDT_TAG_TABLEREF = "TableRef"
SDT_TAG_TABLEGROUP = ""

# Inline cross-ref field markers
FIGUREREF_FIELD_MARKER  = "FigureRef"
TABLEREF_FIELD_MARKER   = "TableRef"

# BX group detection regex (for grouping BX1-*, BX2-*, ... BX25-* paragraph styles)
# Word strips underscores from style IDs (e.g. style name "BX3_T" → ID "BX3T"), so the
# separator is not reliable — match BX + digits only.
BX_STYLE_RE = re.compile(r'^BX(\d+)', re.IGNORECASE)

# ── Step 9: Symbol / Math → Unicode entity map ────────────────────────────────
# Replicates SymbolMathV4.Symbol2Unicode + Other_ent_unicode VBA modules.
# Keys are the literal Unicode characters; values are HTML numeric entities.
# Covers: mathematical operators, Greek uppercase and lowercase letters,
#         and the micro-sign (µ → &#x03BC;).
SYMBOL_MATH_MAP: dict[str, str] = {
    # ── General mathematical operators ───────────────────────────────────────
    "\u002B": "&#x002B;",   # +
    "\u00D7": "&#x00D7;",   # × (multiplication sign)
    "\u00F7": "&#x00F7;",   # ÷ (division sign)
    "\u003D": "&#x003D;",   # =
    "\u003C": "&#x003C;",   # <
    "\u003E": "&#x003E;",   # >
    "\u00B1": "&#x00B1;",   # ± (plus-minus)
    "\u2200": "&#x2200;",   # ∀ (for all)
    "\u2202": "&#x2202;",   # ∂ (partial differential)
    "\u2203": "&#x2203;",   # ∃ (there exists)
    "\u2205": "&#x2205;",   # ∅ (empty set)
    "\u2207": "&#x2207;",   # ∇ (nabla)
    "\u2208": "&#x2208;",   # ∈ (element of)
    "\u2209": "&#x2209;",   # ∉ (not element of)
    "\u220B": "&#x220B;",   # ∋ (contains)
    "\u220F": "&#x220F;",   # ∏ (n-ary product)
    "\u2211": "&#x2211;",   # ∑ (n-ary summation)
    "\u2212": "&#x2212;",   # − (minus sign)
    "\u2213": "&#x2213;",   # ∓ (minus-or-plus)
    "\u2217": "&#x2217;",   # ∗ (asterisk operator)
    "\u221A": "&#x221A;",   # √ (square root)
    "\u221D": "&#x221D;",   # ∝ (proportional to)
    "\u221E": "&#x221E;",   # ∞ (infinity)
    "\u2220": "&#x2220;",   # ∠ (angle)
    "\u2227": "&#x2227;",   # ∧ (logical and)
    "\u2228": "&#x2228;",   # ∨ (logical or)
    "\u2229": "&#x2229;",   # ∩ (intersection)
    "\u222A": "&#x222A;",   # ∪ (union)
    "\u222B": "&#x222B;",   # ∫ (integral)
    "\u2234": "&#x2234;",   # ∴ (therefore)
    "\u223C": "&#x223C;",   # ∼ (tilde operator)
    "\u2245": "&#x2245;",   # ≅ (approximately equal)
    "\u2248": "&#x2248;",   # ≈ (almost equal)
    "\u2260": "&#x2260;",   # ≠ (not equal)
    "\u2261": "&#x2261;",   # ≡ (identical to)
    "\u2264": "&#x2264;",   # ≤ (less-than or equal)
    "\u2265": "&#x2265;",   # ≥ (greater-than or equal)
    "\u2282": "&#x2282;",   # ⊂ (subset of)
    "\u2283": "&#x2283;",   # ⊃ (superset of)
    "\u2284": "&#x2284;",   # ⊄ (not subset)
    "\u2285": "&#x2285;",   # ⊅ (not superset)
    "\u2286": "&#x2286;",   # ⊆ (subset or equal)
    "\u2287": "&#x2287;",   # ⊇ (superset or equal)
    "\u2295": "&#x2295;",   # ⊕ (circled plus)
    "\u2297": "&#x2297;",   # ⊗ (circled times)
    "\u22A5": "&#x22A5;",   # ⊥ (perpendicular)
    "\u22C5": "&#x22C5;",   # ⋅ (dot operator)

    # ── Greek uppercase ───────────────────────────────────────────────────────
    "\u0393": "&#x0393;",   # Γ
    "\u0394": "&#x0394;",   # Δ
    "\u0395": "&#x0395;",   # Ε
    "\u0396": "&#x0396;",   # Ζ
    "\u0397": "&#x0397;",   # Η
    "\u0398": "&#x0398;",   # Θ
    "\u0399": "&#x0399;",   # Ι
    "\u039A": "&#x039A;",   # Κ
    "\u039B": "&#x039B;",   # Λ
    "\u039C": "&#x039C;",   # Μ
    "\u039D": "&#x039D;",   # Ν
    "\u039E": "&#x039E;",   # Ξ
    "\u039F": "&#x039F;",   # Ο
    "\u03A0": "&#x03A0;",   # Π
    "\u03A1": "&#x03A1;",   # Ρ
    "\u03A3": "&#x03A3;",   # Σ
    "\u03A4": "&#x03A4;",   # Τ
    "\u03A5": "&#x03A5;",   # Υ
    "\u03A6": "&#x03A6;",   # Φ
    "\u03A7": "&#x03A7;",   # Χ
    "\u03A8": "&#x03A8;",   # Ψ
    "\u03A9": "&#x03A9;",   # Ω

    # ── Greek lowercase ───────────────────────────────────────────────────────
    "\u03B1": "&#x03B1;",   # α
    "\u03B2": "&#x03B2;",   # β
    "\u03B3": "&#x03B3;",   # γ
    "\u03B4": "&#x03B4;",   # δ
    "\u03B5": "&#x03B5;",   # ε
    "\u03B6": "&#x03B6;",   # ζ
    "\u03B7": "&#x03B7;",   # η
    "\u03B8": "&#x03B8;",   # θ
    "\u03B9": "&#x03B9;",   # ι
    "\u03BA": "&#x03BA;",   # κ
    "\u03BB": "&#x03BB;",   # λ
    "\u03BC": "&#x03BC;",   # μ (mu)
    "\u03BD": "&#x03BD;",   # ν
    "\u03BE": "&#x03BE;",   # ξ
    "\u03BF": "&#x03BF;",   # ο
    "\u03C0": "&#x03C0;",   # π
    "\u03C1": "&#x03C1;",   # ρ
    "\u03C2": "&#x03C2;",   # ς (final sigma)
    "\u03C3": "&#x03C3;",   # σ
    "\u03C4": "&#x03C4;",   # τ
    "\u03C5": "&#x03C5;",   # υ
    "\u03C6": "&#x03C6;",   # φ
    "\u03C7": "&#x03C7;",   # χ
    "\u03C8": "&#x03C8;",   # ψ
    "\u03C9": "&#x03C9;",   # ω
    "\u03D2": "&#x03D2;",   # ϒ (upsilon hook)
    "\u03D6": "&#x03D6;",   # ϖ (pi symbol)

    # ── Micro sign (alternative to μ, commonly in units like µg, µm) ─────────
    "\u00E6": "&#x03BC;",   # æ → µ (legacy Symbol-font micro-sign mapping)
}

# ── Step 10: Math/Symbol entity → character style mapping ────────────────────

# Reverse of SYMBOL_MATH_MAP: entity string → original Unicode char.
# First-wins on duplicates so &#x03BC; → μ (not æ, the legacy micro mapping).
_ENTITY_TO_CHAR: dict[str, str] = {}
for _c, _e in SYMBOL_MATH_MAP.items():
    if _e not in _ENTITY_TO_CHAR:
        _ENTITY_TO_CHAR[_e] = _c
ENTITY_TO_CHAR: dict[str, str] = _ENTITY_TO_CHAR

# Entity strings that get Symbol* character styles (Greek letters + plus sign).
# &#x002B; appears in both VBA macros; Symbol() runs after Math() so Symbol wins.
SYMBOL_ENTITIES: frozenset[str] = frozenset({
    "&#x002B;",
    "&#x0393;","&#x0394;","&#x0395;","&#x0396;","&#x0397;","&#x0398;","&#x0399;",
    "&#x039A;","&#x039B;","&#x039C;","&#x039D;","&#x039E;","&#x039F;","&#x03A0;",
    "&#x03A1;","&#x03A3;","&#x03A4;","&#x03A5;","&#x03A6;","&#x03A7;","&#x03A8;","&#x03A9;",
    "&#x03B1;","&#x03B2;","&#x03B3;","&#x03B4;","&#x03B5;","&#x03B6;","&#x03B7;","&#x03B8;",
    "&#x03B9;","&#x03BA;","&#x03BB;","&#x03BC;","&#x03BD;","&#x03BE;","&#x03BF;","&#x03C0;",
    "&#x03C1;","&#x03C2;","&#x03C3;","&#x03C4;","&#x03C5;","&#x03C6;","&#x03C7;","&#x03C8;",
    "&#x03C9;","&#x03D2;","&#x03D6;",
})

# Entity strings that get Math* character styles (mathematical operators).
MATH_ENTITIES: frozenset[str] = frozenset({
    "&#x00D7;","&#x00F7;","&#x003D;","&#x003C;","&#x003E;","&#x00B1;",
    "&#x2200;","&#x2202;","&#x2203;","&#x2205;","&#x2207;","&#x2208;","&#x2209;","&#x220B;",
    "&#x220F;","&#x2211;","&#x2212;","&#x2213;","&#x2217;","&#x221A;","&#x221D;","&#x221E;",
    "&#x2220;","&#x2227;","&#x2228;","&#x2229;","&#x222A;","&#x222B;","&#x2234;","&#x223C;",
    "&#x2245;","&#x2248;","&#x2260;","&#x2261;","&#x2264;","&#x2265;",
    "&#x2282;","&#x2283;","&#x2284;","&#x2285;","&#x2286;","&#x2287;",
    "&#x2295;","&#x2297;","&#x22A5;","&#x22C5;",
})

# (bold, italic, superscript, subscript) → style-name suffix
_STYLE_SUFFIX: dict[tuple, str] = {
    (False, False, False, False): "",
    (False, False, True,  False): "Sup",
    (False, False, False, True):  "Sub",
    (True,  False, False, False): "B",
    (False, True,  False, False): "I",
    (True,  True,  False, False): "BI",
    (True,  True,  True,  False): "BISup",
    (True,  True,  False, True):  "BISub",
    (True,  False, True,  False): "BSup",
    (True,  False, False, True):  "BSub",
    (False, True,  True,  False): "ISup",
    (False, True,  False, True):  "ISub",
}

MATH_STYLE_MAP:   dict[tuple, str] = {k: ("Math"   + v) for k, v in _STYLE_SUFFIX.items()}
SYMBOL_STYLE_MAP: dict[tuple, str] = {k: ("Symbol" + v) for k, v in _STYLE_SUFFIX.items()}

# Shading colours for Math/Symbol styles, keyed by the same (b,i,su,sb) tuple.
# Base (no formatting) uses family-specific colours; variants follow ApplyBookcolor rules.
_STYLE_SHADING_BASE: dict[tuple[bool, ...], str] = {
    (False, False, True,  False): "CCFFFF",   # Light Cyan   — Sup
    (False, False, False, True):  "FFCCFF",   # Light Pink   — Sub
    (True,  False, False, False): "CCE0FF",   # Light Blue   — B
    (False, True,  False, False): "CCFFCC",   # Light Green  — I
    (True,  True,  False, False): "FFCCCC",   # Light Red    — BI
    (True,  True,  True,  False): "FFCCCC",   # Light Red    — BISup
    (True,  True,  False, True):  "FFCCCC",   # Light Red    — BISub
    (True,  False, True,  False): "CCE0FF",   # Light Blue   — BSup
    (True,  False, False, True):  "CCE0FF",   # Light Blue   — BSub
    (False, True,  True,  False): "FFFFCC",   # Light Yellow — ISup
    (False, True,  False, True):  "CCFFCC",   # Light Green  — ISub
}
MATH_STYLE_SHADING:   dict[tuple[bool, ...], str] = {**_STYLE_SHADING_BASE, (False, False, False, False): "FFCC99"}
SYMBOL_STYLE_SHADING: dict[tuple[bool, ...], str] = {**_STYLE_SHADING_BASE, (False, False, False, False): "CCFFCC"}


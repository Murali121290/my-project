"""Bias terms (JBL-style inclusive-language flags) and article errors."""
from __future__ import annotations
import re
from typing import Iterable
from manuscript_core.extractor import Segment
from manuscript_core.rules.base import Finding, context_snippet, iter_unmasked_matches
# Each entry is (pattern_source, label, suggestion).
# Patterns are matched case-insensitively with word boundaries.
BIAS_TERMS: tuple[tuple[str, str, str], ...] = (
    (r"elderly", "elderly", "older adults"),
    (r"the\s+aged", "the aged", "older adults"),
    (r"aging\s+dependents", "aging dependents", "older adults who need support"),
    (r"old-old", "old-old", "adults aged 85+"),
    (r"young-old", "young-old", "adults aged 65-74"),
    (r"the\s+poor", "the poor", "people with low incomes"),
    (r"the\s+unemployed", "the unemployed", "people who are unemployed"),
    (r"the\s+blind", "the blind", "people who are blind"),
    (r"the\s+visually\s+impaired", "the visually impaired", "people with visual impairment"),
    (r"the\s+deaf", "the deaf", "people who are deaf"),
    (r"the\s+disabled", "the disabled", "people with disabilities"),
    (r"the\s+handicapped", "the handicapped", "people with disabilities"),
    (r"the\s+infirm", "the infirm", "people in poor health"),
    (r"disabled\s+child", "disabled child", "child with a disability"),
    (r"mentally\s+ill\s+person", "mentally ill person", "person with a mental health condition"),
    (r"retarded", "retarded", "person with an intellectual disability"),
    (r"crippled", "crippled", "person with a physical disability"),
    (r"\blame\b", "lame", "person with a mobility impairment"),
    (r"deformed", "deformed", "person with a physical difference"),
    (r"confined\s+to\s+(?:a\s+)?wheelchair", "confined to (a) wheelchair", "wheelchair user"),
    (r"bound\s+to\s+(?:a\s+)?wheelchair", "bound to (a) wheelchair", "wheelchair user"),
    (r"\balcoholic\b", "alcoholic", "person with alcohol use disorder"),
    (r"\baddict\b", "addict", "person with substance use disorder"),
    (r"\babuser\b", "abuser", "person who uses drugs"),
    (r"drug\s+abuse", "drug abuse", "substance use"),
    (r"\basthmatics\b", "asthmatics", "people with asthma"),
    (r"\bdiabetics\b", "diabetics", "people with diabetes"),
    (r"\bepileptic\b", "epileptic", "person with epilepsy"),
    (r"\bvictim\b", "victim", "survivor / person affected by"),
    (r"\bvictims\b", "victims", "survivors / people affected by"),
    (r"Caucasians?", "Caucasian(s)", "White / specify nationality"),
    # Gendered occupation terms
    (r"chairman", "chairman", "chair / chairperson"),
    (r"chairmen", "chairmen", "chairs"),
    (r"chairwoman", "chairwoman", "chair / chairperson"),
    (r"corpsman", "corpsman", "corps member"),
    (r"fireman", "fireman", "firefighter"),
    (r"firemen", "firemen", "firefighters"),
    (r"layman", "layman", "layperson"),
    (r"laymen", "laymen", "laypeople"),
    (r"mailman", "mailman", "mail carrier"),
    (r"mailmen", "mailmen", "mail carriers"),
    (r"\bmankind\b", "mankind", "humanity / humankind"),
    (r"manmade", "manmade", "artificial / synthetic"),
    (r"\bman-made\b", "man-made", "artificial / synthetic"),
    (r"manpower", "manpower", "workforce / staffing"),
    (r"\bmothering\b", "mothering", "parenting"),
    (r"policeman", "policeman", "police officer"),
    (r"policemen", "policemen", "police officers"),
    (r"policewoman", "policewoman", "police officer"),
    (r"spokesman", "spokesman", "spokesperson"),
    (r"spokesmen", "spokesmen", "spokespeople"),
    (r"spokeswoman", "spokeswoman", "spokesperson"),
    (r"\bstewardess\b", "stewardess", "flight attendant"),
    (r"chairwomen", "chairwomen", "chairs"),
    (r"corpsmen", "corpsmen", "corps members"),
    (r"policewomen", "policewomen", "police officers"),
    (r"spokeswomen", "spokeswomen", "spokespeople"),
    (r"\bhe or she\b", "he or she", "they"),
    (r"\bhe/she\b", "he/she", "they"),
    (r"\bhis/her\b", "his/her", "their"),
    (r"\bhim/her\b", "him/her", "them"),
)
def _build_bias_patterns() -> list[tuple[re.Pattern, str, str]]:
    out = []
    for src, label, suggestion in BIAS_TERMS:
        # Ensure word boundaries unless pattern already contains them.
        pat = src if r"\b" in src else r"\b" + src + r"\b"
        out.append((re.compile(pat, re.IGNORECASE), label, suggestion))
    return out
_BIAS_PATTERNS = _build_bias_patterns()
def run_bias_rules(seg: Segment) -> Iterable[Finding]:
    for pat, label, suggestion in _BIAS_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="bias",
                rule_id=f"bias_{label.replace(' ', '_').replace('(', '').replace(')', '')}",
                rule_label=f"Bias term: {label} â†’ {suggestion}",
                surface=m.group(0),
                canonical=label.lower(),
                chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name,
                source=seg.source,
                page=seg.page,
                para_index=seg.para_index,
                context=context_snippet(seg.text, m.start(), m.end()),
                severity="warn",
                replacement=suggestion,
            )
# ---------------------------------------------------------------------------
# Article a/an errors
# ---------------------------------------------------------------------------
# Common cases from the brief. We check the literal following token.
# We flag BOTH directions: "an" before a consonant-sound word, "a" before
# a vowel-sound word Ã¢â‚¬â€ but only for cases we're reasonably confident about.
AN_WRONG = {
    # Words starting with a consonant sound despite the vowel letter
    "eukaryote",
    "eulogy",
    "european",
    "historic",
    "histogram",
    "history",
    "laryngoscope",
    "mammogram",
    "neurologist",
    "user",
    "ultrasound",
    "unit",
    "university",
    "uniform",
    "unique",
    "urine",
    "xenograft",
}
A_WRONG = {
    # Words starting with a vowel sound despite the consonant letter
    # (honor, hour, etc.) or that start with a vowel letter
    "eye",
    "hour",
    "honor",
    "honest",
    "mmse",
    "nsaid",
    "otoscope",
    "ulcer",
    "x-ray",
    "xray",
    "mri",
    "mra",
    "mrsa",
    "nicu",
    "icu",
    "iv",
}
ARTICLE_PATTERN = re.compile(r"\b(a|an)\s+([A-Za-z][A-Za-z0-9\-]*)", re.IGNORECASE)
def run_article_rules(seg: Segment) -> Iterable[Finding]:
    for m in iter_unmasked_matches(ARTICLE_PATTERN, seg.text, seg.mask):
        article = m.group(1).lower()
        following = m.group(2).lower()
        wrong = False
        expected = None
        if article == "an" and following in AN_WRONG:
            wrong = True
            expected = "a"
        elif article == "a" and following in A_WRONG:
            wrong = True
            expected = "an"
        if not wrong:
            continue
        yield Finding(
            category="article",
            rule_id="article_an_a_mismatch",
            rule_label=f"'{article}' before '{following}' (use '{expected}')",
            surface=m.group(0),
            canonical=f"{article}_{following}",
            chapter_index=seg.chapter_index,
            chapter_name=seg.chapter_name,
            source=seg.source,
            page=seg.page,
            para_index=seg.para_index,
            context=context_snippet(seg.text, m.start(), m.end()),
            severity="error",
            replacement=f"{expected} {following}",
        )
# ---------------------------------------------------------------------------
# Country Style
# ---------------------------------------------------------------------------
COUNTRY_PATTERNS = [
    (re.compile(r"\bU\.S(?!\.)\b"), "U.S", "country_style"),
    (re.compile(r"\bU\.S\.\b"), "U.S.", "country_style"),
    (re.compile(r"\bUS\b"), "US", "country_style"),
    (re.compile(r"\bUnited States\b", re.IGNORECASE), "United States", "country_style"),
]
def run_country_style(seg: Segment) -> Iterable[Finding]:
    for pat, canonical, rule_id in COUNTRY_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="te_point", rule_id=rule_id, rule_label="US country style",
                surface=m.group(0), canonical="country_style", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
            )
# ---------------------------------------------------------------------------
# Death Euphemism
# ---------------------------------------------------------------------------
DEATH_PATTERNS = [
    (re.compile(r"\bexpired\b", re.IGNORECASE), "expired"),
    (re.compile(r"\bpassed away\b", re.IGNORECASE), "passed away"),
    (re.compile(r"\bpassed\b", re.IGNORECASE), "passed"),
    (re.compile(r"\bsuccumbed\b", re.IGNORECASE), "succumbed"),
    (re.compile(r"\bsacrificed\b", re.IGNORECASE), "sacrificed"),
]
def run_death_euphemism(seg: Segment) -> Iterable[Finding]:
    for pat, label in DEATH_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="bias", rule_id="death_euphemism", rule_label=f"Death euphemism: {label} -> died",
                surface=m.group(0), canonical="death_euphemism", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
                severity="warn"
            )
# ---------------------------------------------------------------------------
# Clinical Jargon
# ---------------------------------------------------------------------------
JARGON_TERMS = [
    (re.compile(r"\bpreemie\b", re.IGNORECASE), "preterm infant"),
    (re.compile(r"\bpreop\b", re.IGNORECASE), "preoperative"),
    (re.compile(r"\bpostop\b", re.IGNORECASE), "postoperative"),
    (re.compile(r"\bprepped\b", re.IGNORECASE), "prepared"),
    (re.compile(r"\blab\b", re.IGNORECASE), "laboratory"),
    (re.compile(r"\bflu\b", re.IGNORECASE), "influenza"),
    (re.compile(r"\bexam\b", re.IGNORECASE), "examination"),
    (re.compile(r"\bsymptomatology\b", re.IGNORECASE), "symptoms"),
    (re.compile(r"\bin order to\b", re.IGNORECASE), "to"),
    (re.compile(r"\byears of age\b", re.IGNORECASE), "aged X years"),
    (re.compile(r"\bmonths of age\b", re.IGNORECASE), "aged X months"),
    (re.compile(r"\bwith the exception of\b", re.IGNORECASE), "except for"),
    (re.compile(r"\bcomprised of\b", re.IGNORECASE), "comprises / composed of"),
    (re.compile(r"\bdie from\b", re.IGNORECASE), "die of"),
    (re.compile(r"\bdied from\b", re.IGNORECASE), "died of"),
    (re.compile(r"\bthe patient failed treatment\b", re.IGNORECASE), "treatment failed"),
]
def run_clinical_jargon(seg: Segment) -> Iterable[Finding]:
    for pat, suggestion in JARGON_TERMS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="bias", rule_id="clinical_jargon", rule_label=f"Clinical jargon -> {suggestion}",
                surface=m.group(0), canonical="clinical_jargon", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
                severity="warn"
            )
# ---------------------------------------------------------------------------
# Subject Terms
# ---------------------------------------------------------------------------
SUBJECT_PATTERNS = [
    (re.compile(r"\bsubject[s]?\b", re.IGNORECASE), "participant(s)"),
    (re.compile(r"\bcase[s]?\b", re.IGNORECASE), "patient(s) / participant(s)"),
]
def run_subject_terms(seg: Segment) -> Iterable[Finding]:
    for pat, suggestion in SUBJECT_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="bias", rule_id="subject_terms", rule_label=f"Subject term -> {suggestion}",
                surface=m.group(0), canonical="subject_terms", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
                severity="info"
            )
# ---------------------------------------------------------------------------
# Eponym Style
# ---------------------------------------------------------------------------
EPONYM_PATTERNS = [
    (re.compile(r"\b\w+\'s (?:disease|syndrome|sign|test|node)\b", re.IGNORECASE), "with 's"),
    (re.compile(r"\b\w+(?<!s) (?:disease|syndrome|sign|test|node)\b", re.IGNORECASE), "without 's"),
    (re.compile(r"\bSt\.? John wort\b", re.IGNORECASE), "St John's wort (suggestion)"),
]
def run_eponym_style(seg: Segment) -> Iterable[Finding]:
    for pat, label in EPONYM_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="te_point", rule_id="eponym_style", rule_label="Eponym style",
                surface=m.group(0), canonical="eponym_style", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
            )
# ---------------------------------------------------------------------------
# Wrong Usage
# ---------------------------------------------------------------------------
WRONG_USAGE_TERMS = [
    (re.compile(r"\bUnite States\b"), "United States"),
    (re.compile(r"\bpubic\b(?!\s+(?:hair|lice|region|symphysis))", re.IGNORECASE), "public"),
    (re.compile(r"\bheath\b", re.IGNORECASE), "health"),
    (re.compile(r"\bmanger\b", re.IGNORECASE), "manager"),
    (re.compile(r"\bBrasil\b"), "Brazil"),
    (re.compile(r"\bTrendelenberg\b"), "Trendelenburg"),
    (re.compile(r"\bHbA1C\b"), "HbA1c"),
    (re.compile(r"\bHBA1c\b"), "HbA1c"),
]
def run_wrong_usage(seg: Segment) -> Iterable[Finding]:
    for pat, suggestion in WRONG_USAGE_TERMS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="error", rule_id="wrong_usage", rule_label=f"Wrong usage -> {suggestion}",
                surface=m.group(0), canonical="wrong_usage", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
                severity="error", replacement=suggestion
            )
# ---------------------------------------------------------------------------
# Punctuation Pattern
# ---------------------------------------------------------------------------
PUNC_PATTERNS = [
    (re.compile(r'\."'), '."'),
    (re.compile(r'"\.'), '".'),
    (re.compile(r'\,"'), ',"'),
    (re.compile(r'"\,'), '",'),
    (re.compile(r'\?"'), '?"'),
    (re.compile(r'"\?'), '"?'),
    (re.compile(r'\!"'), '!"'),
    (re.compile(r'"\!'), '"!'),
]
def run_punctuation_style(seg: Segment) -> Iterable[Finding]:
    for pat, label in PUNC_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="te_point", rule_id="punctuation_style", rule_label="Punctuation inside/outside quotes",
                surface=m.group(0), canonical="punctuation_style", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
            )
# ---------------------------------------------------------------------------
# Sic and Special
# ---------------------------------------------------------------------------
SIC_PATTERNS = [
    (re.compile(r"\[sic\]", re.IGNORECASE), "[sic]"),
    (re.compile(r'so-called\s+"[^"]+"', re.IGNORECASE), 'so-called "..."'),
    (re.compile(r'"yes"', re.IGNORECASE), '"yes"'),
    (re.compile(r'"no"', re.IGNORECASE), '"no"'),
]
def run_sic_special(seg: Segment) -> Iterable[Finding]:
    for pat, label in SIC_PATTERNS:
        for m in iter_unmasked_matches(pat, seg.text, seg.mask):
            yield Finding(
                category="te_point", rule_id="sic_special", rule_label="Sic / Quoted Special",
                surface=m.group(0), canonical="sic_special", chapter_index=seg.chapter_index,
                chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
                para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
                severity="info"
            )
def run_pronoun_check(seg: Segment) -> Iterable[Finding]:
    # Placeholder for single pronouns (low severity)
    # the "he or she" are already in bias terms
    for m in iter_unmasked_matches(re.compile(r"\b(he|she|his|her|him)\b", re.IGNORECASE), seg.text, seg.mask):
        yield Finding(
            category="bias", rule_id="pronoun_check", rule_label="Generic pronoun",
            surface=m.group(0), canonical="pronoun_check", chapter_index=seg.chapter_index,
            chapter_name=seg.chapter_name, source=seg.source, page=seg.page,
            para_index=seg.para_index, context=context_snippet(seg.text, m.start(), m.end()),
            severity="info"
        )

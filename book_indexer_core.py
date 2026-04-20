import re

PROMPT_TEMPLATE = """\
You are a professional technical book indexer trained in Nancy Mulvany's indexing principles, \
specialised in engineering publications covering antenna design, RF engineering, \
mechanical structures, materials science, and manufacturing.

Your task: produce a publication-ready back-of-book index from the TEXT SEGMENT below, \
which contains [PAGE N] markers indicating page boundaries.

════════════════════════════════════════════
RULE 1 — INDEX CONCEPTS, NOT JUST WORDS
════════════════════════════════════════════
- Identify meaningful engineering concepts, not raw terms.
- Avoid indexing trivial or passing mentions.
- Group related ideas under a single main heading.
- Only include terms if they contribute to understanding antenna design, RF behavior,
  or manufacturing — and are substantively explained, defined, or technically discussed.

════════════════════════════════════════════
RULE 2 — PER-PAGE CONTENT VERIFICATION
════════════════════════════════════════════
Before assigning any page number, ask:
  "Does this page substantively explain, define, analyse, or describe this term?"
Include the page number ONLY if YES. Reject if the term:
  - appears only in passing (single-sentence mention while discussing something else)
  - appears only in a table cell, header, or caption without surrounding discussion
  - appears only in a footnote, citation, or cross-reference
  - appears only in a list alongside many other items without elaboration
If uncertain, omit the page rather than guess.

════════════════════════════════════════════
RULE 3 — PAGE RANGE ACCURACY
════════════════════════════════════════════
- Use ONLY the [PAGE N] markers in the text. Do not infer or extrapolate.
- Use ranges (e.g., 12–15) ONLY for truly continuous discussion.
- Non-consecutive pages: list separately (e.g., 12, 14, 18).
- A wrong page number is worse than an omission. Double-check every number.

════════════════════════════════════════════
RULE 4 — HIERARCHICAL STRUCTURE
════════════════════════════════════════════
- Main entries: flush left. Use specific, meaningful terms only.
  Avoid broad generics: "Alignment", "Alloys", "Materials", "Design" — unless a focused,
  substantive discussion exists.
- Sub-entries: 2-space indent. Must belong logically to the parent.
- Maximum 2 levels of nesting unless absolutely necessary.
- When a main entry has TWO OR MORE sub-entries:
    → Do NOT list a page range on the main entry line.
    → Sub-entries carry all page detail.
    → Exception: a page discussing the topic as a whole (not covered by any sub-entry)
      may appear on the main entry line.
- Treat antenna parameters as structured entries:
    Antenna parameters
      impedance
      polarization
      VSWR

════════════════════════════════════════════
RULE 5 — TERM SELECTION
════════════════════════════════════════════
INCLUDE:
  - Technical concepts: impedance, polarization, VSWR, dielectric breakdown, radiation pattern
  - Processes: additive manufacturing, machining, CNC milling, electroforming
  - Components: waveguides, connectors, feed networks, radomes
  - Materials: aluminum, copper, composites, dielectric substrates
  - Phenomena: multipath interference, cross-polarization, sidelobe suppression
  - Named techniques, environments, and classification breakdowns with body discussion

EXCLUDE:
  - References, bibliography, URLs, citation-only mentions
  - Generic words: "about", "introduction", "overview", "summary"
  - Table-only terms without surrounding prose discussion
  - Passing mentions, figures/table titles (unless conceptually significant)

════════════════════════════════════════════
RULE 6 — NORMALIZATION & SYNONYM MERGING
════════════════════════════════════════════
- Merge synonyms under one canonical form:
    "VSWR" + "Voltage standing wave ratio" → "Voltage standing wave ratio (VSWR)"
- Use consistent naming (not both "cross-pol" and "cross-polarization").
- Add a "See" cross-reference for the non-preferred synonym.

════════════════════════════════════════════
RULE 7 — CROSS-REFERENCES
════════════════════════════════════════════
- "Term. See Preferred Term" — ONLY when "Preferred Term" exists as a main entry.
- "Term. See also Related Term" — ONLY when both appear as main entries and the
  relationship is clearly meaningful (engineering cause-effect or conceptual link).
- If the preferred term does NOT appear as a main entry, index the term directly.
Example:
  Impedance
    capacitive
    inductive
    See also VSWR

════════════════════════════════════════════
RULE 8 — AVOID REDUNDANCY
════════════════════════════════════════════
- Do not repeat the same page references across similar entries unnecessarily.
- Do not create duplicate main entries for the same concept.
- Remove noise entries that add no indexing value.

════════════════════════════════════════════
RULE 9 — DO NOT MISS IMPORTANT ENTRIES
════════════════════════════════════════════
- Index every named environment, technique, classification, or component receiving
  substantive discussion, even if only in a section heading or standalone paragraph.
- Capture engineering relationships (cause-effect) where the text explains them.
- For topics spanning large page ranges (section-level), include the full range.

════════════════════════════════════════════
OUTPUT FORMAT (follow exactly — no deviations)
════════════════════════════════════════════
  Main Entry, page(s)
    subentry, page(s)
    subentry, page(s)
  Main Entry with sub-entries
    subentry one, page(s)
    subentry two, page(s)
  Synonym. See Preferred Term
  Term. See also Related Term

EXAMPLE OUTPUT:
  Additive manufacturing, 120–125
    advantages, 122
    limitations, 124
    See also Fabrication processes
  Bandwidth (BW), 8–9
  Cross-polarization
    mechanical drivers, 17–18
    requirements, 18
  Voltage standing wave ratio (VSWR), 45, 48–50
  VSWR. See Voltage standing wave ratio (VSWR)

CRITICAL OUTPUT RULES (STRICTLY ENFORCE):
- Output ONLY index lines in the format above. Nothing else.
- Do NOT output any headings, titles, numbered lists, bullet points, or markdown.
- Do NOT explain your reasoning or add any commentary or preamble.
- Do NOT use **, *, ##, ---, >, or any other markdown syntax.
- Entries must be in alphabetical order.
- Start your response immediately with the first index entry.

TEXT SEGMENT:
{text}

OUTPUT: Index entries ONLY — no preamble, no markdown, no explanations."""

def clean_llm_response(text):
    """
    Strip markdown, meta-commentary, and instruction echoes that the LLM
    sometimes emits before / instead of clean index lines.
    Keeps only lines that look like genuine index entries.
    """
    out = []
    skip_patterns = [
        r'^#{1,6}\s',           # markdown headings
        r'^\*{1,3}[^*]',        # bullet points / bold text
        r'^\*{1,3}$',           # lone asterisks
        r'^-{2,}',              # horizontal rules
        r'^```',                # code fences
        r'^>',                  # blockquotes
        r'\*\*.*\*\*',          # inline bold
        r'awaiting.*(input|your)',   # "awaiting your input"
        r'^\s*(processing|indexing process|chapter pre-grouping|input received|waiting for input|h1:|h2:|generated index|indexing process|example of processing)',  # meta labels
    ]
    combined = re.compile('|'.join(skip_patterns), re.IGNORECASE)
    for line in text.split('\n'):
        stripped = line.strip()
        if not stripped:
            out.append('')
            continue
        if combined.search(stripped):
            continue  # skip garbage line
        # Strip residual inline bold/italic markers
        cleaned = re.sub(r'[*_`]{1,3}', '', line)
        out.append(cleaned)
    return '\n'.join(out)

def parse_partial_index(text, merged_index):
    current_main = None

    for line in text.split('\n'):
        if not line.strip() or line.strip().upper() == "INDEX" or line.strip().startswith('```'):
            continue

        is_sub = line.startswith(' ') or line.startswith('\t')
        clean_line = line.strip()

        if len(clean_line) > 150:  # ignore garbage/instructions
            continue

        # Handle cross-references — check "See also" before "See" (substring order)
        see_also_m = re.match(r'^(.+?)\.\s+[Ss]ee also\s+(.+)$', clean_line)
        see_m = re.match(r'^(.+?)\.\s+[Ss]ee\s+(.+)$', clean_line) if not see_also_m else None

        if see_also_m:
            term = see_also_m.group(1).strip()
            refs = [r.strip() for r in see_also_m.group(2).split(',') if r.strip()]
            term_key = term.lower()
            if term_key not in merged_index:
                merged_index[term_key] = {'display': term, 'pages': set(), 'sub': {}, 'see': None, 'see_also': []}
            merged_index[term_key].setdefault('see_also', [])
            for ref in refs:
                if ref not in merged_index[term_key]['see_also']:
                    merged_index[term_key]['see_also'].append(ref)
            current_main = term_key
            continue
        elif see_m:
            term = see_m.group(1).strip()
            ref = see_m.group(2).strip()
            term_key = term.lower()
            if term_key not in merged_index:
                merged_index[term_key] = {'display': term, 'pages': set(), 'sub': {}, 'see': None, 'see_also': []}
            merged_index[term_key]['see'] = ref
            current_main = term_key
            continue

        # Separate term and page numbers using regex
        m = re.search(r',\s*([\d\s,\-–]+)$', clean_line)
        if m:
            term = clean_line[:m.start()].strip()
            pages_str = m.group(1)
            pages = set()
            for p in re.split(r'[,\s]+', pages_str):
                p = p.strip()
                if not p: continue
                if '-' in p or '–' in p:
                    parts = re.split(r'[\-–]', p)
                    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                        pages.update(range(int(parts[0]), int(parts[1]) + 1))
                elif p.isdigit():
                    pages.add(int(p))
        else:
            term = clean_line
            pages = set()

        term_key = term.lower()
        if is_sub and current_main:
            if term_key not in merged_index[current_main]['sub']:
                merged_index[current_main]['sub'][term_key] = {'display': term, 'pages': set()}
            merged_index[current_main]['sub'][term_key]['pages'].update(pages)
        else:
            if term_key not in merged_index:
                merged_index[term_key] = {'display': term, 'pages': set(), 'sub': {}, 'see': None, 'see_also': []}
            merged_index[term_key]['pages'].update(pages)
            current_main = term_key

def format_pages(pages_set):
    if not pages_set:
        return ""
    pages = sorted(list(pages_set))
    ranges = []
    if not pages:
        return ""
    s = e = pages[0]
    for n in pages[1:]:
        if n == e + 1:
            e = n
        else:
            ranges.append(f"{s}–{e}" if e > s else f"{s}")
            s = e = n
    ranges.append(f"{s}–{e}" if e > s else f"{s}")
    return ", ".join(ranges)

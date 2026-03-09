import re
import logging
from extractor import CAPTION_START_REGEX, CAPTION_PATTERNS, _extract_inline_credit, extract_credit_from_text, extract_figures_tables, extract_boxes, CREDIT_KEYWORDS_REGEX
from extractor import CAPTION_START_REGEX, CAPTION_PATTERNS, _extract_inline_credit, extract_credit_from_text, extract_figures_tables

logging.basicConfig(level=logging.DEBUG)

paragraphs1 = [
    "Figures 10.8. Specialized sensory endings in skeletal muscle and tendon. Sensory axons are shown in shades of blue, fusimotor axons in red, muscle fibers in yellow, and connective tissue in black and gray. (A) A Golgi tendon organ. (B) Neuromuscular spindle in transverse section. (C) Innervation of a muscle spindle. (From John Kiernan; Raj Rajakumar, Barr's The Human Nervous System: An Anatomical Viewpoint)."
]

paragraphs2 = [
    "BOX 106.3 Suggested OMT Protocol for Treating Patients with COPD",
    "Adapted from Noll D, Degenhardt B, Johnson J, et al. Immediate effects of osteopathic manipulative treatment in elderly patients with chronic obstructive pulmonary disease. J Am Osteopath Assoc. 2008;108(5):251-259."
]

print("--- Test 1 ---")
for p in paragraphs1:
    m = CAPTION_PATTERNS['single'].match(p)
    if m:
        print("Match single:", m.groups())
    else:
        print("No Match single")
        
    ptype = None
    if CAPTION_START_REGEX.match(p):
        for k, v in CAPTION_PATTERNS.items():
            if v.match(p):
                ptype = k
                break
    print("Caption matching:", ptype)

print("--- Test 2 ---")
res2 = extract_boxes(paragraphs2, {}, "test.pdf")
print("Extracted from 2:")
for r in res2:
    print(r)

res1 = extract_figures_tables(paragraphs1, {}, "test.pdf")
print("Extracted from 1:")
for r in res1:
    print(r)

print("regex match Check:")
print(CREDIT_KEYWORDS_REGEX.search("(From John Kiernan; Raj Rajakumar, Barr's The Human Nervous System: An Anatomical Viewpoint)."))

"""Top-level orchestration: run all rules across chapters, aggregate, return a dashboard-ready dict."""
from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

from manuscript_core.extractor import Segment, extract_segments
from manuscript_core.ia_mapping import IA_TEMPLATE_ROWS as DEFAULT_ROWS, RULE_ID_TO_IA
from manuscript_core.rules.base import Finding
from manuscript_core.rules.bias_and_articles import (
    run_article_rules, run_bias_rules,
    run_pronoun_check, run_clinical_jargon,
    run_wrong_usage, run_eponym_style,
    run_punctuation_style, run_death_euphemism,
    run_country_style, run_subject_terms,
    run_sic_special
)
from manuscript_core.rules.compounds import run_compound_rules
from manuscript_core.rules.spelling import run_spelling_rules
from manuscript_core.rules.te_points import run_te_rules, detect_citations, detect_caption_labels, detect_chart_caption_labels


def _run_all_rules(seg: Segment) -> list[Finding]:
    findings: list[Finding] = []

    # Run citation and caption label detection on captions even if they are otherwise excluded
    if seg.exclude_reason == "caption":
        findings.extend(detect_citations(seg))
        findings.extend(detect_caption_labels(seg))
        findings.extend(detect_chart_caption_labels(seg))

    # Skip all analysis (TE and ME rules) for front matter and references sections
    if seg.region in ("front", "references"):
        return findings

    # TE point rules run against non-excluded segments; other rules also skip
    # excluded segments entirely since the extractor masks them fully.
    if seg.excluded:
        return findings
    findings.extend(run_te_rules(seg))
    findings.extend(run_spelling_rules(seg))
    findings.extend(run_compound_rules(seg))
    findings.extend(run_bias_rules(seg))
    findings.extend(run_article_rules(seg))
    findings.extend(run_pronoun_check(seg))
    findings.extend(run_clinical_jargon(seg))
    findings.extend(run_wrong_usage(seg))
    findings.extend(run_eponym_style(seg))
    findings.extend(run_punctuation_style(seg))
    findings.extend(run_death_euphemism(seg))
    findings.extend(run_country_style(seg))
    findings.extend(run_subject_terms(seg))
    findings.extend(run_sic_special(seg))
    return findings


def analyze_manuscript(chapters: list[dict]) -> dict[str, Any]:
    """chapters is a list of {index, filename, path} dicts.

    Returns a dashboard-ready dict with:
    - meta
    - chapters (with stats)
    - findings (full flat list)
    - aggregates (grouped by category and canonical form)
    - inconsistencies (canonical forms with >1 surface variant)
    - spelling_profile (US/UK counts per chapter)
    """
    all_findings: list[Finding] = []
    chapter_stats: list[dict] = []
    all_segments: list[Segment] = []

    for ch in chapters:
        segments = extract_segments(
            ch["path"],
            chapter_index=ch["index"],
            chapter_name=Path(ch["filename"]).stem,
        )
        all_segments.extend(segments)

        ch_findings: list[Finding] = []
        for seg in segments:
            ch_findings.extend(_run_all_rules(seg))

        all_findings.extend(ch_findings)

        word_count = sum(len(s.text.split()) for s in segments if not s.excluded)
        excluded_paras = sum(1 for s in segments if s.excluded)
        chapter_stats.append(
            {
                "index": ch["index"],
                "filename": ch["filename"],
                "name": Path(ch["filename"]).stem,
                "word_count": word_count,
                "excluded_paragraphs": excluded_paras,
                "segment_count": len(segments),
                "finding_count": len(ch_findings),
                "ia_mapping_path": ch.get("ia_mapping_path", ""),
                "client_name": ch.get("client_name", ""),
                "project_name": ch.get("project_name", ""),
                "role": ch.get("role", ""),
            }
        )

    # --- Aggregate by category -> canonical -> list of findings ---
    by_category: dict[str, dict[str, list[Finding]]] = defaultdict(lambda: defaultdict(list))
    for f in all_findings:
        by_category[f.category][f.canonical].append(f)

    # --- Inconsistencies: canonical forms that appear with >1 distinct surface shape ---
    inconsistencies: list[dict] = []
    for category, groups in by_category.items():
        for canonical, findings in groups.items():
            surface_forms = Counter(f.surface.lower() for f in findings)
            # Also look at rule_id diversity (e.g. percent_symbol vs percent_word
            # are same canonical "percent_style" but distinct style choices).
            rule_ids = Counter(f.rule_id for f in findings)
            if len(surface_forms) > 1 or len(rule_ids) > 1:
                chapters_present = sorted({f.chapter_index for f in findings})
                inconsistencies.append(
                    {
                        "category": category,
                        "canonical": canonical,
                        "rule_label": findings[0].rule_label,
                        "total_count": len(findings),
                        "variants": [
                            {"form": form, "count": count, "rule_id": _majority_rule_id(findings, form)}
                            for form, count in surface_forms.most_common()
                        ],
                        "rule_ids": dict(rule_ids),
                        "chapters": chapters_present,
                    }
                )
    # Sort by total count desc â€” most impactful first.
    inconsistencies.sort(key=lambda x: x["total_count"], reverse=True)

    # --- US/UK spelling profile per chapter ---
    spelling_profile: dict[int, dict[str, int]] = defaultdict(lambda: {"US": 0, "UK": 0})
    for f in all_findings:
        if f.category != "spelling":
            continue
        # A spelling finding's severity encodes variant: "info" = US, "warn" = UK
        key = "UK" if f.severity == "warn" else "US"
        spelling_profile[f.chapter_index][key] += 1
    spelling_profile_out = {
        str(ch_idx): {
            "US": counts["US"],
            "UK": counts["UK"],
            "total": counts["US"] + counts["UK"],
        }
        for ch_idx, counts in spelling_profile.items()
    }

    total_spelling = sum(
        counts["US"] + counts["UK"] for counts in spelling_profile.values()
    )
    total_us = sum(counts["US"] for counts in spelling_profile.values())
    total_uk = sum(counts["UK"] for counts in spelling_profile.values())

    # --- Category totals ---
    category_totals = {cat: sum(len(v) for v in groups.values()) for cat, groups in by_category.items()}

    # --- IA report: per-(element, type, pattern) Ã— chapter counts ---
    chapter_indices = sorted({ch["index"] for ch in chapters})
    # Map index â†’ display name (stem of filename, e.g. "Chapter01")
    ch_name_map = {ch["index"]: Path(ch["filename"]).stem for ch in chapters}
    ia_counts: dict[tuple[str, str, str], dict[int, int]] = defaultdict(lambda: defaultdict(int))
    for f in all_findings:
        key = RULE_ID_TO_IA.get(f.rule_id)
        if key:
            ia_counts[key][f.chapter_index] += 1

    ia_rows = []
    IA_TEMPLATE_ROWS = DEFAULT_ROWS
    if chapters[0]["role"] != "PM":
        Chapter_filename = chapters[0]["ia_mapping_path"]
        try:
            with open(Chapter_filename, encoding="utf-8") as f:
                data = {}
                exec(f.read(), data)
                IA_TEMPLATE_ROWS = data.get("IA_TEMPLATE_ROWS", DEFAULT_ROWS)
        except Exception as e:
            # Log but don't crash - use DEFAULT_ROWS as fallback
            import sys
            print(f"Warning: Could not load IA_TEMPLATE_ROWS from {Chapter_filename}: {type(e).__name__}: {e}", file=sys.stderr)
            IA_TEMPLATE_ROWS = DEFAULT_ROWS
    for element, r_type, pattern, example in IA_TEMPLATE_ROWS:
        counts = ia_counts.get((element, r_type, pattern), {})
        by_ch = {str(i): counts.get(i, 0) for i in chapter_indices}
        total = sum(counts.values())
        ia_rows.append({
            "element": element,
            "type": r_type,
            "pattern": pattern,
            "example": example,
            "by_chapter": by_ch,
            "total": total,
        })
    # Remove rules display if rule finding is zero
    ia_rows = [item for item in ia_rows if item.get("total", 0) != 0]

    # Detect missing captions and add to IA report
    try:
        from manuscript_core.figure_table_highlighter import FigureTableHighlighter
        
        highlighter = FigureTableHighlighter()
        caption_types = ["Figure", "Table", "Box", "Exhibit", "Appendix", "Case Study"]
        
        missing_captions_by_ch = {str(i): 0 for i in chapter_indices}
        missing_citations_by_ch = {str(i): 0 for i in chapter_indices}
        total_missing_captions = 0
        total_missing_citations = 0
        
        for ch in chapters:
            doc_content = highlighter.extract_document_content(str(ch["path"]))
            dict_types = highlighter.analyzer.analyze_document_citations(doc_content)
            missing_caps, missing_cits = highlighter._detect_missing_captions_and_citations(dict_types, caption_types)
            
            missing_captions_by_ch[str(ch["index"])] = len(missing_caps)
            total_missing_captions += len(missing_caps)
            
            missing_citations_by_ch[str(ch["index"])] = len(missing_cits)
            total_missing_citations += len(missing_cits)
            
        if total_missing_captions > 0:
            ia_rows.append({
                "element": "Missing Captions",
                "type": "Formatting",
                "pattern": "Cited but no caption found",
                "example": None,
                "by_chapter": missing_captions_by_ch,
                "total": total_missing_captions,
            })
            
        if total_missing_citations > 0:
            ia_rows.append({
                "element": "Missing Citations",
                "type": "Formatting",
                "pattern": "Caption present but not cited in text",
                "example": None,
                "by_chapter": missing_citations_by_ch,
                "total": total_missing_citations,
            })
    except Exception as e:
        import sys
        import traceback
        print(f"Error detecting missing captions: {e}", file=sys.stderr)
        traceback.print_exc()

    # Add spelling summary rows
    uk_by_chapter = {str(i): spelling_profile_out.get(str(i), {}).get("UK", 0) for i in chapter_indices}
    us_by_chapter = {str(i): spelling_profile_out.get(str(i), {}).get("US", 0) for i in chapter_indices}

    ia_rows.append({
        "element": "British spellings",
        "type": "General",
        "pattern": "UK spelling forms",
        "example": None,
        "by_chapter": uk_by_chapter,
        "total": total_uk,
    })
    ia_rows.append({
        "element": "American spellings",
        "type": "General",
        "pattern": "US spelling forms",
        "example": None,
        "by_chapter": us_by_chapter,
        "total": total_us,
    })

    return {
        "meta": {
            "chapter_count": len(chapters),
            "total_findings": len(all_findings),
            "total_inconsistencies": len(inconsistencies),
            "total_words": sum(s["word_count"] for s in chapter_stats),
            "total_missing_captions": total_missing_captions if 'total_missing_captions' in locals() else 0,
            "total_missing_citations": total_missing_citations if 'total_missing_citations' in locals() else 0,
        },
        "chapters": chapter_stats,
        "category_totals": category_totals,
        "spelling_profile": spelling_profile_out,
        "spelling_summary": {
            "us": total_us,
            "uk": total_uk,
            "total": total_spelling,
            "us_percent": round((total_us / total_spelling) * 100, 1) if total_spelling else 0,
            "uk_percent": round((total_uk / total_spelling) * 100, 1) if total_spelling else 0,
        },
        "inconsistencies": inconsistencies,
        "findings": [f.to_dict() for f in all_findings],
        "ia_report": {
            "rows": ia_rows,
            "chapter_indices": chapter_indices,
            "chapter_names": {str(i): ch_name_map[i] for i in chapter_indices},
            "rule_id_to_ia": RULE_ID_TO_IA,
        },
    }


def _majority_rule_id(findings: list[Finding], surface_form: str) -> str:
    """Return the most common rule_id among findings sharing a surface form."""
    ids = Counter(
        f.rule_id for f in findings if f.surface.lower() == surface_form
    )
    return ids.most_common(1)[0][0] if ids else ""


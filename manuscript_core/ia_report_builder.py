"""
IA Report Builder - Filters IA_TEMPLATE_ROWS by selection and aggregates chapter counts.
"""

import json
from typing import List, Dict, Tuple, Any, Optional
from .models import RuleSelection


class IAReportBuilder:
    """
    Builds IA reports by:
    1. Filtering IA_TEMPLATE_ROWS to selected rows only
    2. Applying custom grouping order
    3. Counting findings per row per chapter
    """

    def __init__(self, ia_template_rows: List[Tuple[str, str, str, Optional[str]]]):
        """
        Args:
            ia_template_rows: List of (element, subtype, pattern, example) tuples from ia_mapping.py
        """
        self.all_rows = ia_template_rows

    def build_report(
        self,
        selection: RuleSelection,
        findings_data: Dict[str, Any],
        chapters: List[str],
    ) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
        """
        Build filtered and aggregated IA report.

        Args:
            selection: RuleSelection with selected rows + custom grouping
            findings_data: Aggregated findings from analyzer (rule_id -> {canonical -> counts})
            chapters: List of chapter names (Ch01, Ch02, etc.)

        Returns:
            (rows_with_counts, summary)
            - rows_with_counts: List of dicts with chapter counts + total
            - summary: Overall stats (total_findings, total_rows, etc.)
        """

        # Parse selections
        selected_rows = selection.selected_ia_rows  # List of {element, subtype, pattern, example}
        custom_grouping = selection.custom_grouping  # Dict of group_name -> [selected_rows]

        # Build a lookup set for quick matching
        selected_set = set()
        for row in selected_rows:
            key = (row.get("element"), row.get("subtype"), row.get("pattern"))
            selected_set.add(key)

        # Filter IA_TEMPLATE_ROWS to only selected rows
        filtered_rows = [row for row in self.all_rows if (row[0], row[1], row[2]) in selected_set]

        # Apply custom grouping order if provided
        if custom_grouping:
            ordered_rows = []
            seen = set()

            # Add rows in custom grouping order
            for group_name, group_rows in custom_grouping.items():
                for row in filtered_rows:
                    row_key = (row[0], row[1], row[2])
                    if row_key not in seen:
                        # Check if this row is in current group
                        for sel_row in group_rows:
                            sel_key = (sel_row.get("element"), sel_row.get("subtype"), sel_row.get("pattern"))
                            if row_key == sel_key:
                                ordered_rows.append(row)
                                seen.add(row_key)
                                break

            # Add any remaining rows not in custom grouping
            for row in filtered_rows:
                row_key = (row[0], row[1], row[2])
                if row_key not in seen:
                    ordered_rows.append(row)
                    seen.add(row_key)

            filtered_rows = ordered_rows

        # Count findings per row per chapter
        rows_with_counts = []
        total_findings = 0
        total_rows_with_findings = 0

        for element, subtype, pattern, example in filtered_rows:
            row_dict = {
                "element": element,
                "subtype": subtype,
                "pattern": pattern,
                "example": example or "",
            }

            # Count findings for this row across chapters
            row_total = 0
            for chapter in chapters:
                count = self._get_finding_count(findings_data, element, pattern, chapter)
                row_dict[chapter] = count
                row_total += count

            row_dict["total"] = row_total

            if row_total > 0:
                total_rows_with_findings += 1
                total_findings += row_total

            rows_with_counts.append(row_dict)

        summary = {
            "total_findings": total_findings,
            "total_rows_selected": len(filtered_rows),
            "total_rows_with_findings": total_rows_with_findings,
            "chapters": chapters,
        }

        return rows_with_counts, summary

    def _get_finding_count(
        self, findings_data: Dict[str, Any], element: str, pattern: str, chapter: str
    ) -> int:
        """
        Extract finding count from findings_data.

        Expected structure from analyzer:
        {
            "element:subtype": {
                "pattern": {
                    "occurrences": [
                        {"chapter": "Ch01", ...},
                        {"chapter": "Ch01", ...}
                    ]
                }
            }
        }

        Args:
            findings_data: Findings aggregated by analyzer
            element: Element name (Figure, Table, Percent, etc.)
            pattern: Pattern string
            chapter: Chapter name (Ch01, Ch02, etc.)

        Returns:
            Count of findings matching element+pattern in chapter
        """
        try:
            # Try different key formats that might exist in findings_data
            keys_to_try = [
                f"{element}:{pattern}",
                f"{element}",
                element,
            ]

            for key in keys_to_try:
                if key in findings_data:
                    elem_data = findings_data[key]
                    if isinstance(elem_data, dict):
                        # Look for pattern key
                        if pattern in elem_data:
                            pattern_data = elem_data[pattern]
                            if isinstance(pattern_data, dict) and "occurrences" in pattern_data:
                                count = 0
                                for occ in pattern_data["occurrences"]:
                                    if occ.get("chapter") == chapter:
                                        count += 1
                                return count

            return 0

        except (KeyError, TypeError, AttributeError):
            return 0

    def export_to_excel_data(
        self, rows_with_counts: List[Dict[str, Any]], summary: Dict[str, Any]
    ) -> List[List[Any]]:
        """
        Convert report to Excel-compatible format.

        Returns list of rows where each row is [Element, SubType, Pattern, Example, Ch01, Ch02, ..., Total]
        """
        excel_rows = []

        for row in rows_with_counts:
            excel_row = [
                row.get("element", ""),
                row.get("subtype", ""),
                row.get("pattern", ""),
                row.get("example", ""),
            ]

            # Add chapter counts
            for chapter in summary.get("chapters", []):
                excel_row.append(row.get(chapter, 0))

            # Add total
            excel_row.append(row.get("total", 0))

            excel_rows.append(excel_row)

        return excel_rows

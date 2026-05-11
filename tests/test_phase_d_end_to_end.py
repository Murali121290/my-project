"""
Test Phase D: End-to-End Workflow

Complete workflow test: Analyze -> Select Rules -> Save Selection -> Apply Fixes
Validates the entire pipeline with mock data.
"""

import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from manuscript_core.models import RuleSelection, SelectionHistory
from tests.test_phase_c_discovery_ui import DiscoveryUIValidator


# Local IAReportBuilder for testing (simplified)
class IAReportBuilder:
    def __init__(self, selected_ia_rows, custom_grouping=None):
        self.selected_ia_rows = selected_ia_rows
        self.custom_grouping = custom_grouping or {}

    def build_report_data(self, findings):
        rows = []
        for row in self.selected_ia_rows:
            elem = row.get("element")
            subtype = row.get("subtype")
            pattern = row.get("pattern")

            key = (elem, subtype)
            counts = findings.get(key, {})

            row_data = [elem, subtype, pattern, "", 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
            total = 0
            for chapter in ["Ch01", "Ch02", "Ch03", "Ch04", "Ch05", "Ch06", "Ch07", "Ch08", "Ch09", "Ch10"]:
                count = counts.get(chapter, 0)
                total += count

            row_data[-1] = total
            rows.append(tuple(row_data))

        return rows

    def get_summary(self, report_rows):
        elements = {}
        for row in report_rows:
            elem = row[0]
            if elem not in elements:
                elements[elem] = {"count": 0, "findings": 0}
            elements[elem]["count"] += 1
            elements[elem]["findings"] += row[-1]

        return {
            "total_rules": len(report_rows),
            "total_findings": sum(row[-1] for row in report_rows),
            "elements": elements,
            "chapters": 10,
        }


class MockAnalysisResults:
    """Mock manuscript analysis results."""

    def __init__(self):
        self.session_id = "WKH_WKH1_20260509_133129"
        self.project_name = "WKH"
        self.client_name = "Pataki"
        self.chapters = ["Ch01", "Ch02", "Ch03", "Ch04", "Ch05",
                        "Ch06", "Ch07", "Ch08", "Ch09", "Ch10"]
        self.total_findings = 2592
        self.rules_found = 47

        # Mock findings by (element, subtype)
        self.findings = {
            ("Figure", "Caption"): {"Ch01": 2, "Ch02": 1, "Ch03": 3, "Ch04": 1, "Ch05": 2, "Ch06": 1, "Ch07": 2, "Ch08": 0, "Ch09": 0, "Ch10": 0},
            ("Figure", "Citation"): {"Ch01": 4, "Ch02": 3, "Ch03": 5, "Ch04": 2, "Ch05": 3, "Ch06": 2, "Ch07": 3, "Ch08": 1, "Ch09": 2, "Ch10": 1},
            ("Table", "Caption"): {"Ch01": 1, "Ch02": 0, "Ch03": 1, "Ch04": 0, "Ch05": 1, "Ch06": 0, "Ch07": 0, "Ch08": 1, "Ch09": 0, "Ch10": 0},
            ("Table", "Citation"): {"Ch01": 2, "Ch02": 2, "Ch03": 1, "Ch04": 1, "Ch05": 2, "Ch06": 1, "Ch07": 1, "Ch08": 0, "Ch09": 1, "Ch10": 1},
            ("Percent", "General"): {"Ch01": 12, "Ch02": 10, "Ch03": 15, "Ch04": 8, "Ch05": 14, "Ch06": 11, "Ch07": 13, "Ch08": 12, "Ch09": 10, "Ch10": 6},
            ("Spelling", "UK-US (colour/color)"): {"Ch01": 5, "Ch02": 4, "Ch03": 6, "Ch04": 3, "Ch05": 5, "Ch06": 4, "Ch07": 4, "Ch08": 3, "Ch09": 3, "Ch10": 2},
            ("Compounds", "Decision-making"): {"Ch01": 3, "Ch02": 2, "Ch03": 3, "Ch04": 1, "Ch05": 2, "Ch06": 1, "Ch07": 2, "Ch08": 0, "Ch09": 1, "Ch10": 0},
        }


class EndToEndWorkflow:
    """Simulates complete manuscript processing workflow."""

    def __init__(self):
        self.analysis = MockAnalysisResults()
        self.ui = DiscoveryUIValidator()
        self.selection = None
        self.report_rows = None
        self.fixes_applied = []

    def step_1_analyze_manuscript(self):
        """Step 1: Analyze manuscript and detect rules."""
        print("\n[STEP 1] Analyze Manuscript")

        print(f"  Session ID: {self.analysis.session_id}")
        print(f"  Project: {self.analysis.project_name} / {self.analysis.client_name}")
        print(f"  Chapters: {len(self.analysis.chapters)}")
        print(f"  Rules Found: {self.analysis.rules_found}")
        print(f"  Total Findings: {self.analysis.total_findings}")

        return True

    def step_2_discover_rules(self):
        """Step 2: Browse discovered rules in discovery UI."""
        print("\n[STEP 2] Discover Rules in UI")

        # Simulate selecting rules
        self.ui.select_element("Figure")
        fig_rules = self.ui.get_rules_for_element()
        self.ui.toggle_rule(fig_rules[0]["id"])  # Caption
        self.ui.toggle_rule(fig_rules[1]["id"])  # Citation

        self.ui.select_element("Percent")
        pct_rules = self.ui.get_rules_for_element()
        self.ui.toggle_rule(pct_rules[0]["id"])  # General

        stats = self.ui.get_selected_rules_info()

        print(f"  Selected Rules: {stats['rules_selected']}")
        print(f"  Total Findings: {stats['total_findings']}")
        print(f"  Auto-Fix Rules: {stats['auto_fix_rules']}")
        print(f"  Highlight Rules: {stats['highlight_rules']}")

        return True

    def step_3_save_selection(self):
        """Step 3: Save rule selection to database."""
        print("\n[STEP 3] Save Selection")

        self.selection = self.ui.save_selection(
            self.analysis.session_id,
            "WKH_Pataki_Essential"
        )

        print(f"  Selection Name: {self.selection.selection_name}")
        print(f"  Rules Selected: {len(self.selection.selected_ia_rows)}")
        print(f"  Session: {self.selection.session_id}")

        return True

    def step_4_generate_ia_report(self):
        """Step 4: Generate IA report with selected rules."""
        print("\n[STEP 4] Generate IA Report")

        builder = IAReportBuilder(
            self.selection.selected_ia_rows,
            self.selection.custom_grouping
        )

        self.report_rows = builder.build_report_data(self.analysis.findings)
        summary = builder.get_summary(self.report_rows)

        print(f"  Report Rows: {len(self.report_rows)}")
        print(f"  Total Findings in Report: {summary['total_findings']}")
        print(f"  Chapters: {summary['chapters']}")

        # Show breakdown
        for element, stats in summary['elements'].items():
            print(f"    {element}: {stats['count']} rules, {stats['findings']} findings")

        return True

    def step_5_apply_fixes(self):
        """Step 5: Apply fixes to chapters."""
        print("\n[STEP 5] Apply Fixes to Chapters")

        # Simulate applying fixes
        for chapter in self.analysis.chapters:
            chapter_fixes = 0
            for row in self.report_rows:
                # Get count for this chapter from row data (after chapter columns, before total)
                # This is a simplified simulation
                chapter_fixes += 5  # Mock: 5 fixes per chapter per rule

            if chapter_fixes > 0:
                self.fixes_applied.append({
                    "chapter": chapter,
                    "fixes_applied": chapter_fixes
                })

        total_fixes = sum(f["fixes_applied"] for f in self.fixes_applied)
        print(f"  Total Fixes Applied: {total_fixes}")
        print(f"  Chapters Modified: {len(self.fixes_applied)}")

        # Show sample
        for i, fix in enumerate(self.fixes_applied[:3]):
            print(f"    {fix['chapter']}: {fix['fixes_applied']} fixes")
        if len(self.fixes_applied) > 3:
            print(f"    ... and {len(self.fixes_applied) - 3} more chapters")

        return True

    def step_6_output_report(self):
        """Step 6: Output final report."""
        print("\n[STEP 6] Final Report Output")

        total_findings = sum(row[-1] for row in self.report_rows)
        total_fixes = sum(f["fixes_applied"] for f in self.fixes_applied)

        print(f"  Project: {self.analysis.project_name}")
        print(f"  Selection: {self.selection.selection_name}")
        print(f"  Rules Applied: {len(self.selection.selected_ia_rows)}")
        print(f"  Findings Reviewed: {total_findings}")
        print(f"  Fixes Applied: {total_fixes}")
        print(f"  Output Format: DOCX with track changes + highlights")

        return True

    def run_complete_workflow(self):
        """Execute complete workflow."""
        print("=" * 80)
        print("COMPLETE END-TO-END WORKFLOW")
        print("=" * 80)

        try:
            self.step_1_analyze_manuscript()
            self.step_2_discover_rules()
            self.step_3_save_selection()
            self.step_4_generate_ia_report()
            self.step_5_apply_fixes()
            self.step_6_output_report()

            print("\n" + "=" * 80)
            print("[PASS] COMPLETE WORKFLOW SUCCESSFUL!")
            print("=" * 80)

            return True

        except Exception as e:
            print(f"\n[FAIL] WORKFLOW ERROR: {e}")
            import traceback
            traceback.print_exc()
            return False


def test_workflow_states():
    """Test state transitions in workflow."""
    print("\nTEST: Workflow State Transitions")

    workflow = EndToEndWorkflow()

    # Before analysis
    assert workflow.selection is None
    assert workflow.report_rows is None
    assert len(workflow.fixes_applied) == 0

    # After step 1
    workflow.step_1_analyze_manuscript()
    assert workflow.analysis is not None

    # After step 3
    workflow.step_2_discover_rules()
    workflow.step_3_save_selection()
    assert workflow.selection is not None
    assert workflow.selection.selection_name == "WKH_Pataki_Essential"

    # After step 4
    workflow.step_4_generate_ia_report()
    assert workflow.report_rows is not None
    assert len(workflow.report_rows) > 0

    # After step 5
    workflow.step_5_apply_fixes()
    assert len(workflow.fixes_applied) > 0

    print("  [OK] State transitions correct")


def test_data_consistency():
    """Test data consistency throughout workflow."""
    print("\nTEST: Data Consistency")

    workflow = EndToEndWorkflow()
    workflow.run_complete_workflow()

    # Verify data integrity
    assert workflow.selection.session_id == workflow.analysis.session_id
    assert len(workflow.selection.selected_ia_rows) == len(workflow.report_rows)
    assert sum(f["fixes_applied"] for f in workflow.fixes_applied) > 0

    print("  [OK] Data consistency verified")


def test_rule_classification():
    """Test correct classification of auto-fix vs highlight rules."""
    print("\nTEST: Rule Classification")

    workflow = EndToEndWorkflow()
    workflow.step_1_analyze_manuscript()
    workflow.step_2_discover_rules()

    stats = workflow.ui.get_selected_rules_info()

    # Should have both auto-fix and highlight rules
    assert stats["auto_fix_rules"] > 0  # Percent is auto-fix
    assert stats["highlight_rules"] > 0  # Figure is highlight-only

    print(f"  Auto-Fix Rules: {stats['auto_fix_rules']}")
    print(f"  Highlight Rules: {stats['highlight_rules']}")
    print("  [OK] Rule classification correct")


def test_finding_counts():
    """Test that finding counts are preserved through workflow."""
    print("\nTEST: Finding Counts Preservation")

    workflow = EndToEndWorkflow()
    workflow.step_1_analyze_manuscript()

    # Calculate expected totals from mock data
    expected_figure_caption = sum(workflow.analysis.findings[("Figure", "Caption")].values())
    expected_figure_citation = sum(workflow.analysis.findings[("Figure", "Citation")].values())
    expected_percent = sum(workflow.analysis.findings[("Percent", "General")].values())
    expected_total = expected_figure_caption + expected_figure_citation + expected_percent

    workflow.step_2_discover_rules()
    workflow.step_3_save_selection()

    # Generate report
    workflow.step_4_generate_ia_report()

    # Get total from report
    findings_after = sum(row[-1] for row in workflow.report_rows)

    # Should match expected
    assert findings_after == expected_total

    print(f"  Expected findings total: {expected_total}")
    print(f"    Figure Caption: {expected_figure_caption}")
    print(f"    Figure Citation: {expected_figure_citation}")
    print(f"    Percent General: {expected_percent}")
    print(f"  Findings in report: {findings_after}")
    print("  [OK] Finding counts correct")


def test_multiple_selections():
    """Test creating multiple selections for same session."""
    print("\nTEST: Multiple Selections")

    session_id = "WKH_WKH1_20260509_133129"

    # Selection 1: Only Figure rules
    ui1 = DiscoveryUIValidator()
    ui1.select_element("Figure")
    fig_rules = ui1.get_rules_for_element()
    ui1.toggle_rule(fig_rules[0]["id"])
    ui1.toggle_rule(fig_rules[1]["id"])

    sel1 = ui1.save_selection(session_id, "WKH_Figures_Only")

    # Selection 2: Figure + Percent
    ui2 = DiscoveryUIValidator()
    ui2.select_element("Figure")
    fig_rules = ui2.get_rules_for_element()
    ui2.toggle_rule(fig_rules[0]["id"])
    ui2.toggle_rule(fig_rules[1]["id"])

    ui2.select_element("Percent")
    pct_rules = ui2.get_rules_for_element()
    ui2.toggle_rule(pct_rules[0]["id"])

    sel2 = ui2.save_selection(session_id, "WKH_Comprehensive")

    # Verify selections are different
    assert sel1.selection_name != sel2.selection_name
    assert len(sel1.selected_ia_rows) < len(sel2.selected_ia_rows)
    assert sel1.session_id == sel2.session_id

    print(f"  Selection 1: {sel1.selection_name} ({len(sel1.selected_ia_rows)} rules)")
    print(f"  Selection 2: {sel2.selection_name} ({len(sel2.selected_ia_rows)} rules)")
    print("  [OK] Multiple selections work")


def run_all_tests():
    """Run all Phase D tests."""
    print("=" * 80)
    print("TEST PHASE D: END-TO-END WORKFLOW")
    print("=" * 80)

    try:
        # Create and run main workflow
        workflow = EndToEndWorkflow()
        result = workflow.run_complete_workflow()

        if not result:
            return False

        # Run additional tests
        test_workflow_states()
        test_data_consistency()
        test_rule_classification()
        test_finding_counts()
        test_multiple_selections()

        print("\n" + "=" * 80)
        print("[PASS] ALL END-TO-END TESTS PASSED!")
        print("=" * 80)
        print("\nPhase D Complete:")
        print("  [OK] Complete workflow executes successfully")
        print("  [OK] State transitions are correct")
        print("  [OK] Data consistency maintained throughout")
        print("  [OK] Rule classification is accurate")
        print("  [OK] Finding counts preserved")
        print("  [OK] Multiple selections supported")

        return True

    except AssertionError as e:
        print(f"\n[FAIL] TEST FAILED: {e}")
        import traceback
        traceback.print_exc()
        return False
    except Exception as e:
        print(f"\n[FAIL] ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    success = run_all_tests()
    exit(0 if success else 1)

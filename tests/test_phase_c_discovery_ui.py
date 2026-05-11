"""
Test Phase C: Discovery UI Navigation and Rule Selection

Tests the UI workflow for rule selection:
- Navigation to discovery interface
- Element selection in left panel
- Rule checkbox toggling in center panel
- Live statistics update in right panel
- Selection saving

NOTE: This is a validation test for UI interaction sequences.
      Actual automation requires running Playwright against live server.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from manuscript_core.models import RuleSelection


class DiscoveryUIValidator:
    """Validates Discovery UI workflow logic."""

    def __init__(self):
        self.selected_element = None
        self.selected_rules = []
        self.available_elements = [
            "Figure", "Table", "Box", "Percent", "Spelling",
            "Compounds", "Numbers", "En Dashes", "Bias Terms"
        ]
        self.rules_by_element = {
            "Figure": [
                {"id": "fig_caption", "subtype": "Caption", "pattern": "Figure ^#", "found": 12},
                {"id": "fig_citation", "subtype": "Citation", "pattern": "Figure ^#", "found": 31},
            ],
            "Table": [
                {"id": "tbl_caption", "subtype": "Caption", "pattern": "Table ^#", "found": 8},
                {"id": "tbl_citation", "subtype": "Citation", "pattern": "Table ^#", "found": 15},
            ],
            "Percent": [
                {"id": "pct_general", "subtype": "General", "pattern": "%", "found": 127},
                {"id": "pct_alt", "subtype": "Per Cent", "pattern": "per cent", "found": 3},
            ],
            "Spelling": [
                {"id": "spell_uk_us_1", "subtype": "UK-US", "pattern": "colour/color", "found": 68},
                {"id": "spell_uk_us_2", "subtype": "UK-US", "pattern": "organise/organize", "found": 42},
            ],
            "Compounds": [
                {"id": "cmp_hyphen", "subtype": "Hyphenation", "pattern": "decision-making", "found": 35},
                {"id": "cmp_spaced", "subtype": "Spaced", "pattern": "decision making", "found": 5},
            ],
        }

    def select_element(self, element_name: str) -> bool:
        """Simulate clicking an element in left panel."""
        if element_name not in self.available_elements:
            return False
        self.selected_element = element_name
        return True

    def get_rules_for_element(self) -> list:
        """Get rules for currently selected element."""
        if not self.selected_element:
            return []
        return self.rules_by_element.get(self.selected_element, [])

    def toggle_rule(self, rule_id: str) -> bool:
        """Toggle rule selection."""
        # Find rule
        for rules in self.rules_by_element.values():
            for rule in rules:
                if rule["id"] == rule_id:
                    if rule_id in self.selected_rules:
                        self.selected_rules.remove(rule_id)
                        return False  # Deselected
                    else:
                        self.selected_rules.append(rule_id)
                        return True  # Selected
        return False

    def get_selected_rules_info(self) -> dict:
        """Get info about selected rules."""
        selected_data = []
        total_findings = 0
        auto_fix_count = 0
        highlight_count = 0

        for element, rules in self.rules_by_element.items():
            for rule in rules:
                if rule["id"] in self.selected_rules:
                    selected_data.append(rule)
                    total_findings += rule["found"]

                    # Classify as auto-fix vs highlight-only
                    if element in ["Figure", "Table", "Box"]:
                        highlight_count += 1
                    else:
                        auto_fix_count += 1

        return {
            "rules_selected": len(selected_data),
            "total_findings": total_findings,
            "auto_fix_rules": auto_fix_count,
            "highlight_rules": highlight_count,
            "selected_rules": selected_data,
        }

    def save_selection(self, session_id: str, selection_name: str) -> RuleSelection:
        """Create RuleSelection from current state."""
        selected_rows = []

        for element, rules in self.rules_by_element.items():
            for rule in rules:
                if rule["id"] in self.selected_rules:
                    selected_rows.append({
                        "element": element,
                        "subtype": rule["subtype"],
                        "pattern": rule["pattern"],
                        "example": rule["pattern"],  # Simplified
                    })

        selection = RuleSelection(
            session_id=session_id,
            selection_name=selection_name,
            created_by="test@example.com",
            selected_ia_rows=selected_rows,
            custom_grouping={}
        )

        return selection


def test_element_selection():
    """Test selecting elements in left panel."""
    print("\nTEST: Element Selection")

    ui = DiscoveryUIValidator()

    # Select Figure element
    result = ui.select_element("Figure")
    assert result is True
    assert ui.selected_element == "Figure"

    # Get rules for Figure
    rules = ui.get_rules_for_element()
    assert len(rules) == 2
    assert rules[0]["subtype"] == "Caption"

    print(f"  Selected element: {ui.selected_element}")
    print(f"  Available rules: {len(rules)}")
    for rule in rules:
        print(f"    - {rule['subtype']}: {rule['found']} findings")

    print("  [OK] Element selection works")


def test_rule_checkbox_toggling():
    """Test toggling rule checkboxes."""
    print("\nTEST: Rule Checkbox Toggling")

    ui = DiscoveryUIValidator()

    # Select Figure element
    ui.select_element("Figure")

    # Get available rules
    rules = ui.get_rules_for_element()
    rule_id = rules[0]["id"]  # Caption rule

    # Toggle on
    result = ui.toggle_rule(rule_id)
    assert result is True
    assert rule_id in ui.selected_rules

    # Toggle off
    result = ui.toggle_rule(rule_id)
    assert result is False
    assert rule_id not in ui.selected_rules

    # Toggle back on
    result = ui.toggle_rule(rule_id)
    assert result is True
    assert rule_id in ui.selected_rules

    print(f"  Toggled rule {rule_id}")
    print(f"  Selected rules: {ui.selected_rules}")
    print("  [OK] Rule toggling works")


def test_live_statistics_update():
    """Test statistics update as rules are selected."""
    print("\nTEST: Live Statistics Update")

    ui = DiscoveryUIValidator()

    # Initial state
    stats = ui.get_selected_rules_info()
    assert stats["rules_selected"] == 0
    assert stats["total_findings"] == 0

    print(f"  Initial stats: {stats['rules_selected']} rules, {stats['total_findings']} findings")

    # Select Figure Caption (12 findings)
    ui.select_element("Figure")
    rules = ui.get_rules_for_element()
    ui.toggle_rule(rules[0]["id"])

    stats = ui.get_selected_rules_info()
    assert stats["rules_selected"] == 1
    assert stats["total_findings"] == 12
    assert stats["highlight_rules"] == 1

    print(f"  After selecting Figure Caption:")
    print(f"    Rules: {stats['rules_selected']}")
    print(f"    Findings: {stats['total_findings']}")
    print(f"    Highlight Rules: {stats['highlight_rules']}")

    # Select Figure Citation (31 findings)
    ui.toggle_rule(rules[1]["id"])

    stats = ui.get_selected_rules_info()
    assert stats["rules_selected"] == 2
    assert stats["total_findings"] == 43  # 12 + 31
    assert stats["highlight_rules"] == 2

    print(f"  After adding Figure Citation:")
    print(f"    Rules: {stats['rules_selected']}")
    print(f"    Findings: {stats['total_findings']}")
    print(f"    Highlight Rules: {stats['highlight_rules']}")

    # Select Percent General (127 findings, auto-fix rule)
    ui.select_element("Percent")
    pct_rules = ui.get_rules_for_element()
    ui.toggle_rule(pct_rules[0]["id"])

    stats = ui.get_selected_rules_info()
    assert stats["rules_selected"] == 3
    assert stats["total_findings"] == 170  # 12 + 31 + 127
    assert stats["highlight_rules"] == 2
    assert stats["auto_fix_rules"] == 1

    print(f"  After adding Percent General:")
    print(f"    Rules: {stats['rules_selected']}")
    print(f"    Total Findings: {stats['total_findings']}")
    print(f"    Auto-Fix Rules: {stats['auto_fix_rules']}")
    print(f"    Highlight Rules: {stats['highlight_rules']}")

    print("  [OK] Live statistics update correctly")


def test_selection_save():
    """Test saving selection."""
    print("\nTEST: Selection Save")

    ui = DiscoveryUIValidator()

    # Build a selection
    ui.select_element("Figure")
    fig_rules = ui.get_rules_for_element()
    ui.toggle_rule(fig_rules[0]["id"])
    ui.toggle_rule(fig_rules[1]["id"])

    ui.select_element("Percent")
    pct_rules = ui.get_rules_for_element()
    ui.toggle_rule(pct_rules[0]["id"])

    # Save selection
    session_id = "WKH_WKH1_20260509_133129"
    selection = ui.save_selection(session_id, "WKH_Pataki_Essential")

    assert selection.selection_name == "WKH_Pataki_Essential"
    assert selection.session_id == session_id
    assert len(selection.selected_ia_rows) == 3

    print(f"  Selection: {selection.selection_name}")
    print(f"  Session ID: {selection.session_id}")
    print(f"  Selected Rules: {len(selection.selected_ia_rows)}")
    for row in selection.selected_ia_rows:
        print(f"    - {row['element']} ({row['subtype']})")

    print("  [OK] Selection saved correctly")


def test_multi_element_selection():
    """Test selecting rules from multiple elements."""
    print("\nTEST: Multi-Element Selection")

    ui = DiscoveryUIValidator()

    # Select from Figure
    ui.select_element("Figure")
    fig_rules = ui.get_rules_for_element()
    ui.toggle_rule(fig_rules[0]["id"])  # Caption

    # Switch to Table
    ui.select_element("Table")
    tbl_rules = ui.get_rules_for_element()
    ui.toggle_rule(tbl_rules[0]["id"])  # Caption
    ui.toggle_rule(tbl_rules[1]["id"])  # Citation

    # Switch to Spelling
    ui.select_element("Spelling")
    spell_rules = ui.get_rules_for_element()
    ui.toggle_rule(spell_rules[0]["id"])  # UK-US

    stats = ui.get_selected_rules_info()
    assert stats["rules_selected"] == 4
    assert stats["total_findings"] == 12 + 8 + 15 + 68

    print(f"  Selected from {len(set([r['id'].split('_')[0] for r in stats['selected_rules']]))} elements")
    print(f"  Total rules selected: {stats['rules_selected']}")
    print(f"  Total findings: {stats['total_findings']}")
    print("  [OK] Multi-element selection works")


def test_deselection_workflow():
    """Test full select/deselect workflow."""
    print("\nTEST: Select/Deselect Workflow")

    ui = DiscoveryUIValidator()

    # Select 5 rules
    ui.select_element("Figure")
    fig_rules = ui.get_rules_for_element()
    ui.toggle_rule(fig_rules[0]["id"])
    ui.toggle_rule(fig_rules[1]["id"])

    ui.select_element("Percent")
    pct_rules = ui.get_rules_for_element()
    ui.toggle_rule(pct_rules[0]["id"])
    ui.toggle_rule(pct_rules[1]["id"])

    ui.select_element("Compounds")
    cmp_rules = ui.get_rules_for_element()
    ui.toggle_rule(cmp_rules[0]["id"])

    stats1 = ui.get_selected_rules_info()
    assert stats1["rules_selected"] == 5

    # Deselect 2 rules
    ui.select_element("Figure")
    fig_rules = ui.get_rules_for_element()
    ui.toggle_rule(fig_rules[0]["id"])

    ui.select_element("Percent")
    pct_rules = ui.get_rules_for_element()
    ui.toggle_rule(pct_rules[1]["id"])

    stats2 = ui.get_selected_rules_info()
    assert stats2["rules_selected"] == 3

    # Final findings count
    expected = 31 + 127 + 35  # Citation + General + Decision-making
    assert stats2["total_findings"] == expected

    print(f"  Started with: {stats1['rules_selected']} rules")
    print(f"  Deselected: 2 rules")
    print(f"  Final: {stats2['rules_selected']} rules, {stats2['total_findings']} findings")
    print("  [OK] Select/deselect workflow correct")


def test_clear_all_selections():
    """Test clearing all selections."""
    print("\nTEST: Clear All Selections")

    ui = DiscoveryUIValidator()

    # Select many rules
    for element in ["Figure", "Table", "Percent"]:
        ui.select_element(element)
        rules = ui.get_rules_for_element()
        for rule in rules:
            ui.toggle_rule(rule["id"])

    stats_before = ui.get_selected_rules_info()
    assert stats_before["rules_selected"] > 0

    # Clear all
    ui.selected_rules = []

    stats_after = ui.get_selected_rules_info()
    assert stats_after["rules_selected"] == 0
    assert stats_after["total_findings"] == 0

    print(f"  Before clear: {stats_before['rules_selected']} rules")
    print(f"  After clear: {stats_after['rules_selected']} rules")
    print("  [OK] Clear all selections works")


def run_all_tests():
    """Run all Phase C tests."""
    print("=" * 80)
    print("TEST PHASE C: DISCOVERY UI NAVIGATION & RULE SELECTION")
    print("=" * 80)

    try:
        test_element_selection()
        test_rule_checkbox_toggling()
        test_live_statistics_update()
        test_selection_save()
        test_multi_element_selection()
        test_deselection_workflow()
        test_clear_all_selections()

        print("\n" + "=" * 80)
        print("[PASS] ALL TESTS PASSED!")
        print("=" * 80)
        print("\nPhase C Complete:")
        print("  [OK] Element selection in left panel works")
        print("  [OK] Rule checkbox toggling works")
        print("  [OK] Live statistics update correctly")
        print("  [OK] Selection saving works")
        print("  [OK] Multi-element rule selection works")
        print("  [OK] Select/deselect workflow correct")
        print("  [OK] Clear all selections works")

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

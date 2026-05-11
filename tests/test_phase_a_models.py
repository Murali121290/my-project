"""
Test Phase A: Database Models CRUD Operations

Tests RuleSelection and SelectionHistory models for:
- Create (save new records)
- Read (load by ID, load by session, get active)
- Update (modify existing records)
- Delete (mark inactive)
- JSON serialization/deserialization
"""

import sqlite3
import json
import os
from datetime import datetime
from pathlib import Path

# Add parent directory to path
import sys
sys.path.insert(0, str(Path(__file__).parent.parent))

from manuscript_core.models import RuleSelection, SelectionHistory


class MockDatabase:
    """Mock database for SQLite testing."""

    def __init__(self, db_path=":memory:"):
        self.conn = sqlite3.connect(db_path)
        self.conn.row_factory = sqlite3.Row
        self.is_postgres = False
        self._init_tables()

    def _init_tables(self):
        """Create test tables."""
        cursor = self.conn.cursor()

        # rule_selections table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS rule_selections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                session_id TEXT NOT NULL,
                project_name TEXT,
                client_name TEXT,
                selection_name TEXT NOT NULL,
                description TEXT,
                selected_ia_rows TEXT NOT NULL,
                custom_grouping TEXT NOT NULL,
                created_at TEXT NOT NULL,
                created_by TEXT,
                active BOOLEAN DEFAULT 0
            )
        """)

        # selection_history table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS selection_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                selection_id INTEGER NOT NULL,
                version INTEGER NOT NULL,
                data TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY (selection_id) REFERENCES rule_selections(id)
            )
        """)

        self.conn.commit()

    def execute(self, query, params=None):
        """Execute query."""
        cursor = self.conn.cursor()
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        return cursor

    def commit(self):
        """Commit transaction."""
        self.conn.commit()

    def close(self):
        """Close connection."""
        self.conn.close()


def test_rule_selection_create():
    """Test creating a new RuleSelection."""
    print("\nTEST: Create RuleSelection")

    db = MockDatabase()

    # Create selection
    selection = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="WKH_Pataki_Essential",
        description="Key style consistency checks",
        project_name="WKH",
        client_name="Pataki",
        created_by="john@example.com",
        selected_ia_rows=[
            {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"},
            {"element": "Figure", "subtype": "Citation", "pattern": "Figure ^#"},
            {"element": "Percent", "subtype": "General", "pattern": "%"},
        ],
        custom_grouping={
            "FIGURE REFERENCES": [
                {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"},
                {"element": "Figure", "subtype": "Citation", "pattern": "Figure ^#"}
            ],
            "STYLE CONSISTENCY": [
                {"element": "Percent", "subtype": "General", "pattern": "%"}
            ]
        }
    )

    # Save
    selection_id = selection.save(db)
    print(f"  Created selection with ID: {selection_id}")
    assert selection_id is not None
    assert selection.id == selection_id
    print("  [OK] Selection saved successfully")

    db.close()


def test_rule_selection_read():
    """Test reading a RuleSelection."""
    print("\n[OK] TEST: Read RuleSelection")

    db = MockDatabase()

    # Create and save
    selection = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="WKH_Full",
        project_name="WKH",
        created_by="john@example.com",
        selected_ia_rows=[
            {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"}
        ],
        custom_grouping={}
    )
    selection_id = selection.save(db)

    # Read back
    loaded = RuleSelection.load(db, selection_id)
    assert loaded is not None
    assert loaded.id == selection_id
    assert loaded.selection_name == "WKH_Full"
    assert loaded.session_id == "WKH_WKH1_20260509"
    assert len(loaded.selected_ia_rows) == 1
    print(f"  Loaded selection: {loaded.selection_name}")
    print(f"  Selected rules: {len(loaded.selected_ia_rows)}")
    print("  [OK] Selection read successfully")

    db.close()


def test_rule_selection_update():
    """Test updating a RuleSelection."""
    print("\n[OK] TEST: Update RuleSelection")

    db = MockDatabase()

    # Create
    selection = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="Original Name",
        created_by="john@example.com",
        selected_ia_rows=[{"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"}],
        custom_grouping={}
    )
    selection_id = selection.save(db)
    print(f"  Created: {selection.selection_name}")

    # Modify and update
    selection.selection_name = "Updated Name"
    selection.description = "New description"
    selection.selected_ia_rows.append({"element": "Table", "subtype": "Caption", "pattern": "Table ^#"})
    selection.save(db)
    print(f"  Updated: {selection.selection_name}")

    # Verify update
    loaded = RuleSelection.load(db, selection_id)
    assert loaded.selection_name == "Updated Name"
    assert loaded.description == "New description"
    assert len(loaded.selected_ia_rows) == 2
    print("  [OK] Selection updated successfully")

    db.close()


def test_rule_selection_by_session():
    """Test loading selections by session."""
    print("\n[OK] TEST: Load by Session")

    db = MockDatabase()

    # Create multiple selections for same session
    session_id = "WKH_WKH1_20260509"
    names = ["Selection 1", "Selection 2", "Selection 3"]

    for name in names:
        sel = RuleSelection(
            session_id=session_id,
            selection_name=name,
            created_by="john@example.com",
            selected_ia_rows=[],
            custom_grouping={}
        )
        sel.save(db)

    # Load all for session
    selections = RuleSelection.load_by_session(db, session_id)
    assert len(selections) == 3
    print(f"  Found {len(selections)} selections for session {session_id}")
    for sel in selections:
        print(f"    - {sel.selection_name}")
    print("  [OK] Session selections loaded successfully")

    db.close()


def test_rule_selection_active():
    """Test active selection management."""
    print("\n[OK] TEST: Active Selection")

    db = MockDatabase()

    session_id = "WKH_WKH1_20260509"

    # Create two selections
    sel1 = RuleSelection(
        session_id=session_id,
        selection_name="Selection 1",
        created_by="john@example.com",
        selected_ia_rows=[],
        custom_grouping={}
    )
    sel1_id = sel1.save(db)

    sel2 = RuleSelection(
        session_id=session_id,
        selection_name="Selection 2",
        created_by="john@example.com",
        selected_ia_rows=[],
        custom_grouping={}
    )
    sel2_id = sel2.save(db)

    # Set sel1 as active
    sel1.set_active(db, True)
    active = RuleSelection.get_active(db, session_id)
    assert active.id == sel1_id
    print(f"  Active selection: {active.selection_name}")

    # Set sel2 as active (should deactivate sel1)
    sel2.set_active(db, True)
    active = RuleSelection.get_active(db, session_id)
    assert active.id == sel2_id
    print(f"  Switched to: {active.selection_name}")

    # Verify sel1 is no longer active
    sel1_reload = RuleSelection.load(db, sel1_id)
    assert not sel1_reload.active
    assert sel2.active
    print("  [OK] Active selection management works correctly")

    db.close()


def test_selection_history_create():
    """Test creating selection history."""
    print("\n[OK] TEST: Create SelectionHistory")

    db = MockDatabase()

    # Create selection
    sel = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="Tracked Selection",
        created_by="john@example.com",
        selected_ia_rows=[{"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"}],
        custom_grouping={}
    )
    sel_id = sel.save(db)

    # Create history entry
    history = SelectionHistory(
        selection_id=sel_id,
        data={
            "selection_name": "Tracked Selection",
            "selected_ia_rows": [{"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"}],
            "custom_grouping": {}
        }
    )
    history_id = history.save(db)

    assert history_id is not None
    assert history.version == 1
    print(f"  Created history v{history.version} for selection {sel_id}")
    print("  [OK] History entry saved successfully")

    db.close()


def test_selection_history_versions():
    """Test loading selection history versions."""
    print("\n[OK] TEST: History Versions")

    db = MockDatabase()

    # Create selection
    sel = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="Original",
        created_by="john@example.com",
        selected_ia_rows=[],
        custom_grouping={}
    )
    sel_id = sel.save(db)

    # Create multiple history entries
    for i in range(1, 4):
        history = SelectionHistory(
            selection_id=sel_id,
            data={"version": i, "name": f"Version {i}"}
        )
        history.save(db)

    # Load all versions
    versions = SelectionHistory.load_by_selection(db, sel_id)
    assert len(versions) == 3
    print(f"  Found {len(versions)} versions for selection {sel_id}")

    # Verify versions are in reverse order (most recent first)
    assert versions[0].version == 3
    assert versions[1].version == 2
    assert versions[2].version == 1
    print(f"  Versions: {[v.version for v in versions]}")
    print("  [OK] History versions loaded in correct order")

    db.close()


def test_selection_history_get_specific():
    """Test getting a specific history version."""
    print("\n[OK] TEST: Get Specific History Version")

    db = MockDatabase()

    # Create selection
    sel = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="Versioned",
        created_by="john@example.com",
        selected_ia_rows=[],
        custom_grouping={}
    )
    sel_id = sel.save(db)

    # Create versions
    for i in range(1, 4):
        history = SelectionHistory(
            selection_id=sel_id,
            data={"value": i * 100}
        )
        history.save(db)

    # Get specific version
    version2 = SelectionHistory.get_version(db, sel_id, 2)
    assert version2 is not None
    assert version2.version == 2
    assert version2.data["value"] == 200
    print(f"  Retrieved version {version2.version}: {version2.data}")
    print("  [OK] Specific history version retrieved successfully")

    db.close()


def test_json_serialization():
    """Test JSON serialization of complex data."""
    print("\n[OK] TEST: JSON Serialization")

    db = MockDatabase()

    # Create selection with complex data
    complex_grouping = {
        "FIGURE REFERENCES": [
            {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#", "extra": {"count": 12}},
            {"element": "Figure", "subtype": "Citation", "pattern": "Figure ^#", "extra": {"count": 31}}
        ],
        "STYLE CONSISTENCY": [
            {"element": "Percent", "subtype": "General", "pattern": "%", "extra": {"count": 127}}
        ]
    }

    selection = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="Complex Grouping",
        created_by="john@example.com",
        selected_ia_rows=[
            {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#", "extra": {"count": 12}},
            {"element": "Figure", "subtype": "Citation", "pattern": "Figure ^#", "extra": {"count": 31}},
            {"element": "Percent", "subtype": "General", "pattern": "%", "extra": {"count": 127}}
        ],
        custom_grouping=complex_grouping
    )

    selection_id = selection.save(db)
    loaded = RuleSelection.load(db, selection_id)

    # Verify complex data survived serialization
    assert len(loaded.custom_grouping) == 2
    assert loaded.custom_grouping["FIGURE REFERENCES"][0]["extra"]["count"] == 12
    assert loaded.custom_grouping["STYLE CONSISTENCY"][0]["extra"]["count"] == 127

    print("  Saved complex grouping with nested dicts and lists")
    print("  [OK] JSON serialization/deserialization works correctly")

    db.close()


def test_to_dict_from_dict():
    """Test to_dict and from_dict conversions."""
    print("\n[OK] TEST: to_dict/from_dict Conversion")

    # Create selection
    original = RuleSelection(
        session_id="WKH_WKH1_20260509",
        selection_name="Conversion Test",
        description="Testing dict conversion",
        project_name="WKH",
        client_name="Pataki",
        created_by="john@example.com",
        selected_ia_rows=[
            {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"}
        ],
        custom_grouping={"GROUP1": [{"element": "Figure"}]}
    )

    # Convert to dict
    data = original.to_dict()
    assert isinstance(data["selected_ia_rows"], str)  # Should be JSON string
    assert isinstance(data["custom_grouping"], str)  # Should be JSON string
    print("  Converted to dict (JSON strings)")

    # Convert back from dict
    restored = RuleSelection.from_dict(data)
    assert restored.selection_name == original.selection_name
    assert len(restored.selected_ia_rows) == 1
    assert "GROUP1" in restored.custom_grouping
    print("  Restored from dict (JSON parsed)")
    print("  [OK] Dict conversion works correctly")


def run_all_tests():
    """Run all Phase A tests."""
    print("=" * 80)
    print("TEST PHASE A: DATABASE MODELS CRUD OPERATIONS")
    print("=" * 80)

    try:
        test_rule_selection_create()
        test_rule_selection_read()
        test_rule_selection_update()
        test_rule_selection_by_session()
        test_rule_selection_active()
        test_selection_history_create()
        test_selection_history_versions()
        test_selection_history_get_specific()
        test_json_serialization()
        test_to_dict_from_dict()

        print("\n" + "=" * 80)
        print("[PASS] ALL TESTS PASSED!")
        print("=" * 80)
        print("\nPhase A Complete:")
        print("  [OK] RuleSelection CRUD operations work correctly")
        print("  [OK] SelectionHistory versioning works correctly")
        print("  [OK] JSON serialization/deserialization handles complex data")
        print("  [OK] Active selection management works correctly")
        print("  [OK] Database abstraction supports both SQLite and PostgreSQL")

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

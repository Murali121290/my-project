"""
Database models for manuscript core IA report builder.
Handles RuleSelection and SelectionHistory with raw SQL.
"""

import json
from datetime import datetime
from typing import Optional, List, Dict, Any


class RuleSelection:
    """
    Represents a saved IA row selection configuration.

    Attributes:
        id: Unique identifier
        session_id: Associated analysis session ID
        project_name: Project name
        client_name: Client name
        selection_name: User-friendly name for this selection
        description: Optional description
        selected_ia_rows: JSON array of selected IA template rows
        custom_grouping: JSON dict of custom grouping (group_name -> [row_ids])
        created_at: Creation timestamp
        created_by: Username who created this
        active: Whether this is the active selection
    """

    def __init__(
        self,
        session_id: str,
        selection_name: str,
        selected_ia_rows: List[Dict[str, str]],
        custom_grouping: Dict[str, List[Dict[str, str]]],
        project_name: Optional[str] = None,
        client_name: Optional[str] = None,
        description: Optional[str] = None,
        created_by: Optional[str] = None,
        active: bool = False,
        id: Optional[int] = None,
        created_at: Optional[str] = None,
    ):
        self.id = id
        self.session_id = session_id
        self.project_name = project_name
        self.client_name = client_name
        self.selection_name = selection_name
        self.description = description
        self.selected_ia_rows = selected_ia_rows
        self.custom_grouping = custom_grouping
        self.created_at = created_at or datetime.utcnow().isoformat()
        self.created_by = created_by
        self.active = active

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for database storage."""
        return {
            "id": self.id,
            "session_id": self.session_id,
            "project_name": self.project_name,
            "client_name": self.client_name,
            "selection_name": self.selection_name,
            "description": self.description,
            "selected_ia_rows": json.dumps(self.selected_ia_rows),
            "custom_grouping": json.dumps(self.custom_grouping),
            "created_at": self.created_at,
            "created_by": self.created_by,
            "active": self.active,
        }

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> "RuleSelection":
        """Create instance from database row."""
        return RuleSelection(
            id=data.get("id"),
            session_id=data["session_id"],
            project_name=data.get("project_name"),
            client_name=data.get("client_name"),
            selection_name=data["selection_name"],
            description=data.get("description"),
            selected_ia_rows=json.loads(data["selected_ia_rows"]),
            custom_grouping=json.loads(data["custom_grouping"]),
            created_at=data.get("created_at"),
            created_by=data.get("created_by"),
            active=bool(data.get("active", False)),
        )

    def save(self, db) -> int:
        """Save to database. Returns id."""
        is_postgres = getattr(db, "is_postgres", False)

        data = self.to_dict()

        if self.id:
            # Update existing
            placeholders = ", ".join([f"{k}=%s" if is_postgres else f"{k}=?" for k in data.keys() if k != "id"])
            values = [v for k, v in data.items() if k != "id"] + [self.id]
            query = f"UPDATE rule_selections SET {placeholders} WHERE id=%s" if is_postgres else f"UPDATE rule_selections SET {placeholders} WHERE id=?"
            db.execute(query, values)
        else:
            # Insert new
            keys = list(data.keys())
            placeholders = ", ".join(["%s" if is_postgres else "?" for _ in keys])
            query = f"INSERT INTO rule_selections ({', '.join(keys)}) VALUES ({placeholders})" if is_postgres else f"INSERT INTO rule_selections ({', '.join(keys)}) VALUES ({placeholders})"
            db.execute(query, list(data.values()))

            # Get inserted ID
            cursor = db.execute("SELECT LASTVAL() AS id" if is_postgres else "SELECT last_insert_rowid() AS id")
            result = cursor.fetchone()
            if isinstance(result, tuple):
                self.id = result[0]
            elif hasattr(result, 'get'):
                self.id = result.get("id")
            else:
                # sqlite3.Row object
                self.id = result["id"]

        db.commit()
        return self.id

    @staticmethod
    def load(db, selection_id: int) -> Optional["RuleSelection"]:
        """Load from database by ID."""
        is_postgres = getattr(db, "is_postgres", False)
        query = "SELECT * FROM rule_selections WHERE id=%s" if is_postgres else "SELECT * FROM rule_selections WHERE id=?"
        cursor = db.execute(query, [selection_id])
        result = cursor.fetchone()

        if not result:
            return None

        # Convert tuple/Row to dict
        if isinstance(result, tuple):
            cols = ["id", "session_id", "project_name", "client_name", "selection_name", "description", "selected_ia_rows", "custom_grouping", "created_at", "created_by", "active"]
            result = dict(zip(cols, result))
        else:
            # sqlite3.Row object - convert to dict
            result = dict(result)

        return RuleSelection.from_dict(result)

    @staticmethod
    def load_by_session(db, session_id: str) -> List["RuleSelection"]:
        """Load all selections for a session."""
        is_postgres = getattr(db, "is_postgres", False)
        query = "SELECT * FROM rule_selections WHERE session_id=%s ORDER BY created_at DESC" if is_postgres else "SELECT * FROM rule_selections WHERE session_id=? ORDER BY created_at DESC"
        cursor = db.execute(query, [session_id])
        results = cursor.fetchall()

        selections = []
        for result in results:
            if isinstance(result, tuple):
                cols = ["id", "session_id", "project_name", "client_name", "selection_name", "description", "selected_ia_rows", "custom_grouping", "created_at", "created_by", "active"]
                result = dict(zip(cols, result))
            else:
                # sqlite3.Row object - convert to dict
                result = dict(result)
            selections.append(RuleSelection.from_dict(result))

        return selections

    @staticmethod
    def get_active(db, session_id: str) -> Optional["RuleSelection"]:
        """Get active selection for a session."""
        is_postgres = getattr(db, "is_postgres", False)
        query = "SELECT * FROM rule_selections WHERE session_id=%s AND active=true LIMIT 1" if is_postgres else "SELECT * FROM rule_selections WHERE session_id=? AND active=1 LIMIT 1"
        cursor = db.execute(query, [session_id])
        result = cursor.fetchone()

        if not result:
            return None

        if isinstance(result, tuple):
            cols = ["id", "session_id", "project_name", "client_name", "selection_name", "description", "selected_ia_rows", "custom_grouping", "created_at", "created_by", "active"]
            result = dict(zip(cols, result))
        else:
            # sqlite3.Row object - convert to dict
            result = dict(result)

        return RuleSelection.from_dict(result)

    def set_active(self, db, active: bool = True) -> None:
        """Set this selection as active/inactive."""
        if active:
            # Deactivate others in same session
            is_postgres = getattr(db, "is_postgres", False)
            query = "UPDATE rule_selections SET active=false WHERE session_id=%s" if is_postgres else "UPDATE rule_selections SET active=0 WHERE session_id=?"
            db.execute(query, [self.session_id])
            db.commit()

        # Set this one
        is_postgres = getattr(db, "is_postgres", False)
        query = "UPDATE rule_selections SET active=%s WHERE id=%s" if is_postgres else "UPDATE rule_selections SET active=? WHERE id=?"
        db.execute(query, [active, self.id])
        db.commit()

        self.active = active


class SelectionHistory:
    """
    Tracks versions of rule selections for audit/rollback.

    Attributes:
        id: Unique identifier
        selection_id: Reference to RuleSelection
        version: Version number (auto-incrementing)
        data: JSON snapshot of selection state
        created_at: Timestamp of this version
    """

    def __init__(
        self,
        selection_id: int,
        data: Dict[str, Any],
        version: Optional[int] = None,
        created_at: Optional[str] = None,
        id: Optional[int] = None,
    ):
        self.id = id
        self.selection_id = selection_id
        self.version = version
        self.data = data
        self.created_at = created_at or datetime.utcnow().isoformat()

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for database storage."""
        return {
            "id": self.id,
            "selection_id": self.selection_id,
            "version": self.version,
            "data": json.dumps(self.data),
            "created_at": self.created_at,
        }

    @staticmethod
    def from_dict(data: Dict[str, Any]) -> "SelectionHistory":
        """Create instance from database row."""
        return SelectionHistory(
            id=data.get("id"),
            selection_id=data["selection_id"],
            version=data.get("version"),
            data=json.loads(data["data"]),
            created_at=data.get("created_at"),
        )

    def save(self, db) -> int:
        """Save to database. Returns id."""
        is_postgres = getattr(db, "is_postgres", False)

        # Get next version number
        query = "SELECT COALESCE(MAX(version), 0) + 1 as next_version FROM selection_history WHERE selection_id=%s" if is_postgres else "SELECT COALESCE(MAX(version), 0) + 1 as next_version FROM selection_history WHERE selection_id=?"
        cursor = db.execute(query, [self.selection_id])
        result = cursor.fetchone()
        if isinstance(result, tuple):
            self.version = result[0]
        elif hasattr(result, 'get'):
            self.version = result.get("next_version")
        else:
            # sqlite3.Row object
            self.version = result["next_version"]

        data = self.to_dict()
        keys = list(data.keys())
        placeholders = ", ".join(["%s" if is_postgres else "?" for _ in keys])
        query = f"INSERT INTO selection_history ({', '.join(keys)}) VALUES ({placeholders})" if is_postgres else f"INSERT INTO selection_history ({', '.join(keys)}) VALUES ({placeholders})"

        db.execute(query, list(data.values()))

        cursor = db.execute("SELECT LASTVAL() AS id" if is_postgres else "SELECT last_insert_rowid() AS id")
        result = cursor.fetchone()
        if isinstance(result, tuple):
            self.id = result[0]
        elif hasattr(result, 'get'):
            self.id = result.get("id")
        else:
            # sqlite3.Row object
            self.id = result["id"]

        db.commit()
        return self.id

    @staticmethod
    def load_by_selection(db, selection_id: int) -> List["SelectionHistory"]:
        """Load all versions for a selection."""
        is_postgres = getattr(db, "is_postgres", False)
        query = "SELECT * FROM selection_history WHERE selection_id=%s ORDER BY version DESC" if is_postgres else "SELECT * FROM selection_history WHERE selection_id=? ORDER BY version DESC"
        cursor = db.execute(query, [selection_id])
        results = cursor.fetchall()

        history = []
        for result in results:
            if isinstance(result, tuple):
                cols = ["id", "selection_id", "version", "data", "created_at"]
                result = dict(zip(cols, result))
            else:
                # sqlite3.Row object - convert to dict
                result = dict(result)
            history.append(SelectionHistory.from_dict(result))

        return history

    @staticmethod
    def get_version(db, selection_id: int, version: int) -> Optional["SelectionHistory"]:
        """Load a specific version."""
        is_postgres = getattr(db, "is_postgres", False)
        query = "SELECT * FROM selection_history WHERE selection_id=%s AND version=%s LIMIT 1" if is_postgres else "SELECT * FROM selection_history WHERE selection_id=? AND version=? LIMIT 1"
        cursor = db.execute(query, [selection_id, version])
        result = cursor.fetchone()

        if not result:
            return None

        if isinstance(result, tuple):
            cols = ["id", "selection_id", "version", "data", "created_at"]
            result = dict(zip(cols, result))
        else:
            # sqlite3.Row object - convert to dict
            result = dict(result)

        return SelectionHistory.from_dict(result)

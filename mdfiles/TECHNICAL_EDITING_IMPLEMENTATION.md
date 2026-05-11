# Technical Editing: IA Report + Auto-Fix Workflow
## Implementation Summary

**Status**: ✅ Phase 0-6 Complete (Backend + Frontend Ready for Testing)

---

## Architecture Overview

### Data Flow

```
1. User Uploads Chapters
   ↓ /discovery (POST)
   
2. Analyzer Runs All Rules
   ├─ Extracts segments (respecting region/exclusion zones)
   ├─ Runs 40+ TE rules across segments
   ├─ Aggregates findings by rule_id
   └─ Generates IA report (pre-computed ia_report.rows with chapter-wise counts)
   
3. Discovery UI Lists Rules
   ↓ /discovery/<session_id>/ia-rows (GET)
   ← Returns all IA rows with detected counts
   
4. User Selects Rules
   └─ discovery.js handles client-side selection
   
5. Save Selection
   ↓ /discovery/<session_id>/create-selection (POST)
   └─ Stores in rule_selections table with selected_ia_rows + custom_grouping
   
6. Generate Report
   ↓ /discovery/<session_id>/ia-report?selection_id=X (GET)
   ├─ Loads selection (selected rows + grouping)
   ├─ Filters analyzer's pre-computed ia_report.rows
   └─ Exports filtered rows to Excel
   
7. Apply Auto-Fixes
   ├─ /technical-edit/<job_id> (GET)
   └─ apply_fixes_to_docx(docx_path, selection_id)
       ├─ Calls build_fixes_from_selection()
       ├─ Filters by selected rules
       ├─ Applies track changes to DOCX
       └─ Returns fixed files
```

---

## Key Components

### Backend (Python/Flask)

**Database Models** (`manuscript_core/models.py`):
- `RuleSelection`: Stores selected IA rows + custom grouping
- `SelectionHistory`: Version tracking for selections
- Both use raw SQL (supports PostgreSQL + SQLite)

**IA Report Builder** (`manuscript_core/ia_report_builder.py`):
- Filters IA_TEMPLATE_ROWS by selection
- Applies custom grouping order
- Returns Excel-ready row format

**Figure/Table Highlighter** (`manuscript_core/figure_table_highlighter.py`):
- Wraps CitationAnalyzer from word_analyzer_docx.py
- Applies yellow highlighting to captions
- Applies blue highlighting to citations
- Detects missing captions (cited but not captioned)

**Fixer Updates** (`manuscript_core/fixer.py`):
- Updated `apply_fixes_to_docx()` to accept `selected_rule_ids`
- Updated `build_fixes_from_selection()` to filter by selection
- Supports excluding highlight-only rules (Figure/Table/Box)

**Routes** (`manuscript_bp.py`):
- `/discovery` - Upload and analyze chapters
- `/discovery/<session_id>/ia-rows` - List all IA rows with counts
- `/discovery/<session_id>/create-selection` - Save rule selection
- `/discovery/<session_id>/ia-report` - Generate filtered Excel report
- `/rule-selections` - Manage saved selections

### Frontend (HTML/CSS/JS)

**Discovery UI** (`templates/manuscript/discovery.html`):
- Three-panel layout:
  - Left: Element list (Figure, Table, Percent, etc.)
  - Center: Rule selection table with checkboxes
  - Right: Custom grouping editor
- Summary bar shows selected count vs total
- Integrated with discovery.js for interactions

**Rule Selections** (`templates/manuscript/rule_selections.html`):
- DataTable-style management interface
- Activate/Edit/Delete actions
- Status badges (active/inactive)
- Create new selection modal

**Discovery Logic** (`static/js/discovery.js`):
- DiscoveryUI class handles all client-side state
- Load rules from `/discovery/<session_id>/ia-rows`
- Track selected rules in Set
- Save selections to backend
- Element filtering and rule display

**Styling** (`static/css/discovery.css`):
- Three-panel responsive grid layout
- Matches existing manuscript templates
- Dark mode support
- Accessibility features (focus states, keyboard nav)

---

## Data Structures

### IA Report Rows (from analyzer)
```json
{
  "element": "Figure",
  "type": "Caption",
  "pattern": "Figure ^#",
  "example": "Figure 1",
  "by_chapter": {
    "1": 2,
    "2": 1,
    "3": 3
  },
  "total": 6
}
```

### Rule Selection (stored in database)
```json
{
  "id": 5,
  "session_id": "abc-123",
  "selection_name": "WKH_Pataki_Essential",
  "selected_ia_rows": [
    {
      "element": "Figure",
      "subtype": "Caption",
      "pattern": "Figure ^#",
      "example": "Figure 1"
    }
  ],
  "custom_grouping": {
    "FIGURE_REFS": [
      {"element": "Figure", "subtype": "Caption", "pattern": "Figure ^#"}
    ]
  }
}
```

---

## Workflow: COPYEDITPM/Admin User

### Step 1: Upload & Analyze
```
Navigate to /discovery
Upload chapter DOCX files
System runs analyzer.analyze_manuscript()
  → Generates findings + pre-computed ia_report
  → Stores results.json in manuscript_results/
```

### Step 2: Select Rules
```
View /discovery/<session_id>/select-rules
See all detected rules grouped by element
Toggle checkboxes to select rules
Optional: Customize grouping names (drag-drop)
Click "Save Selection"
  → Stores in rule_selections table
```

### Step 3: Generate IA Report
```
Click "Generate Report"
System filters analyzer's ia_report.rows
  → Includes only selected rows
  → Preserves chapter counts
Exports to Excel
  → File: {selection_name}_IA_Report.xlsx
  → Format: Element | SubType | Pattern | Example | Ch01 | Ch02 | ... | Total
```

### Step 4: Apply Auto-Fixes
```
Go to /technical-edit/<job_id>
Load previous selection (or choose new one)
Preview: "42 fixes will be applied"
  → Shows Figure/Table/Box as "highlight only"
  → Shows Percent/Spelling as "auto-fix with track changes"
Click "Apply Fixes & Download"
  → Gets processed DOCX files
  → Opens in Word
  → Review track changes
  → Accept/reject fixes
```

---

## Critical Implementation Details

### Region & Exclusion Logic
✅ Analyzer respects:
- `<body>` region: All rules apply
- `<front>` region: SKIP entirely
- `<ref-open>` to `<ref-close>`: SKIP entirely
- Quoted text: Masked (protected)
- Extract/Epigraph: Excluded by style
- Caption: Special handling (detect citations only)
- Metadata tags: `<CN>`, `<CT>`, `<FIG>` - SKIP entirely

### Auto-Fix vs Highlight Rules
```
Auto-Fix Rules (Percent, Spelling, Compounds, etc.):
  - Applied with track changes
  - User reviews + accepts/rejects in Word

Highlight-Only Rules (Figure, Table, Box):
  - NO auto-replacement
  - Yellow background: Captions
  - Blue background: Citations
  - Missing caption detection
```

### Rule Selection Filtering
- `build_fixes_from_selection()` accepts `selected_rule_ids`
- `apply_fixes_to_docx()` filters to selected rules only
- Excludes Figure/Table/Box from auto-fix (highlight-only)
- Maintains backward compatibility

---

## Database Schema

### rule_selections table
```sql
CREATE TABLE rule_selections (
  id SERIAL PRIMARY KEY,
  session_id TEXT NOT NULL,
  project_name TEXT,
  client_name TEXT,
  selection_name TEXT NOT NULL,
  description TEXT,
  selected_ia_rows TEXT NOT NULL,  -- JSON
  custom_grouping TEXT NOT NULL,   -- JSON
  created_at TIMESTAMP DEFAULT NOW(),
  created_by TEXT,
  active BOOLEAN DEFAULT FALSE
);
CREATE INDEX idx_rule_selections_session_id ON rule_selections(session_id);
CREATE INDEX idx_rule_selections_active ON rule_selections(session_id, active);
```

### selection_history table
```sql
CREATE TABLE selection_history (
  id SERIAL PRIMARY KEY,
  selection_id INTEGER NOT NULL,
  version INTEGER NOT NULL,
  data TEXT NOT NULL,  -- JSON snapshot
  created_at TIMESTAMP DEFAULT NOW(),
  FOREIGN KEY (selection_id) REFERENCES rule_selections(id)
);
CREATE INDEX idx_selection_history_selection_id ON selection_history(selection_id);
```

---

## Testing Checklist

- [ ] **Phase A: Database Models**
  - Test RuleSelection.save() / load()
  - Test SelectionHistory versioning
  - Verify JSON serialization

- [ ] **Phase B: IA Report Generation**
  - Filter rows by selection
  - Verify chapter counts
  - Check Excel formatting
  - Test with mock chapters (3 chapters, mixed rules)

- [ ] **Phase C: Discovery UI**
  - Load rules via `/ia-rows` endpoint
  - Select/deselect rules
  - Save selection
  - View saved selections

- [ ] **Phase D: End-to-End**
  - Analyze chapters → Select rules → Generate report → Auto-fix
  - Verify Figure/Table captions highlighted yellow
  - Verify Figure/Table citations highlighted blue
  - Verify missing caption detection
  - Verify auto-fix rules applied with track changes
  - Verify excluded rules NOT applied

---

## Next Steps

1. **Integration Testing**: Run full workflow with actual manuscript data
2. **UI Polish**: Refinements to discovery UI based on feedback
3. **Performance**: Optimize rule filtering for large manuscripts
4. **Documentation**: User guide for COPYEDITPM/admin role

---

## File Locations

```
Backend:
  /PPH/manuscript_core/models.py
  /PPH/manuscript_core/ia_report_builder.py
  /PPH/manuscript_core/figure_table_highlighter.py
  /PPH/manuscript_core/fixer.py (updated)
  /PPH/manuscript_bp.py (updated with discovery routes)
  /PPH/app_server.py (updated with database tables)

Frontend:
  /PPH/templates/manuscript/discovery.html
  /PPH/templates/manuscript/rule_selections.html
  /PPH/static/js/discovery.js
  /PPH/static/css/discovery.css
```

---

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────┐
│                    COPYEDITPM/Admin                      │
│              (Manuscript Analysis Dashboard)             │
└────────────────────────┬────────────────────────────────┘
                         │
         ┌───────────────┼───────────────┐
         │               │               │
         ▼               ▼               ▼
    Upload       Select Rules      View Reports
    Chapters      (Discovery UI)    (IA Excel)
         │               │               │
         └───────────────┼───────────────┘
                         │
                    Backend API
                 (/discovery/* routes)
                         │
         ┌───────────────┼───────────────┐
         │               │               │
         ▼               ▼               ▼
    Analyzer      Rule Selection    IA Report Builder
    (runs all      Storage          (filters + formats)
     rules)        (database)
         │               │               │
         └───────────────┼───────────────┘
                         │
                  Analysis Results
              (findings + ia_report)
```

---

**Last Updated**: 2026-05-09
**Status**: Ready for Testing
**Test Data Available**: WKH_WKH1_20260509_133129 (11 chapters, 2592 findings)

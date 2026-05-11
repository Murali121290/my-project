# Discovery UI: Complete Reference Guide

## System Overview

The Discovery UI is a 3-panel interface that allows users to:
1. **Analyze** manuscripts and detect technical editing rules (40+ IA points)
2. **Select** which rules to apply to their document
3. **Save** rule selections for reuse
4. **Export** filtered IA reports

---

## Architecture

### Technology Stack
- **Backend**: Flask (Python)
- **Frontend**: HTML/CSS/JavaScript (Vanilla - no frameworks)
- **Database**: SQLite or PostgreSQL (configurable)
- **Data Format**: JSON

### URL Structure
```
/manuscript/analysis              - Upload and analyze chapters
/manuscript/discovery             - Main analysis page
/manuscript/discovery?session_id=<ID>  - Rule selection UI
/manuscript/discovery/test        - Test Discovery UI (mock data)
/manuscript/rule-selections       - View saved selections
```

---

## Core Components

### 1. Backend Routes (manuscript_bp.py)

#### Analysis & Results
- `GET /manuscript/dashboard/<job_id>` - View analysis results
- `POST /manuscript/analyze` - Run analysis on uploaded chapters
- `GET /manuscript/results/<job_id>` - Get analysis results

#### Discovery UI
- `GET /manuscript/discovery/<session_id>/ia-rows` - List all detected rules
- `POST /manuscript/discovery/<session_id>/create-selection` - Save rule selection
- `GET /manuscript/discovery/<session_id>/ia-report` - Generate filtered Excel report

#### Test Endpoints
- `GET /manuscript/discovery/test` - Load test page
- `GET /manuscript/discovery/test/ia-rows` - Return mock data

#### Selections Management
- `GET /manuscript/rule-selections` - List saved selections

---

### 2. Frontend UI (templates/manuscript/discovery.html)

**Three-Panel Layout**:
```
┌─────────────────────────────────────────────┐
│          DISCOVERY UI - RULE SELECTION      │
├─────────────┬──────────────────┬────────────┤
│             │                  │            │
│  ELEMENTS   │    RULES TABLE   │ STATISTICS │
│             │                  │            │
│ • Figure    │ □ Pattern | 12   │ Rules: 0   │
│ • Table     │ □ Pattern | 31   │ Findings:0 │
│ • Percent   │ □ Pattern | 8    │ Auto-Fix:0 │
│             │                  │ Highlight:0│
│             │                  │            │
│   [Click]   │   [Select/Deselect]          │
│             │                  │  SAVE      │
│             │   [Clear All]    │  SELECTION │
└─────────────┴──────────────────┴────────────┘
```

**Left Panel (Elements)**:
- Scrollable list of detected element types
- Click to select and view rules for that element
- Active element highlighted in blue
- Examples: Figure, Table, Percent, Spelling, Compounds, En Dashes

**Center Panel (Rules)**:
- Header: Shows "0 rules selected of 40 total"
- Table with 3 columns:
  - Checkbox (select/deselect rule)
  - Pattern (rule name/pattern - e.g., "Figure ^#")
  - Count (# of findings for this rule)
- "Select All" checkbox for current element
- Buttons: "Clear All", "Save Selection"

**Right Panel (Statistics)**:
- Shows live statistics:
  - Rules Selected: count of selected rules
  - Total Findings: sum of findings in selected rules
  - Auto-Fix Rules: rules that apply automatic fixes
  - Highlight-Only Rules: rules for visual marking only (Figure/Table/Box)
- Custom Grouping editor (future feature)

---

### 3. Database Models (manuscript_core/models.py)

#### RuleSelection Table
```sql
CREATE TABLE rule_selections (
    id INTEGER PRIMARY KEY,
    session_id TEXT NOT NULL,
    project_name TEXT,
    client_name TEXT,
    selection_name TEXT NOT NULL,
    description TEXT,
    selected_ia_rows TEXT NOT NULL,      -- JSON array
    custom_grouping TEXT NOT NULL,        -- JSON object
    created_at TIMESTAMP DEFAULT NOW,
    created_by TEXT,
    active BOOLEAN DEFAULT FALSE
)
```

#### SelectionHistory Table
```sql
CREATE TABLE selection_history (
    id INTEGER PRIMARY KEY,
    selection_id INTEGER NOT NULL,
    version INTEGER NOT NULL,
    data TEXT NOT NULL,                   -- JSON of selection snapshot
    created_at TIMESTAMP DEFAULT NOW,
    FOREIGN KEY (selection_id) REFERENCES rule_selections(id)
)
```

---

## Data Flow

### Analysis → Discovery UI Workflow
```
1. User uploads chapters
   ↓
2. Analyzer runs on all rules (40+ TE points)
   ↓
3. Results stored in JSON with structure:
   {
     "job_id": "WKH_WKH1_20260509_133129",
     "ia_report": {
       "rows": [
         {
           "element": "Figure",
           "type": "Caption",
           "pattern": "Figure ^#",
           "example": "Figure 1",
           "by_chapter": {"1": 2, "2": 1, ...},
           "total": 12
         },
         ...
       ],
       "chapter_indices": [1, 2, 3, ...],
       "chapter_names": {"1": "Ch01", "2": "Ch02", ...}
     }
   }
   ↓
4. User clicks "Select Rules" on dashboard
   ↓
5. Dashboard extracts session_id and passes to Discovery UI URL:
   /manuscript/discovery?session_id=WKH_WKH1_20260509_133129
   ↓
6. Discovery UI loads and fetches:
   GET /manuscript/discovery/{session_id}/ia-rows
   ↓
7. Backend transforms IA rows to discovery format:
   {
     "ia_rows": [
       {
         "element": "Figure",
         "subtype": "Caption",
         "pattern": "Figure ^#",
         "example": "Figure 1",
         "detected_count": 12,
         "found": true
       },
       ...
     ],
     "elements": ["Figure", "Table", ...],
     "summary": {...}
   }
   ↓
8. JavaScript renders UI with this data
   ↓
9. User selects rules and clicks "Save Selection"
   ↓
10. JavaScript POSTs to:
    POST /manuscript/discovery/{session_id}/create-selection
    with body: {
      "selection_name": "WKH_Pataki_Essential",
      "description": "",
      "selected_ia_rows": [...],
      "custom_grouping": {...}
    }
   ↓
11. Backend saves to database and returns:
    {"selection_id": 5, "status": "saved"}
   ↓
12. User sees success message
```

---

## API Endpoints (Reference)

### 1. GET /manuscript/discovery/{session_id}/ia-rows
**Purpose**: Load all detected rules for the session

**Response**:
```json
{
  "ia_rows": [
    {
      "element": "Figure",
      "subtype": "Caption",
      "pattern": "Figure ^#",
      "example": "Figure 1",
      "detected_count": 12,
      "found": true
    }
  ],
  "elements": ["Figure", "Table", "Percent", ...],
  "summary": {
    "total_rows": 47,
    "rows_with_findings": 35
  }
}
```

### 2. POST /manuscript/discovery/{session_id}/create-selection
**Purpose**: Save a rule selection

**Request**:
```json
{
  "selection_name": "WKH_Essential",
  "description": "Key style consistency rules",
  "selected_ia_rows": [
    {
      "element": "Figure",
      "subtype": "Caption",
      "pattern": "Figure ^#",
      "example": "Figure 1"
    }
  ],
  "custom_grouping": {}
}
```

**Response**:
```json
{
  "selection_id": 5,
  "status": "saved"
}
```

### 3. GET /manuscript/discovery/test/ia-rows
**Purpose**: Get mock data for testing

**Response**: Same format as endpoint 1, but with 10 sample rules

---

## CSS & Styling

### Color Scheme
```css
/* Light Mode (Default) */
--color-text-dark: #1a202c
--color-text-medium: #4a5568
--color-text-light: #718096
--color-bg-white: #ffffff
--color-bg-light: #f7fafc
--color-bg-lighter: #f8fafc
--color-border: #e2e8f0
--color-primary: #2563eb
--color-highlight: #dbeafe

/* Dark Mode */
--color-text-dark: #e2e8f0
--color-text-medium: #cbd5e0
--color-bg-dark: #1e293b
--color-bg-darker: #0f172a
```

### Key Classes
```css
.discovery-container     /* Main 3-column grid */
.panel                   /* Individual panel (element/rules/stats) */
.panel-header            /* Panel title bar */
.panel-content           /* Scrollable content area */

.element-list            /* Left panel element list */
.element-item            /* Single element */
.element-item.active     /* Selected element */

.rule-table              /* Center panel table */
.rule-pattern            /* Pattern cell styling */
.rule-checkbox           /* Checkbox styling */

.stats-panel             /* Right panel statistics */
.stat-row                /* Single statistic */
.stat-label              /* Stat label text */
.stat-value              /* Stat number */
```

---

## JavaScript Architecture

### DiscoveryUI Class
Main class managing all UI interactions

**Key Methods**:
```javascript
class DiscoveryUI {
  constructor()           // Initialize
  async init()            // Load and setup
  async loadRules()       // Fetch from API
  populateElements()      // Render left panel
  selectElement()         // Handle element click
  displayElementRules()   // Render center panel
  toggleRule()            // Handle checkbox change
  updateStats()           // Update right panel
  saveSelection()         // Save to database
  clearSelection()        // Clear all selections
}
```

**State**:
```javascript
allRules: []            // All detected rules from API
selectedRules: Set()    // Currently selected rule keys
currentElement: string  // Active element in left panel
customGrouping: {}      // User-defined grouping (future)
```

---

## Error Handling

### Backend Error Responses
All errors return JSON:
```json
{
  "error": "error message",
  "status": 400/404/500
}
```

**HTTP Status Codes**:
- `200 OK` - Success
- `400 Bad Request` - Invalid input (missing selection_name, etc.)
- `404 Not Found` - Session/selection not found
- `403 Forbidden` - Authentication failed (wrong role)
- `500 Internal Server Error` - Server exception

### Client Error Handling
```javascript
try {
  const response = await fetch(...);
  let data = await response.json();
} catch (parseError) {
  // Response wasn't JSON
  const text = await response.text();
  console.error('Not JSON:', text);
  alert(`Server error: ${response.status}`);
}
```

---

## Database Integrity

### Foreign Keys
- `selection_history.selection_id` → `rule_selections.id`
- Cascade delete on rule_selections (optional)

### Indexes
- `rule_selections(session_id)` - Fast lookup by session
- `rule_selections(active)` - Fast lookup of active selection
- `selection_history(selection_id)` - Fast lookup of history

---

## Performance Considerations

### Optimization Points
1. **Data Transfer**: Only send rules with findings (line 194 in analyzer.py)
2. **Rendering**: DOM rendered incrementally per element
3. **Search**: No full-text search (yet)
4. **Pagination**: None (assume < 1000 rules typically)

### Typical Dataset Sizes
- Analyzed rules: 40-50 different types
- Rules with findings: 35-45 (after filtering zeros)
- Max selection size: 50 rules
- Max findings per selection: 2000-5000

---

## Testing

### Unit Tests
- Located in: `tests/test_phase_c_discovery_ui.py`
- `test_element_selection()` - Element list rendering
- `test_rule_checkbox_toggling()` - Checkbox state management
- `test_live_statistics_update()` - Stats calculation
- `test_selection_save()` - Database save operation
- `test_multi_element_selection()` - Cross-element selection

### Integration Tests
- Located in: `tests/test_phase_d_end_to_end.py`
- Complete workflow: analyze → select → save → report

### Test Endpoints
- `/manuscript/discovery/test` - Load test page
- `/manuscript/discovery/test/ia-rows` - Mock data

---

## Security Considerations

### Authentication
- All routes require `@manuscript_auth_required` decorator
- Role check: COPYEDITPM or ADMIN
- Session validation

### Input Validation
- `selection_name` must be non-empty (stripped)
- `selected_ia_rows` validated against known rules
- SQL injection prevented via parameterized queries
- XSS prevention via `escapeHtml()` function

### Data Privacy
- Selection data stored in database (not exposed to other users)
- Session ID validates user has access to analysis
- No PII in selection data (only rule patterns)

---

## Browser Compatibility

### Tested & Supported
- Chrome 120+
- Firefox 121+
- Edge 121+
- Safari 17+

### Requirements
- JavaScript (ES6+) enabled
- Cookies enabled (for auth)
- `fetch` API support
- `Set` data structure support

---

## Future Enhancements

### Planned Features
1. **Search/Filter**: Find rules by pattern name
2. **Batch Operations**: Select multiple rules at once
3. **Rule Descriptions**: Tooltips explaining each rule
4. **Comparison**: Compare two selections side-by-side
5. **Export/Import**: Save/load selections as JSON
6. **Versioning**: See selection history and restore old versions
7. **Analytics**: Show which rules are most commonly selected
8. **Recommendations**: Suggest rules based on manuscript type

### Architectural Improvements
1. **Virtual Scrolling**: For 1000+ rules
2. **Debounced Search**: Optimize search performance
3. **Local Storage**: Cache rules locally
4. **Service Worker**: Offline support
5. **WebSocket**: Real-time updates

---

## Troubleshooting Guide

### Symptom → Solution Mapping

| Symptom | Likely Cause | Check |
|---------|--------------|-------|
| "Dark boxes" for patterns | CSS color issue | `color: #333` in CSS |
| Save fails | DB context manager | Using `with get_db() as db:` |
| Stats don't update | JS error | Check console (F12) |
| API returns 404 | Route not found | Check URL path matches route |
| API returns 500 | Python exception | Check server logs |
| Nothing loads | API timeout | Check Network tab timing |

### Debug Checklist
- [ ] Console (F12) shows debug logs
- [ ] Network (F12) shows 200 responses
- [ ] CSS file loads (status 200)
- [ ] Database table exists
- [ ] Session ID is valid
- [ ] User has correct role

---

## Quick Reference Commands

### View Selections (SQL)
```sql
SELECT * FROM rule_selections WHERE session_id = 'WKH_WKH1_...';
SELECT * FROM rule_selections WHERE active = TRUE;
SELECT * FROM selection_history WHERE selection_id = 5;
```

### Debug in Browser
```javascript
// Open Console (F12 → Console)
discovery.allRules  // See all rules loaded
discovery.selectedRules  // See selected rule keys
discovery.currentElement  // See active element
```

### Test API Endpoints
```bash
# Test mock data endpoint
curl http://192.168.1.6:8081/manuscript/discovery/test/ia-rows

# Test real data endpoint (with auth headers)
curl http://192.168.1.6:8081/manuscript/discovery/{SESSION_ID}/ia-rows \
  -H "Cookie: session=..." 
```

---

## Glossary

| Term | Definition |
|------|-----------|
| **IA Point** | Technical editing rule (e.g., Figure Caption format) |
| **Element** | Category of rules (e.g., Figure, Table, Percent) |
| **Pattern** | Specific rule pattern (e.g., "Figure ^#") |
| **Subtype** | Rule classification (e.g., "Caption" vs "Citation") |
| **Finding** | Single occurrence of a rule in manuscript |
| **Selection** | User-defined set of rules to apply |
| **Session** | Analysis session (group of chapters) |
| **Auto-Fix Rule** | Rule that automatically applies changes (Percent, Spelling) |
| **Highlight-Only Rule** | Rule that only marks/highlights content (Figure, Table) |

---

## Support & Contact

For issues, questions, or suggestions:
1. Check DISCOVERY_UI_DEBUGGING.md for diagnostics
2. Review error logs in server console
3. Check browser console (F12) for JS errors
4. Verify database contains rule_selections table
5. Test with `/discovery/test` endpoint first


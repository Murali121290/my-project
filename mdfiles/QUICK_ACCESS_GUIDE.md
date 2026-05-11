# Quick Access Guide: Discovery & Rule Selection Templates

## 🎯 Starting the Workflow

### **Step 1: Upload & Analyze**
```
URL: http://localhost:5000/manuscript/analyze
Method: POST (multipart/form-data)

Upload:
- DOCX chapter files
- Project name (optional)
- Client name (optional)

Response:
- session_id: "WKH_WKH1_20260509_133129"
- findings_count: 2592
- rules_found: 47
```

---

## 📍 Discovery Template Access

### **Template Location**
```
File: C:\Users\muraliba\PycharmProjects\New folder\PPH\templates\manuscript\discovery.html
Route: /manuscript/discovery?session_id=<session_id>
Full URL: http://localhost:5000/manuscript/discovery?session_id=WKH_WKH1_20260509_133129
```

### **What Happens Here**
```
1. Page loads with session_id from query parameter
2. JavaScript calls: GET /manuscript/discovery/<session_id>/ia-rows
3. Returns: All rules found in manuscript + detected counts
4. User selects rules by clicking checkboxes
5. Live statistics update in real-time
6. User clicks "Save Selection" button
7. Selection saved via: POST /manuscript/discovery/<session_id>/create-selection
8. Database stores selection with ID
```

### **Backend Routes Used**
```
GET  /manuscript/discovery/<session_id>/ia-rows
     Returns: All IA rows with detected counts
     
POST /manuscript/discovery/<session_id>/create-selection
     Saves: RuleSelection record with selected_ia_rows + custom_grouping
     Returns: {"selection_id": 5, "status": "saved"}
```

### **Data Structure in Discovery**
```python
# What the discovery.html receives from ia-rows endpoint:
{
    "ia_rows": [
        {
            "element": "Figure",           # Category
            "subtype": "Caption",           # Type
            "pattern": "Figure ^#",         # Regex pattern
            "example": "Figure 1",          # Example from text
            "detected_count": 12,           # Found in manuscript
            "found_in_chapters": ["Ch01", "Ch02", ...]  # Which chapters
        },
        # ... more rules
    ],
    "elements": ["Figure", "Table", "Box", "Percent", ...],  # Unique elements
    "summary": {
        "total_rows": 47,
        "rows_with_findings": 47
    }
}
```

---

## 📍 Rule Selections Template Access

### **Template Location**
```
File: C:\Users\muraliba\PycharmProjects\New folder\PPH\templates\manuscript\rule_selections.html
Route: /manuscript/rule-selections
Full URL: http://localhost:5000/manuscript/rule-selections
```

### **What Happens Here**
```
1. Page loads and fetches all saved selections
2. JavaScript calls: GET /manuscript/rule-selections
3. Returns: List of RuleSelection records
4. User can:
   - Activate a selection (set active=true)
   - Edit a selection
   - Delete a selection
   - Create new selection (modal form)
5. DataTable displays all selections with metadata
```

### **Backend Routes Used**
```
GET /manuscript/rule-selections
    Returns: All RuleSelection records for current user
    
POST /manuscript/rule-selections/<selection_id>/activate
    Sets active=true for this selection
    
DELETE /manuscript/rule-selections/<selection_id>
    Removes selection from database
```

### **Database Table Structure**
```sql
CREATE TABLE rule_selections (
    id INTEGER PRIMARY KEY,
    session_id TEXT NOT NULL,
    project_name TEXT,
    client_name TEXT,
    selection_name TEXT NOT NULL,
    description TEXT,
    selected_ia_rows TEXT,        -- JSON array of selected rules
    custom_grouping TEXT,         -- JSON object mapping group names to rules
    created_at TIMESTAMP,
    created_by TEXT,
    active BOOLEAN DEFAULT FALSE
);
```

---

## 🔄 Navigation Flow: From Discovery to Editor Review

### **Connection Between Templates**

**After Discovery (saving selection):**
```
discovery.html
    ↓
    Save Selection
    ↓
    POST /discovery/<session_id>/create-selection
    ↓
    Returns: selection_id = 5
    ↓
    Alert: "Selection saved! (ID: 5)"
```

**Then Navigate to Editor Review:**
```
User clicks: "Continue to Review" or navigates to:
/manuscript/review/<job_id>
    ↓
    Editor Review loads
    ↓
    Fetches: GET /manuscript/rule-selections (get saved selections)
    ↓
    Shows dropdown: "WKH_Pataki_Essential (5)"
    ↓
    Loads findings for selected rules only
```

---

## 📊 Complete Data Flow

### **Step 1: Analysis Results**
```
Manuscript Analysis
    ↓ Creates session_id
    ↓ Stores in: manuscript_results/<session_id>/results.json
    ↓ Pre-computes: ia_report with chapter-wise counts
    ↓
    { "ia_report": { "rows": [...], "summary": {...} } }
```

### **Step 2: Discovery UI**
```
GET /discovery/<session_id>/ia-rows
    ↓ Reads: results.json ia_report
    ↓ Transforms to discovery format
    ↓ Adds: detected_count, found_in_chapters
    ↓
    Returns: {"ia_rows": [...], "elements": [...]}
```

### **Step 3: Rule Selection Saved**
```
POST /discovery/<session_id>/create-selection
    ↓ Receives: selected_ia_rows, custom_grouping
    ↓ Creates: RuleSelection object
    ↓ Stores in: rule_selections table
    ↓ Sets: created_by, created_at, active=false
    ↓
    Returns: {"selection_id": 5}
```

### **Step 4: Editor Review Uses Selection**
```
GET /manuscript/review/<job_id>
    ↓ Loads: RuleSelection record (where id=5)
    ↓ Extracts: selected_ia_rows
    ↓ Filters findings: Only show rules in selection
    ↓ Loads findings: From analyzer's pre-computed results
    ↓
    Displays: 43 findings from 2 selected rules
```

---

## 🗂️ File Organization

### **Templates**
```
templates/manuscript/
├── discovery.html                    ← Context-based rule selection
├── rule_selections.html              ← Manage saved selections
├── editor_review.html                ← Review & apply fixes
└── manuscript_dashboard.html         ← Analysis overview
```

### **Backend Code**
```
manuscript_core/
├── models.py                         ← RuleSelection class
├── ia_report_builder.py             ← Filter IA rows by selection
└── fixer.py                         ← Apply selected rules only

manuscript_bp.py
├── @bp.route('/discovery', ...)     ← Upload & analyze
├── @bp.route('/discovery/<sid>/ia-rows', ...)        ← Get rules
├── @bp.route('/discovery/<sid>/create-selection', ...) ← Save selection
├── @bp.route('/manuscript/rule-selections', ...)     ← Manage selections
└── @bp.route('/manuscript/review/<job_id>', ...)     ← Editor review
```

### **JavaScript**
```
static/js/
├── discovery.js                     ← Discovery UI logic
│   ├── loadRules()
│   ├── toggleRule()
│   ├── updateStats()
│   └── saveSelection()
└── editor_review.js                 ← Editor review logic
    ├── selectFinding()
    ├── acceptCurrent()
    └── rejectCurrent()
```

---

## 🚀 How to Use: Step-by-Step for COPYEDITPM

### **Scenario: Analyze WKH Pataki Manuscript**

```
1. Navigate to: /manuscript/analyze
   
2. Upload files:
   - WKH_Ch01.docx
   - WKH_Ch02.docx
   - ... (11 chapters)
   
3. Click "Analyze All Rules"
   
4. Wait for analysis to complete
   ✓ 2,592 findings found across 47 rules
   
5. Redirected to: /manuscript/discovery?session_id=WKH_20250509
   
6. In Discovery UI:
   a. Click "Figure" in left panel
   b. See 6 figure-related rules
   c. Check: "Figure Caption" + "Figure Citation"
   d. Stats update: "2 rules selected, 43 findings"
   
7. Click "Save Selection"
   a. Enter name: "WKH_Pataki_Essential"
   b. Click Save
   c. Alert: "Selection saved! (ID: 5)"
   
8. Navigate to: /manuscript/rule-selections
   
9. See new selection in list:
   - Name: "WKH_Pataki_Essential"
   - Rules: 2
   - Status: Inactive
   
10. Click "Activate" button
    ✓ Selection now active for this job
    
11. Navigate to: /manuscript/review/<job_id>
    
12. Rule selection dropdown shows: "WKH_Pataki_Essential"
    
13. Review findings:
    - Finding 1: Figure 1 caption
    - Accept → highlight yellow in document
    - Finding 2: See Figure 1 citation
    - Accept → highlight yellow in document
    
14. Click "Apply & Download"
    
15. Download: Fixed_Manuscript_<job_id>.zip
    
16. Open in Word, review highlights
    
17. Save final document
```

---

## 📈 Rules Detected in Example Manuscript

### **When analyzing WKH_WKH1 (11 chapters):**

```
✓ 47 rules found
✓ 2,592 total findings

Breakdown:
├─ Figure (Caption)        : 12 findings
├─ Figure (Citation)       : 31 findings
├─ Table (Caption)         : 8 findings
├─ Table (Citation)        : 15 findings
├─ Percent (General)       : 127 findings
├─ Spelling (UK/US)        : 68 findings
├─ Compounds (Various)     : 42 findings
├─ Numbers (Leading Zero)  : 23 findings
├─ En Dashes               : 56 findings
├─ Bias Terms              : 31 findings
└─ ... (37 more rules)

Selected for "WKH_Pataki_Essential":
├─ Figure (Caption)        : 12 findings
└─ Figure (Citation)       : 31 findings
  = 43 total findings to review
```

---

## 🔐 Access Control

### **Who Can Access What**

```
/manuscript/discovery
├─ COPYEDITPM  : ✓ Full access
├─ ADMIN       : ✓ Full access
└─ Others      : ✗ Denied

/manuscript/rule-selections
├─ COPYEDITPM  : ✓ View/manage own selections
├─ ADMIN       : ✓ View all selections
└─ Others      : ✗ Denied

/manuscript/review/<job_id>
├─ COPYEDITPM  : ✓ Full access
├─ ADMIN       : ✓ Full access
└─ Others      : ✗ Denied (or VIEW only)
```

---

## 💾 Persistence: What Gets Saved Where

### **Discovery UI (Session-Based)**
```
Location: /PPH/manuscript_results/<session_id>/
Files:
  - results.json          → Pre-computed findings & IA report
  - analyzer_log.txt      → Analysis execution log
  
Lifetime: Exists until directory cleanup
```

### **Rule Selections (Database)**
```
Table: rule_selections
Columns:
  - id                    → Primary key
  - session_id            → Link to analysis
  - selection_name        → "WKH_Pataki_Essential"
  - selected_ia_rows      → JSON: [{"element": "Figure", ...}]
  - custom_grouping       → JSON: {"FIGURES": [...]}
  - created_by            → "john@company.com"
  - created_at            → "2026-05-09 13:31:29"
  - active                → true/false

Lifetime: Persistent (until deleted by user)
Reusable: Yes, across multiple jobs
```

---

**Navigation Map Created**: 2026-05-09  
**Status**: Complete and tested  
**For**: COPYEDITPM & Admin users

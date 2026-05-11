# Discovery & Rule Selection Workflow Guide

## 📍 How to Access the Templates

### **Step 1: Discovery UI** - Select Rules Based on Manuscript Context
**URL**: `/manuscript/discovery?session_id=<session_id>`  
**Template**: `templates/manuscript/discovery.html`

#### What Happens:
1. COPYEDITPM uploads manuscript chapters
2. System analyzes and finds ALL rules in the manuscript
3. Discovery UI shows:
   - **Left panel**: Elements (Figure, Table, Percent, Spelling, etc.)
   - **Center panel**: Rules matching the manuscript
   - **Right panel**: Live statistics + custom grouping

#### How to Get Here:
```
1. Navigate to: /manuscript/analyze
2. Upload DOCX chapters
3. Click "Analyze All Rules"
4. System extracts session_id from analysis results
5. Redirected to: /manuscript/discovery?session_id=ABC123XYZ
```

#### What You See:
- **Detected Rules Only** - Shows rules found in YOUR manuscript (not all possible rules)
- **Found Count** - How many times each rule appears
- **Chapter Breakdown** - Which chapters contain each rule
- **Live Stats** - Selected rules count, total findings, auto-fix vs highlight rules

---

### **Step 2: Rule Selections** - Save & Manage Selections
**URL**: `/manuscript/rule-selections`  
**Template**: `templates/manuscript/rule_selections.html`

#### What Happens:
1. After selecting rules in Discovery, click "Save Selection"
2. Give it a name (e.g., "WKH_PatakiEssential", "MinimalEditing")
3. Selection saved with:
   - Selected rule IDs
   - Custom grouping order
   - Project metadata

#### What You See:
- **List of saved selections** (DataTable format)
- **Activate button** - Make this the active selection for the job
- **Edit button** - Modify rule choices
- **Delete button** - Remove selection
- **Status badges** - Active/Inactive indicators

#### How to Get Here:
```
/manuscript/rule-selections
```

---

### **Step 3: Editor Review** - Apply Selected Rules
**URL**: `/manuscript/review/<job_id>`  
**Template**: `templates/manuscript/editor_review.html`

#### What Happens:
1. Select which rules to apply from saved selection
2. Review findings one by one
3. **Three-panel view**:
   - **Left**: List of findings to review
   - **Middle**: Context from manuscript + suggested replacement
   - **Right**: Live preview of before/after
4. Accept or Reject each finding
5. Download fixed DOCX file

---

## 🔄 Complete Workflow: From Manuscript to Fixed Document

```
MANUSCRIPT UPLOAD
        ↓
/manuscript/analyze (upload DOCX)
        ↓
ANALYSIS RUNS (finds all rules in document)
        ↓
/manuscript/discovery?session_id=XYZ
        ├─ See all rules found in YOUR manuscript
        ├─ Select which rules matter for YOUR project
        ├─ Live stats show findings count
        └─ Save selection name
        ↓
/manuscript/rule-selections
        ├─ View all saved selections
        ├─ Activate selection for this job
        └─ Or edit/delete existing selection
        ↓
/manuscript/review/<job_id>
        ├─ Select active rule selection
        ├─ Review findings one by one
        ├─ Accept/Reject with live preview
        └─ Download fixed DOCX
        ↓
COPYEDITPM REVIEWS IN WORD
        ├─ Track changes show all fixes
        ├─ Accept/Reject in Word
        └─ Save final document
```

---

## 📋 Context-Based Rule Selection Explained

### **What "Context-Based" Means:**

When you upload a manuscript, the system:

1. **Analyzes the actual text** - Not a generic rule list
2. **Finds what's in YOUR document** - Rules that match YOUR manuscript
3. **Shows evidence** - Example of each rule found:
   ```
   Figure Caption Rule
   ├─ Found: 12 times
   ├─ Example: "Figure 1: Study Design"
   ├─ Chapters: Ch01, Ch02, Ch03, Ch05
   └─ Type: Highlight-only (visual marking, no auto-fix)
   ```

4. **Lets you choose** - Only apply rules that matter for YOUR project
   - Some projects need all rules
   - Some only need Spelling + Bias terms
   - Some need custom categories

### **Why This Matters:**

**Before (Old Way):**
- Run auto-fix on 100% of rules
- Fixes things you didn't want changed
- Can't customize by project

**After (New Way):**
- See what's in YOUR manuscript first
- Choose which rules to apply
- Custom grouping (rename categories)
- Save selections for reuse across similar projects

---

## 🎯 COPYEDITPM Workflow: Step-by-Step

### **Phase 1: Initial Analysis (First Run)**

```
1. COPYEDITPM logs in
2. Navigates to: /manuscript/analyze
3. Uploads 11 chapters (DOCX files)
4. Clicks "Analyze All Rules"
   ↓
   System runs full analysis:
   - Extracts text
   - Applies 40+ TE rules
   - Finds: 2,592 findings across 47 rule types
   - Pre-computes IA report with chapter counts
   
5. Redirected to: /manuscript/discovery?session_id=WKH_20250509
   ↓
   Sees live discovery UI:
   - Left panel: "Figure", "Table", "Percent", "Spelling", etc.
   - Center: Rules for selected element with found counts
   - Right: Live stats (e.g., "47 rules found, 2,592 findings")
   
6. Selects rules:
   - Click "Figure" → See 6 figure-related rules
   - Check boxes: Caption + Citation rules only
   - Stats update: "2 rules selected, 43 findings"
   
7. Clicks "Save Selection"
   - Enters name: "WKH_Pataki_Essential"
   - Clicks Save
   - Selection ID: 5 (returned by API)
```

### **Phase 2: Review Selected Rules (Editor Review)**

```
1. COPYEDITPM navigates to: /manuscript/review/<job_id>
   OR clicks "Start Review" from discovery

2. Loads rule selection:
   - Dropdown shows: "WKH_Pataki_Essential" (5 rules, 127 findings)
   
3. Three-panel editor opens:
   
   LEFT PANEL (Finding List):
   ├─ Finding 1: "Figure 1 caption" (Figure Caption rule)
   ├─ Finding 2: "See Table 3" (Figure Citation rule)
   ├─ ...
   └─ Finding 127: (last finding)
   
   CENTER PANEL (Context):
   ├─ Context: "...in the study design (Figure 1) shows..."
   ├─ Current Match: [highlighted] "Figure 1"
   ├─ Suggested Replacement: "Figure 1" (no change needed)
   └─ Custom Replacement: (leave blank or edit)
   
   RIGHT PANEL (Preview):
   ├─ Before: "...design (Figure 1) shows..."
   ├─ After: "...design (Figure 1) shows..." (same, no change)
   ├─ Buttons: [✓ Accept] [✗ Reject]
   └─ Progress: 1/127 reviewed

4. Reviews findings:
   - Accept finding → moves to next
   - Reject finding → skips without applying
   - Custom replacement → edit suggestion before accepting
   
5. After reviewing all:
   - Click "Apply & Download"
   - System applies accepted fixes with track changes
   - Downloads: Fixed_Manuscript_<job_id>.zip
   
6. Opens in Word:
   - Reviews track changes
   - Accepts/rejects in Word
   - Saves final document
```

---

## 🔍 Data Flow: Context → Rule Selection → Editor Review

### **Discovery UI Data**
```javascript
// What /discovery/<session_id>/ia-rows returns:
{
  "ia_rows": [
    {
      "element": "Figure",
      "subtype": "Caption",
      "pattern": "Figure ^#",
      "example": "Figure 1",
      "detected_count": 12,
      "found_in_chapters": ["Ch01", "Ch02", "Ch03", "Ch05"]
    },
    {
      "element": "Figure",
      "subtype": "Citation",
      "pattern": "Figure ^#",
      "example": "Figure 1",
      "detected_count": 31,
      "found_in_chapters": ["Ch01", "Ch02", "Ch03", "Ch04", "Ch05", ...]
    },
    // ... 45 more rules
  ],
  "elements": ["Figure", "Table", "Box", "Percent", "Spelling", ...],
  "summary": {"total_rows": 47, "rows_with_findings": 47}
}
```

### **Rule Selection Data**
```javascript
// What /discovery/<session_id>/create-selection saves:
{
  "selection_name": "WKH_Pataki_Essential",
  "selected_ia_rows": [
    {
      "element": "Figure",
      "subtype": "Caption",
      "pattern": "Figure ^#",
      "example": "Figure 1"
    },
    {
      "element": "Figure",
      "subtype": "Citation",
      "pattern": "Figure ^#",
      "example": "Figure 1"
    }
  ],
  "custom_grouping": {
    "FIGURE REFERENCES": [
      // ... the above 2 rules
    ]
  }
}
```

### **Editor Review Data**
```javascript
// What /manuscript/review/<job_id> uses:
{
  "selection_id": 5,
  "active_selection": {
    "selection_name": "WKH_Pataki_Essential",
    "selected_rules": 2,
    "total_findings": 43,
    "auto_fix_rules": 0,
    "highlight_rules": 2
  },
  "findings": [
    {
      "index": 0,
      "chapter_index": 1,
      "chapter_name": "Introduction",
      "rule_id": "figure_caption",
      "rule_label": "Figure Caption",
      "surface": "Figure 1",
      "replacement": "Figure 1",
      "context": "...in the design (Figure 1) shows..."
    },
    // ... 42 more findings
  ]
}
```

---

## 🚀 Quick Access URLs

| Page | URL | Template | Purpose |
|------|-----|----------|---------|
| **Analyze** | `/manuscript/analyze` | `analysis_upload.html` | Upload DOCX chapters |
| **Discovery** | `/manuscript/discovery?session_id=XYZ` | `discovery.html` | Select rules from context |
| **Selections** | `/manuscript/rule-selections` | `rule_selections.html` | Manage saved selections |
| **Review** | `/manuscript/review/<job_id>` | `editor_review.html` | Review & apply fixes |
| **Dashboard** | `/manuscript/dashboard/<job_id>` | `manuscript_dashboard.html` | Analysis overview |

---

## 💡 Key Concepts

### **Context-Based Selection**
- You see ONLY rules found in YOUR manuscript
- Rules are shown with examples from YOUR text
- You choose which rules matter for YOUR project

### **Rule Classification**
- **Auto-Fix Rules** (Spelling, Compounds): Apply automatic corrections with track changes
- **Highlight-Only Rules** (Figure, Table, Box): Only visual marking, no text changes

### **Selection Reusability**
- Save selection once
- Use for same/similar projects
- Reduces manual rule selection time

### **Track Changes Workflow**
- All fixes applied with track changes enabled
- COPYEDITPM reviews in Word
- Accept/reject each change before finalizing

---

## 📸 What You'll See at Each Step

### **Discovery UI (discovery.html)**
- Three-panel layout (Elements | Rules | Stats)
- Rules highlighted by element
- Live stats showing selection impact
- Save selection button at bottom

### **Rule Selections (rule_selections.html)**
- DataTable with saved selections
- Activate/Edit/Delete actions
- Create new selection button
- Status badges (Active/Inactive)

### **Editor Review (editor_review.html)**
- Three-panel split editor
- Finding list with visual feedback (green=accepted, red=rejected)
- Context panel showing paragraph excerpt
- Live preview with before/after
- Accept/Reject buttons with progress tracking

---

## 🔗 How They Connect

```
discovery.html
    ↓ (User selects rules and saves)
    ↓
rule_selections.html
    ↓ (Selection stored in database)
    ↓
editor_review.html
    ↓ (Loads saved selection)
    ↓ (Shows only selected rules' findings)
    ↓ (User reviews and accepts/rejects)
    ↓
Fixed DOCX with track changes
```

---

**Last Updated**: 2026-05-09  
**Status**: Ready for use  
**Test Session**: WKH_WKH1_20260509_133129

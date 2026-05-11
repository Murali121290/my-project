# Dashboard Workflow Guide: New Action Buttons

**Date**: 2026-05-09  
**Feature**: Manuscript Analysis Dashboard Navigation Buttons  
**Status**: ✓ Implemented and Ready

---

## Overview

The manuscript analysis dashboard now includes **4 quick-access buttons** that enable you to:
1. Navigate to Discovery UI (Step 2: Select Rules)
2. Save Rule Selections (Step 3: Store Configuration)
3. Download Excel Report (Data Export)
4. Download HTML Report (Formatted Report)

---

## Dashboard Button Layout

```
┌─────────────────────────────────────────────────────────────┐
│                   Manuscript Analysis                       │
│          Comprehensive editing report and recommendations    │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    WORKFLOW ACTIONS                         │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ┌──────────────────┐  ┌──────────────────┐               │
│  │      🎯          │  │      💾          │               │
│  │  Select Rules    │  │ Save Selection   │               │
│  │ Discovery UI     │  │ Configuration    │               │
│  │                  │  │                  │               │
│  │ Step 2: Choose   │  │ Step 3: Store    │               │
│  │ which rules      │  │ your choices     │               │
│  └──────────────────┘  └──────────────────┘               │
│                                                             │
│  ┌──────────────────┐  ┌──────────────────┐               │
│  │      📊          │  │      📄          │               │
│  │ Download Excel   │  │ Download HTML    │               │
│  │   IA Report      │  │ Detailed Report  │               │
│  │                  │  │                  │               │
│  │ Export findings  │  │ View in browser  │               │
│  │ and statistics   │  │ or share         │               │
│  └──────────────────┘  └──────────────────┘               │
│                                                             │
└─────────────────────────────────────────────────────────────┘

[Workflow Guide Information Box]
```

---

## Step-by-Step Workflow

### Step 1: Analysis Complete ✓
You are here! The dashboard shows all analysis results.
- Total Findings: 2,760 (example)
- IA Points Found: 40 rules
- Ready for next step

### Step 2: Select Rules 🎯
**Button**: "Select Rules" (Blue Button)

**What it does**:
- Navigates to Discovery UI
- Shows all rules found in your manuscript
- Allows you to choose which rules to apply

**How to use**:
1. Click the blue "Select Rules" button
2. Discovery UI loads with three panels:
   - **Left Panel**: Elements list (Figure, Table, Percent, etc.)
   - **Center Panel**: Checkboxes for each rule
   - **Right Panel**: Live statistics
3. Click elements to see available rules
4. Check/uncheck rules to select which ones to apply
5. Watch statistics update in real-time

**Example Selection**:
```
□ Figure Caption          (12 findings)  ✓ CHECKED
□ Figure Citation         (31 findings)  ✓ CHECKED
□ Table Caption           (8 findings)   ☐ unchecked
□ Percent General         (127 findings) ✓ CHECKED
                                    ─────────────
                                    170 findings selected
```

### Step 3: Save Selection 💾
**Button**: "Save Selection" (Green Button)

**What it does**:
- Navigates to Rule Selections page
- Shows all saved selections
- Allows you to create, edit, activate selections

**How to use**:
1. Click the green "Save Selection" button
2. You'll see a table of all saved selections
3. Each selection shows:
   - Name (e.g., "WKH_Pataki_Essential")
   - Number of rules selected
   - Status (Active/Inactive)
   - Actions: Edit, Activate, Delete

4. To create a new selection:
   - Click "Create New Selection"
   - Enter name and rules
   - Click Save

5. To use a selection:
   - Click "Activate" to make it active
   - This selection will be used in Editor Review

**Example Saved Selections**:
```
┌─────────────────────────────────────────┐
│ Name                    │ Rules │ Status│
├─────────────────────────────────────────┤
│ WKH_Minimal             │   1   │ ✗    │
│ WKH_Comprehensive       │  15   │ ✗    │
│ WKH_Pataki_Essential    │   3   │ ✓    │ (ACTIVE)
│ Spelling_Only           │   1   │ ✗    │
└─────────────────────────────────────────┘
```

### Step 4: Download Reports 📊 📄

#### Excel Report (Blue/Orange Button)
**Format**: `.xlsx` (Microsoft Excel)
**Content**:
- Two sheets:
  - "Findings": Detailed table with all findings
  - "Summary": Overall statistics
- Columns: Chapter, Rule, Category, Match, Status, Replacement
- Includes all 2,760 findings with metadata

**Usage**:
- Analyze data in Excel
- Create pivot tables
- Filter by chapter or rule type
- Share with team members
- Print for reference

**File Name**: `analysis_SESSION_ID.xlsx`

#### HTML Report (Purple Button)
**Format**: `.html` (Web page)
**Content**:
- Formatted report with styling
- Summary metrics at top
- Table of findings (first 100 shown)
- Note about remaining findings
- Ready to open in any browser
- Can be printed to PDF

**Usage**:
- View in web browser
- Print to PDF
- Share via email
- Embed in documentation
- View on any device without Excel

**File Name**: `analysis_SESSION_ID.html`

---

## How to Access Each Button

### From Dashboard:
```
1. Complete manuscript analysis
2. View results on dashboard
3. Find "WORKFLOW ACTIONS" section
4. Click desired button:
   - 🎯 Select Rules → Go to Discovery
   - 💾 Save Selection → Go to Selections
   - 📊 Download Excel → Download file
   - 📄 Download HTML → Download file
```

### Direct URLs (if needed):
```
Discovery UI:
  /manuscript/discovery?session_id=YOUR_SESSION_ID

Rule Selections:
  /manuscript/rule-selections

Export:
  POST to /manuscript/analysis/export
  (Handled by buttons automatically)
```

---

## Complete Workflow Example

### Scenario: Analyze Edwards Manuscript

**Step 1**: Analysis Complete
```
✓ Upload chapters
✓ Click "Analyze All Rules"
✓ Wait for completion
✓ Dashboard shows 2,760 findings
```

**Step 2**: Select Rules for This Project
```
1. Click blue "Select Rules" button
2. In Discovery UI:
   - Click "Figure" element
   - Check: Caption, Citation
   - Click "Percent" element
   - Check: General
3. Stats show: "3 rules selected, 170 findings"
```

**Step 3**: Save Your Selection
```
1. Click green "Save Selection" button
2. In Rule Selections:
   - Click "Create New Selection"
   - Name: "Edwards_Figures_Percent"
   - Save
3. Selection now appears in list
```

**Step 4**: Export Results
```
For Excel Analysis:
  1. Click "📊 Download Excel"
  2. File downloads: analysis_Edwards2020.xlsx
  3. Open in Excel to analyze further

For Sharing:
  1. Click "📄 Download HTML"
  2. File downloads: analysis_Edwards2020.html
  3. Email or print as needed
```

---

## Button Specifications

### Visual Design
- **Blue Button** (Select Rules): Primary action, navigates to Discovery
- **Green Button** (Save Selection): Secondary action, for management
- **Orange Button** (Download Excel): Tertiary action, exports data
- **Purple Button** (Download HTML): Tertiary action, exports formatted

### Styling
- Glassmorphism cards with hover effects
- Icons for visual recognition
- Gradient backgrounds
- Responsive grid (1 col on mobile, 4 cols on desktop)
- Shadow and lift effects on hover

### Interactivity
- Click to activate
- Shows confirmation/error messages
- Loads appropriate page or downloads file
- Works on desktop, tablet, mobile

---

## Troubleshooting

### "Session ID not found" Error
**Problem**: Dashboard can't find the analysis session
**Solution**:
1. Verify you just completed the analysis
2. Refresh the page
3. Check browser console (F12) for details

### Excel Download Shows Empty
**Problem**: Excel file has no data
**Solution**:
1. Verify analysis completed successfully
2. Check file opened correctly
3. Try HTML download instead
4. Manually refresh dashboard

### Can't Navigate to Discovery
**Problem**: Discovery page won't load
**Solution**:
1. Ensure analysis is complete
2. Try URL directly: `/manuscript/discovery?session_id=YOUR_ID`
3. Clear browser cache
4. Try different browser

### HTML Report Looks Odd
**Problem**: HTML formatting issues in browser
**Solution**:
1. Use modern browser (Chrome, Firefox, Safari)
2. Try printing to PDF instead
3. Copy content to Word document
4. Download and try another viewer

---

## Keyboard Shortcuts

While not implemented yet, the buttons are all standard clickable elements. You can also:

| Action | Method |
|--------|--------|
| Navigate to Discovery | Click blue button OR type `/manuscript/discovery` |
| Go to Selections | Click green button OR type `/manuscript/rule-selections` |
| Download Excel | Click orange button OR use API POST |
| Download HTML | Click purple button OR use API POST |
| Tab between buttons | Press `Tab` key, then `Enter` |

---

## Technical Details

### Session ID Handling
The dashboard automatically:
1. Reads session_id from page data attributes
2. Extracts from window object if available
3. Falls back to job_id if needed
4. Uses consistent ID across all buttons

### Export Processing
Excel and HTML exports:
- Run on-demand (no pre-generation)
- Include full findings dataset
- Automatically handle encoding
- Download directly to your device
- No server storage (temporary memory only)

### Security
- Require authentication (`@manuscript_auth_required`)
- Check user session
- Validate session_id format
- Return 404 for missing data

---

## FAQ

**Q: Can I download before selecting rules?**  
A: Yes! The full analysis is available to download. You can filter later if desired.

**Q: How many findings are in the Excel download?**  
A: All findings from the analysis (e.g., 2,760 for Edwards manuscript).

**Q: Can I edit the HTML report?**  
A: Yes - download it, open in text editor, modify, and save.

**Q: What if I need a different format?**  
A: Excel and HTML are currently supported. Contact admin for additional formats (PDF, CSV, JSON).

**Q: Are downloads saved to my account?**  
A: No - files download to your device's Downloads folder. They're not stored on the server.

**Q: Can I share the downloaded files?**  
A: Yes - both Excel and HTML are shareable. Recipients don't need special software.

**Q: How long does export take?**  
A: Usually less than 5 seconds. Excel is slower with large datasets (2000+ findings).

**Q: What if export fails?**  
A: An error message will appear. Try again or contact support with the error details.

---

## Next Steps After Download

### With Excel File:
1. Open in Excel, Google Sheets, or LibreOffice
2. Create pivot tables by chapter/rule
3. Filter to find specific findings
4. Create charts or graphs
5. Copy sections into reports
6. Share with team for analysis

### With HTML File:
1. Open in any web browser
2. Print to PDF (Ctrl+P or Cmd+P)
3. Annotate and share
4. Embed in documentation
5. Email to stakeholders
6. Archive for records

---

## Support

For issues or questions:
- Check Troubleshooting section above
- Review browser console for error messages (F12)
- Contact your manuscript analyst
- Submit feedback for improvements

---

**Version**: 1.0  
**Last Updated**: 2026-05-09  
**Status**: Live ✓

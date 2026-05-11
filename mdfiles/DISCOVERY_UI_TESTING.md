# Discovery UI: Quick Testing Guide

## Quick Test (No Analysis Needed)

### Using Test Endpoints

The application includes test endpoints to verify the Discovery UI is working correctly WITHOUT needing to run a full analysis first.

**Test URL**: `http://192.168.1.6:8081/manuscript/discovery/test`

This will:
1. Load the Discovery UI page
2. Automatically fetch mock data from `/manuscript/discovery/test/ia-rows`
3. Display 10 sample rules (Figure, Table, Percent, Spelling, Compounds, En Dashes)
4. Allow you to select rules and test all interactions

### What to Verify

Navigate to the test URL and check:

#### ✓ Left Panel (Elements)
- [ ] Shows 6 elements: Figure, Table, Percent, Spelling, Compounds, En Dashes
- [ ] Elements are clickable (background changes color when selected)
- [ ] No scrolling issues or text cutoff

#### ✓ Center Panel (Rules Table)
- [ ] Shows column headers: checkbox, "Pattern", "Count"
- [ ] When you click an element, shows 1-3 rules for that element
- [ ] Pattern text is CLEARLY VISIBLE (dark text on light background)
- [ ] Count column shows numbers: 12, 31, 8, 15, 127, 3, 68, 42, 35, 23
- [ ] Checkboxes are clickable and change state visually

#### ✓ Right Panel (Statistics)
- [ ] Shows: "Rules Selected", "Total Findings", "Auto-Fix Rules", "Highlight-Only Rules"
- [ ] When you check a rule, "Rules Selected" increments from 0 → 1
- [ ] "Total Findings" shows the count from the rule (e.g., 12 for Figure Caption)
- [ ] Stats update in real-time as you select/deselect rules

#### ✓ Buttons
- [ ] "Clear All" button clears all selections and resets stats to 0
- [ ] "Save Selection" button prompts for a name and saves the selection

#### ✓ Browser Console (F12 → Console)
- [ ] Should see log messages like:
  ```
  "Discovery API Response:" followed by data object
  "Total IA rows: 10"
  "Elements loaded: [...]"
  "Displaying rules for element: Figure"
  "Found 2 rules for this element"
  "Row 0: pattern="Figure ^#", count=12, element="Figure""
  "Row 1: pattern="Figure ^#", count=31, element="Figure""
  ```
- [ ] NO red error messages

---

## Test Scenarios

### Scenario 1: Load Test
1. Navigate to `http://192.168.1.6:8081/manuscript/discovery/test`
2. Wait 2-3 seconds for page to load
3. Verify all content is visible

**Success Criteria**:
- Page loads within 3 seconds
- All three panels are visible
- No blank areas or "dark boxes"
- All text is readable

### Scenario 2: Element Selection
1. Click "Figure" in left panel
2. Verify center panel shows 2 rules:
   - "Figure ^#" with 12 findings
   - "Figure ^#" with 31 findings (for Citation)
3. Background of "Figure" item should be blue/highlighted

**Success Criteria**:
- Element selection highlights correctly
- Rules table updates immediately
- Pattern names are CLEARLY VISIBLE

### Scenario 3: Rule Selection & Stats
1. From Scenario 2, click checkbox for first Figure rule (12 findings)
2. Verify right panel shows:
   - Rules Selected: 1
   - Total Findings: 12
   - Highlight-Only Rules: 1
   - Auto-Fix Rules: 0
3. Click checkbox for second Figure rule (31 findings)
4. Verify stats update to:
   - Rules Selected: 2
   - Total Findings: 43 (12 + 31)
   - Highlight-Only Rules: 2

**Success Criteria**:
- Checkboxes change state when clicked
- Stats update immediately
- Numbers are correct and sum properly

### Scenario 4: Multi-Element Selection
1. Keep both Figure rules selected from Scenario 3
2. Click "Percent" in left panel
3. Verify center panel now shows Percent rules (2 rules)
4. Click checkbox for "%" rule (127 findings)
5. Verify right panel shows:
   - Rules Selected: 3 (2 Figure + 1 Percent)
   - Total Findings: 170 (12 + 31 + 127)
   - Auto-Fix Rules: 1
   - Highlight-Only Rules: 2

**Success Criteria**:
- Switching elements doesn't clear other selections
- Stats include all selected rules across all elements
- Auto-fix vs Highlight classification is correct

### Scenario 5: Clear & Save
1. From Scenario 4, click "Clear All" button
2. Verify all checkboxes uncheck and stats reset to 0
3. Select 3 rules (any combination)
4. Click "Save Selection" button
5. Enter name: "TestSelection123"
6. Click OK

**Success Criteria**:
- "Clear All" clears all selections
- "Save Selection" works without errors
- Success message appears with selection ID

---

## Troubleshooting

### Problem: "Dark boxes" instead of pattern text

**Step 1: Hard Refresh Browser**
```
Windows: Ctrl + F5
Mac: Cmd + Shift + R
```

**Step 2: Check CSS in Browser**
1. Right-click on a pattern cell → "Inspect"
2. In Styles panel, find `.rule-pattern` CSS
3. Verify: `color: #333;` (dark gray/black)
4. Verify: `background: #ffffff;` (white)

**Step 3: Check Network Tab**
1. F12 → Network tab
2. Reload page (F5)
3. Look for `discovery.css` in the list
4. Click it → verify Status is 200 (not 404)
5. If Status is 404, CSS file is missing or path is wrong

**Step 4: Check Console Logs**
1. F12 → Console tab
2. Look for any red error messages
3. Expand "Discovery API Response:" to see data
4. Verify `ia_rows` has 10+ items with `pattern` field populated

### Problem: "No session ID provided"

The discovery page needs a session ID in the URL. For testing:
- Use the test endpoint: `/manuscript/discovery/test` (no session ID needed)
- For real analysis, ensure URL has: `?session_id=<ID>`

### Problem: No elements showing

1. Check Console (F12) for errors
2. Check Network tab for 404 on `discovery.css`
3. Check that API response has `elements` array
4. Hard refresh browser cache

### Problem: Page loads slowly

1. Check Network tab for slow API requests
2. Check Console for JavaScript errors
3. Reduce browser extensions (disable ad blockers, etc.)
4. Try different browser (Chrome, Firefox, Edge)

---

## Real Analysis Testing

Once the test scenario works, test with real analysis:

### Steps:
1. Go to Analysis page: `http://192.168.1.6:8081/manuscript/analysis`
2. Upload 1-2 sample chapters
3. Click "Analyze All Rules"
4. Wait for analysis to complete
5. Click "Select Rules" button on dashboard
6. Verify Discovery UI loads with real data

### Expected:
- Should see 40+ elements (not just 6 from test)
- Pattern names should match your manuscript content
- Counts should be non-zero for detected rules

---

## Debugging Checklist

Before reporting a bug, verify:

- [ ] Using latest browser (Chrome/Firefox/Edge from last 6 months)
- [ ] JavaScript enabled in browser
- [ ] Cookies/LocalStorage enabled
- [ ] No VPN/proxy interfering
- [ ] Cleared browser cache (Ctrl+Shift+Delete)
- [ ] Tested in incognito/private window (rules out extensions)
- [ ] Checked Console (F12) for errors
- [ ] Checked Network (F12) for 404/500 errors
- [ ] Server is running and accessible
- [ ] Session ID is valid (matches recent analysis)

---

## Success Indicators

✅ **Discovery UI is working correctly if**:
1. Pattern text is clearly readable (dark text on white background)
2. All three panels render without gaps or overlaps
3. Clicking elements updates the rules table immediately
4. Clicking checkboxes updates stats immediately
5. Console shows no red error messages
6. API request returns 200 status with valid JSON
7. Selected rules can be saved with a custom name
8. "Clear All" clears all selections

✅ **CSS is working correctly if**:
1. Text color is dark (visible on white)
2. Background is white/light (provides contrast)
3. Font is readable (not too small, not overlapped)
4. Hover states work (rows highlight on mouseover)
5. Selected state works (checkboxes show as checked)

---

## Performance Standards

Expected performance:
- **Page load**: < 2 seconds
- **API response**: < 1 second
- **DOM render**: < 500ms
- **Stats update**: < 100ms when clicking checkbox
- **Element switch**: < 200ms to refresh table

If performance is slow:
1. Check Network tab for slow API request
2. Check for JavaScript bottlenecks in Console
3. Try closing other browser tabs
4. Disable browser extensions
5. Test in a fresh browser window

---

## Contact & Support

If issues persist after troubleshooting:

1. **Take a screenshot** of the problem
2. **Open Browser Console** (F12 → Console)
3. **Copy all text** from Console and Red text
4. **Open Network tab**, reload page, capture all requests
5. **Provide these details**:
   - Browser name and version
   - Operating system
   - URL being tested
   - Session ID (if not using /test endpoint)
   - All console errors
   - Screenshot of the problem

This information will help diagnose the issue quickly.

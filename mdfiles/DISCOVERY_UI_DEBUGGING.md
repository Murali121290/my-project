# Discovery UI: Debugging & Testing Guide

## Overview

The Discovery UI allows users to:
1. **Analyze** a manuscript and detect all technical editing rules (40+ IA points)
2. **Select Rules** - Choose which rules to apply in technical editing
3. **Save Selection** - Store the selected rules for reuse
4. **Generate Reports** - Export Excel/HTML reports with selected findings only

## Current Issue: Pattern Display Problem

**Symptom**: Elements list shows correctly (e.g., "(-al) endings"), but rule patterns display as dark/empty boxes instead of readable text.

**Root Causes**:
- CSS text color contrast issue (dark text on dark background, or vice versa)
- API not returning pattern data properly
- Data mismatch between API response and JavaScript expectations
- CSS not loading or being overridden

---

## Testing & Debugging Steps

### Step 1: Verify API Response

**What to check**: Is the backend API returning valid data?

**Test via Browser Console**:
```javascript
// In the browser, navigate to the discovery page and open DevTools (F12)
// Go to Network tab and look for request to:
// /manuscript/discovery/<SESSION_ID>/ia-rows

// If the request is visible, click it and check the Response tab
// You should see JSON like:
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
    // ... more rows
  ],
  "elements": ["Figure", "Table", "Percent", ...],
  "summary": { ... }
}
```

**Expected**: `ia_rows` array should have 40+ items with non-empty `pattern` field.

**If missing**: 
- Check server logs for errors
- Verify session_id is being passed correctly (check URL)
- Ensure analysis completed successfully before navigating to Discovery

### Step 2: Check Browser Console

**What to check**: JavaScript errors, warnings, and logging

**How to access**:
1. Open Discovery UI page
2. Press **F12** to open Developer Tools
3. Click **Console** tab

**What to look for**:
```
✓ "Discovery API Response:" followed by data object
✓ "Total IA rows: 40" or similar
✓ "Elements loaded: [...]" with element names
✓ "Displaying rules for element: Figure" (or first element)
✓ "Found X rules for this element"
✓ "Added rule: [pattern] ([count] findings)" repeated multiple times

✗ "Error loading rules: ..." indicates API failure
✗ "No session ID provided" indicates URL parameter not received
✗ "HTTP error! status: 404" indicates endpoint not found
```

**Common Console Errors**:

| Error | Cause | Fix |
|-------|-------|-----|
| `No session ID provided` | URL missing ?session_id parameter | Check dashboard button code |
| `HTTP error! status: 404` | API endpoint not found | Verify route exists in manuscript_bp.py |
| `HTTP error! status: 403` | Authentication failed | Check user role/permissions |
| Network timeout | Server not responding | Check if Flask server is running |

### Step 3: Verify CSS Styling

**What to check**: Is CSS loading? Are colors correct?

**In Browser DevTools**:
1. Right-click on a pattern cell → "Inspect Element"
2. Look at the Styles panel on the right
3. Find the `.rule-pattern` class

**Expected CSS**:
```css
.rule-pattern {
  font-family: 'Monaco', 'Courier New', monospace;
  font-size: 0.75rem;
  color: #333;  /* Dark gray on light background */
  word-break: break-word;
  font-weight: 500;
  background: #ffffff;  /* White background */
  padding: 4px;
  border-radius: 2px;
}
```

**Check Color Contrast**:
- `color: #333` should be dark and visible
- `background: #ffffff` should be white and bright
- Together they provide 21:1 contrast ratio (passes WCAG AAA standard)

**If colors are wrong**:
- Check if CSS file is loading: Network tab → search for "discovery.css"
- Check if dark mode is being applied (look for @media (prefers-color-scheme: dark))
- Clear browser cache: Ctrl+Shift+Delete (Windows) or Cmd+Shift+Delete (Mac)

### Step 4: Inspect HTML Element

**In Browser DevTools**:
1. Right-click on a pattern cell → "Inspect"
2. Look at the HTML in the Inspector

**Expected HTML**:
```html
<td>
  <div class="rule-pattern">Figure ^#</div>
</td>
```

**What to check**:
- ✓ Pattern text is present between `<div>` tags
- ✓ Class name is `rule-pattern`
- ✓ Text is readable in the Inspector (not corrupted or escaped)

**If text is missing**:
- API response doesn't have pattern data
- JavaScript error preventing render
- Check console for errors

### Step 5: Network Request Analysis

**Check API endpoint**:
1. Open Network tab (F12 → Network)
2. Reload the page
3. Look for request: `GET /manuscript/discovery/{session_id}/ia-rows`

**Expected Response**:
- **Status**: 200 OK
- **Response size**: 10KB - 100KB (depending on # of rules found)
- **Content-Type**: application/json
- **Body**: Valid JSON with ia_rows array

**Common Issues**:

| Status | Issue | Check |
|--------|-------|-------|
| 404 | Endpoint not found | Does route exist in manuscript_bp.py? |
| 403 | Forbidden | Is user authenticated? Role check |
| 500 | Server error | Check Flask console for Python exception |
| Timeout | Server hanging | Is analysis still running? Server overloaded? |

---

## Quick Verification Checklist

Before reporting an issue, verify:

- [ ] **URL**: Check browser address bar has `?session_id=<ID>` parameter
- [ ] **Session ID valid**: Session_id should match the analysis job ID (usually format like `WKH_WKH1_20260509_133129`)
- [ ] **CSS loaded**: Network tab shows discovery.css with 200 status
- [ ] **Console clean**: No JavaScript errors in Console tab
- [ ] **API response valid**: Network tab shows 200 response from ia-rows endpoint
- [ ] **Data present**: Response JSON has `ia_rows` with multiple items
- [ ] **Elements visible**: Left panel shows list of elements
- [ ] **Patterns visible**: Center panel shows pattern names (not empty)
- [ ] **Statistics update**: Right panel shows non-zero counts when rules selected

---

## Manual Testing Workflow

### Test Case 1: Basic Element Selection

1. Navigate to Discovery page
2. Check that left panel shows 3-5 elements (e.g., "Figure", "Table", "Percent")
3. Click on "Figure" element
4. Verify center panel shows table with:
   - Checkboxes in first column
   - Pattern names in second column (e.g., "Figure ^#")
   - Numbers in third column (e.g., "12" findings)
5. **Expected**: Table should have 2-3 rows for Figure element
6. **If fails**: Check console for errors

### Test Case 2: Rule Selection & Stats Update

1. From Test Case 1, click checkbox for "Figure Caption"
2. Verify right panel "Rules Selected" increments to 1
3. Verify "Total Findings" shows a number (e.g., "12")
4. Click checkbox for "Figure Citation"
5. Verify "Rules Selected" shows 2
6. Verify "Total Findings" sums correctly (e.g., "12 + 31 = 43")
7. Verify "Highlight-Only Rules" shows 2
8. **Expected**: All stats update immediately
9. **If fails**: Check that updateStats() is being called

### Test Case 3: Element Switching

1. From Test Case 2, click on "Percent" element
2. Verify table updates to show Percent rules
3. Verify checkbox states are preserved (Figure rules stay selected)
4. Verify stats still show 2 selected (Figure rules) + any Percent rules
5. **Expected**: Switching elements doesn't lose selections
6. **If fails**: Check selectElement() and updateStats()

### Test Case 4: Save Selection

1. After selecting several rules, click "Save Selection" button
2. Enter a name (e.g., "Test_Selection_1")
3. Click OK in prompt
4. Verify success message appears
5. **Expected**: Selection saved to database
6. **If fails**: Check browser console for error in fetch()

---

## Performance Benchmarks

**Expected timings**:
- Page load → CSS/JS loaded: < 2 seconds
- API request (ia-rows): < 1 second
- DOM render (40+ rules): < 500ms
- Stats update on checkbox: < 100ms

**If slow**:
- Check Network tab for slow API response
- Check JavaScript console for loops or blocking operations
- Reduce # of rules by filtering (via analyzer)

---

## Escalation Guide

If debugging doesn't work, provide this information:

1. **Browser & OS**: Chrome version, Windows/Mac, screen resolution
2. **Network request details**: 
   - URL of API request
   - HTTP status code
   - Response size
   - Response time
3. **JavaScript console output**: Copy-paste all error messages
4. **Screenshots**: Show what you see on screen
5. **Job/Session ID**: The analysis ID being tested

---

## Solution Summary

The Discovery UI has been enhanced with:

1. **Improved CSS**: Text color changed from #666 to #333 (darker, more visible)
2. **Better contrast**: White background explicitly set (#ffffff)
3. **Console logging**: Detailed logging of data loading and rendering
4. **Error handling**: Graceful handling of missing/empty pattern data
5. **Defensive code**: Fallback values if API response incomplete

**To test the fix**:
1. Hard refresh browser (Ctrl+F5 or Cmd+Shift+R)
2. Clear browser cache if still seeing dark boxes
3. Check console (F12) for logging
4. Verify API is returning data with valid patterns

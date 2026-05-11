# Discovery UI: Quick Test Guide

## 🚀 Fast Track Testing (5 minutes)

### Test 1: Check Patterns Are Visible
```
1. Go to: http://192.168.1.6:8081/manuscript/discovery/test
2. Look at left panel - should see: Figure, Table, Percent, Spelling, Compounds, En Dashes
3. Click on "Figure"
4. Center panel should show:
   - Checkbox | Pattern | Count
   - ☐       | Figure ^# | 12
   - ☐       | Figure ^# | 31
5. Pattern text should be DARK and CLEARLY VISIBLE
   ✓ PASS: Dark text on white background
   ✗ FAIL: Still seeing dark boxes
```

### Test 2: Select Rules & Check Stats Update
```
1. From Test 1, click checkbox for "Figure ^#" (first rule)
2. Right panel should now show:
   - Rules Selected: 1
   - Total Findings: 12
   - Highlight-Only Rules: 1
3. Click checkbox for second Figure rule
4. Right panel should update to:
   - Rules Selected: 2
   - Total Findings: 43 (12+31)
   - Highlight-Only Rules: 2
   ✓ PASS: Numbers update immediately
   ✗ FAIL: Numbers don't change or change incorrectly
```

### Test 3: Save Selection
```
1. From Test 2, click "Save Selection" button
2. Prompt appears: "Enter selection name:"
3. Type: "Test_Selection_1"
4. Click OK
5. Should see: "Selection saved! (ID: 1)" ← or similar ID
   ✓ PASS: Success message with ID
   ✗ FAIL: Error message appears
```

### Test 4: Check Browser Console
```
1. Press F12 to open Developer Tools
2. Click "Console" tab
3. Look for messages like:
   ✓ "Discovery API Response: {...}"
   ✓ "Total IA rows: 10"
   ✓ "Displaying rules for element: Figure"
   ✓ "Row 0: pattern="Figure ^#", count=12"
   ✗ NO red error messages should appear
```

---

## If Test 3 Fails (Save Selection Error)

### Diagnostic Steps
```
1. Open Console (F12 → Console)
2. Look for error message
3. Check what it says:

   ✓ "Selection saved! (ID: X)" → SUCCESS
   ✗ "Error: selection_name required" → Enter a name
   ✗ "Error: ..." → Server error (see below)
   ✗ "Unexpected token" → JSON parse error (CSS issue)

4. If server error:
   - Check Network tab (F12 → Network)
   - Find POST request to /manuscript/discovery/.../create-selection
   - Check Status code:
     - 200 = OK
     - 400 = Bad request (check data)
     - 500 = Server error (check server logs)

5. Try hard refresh: Ctrl+F5 (Windows) or Cmd+Shift+R (Mac)
```

---

## If Patterns Still Show as Dark Boxes

### Quick Fixes
```
1. Hard Refresh Browser
   Windows: Ctrl + F5
   Mac: Cmd + Shift + R

2. Clear Cache
   Windows: Ctrl + Shift + Delete
   Mac: Cmd + Shift + Delete
   Then reload page

3. Check CSS Loaded
   F12 → Network tab → Search "discovery.css"
   Should show Status 200 (not 404)

4. Check Dark Mode
   F12 → Console → type:
   window.matchMedia('(prefers-color-scheme: dark)').matches
   If returns "true", dark mode is active (expected CSS handles this)
```

---

## Test with Real Analysis

### Steps
```
1. Go to: http://192.168.1.6:8081/manuscript/analysis
2. Upload 1-2 chapters
3. Click "Analyze All Rules"
4. Wait for completion (progress indicator shows)
5. Click "Select Rules" button on dashboard
6. Should see Discovery UI with real data (40+ rules)

Expected:
✓ Patterns are VISIBLE (dark text)
✓ Elements list shows 10+ items
✓ Clicking elements updates rules table
✓ Stats update when selecting rules
✓ Can save selection

If fails:
✗ Check console (F12) for errors
✗ Check Network tab for 404/500
✗ Check session_id in URL
```

---

## Success Indicators

### ✅ Everything Works If:
```
1. Patterns are clearly readable (dark text, visible)
2. Stats update when clicking checkboxes
3. Save Selection button works without error
4. Console shows logs, no red errors
5. Test endpoints work smoothly
6. Real analysis data loads properly
```

### ❌ Issues To Report If:
```
1. Patterns still appear as dark boxes
   → Take screenshot, send it
   
2. Save Selection fails with error
   → Note the error message, check console
   
3. Stats don't update
   → Check console for JavaScript errors
   
4. Page loads slowly
   → Check Network tab timing
```

---

## Browser Console Shortcuts

### Open Console
- Windows: F12 then click "Console"
- Mac: Cmd+Option+I then click "Console"

### Check for Errors
- Look for red messages starting with "Error", "Uncaught", "Failed"
- Expand arrows to see full error details
- Copy full error text for reporting

### Check Network
- Click "Network" tab
- Reload page (F5)
- Look for requests with Status 404 or 500
- These indicate server problems

### Check CSS
- Click "Network" tab
- Search for "discovery.css"
- Should show Status 200
- If 404, CSS file not found

---

## Quick Problem Solver

| Problem | Cause | Fix |
|---------|-------|-----|
| Dark boxes for patterns | CSS color issue | Hard refresh (Ctrl+F5) or clear cache |
| Save fails: "Unexpected token" | Server error response | Check Network tab, verify session_id |
| Stats don't update | JavaScript error | Hard refresh, check console |
| Nothing loads | API not responding | Check Network tab, restart server |
| Patterns missing | API returned empty data | Use test endpoint instead |

---

## Expected Timings

| Action | Expected Time | Maximum |
|--------|---------------|---------|
| Page load | < 2 seconds | 5 seconds |
| API request | < 500ms | 2 seconds |
| DOM render | < 300ms | 1 second |
| Stats update | < 50ms | 500ms |
| Save selection | < 1 second | 3 seconds |

---

## What to Report if Issues Persist

When reporting a bug, provide:

1. **Screenshot** of the problem
2. **Browser & OS**: Chrome 120, Windows 11, etc.
3. **URL being tested**: Exact URL from address bar
4. **Console errors**: F12 → Console → copy all red text
5. **Network errors**: F12 → Network → check 4xx/5xx status codes
6. **Steps to reproduce**: Exact steps taken before error
7. **Session/Job ID**: If using real analysis

---

## Success Checklist

After fixes, verify:
- [ ] Hard refresh completed (Ctrl+F5)
- [ ] Test page loads (/discovery/test)
- [ ] Patterns are VISIBLE (readable text)
- [ ] Stats update when clicking checkboxes
- [ ] Save Selection works without error
- [ ] Console (F12) shows no red errors
- [ ] Real analysis can also load Discovery
- [ ] All tests pass without issues

**You're good to go when ALL boxes are checked! ✅**


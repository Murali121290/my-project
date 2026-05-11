# Manuscript Templates: Usage & Enhancement Guide

## Template Location & Usage Map

```
C:\Users\muraliba\PycharmProjects\New folder\PPH\templates\manuscript\
├── analysis_upload.html          → /manuscript/analyze (GET/POST) - Upload chapters
├── technical_upload.html         → /manuscript/technical-edit (GET/POST) - Technical editing start
├── manuscript_dashboard.html     → /manuscript/dashboard/<job_id> (GET) - Main analysis dashboard
├── editor_review.html            → /manuscript/review/<job_id> (GET) - Review & approve fixes
├── dashboard_standalone.html     → /download/<job_id>/sheet.html (GET) - Exportable HTML
├── discovery.html                → /manuscript/discovery (GET/POST) - Rule selection UI
├── rule_selections.html          → /manuscript/rule-selections (GET) - Manage selections
├── base.html                     → Parent template for all pages
└── technical_upload.html         → Alternative upload page
```

---

## Current State Analysis

### ✅ Strengths
- Tailwind CSS framework (responsive, modern)
- DataTables integration (sortable, searchable)
- Chart.js metrics visualization
- Color-coded severity badges (error, warn, info)
- Multi-tab interface (Dashboard, Inconsistencies, Findings)
- Filter controls (by chapter, rule, type)

### ⚠️ Limitations

**Dashboard (`manuscript_dashboard.html`):**
- Static metric cards (no real-time updates)
- Limited chart interactions
- Fixed sidebar layout (not responsive on mobile)
- No export options within dashboard
- Limited comparison between chapters
- No drill-down capability from metrics to findings

**Review Editor (`editor_review.html`):**
- Basic table view (limited visual feedback)
- No side-by-side comparison
- No context preview panel
- No undo/redo functionality
- Limited acceptance/rejection indicators
- No bulk operations feedback
- No progress tracking

**Discovery UI (`discovery.html`):**
- Minimal visual feedback during selection
- No preview of selected rules
- No summary statistics during selection
- Limited customization interface

---

## Enhancement Roadmap

### Phase 1: Dashboard Improvements (High Impact)

#### 1.1 Interactive Metric Cards
```javascript
// Add real-time updates and interactions
- Click metric card → Filter table below
- Hover card → Highlight related rows
- Metric breakdown tooltip → Show top rules
- Mini charts in cards → Trend visualization
```

#### 1.2 Chapter Comparison View
```html
<!-- Add comparison mode -->
- Side-by-side chapter statistics
- Chapter heatmap (rule distribution)
- Chapter-specific rule breakdown
- Export per-chapter summaries
```

#### 1.3 Enhanced Visualization
```javascript
// Replace static charts with interactive ones
- Pie charts → Click slice to filter rules
- Bar charts → Hover for exact values
- Heatmap → Show rule density by chapter
- Timeline → Show finding progression
```

### Phase 2: Review Editor Improvements (Critical)

#### 2.1 Side-by-Side Editor
```html
<!-- Left panel: Original text with context
     Right panel: Preview of changes
     Center: Before/After diff view -->
- Syntax highlighting for matches
- Line-by-line comparison
- Confidence score indicator
- Accept/Reject buttons per row
```

#### 2.2 Context Panel
```html
<!-- Show paragraph context for each finding -->
- Full sentence where match occurs
- Highlight match within context
- Show surrounding sentences for context
- Link to exact page in document
```

#### 2.3 Bulk Operations
```javascript
// Enhanced batch processing
- Accept/Reject all by rule
- Accept/Reject all by chapter
- Accept/Reject all with confidence > X%
- Undo last 10 accepts/rejects
- Progress bar during processing
```

#### 2.4 Visual Feedback
```css
/* Enhanced state indicators */
- Green checkmark: Accepted
- Red X: Rejected
- Yellow flag: Skipped/Uncertain
- Spinner: Processing
- Toast notifications: Operation results
```

### Phase 3: Discovery UI Improvements

#### 3.1 Live Preview
```html
<!-- Show real-time selection stats -->
- Selected rules count
- Total findings for selected rules
- Estimated fixes preview
- Rule impact analysis
```

#### 3.2 Rule Preview
```html
<!-- Hover/click rule → Show examples -->
- Example findings matching rule
- Context snippets
- Common variations
- Impact across chapters
```

#### 3.3 Grouping Assistant
```javascript
// Smart grouping suggestions
- Suggest groups by rule category
- Drag-drop to reorder
- Color-coding for groups
- Save group templates
```

---

## Recommended Code Changes

### Enhancement 1: Dashboard - Add Drill-Down Interaction

**Current**: Static metric cards  
**Goal**: Click metric → Filter/highlight findings

```html
<!-- manuscript_dashboard.html modifications -->

<!-- Replace static metric cards with interactive ones -->
<div class="metric-card cursor-pointer hover:shadow-lg transition"
     id="card-findings"
     onclick="filterDashboardTable('all')">
  <div class="p-6">
    <h3 class="text-sm font-medium text-gray-600">Total Findings</h3>
    <p class="text-3xl font-bold text-blue-900 mt-2">{{ meta.total_findings }}</p>
    <p class="text-xs text-gray-500 mt-2">Click to show all findings</p>
  </div>
</div>

<!-- Add mini chart to card -->
<div class="metric-card cursor-pointer" id="card-spelling" onclick="filterDashboardTable('spelling')">
  <div class="p-6">
    <h3 class="text-sm font-medium text-gray-600">Spelling Variants</h3>
    <div style="position: relative; height: 40px; margin-top: 12px;">
      <canvas id="spelling-mini-chart"></canvas>
    </div>
    <p class="text-xs text-gray-500 mt-2 flex items-center gap-1">
      <span class="inline-block w-2 h-2 bg-blue-600 rounded-full"></span>
      {{ meta.total_findings }} issues
    </p>
  </div>
</div>

<script>
// Initialize mini charts
function initMiniCharts() {
  const ctx = document.getElementById('spelling-mini-chart').getContext('2d');
  new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: ['US', 'UK'],
      datasets: [{
        data: [{{ spelling_summary.us }}, {{ spelling_summary.uk }}],
        backgroundColor: ['#2563EB', '#C55A00'],
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } }
    }
  });
}

// Filter table when metric clicked
function filterDashboardTable(category) {
  // Highlight relevant rows in IA report table
  const rows = document.querySelectorAll('table tbody tr');
  rows.forEach(row => {
    if (category === 'all') {
      row.classList.remove('opacity-50');
    } else {
      row.classList.add('opacity-50');
    }
  });
  // Scroll to table
  document.querySelector('table').scrollIntoView({ behavior: 'smooth' });
}

window.addEventListener('DOMContentLoaded', initMiniCharts);
</script>
```

---

### Enhancement 2: Review Editor - Side-by-Side Comparison

**Current**: Inline table with text input  
**Goal**: Side-by-side visual diff with before/after preview

```html
<!-- editor_review.html modifications -->

<!-- Replace table-based editor with split-panel layout -->
<div class="grid grid-cols-3 gap-4 h-screen bg-white">
  
  <!-- Left Panel: Findings List -->
  <div class="col-span-1 border-r border-slate-200 overflow-y-auto">
    <div class="sticky top-0 bg-slate-50 border-b p-4">
      <h2 class="font-semibold text-gray-800 mb-3">Findings</h2>
      <input type="search" id="findings-search" placeholder="Search..."
             class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm">
    </div>
    
    <div id="findings-list" class="divide-y">
      {% for f in findings %}
      <div class="finding-item p-3 cursor-pointer hover:bg-blue-50 border-l-4 border-transparent"
           data-finding-id="{{ loop.index0 }}"
           onclick="selectFinding(this, {{ loop.index0 }})">
        <div class="flex items-start justify-between">
          <div class="flex-1">
            <p class="text-sm font-medium text-gray-800">{{ f.rule_label }}</p>
            <p class="text-xs text-gray-500">Ch {{ f.chapter_index }} · p{{ f.page }}</p>
          </div>
          <input type="checkbox" class="finding-cb rounded accent-blue-600"
                 onchange="event.stopPropagation(); toggleFinding(this, {{ loop.index0 }})">
        </div>
      </div>
      {% endfor %}
    </div>
  </div>
  
  <!-- Middle Panel: Context & Preview -->
  <div class="col-span-1 border-r border-slate-200 overflow-y-auto p-4">
    <div id="context-panel" class="space-y-4">
      <div>
        <h3 class="text-sm font-semibold text-gray-700 mb-2">Context</h3>
        <div id="context-text" class="bg-slate-50 p-3 rounded border border-slate-200 text-sm text-gray-700 font-mono">
          <!-- Context will be shown here -->
        </div>
      </div>
      
      <div>
        <h3 class="text-sm font-semibold text-gray-700 mb-2">Current Match</h3>
        <div class="bg-yellow-50 border border-yellow-200 rounded p-3">
          <span id="current-match" class="text-sm font-semibold text-yellow-900"></span>
        </div>
      </div>
      
      <div>
        <label class="text-sm font-medium text-gray-700 block mb-2">Replacement</label>
        <input type="text" id="replacement-input" 
               class="w-full px-3 py-2 border border-slate-300 rounded-lg text-sm"
               placeholder="Enter replacement...">
      </div>
    </div>
  </div>
  
  <!-- Right Panel: Live Preview -->
  <div class="col-span-1 bg-slate-50 overflow-y-auto p-4">
    <h3 class="text-sm font-semibold text-gray-700 mb-3">Preview</h3>
    
    <div class="space-y-3">
      <div>
        <p class="text-xs font-medium text-gray-500 uppercase tracking-wide mb-1">Before</p>
        <div id="preview-before" class="bg-white p-3 rounded border border-slate-200 text-sm">
          <!-- Before text -->
        </div>
      </div>
      
      <div class="flex justify-center">
        <span class="text-gray-400">↓</span>
      </div>
      
      <div>
        <p class="text-xs font-medium text-gray-500 uppercase tracking-wide mb-1">After</p>
        <div id="preview-after" class="bg-green-50 p-3 rounded border border-green-200 text-sm">
          <!-- After text -->
        </div>
      </div>
    </div>
    
    <div class="mt-4 flex gap-2">
      <button onclick="acceptFinding()" 
              class="flex-1 px-4 py-2 bg-green-600 hover:bg-green-700 text-white rounded-lg text-sm font-medium">
        ✓ Accept
      </button>
      <button onclick="rejectFinding()" 
              class="flex-1 px-4 py-2 bg-red-600 hover:bg-red-700 text-white rounded-lg text-sm font-medium">
        ✗ Reject
      </button>
    </div>
  </div>
</div>

<script>
// State management
let currentFinding = null;
let findings = {{ findings | tojson }};
let decisions = {}; // Track user decisions

function selectFinding(element, idx) {
  // Update selected state
  document.querySelectorAll('.finding-item').forEach(el => {
    el.classList.remove('border-l-blue-600', 'bg-blue-50');
  });
  element.classList.add('border-l-blue-600', 'bg-blue-50');
  
  currentFinding = findings[idx];
  
  // Update context panel
  document.getElementById('context-text').textContent = currentFinding.context;
  document.getElementById('current-match').textContent = currentFinding.surface;
  document.getElementById('replacement-input').value = currentFinding.replacement || '';
  
  // Update preview
  const before = currentFinding.context.replace(
    new RegExp(currentFinding.surface, 'g'),
    `<mark class="match-hl">${currentFinding.surface}</mark>`
  );
  document.getElementById('preview-before').innerHTML = before;
  
  const after = currentFinding.context.replace(
    currentFinding.surface,
    `<mark class="bg-green-200">${currentFinding.replacement}</mark>`
  );
  document.getElementById('preview-after').innerHTML = after;
}

function acceptFinding() {
  if (!currentFinding) return;
  decisions[currentFinding.index] = { decision: 'accept', replacement: document.getElementById('replacement-input').value };
  
  // Visual feedback
  const item = document.querySelector(`[data-finding-id="${currentFinding.index}"]`);
  item.classList.add('bg-green-50');
  item.querySelector('.finding-cb').checked = true;
  
  // Move to next
  const nextIdx = (currentFinding.index + 1) % findings.length;
  selectFinding(document.querySelector(`[data-finding-id="${nextIdx}"]`), nextIdx);
}

function rejectFinding() {
  if (!currentFinding) return;
  decisions[currentFinding.index] = { decision: 'reject' };
  
  // Visual feedback
  const item = document.querySelector(`[data-finding-id="${currentFinding.index}"]`);
  item.classList.add('bg-red-50');
  item.querySelector('.finding-cb').checked = false;
  
  // Move to next
  const nextIdx = (currentFinding.index + 1) % findings.length;
  selectFinding(document.querySelector(`[data-finding-id="${nextIdx}"]`), nextIdx);
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
  if (findings.length > 0) {
    selectFinding(document.querySelector('[data-finding-id="0"]'), 0);
  }
});
</script>
```

---

### Enhancement 3: Discovery UI - Live Preview & Statistics

**Current**: Basic checkboxes  
**Goal**: Real-time statistics and rule preview

```html
<!-- discovery.html modifications -->

<!-- Add statistics panel -->
<div class="grid grid-cols-4 gap-4 mb-6">
  <div class="bg-white p-4 rounded-lg border border-slate-200">
    <p class="text-sm text-gray-600">Rules Selected</p>
    <p class="text-2xl font-bold text-blue-900" id="stats-selected">0</p>
    <p class="text-xs text-gray-500 mt-1" id="stats-of-total">of 0</p>
  </div>
  
  <div class="bg-white p-4 rounded-lg border border-slate-200">
    <p class="text-sm text-gray-600">Total Findings</p>
    <p class="text-2xl font-bold text-orange-600" id="stats-findings">0</p>
    <p class="text-xs text-gray-500 mt-1">Will be included</p>
  </div>
  
  <div class="bg-white p-4 rounded-lg border border-slate-200">
    <p class="text-sm text-gray-600">Auto-Fix Rules</p>
    <p class="text-2xl font-bold text-green-600" id="stats-autofix">0</p>
    <p class="text-xs text-gray-500 mt-1">Will be applied</p>
  </div>
  
  <div class="bg-white p-4 rounded-lg border border-slate-200">
    <p class="text-sm text-gray-600">Highlight-Only</p>
    <p class="text-2xl font-bold text-purple-600" id="stats-highlight">0</p>
    <p class="text-xs text-gray-500 mt-1">Visual markup only</p>
  </div>
</div>

<!-- Rule preview panel -->
<div class="grid grid-cols-3 gap-4">
  <!-- ... existing panels ... -->
  
  <!-- Add rule preview on right -->
  <div class="bg-white rounded-lg border border-slate-200">
    <div class="p-4 border-b border-slate-200">
      <h3 class="font-semibold text-gray-800">Rule Examples</h3>
    </div>
    <div id="rule-preview" class="p-4 space-y-3 max-h-96 overflow-y-auto">
      <p class="text-sm text-gray-500">Hover over a rule to see examples</p>
    </div>
  </div>
</div>

<script>
class DiscoveryUIEnhanced extends DiscoveryUI {
  toggleRule(ruleKey, checked) {
    super.toggleRule(ruleKey, checked);
    this.updateStatistics();
  }

  updateStatistics() {
    const selected = this.selectedRules.size;
    const total = this.allRules.length;
    
    // Count findings and categorize rules
    let totalFindings = 0;
    let autoFixCount = 0;
    let highlightCount = 0;
    
    this.selectedRules.forEach(ruleKey => {
      const [element, pattern] = ruleKey.split(':');
      const rule = this.allRules.find(r => r.element === element && r.pattern === pattern);
      if (rule) {
        totalFindings += rule.detected_count;
        
        // Categorize
        if (['Figure', 'Table', 'Box'].includes(element)) {
          highlightCount += 1;
        } else {
          autoFixCount += 1;
        }
      }
    });
    
    // Update cards
    document.getElementById('stats-selected').textContent = selected;
    document.getElementById('stats-of-total').textContent = `of ${total}`;
    document.getElementById('stats-findings').textContent = totalFindings;
    document.getElementById('stats-autofix').textContent = autoFixCount;
    document.getElementById('stats-highlight').textContent = highlightCount;
  }

  setupRulePreview() {
    document.querySelectorAll('tr').forEach(row => {
      row.addEventListener('mouseenter', () => this.showRulePreview(row));
    });
  }

  showRulePreview(row) {
    const pattern = row.querySelector('.rule-pattern')?.textContent || '';
    const count = row.querySelector('td:last-child')?.textContent || '0';
    
    const preview = document.getElementById('rule-preview');
    preview.innerHTML = `
      <div class="space-y-2">
        <p class="font-medium text-gray-800 text-sm">${pattern}</p>
        <p class="text-xs text-gray-600">${count} findings across manuscript</p>
        <div class="bg-blue-50 p-2 rounded text-xs text-gray-700">
          Example: <em>Typical occurrence of this pattern in text</em>
        </div>
      </div>
    `;
  }
}

// Enhanced initialization
let discovery = new DiscoveryUIEnhanced();
</script>
```

---

## Implementation Priority

### Phase 1 (Week 1): Critical UX Improvements
- ✅ Review editor side-by-side layout
- ✅ Context preview panel
- ✅ Accept/Reject visual feedback
- Estimated effort: 6-8 hours

### Phase 2 (Week 2): Dashboard Enhancements
- ✅ Interactive metric cards
- ✅ Chapter comparison view
- ✅ Enhanced data visualization
- Estimated effort: 8-10 hours

### Phase 3 (Week 3): Discovery Refinements
- ✅ Live statistics panel
- ✅ Rule preview on hover
- ✅ Grouping assistant
- Estimated effort: 4-6 hours

---

## Testing Checklist

- [ ] Editor review loads all findings correctly
- [ ] Side-by-side comparison syncs properly
- [ ] Accept/Reject buttons update UI
- [ ] Preview updates in real-time
- [ ] Dashboard metrics are interactive
- [ ] Charts render without errors
- [ ] Discovery statistics update correctly
- [ ] Mobile responsiveness maintained
- [ ] Keyboard navigation works
- [ ] Accessibility standards (WCAG) met

---

## Browser Compatibility

- Chrome/Edge: 95+
- Firefox: 93+
- Safari: 15+
- Mobile: iOS Safari 15+, Chrome Android 95+

---

## Dependencies

Already included:
- Tailwind CSS (Utility-first styling)
- DataTables (Sortable tables)
- Chart.js (Data visualization)
- Lucide Icons (Icon set)
- jQuery (DOM manipulation)

Optional additions:
- Diff.js (Side-by-side diff display)
- SortableJS (Drag-drop reordering)
- Toast notifications (Feedback)

---

**Last Updated**: 2026-05-09  
**Status**: Ready for implementation  
**Estimated Total Effort**: 18-24 hours  
**Priority**: High (User-facing improvements)

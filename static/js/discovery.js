/**
 * Discovery UI: Rule Selection & Custom Grouping
 * Handles client-side logic for selecting IA rows and organizing custom groupings.
 */

class DiscoveryUI {
  constructor() {
    this.allRules = [];
    this.selectedRules = new Set();
    this.currentElement = null;
    this.customGrouping = {};
    this.init();
  }

  async init() {
    await this.loadRules();
    this.setupEventListeners();
  }

  async loadRules() {
    const sessionId = new URLSearchParams(window.location.search).get('session_id');
    if (!sessionId) {
      this.showError('No session ID provided');
      return;
    }

    try {
      const response = await fetch(`/manuscript/discovery/${sessionId}/ia-rows`);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      const data = await response.json();

      console.log('Discovery API Response:', data);
      console.log('Total IA rows:', data.ia_rows.length);
      if (data.ia_rows.length > 0) {
        console.log('Sample row:', data.ia_rows[0]);
      }

      this.allRules = data.ia_rows;
      const elements = data.elements;

      console.log('Elements loaded:', elements);
      this.populateElements(elements);
      this.updateTotalCount(data.ia_rows.length);

      // Select first element by default
      if (elements.length > 0) {
        this.selectElement(elements[0]);
      }
    } catch (error) {
      console.error('Error loading rules:', error);
      this.showError('Failed to load rules: ' + error.message);
    }
  }

  populateElements(elements) {
    const elementList = document.getElementById('elementList');
    elementList.innerHTML = '';

    elements.forEach(elem => {
      const li = document.createElement('li');
      li.className = 'element-item';
      li.innerHTML = `<span class="element-item-text">${this.escapeHtml(elem)}</span>`;
      li.onclick = () => this.selectElement(elem);
      elementList.appendChild(li);
    });
  }

  selectElement(element) {
    this.currentElement = element;

    // Update UI
    document.querySelectorAll('.element-item').forEach(item => {
      const text = item.querySelector('.element-item-text').textContent;
      if (text === element) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    // Display rules for this element
    this.displayElementRules(element);
    this.updateGroupingEditor();
  }

  displayElementRules(element) {
    const elementRules = this.allRules.filter(r => r.element === element);
    const rulesBody = document.getElementById('rulesBody');
    rulesBody.innerHTML = '';

    console.log(`Displaying rules for element: ${element}`);
    console.log(`Found ${elementRules.length} rules for this element`);

    elementRules.forEach((rule, index) => {
      const ruleKey = this.getRuleKey(rule);
      const tr = document.createElement('tr');

      // Ensure pattern is a string and not empty
      const pattern = rule.pattern || '[No pattern]';
      const patternHtml = this.escapeHtml(pattern);
      const count = rule.detected_count || 0;

      tr.innerHTML = `
        <td>
          <input type="checkbox" class="rule-checkbox"
            ${this.selectedRules.has(ruleKey) ? 'checked' : ''}
            onchange="discovery.toggleRule('${ruleKey}', this.checked)">
        </td>
        <td><div class="rule-pattern">${patternHtml}</div></td>
        <td>${count}</td>
      `;
      rulesBody.appendChild(tr);
      console.log(`Row ${index}: pattern="${pattern}", count=${count}, element="${rule.element}"`);
    });

    this.updateSelectAllCheckbox();
  }

  getRuleKey(rule) {
    return `${rule.element}:${rule.pattern}`;
  }

  toggleRule(ruleKey, checked) {
    if (checked) {
      this.selectedRules.add(ruleKey);
    } else {
      this.selectedRules.delete(ruleKey);
    }
    this.updateSummary();
    this.updateStats();
    this.updateSelectAllCheckbox();
  }

  updateSelectAllCheckbox() {
    const elementRules = this.allRules.filter(r => r.element === this.currentElement);
    const allChecked = elementRules.every(r => this.selectedRules.has(this.getRuleKey(r)));
    const someChecked = elementRules.some(r => this.selectedRules.has(this.getRuleKey(r)));

    const checkbox = document.getElementById('selectAllCheckbox');
    checkbox.checked = allChecked;
    checkbox.indeterminate = someChecked && !allChecked;
  }

  setupEventListeners() {
    const selectAllCheckbox = document.getElementById('selectAllCheckbox');
    selectAllCheckbox.addEventListener('change', () => this.toggleSelectAll());
  }

  toggleSelectAll() {
    const elementRules = this.allRules.filter(r => r.element === this.currentElement);
    const checkbox = document.getElementById('selectAllCheckbox');

    elementRules.forEach(rule => {
      const ruleKey = this.getRuleKey(rule);
      if (checkbox.checked) {
        this.selectedRules.add(ruleKey);
      } else {
        this.selectedRules.delete(ruleKey);
      }
    });

    this.displayElementRules(this.currentElement);
    this.updateSummary();
    this.updateStats();
  }

  updateSummary() {
    const summary = document.querySelector('.rule-count');
    if (summary) {
      summary.textContent = this.selectedRules.size;
    }
  }

  updateStats() {
    // Calculate statistics for selected rules
    let totalFindings = 0;
    let autoFixCount = 0;
    let highlightOnlyCount = 0;

    this.selectedRules.forEach(ruleKey => {
      const [element, pattern] = ruleKey.split(':');
      const rule = this.allRules.find(r => r.element === element && r.pattern === pattern);
      if (rule) {
        totalFindings += (rule.detected_count || 0);

        // Classify as auto-fix or highlight-only
        if (['Figure', 'Table', 'Box'].includes(element)) {
          highlightOnlyCount++;
        } else {
          autoFixCount++;
        }
      }
    });

    // Update stats panel
    document.getElementById('statsSelectedCount').textContent = this.selectedRules.size;
    document.getElementById('statsTotalFindings').textContent = totalFindings.toLocaleString();
    document.getElementById('statsAutoFixRules').textContent = autoFixCount;
    document.getElementById('statsHighlightRules').textContent = highlightOnlyCount;
  }

  updateTotalCount(total) {
    const totalSpan = document.getElementById('totalRules');
    if (totalSpan) {
      totalSpan.textContent = total;
    }
  }

  selectAllGlobal() {
    this.allRules.forEach(rule => {
      this.selectedRules.add(this.getRuleKey(rule));
    });
    this.displayElementRules(this.currentElement);
    this.updateSummary();
    this.updateStats();
  }

  clearSelection() {
    this.selectedRules.clear();
    this.displayElementRules(this.currentElement);
    this.updateSummary();
    this.updateStats();
  }

  updateGroupingEditor() {
    const editor = document.getElementById('groupingEditor');
    if (!editor) return;

    const elementRules = this.allRules.filter(r => r.element === this.currentElement);
    let html = '<div style="font-size: 0.8125rem; color: #718096;">';

    if (elementRules.length === 0) {
      html += 'No rules for this element';
    } else {
      html += `<strong>${this.escapeHtml(this.currentElement)}</strong><br>`;
      html += `${elementRules.length} total rules`;
    }

    html += '</div>';
    editor.innerHTML = html;
  }

  async saveSelection() {
    // Convert to IA row format
    const selectedIARows = [];
    this.selectedRules.forEach(ruleKey => {
      const [element, pattern] = ruleKey.split(':');
      const rule = this.allRules.find(r => r.element === element && r.pattern === pattern);
      if (rule) {
        selectedIARows.push({
          element: rule.element,
          subtype: rule.subtype,
          pattern: rule.pattern,
          example: rule.example,
        });
      }
    });

    if (selectedIARows.length === 0) {
      alert('Please select at least one rule');
      return;
    }

    // Get client_name and project_name from page data or session storage
    const clientName = window.clientData?.client_name ||
                       sessionStorage.getItem('client_name') ||
                       'Default';
    const projectName = window.clientData?.project_name ||
                        sessionStorage.getItem('project_name') ||
                        'Project';

    // Auto-generate selection name from client and project
    const timestamp = new Date().toISOString().split('T')[0].replace(/-/g, '');
    const selectionName = `${projectName}_${clientName}_Rules_${timestamp}`;

    const sessionId = new URLSearchParams(window.location.search).get('session_id');
    console.log(`Saving selection: ${selectionName}, with ${selectedIARows.length} rules`);

    try {
      const response = await fetch(`/manuscript/discovery/${sessionId}/create-selection`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          selection_name: selectionName,
          description: `Rules selected for ${projectName} - ${clientName}`,
          selected_ia_rows: selectedIARows,
          custom_grouping: this.customGrouping,
          project_name: projectName,
          client_name: clientName,
        }),
      });

      console.log(`Response status: ${response.status}`);

      // Try to parse as JSON
      let data;
      try {
        data = await response.json();
      } catch (parseError) {
        const text = await response.text();
        console.error('Response was not JSON:', text);
        alert(`Server error: ${response.status} - ${text.substring(0, 100)}`);
        return;
      }

      if (response.ok) {
        console.log(`Selection saved with ID: ${data.selection_id}`);
        alert(`✓ Rules saved for ${projectName} - ${clientName}\n(Selection ID: ${data.selection_id})\n\nRules selected: ${selectedIARows.length}`);
        
        // After successfully saving a selection, you likely want to be able to navigate to rule selections
        // So we might prompt the user
        if (confirm('Selection saved successfully. Do you want to go to the Rule Selections management page?')) {
            window.location.href = '/manuscript/rule-selections';
        }
      } else {
        console.error('Server error:', data);
        alert('Error: ' + (data.error || 'Unknown error'));
      }
    } catch (error) {
      console.error('Error saving selection:', error);
      alert('Error saving selection: ' + error.message);
    }
  }

  showError(message) {
    const container = document.getElementById('rulesBody');
    if (container) {
      container.innerHTML = `<tr><td colspan="3" style="text-align: center; padding: 2rem; color: #dc2626;">${this.escapeHtml(message)}</td></tr>`;
    }
  }

  escapeHtml(text) {
    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#039;'
    };
    return text.replace(/[&<>"']/g, m => map[m]);
  }
}

// Global instance
let discovery = null;

// Initialize on page load
window.addEventListener('load', () => {
  discovery = new DiscoveryUI();
});

// Global helper functions for inline HTML
function selectElement(element) {
  if (discovery) discovery.selectElement(element);
}

function toggleRule(ruleKey, checked) {
  if (discovery) discovery.toggleRule(ruleKey, checked);
}

function toggleSelectAll() {
  if (discovery) discovery.toggleSelectAll();
}

function selectAllGlobal() {
  if (discovery) discovery.selectAllGlobal();
}

function clearSelection() {
  if (discovery) discovery.clearSelection();
}

function saveSelection() {
  if (discovery) discovery.saveSelection();
}

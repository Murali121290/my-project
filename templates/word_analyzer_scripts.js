/* ── Chapter switcher ── */
function showChapter(id, btn) {
    document.querySelectorAll('.chapter-section').forEach(s => s.style.display = 'none');
    document.getElementById(id).style.display = 'block';
    document.querySelectorAll('.ch-nav-btn').forEach(b => {
        b.style.background = 'white'; b.style.color = 'var(--primary,#4361ee)';
    });
    btn.style.background = 'var(--primary,#4361ee)';
    btn.style.color = 'white';
}

/* ── Scoped tab switcher (called by tab buttons: showTab(prefix, tab, btn)) ── */
function showTab(prefix, tab, btn) {
    // 1-arg call from summary inline links: showTab('citations') — detect by arg count
    if (tab === undefined) {
        // find the active chapter section and delegate
        const activeSection = document.querySelector('.chapter-section:not([style*="none"])');
        if (activeSection) {
            const p = activeSection.id;
            activeSection.querySelectorAll('.tab-content').forEach(t => t.style.display = 'none');
            const target = document.getElementById(p + '_' + prefix);
            if (target) target.style.display = 'block';
            activeSection.querySelectorAll('.nav-tab').forEach(b => { b.style.background=''; b.style.color=''; });
            const matchBtn = Array.from(activeSection.querySelectorAll('.nav-tab'))
                .find(b => b.getAttribute('onclick') && b.getAttribute('onclick').includes("'" + prefix + "'"));
            if (matchBtn) { matchBtn.style.background='var(--primary,#4361ee)'; matchBtn.style.color='white'; }
        }
        return;
    }
    const section = document.getElementById(prefix);
    if (!section) return;
    section.querySelectorAll('.tab-content').forEach(t => t.style.display = 'none');
    const tgt = document.getElementById(prefix + '_' + tab);
    if (tgt) tgt.style.display = 'block';
    section.querySelectorAll('.nav-tab').forEach(b => { b.style.background=''; b.style.color=''; });
    if (btn) { btn.style.background='var(--primary,#4361ee)'; btn.style.color='white'; }
}

/* ── showTabFromRow: called by summary table row clicks ── */
function showTabFromRow(tabId, row) {
    // find the chapter-section ancestor of the clicked row
    let section = row;
    while (section && !section.classList.contains('chapter-section')) section = section.parentElement;
    if (!section) return;
    const prefix = section.id;
    const isAlreadyOpen = document.getElementById(prefix + '_' + tabId).style.display !== 'none'
        && document.getElementById(prefix + '_' + tabId).style.display !== '';
    // toggle: if already visible and tab-section visible, hide it
    const tabContents = section.querySelectorAll('.tab-content');
    const anyVisible = Array.from(tabContents).some(t => t.style.display === 'block');
    if (anyVisible && isAlreadyOpen) {
        tabContents.forEach(t => t.style.display = 'none');
        section.querySelectorAll('.nav-tab').forEach(b => { b.style.background=''; b.style.color=''; });
        if (row) row.classList.remove('summary-row-active');
    } else {
        tabContents.forEach(t => t.style.display = 'none');
        const tgt = document.getElementById(prefix + '_' + tabId);
        if (tgt) tgt.style.display = 'block';
        section.querySelectorAll('.nav-tab').forEach(b => { b.style.background=''; b.style.color=''; });
        const matchBtn = Array.from(section.querySelectorAll('.nav-tab'))
            .find(b => b.getAttribute('onclick') && b.getAttribute('onclick').includes("'" + tabId + "'"));
        if (matchBtn) { matchBtn.style.background='var(--primary,#4361ee)'; matchBtn.style.color='white'; }
        if (row) { row.classList.add('summary-row-active'); tgt.scrollIntoView({behavior:'smooth'}); }
    }
}

/* ── Back-to-top per chapter ── */
window.addEventListener('scroll', function() {
    var btn = document.getElementById('back-to-top');
    if (btn) btn.style.display = window.scrollY > 300 ? 'block' : 'none';
});

/* ── DataTables init ── */
document.addEventListener('DOMContentLoaded', function() {
    // activate first chapter nav button
    const firstBtn = document.querySelector('.ch-nav-btn');
    if (firstBtn) { firstBtn.style.background='var(--primary,#4361ee)'; firstBtn.style.color='white'; }
    // init datatables
    if (typeof $ !== 'undefined' && $.fn.DataTable) {
        $('table.analysis-table, table.element-table, table:not(#globalSummaryTable)').each(function() {
            if ($(this).find('td[colspan], td[rowspan]').length > 0) return;
            try { $(this).DataTable({ pageLength:10, autoWidth:false, ordering:true, responsive:true, columnDefs:[{targets:"_all",defaultContent:""}] }); }
            catch(e) { console.warn('DataTable init failed', e); }
        });

        // Init Global Summary Table explicitly
        if ($('#globalSummaryTable').length > 0) {
            $('#globalSummaryTable').DataTable({
                pageLength: 25,
                autoWidth: false,
                ordering: true,
                dom: 'Bfrtip',
                buttons: [
                    'copyHtml5',
                    'excelHtml5',
                    'csvHtml5'
                ]
            });
        }
    }

    // Animate progress bars
    document.querySelectorAll('[data-w]').forEach(function(bar){
        var w = bar.getAttribute('data-w');
        setTimeout(function(){ bar.style.width = w + '%'; }, 150);
    });
});

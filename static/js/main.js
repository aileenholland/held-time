// ── Table sorting ─────────────────────────────────────────────────────────
const sortState = {};

function sortTable(tableId, colIndex) {
  const table = document.getElementById(tableId);
  if (!table) return;

  const key = `${tableId}-${colIndex}`;
  sortState[key] = sortState[key] === 'asc' ? 'desc' : 'asc';
  const dir = sortState[key];

  const tbody = table.querySelector('tbody');
  const rows = Array.from(tbody.querySelectorAll('tr'));

  rows.sort((a, b) => {
    const aText = a.cells[colIndex]?.innerText.trim() ?? '';
    const bText = b.cells[colIndex]?.innerText.trim() ?? '';
    const aNum = parseFloat(aText.replace(/[$,()%]/g, '').replace(/^\((.+)\)$/, '-$1'));
    const bNum = parseFloat(bText.replace(/[$,()%]/g, '').replace(/^\((.+)\)$/, '-$1'));

    if (!isNaN(aNum) && !isNaN(bNum)) {
      return dir === 'asc' ? aNum - bNum : bNum - aNum;
    }
    return dir === 'asc'
      ? aText.localeCompare(bText)
      : bText.localeCompare(aText);
  });

  rows.forEach(r => tbody.appendChild(r));

  // Update sort indicators
  table.querySelectorAll('th').forEach((th, i) => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (i === colIndex) th.classList.add(dir === 'asc' ? 'sort-asc' : 'sort-desc');
  });
}

// ── Editable cells (COs + Notes) ──────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  document.querySelectorAll('.editable-cell').forEach(input => {
    const original = input.value;

    input.addEventListener('blur', () => {
      const value   = input.value.trim();
      const project = input.dataset.project;
      const field   = input.dataset.field;

      // Skip save if nothing changed
      if (value === original && value === input.dataset.saved) return;

      // Strip $ and commas before saving numeric fields
      const saveVal = field === 'cos' ? value.replace(/[$,]/g, '') : value;

      fetch('/api/override', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ project_number: project, field, value: saveVal }),
      })
        .then(r => r.json())
        .then(data => {
          if (data.ok) {
            input.dataset.saved = value;
            input.classList.add('saved-flash');
            setTimeout(() => input.classList.remove('saved-flash'), 800);
          }
        })
        .catch(() => {});
    });

    // Allow Tab/Enter to move to next editable cell
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
    });
  });
});

// ── Refresh button spin animation ─────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  const btn  = document.getElementById('refresh-btn');
  const icon = document.getElementById('refresh-icon');
  if (btn && icon) {
    btn.addEventListener('click', () => {
      icon.style.transition = 'transform 0.8s linear';
      icon.style.transform  = 'rotate(360deg)';
      btn.style.opacity = '0.7';
      btn.style.pointerEvents = 'none';
      // reset after animation so the page can navigate
      setTimeout(() => {
        icon.style.transform = 'rotate(0deg)';
      }, 850);
    });
  }
});

// ── Collapsible sections ───────────────────────────────────────────────────
function toggleSection(sectionId, btn) {
  const el = document.getElementById(sectionId);
  const chevron = document.getElementById(sectionId.replace('-section', '-chevron'));
  if (!el) return;

  const hidden = el.classList.toggle('hidden');
  if (chevron) chevron.style.transform = hidden ? 'rotate(-90deg)' : 'rotate(0deg)';
}

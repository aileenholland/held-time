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

// ── Collapsible sections ───────────────────────────────────────────────────
function toggleSection(sectionId, btn) {
  const el = document.getElementById(sectionId);
  const chevron = document.getElementById(sectionId.replace('-section', '-chevron'));
  if (!el) return;

  const hidden = el.classList.toggle('hidden');
  if (chevron) chevron.style.transform = hidden ? 'rotate(-90deg)' : 'rotate(0deg)';
}

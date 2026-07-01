'use strict';

/* ---------------------------------------------------------------------
 * Estado
 * ------------------------------------------------------------------- */

const SOURCES = [
  { key: 'lab', label: 'Laboratorio' },
  { key: 'jde', label: 'JD Edwards' },
  { key: 'prov', label: 'Proveedor' },
];

// state.files[key] = { fileName, headerRowIndex, headers: [...], rows: [[...]], selectedCol: index, rowCount }
const state = {
  files: { lab: null, jde: null, prov: null },
  reconciled: null, // { rows: [...], stats: {...}, duplicates: {...} }
};

// Recuerda la última columna elegida por tipo de archivo durante la sesión (memoria en RAM, no en disco)
const lastColumnChoice = { lab: null, jde: null, prov: null };

const GUIA_KEYWORDS = ['guia', 'guía', 'nguia', 'n° guia', 'nro guia', 'numero guia', 'número guía', 'folio', 'documento', 'doc', 'nro doc', 'n° doc'];

/* ---------------------------------------------------------------------
 * Utilidades de normalización
 * ------------------------------------------------------------------- */

function normalizeGuia(raw) {
  if (raw === null || raw === undefined) return null;
  let s;
  if (typeof raw === 'number') {
    if (!isFinite(raw)) return null;
    s = String(Math.trunc(raw));
  } else {
    s = String(raw).trim();
  }
  s = s.replace(/\s+/g, '');
  if (s === '') return null;
  s = s.replace(/\.0+$/, '');
  if (!/^\d+$/.test(s)) return null;
  const stripped = s.replace(/^0+(?=\d)/, '');
  return stripped;
}

function normalizeHeaderText(h) {
  return String(h ?? '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

/* ---------------------------------------------------------------------
 * Lectura de Excel + detección de fila de encabezado
 * ------------------------------------------------------------------- */

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array', cellDates: false });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

function sheetToMatrix(wb) {
  const sheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
}

// Heurística: entre las primeras 15 filas, la fila de encabezado es la que tiene más celdas no vacías
// (y al menos 2), priorizando la más temprana en caso de empate.
function detectHeaderRow(matrix) {
  const scanLimit = Math.min(matrix.length, 15);
  let bestIdx = 0;
  let bestScore = -1;
  for (let i = 0; i < scanLimit; i++) {
    const row = matrix[i] || [];
    const nonEmpty = row.filter((c) => c !== null && c !== undefined && String(c).trim() !== '').length;
    if (nonEmpty > bestScore) {
      bestScore = nonEmpty;
      bestIdx = i;
    }
  }
  return bestIdx;
}

function buildFileState(matrix, fileName) {
  const headerRowIndex = detectHeaderRow(matrix);
  const rawHeaders = matrix[headerRowIndex] || [];
  const headers = rawHeaders.map((h, i) => {
    const t = h !== null && h !== undefined ? String(h).trim() : '';
    return t !== '' ? t : `Columna ${i + 1}`;
  });
  const rows = matrix.slice(headerRowIndex + 1).filter((r) => r && r.some((c) => c !== null && c !== undefined && String(c).trim() !== ''));
  return { fileName, headerRowIndex, headers, rows, selectedCol: null, rowCount: rows.length };
}

function guessGuiaColumn(headers) {
  for (let i = 0; i < headers.length; i++) {
    const h = normalizeHeaderText(headers[i]);
    if (GUIA_KEYWORDS.some((kw) => h.includes(kw))) return i;
  }
  return 0;
}

/* ---------------------------------------------------------------------
 * DOM refs
 * ------------------------------------------------------------------- */

const mappingSection = document.getElementById('mapping-section');
const mappingGrid = document.getElementById('mapping-grid');
const btnProcess = document.getElementById('btn-process');
const processHint = document.getElementById('process-hint');

const resultsSection = document.getElementById('results-section');
const statsGrid = document.getElementById('stats-grid');
const warningsBox = document.getElementById('warnings-box');
const searchInput = document.getElementById('search-input');
const statusFilter = document.getElementById('status-filter');
const ledgerBody = document.getElementById('ledger-body');
const ledgerEmpty = document.getElementById('ledger-empty');
const btnExport = document.getElementById('btn-export');

/* ---------------------------------------------------------------------
 * Carga de archivos por tray
 * ------------------------------------------------------------------- */

SOURCES.forEach(({ key }) => {
  const tray = document.getElementById(`tray-${key}`);
  const drop = tray.querySelector('.tray-drop');
  const input = tray.querySelector('.tray-input');
  const hint = tray.querySelector('.tray-hint');
  const fileBox = tray.querySelector('.tray-file');
  const filenameEl = tray.querySelector('.tray-filename');
  const filerowsEl = tray.querySelector('.tray-filerows');
  const replaceBtn = tray.querySelector('.tray-replace');
  const errorEl = tray.querySelector('.tray-error');

  function openPicker(e) {
    if (e && replaceBtn.contains(e.target)) return;
    input.click();
  }

  drop.addEventListener('click', openPicker);
  drop.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); }
  });
  replaceBtn.addEventListener('click', (e) => { e.stopPropagation(); input.click(); });

  drop.addEventListener('dragover', (e) => { e.preventDefault(); drop.classList.add('dragover'); });
  drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
  drop.addEventListener('drop', (e) => {
    e.preventDefault();
    drop.classList.remove('dragover');
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      handleFile(key, e.dataTransfer.files[0]);
    }
  });

  input.addEventListener('change', () => {
    if (input.files && input.files[0]) handleFile(key, input.files[0]);
    input.value = '';
  });

  async function handleFile(sourceKey, file) {
    errorEl.hidden = true;
    const name = file.name.toLowerCase();
    if (!name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      errorEl.textContent = 'Formato no soportado. Sube un archivo .xlsx o .xls.';
      errorEl.hidden = false;
      return;
    }
    try {
      const wb = await readWorkbook(file);
      const matrix = sheetToMatrix(wb);
      if (!matrix || matrix.length === 0) {
        errorEl.textContent = 'El archivo está vacío o no se pudo leer ninguna fila.';
        errorEl.hidden = false;
        return;
      }
      const fs = buildFileState(matrix, file.name);
      if (fs.headers.length === 0) {
        errorEl.textContent = 'No se detectaron columnas en el archivo.';
        errorEl.hidden = false;
        return;
      }

      let colIndex = null;
      if (lastColumnChoice[sourceKey] !== null) {
        const idx = fs.headers.indexOf(lastColumnChoice[sourceKey]);
        if (idx !== -1) colIndex = idx;
      }
      if (colIndex === null) colIndex = guessGuiaColumn(fs.headers);
      fs.selectedCol = colIndex;

      state.files[sourceKey] = fs;

      hint.hidden = true;
      fileBox.hidden = false;
      filenameEl.textContent = file.name;
      filerowsEl.textContent = `${fs.rowCount} fila${fs.rowCount === 1 ? '' : 's'} detectada${fs.rowCount === 1 ? '' : 's'}`;

      renderMapping();
    } catch (err) {
      console.error(err);
      errorEl.textContent = 'No se pudo leer el archivo. Verifica que sea un Excel válido.';
      errorEl.hidden = false;
    }
  }
});

/* ---------------------------------------------------------------------
 * Sección de mapeo de columnas
 * ------------------------------------------------------------------- */

function renderMapping() {
  const anyLoaded = SOURCES.some(({ key }) => state.files[key]);
  mappingSection.hidden = !anyLoaded;
  if (!anyLoaded) return;

  mappingGrid.innerHTML = '';

  SOURCES.forEach(({ key, label }) => {
    const fs = state.files[key];
    if (!fs) return;

    const card = document.createElement('div');
    card.className = 'map-card';

    const head = document.createElement('div');
    head.className = 'map-card-head';
    const title = document.createElement('span');
    title.textContent = label;
    head.appendChild(title);

    const select = document.createElement('select');
    select.setAttribute('aria-label', `Columna de guía para ${label}`);
    fs.headers.forEach((h, i) => {
      const opt = document.createElement('option');
      opt.value = String(i);
      opt.textContent = h;
      if (i === fs.selectedCol) opt.selected = true;
      select.appendChild(opt);
    });
    select.addEventListener('change', () => {
      fs.selectedCol = Number(select.value);
      lastColumnChoice[key] = fs.headers[fs.selectedCol];
      renderMapping();
    });
    head.appendChild(select);
    card.appendChild(head);

    lastColumnChoice[key] = fs.headers[fs.selectedCol];

    const table = document.createElement('table');
    table.className = 'map-preview';
    const thead = document.createElement('thead');
    const trh = document.createElement('tr');
    fs.headers.forEach((h, i) => {
      const th = document.createElement('th');
      th.textContent = h;
      if (i === fs.selectedCol) th.classList.add('is-guia-col');
      trh.appendChild(th);
    });
    thead.appendChild(trh);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');
    fs.rows.slice(0, 5).forEach((row) => {
      const tr = document.createElement('tr');
      fs.headers.forEach((_, i) => {
        const td = document.createElement('td');
        const v = row[i];
        td.textContent = v === null || v === undefined ? '' : String(v);
        if (i === fs.selectedCol) td.classList.add('is-guia-col');
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    card.appendChild(table);

    // conteo de filas que se descartarían con la columna actual
    let invalidCount = 0;
    fs.rows.forEach((row) => {
      if (normalizeGuia(row[fs.selectedCol]) === null) invalidCount++;
    });
    const note = document.createElement('p');
    note.className = 'map-card-note' + (invalidCount > 0 ? ' is-warn' : '');
    note.textContent = invalidCount > 0
      ? `${invalidCount} de ${fs.rows.length} filas se descartarán (vacías o no numéricas) con esta columna.`
      : `${fs.rows.length} filas válidas con esta columna.`;
    card.appendChild(note);

    mappingGrid.appendChild(card);
  });

  const allLoaded = SOURCES.every(({ key }) => state.files[key]);
  btnProcess.disabled = !allLoaded;
  processHint.textContent = allLoaded
    ? 'Listo para conciliar.'
    : 'Selecciona la columna de guía en los 3 archivos para continuar.';
}

btnProcess.addEventListener('click', () => {
  runReconciliation();
});

/* ---------------------------------------------------------------------
 * Conciliación 3-way
 * ------------------------------------------------------------------- */

function runReconciliation() {
  const perSource = {};
  const duplicates = {};
  const discardedCounts = {};

  SOURCES.forEach(({ key }) => {
    const fs = state.files[key];
    const map = new Map(); // normalized -> raw display value (first occurrence)
    const counts = new Map();
    let discarded = 0;

    fs.rows.forEach((row) => {
      const rawVal = row[fs.selectedCol];
      const norm = normalizeGuia(rawVal);
      if (norm === null) { discarded++; return; }
      if (!map.has(norm)) map.set(norm, rawVal === null || rawVal === undefined ? norm : String(rawVal).trim());
      counts.set(norm, (counts.get(norm) || 0) + 1);
    });

    perSource[key] = map;
    discardedCounts[key] = discarded;
    duplicates[key] = [...counts.entries()].filter(([, c]) => c > 1);
  });

  const unionKeys = new Set();
  SOURCES.forEach(({ key }) => { for (const k of perSource[key].keys()) unionKeys.add(k); });

  const rows = [...unionKeys].sort((a, b) => {
    const na = Number(a), nb = Number(b);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.localeCompare(b);
  }).map((guia) => {
    const presence = {};
    let presentCount = 0;
    SOURCES.forEach(({ key }) => {
      const has = perSource[key].has(guia);
      presence[key] = has ? perSource[key].get(guia) : null;
      if (has) presentCount++;
    });
    let status;
    if (presentCount === 3) status = 'match-all';
    else if (presentCount === 2) status = 'missing-one';
    else status = 'only-one';
    return { guia, presence, presentCount, status };
  });

  const stats = {
    total: rows.length,
    matchAll: rows.filter((r) => r.status === 'match-all').length,
    missingOne: rows.filter((r) => r.status === 'missing-one').length,
    onlyOne: rows.filter((r) => r.status === 'only-one').length,
  };

  state.reconciled = { rows, stats, duplicates, discardedCounts };
  resultsSection.hidden = false;
  renderResults();
  resultsSection.scrollIntoView({ behavior: window.matchMedia('(prefers-reduced-motion: reduce)').matches ? 'auto' : 'smooth', block: 'start' });
}

/* ---------------------------------------------------------------------
 * Render de resultados
 * ------------------------------------------------------------------- */

function sourceLabel(key) {
  return SOURCES.find((s) => s.key === key).label;
}

function missingSourcesText(row) {
  return SOURCES.filter(({ key }) => row.presence[key] === null).map(({ label }) => label).join(', ');
}
function presentSourcesText(row) {
  return SOURCES.filter(({ key }) => row.presence[key] !== null).map(({ label }) => label).join(', ');
}

function statusLabel(row) {
  if (row.status === 'match-all') return 'Conciliado';
  if (row.status === 'missing-one') return `Falta en ${missingSourcesText(row)}`;
  return `Solo en ${presentSourcesText(row)}`;
}

const SEAL_SVG = `<svg class="seal" viewBox="0 0 24 24" aria-hidden="true">
  <circle cx="12" cy="12" r="10" fill="none" stroke="currentColor" stroke-width="1.6"/>
  <path d="M7.5 12.5l3 3 6-6.5" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
</svg>`;

function renderResults() {
  if (!state.reconciled) return;
  const { stats, duplicates, discardedCounts } = state.reconciled;

  statsGrid.innerHTML = `
    <div class="stat-card">
      <span class="stat-num">${stats.total}</span>
      <span class="stat-label">Total de guías</span>
    </div>
    <div class="stat-card stat-emerald">
      <span class="stat-num">${stats.matchAll}</span>
      <span class="stat-label">Coinciden en los 3</span>
    </div>
    <div class="stat-card stat-amber">
      <span class="stat-num">${stats.missingOne}</span>
      <span class="stat-label">Falta en 1 archivo</span>
    </div>
    <div class="stat-card stat-rust">
      <span class="stat-num">${stats.onlyOne}</span>
      <span class="stat-label">Aparece en 1 solo</span>
    </div>
  `;

  const warnItems = [];
  SOURCES.forEach(({ key, label }) => {
    if (discardedCounts[key] > 0) {
      warnItems.push(`${label}: ${discardedCounts[key]} fila(s) descartada(s) por estar vacías o no ser numéricas.`);
    }
    if (duplicates[key].length > 0) {
      const list = duplicates[key].slice(0, 8).map(([g, c]) => `${g} (x${c})`).join(', ');
      const extra = duplicates[key].length > 8 ? ` y ${duplicates[key].length - 8} más` : '';
      warnItems.push(`${label}: guías duplicadas dentro del mismo archivo — ${list}${extra}.`);
    }
  });
  if (warnItems.length > 0) {
    warningsBox.hidden = false;
    warningsBox.innerHTML = `<h3>Avisos</h3><ul>${warnItems.map((w) => `<li>${escapeHtml(w)}</li>`).join('')}</ul>`;
  } else {
    warningsBox.hidden = true;
    warningsBox.innerHTML = '';
  }

  renderTable();
}

function escapeHtml(s) {
  const div = document.createElement('div');
  div.textContent = s;
  return div.innerHTML;
}

function getFilteredRows() {
  const { rows } = state.reconciled;
  const q = searchInput.value.trim().toLowerCase();
  const statusVal = statusFilter.value;
  return rows.filter((r) => {
    if (statusVal !== 'all' && r.status !== statusVal) return false;
    if (q && !r.guia.toLowerCase().includes(q)) return false;
    return true;
  });
}

function renderTable() {
  const filtered = getFilteredRows();
  ledgerBody.innerHTML = '';
  ledgerEmpty.hidden = filtered.length > 0;

  const reduceMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  filtered.forEach((row, i) => {
    const tr = document.createElement('tr');
    if (!reduceMotion) tr.style.animationDelay = `${Math.min(i, 40) * 18}ms`;

    const tdGuia = document.createElement('td');
    tdGuia.className = 'cell-guia';
    tdGuia.textContent = row.guia;
    tr.appendChild(tdGuia);

    SOURCES.forEach(({ key }) => {
      const td = document.createElement('td');
      const val = row.presence[key];
      td.className = 'cell-src ' + (val !== null ? 'is-present' : 'is-absent');
      td.textContent = val !== null ? val : '—';
      tr.appendChild(td);
    });

    const tdStatus = document.createElement('td');
    const pill = document.createElement('span');
    pill.className = `status-pill ${row.status}`;
    if (row.status === 'match-all') pill.innerHTML = SEAL_SVG + '<span>Conciliado</span>';
    else pill.innerHTML = `<span>${escapeHtml(statusLabel(row))}</span>`;
    tdStatus.appendChild(pill);
    tr.appendChild(tdStatus);

    ledgerBody.appendChild(tr);
  });
}

searchInput.addEventListener('input', renderTable);
statusFilter.addEventListener('change', renderTable);

/* ---------------------------------------------------------------------
 * Exportación a Excel
 * ------------------------------------------------------------------- */

btnExport.addEventListener('click', () => {
  if (!state.reconciled) return;
  const { rows } = state.reconciled;

  const toRecord = (r) => ({
    'Guía': r.guia,
    'Laboratorio': r.presence.lab !== null ? r.presence.lab : '',
    'JD Edwards': r.presence.jde !== null ? r.presence.jde : '',
    'Proveedor': r.presence.prov !== null ? r.presence.prov : '',
    'Estado': statusLabel(r),
  });

  const summaryData = rows.map(toRecord);
  const discrepancyData = rows.filter((r) => r.status !== 'match-all').map(toRecord);

  const wb = XLSX.utils.book_new();
  const wsSummary = XLSX.utils.json_to_sheet(summaryData);
  const wsDiscrepancy = XLSX.utils.json_to_sheet(discrepancyData);
  wsSummary['!cols'] = [{ wch: 14 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 24 }];
  wsDiscrepancy['!cols'] = wsSummary['!cols'];
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Resumen');
  XLSX.utils.book_append_sheet(wb, wsDiscrepancy, 'Discrepancias');

  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(wb, `conciliacion_guias_bravoenergy_${stamp}.xlsx`);
});

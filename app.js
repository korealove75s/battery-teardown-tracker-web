'use strict';

// ---------------------------------------------------------------------------
// Sheet (tab) definitions — each tab is its own IndexedDB object store
// ---------------------------------------------------------------------------
const TEARDOWN_COLUMNS = [
  { key: 'lot', label: 'LOT', type: 'text', sticky: 1 },
  { key: 'grade', label: 'Grade', type: 'text' },
  { key: 'count', label: 'Count', type: 'text' },
  { key: 'cellId', label: 'Cell ID', type: 'text', sticky: 2 },
  { key: 'ocv', label: 'OCV (V)', type: 'text' },
  { key: 'acir', label: 'ACIR (mΩ)', type: 'text' },
  { key: 'iv', label: 'IV', type: 'text' },
  { key: 'ir', label: 'IR', type: 'text' },
  { key: 'tdStatus', label: 'Tear Down Status', type: 'text' },
  { key: 'tdDate', label: 'TD Date', type: 'text' },
  { key: 'operator', label: 'Operator', type: 'text' },
  { key: 'tdResult', label: 'TD Result', type: 'long' },
  { key: 'electrodeLayer', label: 'Electrode Layer', type: 'text' },
  { key: 'topBack', label: 'Top/Back', type: 'text' },
  { key: 'x', label: 'X', type: 'text' },
  { key: 'y', label: 'Y', type: 'text' },
  { key: 'images', label: 'VHX/Image', type: 'image' },
  { key: 'furtherAnalysis', label: 'Further Analysis', type: 'long' },
];
const ANALYSIS_COLUMNS = [
  { key: 'lot', label: 'LOT', type: 'text', sticky: 1 },
  { key: 'cellId', label: 'Cell ID', type: 'text', sticky: 2 },
  { key: 'grade', label: 'Grade', type: 'text' },
  { key: 'docv', label: 'dOCV', type: 'text' },
  { key: 'frozenIr', label: 'Frozen IR Result (35MOhm)', type: 'text' },
  { key: 'frozenIrPf', label: 'Frozen IR Pass/Fail', type: 'text' },
  { key: 'voltageDrop', label: 'voltage drop / no drop', type: 'text' },
  { key: 'droppedLayer', label: 'Voltage Dropped Layer', type: 'text' },
  { key: 'docvV', label: 'dOCV (V)', type: 'text' },
  { key: 'spotFound', label: 'Spot Found', type: 'text' },
  { key: 'topBack', label: 'Top/ Back', type: 'text' },
  { key: 'x', label: 'x', type: 'text' },
  { key: 'y', label: 'y', type: 'text' },
  { key: 'shape', label: 'Shape', type: 'text' },
  { key: 'semEds', label: 'SEM/EDS Analysis', type: 'long' },
  { key: 'location', label: 'Location', type: 'text' },
  { key: 'longSide', label: 'Long side', type: 'text' },
  { key: 'shortSide', label: 'Short side', type: 'text' },
  { key: 'height', label: 'Height', type: 'text' },
];
const SHEETS = {
  teardown: {
    label: 'Tear Down',
    store: 'cells',
    columns: TEARDOWN_COLUMNS,
    filters: [['lot', 'LOT'], ['grade', 'Grade'], ['operator', 'Operator'], ['tdStatus', 'TD Status']],
    dateKey: 'tdDate',
    statusKey: 'tdStatus',
    aliases: { 'OCV': 'ocv', 'ACIR': 'acir', 'TD Status': 'tdStatus', 'TopBack': 'topBack' },
    exportName: 'battery-teardown-export',
    sheetName: 'TearDown',
  },
  analysis: {
    label: 'Frozen IR · Spot 분석',
    store: 'analysis',
    columns: ANALYSIS_COLUMNS,
    filters: [['lot', 'LOT'], ['grade', 'Grade'], ['frozenIrPf', 'Frozen IR P/F'], ['voltageDrop', 'Voltage drop'], ['spotFound', 'Spot Found']],
    dateKey: null,
    statusKey: 'frozenIrPf',
    aliases: { 'Frozen IR Result': 'frozenIr', 'Frozen IR': 'frozenIr', 'Frozen IR P/F': 'frozenIrPf', 'Voltage Drop': 'voltageDrop', 'SEM/EDS': 'semEds' },
    exportName: 'battery-analysis-export',
    sheetName: 'Analysis',
  },
};
Object.values(SHEETS).forEach((s) => {
  s.editableKeys = s.columns.filter((c) => c.type !== 'image').map((c) => c.key);
  s.hasImages = s.columns.some((c) => c.type === 'image');
  s.headerMap = {};
  s.columns.forEach((c) => { s.headerMap[normalizeHeader(c.label)] = c.key; });
  Object.entries(s.aliases).forEach(([h, k]) => { s.headerMap[normalizeHeader(h)] = k; });
});

// Active-sheet bindings, reassigned by setActiveSheet()
let activeSheetKey = 'teardown';
let SHEET = SHEETS.teardown;
let COLUMNS = SHEET.columns;
let EDITABLE_KEYS = SHEET.editableKeys;
let HEADER_MAP = SHEET.headerMap;
function curStore() { return SHEET.store; }

function normalizeHeader(h) {
  return String(h || '').toLowerCase().replace(/[\s()./\-\u03a9]/g, '');
}
function emptyFields() {
  return EDITABLE_KEYS.reduce((acc, k) => { acc[k] = ''; return acc; }, {});
}

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
const state = {
  cells: [],
  data: { teardown: [], analysis: [] },
  loaded: false,
  filters: { search: '', date: '', sel: {} },
  formEditingId: null,
  formImages: [],
  formStagedFiles: [],
  detailId: null,
};

// ---------------------------------------------------------------------------
// IndexedDB persistence (per-browser storage — no server, no account needed)
// ---------------------------------------------------------------------------
const IDB_NAME = 'battery-teardown-tracker';
const IDB_VERSION = 2; // v2: added 'analysis' store for the second tab
let idbPromise = null;
function openIdb() {
  if (idbPromise) return idbPromise;
  idbPromise = new Promise((resolve, reject) => {
    if (!('indexedDB' in window)) { reject(new Error('indexeddb_unsupported')); return; }
    const req = indexedDB.open(IDB_NAME, IDB_VERSION);
    req.onupgradeneeded = (e) => {
      const database = e.target.result;
      if (!database.objectStoreNames.contains('cells')) database.createObjectStore('cells', { keyPath: 'id' });
      if (!database.objectStoreNames.contains('images')) database.createObjectStore('images', { keyPath: 'id' });
      if (!database.objectStoreNames.contains('analysis')) database.createObjectStore('analysis', { keyPath: 'id' });
    };
    req.onsuccess = () => {
      const database = req.result;
      database.onversionchange = () => { database.close(); showToast('새 버전이 다른 탭에서 열렸습니다. 이 페이지를 새로고침해 주세요.', true); };
      resolve(database);
    };
    req.onerror = () => reject(req.error);
    req.onblocked = () => showToast('이 사이트가 열려 있는 다른 탭을 닫거나 새로고침해 주세요 (저장소 업데이트 대기 중).', true);
  });
  return idbPromise;
}
function idbStore(storeName, mode) { return openIdb().then((d) => d.transaction(storeName, mode).objectStore(storeName)); }
function idbGetAll(storeName) {
  return idbStore(storeName, 'readonly').then((store) => new Promise((resolve, reject) => {
    const req = store.getAll();
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
}
function idbGet(storeName, id) {
  return idbStore(storeName, 'readonly').then((store) => new Promise((resolve, reject) => {
    const req = store.get(id);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
}
function idbPut(storeName, value) {
  return idbStore(storeName, 'readwrite').then((store) => new Promise((resolve, reject) => {
    const req = store.put(value);
    req.onsuccess = () => resolve(value);
    req.onerror = () => reject(req.error);
  }));
}
function idbDelete(storeName, id) {
  return idbStore(storeName, 'readwrite').then((store) => new Promise((resolve, reject) => {
    const req = store.delete(id);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  }));
}
function genId() {
  return (crypto.randomUUID ? crypto.randomUUID() : 'id-' + Date.now() + '-' + Math.random().toString(36).slice(2));
}

async function initStorage() {
  try {
    await openIdb();
  } catch (err) {
    console.error(err);
    setLiveBadge('unavailable', '저장 불가');
    showLoading(false);
    showEmpty(true, '이 브라우저는 로컬 저장(IndexedDB)을 지원하지 않거나 비활성화되어 있습니다. 시크릿 모드가 아닌 일반 창에서 열어 주세요.');
    return;
  }
  await loadCells();
  setLiveBadge('live', '이 브라우저에 저장됨');
}

function setLiveBadge(mode, text) {
  const el = document.getElementById('liveBadge');
  el.className = 'live-badge ' + mode;
  document.getElementById('liveBadgeText').textContent = text;
}
function showLoading(v) { document.getElementById('loadingState').hidden = !v; }
function showEmpty(v, msg) {
  const el = document.getElementById('emptyState');
  el.hidden = !v;
  if (msg) el.textContent = msg;
}

async function loadCells() {
  for (const [key, sheet] of Object.entries(SHEETS)) {
    const rows = await idbGetAll(sheet.store);
    rows.sort((a, b) => String(a.createdAt || '').localeCompare(String(b.createdAt || '')));
    state.data[key] = rows;
  }
  state.loaded = true;
  setActiveSheet(activeSheetKey);
  showLoading(false);
}

// ---------------------------------------------------------------------------
// Cell mutation helpers
// ---------------------------------------------------------------------------
async function createCell(fields) {
  const now = new Date().toISOString();
  const rec = Object.assign(emptyFields(), fields, { id: genId(), images: [], createdAt: now, updatedAt: now });
  await idbPut(curStore(), rec);
  state.cells.push(rec);
  renderSheetTabs();
  renderStats();
  renderTable();
  return rec.id;
}
async function updateCell(id, fields) {
  const cell = state.cells.find((c) => c.id === id);
  if (!cell) return;
  Object.assign(cell, fields, { updatedAt: new Date().toISOString() });
  await idbPut(curStore(), cell);
}
async function deleteCell(cell) {
  if (cell.images && cell.images.length) {
    for (const img of cell.images) { try { await idbDelete('images', img.id); } catch (err) { console.warn('image delete failed', err); } }
  }
  await idbDelete(curStore(), cell.id);
  state.cells = state.data[activeSheetKey] = state.cells.filter((c) => c.id !== cell.id);
  renderSheetTabs();
  renderStats();
  renderTable();
}
async function addImagesToCell(cell, files) {
  const uploaded = [];
  for (const f of files) {
    const id = genId();
    await idbPut('images', { id, blob: f, contentType: f.type, originalName: f.name });
    imageUrlCache.set(id, URL.createObjectURL(f));
    uploaded.push({ id, originalName: f.name, contentType: f.type, uploadedAt: new Date().toISOString() });
  }
  const newImages = (cell.images || []).concat(uploaded);
  await updateCell(cell.id, { images: newImages });
  const live = state.cells.find((c) => c.id === cell.id);
  return live ? live.images : newImages;
}
async function removeImageFromCell(cell, imageId) {
  try { await idbDelete('images', imageId); } catch (err) { console.warn('image delete failed', err); }
  imageUrlCache.delete(imageId);
  const newImages = (cell.images || []).filter((im) => im.id !== imageId);
  await updateCell(cell.id, { images: newImages });
  const live = state.cells.find((c) => c.id === cell.id);
  return live ? live.images : newImages;
}

// Image blobs are read from IndexedDB asynchronously; cache resolved object
// URLs per image id and attach them to <img> elements once ready.
const imageUrlCache = new Map();
function attachImageSrc(imgEl, imageId) {
  const cached = imageUrlCache.get(imageId);
  if (cached) { imgEl.src = cached; return; }
  idbGet('images', imageId).then((rec) => {
    if (!rec || !rec.blob) return;
    const url = URL.createObjectURL(rec.blob);
    imageUrlCache.set(imageId, url);
    imgEl.src = url;
  }).catch((err) => console.warn('image load failed', err));
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------
function debounce(fn, ms) { let t; return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); }; }
function escapeHtml(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function cssEsc(s) { return (window.CSS && CSS.escape) ? CSS.escape(s) : String(s).replace(/["\\]/g, '\\$&'); }
let toastTimer;
function showToast(msg, isError) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 3200);
}
function looksLikeDateMissingYear(str) {
  if (!str) return false;
  const s = String(str).trim();
  if (!s) return false;
  if (/\d{4}/.test(s)) return false;
  if (/^\d{1,2}\s*[\/.\-]\s*\d{1,2}$/.test(s)) return true;
  if (/^\d{1,2}\s*월\s*\d{1,2}\s*일?$/.test(s)) return true;
  return false;
}
function applyYearToDate(original, year) {
  const s = String(original).trim();
  let m = s.match(/^(\d{1,2})\s*[\/.\-]\s*(\d{1,2})$/);
  if (m) return `${year}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  m = s.match(/^(\d{1,2})\s*월\s*(\d{1,2})\s*일?$/);
  if (m) return `${year}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`;
  return `${year} ${s}`;
}
function isDuplicateLocally(cell) {
  const v = (cell.cellId || '').trim().toLowerCase();
  if (!v) return false;
  return state.cells.some((c) => c.id !== cell.id && (c.cellId || '').trim().toLowerCase() === v);
}

// ---------------------------------------------------------------------------
// TD Date year-resolution modal
// ---------------------------------------------------------------------------
function resolveMissingYearDates(items) {
  return new Promise((resolve) => {
    const overlay = document.getElementById('yearModalOverlay');
    const body = document.getElementById('yearModalBody');
    body.innerHTML = '';
    const info = document.createElement('p');
    info.className = 'field-warning';
    info.textContent = `TD Date에 연도가 없는 항목이 ${items.length}건 있습니다. 임의로 연도를 지정하지 않습니다 — 직접 입력하거나 "원본 유지"를 선택하세요.`;
    body.appendChild(info);
    if (items.length > 1) {
      const bulkRow = document.createElement('div');
      bulkRow.className = 'import-issue-row';
      bulkRow.innerHTML = `<span class="label">전체 일괄 연도 적용</span><input type="number" id="bulkYearInput" placeholder="예: 2026" min="1990" max="2100" /><button class="btn btn-ghost" id="bulkYearApplyBtn" type="button">전체 적용</button>`;
      body.appendChild(bulkRow);
    }
    items.forEach((item, idx) => {
      const row = document.createElement('div');
      row.className = 'import-issue-row';
      row.innerHTML = `<span class="label">${escapeHtml(item.label)} — 원본: "${escapeHtml(item.original)}"</span><input type="number" class="year-input" data-idx="${idx}" placeholder="연도" min="1990" max="2100" /><label style="font-size:11px; display:flex; align-items:center; gap:4px;"><input type="checkbox" class="keep-as-is" data-idx="${idx}" /> 원본 유지</label>`;
      body.appendChild(row);
    });
    const bulkBtn = document.getElementById('bulkYearApplyBtn');
    if (bulkBtn) {
      bulkBtn.onclick = () => {
        const y = document.getElementById('bulkYearInput').value;
        if (!y) return;
        body.querySelectorAll('.year-input').forEach((inp) => { inp.value = y; });
      };
    }
    overlay.classList.add('open');
    function cleanup() {
      overlay.classList.remove('open');
      document.getElementById('yearConfirmBtn').onclick = null;
      document.getElementById('yearCancelBtn').onclick = null;
      document.getElementById('yearModalClose').onclick = null;
    }
    document.getElementById('yearCancelBtn').onclick = () => { cleanup(); resolve(false); };
    document.getElementById('yearModalClose').onclick = () => { cleanup(); resolve(false); };
    document.getElementById('yearConfirmBtn').onclick = () => {
      for (let idx = 0; idx < items.length; idx++) {
        const yearInput = body.querySelector(`.year-input[data-idx="${idx}"]`);
        const keepAsIs = body.querySelector(`.keep-as-is[data-idx="${idx}"]`);
        if (keepAsIs.checked) { items[idx].apply(items[idx].original); }
        else if (yearInput.value) { items[idx].apply(applyYearToDate(items[idx].original, yearInput.value)); }
        else { showToast('모든 항목에 연도를 입력하거나 "원본 유지"를 선택하세요.', true); return; }
      }
      cleanup(); resolve(true);
    };
  });
}

// ---------------------------------------------------------------------------
// Stats row
// ---------------------------------------------------------------------------
function renderStats() {
  document.getElementById('statTotal').textContent = `레코드 ${state.cells.length}건`;
  const row = document.getElementById('statsRow');
  row.querySelectorAll('.stat-chip').forEach((el) => el.remove());
  const counts = new Map();
  state.cells.forEach((c) => {
    const v = (c[SHEET.statusKey] || '').trim();
    if (!v) return;
    counts.set(v, (counts.get(v) || 0) + 1);
  });
  [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8).forEach(([status, n]) => {
    const chip = document.createElement('span');
    chip.className = 'stat-chip';
    chip.innerHTML = `${escapeHtml(status)} <b>${n}</b>`;
    row.appendChild(chip);
  });
}

// ---------------------------------------------------------------------------
// Table rendering (preserves focus/selection across realtime re-renders)
// ---------------------------------------------------------------------------
function renderTableHead() {
  const headRow = document.getElementById('tableHeadRow');
  if (headRow.dataset.built === activeSheetKey) return;
  headRow.innerHTML = '';
  COLUMNS.forEach((col) => {
    const th = document.createElement('th');
    th.textContent = col.label;
    if (col.sticky === 1) th.classList.add('col-sticky-1');
    if (col.sticky === 2) th.classList.add('col-sticky-2');
    headRow.appendChild(th);
  });
  const th = document.createElement('th');
  th.textContent = 'Actions';
  headRow.appendChild(th);
  headRow.dataset.built = activeSheetKey;
}
function getFilteredCells() {
  const f = state.filters;
  return state.cells.filter((c) => {
    if (f.search && !(c.cellId || '').toLowerCase().includes(f.search.toLowerCase())) return false;
    for (const [key, val] of Object.entries(f.sel)) { if (val && c[key] !== val) return false; }
    if (SHEET.dateKey && f.date && !(c[SHEET.dateKey] || '').toLowerCase().includes(f.date.toLowerCase())) return false;
    return true;
  });
}
function renderTable() {
  renderTableHead();
  const tbody = document.getElementById('tableBody');
  const active = document.activeElement;
  let preserve = null;
  if (active && active.classList && active.classList.contains('cell-input')) {
    const tr = active.closest('tr');
    if (tr) preserve = { rowId: tr.dataset.id, field: active.dataset.field, value: active.value, selStart: active.selectionStart, selEnd: active.selectionEnd };
  }
  tbody.innerHTML = '';
  const rows = getFilteredCells();
  const frag = document.createDocumentFragment();
  rows.forEach((cell) => frag.appendChild(createRowElement(cell, preserve)));
  tbody.appendChild(frag);
  updateStickyOffset();
  showEmpty(state.loaded && state.cells.length === 0);
  if (preserve) {
    const el = tbody.querySelector(`tr[data-id="${cssEsc(preserve.rowId)}"] [data-field="${cssEsc(preserve.field)}"]`);
    if (el) {
      el.focus({ preventScroll: true });
      if (typeof preserve.selStart === 'number') { try { el.setSelectionRange(preserve.selStart, preserve.selEnd); } catch (e) {} }
    }
  }
}
// Pin the second frozen column (Cell ID) right after the first (LOT), using LOT's real width
function updateStickyOffset() {
  const first = document.querySelector('#tableHeadRow th.col-sticky-1');
  const width = first ? first.getBoundingClientRect().width : 0;
  document.getElementById('mainTable').style.setProperty('--sticky2-left', width + 'px');
}
function createRowElement(cell, preserve) {
  const tr = document.createElement('tr');
  tr.dataset.id = cell.id;
  tr.className = 'row-clickable';
  COLUMNS.forEach((col) => {
    const td = document.createElement('td');
    if (col.sticky === 1) td.classList.add('col-sticky-1');
    if (col.sticky === 2) td.classList.add('col-sticky-2');
    if (col.type === 'text' || col.type === 'long') {
      const input = document.createElement(col.type === 'long' ? 'textarea' : 'input');
      if (col.type === 'long') { input.rows = 1; input.placeholder = 'Shift+Enter 줄바꿈'; }
      else input.type = 'text';
      input.className = 'cell-input';
      input.dataset.field = col.key;
      const usePreserved = preserve && preserve.rowId === cell.id && preserve.field === col.key;
      input.value = usePreserved ? preserve.value : (cell[col.key] || '');
      if (col.key === 'cellId' && isDuplicateLocally(cell)) input.classList.add('dup-cell');
      td.appendChild(input);
    } else if (col.type === 'image') {
      td.appendChild(renderImagesCell(cell));
    }
    tr.appendChild(td);
  });
  const actionsTd = document.createElement('td');
  const actions = document.createElement('div');
  actions.className = 'row-actions';
  const viewBtn = document.createElement('button');
  viewBtn.className = 'icon-btn'; viewBtn.type = 'button'; viewBtn.title = '상세보기'; viewBtn.textContent = '🔍';
  viewBtn.onclick = (e) => { e.stopPropagation(); openDetailModal(cell); };
  actions.appendChild(viewBtn);
  const delBtn = document.createElement('button');
  delBtn.className = 'icon-btn danger'; delBtn.type = 'button'; delBtn.title = '행 삭제'; delBtn.textContent = '🗑';
  delBtn.onclick = async (e) => {
    e.stopPropagation();
    if (!confirm(`Cell ID "${cell.cellId || '(없음)'}" 행을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.`)) return;
    try { await deleteCell(cell); showToast('행이 삭제되었습니다.'); }
    catch (err) { console.error(err); showToast('삭제 중 오류가 발생했습니다.', true); }
  };
  actions.appendChild(delBtn);
  actionsTd.appendChild(actions);
  tr.appendChild(actionsTd);
  return tr;
}
function renderImagesCell(cell) {
  const wrap = document.createElement('div');
  wrap.className = 'cell-images';
  const images = cell.images || [];
  images.slice(0, 3).forEach((img) => {
    const thumb = document.createElement('img');
    thumb.className = 'thumb'; thumb.alt = img.originalName || ''; thumb.title = img.originalName || '';
    attachImageSrc(thumb, img.id);
    thumb.onclick = (e) => { e.stopPropagation(); openLightbox(images, images.indexOf(img)); };
    wrap.appendChild(thumb);
  });
  if (images.length > 3) {
    const more = document.createElement('span'); more.className = 'thumb-more'; more.textContent = `+${images.length - 3}`;
    wrap.appendChild(more);
  }
  const addBtn = document.createElement('button');
  addBtn.className = 'thumb-add'; addBtn.type = 'button'; addBtn.textContent = '+'; addBtn.title = '이미지 첨부';
  addBtn.onclick = (e) => { e.stopPropagation(); triggerImageUpload(cell); };
  wrap.appendChild(addBtn);
  return wrap;
}
function triggerImageUpload(cell) {
  const input = document.createElement('input');
  input.type = 'file'; input.accept = 'image/*'; input.multiple = true;
  input.onchange = async () => {
    if (!input.files.length) return;
    try { await addImagesToCell(cell, Array.from(input.files)); showToast('이미지 업로드 완료'); }
    catch (err) { console.error(err); showToast('이미지 업로드 실패', true); }
  };
  input.click();
}

// ---------------------------------------------------------------------------
// Inline cell edit
// ---------------------------------------------------------------------------
async function handleCellChange(e) {
  const el = e.target;
  if (!el.classList || !el.classList.contains('cell-input')) return;
  const tr = el.closest('tr');
  const id = tr.dataset.id;
  const field = el.dataset.field;
  const cell = state.cells.find((c) => c.id === id);
  if (!cell) return;
  let value = el.value;
  if (field === 'tdDate' && looksLikeDateMissingYear(value)) {
    let finalVal = value;
    const ok = await resolveMissingYearDates([{ label: 'TD Date', original: value, apply: (v) => { finalVal = v; } }]);
    if (!ok) { el.value = cell.tdDate; return; }
    value = finalVal;
    el.value = value;
  }
  if (value === (cell[field] || '')) return;
  const prev = cell[field];
  cell[field] = value;
  try {
    await updateCell(id, { [field]: value });
    if (field === SHEET.statusKey) renderStats();
    if (SHEET.filters.some(([k]) => k === field)) populateFilterOptions();
    if (field === 'cellId') {
      const dup = state.cells.filter((c) => c.id !== cell.id && (c.cellId || '').trim().toLowerCase() === value.trim().toLowerCase() && value.trim());
      if (dup.length) showToast(`⚠ Cell ID 중복: "${value}" — 기존 ${dup.length}건 존재 (LOT: ${dup.map((d) => d.lot || '-').join(', ')})`, true);
      renderTable();
    }
  } catch (err) {
    console.error(err);
    cell[field] = prev; el.value = prev;
    showToast('저장 실패', true);
  }
}

// ---------------------------------------------------------------------------
// Paste-from-Excel
// ---------------------------------------------------------------------------
async function onTablePaste(e) {
  const target = e.target;
  if (!target.classList || !target.classList.contains('cell-input')) return;
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  if (!text) return;
  // Multi-line prose pasted into a long-text cell stays in that cell
  if (target.tagName === 'TEXTAREA' && !text.includes('\t')) return;
  e.preventDefault();
  const lines = text.replace(/\r/g, '').split('\n');
  if (lines.length && lines[lines.length - 1] === '') lines.pop();
  const grid = lines.map((line) => line.split('\t'));
  if (grid.length === 0) return;
  if (grid.length === 1 && grid[0].length === 1) {
    target.value = grid[0][0];
    target.dispatchEvent(new Event('change', { bubbles: true }));
    return;
  }
  const startColIdx = EDITABLE_KEYS.indexOf(target.dataset.field);
  if (startColIdx === -1) return;
  const startRowId = target.closest('tr').dataset.id;
  const filteredRows = getFilteredCells();
  const startRowIdx = filteredRows.findIndex((r) => r.id === startRowId);
  if (startRowIdx === -1) return;
  await handleGridPaste(grid, filteredRows, startRowIdx, startColIdx);
}
async function handleGridPaste(grid, filteredRows, startRowIdx, startColIdx) {
  const updates = [];
  const creations = [];
  grid.forEach((line, i) => {
    const rowIdx = startRowIdx + i;
    let target = rowIdx < filteredRows.length
      ? { existing: true, id: filteredRows[rowIdx].id, fields: {} }
      : { existing: false, fields: {} };
    line.forEach((val, j) => {
      const colIdx = startColIdx + j;
      if (colIdx >= EDITABLE_KEYS.length) return;
      target.fields[EDITABLE_KEYS[colIdx]] = val;
    });
    (target.existing ? updates : creations).push(target);
  });
  const yearItems = [];
  [...updates, ...creations].forEach((t, idx) => {
    if (t.fields.tdDate && looksLikeDateMissingYear(t.fields.tdDate)) {
      const existingCell = t.existing ? filteredRows.find((r) => r.id === t.id) : null;
      yearItems.push({
        label: t.fields.cellId ? `Cell ID: ${t.fields.cellId}` : (existingCell ? `기존 행 (Cell ID: ${existingCell.cellId || '-'})` : `신규 행 #${idx + 1}`),
        original: t.fields.tdDate,
        apply: (v) => { t.fields.tdDate = v; },
      });
    }
  });
  if (yearItems.length) {
    const ok = await resolveMissingYearDates(yearItems);
    if (!ok) { showToast('붙여넣기가 취소되었습니다.'); return; }
  }
  try {
    for (const u of updates) { await updateCell(u.id, u.fields); }
    for (const c of creations) { await createCell(c.fields); }
    showToast(`${grid.length}행 붙여넣기 완료`);
  } catch (err) {
    console.error(err);
    showToast('붙여넣기 중 오류가 발생했습니다.', true);
  }
}

// ---------------------------------------------------------------------------
// Lightbox
// ---------------------------------------------------------------------------
let lightboxState = { images: [], index: 0 };
function openLightbox(images, index) {
  if (!images || !images.length) return;
  lightboxState = { images, index };
  updateLightbox();
  document.getElementById('lightboxOverlay').classList.add('open');
}
function updateLightbox() {
  const img = lightboxState.images[lightboxState.index];
  attachImageSrc(document.getElementById('lightboxImg'), img.id);
  document.getElementById('lightboxCaption').textContent = `${img.originalName || ''} (${lightboxState.index + 1}/${lightboxState.images.length})`;
}
function lightboxStep(delta) {
  if (!lightboxState.images.length) return;
  lightboxState.index = (lightboxState.index + delta + lightboxState.images.length) % lightboxState.images.length;
  updateLightbox();
}

// ---------------------------------------------------------------------------
// Detail modal
// ---------------------------------------------------------------------------
function openDetailModal(cell) {
  state.detailId = cell.id;
  document.getElementById('detailModalTitle').textContent = `${cell.cellId || '(Cell ID 없음)'} 상세정보`;
  const body = document.getElementById('detailModalBody');
  body.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'detail-grid';
  COLUMNS.filter((c) => c.type !== 'image').forEach((col) => {
    const item = document.createElement('div');
    item.className = 'detail-item' + (col.type === 'long' ? ' full' : '');
    const val = cell[col.key];
    item.innerHTML = `<div class="k">${escapeHtml(col.label)}</div><div class="v ${val ? '' : 'empty'}">${val ? escapeHtml(val) : '(빈 값)'}</div>`;
    grid.appendChild(item);
  });
  body.appendChild(grid);
  if (SHEET.hasImages) {
  const imgSection = document.createElement('div');
  imgSection.className = 'detail-item full';
  const imgHeader = document.createElement('div'); imgHeader.className = 'k'; imgHeader.textContent = 'VHX/Image';
  imgSection.appendChild(imgHeader);
  const imgs = document.createElement('div'); imgs.className = 'detail-images';
  const images = cell.images || [];
  images.forEach((img, i) => {
    const t = document.createElement('img');
    t.className = 'thumb'; t.title = img.originalName || '';
    attachImageSrc(t, img.id);
    t.onclick = () => openLightbox(images, i);
    imgs.appendChild(t);
  });
  if (!images.length) {
    const empty = document.createElement('div'); empty.className = 'v empty'; empty.textContent = '(첨부된 이미지 없음)';
    imgs.appendChild(empty);
  }
  imgSection.appendChild(imgs);
  body.appendChild(imgSection);
  }
  document.getElementById('detailModalOverlay').classList.add('open');
}

// ---------------------------------------------------------------------------
// Form modal
// ---------------------------------------------------------------------------
function openFormModal(cell) {
  state.formEditingId = cell ? cell.id : null;
  state.formImages = cell ? [...(cell.images || [])] : [];
  state.formStagedFiles = [];
  document.getElementById('formModalTitle').textContent = cell ? '셀 정보 편집' : '신규 셀 등록';
  const body = document.getElementById('formModalBody');
  body.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'form-grid';
  COLUMNS.filter((c) => c.type !== 'image').forEach((col) => {
    const field = document.createElement('div');
    field.className = 'form-field' + (col.type === 'long' ? ' full' : '');
    const label = document.createElement('label'); label.textContent = col.label;
    field.appendChild(label);
    const input = document.createElement(col.type === 'long' ? 'textarea' : 'input');
    if (col.type !== 'long') input.type = 'text';
    input.dataset.field = col.key;
    input.value = cell ? (cell[col.key] || '') : '';
    input.id = `formField_${col.key}`;
    field.appendChild(input);
    if (col.key === 'tdDate') {
      const hint = document.createElement('span'); hint.className = 'field-hint';
      hint.textContent = '연도가 없는 날짜(예: 1/26)는 저장 시 연도를 확인합니다.';
      field.appendChild(hint);
    }
    if (col.key === 'ir') {
      const hint = document.createElement('span'); hint.className = 'field-hint';
      hint.textContent = '"O.F." 같은 문자도 그대로 입력할 수 있습니다.';
      field.appendChild(hint);
    }
    if (col.key === 'cellId') input.addEventListener('input', debounce(() => updateDupWarning(input.value), 300));
    grid.appendChild(field);
  });
  body.appendChild(grid);

  if (SHEET.hasImages) {
  const imgField = document.createElement('div');
  imgField.className = 'form-field full';
  const imgLabel = document.createElement('label'); imgLabel.textContent = 'VHX/Image (이미지 첨부)';
  imgField.appendChild(imgLabel);
  const attachArea = document.createElement('div'); attachArea.className = 'image-attach-area'; attachArea.id = 'formImageAttachArea';
  imgField.appendChild(attachArea);
  const addImgBtn = document.createElement('label');
  addImgBtn.className = 'btn btn-secondary file-btn'; addImgBtn.textContent = '이미지 추가';
  addImgBtn.style.marginTop = '6px'; addImgBtn.style.width = 'fit-content';
  const addImgInput = document.createElement('input');
  addImgInput.type = 'file'; addImgInput.accept = 'image/*'; addImgInput.multiple = true; addImgInput.hidden = true;
  addImgInput.onchange = async () => {
    if (state.formEditingId) {
      try {
        const cell = state.cells.find((c) => c.id === state.formEditingId) || { id: state.formEditingId, images: state.formImages };
        const updated = await addImagesToCell(cell, Array.from(addImgInput.files));
        state.formImages = updated;
        renderFormImageAttachArea();
      } catch (err) { console.error(err); showToast('이미지 업로드 실패', true); }
    } else {
      for (const f of addImgInput.files) state.formStagedFiles.push(f);
      renderFormImageAttachArea();
    }
    addImgInput.value = '';
  };
  addImgBtn.appendChild(addImgInput);
  imgField.appendChild(addImgBtn);
  body.appendChild(imgField);
  renderFormImageAttachArea();
  }

  updateDupWarning(cell ? cell.cellId : '');
  document.getElementById('formModalOverlay').classList.add('open');
  document.getElementById('formField_lot').focus();
}
function renderFormImageAttachArea() {
  const area = document.getElementById('formImageAttachArea');
  if (!area) return;
  area.innerHTML = '';
  state.formImages.forEach((img) => {
    const wrap = document.createElement('div'); wrap.style.position = 'relative';
    const t = document.createElement('img'); t.className = 'thumb'; t.title = img.originalName || '';
    attachImageSrc(t, img.id);
    t.onclick = () => openLightbox(state.formImages, state.formImages.indexOf(img));
    const rm = document.createElement('button');
    rm.textContent = '×'; rm.type = 'button'; rm.className = 'icon-btn danger';
    rm.style.cssText = 'position:absolute;top:-8px;right:-8px;background:var(--panel);border-radius:50%;width:18px;height:18px;padding:0;line-height:16px;font-size:12px;border:1px solid var(--border);';
    rm.onclick = async () => {
      if (state.formEditingId) {
        try {
          const cell = state.cells.find((c) => c.id === state.formEditingId) || { id: state.formEditingId, images: state.formImages };
          const updated = await removeImageFromCell(cell, img.id);
          state.formImages = updated;
        } catch (err) { console.error(err); showToast('이미지 삭제 실패', true); return; }
      } else {
        state.formImages = state.formImages.filter((i) => i !== img);
      }
      renderFormImageAttachArea();
    };
    wrap.appendChild(t); wrap.appendChild(rm); area.appendChild(wrap);
  });
  state.formStagedFiles.forEach((f, idx) => {
    const wrap = document.createElement('div'); wrap.style.position = 'relative';
    const t = document.createElement('img'); t.src = URL.createObjectURL(f); t.className = 'thumb'; t.title = f.name + ' (저장 시 업로드)'; t.style.opacity = '0.6';
    const rm = document.createElement('button');
    rm.textContent = '×'; rm.type = 'button'; rm.className = 'icon-btn danger';
    rm.style.cssText = 'position:absolute;top:-8px;right:-8px;background:var(--panel);border-radius:50%;width:18px;height:18px;padding:0;line-height:16px;font-size:12px;border:1px solid var(--border);';
    rm.onclick = () => { state.formStagedFiles.splice(idx, 1); renderFormImageAttachArea(); };
    wrap.appendChild(t); wrap.appendChild(rm); area.appendChild(wrap);
  });
  if (!state.formImages.length && !state.formStagedFiles.length) {
    const empty = document.createElement('span'); empty.className = 'field-hint'; empty.textContent = '첨부된 이미지가 없습니다.';
    area.appendChild(empty);
  }
}
function updateDupWarning(cellIdValue) {
  const warnEl = document.getElementById('dupWarning');
  const v = (cellIdValue || '').trim().toLowerCase();
  if (!v) { warnEl.innerHTML = ''; return; }
  const dup = state.cells.filter((c) => c.id !== state.formEditingId && (c.cellId || '').trim().toLowerCase() === v);
  if (!dup.length) { warnEl.innerHTML = ''; return; }
  warnEl.innerHTML = `⚠ 중복된 Cell ID입니다 (${dup.length}건). <a href="#" id="viewDupLink">기존 기록 보기</a>`;
  document.getElementById('viewDupLink').onclick = (e) => { e.preventDefault(); closeFormModal(); openDetailModal(dup[0]); };
}
function closeFormModal() { document.getElementById('formModalOverlay').classList.remove('open'); }
async function saveFormModal() {
  const body = document.getElementById('formModalBody');
  const fields = {};
  body.querySelectorAll('[data-field]').forEach((el) => { fields[el.dataset.field] = el.value; });
  if (fields.tdDate && looksLikeDateMissingYear(fields.tdDate)) {
    let finalVal = fields.tdDate;
    const ok = await resolveMissingYearDates([{ label: 'TD Date', original: fields.tdDate, apply: (v) => { finalVal = v; } }]);
    if (!ok) return;
    fields.tdDate = finalVal;
  }
  try {
    let id = state.formEditingId;
    if (id) { await updateCell(id, fields); }
    else { id = await createCell(fields); }
    if (state.formStagedFiles.length) {
      await addImagesToCell({ id, images: [] }, state.formStagedFiles);
    }
    closeFormModal();
    showToast(state.formEditingId ? '수정되었습니다.' : '등록되었습니다.');
  } catch (err) {
    console.error(err);
    showToast('저장 중 오류가 발생했습니다.', true);
  }
}

// ---------------------------------------------------------------------------
// Import (CSV / Excel)
// ---------------------------------------------------------------------------
async function handleImportFile(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array' });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });
  if (!rows.length) { showToast('빈 파일입니다.', true); return; }
  const headers = rows[0].map((h) => String(h || '').trim());
  const keyMap = headers.map((h) => HEADER_MAP[normalizeHeader(h)] || null);
  const ignoredHeaders = headers.filter((h, i) => !keyMap[i] && h);
  const parsed = [];
  for (let r = 1; r < rows.length; r++) {
    const raw = rows[r];
    if (!raw || raw.every((v) => String(v || '').trim() === '')) continue;
    const rec = emptyFields();
    headers.forEach((h, i) => { const key = keyMap[i]; if (key && EDITABLE_KEYS.includes(key)) rec[key] = String(raw[i] ?? '').trim(); });
    parsed.push(rec);
  }
  if (!parsed.length) { showToast('가져올 데이터가 없습니다.', true); return; }
  const existingIds = new Set(state.cells.map((c) => (c.cellId || '').trim().toLowerCase()).filter(Boolean));
  const seen = new Set();
  let dupCount = 0;
  parsed.forEach((p) => {
    const v = (p.cellId || '').trim().toLowerCase();
    if (v && (existingIds.has(v) || seen.has(v))) dupCount++;
    if (v) seen.add(v);
  });
  const yearItems = [];
  parsed.forEach((p, idx) => {
    if (p.tdDate && looksLikeDateMissingYear(p.tdDate)) {
      yearItems.push({ label: `#${idx + 1}${p.cellId ? ' Cell ID: ' + p.cellId : ''}`, original: p.tdDate, apply: (v) => { p.tdDate = v; } });
    }
  });
  const ok = await showImportReviewModal({ parsed, ignoredHeaders, dupCount, yearItems });
  if (!ok) { showToast('가져오기가 취소되었습니다.'); return; }
  for (const rec of parsed) { await createCell(rec); }
  showToast(`${parsed.length}건 가져오기 완료`);
}
function showImportReviewModal({ parsed, ignoredHeaders, dupCount, yearItems }) {
  return new Promise((resolve) => {
    const overlay = document.getElementById('importModalOverlay');
    const body = document.getElementById('importModalBody');
    body.innerHTML = '';
    const summary = document.createElement('div');
    summary.className = 'import-summary';
    summary.innerHTML = `<b>[${escapeHtml(SHEET.label)}]</b> 탭에 총 <b>${parsed.length}</b>행을 가져옵니다.<br/>${ignoredHeaders.length ? `인식되지 않아 무시된 열: ${escapeHtml(ignoredHeaders.join(', '))}<br/>` : ''}${dupCount ? `<span style="color:var(--warn)">⚠ 기존 데이터와 Cell ID가 중복되는 행: ${dupCount}건 (그래도 추가되며, 가져오기 후 표에서 중복 표시로 확인할 수 있습니다.)</span><br/>` : ''}${SHEET.hasImages ? 'VHX/Image 열의 실제 이미지 파일은 가져올 수 없어 무시됩니다. 가져온 후 각 행에서 직접 첨부해 주세요.' : ''}`;
    body.appendChild(summary);
    if (yearItems.length) {
      const yearBox = document.createElement('div');
      const warn = document.createElement('p'); warn.className = 'field-warning';
      warn.textContent = `TD Date에 연도가 없는 항목이 ${yearItems.length}건 있습니다. 임의로 연도를 지정하지 않습니다 — 직접 입력하거나 "원본 유지"를 선택하세요.`;
      yearBox.appendChild(warn);
      const bulkRow = document.createElement('div'); bulkRow.className = 'import-issue-row';
      bulkRow.innerHTML = `<span class="label">전체 일괄 연도 적용</span><input type="number" id="importBulkYear" placeholder="예: 2026" min="1990" max="2100" /><button class="btn btn-ghost" type="button" id="importBulkYearBtn">전체 적용</button>`;
      yearBox.appendChild(bulkRow);
      yearItems.forEach((item, idx) => {
        const row = document.createElement('div'); row.className = 'import-issue-row';
        row.innerHTML = `<span class="label">${escapeHtml(item.label)} — 원본: "${escapeHtml(item.original)}"</span><input type="number" class="import-year-input" data-idx="${idx}" placeholder="연도" min="1990" max="2100" /><label style="font-size:11px; display:flex; align-items:center; gap:4px;"><input type="checkbox" class="import-keep-as-is" data-idx="${idx}" /> 원본 유지</label>`;
        yearBox.appendChild(row);
      });
      body.appendChild(yearBox);
      document.getElementById('importBulkYearBtn').onclick = () => {
        const y = document.getElementById('importBulkYear').value;
        if (!y) return;
        body.querySelectorAll('.import-year-input').forEach((inp) => { inp.value = y; });
      };
    }
    const previewWrap = document.createElement('div'); previewWrap.className = 'import-table-wrap';
    const table = document.createElement('table');
    const editableCols = COLUMNS.filter((c) => c.type !== 'image');
    table.innerHTML = `<thead><tr>${editableCols.map((c) => `<th>${escapeHtml(c.label)}</th>`).join('')}</tr></thead>`;
    const tbody = document.createElement('tbody');
    parsed.slice(0, 15).forEach((p) => {
      const tr = document.createElement('tr');
      tr.innerHTML = editableCols.map((c) => `<td>${escapeHtml(p[c.key] || '')}</td>`).join('');
      tbody.appendChild(tr);
    });
    table.appendChild(tbody); previewWrap.appendChild(table);
    if (parsed.length > 15) {
      const more = document.createElement('div'); more.className = 'field-hint'; more.style.padding = '6px 8px';
      more.textContent = `외 ${parsed.length - 15}행...`;
      previewWrap.appendChild(more);
    }
    body.appendChild(previewWrap);
    overlay.classList.add('open');
    function finish(result) {
      overlay.classList.remove('open');
      document.getElementById('importCancelBtn').onclick = null;
      document.getElementById('importModalClose').onclick = null;
      document.getElementById('importConfirmBtn').onclick = null;
      resolve(result);
    }
    document.getElementById('importCancelBtn').onclick = () => finish(false);
    document.getElementById('importModalClose').onclick = () => finish(false);
    document.getElementById('importConfirmBtn').onclick = () => {
      for (let idx = 0; idx < yearItems.length; idx++) {
        const yearInput = body.querySelector(`.import-year-input[data-idx="${idx}"]`);
        const keepAsIs = body.querySelector(`.import-keep-as-is[data-idx="${idx}"]`);
        if (keepAsIs.checked) { yearItems[idx].apply(yearItems[idx].original); }
        else if (yearInput.value) { yearItems[idx].apply(applyYearToDate(yearItems[idx].original, yearInput.value)); }
        else { showToast('연도 미확정 항목이 있습니다. 입력하거나 "원본 유지"를 선택하세요.', true); return; }
      }
      finish(true);
    };
  });
}

// ---------------------------------------------------------------------------
// Export (plain browser download — no sandbox restrictions on a real site)
// ---------------------------------------------------------------------------
function buildExportRows() {
  return state.cells.map((c) => {
    const row = {};
    COLUMNS.forEach((col) => {
      row[col.label] = col.type === 'image' ? (c.images || []).map((im) => im.originalName || im.path).join('; ') : (c[col.key] || '');
    });
    return row;
  });
}
function dateStamp() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}
function downloadBlob(content, filename, type) {
  const blob = new Blob([type.includes('csv') ? '\uFEFF' + content : content], { type });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function exportCsv() {
  const ws = XLSX.utils.json_to_sheet(buildExportRows(), { header: COLUMNS.map((c) => c.label) });
  const csv = XLSX.utils.sheet_to_csv(ws);
  downloadBlob(csv, `${SHEET.exportName}-${dateStamp()}.csv`, 'text/csv;charset=utf-8;');
}
function exportXlsx() {
  const ws = XLSX.utils.json_to_sheet(buildExportRows(), { header: COLUMNS.map((c) => c.label) });
  const wbx = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbx, ws, SHEET.sheetName);
  XLSX.writeFile(wbx, `${SHEET.exportName}-${dateStamp()}.xlsx`);
}

// ---------------------------------------------------------------------------
// Filters
// ---------------------------------------------------------------------------
function uniqueValues(key) { return [...new Set(state.cells.map((c) => c[key]).filter((v) => v && v.trim()))].sort(); }
function fillSelect(sel, values) {
  const current = sel.value;
  const firstOption = sel.options[0];
  sel.innerHTML = ''; sel.appendChild(firstOption);
  values.forEach((v) => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
  if (values.includes(current)) sel.value = current;
}
function populateFilterOptions() {
  document.querySelectorAll('#filterSelects select').forEach((sel) => fillSelect(sel, uniqueValues(sel.dataset.key)));
}
function buildFilterControls() {
  const wrap = document.getElementById('filterSelects');
  wrap.innerHTML = '';
  SHEET.filters.forEach(([key, label]) => {
    const sel = document.createElement('select');
    sel.dataset.key = key;
    sel.innerHTML = `<option value="">${escapeHtml(label)} (전체)</option>`;
    sel.onchange = () => { state.filters.sel[key] = sel.value; renderTable(); };
    wrap.appendChild(sel);
  });
  document.getElementById('filterDate').hidden = !SHEET.dateKey;
}

// ---------------------------------------------------------------------------
// Sheet tabs
// ---------------------------------------------------------------------------
function renderSheetTabs() {
  const nav = document.getElementById('sheetTabs');
  nav.innerHTML = '';
  Object.entries(SHEETS).forEach(([key, sheet]) => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'sheet-tab' + (key === activeSheetKey ? ' active' : '');
    btn.innerHTML = `${escapeHtml(sheet.label)}<span class="count">${state.loaded ? state.data[key].length : '…'}</span>`;
    btn.onclick = () => setActiveSheet(key);
    nav.appendChild(btn);
  });
}
function setActiveSheet(key) {
  if (!SHEETS[key]) return;
  activeSheetKey = key;
  SHEET = SHEETS[key];
  COLUMNS = SHEET.columns;
  EDITABLE_KEYS = SHEET.editableKeys;
  HEADER_MAP = SHEET.headerMap;
  try { localStorage.setItem('btt.activeSheet', key); } catch (e) {}
  state.cells = state.data[key];
  state.filters = { search: '', date: '', sel: {} };
  document.getElementById('searchInput').value = '';
  document.getElementById('filterDate').value = '';
  buildFilterControls();
  populateFilterOptions();
  renderSheetTabs();
  renderStats();
  renderTable();
}

// ---------------------------------------------------------------------------
// Wiring
// ---------------------------------------------------------------------------
function wireModalClose(overlayId, closeBtnIds) {
  const overlay = document.getElementById(overlayId);
  closeBtnIds.forEach((id) => { const el = document.getElementById(id); if (el) el.onclick = () => overlay.classList.remove('open'); });
  overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.classList.remove('open'); });
}
function init() {
  document.getElementById('tableBody').addEventListener('change', handleCellChange);
  document.getElementById('tableBody').addEventListener('paste', onTablePaste, true);
  document.getElementById('tableBody').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey && !e.isComposing && e.target.classList && e.target.classList.contains('cell-input')) { e.preventDefault(); e.target.blur(); }
  });
  document.getElementById('tableBody').addEventListener('click', (e) => {
    const tr = e.target.closest('tr');
    if (!tr) return;
    if (e.target.closest('input, textarea, button, img')) return;
    const cell = state.cells.find((c) => c.id === tr.dataset.id);
    if (cell) openDetailModal(cell);
  });

  document.getElementById('newRecordBtn').onclick = () => openFormModal(null);
  document.getElementById('addRowBtn').onclick = async () => {
    try { await createCell({}); showToast('빈 행이 추가되었습니다.'); }
    catch (err) { console.error(err); showToast('행 추가 실패', true); }
  };
  document.getElementById('formSaveBtn').onclick = saveFormModal;
  wireModalClose('formModalOverlay', ['formModalClose', 'formCancelBtn']);
  wireModalClose('detailModalOverlay', ['detailModalClose', 'detailCloseBtn']);
  wireModalClose('lightboxOverlay', ['lightboxClose']);

  document.getElementById('detailDeleteBtn').onclick = async () => {
    const cell = state.cells.find((c) => c.id === state.detailId);
    if (!cell) return;
    if (!confirm(`Cell ID "${cell.cellId || '(없음)'}" 행을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.`)) return;
    try {
      await deleteCell(cell);
      document.getElementById('detailModalOverlay').classList.remove('open');
      showToast('행이 삭제되었습니다.');
    } catch (err) { console.error(err); showToast('삭제 중 오류가 발생했습니다.', true); }
  };
  document.getElementById('detailEditBtn').onclick = () => {
    const cell = state.cells.find((c) => c.id === state.detailId);
    document.getElementById('detailModalOverlay').classList.remove('open');
    if (cell) openFormModal(cell);
  };

  document.getElementById('lightboxPrev').onclick = () => lightboxStep(-1);
  document.getElementById('lightboxNext').onclick = () => lightboxStep(1);
  document.addEventListener('keydown', (e) => {
    if (!document.getElementById('lightboxOverlay').classList.contains('open')) return;
    if (e.key === 'ArrowLeft') lightboxStep(-1);
    if (e.key === 'ArrowRight') lightboxStep(1);
    if (e.key === 'Escape') document.getElementById('lightboxOverlay').classList.remove('open');
  });

  document.getElementById('searchInput').oninput = debounce((e) => { state.filters.search = e.target.value; renderTable(); }, 200);
  document.getElementById('filterDate').oninput = debounce((e) => { state.filters.date = e.target.value; renderTable(); }, 200);
  document.getElementById('clearFiltersBtn').onclick = () => {
    state.filters = { search: '', date: '', sel: {} };
    document.getElementById('searchInput').value = '';
    document.querySelectorAll('#filterSelects select').forEach((sel) => { sel.value = ''; });
    document.getElementById('filterDate').value = '';
    renderTable();
  };

  document.getElementById('importInput').addEventListener('change', async (e) => {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    try { await handleImportFile(file); }
    catch (err) { console.error(err); showToast('가져오기 처리 중 오류가 발생했습니다.', true); }
  });
  document.getElementById('exportCsvBtn').onclick = exportCsv;
  document.getElementById('exportXlsxBtn').onclick = exportXlsx;

  let savedSheet = null;
  try { savedSheet = localStorage.getItem('btt.activeSheet'); } catch (e) {}
  setActiveSheet(SHEETS[savedSheet] ? savedSheet : 'teardown');
  initStorage();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

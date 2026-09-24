'use strict';

const M = window.LvReportModel;
const P = window.LvReportPpt;
const $ = (id) => document.getElementById(id);
const COLS = ['id', 'x', 'y', 'topBack', 'layer', 'element', 'lot'];

const state = {
  fileName: '',
  sheets: [],        // [{ name, header: { row, x, y }, rowCount, xyCount }]
  sheet: '',
  rows: [],          // raw rows of the chosen sheet
  headerRow: 0,
  map: {},           // column key -> column index (or -1)
  points: [],        // [{ id, x, y, topBack, layer, element, lot }]
  colors: {},        // element -> hex
  hidden: new Set(), // elements switched off
  view: 'cumulative',
};

// ---------------------------------------------------------------------------
// Worker
// ---------------------------------------------------------------------------
const worker = new Worker('worker.js?v=3');
let seq = 0;
const pending = new Map();
worker.onmessage = (e) => {
  const p = pending.get(e.data.id);
  if (!p) return;
  pending.delete(e.data.id);
  e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error));
};
worker.onerror = (e) => { console.error(e); pending.forEach((p) => p.reject(new Error('The file reader failed to start. Check your internet connection and reload.'))); pending.clear(); };
function callWorker(msg, transfer) {
  return new Promise((resolve, reject) => { const id = ++seq; pending.set(id, { resolve, reject }); worker.postMessage(Object.assign({ id }, msg), transfer || []); });
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
let toastTimer;
function showToast(msg, isError) {
  const t = $('toast'); t.textContent = msg; t.className = 'toast show' + (isError ? ' error' : '');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 4000);
}
function busy(text) { $('busyText').textContent = text || 'Working…'; $('busy').hidden = !text; }
function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function norm(v) { return String(v == null ? '' : v).toLowerCase().replace(/[\s\r\n]+/g, ''); }
function numOf(v) { const t = String(v == null ? '' : v).replace(/mm/i, '').replace(/,/g, '').trim(); const n = Number(t); return t !== '' && isFinite(n) ? n : null; }
function today() { const d = new Date(); return `${String(d.getFullYear()).slice(2)}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`; }
function dims() {
  return { xMax: Math.max(1, Number($('dimW').value) || 300), yMax: Math.max(1, Number($('dimH').value) || 210), layers: Math.max(2, Number($('dimL').value) || 37) };
}
const colorOf = (el) => (el ? state.colors[el] || 'A6A6A6' : 'FFFFFF');

// ---------------------------------------------------------------------------
// Column detection
// ---------------------------------------------------------------------------
const PATTERNS = {
  id: [/^cellid$/, /cell_?id/, /baseid/, /셀id|셀아이디/, /^id$/],
  x: [/^x(\(mm\)|mm|좌표)?$/],
  y: [/^y(\(mm\)|mm|좌표)?$/],
  topBack: [/top\/?back/, /topback/, /^(top|back)$/, /면$/],
  layer: [/anodesheet/, /droppedlayer/, /^layer/, /층/],
  lot: [/^lotid$/, /^lot$/, /pkglot/, /^lot/],
  element: [/conclusion/, /edsimpurity/, /sem\/?eds/, /foreign|impurity|이물|element|contaminant/, /^ntf$/],
};
function detectColumns(header, dataRows) {
  const h = header.map(norm);
  const map = {};
  COLS.forEach((k) => {
    map[k] = -1;
    if (k === 'element') return;
    for (const re of PATTERNS[k]) { const i = h.findIndex((v) => re.test(v)); if (i >= 0) { map[k] = i; break; } }
  });
  // Foreign material: of the matching columns, take the one with the most element-like values
  let best = -1, bestScore = 0;
  h.forEach((v, i) => {
    if (!PATTERNS.element.some((re) => re.test(v))) return;
    const score = dataRows.filter((r) => M.normElement(r[i]) && String(r[i]).length < 25).length;
    if (score > bestScore) { best = i; bestScore = score; }
  });
  map.element = best;
  return map;
}

// ---------------------------------------------------------------------------
// Loading
// ---------------------------------------------------------------------------
async function loadFile(file) {
  if (!file) return;
  busy(`Reading "${file.name}"…`);
  $('fileStatus').textContent = '';
  try {
    const buf = await file.arrayBuffer();
    const res = await callWorker({ type: 'book', buf }, [buf]);
    const withXY = res.sheets.filter((s) => s.header);
    if (!withXY.length) throw new Error('No sheet with X and Y columns was found.');
    state.fileName = file.name;
    state.sheets = res.sheets;
    const sel = $('sheetSel');
    sel.innerHTML = withXY.map((s) => `<option value="${esc(s.name)}">${esc(s.name)} — ${s.xyCount} row(s) with X/Y</option>`).join('');
    const best = withXY.slice().sort((a, b) => b.xyCount - a.xyCount)[0];
    sel.value = best.name;
    $('fileStatus').innerHTML = `✓ <b>${esc(file.name)}</b> — ${res.sheets.length} sheet(s), ${withXY.length} with X/Y columns`;
    $('pptTitle').value = defaultTitle(file.name, best.name);
    await loadSheet(best.name);
  } catch (err) {
    console.error(err);
    $('fileStatus').innerHTML = `<span class="err">Could not read this file: ${esc(err.message)}</span>`;
  } finally {
    busy(null);
  }
}
function defaultTitle(fileName, sheet) {
  const model = (/(E\d{2}[A-Z])/i.exec(sheet) || /(E\d{2}[A-Z])/i.exec(fileName) || [])[1];
  const sdr = /sdr/i.test(fileName) ? ' SDR' : '';
  return `${model ? model.toUpperCase() : 'E81C'}${sdr} 저전압`;
}
async function loadSheet(name) {
  busy(`Reading sheet "${name}"…`);
  try {
    const info = state.sheets.find((s) => s.name === name);
    const res = await callWorker({ type: 'bookRows', sheet: name });
    state.sheet = name;
    state.rows = res.rows;
    state.headerRow = info.header.row;
    const header = res.rows[state.headerRow] || [];
    const dataRows = res.rows.slice(state.headerRow + 1);
    state.map = detectColumns(header, dataRows);
    state.map.x = info.header.x;
    state.map.y = info.header.y;
    fillColumnSelects(header);
    if (/E69B/i.test(name)) $('preset').value = 'custom';
    buildPoints();
    $('mapping').hidden = false;
    $('settingsCard').hidden = false;
    $('viewCard').hidden = false;
  } finally {
    busy(null);
  }
}
function fillColumnSelects(header) {
  document.querySelectorAll('select[data-col]').forEach((sel) => {
    const k = sel.dataset.col;
    const required = k === 'x' || k === 'y';
    sel.innerHTML = (required ? '' : '<option value="-1">(none)</option>') +
      header.map((h, i) => `<option value="${i}">${esc(String(h).replace(/\s+/g, ' ').trim() || `Column ${i + 1}`)}</option>`).join('');
    sel.value = String(state.map[k]);
  });
}

function buildPoints() {
  const m = state.map;
  const get = (r, k) => (m[k] >= 0 ? r[m[k]] : '');
  const pts = [];
  let skipped = 0;
  state.rows.slice(state.headerRow + 1).forEach((r, i) => {
    const x = numOf(get(r, 'x')), y = numOf(get(r, 'y'));
    if (x === null || y === null) { if (String(get(r, 'x')).trim() || String(get(r, 'y')).trim()) skipped++; return; }
    pts.push({
      id: String(get(r, 'id') || `Row ${state.headerRow + i + 2}`).trim(),
      x, y,
      topBack: M.normTopBack(get(r, 'topBack')),
      layer: numOf(get(r, 'layer')),
      element: M.normElement(get(r, 'element')),
      lot: String(get(r, 'lot') || '').trim(),
    });
  });
  state.points = pts;
  // Stable colors: known elements use the report palette, others take the next free color
  const counts = M.countBy(pts, (p) => p.element);
  state.colors = {};
  counts.forEach(([el], i) => { state.colors[el] = M.ELEMENT_COLORS[el] || M.elementColor(el, i); });
  state.hidden = new Set([...state.hidden].filter((el) => state.colors[el]));
  const cells = new Set(pts.map((p) => p.id)).size;
  const d = dims();
  const out = pts.filter((p) => p.x > d.xMax || p.y > d.yMax || p.x < 0 || p.y < 0).length;
  const noLayer = pts.filter((p) => p.layer == null).length;
  $('pointStatus').innerHTML = `${pts.length} point(s) from ${cells} cell(s)` +
    (skipped ? ` · ${skipped} row(s) skipped (X/Y not a number)` : '') +
    (out ? ` · <b>${out} point(s) outside the ${d.xMax} × ${d.yMax} mm cell</b> (drawn at the edge — check the cell size)` : '') +
    (m.layer < 0 ? ' · no layer column: the side view needs one' : noLayer ? ` · ${noLayer} point(s) without a layer (not on the side view)` : '');
  renderChips();
  render();
}

// ---------------------------------------------------------------------------
// Filtering
// ---------------------------------------------------------------------------
function filtered() {
  const tb = $('fTB').value;
  const q = $('fSearch').value.trim().toLowerCase();
  return state.points.filter((p) =>
    (tb === 'all' || p.topBack === tb) &&
    !state.hidden.has(p.element) &&
    (!q || p.id.toLowerCase().includes(q) || p.lot.toLowerCase().includes(q)));
}
function renderChips() {
  const counts = M.countBy(state.points, (p) => p.element || '(none)');
  $('elementChips').innerHTML = counts.map(([el, n]) => {
    const key = el === '(none)' ? '' : el;
    return `<span class="chip ${state.hidden.has(key) ? 'off' : ''}" data-el="${esc(key)}"><i style="background:#${colorOf(key)}"></i>${esc(el)} <b>${n}</b></span>`;
  }).join('');
  $('elementChips').querySelectorAll('.chip').forEach((c) => {
    c.onclick = () => { const el = c.dataset.el; state.hidden.has(el) ? state.hidden.delete(el) : state.hidden.add(el); renderChips(); render(); };
  });
}

// ---------------------------------------------------------------------------
// SVG drawing — same cell shape as the PowerPoint (tab, grey electrode, orange tab; side view stripes)
// ---------------------------------------------------------------------------
function cellSvg(points, view, opts) {
  const o = Object.assign({ w: 520, h: 220, dot: 7 }, opts);
  const d = dims();
  const tabW = o.w * 0.05;
  const bx = tabW, bw = o.w - tabW * 2, by = 0, bh = o.h;
  // Padding so dots on the very edge of the cell stay fully visible
  const pad = o.dot + 1;
  let s = `<svg viewBox="${-pad} ${-pad} ${o.w + pad * 2} ${o.h + pad * 2}" xmlns="http://www.w3.org/2000/svg" role="img">`;
  if (view === 'side') {
    for (let i = 0; i < 18; i++) {
      const yy = (i + 0.5) * o.h / 18;
      s += `<line x1="${bx}" x2="${bx + bw}" y1="${yy}" y2="${yy}" stroke="#${i % 2 ? 'A6A6A6' : '7F7F7F'}" stroke-width="1.6"/>`;
    }
  } else {
    s += `<rect x="0" y="${o.h * 0.3}" width="${tabW}" height="${o.h * 0.4}" fill="#F2F2F2" stroke="#D9D9D9"/>`;
    s += `<rect x="${bx}" y="${by}" width="${bw}" height="${bh}" rx="${o.h * 0.04}" fill="#A6A6A6" stroke="#8C8C8C"/>`;
    s += `<rect x="${bx + bw}" y="${o.h * 0.3}" width="${tabW}" height="${o.h * 0.4}" fill="#F8CBAD"/>`;
  }
  points.forEach((p) => {
    const yv = view === 'side' ? p.layer : p.y;
    if (yv == null) return;
    const fx = Math.max(0, Math.min(1, p.x / d.xMax));
    const fy = view === 'side' ? Math.max(0, Math.min(1, (yv - 1) / (d.layers - 1))) : Math.max(0, Math.min(1, yv / d.yMax));
    const cx = bx + fx * bw, cy = by + fy * bh;
    const tip = `${p.id}${p.lot ? ' · ' + p.lot : ''} · ${p.element || 'no material'} · X ${p.x}, Y ${p.y}${p.layer != null ? ' · layer ' + p.layer : ''}${p.topBack ? ' · ' + p.topBack : ''}`;
    s += `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="${o.dot}" fill="#${colorOf(p.element)}" stroke="#262626" stroke-width="0.8"><title>${esc(tip)}</title></circle>`;
  });
  return s + '</svg>';
}
function legendHtml(points) {
  return `<div class="legend">${M.countBy(points, (p) => p.element || '(none)').map(([el, n]) => `<span><i style="background:#${colorOf(el === '(none)' ? '' : el)}"></i>${esc(el)} (${n})</span>`).join('')}</div>`;
}
function pairPanel(title, points) {
  return `<div class="panel"><div class="panel-title">${esc(title)} — ${points.length} point(s)</div><div class="panel-body">` +
    `<div class="map-pair"><div class="map-box"><div class="cap">Top view (X, Y)</div>${cellSvg(points, 'plane')}</div>` +
    `<div class="map-box"><div class="cap">Side view (X, layer)</div>${cellSvg(points, 'side')}</div></div>${legendHtml(points)}</div></div>`;
}

// ---------------------------------------------------------------------------
// Views
// ---------------------------------------------------------------------------
function render() {
  const pts = filtered();
  $('cumulativeView').hidden = state.view !== 'cumulative';
  $('individualView').hidden = state.view !== 'individual';
  if (state.view === 'cumulative') renderCumulative(pts); else renderIndividual(pts);
}
function renderCumulative(pts) {
  if (!pts.length) { $('cumulativeView').innerHTML = '<div class="empty-note">No points match the filters.</div>'; return; }
  let html = pairPanel('[이물 위치] All', pts);
  const top = pts.filter((p) => p.topBack === 'Top'), back = pts.filter((p) => p.topBack === 'Back');
  if (top.length) html += pairPanel('[이물 위치 Top View]', top);
  if (back.length) html += pairPanel('[이물 위치 Back View]', back);
  // Small multiples per foreign material (like the report's element slide)
  const els = M.countBy(pts, (p) => p.element || '(none)');
  html += `<div class="panel"><div class="panel-title">[이물별 위치]</div><div class="panel-body"><div class="multiples">` +
    els.map(([el]) => {
      const sub = pts.filter((p) => (p.element || '(none)') === el);
      return `<div class="map-box"><div class="cap">${esc(el)} (${sub.length})</div>${cellSvg(sub, 'plane', { w: 460, h: 180, dot: 8 })}${cellSvg(sub, 'side', { w: 460, h: 150, dot: 8 })}</div>`;
    }).join('') + '</div></div></div>';
  $('cumulativeView').innerHTML = html;
}
function groupByCell(pts) {
  const map = new Map();
  pts.forEach((p) => { if (!map.has(p.id)) map.set(p.id, []); map.get(p.id).push(p); });
  return [...map.entries()];
}
function cellMeta(list) {
  const p = list[0];
  const els = [...new Set(list.map((x) => x.element).filter(Boolean))];
  return [p.lot && `<span>${esc(p.lot)}</span>`, p.topBack && `<span>${esc(p.topBack)}</span>`,
    ...els.map((el) => `<span class="el"><i style="background:#${colorOf(el)}"></i>${esc(el)}</span>`),
    list.some((x) => x.layer != null) && `<span>layer ${esc([...new Set(list.map((x) => x.layer).filter((v) => v != null))].join(', '))}</span>`,
    list.length > 1 && `<span>${list.length} points</span>`].filter(Boolean).join('');
}
function renderIndividual(pts) {
  const cells = groupByCell(pts);
  if (!cells.length) { $('individualView').innerHTML = '<div class="empty-note">No cells match the filters.</div>'; return; }
  $('individualView').innerHTML = `<div class="cell-grid">${cells.map(([id, list]) => `<div class="cell-card" data-id="${esc(id)}"><div class="cell-head"><b>${esc(id)}</b><div class="cell-meta">${cellMeta(list)}</div></div>` +
    `<div class="cell-pair"><div class="map-box"><div class="cap">Top view</div>${cellSvg(list, 'plane', { w: 300, h: 130, dot: 7 })}</div>` +
    `<div class="map-box"><div class="cap">Side view</div>${cellSvg(list, 'side', { w: 300, h: 130, dot: 7 })}</div></div></div>`).join('')}</div>`;
  $('individualView').querySelectorAll('.cell-card').forEach((c) => {
    c.onclick = () => {
      const list = cells.find(([id]) => id === c.dataset.id)[1];
      $('zoomBody').innerHTML = `<div class="cell-head"><b>${esc(c.dataset.id)}</b><div class="cell-meta">${cellMeta(list)}</div></div>` +
        `<div class="cell-pair"><div class="map-box"><div class="cap">Top view (X ${dims().xMax} mm × Y ${dims().yMax} mm)</div>${cellSvg(list, 'plane', { w: 600, h: 260, dot: 9 })}</div>` +
        `<div class="map-box"><div class="cap">Side view (${dims().layers} layers)</div>${cellSvg(list, 'side', { w: 600, h: 260, dot: 9 })}</div></div>` +
        `<p class="hint">${list.map((p) => `X ${p.x}, Y ${p.y}${p.layer != null ? `, layer ${p.layer}` : ''}${p.element ? ` · ${esc(p.element)}` : ''}`).join('<br/>')}</p>` +
        '<div class="modal-actions"><button class="btn btn-ghost" type="button" id="zoomClose">Close</button></div>';
      $('zoomOverlay').classList.add('open');
      $('zoomClose').onclick = () => $('zoomOverlay').classList.remove('open');
    };
  });
}

// ---------------------------------------------------------------------------
// PowerPoint export (same drawing as the weekly report)
// ---------------------------------------------------------------------------
async function exportPpt() {
  const pts = filtered();
  if (!pts.length) { showToast('No points to export.', true); return; }
  const H = P.helpers;
  const pres = new window.PptxGenJS();
  pres.defineLayout({ name: 'A4R', width: P.W, height: 7.5 });
  pres.layout = 'A4R';
  const meta = { team: $('pptTeam').value.trim() || 'ESHG 품질팀', date: $('pptDate').value.trim() };
  const title = $('pptTitle').value.trim() || 'E81C 저전압';
  const D = Object.assign(dims(), { colorFn: (el) => colorOf(el) });
  const legendItems = (list) => M.countBy(list, (p) => p.element || '(none)').map(([el, n]) => ({ label: `${el} (${n})`, color: colorOf(el === '(none)' ? '' : el) }));

  // Slide 1: cumulative — all / top / back rows, each with top view + side view
  const s1 = pres.addSlide();
  H.chrome(s1, pres, `${title} – 이물 발견 위치 (누적)`, meta);
  const rows = [['[이물 위치] All', pts], ['[이물 위치 Top View]', pts.filter((p) => p.topBack === 'Top')], ['[이물 위치 Back View]', pts.filter((p) => p.topBack === 'Back')]].filter(([, l]) => l.length);
  rows.forEach(([label, list], i) => {
    const y = 0.8 + i * 2.2;
    H.panelTitle(s1, pres, `${label} — ${list.length} point(s)`, 0.02, y, 9.45);
    H.electrodeMap(s1, pres, list, { x: 0.25, y: y + 0.35, w: 4.4, h: 1.7 }, 'plane', D);
    H.electrodeMap(s1, pres, list, { x: 4.95, y: y + 0.35, w: 4.4, h: 1.7 }, 'side', D);
    H.txt(s1, 'Top view', { x: 0.25, y: y + 2.05, w: 4.4, h: 0.14, fontSize: 7, align: 'center', color: '595959' });
    H.txt(s1, 'Side view', { x: 4.95, y: y + 2.05, w: 4.4, h: 0.14, fontSize: 7, align: 'center', color: '595959' });
  });
  H.legend(s1, legendItems(pts), 9.6, 0.9, { size: 7, step: 0.17, max: 30 });

  // Slide 2: per foreign material (top view and side view per element)
  const els = M.countBy(pts, (p) => p.element || '(none)').map(([el]) => el);
  const perPage = 8;
  for (let page = 0; page * perPage < els.length; page++) {
    const s = pres.addSlide();
    H.chrome(s, pres, `${title} – 이물별 발견 위치${els.length > perPage ? ` (${page + 1})` : ''}`, meta);
    H.panelTitle(s, pres, '[이물 위치 Top View]', 0.02, 0.8, 5.3);
    H.panelTitle(s, pres, '[이물 측면 위치]', 5.5, 0.8, 5.3);
    els.slice(page * perPage, page * perPage + perPage).forEach((el, i) => {
      const list = pts.filter((p) => (p.element || '(none)') === el);
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.05 + col * 2.65, y = 1.1 + row * 1.55;
      const label = `${el} (${list.length})`;
      H.txt(s, label, { x, y, w: 2.55, h: 0.15, fontSize: 7.5, align: 'center', color: '404040' });
      H.electrodeMap(s, pres, list, { x: x + 0.05, y: y + 0.17, w: 2.45, h: 1.25 }, 'plane', D);
      H.txt(s, label, { x: x + 5.48, y, w: 2.55, h: 0.15, fontSize: 7.5, align: 'center', color: '404040' });
      H.electrodeMap(s, pres, list, { x: x + 5.53, y: y + 0.17, w: 2.45, h: 1.25 }, 'side', D);
    });
  }

  // Individual cells: 6 per slide (2 columns x 3 rows), each with top view + side view
  const cells = groupByCell(pts);
  for (let page = 0; page * 6 < cells.length; page++) {
    const s = pres.addSlide();
    H.chrome(s, pres, `${title} – 셀별 이물 위치 (${page + 1}/${Math.ceil(cells.length / 6)})`, meta);
    cells.slice(page * 6, page * 6 + 6).forEach(([id, list], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = 0.05 + col * 5.4, y = 0.8 + row * 2.2;
      const p = list[0];
      const info = [p.lot, p.topBack, [...new Set(list.map((q) => q.element).filter(Boolean))].join(', '), list.some((q) => q.layer != null) && `layer ${[...new Set(list.map((q) => q.layer).filter((v) => v != null))].join(', ')}`].filter(Boolean).join(' · ');
      H.panelTitle(s, pres, `${id}${info ? '  |  ' + info : ''}`, x, y, 5.3);
      H.electrodeMap(s, pres, list, { x: x + 0.1, y: y + 0.4, w: 2.5, h: 1.45 }, 'plane', D);
      H.electrodeMap(s, pres, list, { x: x + 2.75, y: y + 0.4, w: 2.5, h: 1.45 }, 'side', D);
      H.txt(s, `Top view  X ${list.map((q) => q.x).join('/')}, Y ${list.map((q) => q.y).join('/')}`, { x: x + 0.1, y: y + 1.88, w: 2.5, h: 0.14, fontSize: 6.5, align: 'center', color: '595959' });
      H.txt(s, 'Side view', { x: x + 2.75, y: y + 1.88, w: 2.5, h: 0.14, fontSize: 6.5, align: 'center', color: '595959' });
    });
  }
  const stamp = (meta.date || today()).replace(/[^0-9]/g, '').slice(-6);
  const fileName = `${stamp}_${title} 이물 위치 (${state.sheet}).pptx`.replace(/[\\/:*?"<>|]/g, '-');
  await pres.writeFile({ fileName });
  showToast(`Downloaded ${fileName}`);
}

// ---------------------------------------------------------------------------
// Reset & wiring
// ---------------------------------------------------------------------------
function doReset() {
  Object.assign(state, { fileName: '', sheets: [], sheet: '', rows: [], headerRow: 0, map: {}, points: [], colors: {}, hidden: new Set(), view: 'cumulative' });
  $('fileInput').value = '';
  $('fileStatus').textContent = '';
  $('pointStatus').textContent = '';
  $('mapping').hidden = true; $('settingsCard').hidden = true; $('viewCard').hidden = true;
  $('preset').value = 'E81C'; $('dimW').value = 300; $('dimH').value = 210; $('dimL').value = 37;
  $('fTB').value = 'all'; $('fSearch').value = ''; $('pptTitle').value = ''; $('pptDate').value = today();
  document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.view === 'cumulative'));
  $('resetOverlay').classList.remove('open');
  showToast('Reset complete.');
}
function init() {
  $('pptDate').value = today();
  const drop = $('drop');
  $('fileInput').addEventListener('change', (e) => loadFile(e.target.files[0]));
  ['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('drag'); }));
  drop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) loadFile(f); });
  $('sheetSel').addEventListener('change', (e) => loadSheet(e.target.value));
  document.querySelectorAll('select[data-col]').forEach((sel) => sel.addEventListener('change', () => { state.map[sel.dataset.col] = Number(sel.value); buildPoints(); }));
  $('preset').addEventListener('change', () => { if ($('preset').value === 'E81C') { $('dimW').value = 300; $('dimH').value = 210; $('dimL').value = 37; } buildPoints(); });
  ['dimW', 'dimH', 'dimL'].forEach((id) => $(id).addEventListener('input', () => { $('preset').value = 'custom'; buildPoints(); }));
  $('fTB').addEventListener('change', render);
  $('fSearch').addEventListener('input', render);
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => {
    state.view = t.dataset.view;
    document.querySelectorAll('.tab').forEach((x) => x.classList.toggle('active', x === t));
    render();
  }));
  $('pptBtn').onclick = () => exportPpt().catch((err) => { console.error(err); showToast('Could not create the PowerPoint: ' + err.message, true); });
  $('resetBtn').onclick = () => $('resetOverlay').classList.add('open');
  $('resetCancelBtn').onclick = () => $('resetOverlay').classList.remove('open');
  $('resetOverlay').addEventListener('click', (e) => { if (e.target === $('resetOverlay')) $('resetOverlay').classList.remove('open'); });
  $('zoomOverlay').addEventListener('click', (e) => { if (e.target === $('zoomOverlay')) $('zoomOverlay').classList.remove('open'); });
  $('resetConfirmBtn').onclick = doReset;
}
init();

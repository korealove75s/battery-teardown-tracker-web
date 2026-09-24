'use strict';

const Cv = window.OcvConvert;
const M = window.LvReportModel;
const P = window.LvReportPpt;
const $ = (id) => document.getElementById(id);
const NOTES_KEY = 'lvreport.notes.v1';

const data = {
  ocv: null,        // { fileName, cells, lots, masterName }
  review: null,     // { fileName, byId }
  stats: null,      // { fileName, rows }
  genealogy: null,  // { fileName, byId }
};

// ---------------------------------------------------------------------------
// Worker (shared with the tracker)
// ---------------------------------------------------------------------------
const worker = new Worker('worker.js?v=2');
let seq = 0;
const pending = new Map();
worker.onmessage = (e) => {
  const p = pending.get(e.data.id);
  if (!p) return;
  if (e.data.progress != null && e.data.ok === undefined) { if (p.onProgress) p.onProgress(e.data.progress); return; }
  pending.delete(e.data.id);
  e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error));
};
worker.onerror = (e) => {
  console.error(e);
  pending.forEach((p) => p.reject(new Error('The file reader failed to start. Check your internet connection and reload.')));
  pending.clear();
};
function callWorker(msg, transfer, onProgress) {
  return new Promise((resolve, reject) => {
    const id = ++seq;
    pending.set(id, { resolve, reject, onProgress });
    worker.postMessage(Object.assign({ id }, msg), transfer || []);
  });
}

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
let toastTimer;
function showToast(msg, isError) {
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show' + (isError ? ' error' : '');
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => t.classList.remove('show'), 4000);
}
function busy(text) { $('busyText').textContent = text || 'Working…'; $('busy').hidden = !text; }
function escapeHtml(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function card(key) { return document.querySelector(`.file-card[data-key="${key}"]`); }
function setFileStatus(key, html, state) {
  const c = card(key);
  c.classList.toggle('loaded', state === 'ok');
  c.classList.toggle('error', state === 'error');
  c.querySelector('.file-status').innerHTML = html;
}
function today() {
  const d = new Date();
  return `${String(d.getFullYear()).slice(2)}.${String(d.getMonth() + 1).padStart(2, '0')}.${String(d.getDate()).padStart(2, '0')}`;
}

// Improvement notes <-> textarea ("LOT" line, then text lines; blank line between notes)
function notesToText(notes) { return notes.map((n) => `${n.lot}\n${n.text}`).join('\n\n'); }
function textToNotes(t) {
  return String(t || '').split(/\n\s*\n/).map((block) => block.split('\n').map((l) => l.trimEnd()).filter((l) => l.trim()))
    .filter((lines) => lines.length >= 2).map((lines) => ({ lot: lines[0].trim(), text: lines.slice(1).join('\n') }));
}
function loadNotes() {
  try { const v = localStorage.getItem(NOTES_KEY); if (v != null) return v; } catch (e) {}
  return notesToText(P.DEFAULT_NOTES);
}

// ---------------------------------------------------------------------------
// Loading files
// ---------------------------------------------------------------------------
async function loadFile(key, file) {
  if (!file) return;
  setFileStatus(key, `Reading ${escapeHtml(file.name)}…`, '');
  try {
    if (key === 'ocv') {
      busy(`Reading "${file.name}"…`);
      const buf = await file.arrayBuffer();
      const res = await callWorker({ type: 'source', buf }, [buf]);
      data.ocv = { fileName: file.name, cells: res.cells, lots: res.lots, masterName: res.masterName };
      const withSheet = res.cells.filter((c) => c.trackingSheet).length;
      setFileStatus(key, `✓ ${escapeHtml(file.name)}<br/>${res.cells.length} cells in "${escapeHtml(res.masterName)}", ${withSheet} with an OCV tracking sheet`, 'ok');
      fillLotSelects();
    } else if (key === 'review') {
      busy(`Reading "${file.name}"…`);
      const buf = await file.arrayBuffer();
      const res = await callWorker({ type: 'report', buf }, [buf]);
      if (!res.report) throw new Error('No "Cell ID" header row was found.');
      data.review = { fileName: file.name, byId: res.report };
      const withLoc = Object.values(res.report).filter((r) => (r.location || '').trim()).length;
      setFileStatus(key, `✓ ${escapeHtml(file.name)}<br/>${Object.keys(res.report).length} cells (sheet "${escapeHtml(res.sheetName)}"), ${withLoc} with a Location`, 'ok');
      fillLotSelects(true);
    } else if (key === 'stats') {
      busy(`Reading "${file.name}"…`);
      const buf = await file.arrayBuffer();
      const res = await callWorker({ type: 'lotstats', buf }, [buf]);
      if (!res.stats.length) throw new Error('No lot rows were found (expected: lot, E %, L %, production, E count, HM, L count).');
      data.stats = { fileName: file.name, rows: res.stats };
      setFileStatus(key, `✓ ${escapeHtml(file.name)}<br/>${res.stats.length} lots (${res.stats[0].lot} ~ ${res.stats[res.stats.length - 1].lot})`, 'ok');
    } else if (key === 'genealogy') {
      const c = card(key);
      c.querySelector('.file-status').innerHTML = `Reading ${escapeHtml(file.name)} (${(file.size / 1048576).toFixed(0)} MB)… <div class="progress"><i></i></div>`;
      const bar = c.querySelector('.progress > i');
      const res = await callWorker({ type: 'genealogy', file }, [], (p) => { bar.style.width = `${Math.round(p * 100)}%`; });
      data.genealogy = { fileName: file.name, byId: res.genealogy };
      setFileStatus(key, `✓ ${escapeHtml(file.name)}<br/>process history for ${res.cells} cells`, 'ok');
    }
  } catch (err) {
    console.error(err);
    data[key] = null;
    setFileStatus(key, `Could not read this file: ${escapeHtml(err.message)}`, 'error');
  } finally {
    busy(null);
    updateBuildState();
  }
}

// Lot pickers: default to the reviewed lots (analysis report), else the last 6 lots with tracking sheets
function fillLotSelects(preferReview) {
  if (!data.ocv) return;
  const lots = data.ocv.lots.map((l) => l.lot).filter((l) => M.lotParts(l)).map(M.padLot);
  const unique = [...new Set(lots)].sort();
  const from = $('fromLot'), to = $('toLot');
  const prevFrom = from.value, prevTo = to.value;
  [from, to].forEach((sel) => { sel.innerHTML = unique.map((l) => `<option value="${l}">${l}</option>`).join(''); });
  let dFrom, dTo;
  const reviewLots = data.review ? [...new Set(data.ocv.cells.filter((c) => data.review.byId[c.cellId.toUpperCase()]).map((c) => M.padLot(c.lot)))].sort() : [];
  if (reviewLots.length && (preferReview || !prevFrom)) { dFrom = reviewLots[0]; dTo = reviewLots[reviewLots.length - 1]; }
  else if (prevFrom && unique.includes(prevFrom)) { dFrom = prevFrom; dTo = prevTo; }
  else {
    const withSheet = data.ocv.lots.filter((l) => l.withSheet && M.lotParts(l.lot)).map((l) => M.padLot(l.lot)).sort();
    dTo = withSheet[withSheet.length - 1] || unique[unique.length - 1];
    dFrom = withSheet[Math.max(0, withSheet.length - 6)] || unique[0];
  }
  from.value = dFrom; to.value = dTo;
}

function updateBuildState() {
  const ready = !!(data.ocv && data.stats);
  $('buildBtn').disabled = !ready;
  const missing = [!data.ocv && 'OCV tracking workbook', !data.stats && 'lot summary'].filter(Boolean);
  $('buildHint').textContent = ready
    ? (data.review ? '' : 'Tip: without the analysis report, Location (inside/outside) charts stay empty. ') + (data.genealogy ? '' : 'Without the genealogy CSV, slides 5–7 are left blank.')
    : `Load the ${missing.join(' and ')} first.`;
}

// ---------------------------------------------------------------------------
// Build
// ---------------------------------------------------------------------------
async function build() {
  const settings = {
    from: $('fromLot').value, to: $('toLot').value,
    cumFrom: $('cumFrom').value.trim() || $('fromLot').value, trendFrom: $('trendFrom').value.trim() || $('fromLot').value,
    date: $('reportDate').value.trim(), team: $('team').value.trim(),
    eCriterion: $('eCriterion').value.trim(), lCriterion: $('lCriterion').value.trim(),
    notes: textToNotes($('notes').value),
    trendHeadline: $('trendHeadline').value.trim(), headline: $('headline').value.trim(), cumHeadline: $('cumHeadline').value.trim(),
  };
  if (M.lotKey(settings.from) > M.lotKey(settings.to)) { showToast('"From" lot must be before "to" lot.', true); return; }
  const minMv = Number($('minDrop').value);
  const minDropV = isFinite(minMv) && minMv >= 0 ? minMv / 1000 : Cv.DEFAULT_MIN_DROP_V;
  try {
    // Analyze the tracking sheets of every cell in the cumulative range (same rule as the tracker)
    const lo = M.lotKey(settings.cumFrom) < M.lotKey(settings.from) ? settings.cumFrom : settings.from;
    const cells = data.ocv.cells.filter((c) => c.trackingSheet && M.inRange(c.lot, lo, settings.to));
    busy(`Analyzing ${cells.length} OCV tracking sheets…`);
    const res = await callWorker({ type: 'tracking', sheets: cells.map((c) => c.trackingSheet) });
    const built = {};
    cells.forEach((c) => { built[c.cellId.toUpperCase()] = Cv.buildRow(c, res.analysis[c.trackingSheet], { minDropV }); });
    busy('Building slides…');
    const model = M.assemble({
      master: data.ocv.cells, built,
      review: data.review ? data.review.byId : {},
      stats: data.stats.rows,
      genealogy: data.genealogy ? data.genealogy.byId : {},
      settings,
    });
    const pres = P.build(window.PptxGenJS, model);
    const fileName = `${model.fileTitle}.pptx`;
    await pres.writeFile({ fileName });
    renderSummary(model, fileName);
    showToast(`Downloaded ${fileName}`);
  } catch (err) {
    console.error(err);
    showToast('Could not build the report: ' + err.message, true);
  } finally {
    busy(null);
  }
}
function renderSummary(model, fileName) {
  const s = model.report.summary;
  const warn = [];
  if (!model.counts.reviewed) warn.push('No cell of these lots is in the analysis report, so Location / Shape come out empty.');
  else if (model.counts.reviewed < model.counts.report) warn.push(`${model.counts.report - model.counts.reviewed} cell(s) are not in the analysis report (Location unknown).`);
  if (!model.counts.withGenealogy) warn.push('No process history for these cells — slides 5–7 are empty.');
  const el = $('summary');
  el.hidden = false;
  el.className = 'summary' + (warn.length ? ' warn' : '');
  el.innerHTML = `<b>${escapeHtml(fileName)}</b><br/>` +
    `${escapeHtml(model.label)}: ${s.total} cells analyzed · voltage drop ${s.dropAll} (${M.pct(s.dropAll, s.total)}%) · genuine ${s.genuine} · NTF ${s.ntf} · ` +
    `coating top ${s.loc['Coating Top']} / inside ${s.loc['Coating Inside']} / Al foil ${s.loc['Al Foil Surface']}${s.locUnknown ? ` / unknown ${s.locUnknown}` : ''}<br/>` +
    `Production ${model.report.lotSum.production.toLocaleString()} · E ${model.report.lotSum.eRate.toFixed(2)}% · L ${model.report.lotSum.lRate.toFixed(2)}% · cumulative ${escapeHtml(model.cumLabel)}: ${model.cumulative.summary.total} cells` +
    (warn.length ? `<br/>⚠ ${warn.map(escapeHtml).join('<br/>⚠ ')}` : '');
}

// ---------------------------------------------------------------------------
// Reset & wiring
// ---------------------------------------------------------------------------
function doReset() {
  Object.keys(data).forEach((k) => { data[k] = null; });
  document.querySelectorAll('.file-card').forEach((c) => { c.classList.remove('loaded', 'error'); c.querySelector('.file-status').innerHTML = ''; c.querySelector('input[type=file]').value = ''; });
  $('fromLot').innerHTML = ''; $('toLot').innerHTML = '';
  $('cumFrom').value = 'FD10'; $('trendFrom').value = 'FD10'; $('reportDate').value = today(); $('team').value = 'ESHG 품질팀';
  $('eCriterion').value = '2.42mV'; $('lCriterion').value = '3.5시그마'; $('minDrop').value = '1.5';
  ['trendHeadline', 'headline', 'cumHeadline'].forEach((id) => { $(id).value = ''; });
  try { localStorage.removeItem(NOTES_KEY); } catch (e) {}
  $('notes').value = notesToText(P.DEFAULT_NOTES);
  $('summary').hidden = true;
  $('resetOverlay').classList.remove('open');
  updateBuildState();
  showToast('Reset complete.');
}

function init() {
  $('reportDate').value = today();
  $('notes').value = loadNotes();
  $('notes').addEventListener('input', () => { try { localStorage.setItem(NOTES_KEY, $('notes').value); } catch (e) {} });
  $('notesDefaultBtn').onclick = () => { $('notes').value = notesToText(P.DEFAULT_NOTES); try { localStorage.removeItem(NOTES_KEY); } catch (e) {} };
  document.querySelectorAll('.file-card').forEach((c) => {
    const input = c.querySelector('input[type=file]');
    input.addEventListener('change', () => loadFile(c.dataset.key, input.files[0]));
    c.addEventListener('dragover', (e) => e.preventDefault());
    c.addEventListener('drop', (e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) loadFile(c.dataset.key, f); });
  });
  $('buildBtn').onclick = build;
  $('resetBtn').onclick = () => $('resetOverlay').classList.add('open');
  $('resetCancelBtn').onclick = () => $('resetOverlay').classList.remove('open');
  $('resetOverlay').addEventListener('click', (e) => { if (e.target === $('resetOverlay')) $('resetOverlay').classList.remove('open'); });
  $('resetConfirmBtn').onclick = doReset;
  updateBuildState();
}
init();

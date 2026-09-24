'use strict';

// ---------------------------------------------------------------------------
// Sheet (tab) definitions — each tab is its own collection / object store
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
    label: 'Frozen IR · Spot Analysis',
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
  loadedSheets: new Set(),
  loaded: false,
  filters: { search: '', date: '', sel: {} },
  formEditingId: null,
  formImages: [],
  formStagedFiles: [],
  detailId: null,
};

// ---------------------------------------------------------------------------
// IndexedDB (browser-only mode, and the source for migrating old local data)
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
      database.onversionchange = () => { database.close(); showToast('A newer version was opened in another tab. Please refresh this page.', true); };
      resolve(database);
    };
    req.onerror = () => reject(req.error);
    req.onblocked = () => showToast('Close or refresh other tabs with this site open (storage update pending).', true);
  });
  return idbPromise;
}
function idbStore(storeName, mode) { return openIdb().then((d) => d.transaction(storeName, mode).objectStore(storeName)); }
function idbRequest(storeName, mode, fn) {
  return idbStore(storeName, mode).then((store) => new Promise((resolve, reject) => {
    const req = fn(store);
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  }));
}
function idbGetAll(storeName) { return idbRequest(storeName, 'readonly', (s) => s.getAll()); }
function idbGet(storeName, id) { return idbRequest(storeName, 'readonly', (s) => s.get(id)); }
function idbPut(storeName, value) { return idbRequest(storeName, 'readwrite', (s) => s.put(value)).then(() => value); }
function idbDelete(storeName, id) { return idbRequest(storeName, 'readwrite', (s) => s.delete(id)); }
function idbClear(storeName) { return idbRequest(storeName, 'readwrite', (s) => s.clear()); }
async function idbPutMany(storeName, values) {
  const database = await openIdb();
  await new Promise((resolve, reject) => {
    const tx = database.transaction(storeName, 'readwrite');
    const store = tx.objectStore(storeName);
    values.forEach((v) => store.put(v));
    tx.oncomplete = resolve;
    tx.onerror = () => reject(tx.error);
    tx.onabort = () => reject(tx.error);
  });
}
function genId() {
  return (crypto.randomUUID ? crypto.randomUUID() : 'id-' + Date.now() + '-' + Math.random().toString(36).slice(2));
}

// ---------------------------------------------------------------------------
// Storage backends — same interface, chosen at startup.
//   live: true  => data arrives through realtime listeners (don't mutate state locally after writes)
// ---------------------------------------------------------------------------
const LocalBackend = {
  mode: 'local',
  live: false,
  async init() { await openIdb(); },
  async subscribe(onSheetData) {
    for (const [key, sheet] of Object.entries(SHEETS)) onSheetData(key, await idbGetAll(sheet.store));
  },
  putMany(store, recs) { return idbPutMany(store, recs); },
  // Local records are stored whole, so the caller passes the full record
  update(store, id, fields, fullRec) { return fullRec ? idbPut(store, fullRec) : Promise.resolve(); },
  async updateMany(store, list) { await idbPutMany(store, list.map((u) => u.fullRec).filter(Boolean)); },
  remove(store, id) { return idbDelete(store, id); },
  removeAll(store) { return idbClear(store); },
  async putImage(id, file) {
    await idbPut('images', { id, blob: file, contentType: file.type, originalName: file.name });
    return URL.createObjectURL(file);
  },
  async getImageUrl(id) {
    const rec = await idbGet('images', id);
    return rec && rec.blob ? URL.createObjectURL(rec.blob) : null;
  },
  removeImage(id) { return idbDelete('images', id); },
};

// Firestore documents are capped at 1 MiB, so images are stored as JPEG data URLs kept under this size
const MAX_IMAGE_DATAURL_CHARS = 950 * 1024;
const FIRESTORE_BATCH_LIMIT = 450;
function readAsDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });
}
async function imageToDataUrl(blob) {
  const original = await readAsDataUrl(blob);
  if (original.length <= MAX_IMAGE_DATAURL_CHARS && /^data:image\//.test(original)) return original;
  const img = await new Promise((resolve, reject) => {
    const el = new Image();
    el.onload = () => resolve(el);
    el.onerror = () => reject(new Error('image_decode_failed'));
    el.src = original;
  });
  let maxDim = 2400, quality = 0.9;
  for (let attempt = 0; attempt < 12; attempt++) {
    const scale = Math.min(1, maxDim / Math.max(img.naturalWidth, img.naturalHeight));
    const canvas = document.createElement('canvas');
    canvas.width = Math.max(1, Math.round(img.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(img.naturalHeight * scale));
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    const out = canvas.toDataURL('image/jpeg', quality);
    if (out.length <= MAX_IMAGE_DATAURL_CHARS) return out;
    if (quality > 0.65) quality -= 0.1; else maxDim = Math.round(maxDim * 0.8);
  }
  throw new Error('image_too_large');
}
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

const CloudBackend = {
  mode: 'cloud',
  live: true,
  db: null,
  unsubs: [],
  async init() {
    if (!firebase.apps.length) firebase.initializeApp(window.BTT_FIREBASE_CONFIG);
    this.db = firebase.firestore();
  },
  subscribe(onSheetData, onError) {
    this.unsubscribe();
    Object.entries(SHEETS).forEach(([key, sheet]) => {
      this.unsubs.push(this.db.collection(sheet.store).onSnapshot(
        (snap) => onSheetData(key, snap.docs.map((d) => d.data())),
        (err) => onError && onError(err),
      ));
    });
  },
  unsubscribe() { this.unsubs.forEach((u) => u()); this.unsubs = []; },
  async putMany(store, recs) {
    for (const part of chunk(recs, FIRESTORE_BATCH_LIMIT)) {
      const batch = this.db.batch();
      part.forEach((r) => batch.set(this.db.collection(store).doc(r.id), r));
      await batch.commit();
    }
  },
  update(store, id, fields) { return this.db.collection(store).doc(id).update(fields); },
  async updateMany(store, list) {
    for (const part of chunk(list, FIRESTORE_BATCH_LIMIT)) {
      const batch = this.db.batch();
      part.forEach((u) => batch.update(this.db.collection(store).doc(u.id), u.fields));
      await batch.commit();
    }
  },
  remove(store, id) { return this.db.collection(store).doc(id).delete(); },
  async removeAll(store) {
    const snap = await this.db.collection(store).get();
    for (const part of chunk(snap.docs, FIRESTORE_BATCH_LIMIT)) {
      const batch = this.db.batch();
      part.forEach((d) => batch.delete(d.ref));
      await batch.commit();
    }
  },
  async putImage(id, file) {
    const dataUrl = await imageToDataUrl(file);
    await this.db.collection('images').doc(id).set({ id, dataUrl, originalName: file.name || '', uploadedAt: new Date().toISOString() });
    return dataUrl;
  },
  async getImageUrl(id) {
    const doc = await this.db.collection('images').doc(id).get();
    return doc.exists ? doc.data().dataUrl : null;
  },
  removeImage(id) { return this.db.collection('images').doc(id).delete(); },
};

let backend = LocalBackend;
function cloudConfigured() {
  const cfg = window.BTT_FIREBASE_CONFIG;
  return !!(cfg && cfg.apiKey && cfg.projectId && window.firebase);
}

function onSheetData(key, rows) {
  rows.sort((a, b) => String(a.createdAt || '').localeCompare(String(b.createdAt || '')));
  state.data[key] = rows;
  if (key === activeSheetKey) state.cells = rows;
  state.loadedSheets.add(key);
  const wasLoaded = state.loaded;
  state.loaded = Object.keys(SHEETS).every((k) => state.loadedSheets.has(k));
  if (state.loaded) showLoading(false);
  scheduleRender();
  if (state.loaded && !wasLoaded && backend.mode === 'cloud') checkLocalDataForMigration();
}
// Coalesce bursts of realtime updates into one render per frame
let renderPending = false;
function scheduleRender() {
  if (renderPending) return;
  renderPending = true;
  requestAnimationFrame(() => {
    renderPending = false;
    populateFilterOptions();
    renderSheetTabs();
    renderStats();
    renderTable();
  });
}

async function initStorage() {
  if (cloudConfigured()) {
    backend = CloudBackend;
    try { await CloudBackend.init(); }
    catch (err) {
      console.error(err);
      setLiveBadge('unavailable', 'Connection error');
      showLoading(false);
      showEmpty(true, 'Could not connect to the shared database. Check firebase-config.js.');
      return;
    }
    document.getElementById('appSubtitle').textContent = 'Battery cell teardown analysis log · Shared database (any device)';
    firebase.auth().onAuthStateChanged((user) => {
      if (user) {
        hideLogin();
        document.getElementById('signOutBtn').hidden = false;
        setLiveBadge('connecting', 'Connecting');
        CloudBackend.subscribe((key, rows) => {
          onSheetData(key, rows);
          setLiveBadge('live', 'Synced · shared');
        }, (err) => {
          console.error(err);
          if (err && err.code === 'permission-denied') {
            setLiveBadge('unavailable', 'No access');
            showToast('Access denied by the database. Sign in with the team password again.', true);
          } else {
            setLiveBadge('error', 'Sync error');
            showToast('Lost connection to the shared database. Refresh the page.', true);
          }
        });
      } else {
        CloudBackend.unsubscribe();
        state.loadedSheets.clear();
        state.loaded = false;
        Object.keys(state.data).forEach((k) => { state.data[k] = []; });
        state.cells = state.data[activeSheetKey];
        imageUrlCache.clear();
        document.getElementById('signOutBtn').hidden = true;
        setLiveBadge('connecting', 'Sign-in required');
        showLoading(false);
        scheduleRender();
        showLogin();
      }
    });
    return;
  }
  try {
    await LocalBackend.init();
  } catch (err) {
    console.error(err);
    setLiveBadge('unavailable', 'Storage unavailable');
    showLoading(false);
    showEmpty(true, 'This browser does not support local storage (IndexedDB) or it is disabled. Open the page in a normal (non-private) window.');
    return;
  }
  document.getElementById('appSubtitle').textContent = 'Battery cell teardown analysis log · Saved in this browser only (shared database not configured)';
  await LocalBackend.subscribe(onSheetData);
  setLiveBadge('live', 'Saved in this browser');
}

// ---------------------------------------------------------------------------
// Team sign-in (cloud mode)
// ---------------------------------------------------------------------------
function showLogin() {
  document.getElementById('loginError').textContent = '';
  document.getElementById('loginOverlay').classList.add('open');
  setTimeout(() => document.getElementById('loginPassword').focus(), 0);
}
function hideLogin() {
  document.getElementById('loginOverlay').classList.remove('open');
  document.getElementById('loginPassword').value = '';
}
async function submitLogin(e) {
  e.preventDefault();
  const pw = document.getElementById('loginPassword').value;
  const errEl = document.getElementById('loginError');
  const btn = document.getElementById('loginBtn');
  if (!pw) { errEl.textContent = 'Enter the team password.'; return; }
  btn.disabled = true;
  errEl.textContent = '';
  try {
    await firebase.auth().setPersistence(firebase.auth.Auth.Persistence.LOCAL);
    await firebase.auth().signInWithEmailAndPassword(window.BTT_TEAM_EMAIL, pw);
  } catch (err) {
    console.error(err);
    const code = err && err.code;
    errEl.textContent = (code === 'auth/wrong-password' || code === 'auth/invalid-credential' || code === 'auth/invalid-login-credentials')
      ? 'Incorrect password.'
      : code === 'auth/too-many-requests' ? 'Too many attempts. Wait a few minutes and try again.'
        : code === 'auth/network-request-failed' ? 'Network error. Check your connection.'
          : 'Sign-in failed (' + (code || 'unknown error') + ').';
  } finally {
    btn.disabled = false;
  }
}

// ---------------------------------------------------------------------------
// One-time upload of data saved by the old browser-only version
// ---------------------------------------------------------------------------
const MIGRATION_FLAG = 'btt.localMigrationDone';
let pendingLocalData = null;
async function checkLocalDataForMigration() {
  try { if (localStorage.getItem(MIGRATION_FLAG)) return; } catch (e) {}
  try {
    const found = {};
    let total = 0;
    for (const [key, sheet] of Object.entries(SHEETS)) {
      found[key] = await idbGetAll(sheet.store);
      total += found[key].length;
    }
    if (!total) return;
    pendingLocalData = found;
    const parts = Object.entries(SHEETS).map(([key, sheet]) => `${sheet.label}: ${found[key].length}`).join(', ');
    document.getElementById('migrateBannerText').textContent =
      `This browser has ${total} record(s) saved by the earlier browser-only version (${parts}). Upload them to the shared database so they are available on every device?`;
    document.getElementById('migrateBanner').hidden = false;
  } catch (err) {
    console.warn('local data check failed', err);
  }
}
async function uploadLocalData() {
  if (!pendingLocalData) return;
  const btn = document.getElementById('migrateUploadBtn');
  btn.disabled = true;
  btn.textContent = 'Uploading…';
  let imageFailures = 0;
  try {
    for (const [key, sheet] of Object.entries(SHEETS)) {
      const cloudIds = new Set(state.data[key].map((r) => r.id));
      const recs = pendingLocalData[key].filter((r) => !cloudIds.has(r.id));
      for (const rec of recs) {
        const keptImages = [];
        for (const img of rec.images || []) {
          try {
            const stored = await idbGet('images', img.id);
            if (!stored || !stored.blob) { imageFailures++; continue; }
            const file = stored.blob;
            if (!file.name) { try { Object.defineProperty(file, 'name', { value: img.originalName || '' }); } catch (e) {} }
            await CloudBackend.putImage(img.id, file);
            keptImages.push(img);
          } catch (err) { console.warn('image upload failed', err); imageFailures++; }
        }
        rec.images = keptImages;
      }
      await CloudBackend.putMany(sheet.store, recs);
    }
    try { localStorage.setItem(MIGRATION_FLAG, '1'); } catch (e) {}
    document.getElementById('migrateBanner').hidden = true;
    pendingLocalData = null;
    showToast(imageFailures ? `Upload complete. ${imageFailures} image(s) could not be uploaded.` : 'Upload complete. Your records are now shared across devices.', imageFailures > 0);
  } catch (err) {
    console.error(err);
    showToast('Upload failed. Please try again.', true);
  } finally {
    btn.disabled = false;
    btn.textContent = 'Upload to shared database';
  }
}

// ---------------------------------------------------------------------------
// Reset
// ---------------------------------------------------------------------------
function openResetModal() {
  document.getElementById('resetCurrentLabel').textContent = `Current tab only (${SHEET.label}${SHEET.hasImages ? ', including attached images' : ''})`;
  document.getElementById('resetScopeNote').textContent = backend.mode === 'cloud'
    ? 'This deletes the data in the shared database for everyone on the team.'
    : 'This deletes the data saved in this browser.';
  document.querySelector('input[name="resetScope"][value="current"]').checked = true;
  const input = document.getElementById('resetConfirmInput');
  input.value = '';
  document.getElementById('resetConfirmBtn').disabled = true;
  document.getElementById('resetModalOverlay').classList.add('open');
  setTimeout(() => input.focus(), 0);
}
async function confirmReset() {
  if (document.getElementById('resetConfirmInput').value.trim() !== 'RESET') return;
  const scope = document.querySelector('input[name="resetScope"]:checked').value;
  const keys = scope === 'all' ? Object.keys(SHEETS) : [activeSheetKey];
  const btn = document.getElementById('resetConfirmBtn');
  btn.disabled = true;
  btn.textContent = 'Deleting…';
  try {
    for (const key of keys) {
      const sheet = SHEETS[key];
      if (sheet.hasImages && scope !== 'all') {
        for (const rec of state.data[key]) {
          for (const img of rec.images || []) { try { await backend.removeImage(img.id); } catch (err) { console.warn('image delete failed', err); } }
        }
      }
      await backend.removeAll(sheet.store);
    }
    if (scope === 'all') await backend.removeAll('images');
    imageUrlCache.clear();
    if (!backend.live) {
      keys.forEach((k) => onSheetData(k, []));
    }
    document.getElementById('resetModalOverlay').classList.remove('open');
    showToast(scope === 'all' ? 'All tabs have been reset.' : `"${SHEET.label}" has been reset.`);
  } catch (err) {
    console.error(err);
    showToast('Reset failed. Please try again.', true);
  } finally {
    btn.textContent = 'Delete data';
    btn.disabled = document.getElementById('resetConfirmInput').value.trim() !== 'RESET';
  }
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

// ---------------------------------------------------------------------------
// Cell mutation helpers
// ---------------------------------------------------------------------------
function renderAfterLocalChange() {
  populateFilterOptions();
  renderSheetTabs();
  renderStats();
  renderTable();
}
async function createCell(fields) {
  const now = new Date().toISOString();
  const rec = Object.assign(emptyFields(), fields, { id: genId(), images: [], createdAt: now, updatedAt: now });
  await backend.putMany(curStore(), [rec]);
  if (!backend.live) { state.cells.push(rec); renderAfterLocalChange(); }
  return rec.id;
}
// Bulk insert; createdAt is offset per row so load order matches paste order
async function createCells(fieldsList) {
  if (!fieldsList.length) return;
  const base = Date.now();
  const recs = fieldsList.map((fields, i) => {
    const ts = new Date(base + i).toISOString();
    return Object.assign(emptyFields(), fields, { id: genId(), images: [], createdAt: ts, updatedAt: ts });
  });
  await backend.putMany(curStore(), recs);
  if (!backend.live) { state.cells.push(...recs); renderAfterLocalChange(); }
}
async function updateCell(id, fields) {
  const patch = Object.assign({}, fields, { updatedAt: new Date().toISOString() });
  const cell = state.cells.find((c) => c.id === id);
  if (cell) Object.assign(cell, patch);
  else if (!backend.live) return;
  await backend.update(curStore(), id, patch, cell);
}
async function updateCells(list) {
  if (!list.length) return;
  const now = new Date().toISOString();
  const items = list.map(({ id, fields }) => {
    const patch = Object.assign({}, fields, { updatedAt: now });
    const cell = state.cells.find((c) => c.id === id);
    if (cell) Object.assign(cell, patch);
    return { id, fields: patch, fullRec: cell };
  });
  await backend.updateMany(curStore(), items);
}
async function deleteCell(cell) {
  if (cell.images && cell.images.length) {
    for (const img of cell.images) { try { await backend.removeImage(img.id); } catch (err) { console.warn('image delete failed', err); } }
  }
  await backend.remove(curStore(), cell.id);
  state.cells = state.data[activeSheetKey] = state.cells.filter((c) => c.id !== cell.id);
  renderAfterLocalChange();
}
async function addImagesToCell(cell, files) {
  const uploaded = [];
  for (const f of files) {
    const id = genId();
    const url = await backend.putImage(id, f);
    imageUrlCache.set(id, url);
    uploaded.push({ id, originalName: f.name, contentType: f.type, uploadedAt: new Date().toISOString() });
  }
  const live = state.cells.find((c) => c.id === cell.id);
  const newImages = ((live || cell).images || []).concat(uploaded);
  await updateCell(cell.id, { images: newImages });
  return newImages;
}
async function removeImageFromCell(cell, imageId) {
  try { await backend.removeImage(imageId); } catch (err) { console.warn('image delete failed', err); }
  imageUrlCache.delete(imageId);
  const live = state.cells.find((c) => c.id === cell.id);
  const newImages = ((live || cell).images || []).filter((im) => im.id !== imageId);
  await updateCell(cell.id, { images: newImages });
  return newImages;
}

// Image data is fetched asynchronously; cache resolved URLs per image id
// and attach them to <img> elements once ready.
const imageUrlCache = new Map();
function attachImageSrc(imgEl, imageId) {
  const cached = imageUrlCache.get(imageId);
  if (cached) { imgEl.src = cached; return; }
  backend.getImageUrl(imageId).then((url) => {
    if (!url) return;
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
    info.textContent = `${items.length} TD Date value(s) have no year. The year is never guessed — enter it, or choose "Keep as is".`;
    body.appendChild(info);
    if (items.length > 1) {
      const bulkRow = document.createElement('div');
      bulkRow.className = 'import-issue-row';
      bulkRow.innerHTML = `<span class="label">Apply one year to all</span><input type="number" id="bulkYearInput" placeholder="e.g. 2026" min="1990" max="2100" /><button class="btn btn-ghost" id="bulkYearApplyBtn" type="button">Apply to all</button>`;
      body.appendChild(bulkRow);
    }
    items.forEach((item, idx) => {
      const row = document.createElement('div');
      row.className = 'import-issue-row';
      row.innerHTML = `<span class="label">${escapeHtml(item.label)} — original: "${escapeHtml(item.original)}"</span><input type="number" class="year-input" data-idx="${idx}" placeholder="Year" min="1990" max="2100" /><label style="font-size:11px; display:flex; align-items:center; gap:4px;"><input type="checkbox" class="keep-as-is" data-idx="${idx}" /> Keep as is</label>`;
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
        else { showToast('Enter a year for every item, or choose "Keep as is".', true); return; }
      }
      cleanup(); resolve(true);
    };
  });
}

// ---------------------------------------------------------------------------
// Stats row
// ---------------------------------------------------------------------------
function renderStats() {
  document.getElementById('statTotal').textContent = `${state.cells.length} record${state.cells.length === 1 ? '' : 's'}`;
  const row = document.getElementById('statsRow');
  row.querySelectorAll('.stat-chip').forEach((el) => el.remove());
  const counts = new Map();
  state.cells.forEach((c) => {
    const v = (c[SHEET.statusKey] || '').trim();
    if (!v) return;
    counts.set(v, (counts.get(v) || 0) + 1);
  });
  if (activeSheetKey === 'analysis') {
    const needs = state.cells.filter(ocvNeedsInput).length;
    if (needs) {
      const chip = document.createElement('span');
      chip.className = 'stat-chip needs-input';
      chip.title = 'Show only rows that need input';
      chip.innerHTML = `⚠ Needs input <b>${needs}</b>`;
      chip.onclick = () => { const cb = document.getElementById('needsInputOnly'); cb.checked = true; cb.dispatchEvent(new Event('change')); };
      row.insertBefore(chip, row.querySelector('.paste-hint'));
    }
  }
  [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8).forEach(([status, n]) => {
    const chip = document.createElement('span');
    chip.className = 'stat-chip';
    chip.innerHTML = `${escapeHtml(status)} <b>${n}</b>`;
    row.insertBefore(chip, row.querySelector('.paste-hint'));
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
    if (f.needsInput && !ocvNeedsInput(c)) return false;
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
      if (col.type === 'long') { input.rows = 1; input.placeholder = 'Shift+Enter for new line'; }
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
  viewBtn.className = 'icon-btn'; viewBtn.type = 'button'; viewBtn.title = 'View details'; viewBtn.textContent = '🔍';
  viewBtn.onclick = (e) => { e.stopPropagation(); openDetailModal(cell); };
  actions.appendChild(viewBtn);
  const delBtn = document.createElement('button');
  delBtn.className = 'icon-btn danger'; delBtn.type = 'button'; delBtn.title = 'Delete row'; delBtn.textContent = '🗑';
  delBtn.onclick = async (e) => {
    e.stopPropagation();
    if (!confirm(`Delete the row for Cell ID "${cell.cellId || '(none)'}"? This cannot be undone.`)) return;
    try { await deleteCell(cell); showToast('Row deleted.'); }
    catch (err) { console.error(err); showToast('An error occurred while deleting.', true); }
  };
  actions.appendChild(delBtn);
  actionsTd.appendChild(actions);
  tr.appendChild(actionsTd);
  applyOcvHighlights(tr, cell);
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
  addBtn.className = 'thumb-add'; addBtn.type = 'button'; addBtn.textContent = '+'; addBtn.title = 'Attach images';
  addBtn.onclick = (e) => { e.stopPropagation(); triggerImageUpload(cell); };
  wrap.appendChild(addBtn);
  return wrap;
}
function triggerImageUpload(cell) {
  const input = document.createElement('input');
  input.type = 'file'; input.accept = 'image/*'; input.multiple = true;
  input.onchange = async () => {
    if (!input.files.length) return;
    try { await addImagesToCell(cell, Array.from(input.files)); showToast('Images uploaded.'); }
    catch (err) { console.error(err); showToast('Image upload failed.', true); }
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
    if (field === SHEET.statusKey || cell.ocv) renderStats();
    applyOcvHighlights(tr, cell);
    if (SHEET.filters.some(([k]) => k === field)) populateFilterOptions();
    if (field === 'cellId') {
      const dup = state.cells.filter((c) => c.id !== cell.id && (c.cellId || '').trim().toLowerCase() === value.trim().toLowerCase() && value.trim());
      if (dup.length) showToast(`⚠ Duplicate Cell ID: "${value}" — ${dup.length} existing record(s) (LOT: ${dup.map((d) => d.lot || '-').join(', ')})`, true);
      renderTable();
    }
  } catch (err) {
    console.error(err);
    cell[field] = prev; el.value = prev;
    showToast('Save failed.', true);
  }
}

// ---------------------------------------------------------------------------
// Paste-from-Excel
// ---------------------------------------------------------------------------
// Excel copies as TSV; cells containing tabs/newlines/quotes are wrapped in "..." with "" escapes
function parseClipboardGrid(text) {
  const s = String(text).replace(/\r\n?/g, '\n');
  const rows = [];
  let row = [], cell = '', i = 0, inQuotes = false;
  while (i < s.length) {
    const ch = s[i];
    if (inQuotes) {
      if (ch === '"' && s[i + 1] === '"') { cell += '"'; i += 2; continue; }
      if (ch === '"' && (i + 1 === s.length || s[i + 1] === '\t' || s[i + 1] === '\n')) { inQuotes = false; i++; continue; }
      cell += ch; i++; continue;
    }
    if (ch === '"' && cell === '') { inQuotes = true; i++; continue; }
    if (ch === '\t') { row.push(cell); cell = ''; i++; continue; }
    if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; i++; continue; }
    cell += ch; i++;
  }
  if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
  return rows.filter((r) => r.some((v) => v.trim() !== ''));
}
// If the first copied row is this tab's header row, return a per-column key map
function detectHeaderKeys(firstRow) {
  const keys = firstRow.map((h) => HEADER_MAP[normalizeHeader(h)] || null);
  const nonEmpty = firstRow.filter((h) => h.trim()).length;
  const matched = keys.filter((k) => k && EDITABLE_KEYS.includes(k)).length;
  return matched >= 2 && matched * 2 >= nonEmpty ? keys : null;
}
async function onTablePaste(e) {
  const target = e.target;
  if (!target.classList || !target.classList.contains('cell-input')) return;
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  if (!text) return;
  // Multi-line prose pasted into a long-text cell stays in that cell
  if (target.tagName === 'TEXTAREA' && !text.includes('\t')) return;
  e.preventDefault();
  const grid = parseClipboardGrid(text);
  if (grid.length === 0) return;
  if (grid.length === 1 && grid[0].length === 1) {
    target.value = grid[0][0];
    target.dispatchEvent(new Event('change', { bubbles: true }));
    return;
  }
  const headerKeys = detectHeaderKeys(grid[0]);
  if (headerKeys) { await handleGridPaste(grid.slice(1), [], 0, 0, headerKeys); return; }
  const startColIdx = EDITABLE_KEYS.indexOf(target.dataset.field);
  if (startColIdx === -1) return;
  const startRowId = target.closest('tr').dataset.id;
  const filteredRows = getFilteredCells();
  const startRowIdx = filteredRows.findIndex((r) => r.id === startRowId);
  if (startRowIdx === -1) return;
  await handleGridPaste(grid, filteredRows, startRowIdx, startColIdx);
}
// Ctrl+V outside any cell (e.g. on an empty table): append the rows at the bottom from the first column
async function onDocumentPaste(e) {
  if (e.defaultPrevented) return;
  if (e.target.closest && e.target.closest('input, textarea, select, [contenteditable], .modal-overlay')) return;
  if (document.querySelector('.modal-overlay.open')) return;
  const text = (e.clipboardData || window.clipboardData).getData('text/plain');
  if (!text || (!text.includes('\t') && !text.includes('\n'))) return;
  e.preventDefault();
  const grid = parseClipboardGrid(text);
  if (!grid.length) return;
  const headerKeys = detectHeaderKeys(grid[0]);
  await handleGridPaste(headerKeys ? grid.slice(1) : grid, [], 0, 0, headerKeys);
}
async function handleGridPaste(grid, filteredRows, startRowIdx, startColIdx, headerKeys) {
  if (!grid.length) { showToast('No data rows to paste.', true); return; }
  const updates = [];
  const creations = [];
  grid.forEach((line, i) => {
    const rowIdx = startRowIdx + i;
    let target = rowIdx < filteredRows.length
      ? { existing: true, id: filteredRows[rowIdx].id, fields: {} }
      : { existing: false, fields: {} };
    line.forEach((val, j) => {
      const key = headerKeys ? headerKeys[j] : EDITABLE_KEYS[startColIdx + j];
      if (!key || !EDITABLE_KEYS.includes(key)) return;
      target.fields[key] = val.trim();
    });
    (target.existing ? updates : creations).push(target);
  });
  const yearItems = [];
  [...updates, ...creations].forEach((t, idx) => {
    if (t.fields.tdDate && looksLikeDateMissingYear(t.fields.tdDate)) {
      const existingCell = t.existing ? filteredRows.find((r) => r.id === t.id) : null;
      yearItems.push({
        label: t.fields.cellId ? `Cell ID: ${t.fields.cellId}` : (existingCell ? `Existing row (Cell ID: ${existingCell.cellId || '-'})` : `New row #${idx + 1}`),
        original: t.fields.tdDate,
        apply: (v) => { t.fields.tdDate = v; },
      });
    }
  });
  if (yearItems.length) {
    const ok = await resolveMissingYearDates(yearItems);
    if (!ok) { showToast('Paste cancelled.'); return; }
  }
  try {
    await updateCells(updates);
    if (creations.length) await createCells(creations.map((c) => c.fields));
    populateFilterOptions();
    renderStats();
    renderTable();
    const parts = [];
    if (updates.length) parts.push(`filled ${updates.length} existing row(s)`);
    if (creations.length) parts.push(`added ${creations.length} new row(s)`);
    showToast(`Paste complete — ${parts.join(', ')}${headerKeys ? ' (columns matched by Excel header)' : ''}`);
  } catch (err) {
    console.error(err);
    showToast('An error occurred while pasting.', true);
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
  document.getElementById('detailModalTitle').textContent = `${cell.cellId || '(No Cell ID)'} — Details`;
  const body = document.getElementById('detailModalBody');
  body.innerHTML = '';
  const grid = document.createElement('div');
  grid.className = 'detail-grid';
  COLUMNS.filter((c) => c.type !== 'image').forEach((col) => {
    const item = document.createElement('div');
    item.className = 'detail-item' + (col.type === 'long' ? ' full' : '');
    const val = cell[col.key];
    item.innerHTML = `<div class="k">${escapeHtml(col.label)}</div><div class="v ${val ? '' : 'empty'}">${val ? escapeHtml(val) : '(empty)'}</div>`;
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
    const empty = document.createElement('div'); empty.className = 'v empty'; empty.textContent = '(No images attached)';
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
  document.getElementById('formModalTitle').textContent = cell ? 'Edit cell record' : 'New cell record';
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
      hint.textContent = 'Dates without a year (e.g. 1/26) will ask for the year when saved.';
      field.appendChild(hint);
    }
    if (col.key === 'ir') {
      const hint = document.createElement('span'); hint.className = 'field-hint';
      hint.textContent = 'Text such as "O.F." is saved exactly as typed.';
      field.appendChild(hint);
    }
    if (col.key === 'cellId') input.addEventListener('input', debounce(() => updateDupWarning(input.value), 300));
    grid.appendChild(field);
  });
  body.appendChild(grid);

  if (SHEET.hasImages) {
  const imgField = document.createElement('div');
  imgField.className = 'form-field full';
  const imgLabel = document.createElement('label'); imgLabel.textContent = 'VHX/Image (attachments)';
  imgField.appendChild(imgLabel);
  const attachArea = document.createElement('div'); attachArea.className = 'image-attach-area'; attachArea.id = 'formImageAttachArea';
  imgField.appendChild(attachArea);
  const addImgBtn = document.createElement('label');
  addImgBtn.className = 'btn btn-secondary file-btn'; addImgBtn.textContent = 'Add images';
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
      } catch (err) { console.error(err); showToast('Image upload failed.', true); }
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
        } catch (err) { console.error(err); showToast('Image delete failed.', true); return; }
      } else {
        state.formImages = state.formImages.filter((i) => i !== img);
      }
      renderFormImageAttachArea();
    };
    wrap.appendChild(t); wrap.appendChild(rm); area.appendChild(wrap);
  });
  state.formStagedFiles.forEach((f, idx) => {
    const wrap = document.createElement('div'); wrap.style.position = 'relative';
    const t = document.createElement('img'); t.src = URL.createObjectURL(f); t.className = 'thumb'; t.title = f.name + ' (uploaded on save)'; t.style.opacity = '0.6';
    const rm = document.createElement('button');
    rm.textContent = '×'; rm.type = 'button'; rm.className = 'icon-btn danger';
    rm.style.cssText = 'position:absolute;top:-8px;right:-8px;background:var(--panel);border-radius:50%;width:18px;height:18px;padding:0;line-height:16px;font-size:12px;border:1px solid var(--border);';
    rm.onclick = () => { state.formStagedFiles.splice(idx, 1); renderFormImageAttachArea(); };
    wrap.appendChild(t); wrap.appendChild(rm); area.appendChild(wrap);
  });
  if (!state.formImages.length && !state.formStagedFiles.length) {
    const empty = document.createElement('span'); empty.className = 'field-hint'; empty.textContent = 'No images attached.';
    area.appendChild(empty);
  }
}
function updateDupWarning(cellIdValue) {
  const warnEl = document.getElementById('dupWarning');
  const v = (cellIdValue || '').trim().toLowerCase();
  if (!v) { warnEl.innerHTML = ''; return; }
  const dup = state.cells.filter((c) => c.id !== state.formEditingId && (c.cellId || '').trim().toLowerCase() === v);
  if (!dup.length) { warnEl.innerHTML = ''; return; }
  warnEl.innerHTML = `⚠ Duplicate Cell ID (${dup.length} existing). <a href="#" id="viewDupLink">View existing record</a>`;
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
    showToast(state.formEditingId ? 'Record updated.' : 'Record added.');
  } catch (err) {
    console.error(err);
    showToast('An error occurred while saving.', true);
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
  if (!rows.length) { showToast('The file is empty.', true); return; }
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
  if (!parsed.length) { showToast('No data to import.', true); return; }
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
  if (!ok) { showToast('Import cancelled.'); return; }
  await createCells(parsed);
  showToast(`Imported ${parsed.length} record(s).`);
}
function showImportReviewModal({ parsed, ignoredHeaders, dupCount, yearItems }) {
  return new Promise((resolve) => {
    const overlay = document.getElementById('importModalOverlay');
    const body = document.getElementById('importModalBody');
    body.innerHTML = '';
    const summary = document.createElement('div');
    summary.className = 'import-summary';
    summary.innerHTML = `Importing <b>${parsed.length}</b> row(s) into the <b>[${escapeHtml(SHEET.label)}]</b> tab.<br/>${ignoredHeaders.length ? `Unrecognized columns (ignored): ${escapeHtml(ignoredHeaders.join(', '))}<br/>` : ''}${dupCount ? `<span style="color:var(--warn)">⚠ Rows whose Cell ID duplicates existing data: ${dupCount} (they are still added and are highlighted as duplicates in the table).</span><br/>` : ''}${SHEET.hasImages ? 'Image files in the VHX/Image column cannot be imported and are ignored. Attach them on each row after importing.' : ''}`;
    body.appendChild(summary);
    if (yearItems.length) {
      const yearBox = document.createElement('div');
      const warn = document.createElement('p'); warn.className = 'field-warning';
      warn.textContent = `${yearItems.length} TD Date value(s) have no year. The year is never guessed — enter it, or choose "Keep as is".`;
      yearBox.appendChild(warn);
      const bulkRow = document.createElement('div'); bulkRow.className = 'import-issue-row';
      bulkRow.innerHTML = `<span class="label">Apply one year to all</span><input type="number" id="importBulkYear" placeholder="e.g. 2026" min="1990" max="2100" /><button class="btn btn-ghost" type="button" id="importBulkYearBtn">Apply to all</button>`;
      yearBox.appendChild(bulkRow);
      yearItems.forEach((item, idx) => {
        const row = document.createElement('div'); row.className = 'import-issue-row';
        row.innerHTML = `<span class="label">${escapeHtml(item.label)} — original: "${escapeHtml(item.original)}"</span><input type="number" class="import-year-input" data-idx="${idx}" placeholder="Year" min="1990" max="2100" /><label style="font-size:11px; display:flex; align-items:center; gap:4px;"><input type="checkbox" class="import-keep-as-is" data-idx="${idx}" /> Keep as is</label>`;
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
      more.textContent = `...and ${parsed.length - 15} more row(s)`;
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
        else { showToast('Some years are still missing. Enter them, or choose "Keep as is".', true); return; }
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
  if (activeSheetKey === 'analysis') { exportAnalysisReport(); return; }
  const ws = XLSX.utils.json_to_sheet(buildExportRows(), { header: COLUMNS.map((c) => c.label) });
  const wbx = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbx, ws, SHEET.sheetName);
  XLSX.writeFile(wbx, `${SHEET.exportName}-${dateStamp()}.xlsx`);
}

// ---------------------------------------------------------------------------
// OCV tracking workbook import (Frozen IR · Spot Analysis tab)
// Rows are added or updated by Cell ID. Each imported row keeps its auto-filled values in
// rec.ocv.auto, so a later import only overwrites fields the team has not edited.
// ---------------------------------------------------------------------------
const OcvC = window.OcvConvert;
// convert.js column key -> Analysis tab field
const OCV_FIELD = { frozenPf: 'frozenIrPf', layer: 'droppedLayer', spot: 'spotFound', sem: 'semEds' };
const ocvField = (k) => OCV_FIELD[k] || k;
const OCV_KEYS = OcvC ? OcvC.COLUMNS.map((c) => c.key) : [];
const ocvState = { source: null, lots: new Set() };

let ocvWorker = null;
let ocvSeq = 0;
const ocvPending = new Map();
function callOcvWorker(msg, transfer) {
  if (!ocvWorker) {
    ocvWorker = new Worker('worker.js?v=5');
    ocvWorker.onmessage = (e) => {
      const p = ocvPending.get(e.data.id);
      if (!p) return;
      ocvPending.delete(e.data.id);
      e.data.ok ? p.resolve(e.data) : p.reject(new Error(e.data.error));
    };
    ocvWorker.onerror = (e) => {
      console.error(e);
      ocvPending.forEach((p) => p.reject(new Error('The file reader failed to start. Check your internet connection and reload.')));
      ocvPending.clear();
    };
  }
  return new Promise((resolve, reject) => {
    const id = ++ocvSeq;
    ocvPending.set(id, { resolve, reject });
    ocvWorker.postMessage(Object.assign({ id }, msg), transfer || []);
  });
}

// Which Analysis fields of an imported row still need input, and which auto values to verify
function ocvRowStatus(cell) {
  if (!cell.ocv || !OcvC) return null;
  const values = {};
  OCV_KEYS.forEach((k) => { values[k] = (cell[ocvField(k)] || '').trim(); });
  const missing = new Set(OcvC.requiredKeys(values).filter((k) => !values[k]).map(ocvField));
  const verify = new Set();
  const src = cell.ocv.src || {};
  const auto = cell.ocv.auto || {};
  Object.keys(src).forEach((f) => { if (src[f] === 'calc' && cell[f] && cell[f] === auto[f]) verify.add(f); });
  return { missing, verify };
}
function ocvCellTitle(cell, field) {
  if (!cell.ocv) return '';
  const auto = cell.ocv.auto || {};
  const src = (cell.ocv.src || {})[field];
  const note = (cell.ocv.note || {})[field];
  if (!cell[field] || cell[field] !== auto[field]) return '';
  const base = { master: 'From Master E & L', sheet: 'From the OCV tracking sheet', calc: 'Tracking sheet and Master disagree — please verify' }[src] || '';
  return [base, note && note !== base ? note : ''].filter(Boolean).join(' — ');
}
// Apply highlight classes / tooltips to one rendered row of the Analysis tab
function applyOcvHighlights(tr, cell) {
  if (activeSheetKey !== 'analysis') return;
  const status = ocvRowStatus(cell);
  tr.querySelectorAll('[data-field]').forEach((input) => {
    const td = input.closest('td');
    const f = input.dataset.field;
    td.classList.toggle('ocv-missing', !!(status && status.missing.has(f)));
    td.classList.toggle('ocv-calc', !!(status && status.verify.has(f)));
    const title = status ? (status.missing.has(f) ? 'Needs manual input' : ocvCellTitle(cell, f)) : '';
    if (title) input.title = title; else input.removeAttribute('title');
  });
}
function ocvNeedsInput(cell) {
  const s = ocvRowStatus(cell);
  return !!(s && s.missing.size);
}

function openOcvModal() {
  document.getElementById('ocvModalOverlay').classList.add('open');
  renderOcvOptions();
}
function closeOcvModal() { document.getElementById('ocvModalOverlay').classList.remove('open'); }
function setOcvStatus(html, isError) {
  const el = document.getElementById('ocvStatus');
  el.hidden = !html;
  el.className = 'ocv-status' + (isError ? ' error' : '');
  el.innerHTML = html || '';
}
async function loadOcvWorkbook(file) {
  if (!file) return;
  setOcvStatus(`Reading "${escapeHtml(file.name)}"…`);
  document.getElementById('ocvOptions').hidden = true;
  try {
    const buf = await file.arrayBuffer();
    const res = await callOcvWorker({ type: 'source', buf }, [buf]);
    if (!res.cells.length) throw new Error('No cell rows were found in the Master sheet.');
    ocvState.source = { fileName: file.name, cells: res.cells, lots: res.lots };
    ocvState.lots.clear();
    const warn = res.missingHeaders.length ? `<br/>⚠ Columns not found in the Master sheet (left for manual input): ${escapeHtml(res.missingHeaders.join(', '))}` : '';
    setOcvStatus(`Loaded <b>${escapeHtml(file.name)}</b> — ${res.cells.length} cells in "${escapeHtml(res.masterName)}", ${res.sheetCount} sheets. Choose LOTs:${warn}`);
    document.getElementById('ocvOptions').hidden = false;
    renderOcvOptions();
  } catch (err) {
    console.error(err);
    ocvState.source = null;
    setOcvStatus(`Could not read this file: ${escapeHtml(err.message)}`, true);
  }
}
function ocvSelectedCells() {
  if (!ocvState.source) return [];
  const onlyWithSheet = document.getElementById('ocvOnlyWithSheet').checked;
  return ocvState.source.cells.filter((c) => ocvState.lots.has(c.lot || '(no LOT)') && (!onlyWithSheet || c.trackingSheet));
}
function renderOcvOptions() {
  const grid = document.getElementById('ocvLotGrid');
  if (!ocvState.source) { grid.innerHTML = ''; updateOcvHint(); return; }
  const onlyWithSheet = document.getElementById('ocvOnlyWithSheet').checked;
  grid.innerHTML = '';
  ocvState.source.lots.forEach((l) => {
    const count = onlyWithSheet ? l.withSheet : l.total;
    const el = document.createElement('label');
    el.className = 'ocv-lot' + (ocvState.lots.has(l.lot) ? ' on' : '') + (count ? '' : ' empty');
    el.innerHTML = `<input type="checkbox" ${ocvState.lots.has(l.lot) ? 'checked' : ''} /> <b>${escapeHtml(l.lot)}</b><small>${count}</small>`;
    el.querySelector('input').onchange = (e) => {
      if (e.target.checked) ocvState.lots.add(l.lot); else ocvState.lots.delete(l.lot);
      el.classList.toggle('on', e.target.checked);
      updateOcvHint();
    };
    grid.appendChild(el);
  });
  updateOcvHint();
}
function updateOcvHint() {
  const cells = ocvSelectedCells();
  const existing = new Set(state.data.analysis.map((r) => (r.cellId || '').trim().toUpperCase()));
  const upd = cells.filter((c) => existing.has(c.cellId.toUpperCase())).length;
  document.getElementById('ocvImportConfirm').disabled = !cells.length;
  document.getElementById('ocvSelHint').textContent = cells.length
    ? `${cells.length} cell${cells.length === 1 ? '' : 's'}: ${cells.length - upd} new, ${upd} already in this tab`
    : (ocvState.source ? 'Select one or more LOTs' : '');
}
async function runOcvImport() {
  const cells = ocvSelectedCells();
  if (!cells.length) return;
  const btn = document.getElementById('ocvImportConfirm');
  btn.disabled = true;
  const minMv = Number(document.getElementById('ocvMinDrop').value);
  const minDropV = isFinite(minMv) && minMv >= 0 ? minMv / 1000 : OcvC.DEFAULT_MIN_DROP_V;
  try {
    const withSheet = cells.filter((c) => c.trackingSheet);
    setOcvStatus(`Analyzing ${withSheet.length} OCV tracking sheets…`);
    const res = await callOcvWorker({ type: 'tracking', sheets: withSheet.map((c) => c.trackingSheet) });
    const importedAt = new Date().toISOString();
    const byId = new Map();
    state.data.analysis.forEach((r) => { const k = (r.cellId || '').trim().toUpperCase(); if (k && !byId.has(k)) byId.set(k, r); });
    const creations = [];
    const updates = [];
    let keptEdits = 0;
    cells.forEach((cell) => {
      const built = OcvC.buildRow(cell, cell.trackingSheet ? res.analysis[cell.trackingSheet] : null, { minDropV });
      const auto = {}, src = {}, note = {};
      OCV_KEYS.forEach((k) => {
        const f = ocvField(k);
        auto[f] = built[k].value;
        src[f] = built[k].source;
        if (built[k].note) note[f] = built[k].note;
      });
      const ocv = { auto, src, note, file: ocvState.source.fileName, importedAt, minDropV };
      const existing = byId.get(cell.cellId.toUpperCase());
      if (!existing) {
        const fields = { ocv };
        Object.keys(auto).forEach((f) => { if (auto[f]) fields[f] = auto[f]; });
        creations.push(fields);
        return;
      }
      const prevAuto = (existing.ocv && existing.ocv.auto) || {};
      const fields = { ocv };
      Object.keys(auto).forEach((f) => {
        const cur = (existing[f] || '').trim();
        if (!auto[f] || cur === auto[f]) return;
        // Fill blanks and refresh values that still equal the previous import; keep the team's edits
        if (cur === '' || (Object.prototype.hasOwnProperty.call(prevAuto, f) && cur === prevAuto[f])) fields[f] = auto[f];
        else keptEdits++;
      });
      updates.push({ id: existing.id, fields });
    });
    setOcvStatus('Saving…');
    if (creations.length) await createCells(creations);
    if (updates.length) await updateCells(updates);
    populateFilterOptions();
    renderStats();
    renderTable();
    closeOcvModal();
    const needs = state.data.analysis.filter(ocvNeedsInput).length;
    showToast(`Imported ${cells.length} cells (${creations.length} new, ${updates.length} updated${keptEdits ? `, ${keptEdits} edited value(s) kept` : ''}). ${needs} row(s) need input.`);
  } catch (err) {
    console.error(err);
    setOcvStatus('Import failed: ' + escapeHtml(err.message), true);
  } finally {
    btn.disabled = false;
  }
}

// Excel export for the Analysis tab in the report layout (headers on row 2, No. in column B,
// yellow = needs input, light blue = verify), plus a Legend sheet.
const OCV_NUMERIC = new Set(['docv', 'frozenIr', 'droppedLayer', 'docvV', 'x', 'y', 'longSide', 'shortSide', 'height']);
function exportAnalysisReport() {
  const rows = getFilteredCells();
  const THIN = { style: 'thin', color: { rgb: 'BFBFBF' } };
  const BORDER = { top: THIN, bottom: THIN, left: THIN, right: THIN };
  const FILL_MISSING = { patternType: 'solid', fgColor: { rgb: 'FFFF00' } };
  const FILL_CALC = { patternType: 'solid', fgColor: { rgb: 'DDEBF7' } };
  const ws = {};
  const put = (r, c, v, s) => {
    const cell = v === '' || v == null ? { t: 's', v: '' } : (typeof v === 'number' ? { t: 'n', v } : { t: 's', v: String(v) });
    if (s) cell.s = s;
    ws[XLSX.utils.encode_cell({ r, c })] = cell;
  };
  put(1, 1, '', { fill: { patternType: 'solid', fgColor: { rgb: 'D9D9D9' } }, border: BORDER });
  COLUMNS.forEach((col, i) => put(1, i + 2, col.label, { font: { bold: true }, border: BORDER, alignment: { wrapText: true, vertical: 'center' } }));
  rows.forEach((cell, ri) => {
    const r = ri + 2;
    const status = ocvRowStatus(cell);
    put(r, 1, ri + 1, { border: BORDER });
    COLUMNS.forEach((col, i) => {
      let v = (cell[col.key] || '').trim();
      if (v && OCV_NUMERIC.has(col.key) && !isNaN(Number(v))) v = Number(v);
      const style = { border: BORDER };
      if (status && status.missing.has(col.key)) style.fill = FILL_MISSING;
      else if (status && status.verify.has(col.key)) style.fill = FILL_CALC;
      put(r, i + 2, v, style);
    });
  });
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 1, c: 1 }, e: { r: Math.max(rows.length, 1) + 1, c: COLUMNS.length + 1 } });
  ws['!cols'] = [{ wch: 2 }, { wch: 5 }].concat(COLUMNS.map((c) => ({ wch: c.key === 'cellId' ? 13 : Math.max(8, Math.min(26, c.label.length + 2)) })));

  const imports = rows.map((c) => c.ocv).filter(Boolean);
  const last = imports.sort((a, b) => String(b.importedAt).localeCompare(String(a.importedAt)))[0];
  const legend = XLSX.utils.aoa_to_sheet([
    ['Legend'],
    ['Yellow', 'Needs manual input (not found in the OCV tracking workbook)'],
    ['Light blue', 'Voltage drop from the OCV tracking sheet disagrees with Master E & L — please verify'],
    ['Voltage drop', `Per layer dOCV = biggest fall between the tracking dates (C, D, E). Drop = a layer > 2.6σ above the others and ≥ ${last ? (last.minDropV * 1000).toFixed(1) : '1.5'} mV; otherwise NTF.`],
    ['Frozen IR P/F', 'NG when Frozen IR < 35 MΩ, OK otherwise (when Master E & L has no result)'],
    [],
    ['Source file', last ? last.file : '(rows entered by hand)'],
    ['Downloaded', new Date().toLocaleString()],
  ]);
  legend.A1.s = { font: { bold: true } };
  legend.A2.s = { fill: FILL_MISSING };
  legend.A3.s = { fill: FILL_CALC };
  legend['!cols'] = [{ wch: 16 }, { wch: 110 }];

  const lots = [...new Set(rows.map((c) => (c.lot || '').trim()).filter(Boolean))].sort((a, b) => OcvC.lotSortKey(a).localeCompare(OcvC.lotSortKey(b)));
  const d = new Date();
  const stamp = String(d.getFullYear()).slice(2) + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0');
  const range = !lots.length ? 'Analysis' : lots.length === 1 ? OcvC.padLot(lots[0]) : `${OcvC.padLot(lots[0])}~${OcvC.padLot(lots[lots.length - 1])}`;
  const sheetName = `${stamp} ${range}`.replace(/[\[\]:*?\/\\]/g, '-').slice(0, 31);
  const wbx = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wbx, ws, sheetName);
  XLSX.utils.book_append_sheet(wbx, legend, 'Legend');
  XLSX.writeFile(wbx, `${sheetName}.xlsx`);
  const remaining = rows.reduce((n, c) => { const s = ocvRowStatus(c); return n + (s ? s.missing.size : 0); }, 0);
  showToast(remaining ? `Downloaded. ${remaining} highlighted cell(s) still need input.` : 'Downloaded.');
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
    sel.innerHTML = `<option value="">${escapeHtml(label)} (All)</option>`;
    sel.onchange = () => { state.filters.sel[key] = sel.value; renderTable(); };
    wrap.appendChild(sel);
  });
  document.getElementById('filterDate').hidden = !SHEET.dateKey;
  const isAnalysis = activeSheetKey === 'analysis';
  document.getElementById('needsInputWrap').hidden = !isAnalysis;
  document.getElementById('ocvImportBtn').hidden = !isAnalysis;
  document.getElementById('needsInputOnly').checked = false;
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
  document.addEventListener('paste', onDocumentPaste);
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
    const countEl = document.getElementById('addRowCount');
    const n = Math.min(1000, Math.max(1, parseInt(countEl.value, 10) || 1));
    countEl.value = n;
    try {
      await createCells(Array.from({ length: n }, () => ({})));
      showToast(`Added ${n} blank row(s). Click the first cell to fill and press Ctrl+V.`);
    } catch (err) { console.error(err); showToast('Failed to add rows.', true); }
  };
  document.getElementById('formSaveBtn').onclick = saveFormModal;
  wireModalClose('formModalOverlay', ['formModalClose', 'formCancelBtn']);
  wireModalClose('detailModalOverlay', ['detailModalClose', 'detailCloseBtn']);
  wireModalClose('lightboxOverlay', ['lightboxClose']);

  document.getElementById('detailDeleteBtn').onclick = async () => {
    const cell = state.cells.find((c) => c.id === state.detailId);
    if (!cell) return;
    if (!confirm(`Delete the row for Cell ID "${cell.cellId || '(none)'}"? This cannot be undone.`)) return;
    try {
      await deleteCell(cell);
      document.getElementById('detailModalOverlay').classList.remove('open');
      showToast('Row deleted.');
    } catch (err) { console.error(err); showToast('An error occurred while deleting.', true); }
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
    document.getElementById('needsInputOnly').checked = false;
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
    catch (err) { console.error(err); showToast('An error occurred while importing.', true); }
  });
  document.getElementById('exportCsvBtn').onclick = exportCsv;

  document.getElementById('needsInputOnly').addEventListener('change', (e) => { state.filters.needsInput = e.target.checked; renderTable(); });
  document.getElementById('ocvImportBtn').onclick = openOcvModal;
  wireModalClose('ocvModalOverlay', ['ocvModalClose', 'ocvCancelBtn']);
  const ocvDrop = document.getElementById('ocvDrop');
  document.getElementById('ocvFileInput').addEventListener('change', (e) => { loadOcvWorkbook(e.target.files[0]); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach((t) => ocvDrop.addEventListener(t, (e) => { e.preventDefault(); ocvDrop.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((t) => ocvDrop.addEventListener(t, (e) => { e.preventDefault(); ocvDrop.classList.remove('drag'); }));
  ocvDrop.addEventListener('drop', (e) => { const f = e.dataTransfer.files[0]; if (f) loadOcvWorkbook(f); });
  document.getElementById('ocvOnlyWithSheet').addEventListener('change', renderOcvOptions);
  document.getElementById('ocvLotsNone').onclick = () => { ocvState.lots.clear(); renderOcvOptions(); };
  document.getElementById('ocvImportConfirm').onclick = runOcvImport;
  document.getElementById('exportXlsxBtn').onclick = exportXlsx;

  document.getElementById('resetBtn').onclick = openResetModal;
  document.getElementById('resetConfirmInput').oninput = (e) => {
    document.getElementById('resetConfirmBtn').disabled = e.target.value.trim() !== 'RESET';
  };
  document.getElementById('resetConfirmInput').onkeydown = (e) => { if (e.key === 'Enter') confirmReset(); };
  document.getElementById('resetConfirmBtn').onclick = confirmReset;
  wireModalClose('resetModalOverlay', ['resetModalClose', 'resetCancelBtn']);

  document.getElementById('loginForm').addEventListener('submit', submitLogin);
  document.getElementById('signOutBtn').onclick = () => { firebase.auth().signOut(); };
  document.getElementById('migrateUploadBtn').onclick = uploadLocalData;
  document.getElementById('migrateDismissBtn').onclick = () => {
    try { localStorage.setItem(MIGRATION_FLAG, '1'); } catch (e) {}
    document.getElementById('migrateBanner').hidden = true;
  };

  let savedSheet = null;
  try { savedSheet = localStorage.getItem('btt.activeSheet'); } catch (e) {}
  setActiveSheet(SHEETS[savedSheet] ? savedSheet : 'teardown');
  initStorage();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();

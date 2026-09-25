// Parses workbooks off the main thread so the page stays responsive with large (30MB+) files.
importScripts(
  'https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js',
  'https://cdn.jsdelivr.net/npm/fflate@0.8.2/umd/index.js',
  'xlsx-subset.js?v=1',
  'convert.js?v=4',
  'report-model.js?v=5'
);

var source = null; // { handle (XlsxSubset workbook or null), data, names }

function findMasterSheet(names) {
  var norm = function (s) { return String(s).toLowerCase().replace(/[^a-z]/g, ''); };
  for (var i = 0; i < names.length; i++) if (norm(names[i]).indexOf('master') === 0) return names[i];
  return names[0];
}

// Open with the fast partial reader; fall back to a full SheetJS parse for unusual files (e.g. .xls)
function readSheets(names, opts) {
  var input = source.handle ? source.handle.open(names) : source.data;
  return XLSX.read(input, Object.assign({ type: 'array', sheets: names }, opts));
}

// Streams a large text file line by line (the genealogy CSV can be hundreds of MB)
async function streamLines(file, onLine, onProgress) {
  var reader = file.stream().getReader();
  var decoder = new TextDecoder();
  var carry = '', done = 0, lastReport = 0;
  for (;;) {
    var chunk = await reader.read();
    if (chunk.done) break;
    done += chunk.value.length;
    var textPart = carry + decoder.decode(chunk.value, { stream: true });
    var lines = textPart.split(String.fromCharCode(10));
    carry = lines.pop();
    for (var i = 0; i < lines.length; i++) { var ln = lines[i]; if (ln.charCodeAt(ln.length - 1) === 13) ln = ln.slice(0, -1); onLine(ln); }
    if (done - lastReport > 8 * 1024 * 1024) { lastReport = done; onProgress(done / file.size); }
  }
  carry += decoder.decode();
  if (carry) onLine(carry);
}

// ---- Generic workbook reading (position map page) ----
var book = null;
function readBookSheets(names, opts) {
  var input = book.handle ? book.handle.open(names) : book.data;
  return XLSX.read(input, Object.assign({ type: 'array', sheets: names }, opts));
}
// Rows of a sheet, trimmed to the cells that actually hold values (some sheets claim a range of
// a million rows / 16k columns because of formatting)
function sheetRows(ws) {
  if (!ws) return [];
  var dense = ws['!data'] || (Array.isArray(ws) ? ws : null);
  var maxR = -1, maxC = -1;
  if (dense) {
    for (var r = 0; r < dense.length; r++) {
      var row = dense[r];
      if (!row) continue;
      for (var c = 0; c < row.length; c++) {
        var cell = row[c];
        if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '') { if (r > maxR) maxR = r; if (c > maxC) maxC = c; }
      }
    }
  } else {
    Object.keys(ws).forEach(function (k) {
      if (k.charAt(0) === '!') return;
      var a = XLSX.utils.decode_cell(k);
      if (ws[k].v !== undefined && String(ws[k].v).trim() !== '') { if (a.r > maxR) maxR = a.r; if (a.c > maxC) maxC = a.c; }
    });
  }
  if (maxR < 0) return [];
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.min(maxR, 200000), c: Math.min(maxC, 300) } });
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
}
function normHead(v) { return String(v == null ? '' : v).toLowerCase().replace(/[\s\r\n]+/g, ''); }
function isNumLike(v) { var t = String(v == null ? '' : v).replace(/mm/i, '').replace(/,/g, '').trim(); return t !== '' && !isNaN(Number(t)); }
// Header row = first row (within 40) that has both an X and a Y column
function findXYHeader(rows) {
  for (var r = 0; r < Math.min(rows.length, 40); r++) {
    var h = (rows[r] || []).map(normHead);
    var x = h.findIndex(function (v) { return /^x(\(mm\)|좌표|coordinate|mm)?$/.test(v); });
    var y = h.findIndex(function (v) { return /^y(\(mm\)|좌표|coordinate|mm)?$/.test(v); });
    if (x >= 0 && y >= 0) return { row: r, x: x, y: y };
  }
  return null;
}

self.onmessage = function (e) {
  var msg = e.data;
  if (msg.type === 'genealogy') {
    var parser = LvReportModel.createGenealogyParser();
    streamLines(msg.file, parser.line, function (p) { self.postMessage({ id: msg.id, progress: p }); })
      .then(function () {
        var result = parser.result();
        self.postMessage({ id: msg.id, ok: true, genealogy: result, cells: Object.keys(result).length });
      })
      .catch(function (err) { self.postMessage({ id: msg.id, ok: false, error: String(err && err.message || err) }); });
    return;
  }
  try {
    if (msg.type === 'book') {
      // Any workbook: find sheets that have X / Y columns and count rows with numeric X and Y
      var bdata = new Uint8Array(msg.buf);
      var bhandle = null, bnames;
      try { bhandle = XlsxSubset.openWorkbook(fflate, bdata); bnames = bhandle.names; }
      catch (zipErr2) { bnames = XLSX.read(bdata, { type: 'array', bookSheets: true }).SheetNames; }
      book = { handle: bhandle, data: bdata, names: bnames };
      var heads = readBookSheets(bnames, { dense: true, sheetRows: 40 });
      var info = bnames.map(function (n) { return { name: n, header: findXYHeader(sheetRows(heads.Sheets[n])) }; });
      var candidates = info.filter(function (s) { return s.header; }).map(function (s) { return s.name; });
      var full = candidates.length ? readBookSheets(candidates, { dense: true }) : { Sheets: {} };
      info.forEach(function (s) {
        if (!s.header) return;
        var rows = sheetRows(full.Sheets[s.name]);
        s.rowCount = rows.length;
        s.xyCount = rows.slice(s.header.row + 1).filter(function (r) { return isNumLike(r[s.header.x]) && isNumLike(r[s.header.y]); }).length;
      });
      self.postMessage({ id: msg.id, ok: true, sheets: info });
      return;
    }
    if (msg.type === 'bookRows') {
      var rwb2 = readBookSheets([msg.sheet], { dense: true });
      self.postMessage({ id: msg.id, ok: true, rows: sheetRows(rwb2.Sheets[msg.sheet]) });
      return;
    }
    if (msg.type === 'lotstats') {
      var lwb = XLSX.read(new Uint8Array(msg.buf), { type: 'array' });
      var lrows = XLSX.utils.sheet_to_json(lwb.Sheets[lwb.SheetNames[0]], { header: 1, raw: false, defval: '' });
      self.postMessage({ id: msg.id, ok: true, stats: LvReportModel.parseLotStats(lrows).rows });
      return;
    }
    if (msg.type === 'source') {
      var data = new Uint8Array(msg.buf);
      var handle = null, names;
      try {
        handle = XlsxSubset.openWorkbook(fflate, data);
        names = handle.names;
      } catch (zipErr) {
        names = XLSX.read(data, { type: 'array', bookSheets: true }).SheetNames;
      }
      source = { handle: handle, data: data, names: names };
      var masterName = findMasterSheet(names);
      var wb = readSheets([masterName], { dense: true, cellStyles: true });
      var mws = wb.Sheets[masterName];
      var rows = XLSX.utils.sheet_to_json(mws, { header: 1, raw: false, defval: '' });
      // rows[] starts at the sheet's used range, so offset back to sheet coordinates for the fill lookup
      var start = XLSX.utils.decode_range(mws['!ref'] || 'A1').s;
      var grid = mws['!data'] || (Array.isArray(mws) ? mws : null);
      var fillAt = function (r, c) {
        var cell = grid ? (grid[r + start.r] || [])[c + start.c] : mws[XLSX.utils.encode_cell({ r: r + start.r, c: c + start.c })];
        return cell && cell.s;
      };
      var parsed = OcvConvert.parseMaster(rows, names, fillAt);
      self.postMessage({
        id: msg.id, ok: true, masterName: masterName, sheetCount: names.length, layout: parsed.layout,
        cells: parsed.cells, missingHeaders: parsed.missingHeaders, lots: OcvConvert.lotSummary(parsed.cells),
      });
    } else if (msg.type === 'tracking') {
      var result = {};
      if (msg.sheets.length) {
        var twb = readSheets(msg.sheets, { dense: true, sheetRows: 60 });
        msg.sheets.forEach(function (name) { result[name] = OcvConvert.analyzeTrackingSheet(twb.Sheets[name]); });
      }
      self.postMessage({ id: msg.id, ok: true, analysis: result });
    } else if (msg.type === 'report') {
      var rwb = XLSX.read(new Uint8Array(msg.buf), { type: 'array' });
      var found = null, sheetName = null;
      for (var i = 0; i < rwb.SheetNames.length && !found; i++) {
        var rrows = XLSX.utils.sheet_to_json(rwb.Sheets[rwb.SheetNames[i]], { header: 1, raw: false, defval: '' });
        found = OcvConvert.parsePreviousReport(rrows);
        if (found) sheetName = rwb.SheetNames[i];
      }
      self.postMessage({ id: msg.id, ok: true, report: found, sheetName: sheetName });
    }
  } catch (err) {
    self.postMessage({ id: msg.id, ok: false, error: String(err && err.message || err) });
  }
};

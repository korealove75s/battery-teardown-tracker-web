// Weekly low-voltage report: turns the source data into the numbers each slide needs.
// Pure functions shared by the browser (window.LvReportModel) and Node tests (module.exports).
(function (root) {
  'use strict';

  function text(v) { return String(v == null ? '' : v).trim(); }
  function num(v) { const n = Number(text(v)); return text(v) !== '' && isFinite(n) ? n : null; }
  function lotParts(lot) {
    const m = /^([A-Za-z]+?)([A-Za-z])0*(\d+)$/.exec(text(lot));
    return m ? { prefix: m[1].toUpperCase(), month: m[2].toUpperCase(), day: Number(m[3]) } : null;
  }
  // "FH9" / "FH09" -> "FH09"
  function padLot(lot) {
    const p = lotParts(lot);
    return p ? p.prefix + p.month + String(p.day).padStart(2, '0') : text(lot).toUpperCase();
  }
  function lotKey(lot) { return padLot(lot); }
  // Month letter of the lot code: D=April ... I=September
  function lotMonth(lot) { const p = lotParts(lot); return p ? p.month.charCodeAt(0) - 64 : null; }
  function shortLot(lot) { return padLot(lot).slice(1); } // "FH09" -> "H09" (report chart labels)
  function inRange(lot, from, to) { const k = lotKey(lot); return k >= lotKey(from) && k <= lotKey(to); }

  // ---------------------------------------------------------------------------
  // Foreign material (SEM/EDS) and location vocabulary
  // ---------------------------------------------------------------------------
  function normElement(s) {
    const t = text(s);
    if (!t) return '';
    const u = t.toUpperCase();
    if (['NTF', 'OVER.F', 'OVRE.F', 'N/A', 'NA', 'NRCF', '-', 'NONE'].includes(u)) return '';
    if (/pin\s*hole/i.test(t)) return 'Pin Hole';
    return t.split(/\s*(?:&|,|\/|\+|and)\s*/i).map((p) => p.trim()).filter(Boolean)
      .map((p) => p.charAt(0).toUpperCase() + p.slice(1).toLowerCase()).join(' & ');
  }
  function elementGroup(el) {
    if (!el) return '';
    if (/\bCu\b/.test(el)) return 'Copper';
    if (/\bFe\b/.test(el)) return 'Steel';
    return 'Other';
  }
  const GROUPS = ['Steel', 'Copper', 'Other'];
  function normLocation(s) {
    const t = text(s);
    if (/foil|al\s*surface/i.test(t)) return 'Al Foil Surface';
    if (/inside|내부/i.test(t)) return 'Coating Inside';
    if (/top|외부/i.test(t)) return 'Coating Top';
    return '';
  }
  const LOCATIONS = ['Al Foil Surface', 'Coating Inside', 'Coating Top'];
  function normTopBack(s) {
    const t = text(s).toLowerCase();
    return t === 'top' ? 'Top' : t === 'back' ? 'Back' : '';
  }
  function normDrop(s) {
    const t = text(s).toUpperCase();
    if (t === 'DROP') return 'Drop';
    if (t === 'NRCF') return 'NRCF';
    if (t === 'NTF' || t === 'NO DROP') return 'NTF';
    return '';
  }

  // Colors follow the legend of the existing report
  const ELEMENT_COLORS = {
    'Al': 'E7E6E6', 'Al & Mo': 'D9D9D9', 'Cell Missing': '7030A0', 'Cu': 'F4462B', 'Cu & Fe': 'F07C6A', 'Cu & Ni': 'F6A8A0',
    'Cu & Sn': 'F9CFC9', 'Cu & Zn': 'F48B29', 'Cu & Br': 'C0504D', 'Al & Cu': 'FFB3A7', 'Fe': '2E3FF5', 'Fe & Ni': '5B6EF5',
    'Fe & Ti': '8FA0F7', 'Fe & Cr': '4A5BD0', 'Ni': '404040', 'Ni & Al': '808080', 'Pin Hole': '1F3864',
  };
  const EXTRA_COLORS = ['9E480E', '997300', '43682B', '264478', '7F6084', '4BACC6', 'C00000', '00B050'];
  function elementColor(el, i) { return ELEMENT_COLORS[el] || EXTRA_COLORS[(i || 0) % EXTRA_COLORS.length]; }
  const LOCATION_COLORS = { 'Al Foil Surface': '404040', 'Coating Inside': '33E68C', 'Coating Top': '33A9F5' };
  const TOPBACK_COLORS = { Back: '33E68C', Top: '33A9F5' };

  // ---------------------------------------------------------------------------
  // Lot statistics (production and E/L defect counts per PKG lot) — "test2" format
  // Accepts either one text column ("FH09  0.09%  0.25%  7736  7  8  19") or separate columns.
  // ---------------------------------------------------------------------------
  function parseLotStats(rows) {
    const out = [];
    let headerSeen = false;
    rows.forEach((row) => {
      let cells = (row || []).map(text).filter((v) => v !== '');
      if (cells.length === 1) cells = cells[0].split(/\s{2,}|\t/).map(text);
      if (!cells.length) return;
      if (/count|grade|lot/i.test(cells.join(' ')) && !lotParts(cells[0])) { headerSeen = true; return; }
      if (!lotParts(cells[0]) || cells.length < 7) return;
      const pct = (v) => num(String(v).replace('%', ''));
      out.push({ lot: padLot(cells[0]), eRate: pct(cells[1]), lRate: pct(cells[2]), production: num(cells[3]) || 0, eCount: num(cells[4]) || 0, hm: num(cells[5]) || 0, lCount: num(cells[6]) || 0 });
    });
    out.sort((a, b) => a.lot.localeCompare(b.lot));
    return { rows: out, headerSeen };
  }
  function sumLots(stats, from, to) {
    const rows = stats.filter((r) => inRange(r.lot, from, to));
    const production = rows.reduce((s, r) => s + r.production, 0);
    const eCount = rows.reduce((s, r) => s + r.eCount, 0);
    const lCount = rows.reduce((s, r) => s + r.lCount, 0);
    return {
      lots: rows.length, production, eCount, lCount,
      eRate: production ? eCount / production * 100 : 0,
      lRate: production ? lCount / production * 100 : 0,
      totalRate: production ? (eCount + lCount) / production * 100 : 0,
    };
  }

  // ---------------------------------------------------------------------------
  // Genealogy CSV (cell -> mixing/coating/roll/slitting/DNC/AZS/PKG), streamed line by line
  // ---------------------------------------------------------------------------
  function createGenealogyParser() {
    let idx = null;
    const cells = new Map();
    const col = (header, name) => header.findIndex((h) => h.trim().toLowerCase() === name.toLowerCase());
    function polarityOfLot(lot) { const t = text(lot).toUpperCase(); return t.startsWith('4C') ? 'C' : t.startsWith('4A') ? 'A' : ''; }
    function line(l) {
      if (!l) return;
      const c = l.split(',');
      if (!idx) {
        idx = {
          id: col(c, 'BASE_ID'),
          mixLot: col(c, '01_Mixing_Lot ID'), coatLot: col(c, '02_Coating_Lot ID'), rollLot: col(c, '03_Roll Pressing_Lot ID'),
          slitLot: col(c, '04_Slitting_Lot ID'), dncLot: col(c, '05_DNC_Lot ID'), dncEq: col(c, '05_DNC_Equipment ID'),
          azsLot: col(c, '06_AZS Stacking_Lot ID'), azsEq: col(c, '06_AZS Stacking_Equipment ID'), azsPos: col(c, '06_AZS Stacking_Position'),
          pkgLot: col(c, '07_Packaging_Lot ID'), pkgEq: col(c, '07_Packaging_Equipment ID'), pkgDate: col(c, '07_Packaging_Finished Date'),
        };
        if (idx.id < 0 || idx.pkgLot < 0) throw new Error('This CSV does not look like the genealogy export (BASE_ID / 07_Packaging_Lot ID columns not found).');
        return;
      }
      const id = text(c[idx.id]).toUpperCase();
      if (!id) return;
      let g = cells.get(id);
      if (!g) {
        g = { azsPos: text(c[idx.azsPos]), azsLot: text(c[idx.azsLot]), azsEq: text(c[idx.azsEq]), pkgLot: text(c[idx.pkgLot]), pkgDate: text(c[idx.pkgDate]), dnc: new Set(), C: { mix: new Set(), coat: new Set(), roll: new Set(), slit: new Set() }, A: { mix: new Set(), coat: new Set(), roll: new Set(), slit: new Set() } };
        cells.set(id, g);
      }
      [['mix', idx.mixLot], ['coat', idx.coatLot], ['roll', idx.rollLot], ['slit', idx.slitLot]].forEach(([k, i]) => {
        const v = text(c[i]);
        const p = polarityOfLot(v);
        if (p) g[p][k].add(v);
      });
      const dncLot = text(c[idx.dncLot]).toUpperCase();
      const dncEq = text(c[idx.dncEq]);
      // DNC lot e.g. "3FH913NCA1" (cathode, +) / "3FH914NAN1" (anode, -)
      const tail = dncLot.slice(6);
      if (dncEq) g.dnc.add(dncEq + '|' + (tail.startsWith('NC') ? '+' : tail.startsWith('NA') ? '-' : ''));
    }
    function result() {
      const out = {};
      cells.forEach((g, id) => {
        const conv = (o) => ({ mix: [...o.mix], coat: [...o.coat], roll: [...o.roll], slit: [...o.slit] });
        out[id] = { azsPos: g.azsPos, azsLot: g.azsLot, azsEq: g.azsEq, pkgLot: g.pkgLot, pkgDate: g.pkgDate, dnc: [...g.dnc].map((s) => { const [eq, pol] = s.split('|'); return { eq, pol }; }), C: conv(g.C), A: conv(g.A) };
      });
      return out;
    }
    return { line, result };
  }

  // ---------------------------------------------------------------------------
  // Cell records: Master E & L + drop analysis + reviewed analysis rows + genealogy
  // ---------------------------------------------------------------------------
  // built: OcvConvert.buildRow() output (fields with .value) or null; review: analysis-report row (test1 format)
  function makeRecord(cell, built, review, genealogy) {
    const b = (k) => (built && built[k] ? text(built[k].value) : '');
    const r = (k) => (review ? text(review[k]) : '');
    const pick = (k) => r(k) || b(k);
    const id = text(cell.cellId).toUpperCase();
    let vd = normDrop(r('voltageDrop')) || normDrop(b('voltageDrop'));
    if (!vd) {
      const ntf = text(cell.ntf).toUpperCase();
      vd = ntf === 'NTF' ? 'NTF' : (text(cell.anodeSheet) || text(cell.voltageDrop)) ? 'Drop' : '';
    }
    const element = normElement(pick('sem') || (!built ? (cell.eds || (/^(NTF|OVER\.F)$/i.test(cell.ntf) ? '' : cell.ntf)) : ''));
    return {
      id,
      lot: padLot(cell.lot),
      voltageDrop: vd,
      element,
      group: elementGroup(element),
      location: normLocation(r('location')),
      topBack: normTopBack(pick('topBack') || (!built ? cell.topBack : '')),
      ...swapXY(num(pick('x') || (!built ? cell.x : '')), num(pick('y') || (!built ? cell.y : ''))),
      layer: num(pick('layer') || (!built ? cell.anodeSheet : '')),
      shape: r('shape'),
      reviewed: !!review,
      g: genealogy || null,
    };
  }

  // E81C electrode is 98 mm tall: Y above that means X and Y were entered the wrong way round
  const Y_MAX = 98;
  function swapXY(x, y) {
    if (x != null && y != null && y > Y_MAX && x <= Y_MAX) return { x: y, y: x, xySwapped: true };
    return { x, y, xySwapped: false };
  }
  function countBy(items, keyFn) {
    const m = new Map();
    items.forEach((it) => { const k = keyFn(it); if (k === '' || k == null) return; m.set(k, (m.get(k) || 0) + 1); });
    return [...m.entries()].sort((a, b) => b[1] - a[1] || String(a[0]).localeCompare(String(b[0])));
  }
  function pct(n, d, digits) { return d ? (n / d * 100).toFixed(digits == null ? 0 : digits) : '0'; }

  // Summary for one set of analyzed cells (slide 2 / slide 8)
  function analysisSummary(records) {
    const total = records.length;
    const dropAll = records.filter((r) => r.voltageDrop === 'Drop' || r.voltageDrop === 'NRCF');
    const genuine = records.filter((r) => r.voltageDrop === 'Drop');
    const nrcf = dropAll.length - genuine.length;
    const ntf = total - dropAll.length;
    const loc = {};
    LOCATIONS.forEach((l) => { loc[l] = genuine.filter((r) => r.location === l).length; });
    const locUnknown = genuine.filter((r) => !r.location).length;
    const elements = countBy(genuine, (r) => r.element);
    const elementsByLocation = {};
    LOCATIONS.forEach((l) => { elementsByLocation[l] = countBy(genuine.filter((r) => r.location === l), (r) => r.element); });
    const topBack = countBy(genuine, (r) => r.topBack);
    const withElement = genuine.filter((r) => r.element).length;
    return { total, dropAll: dropAll.length, genuine: genuine.length, nrcf, ntf, loc, locUnknown, elements, elementsByLocation, topBack, withElement, genuineRecords: genuine };
  }

  // Per-lot element counts split into Steel / Copper / Other panels
  function lotElementPanels(genuine, lots) {
    const panels = {};
    GROUPS.forEach((gname) => {
      const els = [...new Set(genuine.filter((r) => r.group === gname).map((r) => r.element))].sort();
      panels[gname] = els.map((el) => ({ name: el, labels: lots.map(shortLot), values: lots.map((lot) => genuine.filter((r) => r.lot === lot && r.element === el).length) }));
    });
    return panels;
  }

  // ---------------------------------------------------------------------------
  // Slide 1: monthly trend
  // ---------------------------------------------------------------------------
  const MONTH_NAMES = { 4: '4월', 5: '5월', 6: '6월', 7: '7월', 8: '8월', 9: '9월', 10: '10월', 11: '11월', 12: '12월', 1: '1월', 2: '2월', 3: '3월' };
  function monthlyTrend(stats, fromLot, toLot, locationRecords) {
    const lots = stats.filter((r) => inRange(r.lot, fromLot, toLot));
    const months = [];
    lots.forEach((r, i) => {
      const m = lotMonth(r.lot);
      let cur = months[months.length - 1];
      if (!cur || cur.month !== m) { cur = { month: m, name: MONTH_NAMES[m] || `${m}월`, first: i, last: i, lots: [] }; months.push(cur); }
      cur.last = i;
      cur.lots.push(r);
    });
    months.forEach((mo) => {
      const prod = mo.lots.reduce((s, r) => s + r.production, 0);
      const e = mo.lots.reduce((s, r) => s + r.eCount, 0);
      const l = mo.lots.reduce((s, r) => s + r.lCount, 0);
      const lotSet = new Set(mo.lots.map((r) => r.lot));
      const recs = locationRecords.filter((r) => lotSet.has(r.lot) && r.voltageDrop === 'Drop' && r.location);
      const inside = recs.filter((r) => r.location === 'Coating Inside').length;
      const outside = recs.filter((r) => r.location !== 'Coating Inside').length;
      Object.assign(mo, { production: prod, eCount: e, lCount: l, eRate: prod ? e / prod * 100 : 0, lRate: prod ? l / prod * 100 : 0, totalRate: prod ? (e + l) / prod * 100 : 0, inside, outside, located: inside + outside });
    });
    // Share of coating-inside foreign material per lot (only lots with located cells)
    const insideShare = lots.map((r) => {
      const recs = locationRecords.filter((x) => x.lot === r.lot && x.voltageDrop === 'Drop' && x.location);
      if (!recs.length) return null;
      return recs.filter((x) => x.location === 'Coating Inside').length / recs.length * 100;
    });
    return {
      lots: lots.map((r) => r.lot),
      labels: lots.map((r) => r.lot),
      eRate: lots.map((r) => (r.production ? r.eCount / r.production * 100 : 0)),
      lRate: lots.map((r) => (r.production ? r.lCount / r.production * 100 : 0)),
      production: lots.map((r) => r.production),
      insideShare,
      months,
    };
  }
  function trendHeadline(months) {
    if (!months.length) return '';
    const parts = months.map((m) => `${m.name} ${m.eRate.toFixed(2)}%`).join(', ');
    let tail = '';
    if (months.length >= 2) {
      const a = months[months.length - 2], b = months[months.length - 1];
      const d = b.eRate - a.eRate;
      tail = Math.abs(d) < 0.005 ? `, ${b.name}은 ${a.name}와 유사한 수준을 유지하고 있음.`
        : d > 0 ? `, ${b.name}은 ${a.name} 대비 ${d.toFixed(2)}%p 증가하였음.` : `, ${b.name}은 ${a.name} 대비 ${(-d).toFixed(2)}%p 감소하였음.`;
    }
    return `E81C 저전압 발생률은 ${parts}로${tail ? tail.slice(1).replace(/^ /, ' ') : '.'}`;
  }

  // ---------------------------------------------------------------------------
  // Slide 5: AZS stacker / DNC (coating-top = outside foreign material)
  // ---------------------------------------------------------------------------
  function isoWeek(dateStr) {
    const d = new Date(String(dateStr).replace(' ', 'T'));
    if (isNaN(d)) return null;
    const t = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
    const day = t.getUTCDay() || 7;
    t.setUTCDate(t.getUTCDate() + 4 - day);
    const y0 = new Date(Date.UTC(t.getUTCFullYear(), 0, 1));
    return Math.ceil(((t - y0) / 86400000 + 1) / 7);
  }
  const MONTH_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function stackerTables(allGenealogy, toLot, monthsBack) {
    // All E-grade NG cells in the genealogy file, up to the report's last PKG lot
    const cells = Object.entries(allGenealogy).map(([id, g]) => ({ id, g, lot: padLot(g.pkgLot.slice(3, 7)) }))
      .filter((c) => c.g.azsPos && lotParts(c.lot) && lotKey(c.lot) <= lotKey(toLot));
    const dated = cells.map((c) => ({ ...c, date: new Date(String(c.g.pkgDate).replace(' ', 'T')) })).filter((c) => !isNaN(c.date));
    if (!dated.length) return null;
    const maxDate = new Date(Math.max(...dated.map((c) => c.date)));
    const startMonth = new Date(maxDate.getFullYear(), maxDate.getMonth() - (monthsBack - 1), 1);
    const win = dated.filter((c) => c.date >= startMonth);
    const positions = [...new Set(win.map((c) => c.g.azsPos))].sort();
    const monthKeys = [...new Set(win.map((c) => c.date.getFullYear() * 12 + c.date.getMonth()))].sort((a, b) => a - b);
    const byMonth = monthKeys.map((k) => ({ label: MONTH_EN[k % 12], values: positions.map((p) => win.filter((c) => c.g.azsPos === p && c.date.getFullYear() * 12 + c.date.getMonth() === k).length) }));
    const weeks = [...new Set(win.map((c) => isoWeek(c.g.pkgDate)).filter(Boolean))].sort((a, b) => a - b);
    const byWeek = weeks.map((w) => ({ label: String(w), values: positions.map((p) => win.filter((c) => c.g.azsPos === p && isoWeek(c.g.pkgDate) === w).length) }));
    return { positions, byMonth, byWeek, weeks, count: win.length };
  }
  function outsideEquipment(genuineOutside) {
    const recs = genuineOutside.filter((r) => r.g);
    const positions = [...new Set(recs.map((r) => r.g.azsPos).filter(Boolean))].sort();
    const elements = [...new Set(recs.map((r) => r.element).filter(Boolean))].sort();
    const stackerByElement = elements.map((el) => ({ name: el, labels: positions, values: positions.map((p) => recs.filter((r) => r.g.azsPos === p && r.element === el).length) }));
    const dncKeys = [...new Set(recs.flatMap((r) => r.g.dnc.map((d) => d.eq + '|' + d.pol)))].sort();
    const dncLabel = (k) => { const [eq, pol] = k.split('|'); return eq.replace(/^K1A/, '') + (pol ? `(${pol})` : ''); };
    const dncByElement = elements.map((el) => ({ name: el, labels: dncKeys.map(dncLabel), values: dncKeys.map((k) => recs.filter((r) => r.element === el && r.g.dnc.some((d) => d.eq + '|' + d.pol === k)).length) }));
    const lots = [...new Set(recs.map((r) => r.lot))].sort();
    const pkgByStacker = { cols: positions, rows: lots.map((lot) => ({ label: shortLot(lot), values: positions.map((p) => recs.filter((r) => r.lot === lot && r.g.azsPos === p).length) })) };
    const lines = [...new Set(recs.map((r) => r.g.azsPos.charAt(0)))].sort();
    const pkgByLine = { cols: lines.map((l) => `1-${l}`), rows: lots.map((lot) => ({ label: shortLot(lot), values: lines.map((l) => recs.filter((r) => r.lot === lot && r.g.azsPos.charAt(0) === l).length) })) };
    return { positions, elements, stackerByElement, dncByElement, pkgByStacker, pkgByLine, count: recs.length };
  }

  // ---------------------------------------------------------------------------
  // Slides 6/7: electrode lots (C = cathode 양극, A = anode 음극)
  // ---------------------------------------------------------------------------
  function electrodeAnalysis(records, genuine, pol, lots) {
    const withG = records.filter((r) => r.g);
    const gen = genuine.filter((r) => r.g);
    const coatLots = (r) => r.g[pol].coat;
    // Share of coating batch (lot chars 3-5, e.g. "F7N") / batch lot (chars 3-7, e.g. "F7N18") per PKG lot
    function shareBy(codeFn) {
      const codes = [...new Set(withG.flatMap((r) => coatLots(r).map(codeFn)))].sort();
      const series = codes.map((code) => ({
        name: code,
        labels: lots.map(shortLot),
        values: lots.map((lot) => {
          const cellsInLot = withG.filter((r) => r.lot === lot);
          const hits = cellsInLot.reduce((s, r) => s + coatLots(r).filter((l) => codeFn(l) === code).length, 0);
          const all = cellsInLot.reduce((s, r) => s + coatLots(r).length, 0);
          return all ? hits / all * 100 : 0;
        }),
      }));
      return series;
    }
    function perLot(getLots, labelFn) {
      const allLots = [...new Set(gen.flatMap(getLots))].sort();
      const elements = [...new Set(gen.map((r) => r.element).filter(Boolean))].sort();
      const series = elements.map((el) => ({ name: el, labels: allLots.map(labelFn), values: allLots.map((l) => gen.filter((r) => r.element === el && getLots(r).includes(l)).length) }));
      const totals = allLots.map((l) => gen.filter((r) => getLots(r).includes(l)).length);
      const avg = totals.length ? totals.reduce((s, v) => s + v, 0) / totals.length : 0;
      return { lots: allLots, series, avg };
    }
    const batch = (l) => l.slice(2, 5);
    const batchLot = (l) => l.slice(2, 7);
    return {
      batchVsPkg: shareBy(batch),
      batchLotVsPkg: shareBy(batchLot),
      pkgVsForeign: perLot((r) => [r.lot], shortLot),
      slitVsForeign: perLot((r) => r.g[pol].slit, (l) => l.slice(2)),
      rollVsForeign: perLot((r) => r.g[pol].roll, (l) => l.slice(2)),
      coatVsForeign: perLot((r) => r.g[pol].coat, (l) => l.slice(2)),
    };
  }

  // ---------------------------------------------------------------------------
  // Headline sentence for the 5-analysis slides
  // ---------------------------------------------------------------------------
  function analysisHeadline(label, s, extra) {
    const top = s.elements.slice(0, 4).map(([el, n]) => `${el} ${n}셀 (${pct(n, s.genuine)}%)`).join(', ');
    return `양산 ${label} 랏 저전압 불량 ${s.total}셀${extra || ''} 분석 결과, 전압 강하된 셀은 ${s.dropAll}셀 (${pct(s.dropAll, s.total)}%)이며` +
      (top ? `, 주요 이물은 ${top} 순으로 분석 되었음` : '');
  }

  // ---------------------------------------------------------------------------
  // Assemble everything the deck needs
  //   master: parsed Master E & L cells; built: { CELLID: buildRow() } for cells with tracking sheets
  //   review: { CELLID: analysis-report row }; stats: parseLotStats().rows; genealogy: { CELLID: {...} }
  //   settings: { from, to, cumFrom, trendFrom, team, date, eCriterion, lCriterion, notes, trendHeadline, headline }
  // ---------------------------------------------------------------------------
  function assemble(input) {
    const { master, built, review, stats, genealogy, settings } = input;
    const gen = genealogy || {};
    const from = padLot(settings.from), to = padLot(settings.to), cumFrom = padLot(settings.cumFrom || from);
    const key = (c) => text(c.cellId).toUpperCase();
    const mk = (c) => makeRecord(c, built[key(c)] || null, review[key(c)] || null, gen[key(c)] || null);
    const label = `${from} ~ ${to}`;
    const kfmt = (n) => (n >= 1000 ? `${(n / 1000).toFixed(1)}K` : String(n));

    // This week's lots: cells with a tracking sheet or a reviewed analysis row
    const reportRecords = master.filter((c) => inRange(c.lot, from, to) && (c.trackingSheet || review[key(c)])).map(mk);
    const lots = [...new Set(reportRecords.map((r) => r.lot))].sort();
    const summary = analysisSummary(reportRecords);
    const lotSum = sumLots(stats, from, to);
    const funnelHeader = (lab, s) => `E81C ${lab} ${kfmt(s.production)}\nE 등급 불량률 ${s.eRate.toFixed(2)}% (판정, : ${settings.eCriterion || '2.42mV'})\nL등급 불량률 ${s.lRate.toFixed(2)}% (판정 : ${settings.lCriterion || '3.5시그마'})`;

    // Cumulative: every analyzed cell from the cumulative start lot
    const cumRecords = master.filter((c) => inRange(c.lot, cumFrom, to) && (built[key(c)] || review[key(c)] || text(c.ntf) || text(c.anodeSheet) || text(c.voltageDrop))).map(mk);
    const cumSummary = analysisSummary(cumRecords);
    const cumLots = [...new Set(cumRecords.map((r) => r.lot))].sort();
    // The cumulative range starts at the first lot that actually has analyzed cells
    const cumStart = cumLots.length && lotKey(cumLots[0]) > cumFrom ? cumLots[0] : cumFrom;
    const cumSum = sumLots(stats, cumStart, to);

    // Trend: location shares come from every reviewed row (any lot)
    const reviewedRecords = master.filter((c) => review[key(c)]).map(mk);
    const trend = monthlyTrend(stats, padLot(settings.trendFrom || cumFrom), to, reviewedRecords);
    const lastMonth = trend.months[trend.months.length - 1];
    if (lastMonth && stats.some((r) => lotMonth(r.lot) === lastMonth.month && lotKey(r.lot) > to && lotParts(r.lot).month === lotParts(to).month)) {
      lastMonth.label = `${lastMonth.name} (~${shortLot(to)})`;
    }

    return {
      label,
      cumLabel: `${cumStart} ~ ${to}`,
      fileTitle: `${(settings.date || '').replace(/[^0-9]/g, '').slice(-6)}_E81C 양산랏(${from} ~ ${to}) 저전압 현황`,
      meta: { team: settings.team || 'ESHG 품질팀', date: settings.date || '' },
      notes: settings.notes || [],
      trend,
      trendHeadline: settings.trendHeadline || trendHeadline(trend.months),
      report: {
        summary, lots, lotSum,
        funnelHeader: funnelHeader(label, lotSum),
        headline: settings.headline || analysisHeadline(label, summary),
        lotPanels: lotElementPanels(summary.genuineRecords, lots),
      },
      cumulative: {
        summary: cumSummary, lots: cumLots, lotSum: cumSum,
        funnelHeader: funnelHeader(`${cumStart} ~ ${to}`, cumSum),
        headline: settings.cumHeadline || analysisHeadline(`${cumStart} ~ ${to}`, cumSummary, ' (E등급)'),
        lotPanels: lotElementPanels(cumSummary.genuineRecords, cumLots),
      },
      stacker: stackerTables(gen, to, 2),
      outside: outsideEquipment(summary.genuineRecords.filter((r) => r.location === 'Coating Top' && r.g && r.g.azsPos)),
      cathode: electrodeAnalysis(reportRecords, summary.genuineRecords, 'C', lots),
      anode: electrodeAnalysis(reportRecords, summary.genuineRecords, 'A', lots),
      counts: { report: reportRecords.length, reviewed: reportRecords.filter((r) => r.reviewed).length, withGenealogy: reportRecords.filter((r) => r.g).length, cumulative: cumRecords.length },
    };
  }

  const api = {
    assemble,
    text, num, padLot, lotKey, lotMonth, shortLot, inRange, lotParts,
    normElement, elementGroup, normLocation, normTopBack, normDrop, GROUPS, LOCATIONS,
    ELEMENT_COLORS, elementColor, LOCATION_COLORS, TOPBACK_COLORS,
    parseLotStats, sumLots, createGenealogyParser, makeRecord, analysisSummary, lotElementPanels,
    monthlyTrend, trendHeadline, stackerTables, outsideEquipment, electrodeAnalysis, analysisHeadline, countBy, pct, isoWeek,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LvReportModel = api;
})(typeof self !== 'undefined' ? self : this);

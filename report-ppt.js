// Builds the weekly low-voltage PowerPoint (A4, same layout as the "E81C 양산랏 저전압 현황" deck)
// from a report model. Works in the browser (window.LvReportPpt, needs PptxGenJS) and in Node.
(function (root) {
  'use strict';
  const M = root.LvReportModel || (typeof require !== 'undefined' ? require('./report-model.js') : null);

  const FONT = 'Malgun Gothic';
  const W = 10.833;
  const GRAY_LINE = '7F7F7F';
  const PANEL_FILL = 'F2F2F2';

  // Static appendix data (C4 sample ~ SOP history, FA22 ~ FD15), taken from the original deck
  const HISTORY = {
    groups: [['C4', 23], ['Pre-Stress', 12], ['FC12', 1], ['FC16', 1], ['FC17', 1], ['FC18', 1], ['FC19', 1], ['Stress Run', 13]],
    lots: ['FA22', 'FA23', 'FA26', 'FA27', 'FA28', 'FA29', 'FA30', 'FB02', 'FB03', 'FB04', 'FB05', 'FB06', 'FB07', 'FB09', 'FB10', 'FB11', 'FB12', 'FB13', 'FB14', 'FB16', 'FB17', 'FB18', 'FB19', 'FB20', 'FB23', 'FB24', 'FB25', 'FB26', 'FC03', 'FC04', 'FC05', 'FC06', 'FC09', 'FC10', 'FC11', 'FC12', 'FC16', 'FC17', 'FC18', 'FC19', 'FC30', 'FC31', 'FD01', 'FD02', 'FD06', 'FD07', 'FD09', 'FD10', 'FD11', 'FD12', 'FD13', 'FD14', 'FD15'],
    e: [0.14, 0.00, 0.00, 0.26, 29.00, 0.23, 0.19, 0.12, 0.08, 0.26, 0.51, 0.22, 0.69, 0.24, 0.57, 1.83, 0.13, 0.17, 1.47, 0.00, 0.00, 0.00, 0.00, 0.15, 0.00, 0.48, 0.00, 0.09, 0.00, 12.30, 1.53, 0.00, 0.00, 0.47, 0.30, 0.00, 9.79, 0.00, 0.00, 0.61, 0.17, 1.33, 0.43, 0.22, 0.52, 0.10, 0.28, 0.48, 0.33, 0.37, 0.11, 0.24, null],
    l: [0.29, 0.60, 0.69, 0.78, 0.58, 0.52, 0.80, 0.18, 0.76, 0.68, 0.45, 0.79, 0.99, 0.74, 0.64, 0.36, 0.13, 0.31, 0.53, 0.00, 1.67, 0.96, 0.68, 0.49, 1.13, 0.36, 1.61, 0.23, 0.32, 0.00, 0.37, 1.05, 0.26, 0.53, 0.25, 1.05, 0.55, 1.46, 1.16, 0.51, 0.37, 0.03, 0.00, 0.00, 0.07, 0.00, 0.00, 0.29, 0.34, 0.22, 0.45, 0.27, null],
    prod: [693, 496, 1736, 385, 1030, 3064, 2128, 1714, 1188, 1912, 1560, 3156, 1013, 3768, 2962, 3872, 3936, 2926, 2994, 216, 240, 416, 1033, 2674, 265, 2514, 124, 3475, 628, 854, 3005, 666, 380, 3191, 1969, 476, 725, 480, 519, 978, 5171, 5937, 6549, 8516, 8736, 6937, 722, 8302, 9534, 8004, 6456, 9058, 6324],
    totalE: 0.79, totalL: 0.58, totalProd: 155607,
  };

  // Default improvement notes for the trend slide (editable on the page)
  const DEFAULT_NOTES = [
    { lot: 'FD13', text: '(DNC) Magazine 세정기\nBlow/Suction 유속 부족 개선' },
    { lot: 'FD30', text: '(DNC) 노칭/커터부 전극\n혼입 방지를 위한 칸막이 적용' },
    { lot: 'FE04', text: '(DNC) 고성능 Magnet 적용' },
    { lot: 'FE07', text: '(코터) 코터 에어건의 CDA\n밸브 재질 변경 (Cu + Zn → Zn)\n(코터) 전극 폐기물 작업공간 격리' },
    { lot: 'FF02', text: '(AZS, PKG) Magnet 적용\n5000G → 10000G 고성능' },
    { lot: 'FF19', text: '(믹서) 슬러리 이송탱크, 공급탱크\n필터 사이즈 추가 강화\n(믹서) 선분산액 필터 강화' },
    { lot: 'FF23', text: '(DNC) Magazine 세정기 구동 부\n이물받이 적용 및 구동부 코팅' },
    { lot: 'FG13', text: '(AZS) ESC 설치' },
    { lot: 'FG17', text: '(조립) 조립공정 CDA 멤브레인\n필터 설치' },
  ];

  // ---------------------------------------------------------------------------
  // Small drawing helpers
  // ---------------------------------------------------------------------------
  function txt(slide, text, o) {
    slide.addText(text, Object.assign({ fontFace: FONT, fontSize: 10, color: '000000', margin: 0, isTextBox: true, valign: 'middle' }, o));
  }
  function rect(slide, pres, o) {
    slide.addShape(pres.shapes.RECTANGLE, Object.assign({ line: { color: 'FFFFFF', width: 0 } }, o));
  }
  function hline(slide, pres, x, y, w, color, width, dash) {
    slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: color || GRAY_LINE, width: width || 1, dashType: dash || 'solid' } });
  }
  function vline(slide, pres, x, y, h, color, width, dash) {
    slide.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color: color || GRAY_LINE, width: width || 1, dashType: dash || 'solid' } });
  }
  function chrome(slide, pres, title, meta, suffix) {
    const runs = [{ text: title, options: { fontSize: suffix || title.length > 44 ? 17 : 20, bold: true } }];
    if (suffix) runs.push({ text: ' ' + suffix, options: { fontSize: 11, bold: true } });
    txt(slide, runs, { x: 0.11, y: 0.11, w: 9.55, h: 0.45, valign: 'bottom' });
    hline(slide, pres, 0, 0.56, W, GRAY_LINE, 2.25);
    txt(slide, `${meta.team}\n${meta.date}`, { x: 9.7, y: 0.02, w: 1.1, h: 0.48, fontSize: 9, align: 'center' });
  }
  function panelTitle(slide, pres, title, x, y, w) {
    rect(slide, pres, { x, y, w, h: 0.22, fill: { color: PANEL_FILL } });
    hline(slide, pres, x, y + 0.22, w, 'BFBFBF', 0.75);
    txt(slide, title, { x, y, w, h: 0.22, fontSize: 9.5, bold: true, align: 'center', color: '404040' });
  }
  // Vertical label centered on (cx, cy); the box is laid out horizontally, then rotated, so it never wraps
  function vtext(slide, label, cx, cy, len, o) {
    txt(slide, label, Object.assign({ x: cx - len / 2, y: cy - 0.08, w: len, h: 0.16, fontSize: 6, rotate: 90, align: 'center', color: '595959' }, o || {}));
  }
  function legend(slide, items, x, y, opts) {
    const o = Object.assign({ size: 6.5, step: 0.14, dot: 0.08, max: 12 }, opts);
    items.slice(0, o.max).forEach((it, i) => {
      slide.addShape('ellipse', { x, y: y + i * o.step + (o.step - o.dot) / 2, w: o.dot, h: o.dot, fill: { color: it.color }, line: { color: it.color === 'FFFFFF' || it.color === 'E7E6E6' || it.color === 'D9D9D9' ? 'A6A6A6' : it.color, width: 0.5 } });
      txt(slide, it.label, { x: x + o.dot + 0.05, y: y + i * o.step, w: 1.0, h: o.step, fontSize: o.size, color: '404040' });
    });
  }
  function elementLegend(slide, elements, x, y, opts) {
    legend(slide, elements.map((el, i) => ({ label: el, color: M.elementColor(el, i) })), x, y, opts);
  }
  function heat(v, max) {
    if (!v || !max) return 'FFFFFF';
    const t = Math.min(1, v / max);
    const lerp = (a, b) => Math.round(a + (b - a) * t);
    const r = lerp(0xF6, 0xA5), g = lerp(0xE3, 0x2A), b = lerp(0xE0, 0x2A);
    return [r, g, b].map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase();
  }

  // Electrode drawing with foreign-material points: view "plane" (x, y) or "side" (x, layer)
  const ELECTRODE = { xMax: 300, yMax: 210, layers: 37 };
  function electrodeMap(slide, pres, records, box, view) {
    const { x, y, w, h } = box;
    const tabW = w * 0.05;
    const body = { x: x + tabW, y, w: w - tabW * 2, h };
    if (view === 'side') {
      const n = 18;
      for (let i = 0; i < n; i++) hline(slide, pres, body.x, y + (i + 0.5) * h / n, body.w, i % 2 ? 'A6A6A6' : '7F7F7F', 0.75);
    } else {
      rect(slide, pres, { x, y: y + h * 0.3, w: tabW, h: h * 0.4, fill: { color: 'F2F2F2' }, line: { color: 'D9D9D9', width: 0.5 } });
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: body.x, y: body.y, w: body.w, h: body.h, rectRadius: 0.03, fill: { color: 'A6A6A6' }, line: { color: '8C8C8C', width: 0.5 } });
      rect(slide, pres, { x: body.x + body.w, y: y + h * 0.15, w: tabW, h: h * 0.4, fill: { color: 'F8CBAD' } });
    }
    const d = Math.max(0.045, Math.min(0.07, h / 12));
    records.forEach((r) => {
      if (r.x == null) return;
      const yy = view === 'side' ? r.layer : r.y;
      if (yy == null) return;
      const fx = Math.max(0, Math.min(1, r.x / ELECTRODE.xMax));
      const fy = view === 'side' ? Math.max(0, Math.min(1, (yy - 1) / (ELECTRODE.layers - 1))) : Math.max(0, Math.min(1, yy / ELECTRODE.yMax));
      const color = M.elementColor(r.element || 'Unknown');
      slide.addShape(pres.shapes.OVAL, { x: body.x + fx * body.w - d / 2, y: body.y + fy * body.h - d / 2, w: d, h: d, fill: { color: r.element ? color : 'FFFFFF' }, line: { color: '262626', width: 0.25 } });
    });
  }
  function sortedElements(records) {
    const order = Object.keys(M.ELEMENT_COLORS);
    return [...new Set(records.map((r) => r.element).filter(Boolean))].sort((a, b) => (order.indexOf(a) + 1 || 99) - (order.indexOf(b) + 1 || 99) || a.localeCompare(b));
  }

  // Pie with "Name, n (p%)" labels; legend drawn separately
  function pie(slide, pres, entries, box, colorFn, opts) {
    const total = entries.reduce((s, [, n]) => s + n, 0);
    if (!total) {
      txt(slide, '(데이터 없음)', Object.assign({ fontSize: 9, color: '7F7F7F', align: 'center' }, box));
      return;
    }
    const labels = entries.map(([k, n]) => (n / total < 0.04 ? '' : `${k}, ${n}
(${(n / total * 100).toFixed(1)}%)`));
    slide.addChart(pres.charts.PIE, [{ name: 'n', labels, values: entries.map(([, n]) => n) }], Object.assign({
      x: box.x, y: box.y, w: box.w, h: box.h,
      chartColors: entries.map(([k], i) => colorFn(k, i)),
      showLegend: false, showLabel: true, showValue: false, showPercent: false,
      dataLabelFontSize: 7, dataLabelFontFace: FONT, dataLabelColor: '262626', dataLabelPosition: 'bestFit',
      firstSliceAng: 0, dataBorder: { pt: 0.75, color: 'FFFFFF' },
    }, opts || {}));
  }
  function stackedBars(slide, pres, series, box, opts) {
    const data = series.filter((s) => s.values.some((v) => v));
    if (!data.length) { txt(slide, '(데이터 없음)', Object.assign({ fontSize: 9, color: '7F7F7F', align: 'center' }, box)); return; }
    slide.addChart(pres.charts.BAR, data, Object.assign({
      x: box.x, y: box.y, w: box.w, h: box.h, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 40,
      chartColors: data.map((s, i) => (opts && opts.colorFn ? opts.colorFn(s.name, i) : M.elementColor(s.name, i))),
      showLegend: false, catAxisLabelFontSize: 6.5, catAxisLabelFontFace: FONT, valAxisLabelFontSize: 6.5, valAxisLabelFontFace: FONT,
      valGridLine: { color: 'E7E6E6', size: 0.5 }, catGridLine: { style: 'none' },
    }, opts || {}));
  }
  function stackedWithAvg(slide, pres, block, box) {
    const data = block.series.filter((s) => s.values.some((v) => v));
    if (!data.length) { txt(slide, '(데이터 없음)', Object.assign({ fontSize: 9, color: '7F7F7F', align: 'center' }, box)); return; }
    const labels = data[0].labels;
    slide.addChart([
      { type: pres.charts.BAR, data, options: { barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 30, chartColors: data.map((s, i) => M.elementColor(s.name, i)) } },
      { type: pres.charts.LINE, data: [{ name: `Avg ${block.avg.toFixed(1)}`, labels, values: labels.map(() => Math.round(block.avg * 100) / 100) }], options: { chartColors: ['404040'], lineSize: 1, lineDash: 'dash', lineDataSymbol: 'none' } },
    ], {
      x: box.x, y: box.y, w: box.w, h: box.h, showLegend: false,
      catAxisLabelFontSize: 5.5, catAxisLabelFontFace: FONT, catAxisLabelRotate: 270, valAxisLabelFontSize: 6.5,
      valGridLine: { color: 'E7E6E6', size: 0.5 }, catGridLine: { style: 'none' },
    });
    txt(slide, `Avg(Y): ${block.avg.toFixed(1)}`, { x: box.x + box.w - 1.0, y: box.y + 0.02, w: 0.95, h: 0.15, fontSize: 6.5, color: '595959', align: 'right' });
  }
  function simpleTable(slide, rows, box, opts) {
    const o = Object.assign({ fontSize: 6.5, heat: true }, opts);
    const max = Math.max(0, ...rows.slice(1).flatMap((r) => r.slice(1, -1).map((c) => (typeof c === 'number' ? c : 0))));
    const cells = rows.map((r, ri) => r.map((c, ci) => {
      const isHeader = ri === 0 || ci === 0;
      const isTotal = ri === rows.length - 1 || ci === r.length - 1;
      const v = typeof c === 'number' ? c : null;
      return {
        text: v === null ? String(c) : (v ? String(v) : '-'),
        options: {
          fontFace: FONT, fontSize: ri === 0 && r.length > 12 ? Math.min(o.fontSize, r.length > 16 ? 4.5 : 5.2) : o.fontSize, align: ci === 0 ? 'left' : 'center', valign: 'middle', bold: isHeader && ri === 0,
          color: '262626', fill: { color: !isHeader && !isTotal && o.heat ? heat(v, max) : 'FFFFFF' },
          border: [{ type: 'none' }, { type: 'none' }, { pt: 0.5, color: 'E7E6E6' }, { type: 'none' }],
          margin: 0.01,
        },
      };
    }));
    slide.addTable(cells, { x: box.x, y: box.y, w: box.w, colW: box.colW, rowH: box.rowH || 0.16 });
  }

  // ---------------------------------------------------------------------------
  // Slide 1: trend
  // ---------------------------------------------------------------------------
  function trendSlide(pres, model) {
    const slide = pres.addSlide();
    const t = model.trend;
    txt(slide, '저전압 발생 및 개선 현황', { x: 0.03, y: 0.09, w: 9.59, h: 0.4, fontSize: 20, bold: true });
    hline(slide, pres, 0, 0.51, 4.15, '000000', 0.75);
    txt(slide, model.trendHeadline, { x: 0.12, y: 0.58, w: 10.6, h: 0.72, fontSize: 13, bold: true, valign: 'top' });

    const X0 = 0.95, CW = 9.65, n = t.lots.length || 1;
    const cx = (i) => X0 + (i + 0.5) / n * CW;
    const edge = (i) => X0 + i / n * CW;
    const layout = { x: 0, y: 0.04, w: 1, h: 0.92 };

    // Month headers and boundaries
    t.months.forEach((mo, i) => {
      const x1 = edge(mo.first), x2 = edge(mo.last + 1);
      slide.addShape(pres.shapes.RECTANGLE, { x: x1, y: 1.44, w: x2 - x1, h: 0.24, fill: { color: 'FFFFFF' }, line: { color: '000000', width: 1 } });
      txt(slide, mo.label || mo.name, { x: x1, y: 1.44, w: x2 - x1, h: 0.24, fontSize: 11, bold: true, align: 'center' });
      if (i > 0) vline(slide, pres, x1, 1.68, 3.3, '595959', 1, 'dash');
    });

    // Row labels
    [['저전압\n불량률', 1.72, 1.3], ['외부 /\n내부', 3.07, 0.52], ['개선\n사항\n및\n생산량', 3.63, 1.3]].forEach(([label, y, h]) => {
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.04, y, w: 0.52, h, rectRadius: 0.06, fill: { color: 'FFFFFF' }, line: { color: 'A6A6A6', width: 0.75 } });
      txt(slide, label, { x: 0.04, y, w: 0.52, h, fontSize: 8, align: 'center', color: label.startsWith('외부') ? 'ED7D31' : '404040' });
    });

    // Rate lines (E, L)
    const rMax = Math.max(0.5, Math.ceil(Math.max(...t.eRate, ...t.lRate, 0) * 10) / 10);
    slide.addChart(pres.charts.LINE, [
      { name: 'E', labels: t.labels, values: t.eRate.map((v) => Math.round(v * 1000) / 1000) },
      { name: 'L', labels: t.labels, values: t.lRate.map((v) => Math.round(v * 1000) / 1000) },
    ], {
      x: X0, y: 1.72, w: CW, h: 1.3, layout, chartColors: ['4472C4', 'ED7D31'], lineSize: 1.25, lineDataSymbol: 'circle', lineDataSymbolSize: 3,
      catAxisHidden: true, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: rMax, valGridLine: { color: 'E7E6E6', size: 0.5, style: 'dash' },
      showLegend: false,
    });
    const yOf = (v) => 1.72 + 1.3 * (layout.y + layout.h * (1 - v / rMax));
    for (let k = 0; k <= 5; k++) {
      const v = rMax * k / 5;
      txt(slide, String(Math.round(v * 10) / 10), { x: 0.62, y: yOf(v) - 0.06, w: 0.3, h: 0.12, fontSize: 6.5, align: 'right', color: '595959' });
    }
    txt(slide, '불량률(%)', { x: 0.62, y: 1.56, w: 0.6, h: 0.11, fontSize: 6, color: '595959' });
    legend(slide, [{ label: 'E', color: '4472C4' }, { label: 'L', color: 'ED7D31' }], 9.9, 1.74, { size: 7 });
    // E min / max labels per month
    t.months.forEach((mo) => {
      const idx = []; for (let i = mo.first; i <= mo.last; i++) idx.push(i);
      if (!idx.length) return;
      const hi = idx.reduce((a, b) => (t.eRate[b] > t.eRate[a] ? b : a));
      const lo = idx.reduce((a, b) => (t.eRate[b] < t.eRate[a] ? b : a));
      [[hi, -0.17], [lo, 0.04]].forEach(([i, dy]) => {
        txt(slide, t.eRate[i].toFixed(2), { x: cx(i) - 0.25, y: yOf(t.eRate[i]) + dy, w: 0.5, h: 0.13, fontSize: 8, align: 'center', color: '262626' });
      });
    });

    // Inside / outside share (100% stacked)
    const has = t.insideShare.map((v) => v != null);
    slide.addChart(pres.charts.BAR, [
      { name: '내부', labels: t.labels, values: t.insideShare.map((v) => (v == null ? 0 : Math.round(v * 10) / 10)) },
      { name: '외부', labels: t.labels, values: t.insideShare.map((v, i) => (has[i] ? Math.round((100 - v) * 10) / 10 : 0)) },
    ], {
      x: X0, y: 3.07, w: CW, h: 0.52, layout, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 25, chartColors: ['F4B183', 'D0CECE'],
      catAxisHidden: true, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: 'none' }, showLegend: false,
    });
    ['100%', '50%', '0%'].forEach((l, k) => txt(slide, l, { x: 0.6, y: 3.07 + 0.52 * (layout.y + layout.h * k / 2) - 0.06, w: 0.32, h: 0.12, fontSize: 6, align: 'right', color: '595959' }));

    // Production bars + improvement notes
    const pMax = Math.max(10000, Math.ceil(Math.max(...t.production, 0) / 5000) * 5000);
    slide.addChart(pres.charts.BAR, [{ name: '생산량', labels: t.labels.map((l) => l.slice(1)), values: t.production }], {
      x: X0, y: 3.63, w: CW, h: 1.62, layout: { x: 0, y: 0.03, w: 1, h: 0.72 }, barDir: 'col', barGapWidthPct: 25, chartColors: ['D0CECE'],
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: pMax, valGridLine: { style: 'none' }, showLegend: false,
      catAxisLabelFontSize: 5.5, catAxisLabelRotate: 270, catAxisLabelFrequency: 3, catAxisLabelColor: '595959', catAxisLineShow: false,
    });
    const pY = (v) => 3.63 + 1.62 * (0.03 + 0.72 * (1 - v / pMax));
    [pMax, pMax / 2, 0].forEach((v) => txt(slide, v ? `${Math.round(v / 1000)}K` : '0', { x: 0.6, y: pY(v) - 0.06, w: 0.32, h: 0.12, fontSize: 6, align: 'right', color: '595959' }));
    vtext(slide, '생산량', 0.66, pY(pMax * 0.75), 0.4);
    // Improvement notes: flag + text, stacked into rows so texts never overlap
    const NOTE_W = 1.45, LINE_H = 0.11;
    const rowsEnd = [];
    (model.notes || [])
      .map((note) => ({ note, i: t.lots.indexOf(M.padLot(note.lot)) }))
      .filter((n) => n.i >= 0)
      .sort((a, b) => a.i - b.i)
      .forEach(({ note, i }) => {
        const lines = String(note.text).split('\n');
        const x = Math.min(cx(i) + 0.03, W - NOTE_W - 0.05);
        let row = rowsEnd.findIndex((end) => end < x);
        if (row < 0) { row = rowsEnd.length; rowsEnd.push(0); }
        rowsEnd[row] = x + NOTE_W;
        const flagY = 3.66 + (row % 4) * 0.33;
        vline(slide, pres, cx(i), flagY, pY(0) - flagY, '000000', 1);
        slide.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: cx(i) - 0.01, y: flagY, w: 0.12, h: 0.1, rotate: 90, fill: { color: '5B9BD5' }, line: { color: '5B9BD5', width: 0.5 } });
        txt(slide, note.text, { x, y: flagY + 0.1, w: NOTE_W, h: LINE_H * lines.length + 0.02, fontSize: 6, bold: true, valign: 'top', color: '262626' });
      });

    // Monthly table
    const cols = t.months.length;
    const pctS = (v) => `${v.toFixed(2)}%`;
    const rows = [
      ['구분', ...t.months.map((m) => m.label || m.name)],
      ['E 등급', ...t.months.map((m) => `${m.eCount}건 (${pctS(m.eRate)})`)],
      ['L 등급', ...t.months.map((m) => `${m.lCount}건 (${pctS(m.lRate)})`)],
      ['전체 불량율(E+L)', ...t.months.map((m) => pctS(m.totalRate))],
      ['내부', ...t.months.map((m) => (m.located ? `${m.inside}건 (${(m.inside / m.located * 100).toFixed(1)}%)` : '-'))],
      ['외부', ...t.months.map((m) => (m.located ? `${m.outside}건 (${(m.outside / m.located * 100).toFixed(1)}%)` : '-'))],
      ['전체 생산량', ...t.months.map((m) => m.production.toLocaleString('en-US'))],
    ];
    const firstW = 0.9;
    const colW = [firstW, ...Array(cols).fill((10.38 - firstW) / Math.max(cols, 1))];
    slide.addTable(rows.map((r, ri) => r.map((c, ci) => ({
      text: c,
      options: { fontFace: FONT, fontSize: 8, align: 'center', valign: 'middle', fill: { color: ri === 0 ? 'D9D9D9' : ci === 0 ? 'F2F2F2' : 'FFFFFF' }, bold: ri === 0, border: { type: 'solid', pt: 0.5, color: 'A6A6A6' } },
    }))), { x: 0.19, y: 5.3, w: 10.38, colW, rowH: [0.26, 0.26, 0.26, 0.3, 0.26, 0.26, 0.28] });
  }

  // ---------------------------------------------------------------------------
  // Slide 2 / 8: "5대 분석 검사 결과"
  // ---------------------------------------------------------------------------
  function funnel(slide, pres, s, header, x, y) {
    txt(slide, header, { x, y, w: 2.9, h: 0.42, fontSize: 7.5, color: '548235', valign: 'top' });
    const top = y + 0.45, H = 1.6;
    const box = (bx, by, bw, bh, label, fill, size) => {
      slide.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: fill }, line: { color: '595959', width: 1 } });
      txt(slide, label, { x: bx, y: by, w: bw, h: bh, fontSize: size || 9, align: 'center', color: '262626' });
    };
    box(x + 0.05, top, 0.6, H, `E등급\n\n${s.total} 셀\n(100%)`, 'FFFFFF', 9.5);
    // Keep both halves tall enough for their text
    const dropH = s.total ? Math.min(H - 0.45, Math.max(0.55, H * s.dropAll / s.total)) : H / 2;
    box(x + 0.7, top, 0.62, dropH, `냉동저항\nNG\n${s.dropAll}셀\n(${M.pct(s.dropAll, s.total)}%)`, 'D9D9D9', 8);
    box(x + 0.7, top + dropH, 0.62, H - dropH, `NTF\n${s.ntf}셀\n(${M.pct(s.ntf, s.total)}%)`, 'D9D9D9', 8);
    box(x + 1.37, top, 0.62, dropH, `진성불량\n${s.genuine}셀\n(${M.pct(s.genuine, s.dropAll)}%)`, 'A6A6A6', 8);
    if (s.nrcf) txt(slide, `NRCF ${s.nrcf}셀 (${M.pct(s.nrcf, s.dropAll)}%)`, { x: x + 1.34, y: top + dropH + 0.02, w: 0.7, h: 0.16, fontSize: 6.5, align: 'center', color: '404040' });
    // Location breakdown of genuine cells (each band at least one text line tall)
    const parts = [
      ['코팅층 외부', s.loc['Coating Top'], 'DEEBF7'],
      ['코팅층 내부', s.loc['Coating Inside'], 'E2F0D9'],
      ['Al Foil Surface', s.loc['Al Foil Surface'], 'F4B183'],
      ['위치 미확인', s.locUnknown, 'F2F2F2'],
    ].filter((p) => p[1]);
    const minH = 0.2;
    const free = Math.max(0, dropH - minH * parts.length);
    let cy = top;
    parts.forEach(([label, n, fill]) => {
      const h = minH + (s.genuine ? free * n / s.genuine : 0);
      const oneLine = h < 0.34;
      box(x + 2.04, cy, 0.86, h, oneLine ? `${label} ${n}셀 (${M.pct(n, s.genuine)}%)` : `${label}\n${n}셀 (${M.pct(n, s.genuine)}%)`, fill, oneLine ? 5.5 : 7);
      cy += h;
    });
    txt(slide, `총 ${s.genuine}셀`, { x: x + 2.04, y: cy + 0.03, w: 0.86, h: 0.2, fontSize: 9, align: 'center' });
  }
  function analysisSlide(pres, model, block, titleText, suffix) {
    const slide = pres.addSlide();
    chrome(slide, pres, titleText, model.meta, suffix);
    const s = block.summary;
    txt(slide, block.headline, { x: 0.11, y: 0.66, w: 10.71, h: 0.6, fontSize: 13.5, valign: 'top' });
    funnel(slide, pres, s, block.funnelHeader, 0.05, 1.28);

    const gen = s.genuineRecords;
    const els = sortedElements(gen);
    panelTitle(slide, pres, '[이물 위치]', 3.0, 1.35, 3.6);
    electrodeMap(slide, pres, gen, { x: 3.05, y: 1.75, w: 2.7, h: 1.25 }, 'plane');
    elementLegend(slide, els, 5.85, 1.7);
    panelTitle(slide, pres, '[이물 측면 위치]', 6.7, 1.35, 4.1);
    electrodeMap(slide, pres, gen, { x: 6.75, y: 1.75, w: 2.95, h: 1.25 }, 'side');
    elementLegend(slide, els, 9.8, 1.7);

    panelTitle(slide, pres, `[${s.genuine} 셀 분석]`, 0.02, 3.42, 3.5);
    pie(slide, pres, s.elements, { x: 0.1, y: 3.68, w: 2.6, h: 1.62 }, (k, i) => M.elementColor(k, i));
    elementLegend(slide, s.elements.map(([k]) => k), 2.72, 3.72, { size: 6, step: 0.13 });

    const locs = M.LOCATIONS.filter((l) => s.loc[l]);
    const locTitle = locs.map((l) => ({ 'Al Foil Surface': `Al Foil ${s.loc[l]}셀`, 'Coating Inside': `코팅층 내부 ${s.loc[l]}셀`, 'Coating Top': `외부 ${s.loc[l]}셀` }[l])).join(', ');
    panelTitle(slide, pres, `[${locTitle || '위치 정보 없음'} 분석]`, 3.6, 3.42, 3.55);
    // Up to 2 locations: one large pie per row; 3: two small on top, one below (like the original)
    const slots = locs.length <= 2
      ? [{ x: 3.65, y: 3.66, w: 2.7, h: 1.7 }, { x: 3.65, y: 5.42, w: 2.7, h: 1.62 }]
      : [{ x: 3.62, y: 3.66, w: 1.55, h: 1.7 }, { x: 5.12, y: 3.66, w: 1.55, h: 1.7 }, { x: 3.65, y: 5.42, w: 2.7, h: 1.62 }];
    locs.forEach((l, i) => {
      const b = slots[i];
      txt(slide, l, { x: b.x, y: b.y, w: b.w, h: 0.14, fontSize: 6.5, align: 'center', color: '595959' });
      pie(slide, pres, s.elementsByLocation[l], { x: b.x, y: b.y + 0.14, w: b.w, h: b.h - 0.14 }, (k, j) => M.elementColor(k, j), { dataLabelFontSize: 6.5 });
    });
    if (!locs.length) txt(slide, 'Analysis 데이터에 Location 값이 없습니다.', { x: 3.62, y: 4.5, w: 3.5, h: 0.3, fontSize: 8, align: 'center', color: '7F7F7F' });
    elementLegend(slide, els, 6.45, 3.72, { size: 6, step: 0.13 });

    panelTitle(slide, pres, '[CT 결과]', 7.22, 3.42, 3.6);
    pie(slide, pres, locs.map((l) => [l, s.loc[l]]), { x: 7.3, y: 3.68, w: 2.5, h: 1.62 }, (k) => M.LOCATION_COLORS[k] || 'A6A6A6');
    legend(slide, locs.map((l) => ({ label: l, color: M.LOCATION_COLORS[l] })), 9.82, 3.72);

    panelTitle(slide, pres, '[LOT 별 주요 이물]', 0.02, 5.4, 3.5);
    const panels = block.lotPanels;
    M.GROUPS.forEach((g, i) => {
      const by = 5.66 + i * 0.6;
      stackedBars(slide, pres, panels[g], { x: 0.05, y: by, w: 2.55, h: i < 2 ? 0.6 : 0.72 }, { catAxisHidden: i < 2, catAxisLabelFontSize: 5, catAxisLabelRotate: block.lots.length > 8 ? 270 : 0, valAxisLabelFontSize: 5.5 });
      vtext(slide, g, 2.68, by + 0.3, 0.5);
    });
    elementLegend(slide, els, 2.8, 5.68, { size: 6, step: 0.12 });

    panelTitle(slide, pres, '[이물 위치 Top/Back 개수]', 7.22, 5.4, 3.6);
    pie(slide, pres, s.topBack, { x: 7.3, y: 5.64, w: 2.5, h: 1.6 }, (k) => M.TOPBACK_COLORS[k] || 'A6A6A6');
    legend(slide, s.topBack.map(([k]) => ({ label: k, color: M.TOPBACK_COLORS[k] || 'A6A6A6' })), 9.82, 5.68);
    return slide;
  }

  // ---------------------------------------------------------------------------
  // Slide 3: foreign material location by group / top-back and inside-outside
  // ---------------------------------------------------------------------------
  function locationSlide(pres, model) {
    const slide = pres.addSlide();
    chrome(slide, pres, `3. E81C 저전압 (${model.label}) 원인 파악 – 이물 발견 위치`, model.meta);
    const gen = model.report.summary.genuineRecords;
    const els = sortedElements(gen);
    const grid = (title, x, view) => {
      panelTitle(slide, pres, title, x, 1.05, 5.1);
      ['Back', 'Top'].forEach((tb, c) => txt(slide, tb, { x: x + 0.05 + c * 2.25, y: 1.3, w: 2.2, h: 0.14, fontSize: 6.5, align: 'center', color: '595959' }));
      M.GROUPS.forEach((g, r) => {
        ['Back', 'Top'].forEach((tb, c) => {
          electrodeMap(slide, pres, gen.filter((x2) => x2.group === g && x2.topBack === tb), { x: x + 0.1 + c * 2.25, y: 1.48 + r * 0.8, w: 2.1, h: 0.68 }, view);
        });
        vtext(slide, g, x + 4.62, 1.48 + r * 0.8 + 0.34, 0.66);
      });
      elementLegend(slide, els, x + 4.85, 1.35, { size: 6 });
    };
    grid('[이물 위치 Top/Back View]', 0.02, 'plane');
    grid('[이물 위치 Top/Back Side View]', 5.45, 'side');
    const inout = (title, x, view) => {
      panelTitle(slide, pres, title, x, 4.05, 5.1);
      M.LOCATIONS.forEach((l, r) => {
        electrodeMap(slide, pres, gen.filter((x2) => x2.location === l), { x: x + 0.9, y: 4.38 + r * 0.95, w: 3.0, h: 0.8 }, view);
        vtext(slide, l, x + 4.2, 4.38 + r * 0.95 + 0.4, 0.9);
      });
      elementLegend(slide, els, x + 4.85, 4.35, { size: 6 });
    };
    inout('[내/외부 이물 위치]', 0.02, 'plane');
    inout('[내/외부 이물 측면 위치]', 5.45, 'side');
  }

  // Slide 4: small multiples per element, Top / Back
  function elementSlide(pres, model) {
    const slide = pres.addSlide();
    chrome(slide, pres, `3. E81C 저전압 (${model.label}) 원인 파악 – 이물 발견 위치`, model.meta);
    const gen = model.report.summary.genuineRecords;
    const block = (title, x, y, tb, view) => {
      panelTitle(slide, pres, title, x, y, 5.1);
      const recs = gen.filter((r) => r.topBack === tb);
      const els = sortedElements(recs);
      els.slice(0, 8).forEach((el, i) => {
        const bx = x + 0.05 + (i % 2) * 2.2, by = y + 0.27 + Math.floor(i / 2) * 0.72;
        txt(slide, el, { x: bx, y: by, w: 2.1, h: 0.13, fontSize: 6.5, align: 'center', color: '595959' });
        electrodeMap(slide, pres, recs.filter((r) => r.element === el), { x: bx + 0.05, y: by + 0.14, w: 2.0, h: 0.52 }, view);
      });
      if (!els.length) txt(slide, '(해당 셀 없음)', { x, y: y + 1, w: 4.5, h: 0.3, fontSize: 8, align: 'center', color: '7F7F7F' });
      elementLegend(slide, els, x + 4.5, y + 0.3, { size: 6 });
    };
    block('[이물 위치 Top View]', 0.02, 1.1, 'Top', 'plane');
    block('[이물 측면 위치 Top View]', 5.45, 1.1, 'Top', 'side');
    block('[이물 위치 Back View]', 0.02, 4.15, 'Back', 'plane');
    block('[이물 측면 위치 Back View]', 5.45, 4.15, 'Back', 'side');
  }

  // ---------------------------------------------------------------------------
  // Slide 5: coating-top (outside) analysis — AZS stacker / DNC
  // ---------------------------------------------------------------------------
  function outsideSlide(pres, model) {
    const slide = pres.addSlide();
    chrome(slide, pres, `4. E81C 저전압 (${model.label}) 원인 파악 – 코팅층 외부 분석`, model.meta);
    const st = model.stacker, oe = model.outside;
    if (st) {
      const tableRows = (list, firstLabel) => {
        const totals = st.positions.map((_, i) => list.reduce((s, r) => s + r.values[i], 0));
        return [
          [firstLabel, ...st.positions, 'Grand total'],
          ...list.map((r) => [r.label, ...r.values, r.values.reduce((a, b) => a + b, 0)]),
          ['Grand total', ...totals, totals.reduce((a, b) => a + b, 0)],
        ];
      };
      const colW = [0.5, ...st.positions.map(() => (3.1 / st.positions.length)), 0.45];
      panelTitle(slide, pres, '[AZS] 월별 E 불량 개수', 0.02, 1.1, 4.1);
      simpleTable(slide, tableRows(st.byMonth, 'Month'), { x: 0.05, y: 1.38, w: 4.05, colW, rowH: 0.16 });
      panelTitle(slide, pres, '[AZS] 주별 E 불량 개수', 0.02, 2.15, 4.1);
      simpleTable(slide, tableRows(st.byWeek, 'Week'), { x: 0.05, y: 2.43, w: 4.05, colW, rowH: 0.16 });
      // Stacker trend: one tiny line per stacker
      panelTitle(slide, pres, '[AZS] 스태커 트랜드', 0.02, 4.1, 4.1);
      const labels = st.byWeek.map((r) => r.label);
      st.positions.slice(0, 18).forEach((p, i) => {
        const y = 4.38 + i * (2.9 / Math.max(st.positions.length, 12));
        const h = 2.9 / Math.max(st.positions.length, 12);
        txt(slide, p, { x: 0.08, y, w: 0.4, h, fontSize: 6, color: '404040' });
        const vals = st.byWeek.map((r) => r.values[i]);
        slide.addChart(pres.charts.LINE, [{ name: p, labels, values: vals }], {
          x: 0.55, y, w: 3.0, h, chartColors: ['7F7F7F'], lineSize: 0.75, lineDataSymbol: 'none', catAxisHidden: true, valAxisHidden: true,
          valGridLine: { style: 'none' }, showLegend: false, layout: { x: 0, y: 0.1, w: 1, h: 0.8 },
        });
        txt(slide, String(vals[vals.length - 1] || 0), { x: 3.6, y, w: 0.3, h, fontSize: 6, align: 'right', color: '404040' });
      });
    } else {
      txt(slide, '공정 이력(CSV) 데이터가 없어 AZS 분석을 만들 수 없습니다.', { x: 0.1, y: 2, w: 4, h: 0.4, fontSize: 9, color: '7F7F7F' });
    }
    if (oe && oe.count) {
      panelTitle(slide, pres, '[AZS] 스태커별 셀 수량', 4.2, 1.1, 3.4);
      stackedBars(slide, pres, oe.stackerByElement, { x: 4.22, y: 1.38, w: 2.75, h: 2.55 }, { catAxisLabelFontSize: 5.5, catAxisLabelRotate: oe.positions.length > 10 ? 270 : 0 });
      elementLegend(slide, oe.elements, 7.0, 1.45, { size: 6 });
      panelTitle(slide, pres, '[DNC] DNC 장비 구분', 7.7, 1.1, 3.12);
      stackedBars(slide, pres, oe.dncByElement, { x: 7.72, y: 1.38, w: 2.45, h: 2.55 }, { catAxisLabelFontSize: 6 });
      elementLegend(slide, oe.elements, 10.2, 1.45, { size: 6 });
      const tbl = (t, x, w, first) => {
        const totals = t.cols.map((_, i) => t.rows.reduce((s, r) => s + r.values[i], 0));
        const rows = [[first, ...t.cols, 'Grand total'], ...t.rows.map((r) => [r.label, ...r.values, r.values.reduce((a, b) => a + b, 0)]), ['Grand total', ...totals, totals.reduce((a, b) => a + b, 0)]];
        simpleTable(slide, rows, { x, y: 4.38, w, colW: [0.45, ...t.cols.map(() => (w - 0.9) / t.cols.length), 0.45], rowH: 0.17 });
      };
      panelTitle(slide, pres, '[AZS] 스태커별 셀 수량', 4.2, 4.1, 3.4);
      tbl(oe.pkgByStacker, 4.22, 3.35, 'PKG Lot');
      panelTitle(slide, pres, '[AZS] 라인 구분', 7.7, 4.1, 3.12);
      tbl(oe.pkgByLine, 7.72, 2.2, 'PKG Lot');
    } else {
      txt(slide, '코팅층 외부(Coating Top) 셀의 공정 이력이 없습니다.', { x: 4.3, y: 2, w: 6, h: 0.4, fontSize: 9, color: '7F7F7F' });
    }
  }

  // ---------------------------------------------------------------------------
  // Slides 6/7: electrode (cathode / anode)
  // ---------------------------------------------------------------------------
  const BATCH_COLORS = ['9C7F1F', '98239B', '5B7FF5', 'F47C62', 'B679F2', '8FB4F9', 'FFE699', 'A9F5C7', 'F9A0F0', '5B9BD5', 'ED7D31', '70AD47', 'FFC000', '4472C4', '264478', '9E480E', '636363', '997300'];
  function electrodeSlide(pres, model, e, title) {
    const slide = pres.addSlide();
    chrome(slide, pres, title, model.meta);
    const share = (series, box) => {
      const data = series.filter((s) => s.values.some((v) => v));
      if (!data.length) { txt(slide, '(데이터 없음)', Object.assign({ fontSize: 9, color: '7F7F7F', align: 'center' }, box)); return; }
      slide.addChart(pres.charts.BAR, data.map((s) => ({ name: s.name, labels: s.labels, values: s.values.map((v) => Math.round(v * 10) / 10) })), {
        x: box.x, y: box.y, w: box.w, h: box.h, barDir: 'col', barGrouping: 'percentStacked', barGapWidthPct: 40,
        chartColors: data.map((_, i) => BATCH_COLORS[i % BATCH_COLORS.length]), showLegend: true, legendPos: 'r', legendFontSize: 6, legendFontFace: FONT,
        catAxisLabelFontSize: 6.5, valAxisLabelFontSize: 6.5, valAxisLabelFormatCode: '0%', valGridLine: { color: 'E7E6E6', size: 0.5 }, catGridLine: { style: 'none' },
      });
    };
    panelTitle(slide, pres, `[전극] 날짜별 ${e.name} 코팅 배치 vs PKG LOT`, 0.02, 1.05, 5.25);
    share(e.data.batchVsPkg, { x: 0.05, y: 1.32, w: 5.2, h: 2.6 });
    panelTitle(slide, pres, `[전극] ${e.name} 코팅 배치 LOT vs PKG LOT`, 0.02, 4.05, 5.25);
    share(e.data.batchLotVsPkg, { x: 0.05, y: 4.32, w: 5.2, h: 2.95 });
    const els = sortedElements(model.report.summary.genuineRecords);
    [['PKG LOT', e.data.pkgVsForeign], ['슬리팅 LOT', e.data.slitVsForeign], ['롤프레싱 LOT', e.data.rollVsForeign], ['코팅 LOT', e.data.coatVsForeign]].forEach(([label, block], i) => {
      const y = 1.05 + i * 1.57;
      panelTitle(slide, pres, `[전극] ${e.name} ${label} vs 이물`, 5.4, y, 5.42);
      stackedWithAvg(slide, pres, block, { x: 5.42, y: y + 0.24, w: 4.6, h: 1.3 });
      elementLegend(slide, els, 10.1, y + 0.3, { size: 5.5, step: 0.12 });
    });
  }

  // ---------------------------------------------------------------------------
  // Slide 9: static history appendix
  // ---------------------------------------------------------------------------
  function historySlide(pres, model) {
    const slide = pres.addSlide();
    chrome(slide, pres, '유첨. E81C 저전압 발생 현황', { team: model.meta.team, date: '26. 4. 25' });
    const h = HISTORY;
    const labels = h.lots;
    slide.addChart([
      { type: pres.charts.BAR, data: [{ name: '생산량', labels, values: h.prod }], options: { barDir: 'col', chartColors: ['A6A6A6'], barGapWidthPct: 60, secondaryValAxis: true, secondaryCatAxis: true } },
      { type: pres.charts.LINE, data: [{ name: '불량률 E', labels, values: h.e.map((v) => (v == null ? 0 : v)) }, { name: '불량률 L', labels, values: h.l.map((v) => (v == null ? 0 : v)) }], options: { chartColors: ['5B9BD5', 'ED7D31'], lineSize: 2, lineDataSymbol: 'none' } },
    ], {
      x: 0.1, y: 1.15, w: 10.6, h: 4.15, showLegend: true, legendPos: 't', legendFontSize: 9, legendFontFace: FONT,
      showTitle: true, title: 'E81C 저전압 불량률 (By PKG Lot, Lot별 dOCV 판정)', titleFontSize: 12, titleFontFace: FONT, titleColor: '595959',
      catAxisLabelFontSize: 6.5, catAxisLabelRotate: 270, valGridLine: { color: 'E7E6E6', size: 0.5 },
      valAxes: [{ showValAxisTitle: false, valAxisMinVal: 0, valAxisMaxVal: 13, valAxisLabelFormatCode: '0.0"%"', valAxisLabelFontSize: 8 }, { showValAxisTitle: false, valAxisMinVal: 0, valAxisMaxVal: 14000, valAxisLabelFontSize: 8, valGridLine: { style: 'none' } }],
      catAxes: [{ catAxisTitle: 'LOT' }, { catAxisHidden: true }],
    });
    [['C4 샘플 (FA22 ~ FB19) 생산량 : 42.4k\nE 등급: 0.26%', 1.0], ['Stress Run 생산량 : 95k (~ FD15)\nE 등급 : 0.21%', 6.9], ['양산 초도\n- FD10 : E 등급 0.48%\n- FD11 : E 등급 0.33%\n- FD12 : E 등급 0.37%\n- FD13 : E 등급 0.11%', 8.9]].forEach(([t, x]) => {
      txt(slide, t, { x, y: 1.62, w: 1.9, h: 0.75, fontSize: 8.5, bold: true, valign: 'top' });
    });
    // Table
    const cellOpt = (ri, extra) => Object.assign({ fontFace: FONT, fontSize: 4.5, align: 'center', valign: 'middle', fill: { color: ri < 2 ? 'F2F2F2' : 'FFFFFF' }, border: { type: 'solid', pt: 0.5, color: 'A6A6A6' }, margin: 0 }, extra || {});
    const head1 = [{ text: 'LOT', options: cellOpt(0, { rowspan: 2 }) }];
    h.groups.forEach(([g, n]) => head1.push({ text: g, options: cellOpt(0, { colspan: n }) }));
    head1.push({ text: 'Total', options: cellOpt(0, { rowspan: 2 }) });
    const head2 = h.lots.map((l) => ({ text: l, options: cellOpt(1) }));
    const rows = [
      ['불량률 E', ...h.e.map((v) => (v == null ? '-' : v.toFixed(2))), h.totalE.toFixed(2)],
      ['불량률 L', ...h.l.map((v) => (v == null ? '-' : v.toFixed(2))), h.totalL.toFixed(2)],
      ['생산량', ...h.prod.map((v) => v.toLocaleString('en-US')), h.totalProd.toLocaleString('en-US')],
    ];
    const n = h.lots.length;
    const colW = [0.45, ...Array(n).fill((10.46 - 0.45 - 0.4) / n), 0.4];
    const body = rows.map((r, ri) => r.map((c) => ({ text: String(c), options: cellOpt(ri + 2) })));
    slide.addTable([head1, head2, ...body], { x: 0.19, y: 5.45, w: 10.46, colW, rowH: [0.2, 0.3, 0.33, 0.33, 0.33] });
    // Improvement notes and phase markers from the original deck
    [['[조립] 스택 버퍼 비산 방지 차폐 적용 완료 (2/4)', 1.05, 3.0], ['[전극] Mixer 활물질 투입 호퍼 Suction 외 4건 완료 (2/23~25)', 3.64, 3.15], ['Mixer/Coater/조립 설비 7건 집중 개선 (2/27)', 5.16, 2.81], ['[조립] 노칭 후 세정 유닛 추가 완료외 1건 (3/31)', 7.03, 2.57]].forEach(([t, x, y]) => {
      txt(slide, t, { x, y, w: 3.8, h: 0.24, fontSize: 8, color: 'FF0000' });
    });
  }

  // ---------------------------------------------------------------------------
  // Build the whole deck
  // ---------------------------------------------------------------------------
  function build(PptxGenJS, model) {
    const pres = new PptxGenJS();
    pres.defineLayout({ name: 'A4R', width: W, height: 7.5 });
    pres.layout = 'A4R';
    pres.author = model.meta.team;
    pres.title = model.fileTitle;
    trendSlide(pres, model);
    analysisSlide(pres, model, model.report, `2. E81C 저전압 (${model.label}) 원인 파악 – 저전압 5대 분석 검사 결과`);
    locationSlide(pres, model);
    elementSlide(pres, model);
    outsideSlide(pres, model);
    electrodeSlide(pres, model, { name: '양극', data: model.cathode }, `5. E81C 저전압 (${model.label}) 원인 파악 – 전극 (양극)`);
    electrodeSlide(pres, model, { name: '음극', data: model.anode }, `5. E81C 저전압 (${model.label}) 원인 파악 – 전극 (음극)`);
    analysisSlide(pres, model, model.cumulative, '유첨. 양산 샘플 저전압 (누적) 원인 파악 – 저전압 5대 분석 검사 결과', `(${model.cumLabel})`);
    historySlide(pres, model);
    return pres;
  }

  const api = { build, DEFAULT_NOTES, HISTORY };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LvReportPpt = api;
})(typeof self !== 'undefined' ? self : this);

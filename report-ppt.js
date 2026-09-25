// Builds the weekly low-voltage PowerPoint (A4, same slide order and panels as the "E81C 양산랏 저전압 현황" deck)
// from a report model. Works in the browser (window.LvReportPpt, needs PptxGenJS + JSZip) and in Node.
(function (root) {
  'use strict';
  const M = root.LvReportModel || (typeof require !== 'undefined' ? require('./report-model.js') : null);

  const FONT = 'Malgun Gothic';
  const W = 10.833;
  const MX = 0.18;               // page side margin
  const R = W - MX;              // right edge of the content area
  const C = {
    ink: '1F1F1F', sub: '595959', mute: '8C8C8C',
    navy: '1F3864', accent: '2F5597', line: 'D0D7E2', panel: 'EEF2F8', grid: 'E4E8EF',
    e: '2F5597', l: 'ED7D31', inside: 'F08C4B', outside: 'D5DAE3', bar: 'C9CED8', green: '3A7D2C',
  };
  const PT_H = 0.27;             // panel title height

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
    slide.addText(text, Object.assign({ fontFace: FONT, fontSize: 10, color: C.ink, margin: 0, isTextBox: true, valign: 'middle' }, o));
  }
  function rect(slide, pres, o) {
    slide.addShape(pres.shapes.RECTANGLE, Object.assign({ line: { color: 'FFFFFF', width: 0 } }, o));
  }
  function hline(slide, pres, x, y, w, color, width, dash) {
    slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: color || C.line, width: width || 1, dashType: dash || 'solid' } });
  }
  function vline(slide, pres, x, y, h, color, width, dash) {
    slide.addShape(pres.shapes.LINE, { x, y, w: 0, h, line: { color: color || C.line, width: width || 1, dashType: dash || 'solid' } });
  }
  // Slide title with a navy accent bar, the team / date on the right and a rule underneath
  function chrome(slide, pres, title, meta, suffix) {
    rect(slide, pres, { x: MX, y: 0.2, w: 0.07, h: 0.38, fill: { color: C.navy } });
    const size = suffix || title.length > 40 ? 17 : 20;
    const runs = [{ text: title, options: { fontSize: size, bold: true } }];
    if (suffix) runs.push({ text: ' ' + suffix, options: { fontSize: 11, bold: true, color: C.sub } });
    txt(slide, runs, { x: MX + 0.16, y: 0.14, w: R - MX - 1.8, h: 0.5 });
    if (meta) txt(slide, `${meta.team}\n${meta.date}`, { x: R - 1.55, y: 0.16, w: 1.55, h: 0.44, fontSize: 9, align: 'right', color: C.sub });
    hline(slide, pres, MX, 0.68, R - MX, C.navy, 1.25);
  }
  function panelTitle(slide, pres, title, x, y, w) {
    rect(slide, pres, { x, y, w, h: PT_H, fill: { color: C.panel } });
    rect(slide, pres, { x, y, w: 0.05, h: PT_H, fill: { color: C.navy } });
    txt(slide, title, { x: x + 0.13, y, w: w - 0.18, h: PT_H, fontSize: 10, bold: true, color: C.navy });
  }
  function noData(slide, text, box) {
    txt(slide, text || '(데이터 없음)', Object.assign({ fontSize: 9, color: C.mute, align: 'center' }, box));
  }
  function legend(slide, items, x, y, opts) {
    const o = Object.assign({ size: 7.5, step: 0.17, dot: 0.1, max: 12, w: 0.95, maxH: 0 }, opts);
    const list = items.slice(0, o.max);
    if (o.maxH && list.length * o.step > o.maxH) {
      o.step = Math.max(0.12, o.maxH / list.length);
      o.size = Math.min(o.size, o.step < 0.15 ? 6.5 : 7);
      o.dot = Math.min(o.dot, o.step * 0.6);
    }
    list.slice(0, o.maxH ? Math.floor(o.maxH / o.step + 0.01) : list.length).forEach((it, i) => {
      const pale = ['FFFFFF', 'E7E6E6', 'D9D9D9'].includes(it.color);
      slide.addShape('ellipse', { x, y: y + i * o.step + (o.step - o.dot) / 2, w: o.dot, h: o.dot, fill: { color: it.color }, line: { color: pale ? 'A6A6A6' : it.color, width: 0.5 } });
      txt(slide, it.label, { x: x + o.dot + 0.05, y: y + i * o.step, w: o.w, h: o.step, fontSize: o.size, color: C.sub });
    });
  }
  function elementLegend(slide, elements, x, y, opts) {
    legend(slide, elements.map((el, i) => ({ label: el, color: M.elementColor(el, i) })), x, y, opts);
  }
  function heat(v, max) {
    if (!v || !max) return 'FFFFFF';
    const t = Math.min(1, v / max);
    const lerp = (a, b) => Math.round(a + (b - a) * t);
    const r = lerp(0xFB, 0xC0), g = lerp(0xE5, 0x39), b = lerp(0xDD, 0x2B);
    return [r, g, b].map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  function heatText(v, max) { return v && max && v / max > 0.6 ? 'FFFFFF' : '262626'; }

  // Electrode drawing with foreign-material points: view "plane" (x, y) or "side" (x, layer)
  // E81C electrode: X 0-320 mm, Y 0-98 mm, 37 anode layers
  const ELECTRODE = { xMax: 320, yMax: 98, layers: 37 };
  // Top view: origin (0, 0) is the bottom-left corner of the electrode, drawn at the real X:Y proportion;
  // side view: X across, layer 1 at the top.
  function electrodeMap(slide, pres, records, box, view, dims) {
    const D = Object.assign({}, ELECTRODE, dims || {});
    const { x, y, w, h } = box;
    let tabW = w * 0.05;
    let body = { x: x + tabW, y, w: w - tabW * 2, h };
    if (view === 'side') {
      const n = 18;
      for (let i = 0; i < n; i++) hline(slide, pres, body.x, y + (i + 0.5) * h / n, body.w, i % 2 ? 'B4BAC4' : '8C939E', 0.75);
    } else {
      const ratio = D.xMax / D.yMax;
      let bw = w / 1.1, bh = bw / ratio;
      if (bh > h) { bh = h; bw = bh * ratio; }
      tabW = bw * 0.05;
      body = { x: x + (w - bw) / 2, y: y + (h - bh) / 2, w: bw, h: bh };
      rect(slide, pres, { x: body.x - tabW, y: body.y + bh * 0.3, w: tabW, h: bh * 0.4, fill: { color: 'F2F2F2' }, line: { color: 'D0D0D0', width: 0.5 } });
      slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: body.x, y: body.y, w: body.w, h: body.h, rectRadius: 0.03, fill: { color: 'AEB4BE' }, line: { color: '8C939E', width: 0.5 } });
      rect(slide, pres, { x: body.x + body.w, y: body.y + bh * 0.3, w: tabW, h: bh * 0.4, fill: { color: 'F8CBAD' } });
    }
    const d = Math.max(0.05, Math.min(0.075, h / 11));
    records.forEach((r) => {
      if (r.x == null) return;
      const yy = view === 'side' ? r.layer : r.y;
      if (yy == null) return;
      const fx = Math.max(0, Math.min(1, r.x / D.xMax));
      const fy = view === 'side' ? Math.max(0, Math.min(1, (yy - 1) / Math.max(1, D.layers - 1))) : 1 - Math.max(0, Math.min(1, yy / D.yMax));
      const color = D.colorFn ? D.colorFn(r.element) : M.elementColor(r.element || 'Unknown');
      slide.addShape(pres.shapes.OVAL, { x: body.x + fx * body.w - d / 2, y: body.y + fy * body.h - d / 2, w: d, h: d, fill: { color: r.element ? color : 'FFFFFF' }, line: { color: '262626', width: 0.25 } });
    });
  }
  function sortedElements(records) {
    const order = Object.keys(M.ELEMENT_COLORS);
    return [...new Set(records.map((r) => r.element).filter(Boolean))].sort((a, b) => (order.indexOf(a) + 1 || 99) - (order.indexOf(b) + 1 || 99) || a.localeCompare(b));
  }

  const AXIS = { catAxisLabelFontFace: FONT, valAxisLabelFontFace: FONT, catAxisLabelColor: C.sub, valAxisLabelColor: C.sub, catAxisLineColor: 'BFBFBF', valAxisLineShow: false };

  // Pie with "Name n (p%)" labels outside the slices; legend drawn separately
  function pie(slide, pres, entries, box, colorFn, opts) {
    const total = entries.reduce((s, [, n]) => s + n, 0);
    if (!total) { noData(slide, null, box); return; }
    const labels = entries.map(([k, n]) => (n / total < 0.05 ? '' : `${k} ${n}\n(${(n / total * 100).toFixed(1)}%)`));
    slide.addChart(pres.charts.PIE, [{ name: 'n', labels, values: entries.map(([, n]) => n) }], Object.assign({
      x: box.x, y: box.y, w: box.w, h: box.h,
      chartColors: entries.map(([k], i) => colorFn(k, i)),
      showLegend: false, showLabel: true, showValue: false, showPercent: false, showLeaderLines: true,
      dataLabelFontSize: 8, dataLabelFontFace: FONT, dataLabelColor: '262626', dataLabelPosition: 'outEnd',
      layout: { x: 0.2, y: 0.14, w: 0.6, h: 0.72 },
      firstSliceAng: 0, dataBorder: { pt: 1, color: 'FFFFFF' },
    }, opts || {}));
  }
  function stackedBars(slide, pres, series, box, opts) {
    const data = series.filter((s) => s.values.some((v) => v));
    if (!data.length) { noData(slide, null, box); return; }
    slide.addChart(pres.charts.BAR, data, Object.assign({
      x: box.x, y: box.y, w: box.w, h: box.h, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 45,
      chartColors: data.map((s, i) => (opts && opts.colorFn ? opts.colorFn(s.name, i) : M.elementColor(s.name, i))),
      showLegend: false, catAxisLabelFontSize: 7, valAxisLabelFontSize: 7,
      valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' },
    }, AXIS, opts || {}));
  }
  function stackedWithAvg(slide, pres, block, box) {
    const data = block.series.filter((s) => s.values.some((v) => v));
    if (!data.length) { noData(slide, null, box); return; }
    const labels = data[0].labels;
    slide.addChart([
      { type: pres.charts.BAR, data, options: { barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 30, chartColors: data.map((s, i) => M.elementColor(s.name, i)) } },
      { type: pres.charts.LINE, data: [{ name: `Avg ${block.avg.toFixed(1)}`, labels, values: labels.map(() => Math.round(block.avg * 100) / 100) }], options: { chartColors: ['404040'], lineSize: 1, lineDash: 'dash', lineDataSymbol: 'none' } },
    ], Object.assign({
      x: box.x, y: box.y, w: box.w, h: box.h, showLegend: false,
      catAxisLabelFontSize: labels.length > 30 ? 6 : 7, catAxisLabelRotate: labels.length > 8 ? 270 : 0, valAxisLabelFontSize: 7,
      valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' },
    }, AXIS));
    txt(slide, `평균 ${block.avg.toFixed(1)}`, { x: box.x + box.w - 1.0, y: box.y - 0.01, w: 0.95, h: 0.16, fontSize: 7.5, color: C.sub, align: 'right' });
  }
  // Count table with a heat-map body; first row = header, first column = labels, last row/column = totals
  function simpleTable(slide, rows, box, opts) {
    const o = Object.assign({ fontSize: 7, heat: true }, opts);
    const max = Math.max(0, ...rows.slice(1, -1).flatMap((r) => r.slice(1, -1).map((c) => (typeof c === 'number' ? c : 0))));
    const cells = rows.map((r, ri) => r.map((c, ci) => {
      const isHead = ri === 0;
      const isTotal = ri === rows.length - 1 || ci === r.length - 1;
      const v = typeof c === 'number' ? c : null;
      const hot = !isHead && ci > 0 && !isTotal && o.heat;
      return {
        text: v === null ? String(c) : (v ? String(v) : '-'),
        options: {
          fontFace: FONT, fontSize: isHead && r.length > 14 ? Math.min(o.fontSize, 6) : o.fontSize, align: ci === 0 ? 'left' : 'center', valign: 'middle',
          bold: isHead || isTotal, color: hot ? heatText(v, max) : isHead ? C.navy : '262626',
          fill: { color: isHead ? C.panel : hot ? heat(v, max) : 'FFFFFF' },
          border: [{ type: 'none' }, { type: 'none' }, { pt: 0.5, color: C.line }, { type: 'none' }],
          margin: r.length > 14 ? 0 : [0, 0.02, 0, 0.02],
        },
      };
    }));
    slide.addTable(cells, { x: box.x, y: box.y, w: box.w, colW: box.colW, rowH: box.rowH || 0.18 });
  }

  // ---------------------------------------------------------------------------
  // Slide 1: trend
  // ---------------------------------------------------------------------------
  function rowLabel(slide, pres, label, y, h, items) {
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: MX, y, w: 0.58, h, rectRadius: 0.05, fill: { color: C.panel }, line: { color: C.line, width: 0.75 } });
    const labelH = items ? Math.min(h - items.length * 0.16 - 0.06, 0.55) : h;
    txt(slide, label, { x: MX, y: y + (items ? 0.04 : 0), w: 0.58, h: labelH, fontSize: 8.5, bold: true, align: 'center', color: C.navy });
    (items || []).forEach((it, i) => {
      const iy = y + h - (items.length - i) * 0.16 - 0.05;
      rect(slide, pres, { x: MX + 0.1, y: iy + 0.045, w: 0.1, h: 0.07, fill: { color: it.color } });
      txt(slide, it.label, { x: MX + 0.24, y: iy, w: 0.32, h: 0.16, fontSize: 7.5, color: C.sub });
    });
  }
  function trendSlide(pres, model) {
    const slide = pres.addSlide();
    const t = model.trend;
    chrome(slide, pres, '저전압 발생 및 개선 현황', model.meta);
    txt(slide, model.trendHeadline, { x: MX, y: 0.76, w: R - MX, h: 0.46, fontSize: 12.5, bold: true, valign: 'top' });

    const X0 = 1.05, CW = R - X0, n = t.lots.length || 1;
    const cx = (i) => X0 + (i + 0.5) / n * CW;
    const edge = (i) => X0 + i / n * CW;
    const layout = { x: 0, y: 0.06, w: 1, h: 0.88 };
    const Y = { mon: 1.26, rate: 1.54, rateH: 1.24, io: 2.84, ioH: 0.44, note: 3.34, prodEnd: 5.32, table: 5.42 };

    // Month headers and boundaries
    t.months.forEach((mo, i) => {
      const x1 = edge(mo.first), x2 = edge(mo.last + 1);
      rect(slide, pres, { x: x1, y: Y.mon, w: x2 - x1, h: 0.24, fill: { color: C.navy }, line: { color: 'FFFFFF', width: 1 } });
      txt(slide, mo.label || mo.name, { x: x1, y: Y.mon, w: x2 - x1, h: 0.24, fontSize: 10.5, bold: true, align: 'center', color: 'FFFFFF' });
      if (i > 0) vline(slide, pres, x1, Y.rate, 3.5, '8C939E', 0.75, 'dash');
    });

    rowLabel(slide, pres, '저전압\n불량률', Y.rate, Y.rateH, [{ label: 'E', color: C.e }, { label: 'L', color: C.l }]);
    rowLabel(slide, pres, '외부 /\n내부', Y.io, Y.ioH);
    rowLabel(slide, pres, '개선\n사항\n및\n생산량', Y.note, Y.prodEnd - Y.note);

    // Rate lines (E, L)
    const rMax = Math.max(0.5, Math.ceil(Math.max(...t.eRate, ...t.lRate, 0) * 10) / 10);
    slide.addChart(pres.charts.LINE, [
      { name: 'E', labels: t.labels, values: t.eRate.map((v) => Math.round(v * 1000) / 1000) },
      { name: 'L', labels: t.labels, values: t.lRate.map((v) => Math.round(v * 1000) / 1000) },
    ], {
      x: X0, y: Y.rate, w: CW, h: Y.rateH, layout, chartColors: [C.e, C.l], lineSize: 1.5, lineDataSymbol: 'circle', lineDataSymbolSize: 3,
      catAxisHidden: true, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: rMax, valGridLine: { color: C.grid, size: 0.5, style: 'dash' },
      showLegend: false,
    });
    const yOf = (v) => Y.rate + Y.rateH * (layout.y + layout.h * (1 - v / rMax));
    for (let k = 0; k <= 5; k++) {
      const v = rMax * k / 5;
      txt(slide, (Math.round(v * 10) / 10).toFixed(1), { x: 0.76, y: yOf(v) - 0.07, w: 0.26, h: 0.14, fontSize: 7, align: 'right', color: C.sub });
    }
    txt(slide, '(%)', { x: 0.76, y: Y.rate - 0.12, w: 0.26, h: 0.12, fontSize: 6.5, align: 'right', color: C.mute });
    // E min / max labels per month
    t.months.forEach((mo) => {
      const idx = []; for (let i = mo.first; i <= mo.last; i++) idx.push(i);
      if (!idx.length) return;
      const hi = idx.reduce((a, b) => (t.eRate[b] > t.eRate[a] ? b : a));
      const lo = idx.reduce((a, b) => (t.eRate[b] < t.eRate[a] ? b : a));
      [[hi, -0.19], [lo, 0.05]].forEach(([i, dy]) => {
        txt(slide, t.eRate[i].toFixed(2), { x: cx(i) - 0.25, y: yOf(t.eRate[i]) + dy, w: 0.5, h: 0.14, fontSize: 8, bold: true, align: 'center', color: C.e });
      });
    });

    // Inside / outside share (100% stacked)
    const has = t.insideShare.map((v) => v != null);
    slide.addChart(pres.charts.BAR, [
      { name: '내부', labels: t.labels, values: t.insideShare.map((v) => (v == null ? 0 : Math.round(v * 10) / 10)) },
      { name: '외부', labels: t.labels, values: t.insideShare.map((v, i) => (has[i] ? Math.round((100 - v) * 10) / 10 : 0)) },
    ], {
      x: X0, y: Y.io, w: CW, h: Y.ioH, layout, barDir: 'col', barGrouping: 'stacked', barGapWidthPct: 25, chartColors: [C.inside, C.outside],
      catAxisHidden: true, valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: 100, valGridLine: { style: 'none' }, showLegend: false,
    });
    ['100%', '50%', '0%'].forEach((l, k) => txt(slide, l, { x: 0.74, y: Y.io + Y.ioH * (layout.y + layout.h * k / 2) - 0.07, w: 0.28, h: 0.14, fontSize: 6.5, align: 'right', color: C.sub }));
    // Monthly inside share written over each month's bars
    t.months.forEach((mo) => {
      if (!mo.located) return;
      const x1 = edge(mo.first), x2 = edge(mo.last + 1);
      txt(slide, `내부 ${(mo.inside / mo.located * 100).toFixed(1)}%`, { x: x1 + (x2 - x1) / 2 - 0.45, y: Y.io + 0.02, w: 0.9, h: 0.16, fontSize: 7.5, bold: true, align: 'center', color: C.ink, fill: { color: 'FFFFFF', transparency: 15 } });
    });

    // Production bars in the lower part of the band, improvement notes above them
    const pMax = Math.max(10000, Math.ceil(Math.max(...t.production, 0) / 5000) * 5000);
    const prodTop = 4.44, prodBottom = 5.06;
    const pLayout = { x: 0, y: (prodTop - 4.3) / (Y.prodEnd - 4.3), w: 1, h: (prodBottom - prodTop) / (Y.prodEnd - 4.3) };
    slide.addChart(pres.charts.BAR, [{ name: '생산량', labels: t.labels.map((l) => l.slice(1)), values: t.production }], Object.assign({
      x: X0, y: 4.3, w: CW, h: Y.prodEnd - 4.3, layout: pLayout, barDir: 'col', barGapWidthPct: 25, chartColors: [C.bar],
      valAxisHidden: true, valAxisMinVal: 0, valAxisMaxVal: pMax, valGridLine: { style: 'none' }, showLegend: false,
      catAxisLabelFontSize: 6.5, catAxisLabelRotate: 270, catAxisLabelFrequency: n > 60 ? 3 : n > 30 ? 2 : 1, catAxisLineShow: false,
    }, AXIS, { catAxisLineShow: false }));
    const pY = (v) => prodTop + (prodBottom - prodTop) * (1 - v / pMax);
    [pMax, 0].forEach((v) => txt(slide, v ? `${Math.round(v / 1000)}K` : '0', { x: 0.74, y: pY(v) - 0.07, w: 0.28, h: 0.14, fontSize: 6.5, align: 'right', color: C.sub }));
    // Notes: flag + text, packed into up to three rows so texts never overlap
    const NOTE_W = 1.62, LINE_H = 0.12, ROW_H = 0.37, ROWS = 3;
    const rowsEnd = [];
    (model.notes || [])
      .map((note) => ({ note, i: t.lots.indexOf(M.padLot(note.lot)) }))
      .filter((it) => it.i >= 0)
      .sort((a, b) => a.i - b.i)
      .forEach(({ note, i }) => {
        const lines = String(note.text).split('\n');
        const x = Math.min(cx(i) + 0.1, R - NOTE_W);
        let row = rowsEnd.findIndex((end) => end < x);
        if (row < 0 && rowsEnd.length < ROWS) { row = rowsEnd.length; rowsEnd.push(0); }
        if (row < 0) row = rowsEnd.indexOf(Math.min(...rowsEnd));
        rowsEnd[row] = x + NOTE_W;
        const flagY = Y.note + 0.02 + row * ROW_H;
        vline(slide, pres, cx(i), flagY, prodBottom - flagY, '404040', 0.75);
        slide.addShape(pres.shapes.ISOSCELES_TRIANGLE, { x: cx(i) - 0.015, y: flagY, w: 0.1, h: 0.09, rotate: 90, fill: { color: C.accent }, line: { color: C.accent, width: 0.5 } });
        txt(slide, note.text, { x, y: flagY + 0.01, w: NOTE_W, h: LINE_H * lines.length + 0.02, fontSize: 7, bold: true, valign: 'top', color: C.ink });
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
    const firstW = 1.15;
    const colW = [firstW, ...Array(cols).fill((R - MX - firstW) / Math.max(cols, 1))];
    slide.addTable(rows.map((r, ri) => r.map((c, ci) => ({
      text: c,
      options: {
        fontFace: FONT, fontSize: 8.5, align: 'center', valign: 'middle', bold: ri === 0 || ci === 0 || ri === 3,
        color: ri === 0 ? 'FFFFFF' : ci === 0 ? C.navy : C.ink,
        fill: { color: ri === 0 ? C.navy : ci === 0 ? C.panel : ri === 3 ? 'F7F9FC' : 'FFFFFF' },
        border: { type: 'solid', pt: 0.5, color: C.line }, margin: 0.02,
      },
    }))), { x: MX, y: Y.table, w: R - MX, colW, rowH: [0.26, 0.25, 0.25, 0.27, 0.25, 0.25, 0.26] });
  }

  // ---------------------------------------------------------------------------
  // Slide 2 / 8: "5대 분석 검사 결과"
  // ---------------------------------------------------------------------------
  function funnel(slide, pres, s, header, x, y) {
    txt(slide, header, { x, y, w: 3.05, h: 0.44, fontSize: 7.5, color: C.green, valign: 'top' });
    const top = y + 0.46, H = 1.3;
    const box = (bx, by, bw, bh, label, fill, size, color) => {
      slide.addShape(pres.shapes.RECTANGLE, { x: bx, y: by, w: bw, h: bh, fill: { color: fill }, line: { color: '7F8792', width: 0.75 } });
      txt(slide, label, { x: bx, y: by, w: bw, h: bh, fontSize: size || 9, align: 'center', color: color || C.ink });
    };
    box(x, top, 0.6, H, `E등급\n\n${s.total} 셀\n(100%)`, 'FFFFFF', 9.5);
    const dropH = s.total ? Math.min(H - 0.45, Math.max(0.62, H * s.dropAll / s.total)) : H / 2;
    box(x + 0.64, top, 0.66, dropH, `냉동저항\nNG\n${s.dropAll}셀\n(${M.pct(s.dropAll, s.total)}%)`, 'DCE1E8', 8.5);
    box(x + 0.64, top + dropH, 0.66, H - dropH, `NTF\n${s.ntf}셀\n(${M.pct(s.ntf, s.total)}%)`, 'EEF0F3', 8.5);
    box(x + 1.34, top, 0.66, dropH, `진성불량\n${s.genuine}셀\n(${M.pct(s.genuine, s.dropAll)}%)`, '8C96A5', 8.5, 'FFFFFF');
    if (s.nrcf) txt(slide, `NRCF ${s.nrcf}셀 (${M.pct(s.nrcf, s.dropAll)}%)`, { x: x + 1.3, y: top + dropH + 0.03, w: 0.74, h: 0.16, fontSize: 7, align: 'center', color: C.sub });
    // Location breakdown of genuine cells (each band tall enough for two lines)
    const parts = [
      ['코팅층 외부', s.loc['Coating Top'], 'DEEBF7'],
      ['코팅층 내부', s.loc['Coating Inside'], 'E2F0D9'],
      ['Al Foil Surface', s.loc['Al Foil Surface'], 'F8CBAD'],
      ['위치 미확인', s.locUnknown, 'F2F2F2'],
    ].filter((p) => p[1]);
    const minH = 0.3;
    const colH = Math.min(H, Math.max(dropH, minH * parts.length));
    const free = Math.max(0, colH - minH * parts.length);
    let cy = top;
    parts.forEach(([label, n, fill]) => {
      const h = minH + (s.genuine ? free * n / s.genuine : 0);
      box(x + 2.04, cy, 0.98, h, `${label}\n${n}셀 (${M.pct(n, s.genuine)}%)`, fill, 7.5);
      cy += h;
    });
    txt(slide, `총 ${s.genuine}셀`, { x: x + 2.04, y: cy + 0.02, w: 0.98, h: 0.2, fontSize: 9, bold: true, align: 'center' });
  }
  function analysisSlide(pres, model, block, titleText, suffix) {
    const slide = pres.addSlide();
    chrome(slide, pres, titleText, model.meta, suffix);
    const s = block.summary;
    txt(slide, block.headline, { x: MX, y: 0.75, w: R - MX, h: 0.52, fontSize: 12.5, valign: 'top' });
    funnel(slide, pres, s, block.funnelHeader, MX, 1.3);

    const gen = s.genuineRecords;
    const els = sortedElements(gen);
    const r1 = 1.32, r2 = 3.32, r3 = 5.42;
    panelTitle(slide, pres, '[이물 위치]', 3.33, r1, 3.55);
    electrodeMap(slide, pres, gen, { x: 3.38, y: r1 + 0.38, w: 2.5, h: 1.3 }, 'plane');
    elementLegend(slide, els, 5.95, r1 + 0.36, { w: 0.85, maxH: r2 - r1 - 0.42 });
    panelTitle(slide, pres, '[이물 측면 위치]', 6.98, r1, R - 6.98);
    electrodeMap(slide, pres, gen, { x: 7.03, y: r1 + 0.38, w: 2.6, h: 1.3 }, 'side');
    elementLegend(slide, els, 9.72, r1 + 0.36, { w: 0.85, maxH: r2 - r1 - 0.42 });

    panelTitle(slide, pres, `[${s.genuine} 셀 분석]`, MX, r2, 3.4);
    pie(slide, pres, s.elements, { x: MX, y: r2 + 0.3, w: 2.5, h: 1.8 }, (k, i) => M.elementColor(k, i));
    elementLegend(slide, s.elements.map(([k]) => k), 2.72, r2 + 0.36, { w: 0.85, maxH: r3 - r2 - 0.42 });

    const locs = M.LOCATIONS.filter((l) => s.loc[l]);
    const locTitle = locs.map((l) => ({ 'Al Foil Surface': `Al Foil ${s.loc[l]}셀`, 'Coating Inside': `코팅층 내부 ${s.loc[l]}셀`, 'Coating Top': `외부 ${s.loc[l]}셀` }[l])).join(', ');
    panelTitle(slide, pres, `[${locTitle || '위치 정보 없음'} 분석]`, 3.68, r2, 3.4);
    // Up to 2 locations: one pie per row; 3: two small side by side on top, one below
    const slots = locs.length <= 2
      ? [{ x: 3.68, y: r2 + 0.3, w: 2.5, h: 1.85 }, { x: 3.68, y: r3 + 0.02, w: 2.5, h: 1.9 }]
      : [{ x: 3.68, y: r2 + 0.3, w: 1.3, h: 1.85 }, { x: 4.95, y: r2 + 0.3, w: 1.3, h: 1.85 }, { x: 3.68, y: r3 + 0.02, w: 2.5, h: 1.9 }];
    locs.forEach((l, i) => {
      const b = slots[i];
      txt(slide, l, { x: b.x, y: b.y, w: b.w, h: 0.16, fontSize: 7.5, bold: true, align: 'center', color: C.sub });
      pie(slide, pres, s.elementsByLocation[l], { x: b.x, y: b.y + 0.14, w: b.w, h: b.h - 0.14 }, (k, j) => M.elementColor(k, j), locs.length > 2 && i < 2 ? { dataLabelFontSize: 7, layout: { x: 0.18, y: 0.2, w: 0.64, h: 0.6 } } : {});
    });
    if (!locs.length) noData(slide, 'Analysis 데이터에 Location 값이 없습니다.', { x: 3.68, y: 4.5, w: 3.4, h: 0.3 });
    elementLegend(slide, els, 6.25, r2 + 0.36, { w: 0.8, maxH: 7.3 - r2 - 0.36 });

    panelTitle(slide, pres, '[CT 결과]', 7.18, r2, R - 7.18);
    pie(slide, pres, locs.map((l) => [l, s.loc[l]]), { x: 7.18, y: r2 + 0.3, w: 2.45, h: 1.8 }, (k) => M.LOCATION_COLORS[k] || 'A6A6A6');
    legend(slide, locs.map((l) => ({ label: l, color: M.LOCATION_COLORS[l] })), 9.62, r2 + 0.36, { w: 0.95 });

    panelTitle(slide, pres, '[LOT 별 주요 이물]', MX, r3, 3.4);
    const panels = block.lotPanels;
    const bandH = [0.5, 0.5, 0.66];
    let by = r3 + 0.32;
    M.GROUPS.forEach((g, i) => {
      txt(slide, g, { x: MX, y: by, w: 0.48, h: 0.4, fontSize: 7.5, bold: true, color: C.sub });
      stackedBars(slide, pres, panels[g], { x: MX + 0.45, y: by, w: 2.15, h: bandH[i] }, { catAxisHidden: i < 2, catAxisLabelFontSize: 6.5, catAxisLabelRotate: block.lots.length > 8 ? 270 : 0, valAxisLabelFontSize: 6.5 });
      by += bandH[i];
    });
    elementLegend(slide, els, 2.8, r3 + 0.36, { size: 7, step: 0.15, w: 0.75, maxH: 7.3 - r3 - 0.36 });

    panelTitle(slide, pres, '[이물 위치 Top/Back 개수]', 7.18, r3, R - 7.18);
    pie(slide, pres, s.topBack, { x: 7.18, y: r3 + 0.3, w: 2.45, h: 1.72 }, (k) => M.TOPBACK_COLORS[k] || 'A6A6A6');
    legend(slide, s.topBack.map(([k]) => ({ label: k, color: M.TOPBACK_COLORS[k] || 'A6A6A6' })), 9.62, r3 + 0.36, { w: 0.9 });
    // Thin separators between the three columns of the lower part
    vline(slide, pres, 3.6, r2, 7.35 - r2, C.line, 0.5);
    vline(slide, pres, 7.1, r2, 7.35 - r2, C.line, 0.5);
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
    const PW = (R - MX - 0.2) / 2;
    const grid = (title, x, view) => {
      panelTitle(slide, pres, title, x, 0.85, PW);
      ['Back', 'Top'].forEach((tb, c) => txt(slide, tb, { x: x + 0.55 + c * 1.84, y: 1.16, w: 1.78, h: 0.16, fontSize: 8, bold: true, align: 'center', color: C.sub }));
      M.GROUPS.forEach((g, r) => {
        const ry = 1.34 + r * 0.76;
        txt(slide, g, { x: x + 0.02, y: ry, w: 0.5, h: 0.64, fontSize: 8, bold: true, color: C.sub });
        ['Back', 'Top'].forEach((tb, c) => {
          electrodeMap(slide, pres, gen.filter((x2) => x2.group === g && x2.topBack === tb), { x: x + 0.55 + c * 1.84, y: ry, w: 1.78, h: 0.64 }, view);
        });
        if (r) hline(slide, pres, x, ry - 0.06, PW - 0.9, C.grid, 0.5);
      });
      elementLegend(slide, els, x + PW - 0.88, 1.2, { w: 0.78, maxH: 2.4 });
    };
    grid('[이물 위치 Top/Back View]', MX, 'plane');
    grid('[이물 위치 Top/Back Side View]', MX + PW + 0.2, 'side');
    const inout = (title, x, view) => {
      panelTitle(slide, pres, title, x, 3.75, PW);
      M.LOCATIONS.forEach((l, r) => {
        const ry = 4.12 + r * 1.06;
        txt(slide, l.replace(/ (?=\S+$)/, '\n'), { x: x + 0.02, y: ry, w: 0.75, h: 0.92, fontSize: 8, bold: true, color: C.sub });
        electrodeMap(slide, pres, gen.filter((x2) => x2.location === l), { x: x + 0.8, y: ry, w: 3.2, h: 0.92 }, view);
      });
      elementLegend(slide, els, x + PW - 0.88, 4.1, { w: 0.78, maxH: 3.15 });
    };
    inout('[내/외부 이물 위치]', MX, 'plane');
    inout('[내/외부 이물 측면 위치]', MX + PW + 0.2, 'side');
  }

  // Slide 4: small multiples per element, Top / Back
  function elementSlide(pres, model) {
    const slide = pres.addSlide();
    chrome(slide, pres, `3. E81C 저전압 (${model.label}) 원인 파악 – 이물 발견 위치`, model.meta);
    const gen = model.report.summary.genuineRecords;
    const PW = (R - MX - 0.2) / 2;
    const block = (title, x, y, tb, view) => {
      panelTitle(slide, pres, title, x, y, PW);
      const recs = gen.filter((r) => r.topBack === tb);
      const els = sortedElements(recs);
      els.slice(0, 8).forEach((el, i) => {
        const bx = x + 0.05 + (i % 2) * 2.05, by = y + 0.36 + Math.floor(i / 2) * 0.7;
        txt(slide, el, { x: bx, y: by, w: 1.95, h: 0.15, fontSize: 8, bold: true, align: 'center', color: C.sub });
        electrodeMap(slide, pres, recs.filter((r) => r.element === el), { x: bx, y: by + 0.16, w: 1.95, h: 0.5 }, view);
      });
      if (!els.length) noData(slide, '(해당 셀 없음)', { x, y: y + 1.2, w: PW - 0.9, h: 0.3 });
      elementLegend(slide, els, x + PW - 0.88, y + 0.4, { w: 0.78, maxH: 2.8 });
    };
    block('[이물 위치 Top View]', MX, 0.85, 'Top', 'plane');
    block('[이물 측면 위치 Top View]', MX + PW + 0.2, 0.85, 'Top', 'side');
    block('[이물 위치 Back View]', MX, 4.1, 'Back', 'plane');
    block('[이물 측면 위치 Back View]', MX + PW + 0.2, 4.1, 'Back', 'side');
  }

  // ---------------------------------------------------------------------------
  // Slide 5: coating-top (outside) analysis — AZS stacker / DNC
  // ---------------------------------------------------------------------------
  function outsideSlide(pres, model) {
    const slide = pres.addSlide();
    chrome(slide, pres, `4. E81C 저전압 (${model.label}) 원인 파악 – 코팅층 외부 분석`, model.meta);
    const st = model.stacker, oe = model.outside;
    const LW = 4.3, MXc = MX + LW + 0.12, MW = 3.0, RXc = MXc + MW + 0.12, RW = R - RXc;
    if (st) {
      const tableRows = (list, firstLabel) => {
        const totals = st.positions.map((_, i) => list.reduce((s, r) => s + r.values[i], 0));
        return [
          [firstLabel, ...st.positions, 'Total'],
          ...list.map((r) => [r.label, ...r.values, r.values.reduce((a, b) => a + b, 0)]),
          ['Total', ...totals, totals.reduce((a, b) => a + b, 0)],
        ];
      };
      const colW = [0.45, ...st.positions.map(() => ((LW - 0.85) / st.positions.length)), 0.4];
      panelTitle(slide, pres, '[AZS] 월별 E 불량 개수', MX, 0.85, LW);
      simpleTable(slide, tableRows(st.byMonth, 'Month'), { x: MX, y: 1.16, w: LW, colW, rowH: 0.18 });
      const weekY = 1.3 + (st.byMonth.length + 2) * 0.18;
      panelTitle(slide, pres, '[AZS] 주별 E 불량 개수', MX, weekY, LW);
      simpleTable(slide, tableRows(st.byWeek, 'Week'), { x: MX, y: weekY + 0.31, w: LW, colW, rowH: 0.18 });
      // Stacker trend: one sparkline per stacker (weekly counts)
      const trendY = Math.max(4.0, weekY + 0.45 + (st.byWeek.length + 2) * 0.18);
      panelTitle(slide, pres, '[AZS] 스태커 트랜드 (주별)', MX, trendY, LW);
      const labels = st.byWeek.map((r) => r.label);
      const shown = st.positions.slice(0, 18);
      const rowH = Math.min(0.2, (7.35 - trendY - 0.32) / Math.max(shown.length, 1));
      shown.forEach((p, i) => {
        const y = trendY + 0.32 + i * rowH;
        txt(slide, p, { x: MX + 0.05, y, w: 0.4, h: rowH, fontSize: 7, bold: true, color: C.sub });
        const vals = st.byWeek.map((r) => r.values[i]);
        slide.addChart(pres.charts.LINE, [{ name: p, labels, values: vals }], {
          x: MX + 0.5, y, w: LW - 1.0, h: rowH, chartColors: [C.accent], lineSize: 1, lineDataSymbol: 'none', catAxisHidden: true, valAxisHidden: true,
          valAxisMinVal: 0, valGridLine: { style: 'none' }, showLegend: false, layout: { x: 0, y: 0.1, w: 1, h: 0.8 },
        });
        txt(slide, String(vals[vals.length - 1] || 0), { x: MX + LW - 0.45, y, w: 0.4, h: rowH, fontSize: 7, bold: true, align: 'right', color: C.ink });
        hline(slide, pres, MX, y + rowH, LW, C.grid, 0.5);
      });
    } else {
      noData(slide, '공정 이력(CSV) 데이터가 없어 AZS 분석을 만들 수 없습니다.', { x: MX, y: 2, w: LW, h: 0.4 });
    }
    if (oe && oe.count) {
      panelTitle(slide, pres, '[AZS] 스태커별 셀 수량', MXc, 0.85, MW);
      stackedBars(slide, pres, oe.stackerByElement, { x: MXc, y: 1.18, w: MW - 0.82, h: 2.6 }, { catAxisLabelFontSize: 6.5, catAxisLabelFrequency: 1, catAxisLabelRotate: oe.positions.length > 8 ? 270 : 0 });
      elementLegend(slide, oe.elements, MXc + MW - 0.8, 1.22, { w: 0.72, maxH: 2.6 });
      panelTitle(slide, pres, '[DNC] DNC 장비 구분', RXc, 0.85, RW);
      stackedBars(slide, pres, oe.dncByElement, { x: RXc, y: 1.18, w: RW - 0.82, h: 2.6 }, { catAxisLabelFontSize: 7, catAxisLabelRotate: 270 });
      elementLegend(slide, oe.elements, RXc + RW - 0.8, 1.22, { w: 0.72, maxH: 2.6 });
      const tbl = (t, x, w, first) => {
        const totals = t.cols.map((_, i) => t.rows.reduce((s, r) => s + r.values[i], 0));
        const rows = [[first, ...t.cols, 'Total'], ...t.rows.map((r) => [r.label, ...r.values, r.values.reduce((a, b) => a + b, 0)]), ['Total', ...totals, totals.reduce((a, b) => a + b, 0)]];
        simpleTable(slide, rows, { x, y: 4.31, w, colW: [0.42, ...t.cols.map(() => (w - 0.82) / t.cols.length), 0.4], rowH: 0.2 });
      };
      const BW = MW + 0.12 + RW * 0.45;
      panelTitle(slide, pres, '[AZS] PKG LOT × 스태커 셀 수량', MXc, 4.0, BW);
      tbl(oe.pkgByStacker, MXc, BW, 'PKG');
      panelTitle(slide, pres, '[AZS] 라인 구분', MXc + BW + 0.12, 4.0, R - (MXc + BW + 0.12));
      tbl(oe.pkgByLine, MXc + BW + 0.12, R - (MXc + BW + 0.12), 'PKG');
    } else {
      noData(slide, '코팅층 외부(Coating Top) 셀의 공정 이력이 없습니다.', { x: MXc, y: 2, w: R - MXc, h: 0.4 });
    }
  }

  // ---------------------------------------------------------------------------
  // Slides 6/7: electrode (cathode / anode)
  // ---------------------------------------------------------------------------
  const BATCH_COLORS = ['9C7F1F', '98239B', '5B7FF5', 'F47C62', 'B679F2', '8FB4F9', 'FFD966', '7FD6A4', 'F9A0F0', '5B9BD5', 'ED7D31', '70AD47', 'FFC000', '4472C4', '264478', '9E480E', '636363', '997300'];
  function electrodeSlide(pres, model, e, title) {
    const slide = pres.addSlide();
    chrome(slide, pres, title, model.meta);
    const LW = 5.15;
    const share = (series, box) => {
      const data = series.filter((s) => s.values.some((v) => v));
      if (!data.length) { noData(slide, null, box); return; }
      slide.addChart(pres.charts.BAR, data.map((s) => ({ name: s.name, labels: s.labels, values: s.values.map((v) => Math.round(v * 10) / 10) })), Object.assign({
        x: box.x, y: box.y, w: box.w, h: box.h, barDir: 'col', barGrouping: 'percentStacked', barGapWidthPct: 40,
        chartColors: data.map((_, i) => BATCH_COLORS[i % BATCH_COLORS.length]), showLegend: true, legendPos: 'r', legendFontSize: data.length > 14 ? 6.5 : 7, legendFontFace: FONT, legendColor: C.sub,
        catAxisLabelFontSize: 7, valAxisLabelFontSize: 7, valAxisLabelFormatCode: '0%', valGridLine: { color: C.grid, size: 0.5 }, catGridLine: { style: 'none' },
      }, AXIS));
    };
    panelTitle(slide, pres, `[전극] 날짜별 ${e.name} 코팅 배치 vs PKG LOT`, MX, 0.85, LW);
    share(e.data.batchVsPkg, { x: MX, y: 1.16, w: LW, h: 2.75 });
    panelTitle(slide, pres, `[전극] ${e.name} 코팅 배치 LOT vs PKG LOT`, MX, 4.02, LW);
    share(e.data.batchLotVsPkg, { x: MX, y: 4.33, w: LW, h: 3.0 });
    const els = sortedElements(model.report.summary.genuineRecords);
    const RX = MX + LW + 0.15, RWd = R - RX;
    [['PKG LOT', e.data.pkgVsForeign], ['슬리팅 LOT', e.data.slitVsForeign], ['롤프레싱 LOT', e.data.rollVsForeign], ['코팅 LOT', e.data.coatVsForeign]].forEach(([label, block], i) => {
      const y = 0.85 + i * 1.63;
      panelTitle(slide, pres, `[전극] ${e.name} ${label} vs 이물`, RX, y, RWd);
      stackedWithAvg(slide, pres, block, { x: RX, y: y + 0.3, w: RWd - 0.8, h: 1.3 });
      elementLegend(slide, els, R - 0.78, y + 0.34, { size: 7, step: 0.15, w: 0.72, maxH: 1.25 });
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
    const cy = 0.85, ch = 3.5;
    slide.addChart([
      { type: pres.charts.BAR, data: [{ name: '생산량', labels, values: h.prod }], options: { barDir: 'col', chartColors: ['C9CED8'], barGapWidthPct: 60, secondaryValAxis: true, secondaryCatAxis: true } },
      { type: pres.charts.LINE, data: [{ name: '불량률 E', labels, values: h.e.map((v) => (v == null ? 0 : v)) }, { name: '불량률 L', labels, values: h.l.map((v) => (v == null ? 0 : v)) }], options: { chartColors: [C.e, C.l], lineSize: 2, lineDataSymbol: 'none' } },
    ], Object.assign({
      x: MX, y: cy, w: R - MX, h: ch, showLegend: true, legendPos: 't', legendFontSize: 9, legendFontFace: FONT, legendColor: C.sub,
      showTitle: true, title: 'E81C 저전압 불량률 (By PKG Lot, Lot별 dOCV 판정)', titleFontSize: 12, titleFontFace: FONT, titleColor: C.navy,
      catAxisLabelFontSize: 7, catAxisLabelRotate: 270, valGridLine: { color: C.grid, size: 0.5 },
      valAxes: [{ showValAxisTitle: false, valAxisMinVal: 0, valAxisMaxVal: 13, valAxisLabelFormatCode: '0.0"%"', valAxisLabelFontSize: 8 }, { showValAxisTitle: false, valAxisMinVal: 0, valAxisMaxVal: 14000, valAxisLabelFontSize: 8, valGridLine: { style: 'none' } }],
      catAxes: [{ catAxisTitle: 'LOT' }, { catAxisHidden: true }],
    }, AXIS));
    const sy = (oldY) => cy + (oldY - 1.15) * ch / 4.15; // positions below were measured on the original 4.15" chart
    [['C4 샘플 (FA22 ~ FB19) 생산량 : 42.4k\nE 등급: 0.26%', 1.0], ['Stress Run 생산량 : 95k (~ FD15)\nE 등급 : 0.21%', 6.9], ['양산 초도\n- FD10 : E 등급 0.48%\n- FD11 : E 등급 0.33%\n- FD12 : E 등급 0.37%\n- FD13 : E 등급 0.11%', 8.9]].forEach(([t, x]) => {
      txt(slide, t, { x, y: sy(1.62), w: 1.75, h: 0.72, fontSize: 8, bold: true, valign: 'top', color: C.ink });
    });
    [['[조립] 스택 버퍼 비산 방지 차폐 적용 완료 (2/4)', 1.05, 3.0], ['[전극] Mixer 활물질 투입 호퍼 Suction 외 4건 완료 (2/23~25)', 3.64, 3.15], ['Mixer/Coater/조립 설비 7건 집중 개선 (2/27)', 5.16, 2.81], ['[조립] 노칭 후 세정 유닛 추가 완료외 1건 (3/31)', 7.03, 2.57]].forEach(([t, x, y]) => {
      txt(slide, t, { x, y: sy(y), w: 3.8, h: 0.2, fontSize: 8, bold: true, color: 'C00000' });
    });

    // Table split into two halves so every value stays readable
    const groupOf = [];
    h.groups.forEach(([g, cnt]) => { for (let i = 0; i < cnt; i++) groupOf.push(g); });
    const cellOpt = (ri, extra) => Object.assign({
      fontFace: FONT, fontSize: 7, align: 'center', valign: 'middle', color: ri < 2 ? C.navy : C.ink, bold: ri < 2,
      fill: { color: ri === 0 ? 'DCE3EF' : ri === 1 ? C.panel : 'FFFFFF' }, border: { type: 'solid', pt: 0.5, color: C.line }, margin: 0,
    }, extra || {});
    const half = (from, to, withTotal, y) => {
      const lots = h.lots.slice(from, to);
      const head1 = [{ text: 'LOT', options: cellOpt(0, { rowspan: 2 }) }];
      for (let i = from; i < to;) {
        let j = i; while (j < to && groupOf[j] === groupOf[i]) j++;
        head1.push({ text: groupOf[i], options: cellOpt(0, { colspan: j - i }) });
        i = j;
      }
      if (withTotal) head1.push({ text: 'Total', options: cellOpt(0, { rowspan: 2 }) });
      const head2 = lots.map((l) => ({ text: l, options: cellOpt(1) }));
      const body = [
        ['불량률 E', ...h.e.slice(from, to).map((v) => (v == null ? '-' : v.toFixed(2))), ...(withTotal ? [h.totalE.toFixed(2)] : [])],
        ['불량률 L', ...h.l.slice(from, to).map((v) => (v == null ? '-' : v.toFixed(2))), ...(withTotal ? [h.totalL.toFixed(2)] : [])],
        ['생산량', ...h.prod.slice(from, to).map((v) => v.toLocaleString('en-US')), ...(withTotal ? [h.totalProd.toLocaleString('en-US')] : [])],
      ].map((r) => r.map((c, ci) => ({ text: String(c), options: cellOpt(2, ci === 0 ? { bold: true, color: C.navy, fill: { color: C.panel } } : {}) })));
      const firstW = 0.6, totalW = withTotal ? 0.55 : 0;
      const n = lots.length;
      const colW = [firstW, ...Array(n).fill((R - MX - firstW - totalW) / n), ...(withTotal ? [totalW] : [])];
      slide.addTable([head1, head2, ...body], { x: MX, y, w: R - MX, colW, rowH: [0.2, 0.2, 0.22, 0.22, 0.22] });
    };
    half(0, 27, false, 4.5);
    half(27, h.lots.length, true, 5.66);
  }

  // ---------------------------------------------------------------------------
  // Font clean-up of the written file: Korean (Hangul) font for every run, chart and the theme,
  // so the deck shows Malgun Gothic instead of a fallback font in every viewer.
  // ---------------------------------------------------------------------------
  const FONT_TAGS = `<a:latin typeface="${FONT}"/><a:ea typeface="${FONT}"/><a:cs typeface="${FONT}"/>`;
  function fixXml(name, xml) {
    if (/^ppt\/theme\//.test(name)) {
      return xml.replace(/<a:latin typeface="[^"]*"[^>]*\/>/g, `<a:latin typeface="${FONT}"/>`)
        .replace(/<a:ea typeface="[^"]*"[^>]*\/>/g, `<a:ea typeface="${FONT}"/>`)
        .replace(/<a:font script="Hang" typeface="[^"]*"\/>/g, '')
        .replace(/(<a:ea typeface="[^"]*"\/><a:cs typeface="[^"]*"\/>)/g, `$1<a:font script="Hang" typeface="맑은 고딕"/>`);
    }
    let out = xml.replace(/charset="-122"/g, 'charset="-127"').replace(/lang="en-US"/g, 'lang="ko-KR" altLang="en-US"').replace(/altLang="en-US" altLang="en-US"/g, 'altLang="en-US"');
    if (/^ppt\/charts\//.test(name)) {
      out = out.replace(/<a:(ea|cs)\s+typeface="[^"]*"[^>]*\/>/g, '').replace(/<a:latin\s+typeface="[^"]*"[^>]*\/>/g, FONT_TAGS);
    }
    return out;
  }
  async function fixFonts(zip) {
    const names = Object.keys(zip.files).filter((n) => /^ppt\/(slides|charts|theme|slideMasters|slideLayouts)\/[^/]+\.xml$/.test(n));
    for (const n of names) zip.file(n, fixXml(n, await zip.file(n).async('string')));
    return zip;
  }
  // Writes the deck with the font clean-up applied. type: 'blob' (browser) or 'nodebuffer' (Node)
  async function write(pres, JSZip, type) {
    const buf = await pres.write({ outputType: 'arraybuffer' });
    const zip = await fixFonts(await JSZip.loadAsync(buf));
    return zip.generateAsync({ type: type || 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', compression: 'DEFLATE' });
  }

  // ---------------------------------------------------------------------------
  // Build the whole deck
  // ---------------------------------------------------------------------------
  function build(PptxGenJS, model) {
    const pres = new PptxGenJS();
    pres.defineLayout({ name: 'A4R', width: W, height: 7.5 });
    pres.layout = 'A4R';
    pres.theme = { headFontFace: FONT, bodyFontFace: FONT };
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

  const api = { build, write, fixFonts, DEFAULT_NOTES, HISTORY, W, FONT, helpers: { txt, rect, hline, vline, chrome, panelTitle, legend, electrodeMap } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.LvReportPpt = api;
})(typeof self !== 'undefined' ? self : this);

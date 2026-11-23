/*
 * Builds Excel workbooks from parsed gamma reports (see parser.js).
 * Needs SheetJS (global XLSX in the browser, or passed in from Node).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory;
  else root.GammaWorkbook = factory(root.XLSX, root.GammaParser);
})(typeof self !== 'undefined' ? self : this, function (XLSX, GammaParser) {
  'use strict';

  const DATE_FMT = 'dd/mm/yyyy hh:mm:ss';
  const DAY_FMT = 'dd/mm/yyyy';

  // A cell is a plain value (string/number/null) or a date object from the parser.
  function dateCell(dt) {
    return { t: 'n', v: GammaParser.excelSerial(dt), z: dt.hasTime ? DATE_FMT : DAY_FMT, w: undefined };
  }

  function makeSheet(rows) {
    const ws = {};
    let maxC = 0;
    const widths = [];
    rows.forEach((row, r) => {
      (row || []).forEach((v, c) => {
        if (v === null || v === undefined || v === '') return;
        let cell, len;
        if (typeof v === 'object' && v.year) { cell = dateCell(v); len = 19; }
        else if (typeof v === 'number') { cell = { t: 'n', v }; len = String(v).length; }
        else { cell = { t: 's', v: String(v) }; len = String(v).length; }
        ws[XLSX.utils.encode_cell({ r, c })] = cell;
        widths[c] = Math.max(widths[c] || 0, len);
        if (c > maxC) maxC = c;
      });
    });
    ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(rows.length - 1, 0), c: maxC } });
    ws['!cols'] = Array.from({ length: maxC + 1 }, (_, c) => ({ wch: Math.min(Math.max((widths[c] || 6) + 2, 8), 60) }));
    return ws;
  }

  // Excel sheet names: max 31 chars, no []:*?/\ , unique (case-insensitive).
  function addSheet(wb, ws, name) {
    let base = String(name).replace(/[\[\]:*?\/\\]/g, '_').replace(/^'+|'+$/g, '').trim() || 'Sheet';
    base = base.slice(0, 31);
    let n = base, k = 2;
    const taken = new Set(wb.SheetNames.map((s) => s.toLowerCase()));
    while (taken.has(n.toLowerCase())) {
      const suffix = ' (' + k++ + ')';
      n = base.slice(0, 31 - suffix.length) + suffix;
    }
    XLSX.utils.book_append_sheet(wb, ws, n);
    return n;
  }

  // ---------- single report ----------

  function metadataRows(report) {
    const rows = [['Field', 'As printed', 'Value', 'Unit', 'Uncertainty']];
    for (const m of report.metadata) {
      rows.push([m.field, m.raw, m.kind === 'number' || m.kind === 'date' ? m.value : null, m.unit, m.uncertainty]);
    }
    return rows;
  }

  function dataRows(report) {
    const rows = [];
    report.tables.forEach((t, i) => {
      if (i > 0) rows.push([], []);
      rows.push([t.title]);
      rows.push(t.headers);
      for (const r of t.rows) rows.push(r.cells);
    });
    return rows;
  }

  function buildSingle(report) {
    const wb = XLSX.utils.book_new();
    addSheet(wb, makeSheet(metadataRows(report)), 'metadata');
    addSheet(wb, makeSheet(dataRows(report)), 'data');
    return wb;
  }

  // ---------- combined workbook ----------

  function summarySheetName(title) {
    if (/energ/i.test(title)) return 'Peaks (all samples)';
    if (/activ|MDA/i.test(title)) return 'Activities (all samples)';
    return title + ' (all)';
  }

  // One wide row per sample; numeric/date fields become real numbers.
  function wideMetadataRows(reports) {
    const fields = [];
    const info = new Map();
    for (const rep of reports) for (const m of rep.metadata) {
      if (!info.has(m.field)) { info.set(m.field, { kinds: new Set(), units: new Set(), unc: false }); fields.push(m.field); }
      const f = info.get(m.field);
      if (m.kind !== 'empty') f.kinds.add(m.kind);
      if (m.unit) f.units.add(m.unit);
      if (m.uncertainty) f.unc = true;
    }
    const header = ['Sample', 'File'];
    const plan = [];
    for (const field of fields) {
      const f = info.get(field);
      const numeric = f.kinds.size === 1 && (f.kinds.has('number') || f.kinds.has('date')) && f.units.size <= 1;
      if (numeric) {
        const unit = [...f.units][0];
        header.push(unit ? field + ' (' + unit + ')' : field);
        plan.push({ field, mode: 'value' });
        if (f.unc) { header.push(field + ' uncertainty'); plan.push({ field, mode: 'unc' }); }
      } else {
        header.push(field);
        plan.push({ field, mode: 'raw' });
      }
    }
    const rows = [header];
    for (const rep of reports) {
      const byField = new Map(rep.metadata.map((m) => [m.field, m]));
      const row = [rep.sample, rep.fileName];
      for (const p of plan) {
        const m = byField.get(p.field);
        if (!m) row.push(null);
        else if (p.mode === 'value') row.push(m.value);
        else if (p.mode === 'unc') row.push(m.uncertainty);
        else row.push(m.raw);
      }
      rows.push(row);
    }
    return rows;
  }

  // For each table title: one long table with Sample + File columns prepended.
  function longTables(reports) {
    const groups = new Map();
    for (const rep of reports) for (const t of rep.tables) {
      if (!groups.has(t.title)) groups.set(t.title, { headers: [], rows: [] });
      const g = groups.get(t.title);
      for (const h of t.headers) if (!g.headers.includes(h)) g.headers.push(h);
      for (const r of t.rows) {
        const obj = {};
        t.headers.forEach((h, i) => { obj[h] = r.cells[i]; });
        g.rows.push({ sample: rep.sample, file: rep.fileName, obj });
      }
    }
    return [...groups.entries()].map(([title, g]) => ({
      title,
      rows: [['Sample', 'File', ...g.headers], ...g.rows.map((x) => [x.sample, x.file, ...g.headers.map((h) => (x.obj[h] === undefined ? null : x.obj[h]))])],
    }));
  }

  function buildCombined(reports) {
    const wb = XLSX.utils.book_new();
    const longs = longTables(reports);
    // Activities first: it is the table people compare across samples.
    longs.sort((a, b) => (/activ|MDA/i.test(b.title) ? 1 : 0) - (/activ|MDA/i.test(a.title) ? 1 : 0));
    for (const l of longs) {
      const ws = makeSheet(l.rows);
      ws['!autofilter'] = { ref: ws['!ref'] };
      addSheet(wb, ws, summarySheetName(l.title));
    }
    const meta = makeSheet(wideMetadataRows(reports));
    meta['!autofilter'] = { ref: meta['!ref'] };
    addSheet(wb, meta, 'Metadata (all samples)');
    for (const rep of reports) {
      addSheet(wb, makeSheet(dataRows(rep)), rep.sample);
    }
    return wb;
  }

  function toArrayBuffer(wb) {
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  }

  return { buildSingle, buildCombined, toArrayBuffer };
});

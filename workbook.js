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

  // ---------- sample info + one-row-per-sample activities ----------

  const isActivityTable = (t) => /activ|MDA/i.test(t.title) || (t.headers.some((h) => /^isotope$/i.test(h)) && !t.headers.some((h) => /^energy$/i.test(h)));
  const findMeta = (rep, re) => rep.metadata.find((m) => re.test(m.field));

  // [header, value] pairs repeated on every data row: date, detector, mass, mass uncertainty
  function sampleInfo(rep) {
    const out = [['Sample', rep.sample], ['File', rep.fileName]];
    const date = findMeta(rep, /^acquisition date$/i);
    out.push(['Acquisition date', date && date.kind === 'date' ? date.value : date ? date.raw : null]);
    const det = findMeta(rep, /^detector$/i);
    out.push(['Detector', det ? (det.kind === 'number' ? det.value : det.raw) : null]);
    const mass = findMeta(rep, /^mass\s*\/\s*volume$/i) || findMeta(rep, /^mass/i);
    const label = mass ? mass.field : 'Mass/Volume';
    if (mass && mass.kind === 'number') {
      out.push([label + (mass.unit ? ' (' + mass.unit + ')' : ''), mass.value]);
      const u = /^([-+]?[\d.]+(?:[eE][-+]?\d+)?)\s*(.*)$/.exec(mass.uncertainty || '');
      out.push([label + ' uncertainty' + (u && u[2] ? ' (' + u[2] + ')' : ''), u ? Number(u[1]) : null]);
    } else {
      out.push([label, mass ? mass.raw : null], [label + ' uncertainty', null]);
    }
    return out;
  }

  // Activity table flattened: "K-40 Activ (Bq/kg)", "K-40 Uncert. (%)", ...
  function wideActivities(rep) {
    const t = rep.tables.find(isActivityTable);
    if (!t) return [];
    const isoCol = t.headers.findIndex((h) => /^isotope$/i.test(h));
    if (isoCol < 0) return [];
    const out = [];
    for (const r of t.rows) {
      const iso = r.cells[isoCol];
      if (iso === null || iso === undefined) continue;
      t.headers.forEach((h, c) => {
        if (c === isoCol || /^no\.?$/i.test(h)) return;
        if (r.footer && (r.cells[c] === null || r.cells[c] === undefined)) return; // e.g. SUM has no CL/MDA
        out.push([iso + ' ' + h, r.cells[c]]);
      });
    }
    return out;
  }

  // Merge [header, value] lists of several samples into one table (union of columns, first-seen order).
  function unionRows(lists) {
    const headers = [];
    const seen = new Set();
    for (const l of lists) for (const [h] of l) if (!seen.has(h)) { seen.add(h); headers.push(h); }
    const rows = lists.map((l) => { const m = new Map(l); return headers.map((h) => (m.has(h) ? m.get(h) : null)); });
    return { headers, rows };
  }

  // Blocks shown on a data sheet: the one-row-per-sample activity block, then every other table as a list.
  function dataBlocks(reports) {
    const blocks = [];
    const act = reports.map((r) => r.tables.find(isActivityTable)).find(Boolean);
    const wide = unionRows(reports.map((r) => [...sampleInfo(r), ...wideActivities(r)]));
    blocks.push({ title: act ? act.title : 'Activities', headers: wide.headers, rows: wide.rows });
    const titles = [];
    for (const r of reports) for (const t of r.tables) if (!isActivityTable(t) && !titles.includes(t.title)) titles.push(t.title);
    for (const title of titles) {
      const lists = [];
      for (const r of reports) {
        const t = r.tables.find((x) => x.title === title);
        if (!t) continue;
        const info = sampleInfo(r);
        for (const row of t.rows) lists.push([...info, ...t.headers.map((h, i) => [h, row.cells[i]])]);
      }
      const u = unionRows(lists);
      blocks.push({ title, headers: u.headers, rows: u.rows });
    }
    return blocks;
  }

  function blockRows(blocks) {
    const rows = [];
    blocks.forEach((b, i) => {
      if (i > 0) rows.push([], []);
      rows.push([b.title], b.headers, ...b.rows);
    });
    return rows;
  }

  function buildSingle(report) {
    const wb = XLSX.utils.book_new();
    addSheet(wb, makeSheet(metadataRows(report)), 'metadata');
    addSheet(wb, makeSheet(blockRows(dataBlocks([report]))), 'data');
    return wb;
  }

  // ---------- combined workbook ----------

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

  function summarySheetName(title, i) {
    if (i === 0) return 'Data (all samples)';
    if (/energ/i.test(title)) return 'Peaks (all samples)';
    return title + ' (all)';
  }

  function buildCombined(reports) {
    const wb = XLSX.utils.book_new();
    dataBlocks(reports).forEach((b, i) => {
      const ws = makeSheet([b.headers, ...b.rows]);
      ws['!autofilter'] = { ref: ws['!ref'] };
      addSheet(wb, ws, summarySheetName(b.title, i));
    });
    const meta = makeSheet(wideMetadataRows(reports));
    meta['!autofilter'] = { ref: meta['!ref'] };
    addSheet(wb, meta, 'Metadata (all samples)');
    for (const rep of reports) {
      addSheet(wb, makeSheet(blockRows(dataBlocks([rep]))), rep.sample);
    }
    return wb;
  }

  function toArrayBuffer(wb) {
    return XLSX.write(wb, { bookType: 'xlsx', type: 'array', compression: true });
  }

  return { buildSingle, buildCombined, toArrayBuffer, dataBlocks, metadataRows };
});

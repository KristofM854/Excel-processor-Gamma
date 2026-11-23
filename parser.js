/*
 * Gamma report parser.
 *
 * Reads the plain-text analysis report written by the gamma spectrometry
 * software (e.g. "Spain-GL-1 II.TXT") and returns:
 *   - metadata: every line before the first table, as { field, raw, value, unit, uncertainty, kind }
 *   - tables:   every boxed table (+---+ / | ... |) with its title, headers and rows
 *
 * Works in the browser (window.GammaParser) and in Node (require('./parser.js')).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GammaParser = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------- text decoding ----------

  // Lab PCs often write Windows-1252 ("±" as a single byte). Try strict UTF-8
  // first, fall back to Windows-1252 if the bytes are not valid UTF-8.
  function decodeBytes(bytes) {
    try {
      return new TextDecoder('utf-8', { fatal: true }).decode(bytes).replace(/^﻿/, '');
    } catch (e) {
      return new TextDecoder('windows-1252').decode(bytes);
    }
  }

  // ---------- value parsing ----------

  const NUM_RE = /^[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;
  // "14/08/2026 14:35:52" or "14/08/2026" (the software prints day/month/year)
  const DATE_RE = /^(\d{1,2})[\/.](\d{1,2})[\/.](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/;
  // "5.020 g ± 0.10 %", "266939.89 s", "0.02 %", "1"
  const QTY_RE = /^([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)(?:\s*([A-Za-zµ%°][A-Za-zµ%°\/.\s]{0,14}?))?(?:\s*(?:±|\+\/-|\+-)\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)\s*([A-Za-zµ%°\/]{0,10}))?\s*$/;

  function toNumber(s) {
    s = String(s).trim();
    return NUM_RE.test(s) ? Number(s) : null;
  }

  function parseDate(s) {
    const m = DATE_RE.exec(String(s).trim());
    if (!m) return null;
    const [, d, mo, y, h = '0', mi = '0', se = '0'] = m;
    const day = +d, month = +mo;
    if (month < 1 || month > 12 || day < 1 || day > 31) return null;
    return { year: +y, month, day, hour: +h, minute: +mi, second: +se, hasTime: m[4] !== undefined };
  }

  // Excel serial date number (1900 date system), timezone-independent.
  function excelSerial(dt) {
    const ms = Date.UTC(dt.year, dt.month - 1, dt.day, dt.hour, dt.minute, dt.second);
    return ms / 86400000 + 25569;
  }

  // Split a metadata value into number / unit / uncertainty where that is unambiguous.
  function splitValue(raw) {
    const s = raw.trim();
    if (s === '') return { kind: 'empty', value: null, unit: '', uncertainty: '' };
    const dt = parseDate(s);
    if (dt) return { kind: 'date', value: dt, unit: '', uncertainty: '' };
    const m = QTY_RE.exec(s);
    if (m) {
      return {
        kind: 'number',
        value: Number(m[1]),
        unit: (m[2] || '').trim(),
        uncertainty: m[3] !== undefined ? (m[3] + (m[4] ? ' ' + m[4] : '')).trim() : '',
      };
    }
    return { kind: 'text', value: null, unit: '', uncertainty: '' };
  }

  // ---------- table parsing ----------

  const isBorder = (l) => /^\s*\+[-+]+\+\s*$/.test(l);
  const isRow = (l) => /^\s*\|.*\|\s*$/.test(l);

  function splitRow(line) {
    const t = line.trim();
    return t.slice(1, -1).split('|').map((c) => c.trim());
  }

  // Multiplet marker before a number: "/ 238.86" (first peak), "\ 241.93" (last peak)
  const MULTIPLET_RE = /^([\/\\])\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)$/;

  function cellValue(s) {
    if (s === '') return null;
    const n = toNumber(s);
    return n !== null ? n : s;
  }

  function parseTable(title, lines) {
    // lines: from first border to last border (inclusive)
    const sections = [[]];
    for (const l of lines) {
      if (isBorder(l)) sections.push([]);
      else if (isRow(l)) sections[sections.length - 1].push(splitRow(l));
    }
    const nonEmpty = sections.filter((s) => s.length);
    const headerRows = nonEmpty[0] || [];
    const bodySections = nonEmpty.slice(1);
    const nCols = Math.max(0, ...nonEmpty.flat().map((r) => r.length));

    // Combine multi-line headers: "Activ" + "(Bq/kg)" -> "Activ (Bq/kg)"
    const headers = [];
    for (let c = 0; c < nCols; c++) {
      const parts = headerRows.map((r) => (r[c] || '').trim()).filter(Boolean);
      headers.push(parts.join(' ') || 'Column ' + (c + 1));
    }

    // Detect multiplet markers anywhere in the table
    let multipletCol = -1;
    for (const sec of bodySections) for (const r of sec) r.forEach((cell, c) => {
      if (multipletCol < 0 && MULTIPLET_RE.test(cell)) multipletCol = c;
    });

    const rows = [];
    let group = 0, open = false;
    bodySections.forEach((sec, si) => {
      for (const r of sec) {
        const cells = [];
        let marker = '', groupNo = null;
        for (let c = 0; c < nCols; c++) {
          const raw = r[c] === undefined ? '' : r[c];
          const m = c === multipletCol ? MULTIPLET_RE.exec(raw) : null;
          if (m) {
            marker = m[1];
            cells.push(Number(m[2]));
          } else {
            cells.push(cellValue(raw));
          }
        }
        if (multipletCol >= 0) {
          if (marker === '/') { group += 1; open = true; groupNo = group; }
          else if (marker === '\\') { groupNo = open ? group : ++group; open = false; }
          cells.splice(multipletCol + 1, 0, groupNo);
        }
        // footer rows (e.g. SUM) sit in a later section of the box
        rows.push({ cells, footer: si > 0 });
      }
    });

    if (multipletCol >= 0) headers.splice(multipletCol + 1, 0, 'Multiplet');
    return { title, headers, rows, hasMultiplet: multipletCol >= 0 };
  }

  // ---------- whole report ----------

  function parseReport(text, fileName) {
    const lines = text.replace(/\r\n?/g, '\n').split('\n');
    const warnings = [];

    // Locate tables: a run of border/row lines starting with a border.
    const tables = [];
    const tableStarts = [];
    let i = 0;
    while (i < lines.length) {
      if (isBorder(lines[i])) {
        let j = i;
        while (j + 1 < lines.length && (isBorder(lines[j + 1]) || isRow(lines[j + 1]))) j++;
        // title = nearest non-empty line above the box
        let k = i - 1;
        while (k >= 0 && lines[k].trim() === '') k--;
        const title = k >= 0 && !isBorder(lines[k]) && !isRow(lines[k]) ? lines[k].trim() : 'Table ' + (tables.length + 1);
        tableStarts.push(k >= 0 ? k : i);
        tables.push(parseTable(title, lines.slice(i, j + 1)));
        i = j + 1;
      } else i++;
    }
    if (!tables.length) warnings.push('No tables found in this file.');

    // Metadata: everything before the first table's title line.
    const metaEnd = tableStarts.length ? tableStarts[0] : lines.length;
    const metadata = [];
    let spectrum = '';
    let headerNo = 0;
    for (let n = 0; n < metaEnd; n++) {
      const line = lines[n];
      if (line.trim() === '') continue;
      const kv = /^\s*([^:]+?)\s*:\s*(.*?)\s*$/.exec(line);
      const sm = /^\s*Analysis of spectrum\s+(.+?)\s*$/i.exec(line);
      let field, raw;
      if (sm) { field = 'Spectrum'; raw = sm[1]; spectrum = raw; }
      else if (kv && !/^[A-Za-z]:\\/.test(line.trim())) { field = kv[1]; raw = kv[2]; }
      else { field = 'Report header ' + (++headerNo); raw = line.trim(); }
      metadata.push(Object.assign({ field, raw }, splitValue(raw)));
    }

    const sample = (spectrum || fileName || 'sample').replace(/\.[A-Za-z0-9]{1,4}$/, '').trim();
    return { fileName, sample, metadata, tables, warnings };
  }

  return { decodeBytes, parseReport, splitValue, parseDate, excelSerial };
});

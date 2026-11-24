/*
 * Core-level sheets for the combined workbook: Summary, Nuclear data, Profiles (with charts).
 *
 * Summary   : one row per sample; per nuclide Flag / Activity / Unc. / CL / MDA.
 * Nuclear data: half-lives and decay constants (same content as the lab's template).
 * Profiles  : depth profiles as native Excel scatter charts, fed by a formula table
 *             that references Summary and Nuclear data, so edits in Excel flow through.
 *
 * Browser: window.GammaCore. Node: require('./core.js')(XLSX, JSZip, GammaParser).
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory;
  else root.GammaCore = factory(root.XLSX, root.JSZip, root.GammaParser);
})(typeof self !== 'undefined' ? self : this, function (XLSX, JSZip, GammaParser) {
  'use strict';

  const enc = (r, c) => XLSX.utils.encode_cell({ r, c });
  const col = (c) => XLSX.utils.encode_col(c);
  const abs = (r, c) => '$' + col(c) + '$' + (r + 1);
  const LN2 = Math.LN2;
  const DAYS_PER_YEAR = 365.25;

  // ---------- Nuclear data (copied from the lab template) ----------
  // [nuclide, E (keV), probability (%), u, T1/2 (a), u]; lambda = LN(2)/T1/2, u(lambda) = lambda*u(T)/T
  const NUCLEAR_DATA = [
    ['Pb-210', 46.539, 4.252, 0.04, 22.23, 0.12],
    ['Pb-214', 295.224, 18.414, 0.36, 1600, 7],
    ['Pb-214', 351.932, 35.6, 0.07, 1600, 7],
    ['Bi-214', 609.312, 45.49, 0.19, 1600, 7],
    ['Cs-137', 661.657, 84.99, 0.2, 30.05, 0.08],
    ['Bi-214', 1120.29, 14.91, 0.03, 1600, 7],
    ['K-40', 1460.822, 10.55, 0.11, 1250400000, 30000000],
    ['Bi-214', 1764.5, 15.31, 0.05, 1600, 7],
    ['Th-234', 63.3, 3.75, 0.08, '24.1 d', 0.3],
    ['U-238', null, null, null, 4468000000, 5],
    ['Am-241', null, 35.92, 0.17, 432.6, 0.6],
  ];
  const ND_FIRST_ROW = 2; // 0-based row of 'Pb-210'
  const lambdaRef = (nuclide) => {
    const i = NUCLEAR_DATA.findIndex((d) => d[0] === nuclide);
    return { ref: "'Nuclear data'!" + abs(ND_FIRST_ROW + i, 6), value: LN2 / NUCLEAR_DATA[i][4] };
  };

  // ---------- sheet helper ----------
  function Sheet() {
    this.ws = {};
    this.maxR = 0; this.maxC = 0; this.widths = [];
  }
  Sheet.prototype.set = function (r, c, v) {
    if (v === null || v === undefined || v === '') return;
    let cell;
    if (typeof v === 'object' && !Array.isArray(v)) cell = Object.assign({}, v);
    else if (typeof v === 'number') cell = { t: 'n', v };
    else cell = { t: 's', v: String(v) };
    if (cell.t === undefined) cell.t = typeof cell.v === 'number' ? 'n' : 's';
    this.ws[enc(r, c)] = cell;
    const len = cell.t === 's' ? String(cell.v).length : cell.z && /yy/.test(cell.z) ? 16 : 9;
    this.widths[c] = Math.max(this.widths[c] || 0, len);
    this.maxR = Math.max(this.maxR, r); this.maxC = Math.max(this.maxC, c);
  };
  Sheet.prototype.done = function (fixedWidths) {
    this.ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: this.maxR, c: this.maxC } });
    this.ws['!cols'] = Array.from({ length: this.maxC + 1 }, (_, c) =>
      ({ wch: fixedWidths || Math.min(Math.max((this.widths[c] || 6) + 2, 8), 40) }));
    return this.ws;
  };

  const num = (v, f, z) => (v === null || v === undefined || Number.isNaN(v)
    ? { t: 'e', v: 0x2A, f, w: '#N/A' }
    : Object.assign({ t: 'n', v, f }, z ? { z } : {}));
  const dateCell = (dt, z) => ({ t: 'n', v: GammaParser.excelSerial(dt), z });

  // ---------- data model ----------

  // "CS-137E" -> "Cs-137", "AM-241" -> "Am-241"
  function normIso(name) {
    const m = /^([A-Za-z]{1,2})-?(\d{1,3})(m?)/.exec(String(name).trim());
    if (!m) return String(name).trim();
    return m[1].charAt(0).toUpperCase() + m[1].slice(1).toLowerCase() + '-' + m[2] + m[3];
  }

  const PREFERRED = ['K-40', 'Cs-137', 'Pb-210', 'Pb-212', 'Pb-214', 'Ac-228', 'Th-234', 'Am-241'];

  function activityValues(rep) {
    const t = rep.tables.find((x) => /activ|MDA/i.test(x.title));
    const out = new Map();
    if (!t) return out;
    const idx = (re) => t.headers.findIndex((h) => re.test(h));
    const iI = idx(/^isotope$/i), iA = idx(/^activ/i), iU = idx(/uncert.*bq/i), iC = idx(/^CL\b/i), iM = idx(/^MDA\b/i);
    if (iI < 0) return out;
    for (const r of t.rows) {
      if (r.footer) continue;
      const iso = r.cells[iI];
      if (typeof iso !== 'string') continue;
      const get = (i) => (i >= 0 && typeof r.cells[i] === 'number' ? r.cells[i] : null);
      const v = { act: get(iA), unc: get(iU), cl: get(iC), mda: get(iM) };
      v.flag = v.act === null ? (v.mda !== null ? '<' : '') : v.mda !== null && v.act < v.mda ? 'A' : '';
      out.set(normIso(iso), v);
    }
    return out;
  }

  function metaValue(rep, re) {
    const m = rep.metadata.find((x) => re.test(x.field));
    return m || null;
  }

  // Depth from a trailing section range in the sample name: "SAMO 14-2 2-3" -> 2.5
  function guessDepth(sample) {
    const m = /\s(\d+(?:[.,]\d+)?)\s*-\s*(\d+(?:[.,]\d+)?)\s*$/.exec(String(sample));
    if (!m) return null;
    const a = parseFloat(m[1].replace(',', '.')), b = parseFloat(m[2].replace(',', '.'));
    return b > a ? (a + b) / 2 : null;
  }

  function buildModel(reports, core) {
    core = core || {};
    const depths = core.depths || [];
    let samples = reports.map((rep, i) => {
      const date = metaValue(rep, /^acquisition date$/i);
      const mass = metaValue(rep, /^mass\s*\/\s*volume$/i) || metaValue(rep, /^mass/i);
      const live = metaValue(rep, /^acq\.?\s*time \(live\)/i) || metaValue(rep, /live/i);
      const d = depths[i];
      return {
        rep,
        depth: typeof d === 'number' && !Number.isNaN(d) ? d : null,
        date: date && date.kind === 'date' ? date.value : null,
        mass: mass && mass.kind === 'number' ? mass.value : null,
        live: live && live.kind === 'number' ? live.value : null,
        acts: activityValues(rep),
      };
    });
    if (samples.every((s) => s.depth !== null)) samples = samples.slice().sort((a, b) => a.depth - b.depth);
    const found = new Set();
    samples.forEach((s) => s.acts.forEach((_, k) => found.add(k)));
    const nuclides = [...PREFERRED.filter((n) => found.has(n)), ...[...found].filter((n) => !PREFERRED.includes(n))];
    return { samples, nuclides, coreName: core.name || '', sampling: core.samplingDate || null };
  }

  // ---------- Summary ----------
  const S_FIRST = 4;              // first sample row (0-based)
  const S_FIXED = 7;              // #, Sample ID, File, Depth, Measured date, Mass, Time
  const GROUP = ['Flag', 'Activity', 'Unc.', 'CL', 'MDA'];
  const groupStart = (model, nuc) => {
    const i = model.nuclides.indexOf(nuc);
    return i < 0 ? -1 : S_FIXED + i * GROUP.length;
  };

  function summarySheet(model) {
    const s = new Sheet();
    s.set(0, 0, 'Core name'); s.set(0, 1, model.coreName || null);
    s.set(0, 3, 'Sampling date');
    if (model.sampling) s.set(0, 4, dateCell(model.sampling, 'dd/mm/yyyy'));
    ['#', 'Sample ID', 'File', 'Depth', 'Measured date', 'Mass', 'Time'].forEach((h, c) => s.set(1, c, h));
    s.set(2, 6, 'measure');
    [null, null, null, '(cm)', null, '(g)', '(s)'].forEach((h, c) => s.set(3, c, h));
    model.nuclides.forEach((n) => {
      const g = groupStart(model, n);
      s.set(1, g + 1, n);
      GROUP.forEach((h, k) => { s.set(2, g + k, h); if (k > 0) s.set(3, g + k, '(Bq/kg)'); });
    });
    model.samples.forEach((smp, i) => {
      const r = S_FIRST + i;
      s.set(r, 0, i + 1);
      s.set(r, 1, smp.rep.sample);
      s.set(r, 2, smp.rep.fileName);
      if (smp.depth !== null) s.set(r, 3, smp.depth);
      if (smp.date) s.set(r, 4, dateCell(smp.date, 'dd/mm/yyyy hh:mm'));
      s.set(r, 5, smp.mass);
      s.set(r, 6, smp.live);
      model.nuclides.forEach((n) => {
        const v = smp.acts.get(n);
        if (!v) return;
        const g = groupStart(model, n);
        s.set(r, g, v.flag || null);
        s.set(r, g + 1, v.act); s.set(r, g + 2, v.unc); s.set(r, g + 3, v.cl); s.set(r, g + 4, v.mda);
      });
    });
    const last = S_FIRST + model.samples.length - 1;
    const mean = last + 2;
    s.set(mean, 1, 'Mean'); s.set(mean + 1, 1, 'Std. dev.');
    model.nuclides.forEach((n) => {
      const c = groupStart(model, n) + 1;
      const vals = model.samples.map((x) => (x.acts.get(n) || {}).act).filter((x) => typeof x === 'number');
      const range = col(c) + (S_FIRST + 1) + ':' + col(c) + (last + 1);
      if (vals.length >= 1) s.set(mean, c, num(vals.reduce((a, b) => a + b, 0) / vals.length, 'AVERAGE(' + range + ')'));
      if (vals.length >= 2) {
        const m = vals.reduce((a, b) => a + b, 0) / vals.length;
        s.set(mean + 1, c, num(Math.sqrt(vals.reduce((a, b) => a + (b - m) ** 2, 0) / (vals.length - 1)), 'STDEV(' + range + ')'));
      }
    });
    const notes = [
      'Flag  <  : below the detection limit. No activity reported; the MDA is shown.',
      'Flag  A  : activity reported, but below the MDA.',
      'Activity, Unc., CL and MDA are copied from the reports (Unc. = uncertainty in Bq/kg, CL = decision threshold, MDA = minimum detectable activity).',
      'Core name, sampling date (E1) and depth (column D) can be filled in or changed here; the Profiles sheet updates automatically.',
      'Mean and Std. dev. use detected activities only.',
    ];
    notes.forEach((t, k) => s.set(mean + 3 + k, 0, t));
    const ws = s.done();
    ws['!cols'][0] = { wch: 6 }; ws['!cols'][1] = { wch: 22 };
    return ws;
  }

  // ---------- Nuclear data ----------
  function nuclearSheet() {
    const s = new Sheet();
    ['Radionúclido', 'E ', 'Probab.', 'u', 'T1/2', 'u', 'l ', 'u'].forEach((h, c) => s.set(0, c, h));
    [null, '(keV)', '(%)', null, '(a)', null, '(a)', 'u'].forEach((h, c) => s.set(1, c, h));
    NUCLEAR_DATA.forEach((d, i) => {
      const r = ND_FIRST_ROW + i;
      d.forEach((v, c) => s.set(r, c, v));
      if (typeof d[4] === 'number') {
        const lam = LN2 / d[4];
        s.set(r, 6, num(lam, 'LN(2)/E' + (r + 1), '0.0000'));
        s.set(r, 7, num(lam * d[5] / d[4], 'G' + (r + 1) + '*F' + (r + 1) + '/E' + (r + 1), '0.0000'));
      }
    });
    s.set(ND_FIRST_ROW + NUCLEAR_DATA.length + 1, 0, 'l = decay constant = LN(2)/T1/2 (per year). Used in Profiles for the decay correction to the sampling date.');
    const ws = s.done();
    ws['!cols'] = [14, 10, 9, 7, 14, 12, 10, 10].map((w) => ({ wch: w }));
    return ws;
  }

  // ---------- Profiles ----------
  const P_TABLE = 46;   // 0-based row of the profile-data header
  const P_FIRST = P_TABLE + 2;

  function profilesSheet(model) {
    const s = new Sheet();
    const n = model.samples.length;
    const sumRef = (r, c) => 'Summary!' + col(c) + (r + 1);
    const samplingRef = 'Summary!' + abs(0, 4);
    const sampSerial = model.sampling ? GammaParser.excelSerial(model.sampling) : null;

    // value/error expression for one nuclide: MDA without error bar when flagged '<'
    function plotted(nuc, r, smp) {
      const g = groupStart(model, nuc);
      if (g < 0) return { vf: 'NA()', v: null, ef: '0', e: 0 };
      const F = sumRef(r, g), A = sumRef(r, g + 1), U = sumRef(r, g + 2), M = sumRef(r, g + 4);
      const v = smp.acts.get(nuc) || {};
      const value = v.flag === '<' ? v.mda : v.act;
      return {
        vf: 'IF(' + F + '="<",' + M + ',IF(' + A + '="",NA(),' + A + '))',
        v: value === undefined ? null : value,
        ef: 'IF(' + F + '="<",0,N(' + U + '))',
        e: v.flag === '<' ? 0 : v.unc || 0,
      };
    }

    const cols = [
      { key: 'K-40', label: 'K-40' },
      { key: 'Pb-210', label: 'Pb-210' },
      { key: 'Pb-214', label: 'Pb-214' },
      { key: 'exPb', label: 'exPb-210 corr' },
      { key: 'csCorr', label: 'Cs-137 corr' },
      { key: 'Am-241', label: 'Am-241' },
      { key: 'Th-234', label: 'Th-234' },
      { key: 'Pb-212', label: 'Pb-212' },
      { key: 'Ac-228', label: 'Ac-228' },
    ];
    const colOf = {};
    cols.forEach((c, i) => { colOf[c.key] = 2 + 2 * i; });

    s.set(0, 0, 'Depth profiles');
    s.set(1, 0, '* Points without an error bar are below the detection limit and show the MDA.');
    s.set(2, 0, 'exPb-210 corr = (Pb-210 − Pb-214), and Cs-137 corr, are decay-corrected to the sampling date (Summary!E1) with the decay constants in Nuclear data. Both stay empty until a sampling date is entered.');
    s.set(P_TABLE - 1, 0, 'Profile data (calculated from the Summary sheet; #N/A = no value, e.g. depth not entered yet)');
    s.set(P_TABLE, 0, 'Sample'); s.set(P_TABLE, 1, 'Depth');
    s.set(P_TABLE + 1, 1, '(cm)');
    cols.forEach((c) => {
      s.set(P_TABLE, colOf[c.key], c.label); s.set(P_TABLE, colOf[c.key] + 1, '±');
      s.set(P_TABLE + 1, colOf[c.key], '(Bq/kg)'); s.set(P_TABLE + 1, colOf[c.key] + 1, '(Bq/kg)');
    });

    const l210 = lambdaRef('Pb-210'), lCs = lambdaRef('Cs-137');
    model.samples.forEach((smp, i) => {
      const r = P_FIRST + i, sr = S_FIRST + i;
      s.set(r, 0, { t: 's', v: smp.rep.sample, f: sumRef(sr, 1) });
      s.set(r, 1, num(smp.depth, 'IF(' + sumRef(sr, 3) + '="",NA(),' + sumRef(sr, 3) + ')'));
      const dtF = '(' + sumRef(sr, 4) + '-' + samplingRef + ')/' + DAYS_PER_YEAR;
      const dt = sampSerial !== null && smp.date ? (GammaParser.excelSerial(smp.date) - sampSerial) / DAYS_PER_YEAR : null;

      for (const key of ['K-40', 'Pb-210', 'Pb-214', 'Am-241', 'Th-234', 'Pb-212', 'Ac-228']) {
        const p = plotted(key, sr, smp);
        s.set(r, colOf[key], num(p.v, p.vf));
        s.set(r, colOf[key] + 1, num(p.e, p.ef));
      }

      // excess Pb-210, decay-corrected
      const g210 = groupStart(model, 'Pb-210'), g214 = groupStart(model, 'Pb-214');
      const exCol = colOf.exPb;
      if (g210 >= 0 && g214 >= 0) {
        const A1 = sumRef(sr, g210 + 1), A2 = sumRef(sr, g214 + 1), U1 = sumRef(sr, g210 + 2), U2 = sumRef(sr, g214 + 2);
        const a = smp.acts.get('Pb-210') || {}, b = smp.acts.get('Pb-214') || {};
        const ok = typeof a.act === 'number' && typeof b.act === 'number' && dt !== null;
        s.set(r, exCol, num(ok ? (a.act - b.act) * Math.exp(l210.value * dt) : null,
          'IF(OR(' + A1 + '="",' + A2 + '="",' + samplingRef + '=""),NA(),(' + A1 + '-' + A2 + ')*EXP(' + l210.ref + '*' + dtF + '))'));
        s.set(r, exCol + 1, num(ok ? Math.sqrt((a.unc || 0) ** 2 + (b.unc || 0) ** 2) : 0,
          'IF(ISNA(' + col(exCol) + (r + 1) + '),0,SQRT(N(' + U1 + ')^2+N(' + U2 + ')^2))'));
      } else {
        s.set(r, exCol, num(null, 'NA()')); s.set(r, exCol + 1, num(0, '0'));
      }

      // Cs-137, decay-corrected
      const cs = plotted('Cs-137', sr, smp);
      const csOk = cs.v !== null && dt !== null;
      s.set(r, colOf.csCorr, num(csOk ? cs.v * Math.exp(lCs.value * dt) : null,
        'IF(' + samplingRef + '="",NA(),' + cs.vf + '*EXP(' + lCs.ref + '*' + dtF + '))'));
      s.set(r, colOf.csCorr + 1, num(cs.e, cs.ef));
    });

    const ws = s.done(11);
    ws['!cols'][0] = { wch: 18 };

    // chart definitions
    const rng = (c) => "'Profiles'!" + abs(P_FIRST, c) + ':' + abs(P_FIRST + Math.max(n, 1) - 1, c);
    const ser = (key, name) => ({ name, x: rng(colOf[key]), err: rng(colOf[key] + 1), y: rng(1) });
    const charts = [
      { title: 'K-40 (Bq/kg)', series: [ser('K-40', 'K-40')] },
      { title: 'exPb-210 corr (Bq/kg)', series: [ser('exPb', 'exPb-210 corr')] },
      { title: 'Pb-214 and Pb-210 (Bq/kg)', series: [ser('Pb-214', 'Pb-214'), ser('Pb-210', 'Pb-210')] },
      { title: 'Cs-137 corr and Am-241 (Bq/kg)', series: [ser('csCorr', 'Cs-137 corr'), ser('Am-241', 'Am-241')] },
      { title: 'Th-234 and Pb-214 (Bq/kg)', series: [ser('Th-234', 'Th-234'), ser('Pb-214', 'Pb-214')] },
      { title: 'Pb-212 and Ac-228 (Bq/kg)', series: [ser('Pb-212', 'Pb-212'), ser('Ac-228', 'Ac-228')] },
    ];
    // 3 charts per row, each 6 columns x 20 rows, starting at row 4
    charts.forEach((ch, i) => {
      ch.from = { c: (i % 3) * 6, r: 4 + Math.floor(i / 3) * 21 };
      ch.to = { c: ch.from.c + 6, r: ch.from.r + 20 };
    });
    return { ws, charts };
  }

  // ---------- chart XML ----------
  const xmlEsc = (s) => String(s).replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const COLORS = ['1F4E79', 'C55A11'];
  const SYMBOLS = ['circle', 'diamond'];

  function chartXml(ch) {
    const ln = (clr, w) => '<a:ln w="' + w + '"><a:solidFill><a:srgbClr val="' + clr + '"/></a:solidFill></a:ln>';
    const series = ch.series.map((s, i) => {
      const clr = COLORS[i % COLORS.length];
      const ref = (f) => '<c:numRef><c:f>' + xmlEsc(f) + '</c:f></c:numRef>';
      return '<c:ser><c:idx val="' + i + '"/><c:order val="' + i + '"/><c:tx><c:v>' + xmlEsc(s.name) + '</c:v></c:tx>' +
        '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>' +
        '<c:marker><c:symbol val="' + SYMBOLS[i % 2] + '"/><c:size val="6"/><c:spPr><a:solidFill><a:srgbClr val="' + clr + '"/></a:solidFill>' + ln(clr, 9525) + '</c:spPr></c:marker>' +
        '<c:errBars><c:errDir val="x"/><c:errBarType val="both"/><c:errValType val="cust"/><c:noEndCap val="0"/>' +
        '<c:plus>' + ref(s.err) + '</c:plus><c:minus>' + ref(s.err) + '</c:minus><c:spPr>' + ln(clr, 9525) + '</c:spPr></c:errBars>' +
        '<c:xVal>' + ref(s.x) + '</c:xVal><c:yVal>' + ref(s.y) + '</c:yVal><c:smooth val="0"/></c:ser>';
    }).join('');
    const text = (t, sz, bold, rot) => '<c:tx><c:rich><a:bodyPr' + (rot ? ' rot="-5400000" vert="horz"' : '') + '/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="' + sz + '" b="' + (bold ? 1 : 0) + '"/></a:pPr><a:r><a:rPr lang="en-US" sz="' + sz + '" b="' + (bold ? 1 : 0) + '"/><a:t>' + xmlEsc(t) + '</a:t></a:r></a:p></c:rich></c:tx>';
    const axis = (id, cross, pos, orient, title) => '<c:valAx><c:axId val="' + id + '"/><c:scaling><c:orientation val="' + orient + '"/></c:scaling>' +
      '<c:delete val="0"/><c:axPos val="' + pos + '"/>' + (pos === 'l' ? '<c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="E0E0E0"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : '') +
      (title ? '<c:title>' + text(title, 900, false, pos === 'l') + '<c:overlay val="0"/></c:title>' : '') +
      '<c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
      '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>' +
      '<c:crossAx val="' + cross + '"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<c:roundedCorners val="0"/><c:chart><c:title>' + text(ch.title, 1100, true) + '<c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>' +
      '<c:plotArea><c:layout/><c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>' + series +
      '<c:axId val="5001"/><c:axId val="5002"/></c:scatterChart>' +
      axis(5001, 5002, 'b', 'minMax', null) + axis(5002, 5001, 'l', 'maxMin', 'Depth (cm)') + '</c:plotArea>' +
      (ch.series.length > 1 ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '') +
      '<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>';
  }

  function drawingXml(charts) {
    const anchors = charts.map((ch, i) =>
      '<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>' + ch.from.c + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ch.from.r + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
      '<xdr:to><xdr:col>' + ch.to.c + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + ch.to.r + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
      '<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="' + (i + 2) + '" name="Chart ' + (i + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
      '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      '<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId' + (i + 1) + '"/>' +
      '</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>').join('');
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' + anchors + '</xdr:wsDr>';
  }

  const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const R_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';

  // Add the charts to an .xlsx written by SheetJS (which cannot write charts itself).
  async function injectCharts(buf, sheetName, charts) {
    const zip = await JSZip.loadAsync(buf);
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const relsXml = await zip.file('xl/_rels/workbook.xml.rels').async('string');
    const sheetTag = new RegExp('<sheet [^>]*name="' + sheetName + '"[^>]*/>').exec(wbXml);
    const rid = /r:id="([^"]+)"/.exec(sheetTag[0])[1];
    const target = new RegExp('<Relationship [^>]*Id="' + rid + '"[^>]*>').exec(relsXml)[0].match(/Target="([^"]+)"/)[1];
    const sheetPath = 'xl/' + target.replace(/^\/?xl\//, '');
    const sheetFile = sheetPath.split('/').pop();
    const sheetRelsPath = sheetPath.replace(sheetFile, '_rels/' + sheetFile + '.rels');

    // charts + drawing
    charts.forEach((ch, i) => zip.file('xl/charts/chart' + (i + 1) + '.xml', chartXml(ch)));
    zip.file('xl/drawings/drawing1.xml', drawingXml(charts));
    zip.file('xl/drawings/_rels/drawing1.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + REL_NS + '">' +
      charts.map((_, i) => '<Relationship Id="rId' + (i + 1) + '" Type="' + R_TYPE + 'chart" Target="../charts/chart' + (i + 1) + '.xml"/>').join('') + '</Relationships>');

    // sheet -> drawing relationship
    const drawRel = '<Relationship Id="rIdGammaDrawing" Type="' + R_TYPE + 'drawing" Target="../drawings/drawing1.xml"/>';
    const existing = zip.file(sheetRelsPath);
    if (existing) {
      const x = await existing.async('string');
      zip.file(sheetRelsPath, x.replace('</Relationships>', drawRel + '</Relationships>'));
    } else {
      zip.file(sheetRelsPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + REL_NS + '">' + drawRel + '</Relationships>');
    }
    let sx = await zip.file(sheetPath).async('string');
    if (!/xmlns:r=/.test(sx)) sx = sx.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    sx = sx.replace(/<(legacyDrawing|tableParts|extLst)\b|<\/worksheet>/, (m) => '<drawing r:id="rIdGammaDrawing"/>' + m);
    zip.file(sheetPath, sx);

    // content types
    let ct = await zip.file('[Content_Types].xml').async('string');
    const ov = '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' +
      charts.map((_, i) => '<Override PartName="/xl/charts/chart' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>').join('');
    ct = ct.replace('</Types>', ov + '</Types>');
    zip.file('[Content_Types].xml', ct);

    // recalculate formulas when the file is opened
    let wx = wbXml;
    if (/<calcPr\b/.test(wx)) wx = wx.replace(/<calcPr\b/, '<calcPr fullCalcOnLoad="1"');
    else wx = wx.replace(/<(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/, (m) => '<calcPr fullCalcOnLoad="1"/>' + m);
    zip.file('xl/workbook.xml', wx);

    return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
  }

  return { buildModel, summarySheet, nuclearSheet, profilesSheet, injectCharts, guessDepth, normIso, chartXml };
});

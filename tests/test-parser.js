// Run: node tests/test-parser.js
// Uses a small synthetic report (made-up numbers) in the same layout as the real ones.
const assert = require('assert');
const P = require('../parser.js');

const report = [
  '                             LAB/TEST',
  '                       Radionuclides in sediments',
  '',
  '          Analysis of spectrum Test-1 A.SPE',
  '',
  'Directory:                    C:\\Data\\TEST',
  'Acq.time (live):              1000.50 s',
  'Acquisition date:             14/08/2026 14:35:52',
  'Mass/Volume:                  5.020 g ± 0.10 %',
  'FWHM:                         1.09 keV - 2.16 keV',
  'Comment:',
  '',
  '          List by energies (with confirmed isotopes)',
  '+---+---------+-----+-----+---------+',
  '|No.| Energy  |FWHM |FWTM |ISOTOPE  |',
  '+---+---------+-----+-----+---------+',
  '|  1|   46.07 | 1.14|    *|PB-210   |',
  '|  2|/ 238.86 | 1.31| 2.41|PB-212   |',
  '|  3|\\ 241.93 | 1.78|    *|         |',
  '+---+---------+-----+-----+---------+',
  '',
  '          List activity / MDA',
  '+---+---------+-------------+---------------+',
  '|No.| Isotope |    Activ    |      MDA      |',
  '|   |         |   (Bq/kg)   |    (Bq/kg)    |',
  '+---+---------+-------------+---------------+',
  '|  1|BE-7     |             |   2.083E+01   |',
  '|  2|K-40     |  8.518E+02  |   8.760E+01   |',
  '+---+---------+-------------+---------------+',
  '|   |SUM      |  1.367E+03  |               |',
  '+---+---------+-------------+---------------+',
].join('\r\n');

const r = P.parseReport(report, 'Test-1 A.TXT');
const meta = Object.fromEntries(r.metadata.map((m) => [m.field, m]));

assert.strictEqual(r.sample, 'Test-1 A');
assert.strictEqual(meta['Report header 1'].raw, 'LAB/TEST');
assert.strictEqual(meta['Directory'].raw, 'C:\\Data\\TEST');
assert.strictEqual(meta['Acq.time (live)'].value, 1000.5);
assert.strictEqual(meta['Acq.time (live)'].unit, 's');
assert.strictEqual(meta['Mass/Volume'].value, 5.02);
assert.strictEqual(meta['Mass/Volume'].unit, 'g');
assert.strictEqual(meta['Mass/Volume'].uncertainty, '0.10 %');
assert.strictEqual(meta['Acquisition date'].kind, 'date');
assert.strictEqual(P.excelSerial(meta['Acquisition date'].value).toFixed(5), '46248.60824');
assert.strictEqual(meta['FWHM'].kind, 'text');
assert.strictEqual(meta['Comment'].kind, 'empty');

assert.strictEqual(r.tables.length, 2);
const [peaks, act] = r.tables;
assert.strictEqual(peaks.title, 'List by energies (with confirmed isotopes)');
assert.deepStrictEqual(peaks.headers, ['No.', 'Energy', 'Multiplet', 'FWHM', 'FWTM', 'ISOTOPE']);
assert.deepStrictEqual(peaks.rows[0].cells, [1, 46.07, null, 1.14, '*', 'PB-210']);
assert.deepStrictEqual(peaks.rows[1].cells, [2, 238.86, 1, 1.31, 2.41, 'PB-212']);
assert.deepStrictEqual(peaks.rows[2].cells, [3, 241.93, 1, 1.78, '*', null]);

assert.deepStrictEqual(act.headers, ['No.', 'Isotope', 'Activ (Bq/kg)', 'MDA (Bq/kg)']);
assert.deepStrictEqual(act.rows[0].cells, [1, 'BE-7', null, 20.83]);
assert.deepStrictEqual(act.rows[1].cells, [2, 'K-40', 851.8, 87.6]);
assert.deepStrictEqual(act.rows[2], { cells: [null, 'SUM', 1367, null], footer: true });

// Windows-1252 input ("±" as byte 0xB1) decodes correctly
const bytes = Uint8Array.from([0x35, 0x20, 0x67, 0x20, 0xb1, 0x20, 0x31, 0x20, 0x25]);
assert.strictEqual(P.decodeBytes(bytes), '5 g ± 1 %');

console.log('All parser tests passed.');

// ---------- workbook layout ----------
const XLSX = require('../vendor/xlsx.full.min.js');
const JSZip = require('../vendor/jszip.min.js');
const C = require('../core.js')(XLSX, JSZip, P);
const W = require('../workbook.js')(XLSX, P, C);
const [wide, peakList] = W.dataBlocks([r]);
assert.deepStrictEqual(wide.headers.slice(0, 13),
  ['Sample', 'File', 'Detector', 'Acq.time (live) (s)', 'Acq.time (real)', 'Dead time', 'Count rate',
    'Acquisition date', 'Mass/Volume (g)', 'Mass/Volume uncertainty (%)',
    'Calibration date', 'Isotope table', 'Efficiency function']);
assert.strictEqual(wide.rows[0][3], 1000.5);
assert.deepStrictEqual(peakList.headers.slice(0, 6),
  ['Sample', 'File', 'Acquisition date', 'Detector', 'Mass/Volume (g)', 'Mass/Volume uncertainty (%)']);
assert.strictEqual(wide.rows.length, 1); // one row per sample
const col = (h) => wide.rows[0][wide.headers.indexOf(h)];
assert.strictEqual(col('Mass/Volume (g)'), 5.02);
assert.strictEqual(col('Mass/Volume uncertainty (%)'), 0.1);
assert.strictEqual(col('Detector'), null); // not in the synthetic report
assert.strictEqual(col('K-40 Activ (Bq/kg)'), 851.8);
assert.strictEqual(col('BE-7 Activ (Bq/kg)'), null);
assert.strictEqual(col('BE-7 MDA (Bq/kg)'), 20.83);
assert.strictEqual(col('SUM Activ (Bq/kg)'), 1367);
assert.ok(!wide.headers.includes('SUM MDA (Bq/kg)'));
assert.strictEqual(peakList.title, 'List by energies (with confirmed isotopes)');
assert.strictEqual(peakList.rows.length, 3);
assert.strictEqual(peakList.headers[0], 'Sample');

// combined: one row per sample on the data sheet
const r2 = P.parseReport(report.replace('Test-1 A.SPE', 'Test-2.SPE'), 'Test-2.TXT');
const wb = W.buildCombined([r, r2]);
assert.deepStrictEqual(wb.SheetNames.slice(0, 6),
  ['Profiles', 'Summary', 'Nuclear data', 'Data (all samples)', 'Peaks (all samples)', 'Metadata (all samples)']);
const data = XLSX.utils.sheet_to_json(wb.Sheets['Data (all samples)'], { header: 1 });
assert.strictEqual(data.length, 3);

console.log('All workbook tests passed.');

// ---------- Summary / Profiles / Nuclear data ----------
assert.strictEqual(C.normIso('CS-137E'), 'Cs-137');
assert.strictEqual(C.normIso('AM-241'), 'Am-241');
assert.strictEqual(C.guessDepth('SAMO 14-2 2-3'), 2.5);
assert.strictEqual(C.guessDepth('Spain-GL-1 II'), null);
assert.strictEqual(C.guessDepth('SAMO 14-2'), null); // 14-2 is not an increasing range

const rich = P.parseReport(report
  .replace('Comment:', 'Acquisition date:             14/08/2026 14:35:52\r\nComment:')
  .replace('|  2|K-40     |  8.518E+02  |   8.760E+01   |',
    '|  2|K-40     |  8.518E+02  |   8.760E+01   |\r\n' +
    '|  3|PB-210   |  3.000E+02  |   3.000E+01   |\r\n' +
    '|  4|PB-214   |  5.000E+01  |   8.000E+00   |'),
  'Test-1 A.TXT');
const sampling = { year: 2026, month: 3, day: 1, hour: 0, minute: 0, second: 0 };
const model = C.buildModel([rich], { samplingDate: sampling, depths: [2.5] });
assert.deepStrictEqual(model.nuclides.slice(0, 3), ['K-40', 'Pb-210', 'Pb-214']);
assert.strictEqual(model.samples[0].acts.get('Be-7').flag, '<');
const sum = C.summarySheet(model);
assert.strictEqual(sum.B5.v, 'Test-1 A');
assert.strictEqual(sum.D5.v, 2.5);
const prof = C.profilesSheet(model);
assert.strictEqual(prof.charts.length, 6);
// exPb-210 corr = (300 - 50) * exp(ln2/22.23 * dt), dt in years from sampling to measurement
const measured = { year: 2026, month: 8, day: 14, hour: 14, minute: 35, second: 52 };
const dt = (P.excelSerial(measured) - P.excelSerial(sampling)) / 365.25;
const ex = Object.values(prof.ws).find((c) => c && c.f && /^IF\(OR\(/.test(c.f));
assert.ok(Math.abs(ex.v - 250 * Math.exp(Math.LN2 / 22.23 * dt)) < 1e-9);
const nd = C.nuclearSheet();
assert.strictEqual(nd.A3.v, 'Pb-210');
assert.strictEqual(nd.G3.f, 'LN(2)/E3');

(async () => {
  const buf = await W.toArrayBuffer(W.buildCombined([rich], { depths: [2.5] }));
  const zip = await JSZip.loadAsync(buf);
  assert.ok(zip.file('xl/charts/chart6.xml'));
  assert.ok(/<drawing r:id=/.test(await zip.file('xl/worksheets/sheet1.xml').async('string')));
  console.log('All Summary/Profiles tests passed.');
})().catch((e) => { console.error(e); process.exit(1); });

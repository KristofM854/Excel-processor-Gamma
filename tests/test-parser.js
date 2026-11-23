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
const W = require('../workbook.js')(XLSX, P);
const [wide, peakList] = W.dataBlocks([r]);
assert.deepStrictEqual(wide.headers.slice(0, 6),
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
assert.deepStrictEqual(wb.SheetNames.slice(0, 3), ['Data (all samples)', 'Peaks (all samples)', 'Metadata (all samples)']);
const data = XLSX.utils.sheet_to_json(wb.Sheets['Data (all samples)'], { header: 1 });
assert.strictEqual(data.length, 3);

console.log('All workbook tests passed.');

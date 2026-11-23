# Gamma Report to Excel

A web page that converts gamma spectrometry analysis reports (`.TXT`, e.g.
`Spain-GL-1 II.TXT`) into Excel workbooks. It runs entirely in the browser:
there is nothing to install, and files are never uploaded.

**Open the app: https://kristofmoeller.com/Excel-processor-Gamma/**

## For users

1. Open the page link.
2. Drag `.TXT` report files or a whole folder onto the page, or use
   **Choose files** / **Choose folder**.
3. Pick the output:
   - **One Excel file per report**: each workbook has a `metadata` sheet and
     a `data` sheet. Several files download as one ZIP.
   - **One combined workbook**: `Data (all samples)` with one row per
     sample, `Peaks (all samples)`, `Metadata (all samples)`, plus one data
     sheet per sample.
4. Click **Create Excel**. The file appears in your Downloads folder.

## What ends up in the workbook

**metadata**: every line above the first table.

| Column | Meaning |
|---|---|
| Field | Text before the colon (`Mass/Volume`) |
| As printed | Exactly what the report says (`5.020 g ± 0.10 %`) |
| Value | The number, as a real Excel number (`5.02`), or a real Excel date |
| Unit | `g` |
| Uncertainty | `0.10 %` |

Value, Unit and Uncertainty are only filled in when the text is clearly a
number (with optional unit and ± uncertainty) or a `dd/mm/yyyy hh:mm:ss`
date. Anything else (ranges, file names, versions) stays in *As printed* only.

**data**: two blocks, each under its table title.

1. **One row per sample** (title `List activity / MDA`): Sample, File,
   Acquisition date, Detector, Mass/Volume (g), Mass/Volume uncertainty (%),
   then every isotope's values side by side: `K-40 Activ (Bq/kg)`,
   `K-40 Uncert. (%)`, `K-40 Uncert. (Bq/kg)`, `K-40 CL (Bq/kg)`,
   `K-40 MDA (Bq/kg)`, and so on, ending with the `SUM` values. Nuclides
   below the detection limit have empty activity cells but keep CL and MDA.
2. **Peak list** (`List by energies ...`): one row per peak, with the same
   sample columns in front.

Date, detector and mass also stay in the metadata sheet.

- Two-line headers are joined: `Activ` + `(Bq/kg)` → `Activ (Bq/kg)`.
- The report marks overlapping (deconvoluted) peaks with `/` and `\` in front
  of the energy, e.g. `/ 238.86`. That marker is moved into a separate
  **Multiplet** column, where peaks fitted together share a group number, so
  that *Energy* stays a number.
- Empty cells in the report stay empty, not zero. `*` in FWTM is kept as text.

## Publishing the page (GitHub Pages)

1. Merge this branch into `main`.
2. In the repository on GitHub: **Settings → Pages → Build and deployment →
   Source: Deploy from a branch**, branch `main`, folder `/ (root)`, **Save**.
3. After a minute the page is live at
   https://kristofmoeller.com/Excel-processor-Gamma/ (the custom domain of
   `kristofm854.github.io`; the path is case-sensitive). Send that link to
   users. Every later push to `main` updates the page automatically; users
   just reload it.

GitHub Pages on a private repository needs a paid GitHub plan. The page itself
holds no data, so a public repository is fine.

## Development

- `parser.js`: reads a report into metadata + tables.
- `workbook.js`: builds the Excel workbooks.
- `index.html`: the page.
- `vendor/`: SheetJS 0.18.5 (writes `.xlsx`) and JSZip 3.10.1 (ZIP), stored
  locally so the page has no outside dependencies. SheetJS is only used to
  write files, never to read untrusted spreadsheets.

Run the tests with Node 18+:

```bash
node tests/test-parser.js
```

To try the page locally, open `index.html` in a browser.

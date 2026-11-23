# Gamma Report to Excel

A web page that converts gamma spectrometry analysis reports (`.TXT`, e.g.
`Spain-GL-1 II.TXT`) into Excel workbooks. It runs entirely in the browser:
there is nothing to install, and files are never uploaded.

## For users

1. Open the page link.
2. Drag `.TXT` report files or a whole folder onto the page, or use
   **Choose files** / **Choose folder**.
3. Pick the output:
   - **One Excel file per report**: each workbook has a `metadata` sheet and
     a `data` sheet. Several files download as one ZIP.
   - **One combined workbook**: `Activities (all samples)`,
     `Peaks (all samples)` and `Metadata (all samples)`, each with one row per
     sample (and nuclide/peak), plus one data sheet per sample.
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

**data**: each table with its title, a header row and the data rows.

- Two-line headers are joined: `Activ` + `(Bq/kg)` → `Activ (Bq/kg)`.
- The report marks overlapping (deconvoluted) peaks with `/` and `\` in front
  of the energy, e.g. `/ 238.86`. That marker is moved into a separate
  **Multiplet** column, where peaks fitted together share a group number, so
  that *Energy* stays a number.
- Empty cells in the report (e.g. nuclides below the detection limit) stay
  empty, not zero. `*` in FWTM is kept as text. The `SUM` row is kept.

## Publishing the page (GitHub Pages)

1. Merge this branch into `main`.
2. In the repository on GitHub: **Settings → Pages → Build and deployment →
   Source: Deploy from a branch**, branch `main`, folder `/ (root)`, **Save**.
3. After a minute the page is live at
   `https://<github-user>.github.io/<repository-name>/`. Send that link to
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

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
   - **One combined workbook** (for a core or any series of samples):
     `Profiles`, `Summary`, `Nuclear data`, `Data (all samples)`,
     `Peaks (all samples)`, `Metadata (all samples)`, plus one data sheet per
     sample. Optionally enter the core name, sampling date and each sample's
     depth on the page, or skip that and fill them in later in the `Summary`
     sheet (B1, E1 and the Depth column). The profiles update automatically.
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
   Detector, Acq.time (live) (s), Acq.time (real) (s), Dead time (%),
   Count rate (cps), Acquisition date, Mass/Volume (g),
   Mass/Volume uncertainty (%), Calibration date, Isotope table,
   Efficiency function, then every isotope's values side by side: `K-40 Activ (Bq/kg)`,
   `K-40 Uncert. (%)`, `K-40 Uncert. (Bq/kg)`, `K-40 CL (Bq/kg)`,
   `K-40 MDA (Bq/kg)`, and so on, ending with the `SUM` values. Nuclides
   below the detection limit have empty activity cells but keep CL and MDA.
2. **Peak list** (`List by energies ...`): one row per peak, with Sample,
   File, Acquisition date, Detector, Mass/Volume and its uncertainty in front.

Date, detector and mass also stay in the metadata sheet.

- Two-line headers are joined: `Activ` + `(Bq/kg)` → `Activ (Bq/kg)`.
- The report marks overlapping (deconvoluted) peaks with `/` and `\` in front
  of the energy, e.g. `/ 238.86`. That marker is moved into a separate
  **Multiplet** column, where peaks fitted together share a group number, so
  that *Energy* stays a number.
- Empty cells in the report stay empty, not zero. `*` in FWTM is kept as text.

### Combined workbook: Summary, Nuclear data, Profiles

**Summary**: core name and sampling date at the top, then one row per sample
(sorted by depth when every depth is known): #, Sample ID, File, Depth (cm),
Measured date, Mass (g), Time measure (s), and per nuclide
*Flag / Activity / Unc. / CL / MDA* (all Bq/kg). Isotope names are written as
`Cs-137`, `Am-241`, ... (the report's `CS-137E` becomes `Cs-137`).

- Flag `<`: below the detection limit. No activity reported; the MDA is shown.
- Flag `A`: activity reported, but below the MDA.
- CL is the decision threshold (critical level): above it, the nuclide is
  considered detected. MDA is the minimum detectable activity (about 2 × CL).
- Mean and Std. dev. rows use the detected activities of all samples.

**Nuclear data**: gamma energies, emission probabilities and half-lives with
their uncertainties, and the decay constant λ = LN(2)/T½ as a formula. The
content matches the lab's existing template.

**Profiles**: six depth-profile charts (native Excel scatter charts, depth
downwards, error bars = Unc.):

1. K-40
2. exPb-210 corr
3. Pb-214 and Pb-210
4. Cs-137 corr and Am-241
5. Th-234 and Pb-214
6. Pb-212 and Ac-228

The charts read a formula table below them, which references `Summary` and
`Nuclear data`, so any value changed in Excel flows through.

- `exPb-210 corr = (Pb-210 − Pb-214) × EXP(λ(Pb-210) × Δt)` and
  `Cs-137 corr = Cs-137 × EXP(λ(Cs-137) × Δt)`, where Δt is the time from
  sampling to measurement in years (days / 365.25). The error bar for
  exPb-210 is √(u(Pb-210)² + u(Pb-214)²). Both corrected values stay empty
  (`#N/A`) until a sampling date is entered.
- Below the detection limit, the MDA is plotted without an error bar.
- A sample without a depth isn't plotted (`#N/A`) until a depth is entered.

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
- `core.js`: Summary, Nuclear data and Profiles sheets. SheetJS can't write
  charts, so the chart XML is added to the finished `.xlsx` with JSZip.
- `index.html`: the page.
- `vendor/`: SheetJS 0.18.5 (writes `.xlsx`) and JSZip 3.10.1 (ZIP), stored
  locally so the page has no outside dependencies. SheetJS is only used to
  write files, never to read untrusted spreadsheets.

Run the tests with Node 18+:

```bash
node tests/test-parser.js
```

To try the page locally, open `index.html` in a browser.

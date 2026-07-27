# Envelope Merge for Google Sheets

A bound Google Apps Script that turns a sheet of addresses into a print-ready PDF —
**one envelope per row, one page per record**. Sender (`From`) and recipient (`To`)
blocks render side-by-side. Built for bulk postal runs (e.g. 500 institutions) with
configurable page size, margins, fonts, and a row-range selector.

---

## Features

- One page per row, single merged PDF (not hundreds of files).
- `From` (left) and `To` (right) blocks on the same row via a borderless table.
- Header-based column mapping — column order in the sheet does not matter.
- Configurable page size, margins, font size, italic, column split, row range.
- Settings persist between runs (`PropertiesService`).
- Preview mode (first 3 rows) before committing to the full run.
- Blank rows and empty address fields are skipped automatically.

---

## Requirements

- A Google Sheet with these columns (header names, **case-insensitive**, any order):

  | Column          | Purpose                    |
  |-----------------|----------------------------|
  | `INSTITUTE_NAME`| Recipient / addressee      |
  | `DISTRICT`      | District line              |
  | `THANA`         | Thana / upazila line       |
  | `UNION`         | Union line                 |
  | `POST_OFFICE`   | Post office name           |
  | `POST_CODE`     | Postal code                |

  Header row must be row 1. Extra columns (e.g. `SN`) are ignored.

- A Google account with permission to create Docs/Drive files.

---

## Installation

1. Open your Sheet → **Extensions → Apps Script**.
2. Rename the default file to `Code.gs` and paste in `Code.gs`.
3. Click **+ → HTML**, name it exactly **`Sidebar`** (no `.html` extension),
   and paste in `Sidebar.html`.
4. **Save** (Ctrl+S) and reload the Sheet.
5. A new **Envelopes** menu appears → **Open panel**.
6. On first run, approve the authorization prompt
   (Advanced → *Go to project (unsafe)* → Allow — normal for your own script).

---

## Usage

1. **Envelopes → Open panel**.
2. Set page size, margins, font, and the `From` block text.
3. **Preview first 3** — opens a 3-row Doc so you can check alignment.
4. Adjust and re-preview until correct.
5. Set the row range (or leave **End = 0** for all rows).
6. **Generate PDF** → opens one PDF (one page per record).
7. Print with **Scale = 100% / Actual size** (never "Fit to page"),
   paper and orientation matching your configured page size.

---

## Configuration reference

All settings live in `getSettings()` defaults and are editable at runtime.
Distances are in **millimetres** unless noted.

| Key              | Default | Meaning                                                        |
|------------------|---------|----------------------------------------------------------------|
| `widthMm`        | 279     | Page width (279 = 11 in — Letter landscape).                   |
| `heightMm`       | 216     | Page height (216 = 8.5 in).                                    |
| `marginTopMm`    | 51      | Top margin — vertical position of the block (51 = 2 in).       |
| `marginLeftMm`   | 51      | Left margin — horizontal position of the block (51 = 2 in).    |
| `marginMm`       | 10      | Right / bottom margin.                                          |
| `fromColPct`     | 42      | `From` column width as % of content width; `To` gets the rest. |
| `recipientTopMm` | 0       | Extra spacer below the top margin (0 = none).                  |
| `fromText`       | —       | Sender block; one line per `\n`. Line 2 is bold.               |
| `fontSize`       | 12      | Point size for all address text.                               |
| `italic`         | true    | Italicise all `From` / `To` text.                              |
| `startRow`       | 1       | First data row to process (1-based, excludes header).          |
| `endRow`         | 0       | Last data row; 0 = process to the end.                         |

### Common conversions
- 1 inch = 25.4 mm (so 1 in ≈ 25, 2 in ≈ 51, 3 in ≈ 76).
- 1 mm = 2.8346 points (the `MM` constant).

### Page-size presets (Width × Height, mm)
- Letter landscape: 279 × 216
- Letter portrait: 216 × 279
- A4 landscape: 297 × 210
- #10 envelope: 241 × 105
- DL envelope: 220 × 110
- C5: 229 × 162 · C6: 162 × 114

Width > Height ⇒ landscape (Docs infers orientation from the dimensions).

---

## Changing the address layout

The `To` block is assembled in `_appendEnvelope()`:

```javascript
const toLines = [
  { t: 'To',                            bold: false },
  { t: 'The Head Master',               bold: false },
  { t: institute,                       bold: true  },
  { t: 'District: ' + _clean(r[idx.D]), bold: false },
  { t: 'Thana: '    + _clean(r[idx.T]), bold: false },
  { t: 'Union: '    + _clean(r[idx.U]), bold: false },
  { t: 'Post: '     + post,             bold: false }
];
```

Edit line labels, order, or `bold` flags here. Any line ending in an empty
value (e.g. `Union: ` with no data) is dropped automatically.

To map different sheet columns, edit the `idx` object in `_readRows()`.

---

## Handling large runs (the 6-minute limit)

Apps Script caps a single execution at ~6 minutes. A few hundred rows in one PDF
is usually fine, but if `generateEnvelopes` times out, split the run using the
row-range fields:

- Run 1: Start = 1, End = 250
- Run 2: Start = 251, End = 500

Each run produces its own PDF. No code change needed.

---

## Utility functions

Run these directly from the Apps Script editor (function dropdown → Run):

- `resetSettings()` — clears saved settings so the current code defaults apply.
  **Run this after changing any default in `getSettings()`**, otherwise the old
  saved values persist.
- `previewEnvelopes()` — generates the 3-row preview Doc.
- `generateEnvelopes()` — full run with saved settings.

---

## Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| Changed a default but output is unchanged | Old values are saved in `PropertiesService`. Run `resetSettings()` once, or re-enter values in the sidebar. |
| `From` and `To` stacked vertically | Old paragraph-based layout still in place. Replace the entire `Code.gs` with the current version. |
| `setColumnWidth` / `setWidth` "Unexpected error" | Known intermittent Docs API bug. Width calls are wrapped in try/catch and fall back to auto-sized columns — safe to ignore. |
| `Cannot read properties of null (reading 'setBold')` | `Paragraph.setText()` returns void; never chain after it. Fixed in current version. |
| `Missing column(s): ...` | A required header is misspelled or absent. Match the header names in the Requirements table. |
| Printout shifted / wrong size | Print at **100% / Actual size**, and set paper + orientation to match `widthMm` × `heightMm`. |
| PDFs cluttering Drive root | Add a target folder: `DriveApp.getFolderById('<id>').createFile(pdf)` in `generateEnvelopes`. |

---

## Security & privacy notes

- The script only reads the active sheet and writes Docs/PDFs to **your** Drive.
- No external services or network calls are made.
- Generated Docs/PDFs land in your Drive root by default — move or auto-file them
  if the address data is sensitive.
- Authorization is scoped to Sheets, Docs, and Drive for the signed-in user only.

---

## File overview

- `Code.gs` — menu, settings, sheet reading, Doc/PDF generation.
- `Sidebar.html` — configuration UI (page size, margins, font, `From` text, rows).

---

## License

Provided as-is for internal use. Adapt freely.

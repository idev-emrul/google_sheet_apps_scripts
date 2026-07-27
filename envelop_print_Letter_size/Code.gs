/**
 * Envelope Merge — bound Apps Script for Google Sheets
 * One envelope per row (one page per customer) as a single PDF.
 * From block (left) and To block (right) sit side-by-side on the same row.
 *
 * Expected headers (case-insensitive, any column order):
 *   DISTRICT | THANA | UNION | POST_OFFICE | POST_CODE | INSTITUTE_NAME
 *
 * Setup: Extensions -> Apps Script -> paste this as Code.gs, add Sidebar.html.
 * Reload the Sheet; use the "Envelopes" menu.
 */

const MM = 2.8346;                 // 1 mm in points
const PROP_KEY = 'ENV_SETTINGS';

/* ---------- Menu / UI ---------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Envelopes')
    .addItem('Open panel', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Envelope Merge')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

/* ---------- Settings (persisted) ---------- */

function getSettings() {
  const raw = PropertiesService.getDocumentProperties().getProperty(PROP_KEY);
  const defaults = {
    widthMm: 241, heightMm: 105,       // #10 business envelope
    marginMm: 10,
    fromColPct: 42,                    // left column width % (From). To gets the rest.
    recipientTopMm: 25,                // shared top spacer for both blocks
    fromText: 'From\nPathshala\nwww.pathshala-eims.com\nIT Lab Solutions Ltd\nContact: +88 01842 48 52 22',
    fontSize: 12,
    startRow: 1, endRow: 0             // 0 = all rows
  };
  return raw ? Object.assign(defaults, JSON.parse(raw)) : defaults;
}

function saveSettings(cfg) {
  PropertiesService.getDocumentProperties().setProperty(PROP_KEY, JSON.stringify(cfg));
  return true;
}

/* ---------- Helpers ---------- */

function _readRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const values = sheet.getDataRange().getValues();
  const header = values.shift().map(h => String(h).trim().toUpperCase());
  const idx = {
    D:  header.indexOf('DISTRICT'),
    T:  header.indexOf('THANA'),
    U:  header.indexOf('UNION'),
    PO: header.indexOf('POST_OFFICE'),
    PC: header.indexOf('POST_CODE'),
    IN: header.indexOf('INSTITUTE_NAME')
  };
  const missing = Object.keys(idx).filter(k => idx[k] === -1);
  if (missing.length) {
    throw new Error('Missing column(s): ' + missing.join(', ') +
      '. Headers found: ' + header.join(', '));
  }
  return { rows: values, idx };
}

const _clean = v => String(v == null ? '' : v).trim();

/**
 * Fill one table cell with an array of { t, bold } lines.
 */
function _fillCell(cell, lines, cfg) {
  cell.setPaddingTop(0).setPaddingBottom(0).setPaddingLeft(4).setPaddingRight(4);
  lines.forEach((o, i) => {
    // Cell starts with one empty paragraph — reuse it for the first line.
    const p = (i === 0) ? cell.getChild(0).asParagraph() : cell.appendParagraph('');
    p.setText(o.t)
     .setBold(!!o.bold)
     .setFontSize(cfg.fontSize)
     .setSpacingBefore(0)
     .setSpacingAfter(0);
  });
}

/**
 * Append one envelope page: From (left) | To (right) via a borderless table.
 */
function _appendEnvelope(body, r, idx, cfg, isLast) {
  const institute = _clean(r[idx.IN]);
  const post = [_clean(r[idx.PO]), _clean(r[idx.PC])].filter(Boolean).join(' - ');

  // From block — line index 1 (company/product) is bold
  const fromLines = cfg.fromText.split('\n').map((t, i) => ({ t: t, bold: i === 1 }));

  // To block — drop any field line left empty (e.g. "Union: " with no value)
  const toLines = [
    { t: 'To',                            bold: false },
    { t: 'The Head Master',               bold: false },
    { t: institute,                       bold: true  },
    { t: 'District: ' + _clean(r[idx.D]), bold: false },
    { t: 'Thana: '    + _clean(r[idx.T]), bold: false },
    { t: 'Union: '    + _clean(r[idx.U]), bold: false },
    { t: 'Post: '     + post,             bold: false }
  ].filter(o => o.t !== '' && !/:\s*$/.test(o.t));

  // Shared top spacer to push both blocks down the envelope
  if (cfg.recipientTopMm > 0) {
    body.appendParagraph('')
        .setSpacingBefore(cfg.recipientTopMm * MM)
        .setSpacingAfter(0);
  }

  // 1-row, 2-column borderless table
  const table = body.appendTable([['', '']]);
  table.setBorderWidth(0).setBorderColor('#ffffff');

  const contentPt = (cfg.widthMm - 2 * cfg.marginMm) * MM;
  const leftPct = Math.min(Math.max(cfg.fromColPct, 20), 70) / 100;
  table.setColumnWidth(0, Math.round(contentPt * leftPct));
  table.setColumnWidth(1, Math.round(contentPt * (1 - leftPct)));

  const row = table.getRow(0);
  _fillCell(row.getCell(0), fromLines, cfg);
  _fillCell(row.getCell(1), toLines, cfg);

  if (!isLast) body.appendPageBreak();
}

/* ---------- Build doc ---------- */

function _buildDoc(rowSlice, idx, cfg, name) {
  const doc = DocumentApp.create(name);
  const body = doc.getBody();
  body.setPageWidth(cfg.widthMm * MM).setPageHeight(cfg.heightMm * MM);
  body.setMarginTop(cfg.marginMm * MM).setMarginBottom(cfg.marginMm * MM)
      .setMarginLeft(cfg.marginMm * MM).setMarginRight(cfg.marginMm * MM);

  rowSlice.forEach((r, i) =>
    _appendEnvelope(body, r, idx, cfg, i === rowSlice.length - 1));

  doc.saveAndClose();
  return doc;
}

/* ---------- Preview (first 3 rows) ---------- */

function previewEnvelopes(cfg) {
  cfg = cfg || getSettings();
  const { rows, idx } = _readRows();
  const sample = rows.filter(r => _clean(r[idx.IN])).slice(0, 3);
  if (!sample.length) throw new Error('No non-empty rows found.');

  const doc = _buildDoc(sample, idx, cfg, 'Envelope_PREVIEW');
  return { url: doc.getUrl(), count: sample.length };
}

/* ---------- Generate all (or row range) -> single PDF ---------- */

function generateEnvelopes(cfg) {
  cfg = cfg || getSettings();
  saveSettings(cfg);

  const { rows, idx } = _readRows();

  const start = Math.max(1, cfg.startRow | 0);
  const end   = cfg.endRow > 0 ? cfg.endRow : rows.length;
  const slice = rows.slice(start - 1, end).filter(r => _clean(r[idx.IN]));

  if (!slice.length) throw new Error('No rows in selected range.');

  const stamp = new Date().toISOString().slice(0, 10);
  const doc = _buildDoc(slice, idx, cfg, 'Envelopes_' + stamp + '_' + start + '-' + end);

  const pdf = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  const file = DriveApp.createFile(pdf).setName(doc.getName() + '.pdf');

  return { docUrl: doc.getUrl(), pdfUrl: file.getUrl(), count: slice.length };
}

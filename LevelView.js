/**
 * FieldPulse Pricebook 2.0 — Level View
 *
 * Tiered show/disable across the four authoring sheets. A row's "Level" cell
 * declares the lowest tier it appears in (nested): Basic=1, Normal=2, Everything=3.
 * Selecting tier T enables (Col A=TRUE) + unhides every tagged row with rank <= T,
 * and disables (Col A=FALSE) + hides the rest.
 *
 * Rules:
 *   - Gate column is A (boolean) on every sheet.
 *   - Level column located by HEADER TEXT ("Level"), not a fixed letter.
 *   - BLANK Level -> skipped entirely (Col A and visibility untouched).
 *   - Unrecognized Level value -> skipped and reported.
 *   - Each run reports active (A=TRUE) rows carrying a blank/unknown Level.
 *   - EVERYTHING additionally blanket-unhides every row on each sheet (full visual
 *     reset); Col A is still written on tagged rows only.
 *
 * "Hide All Excluded Rows" is standalone: hides rows where A=FALSE on all four sheets.
 *
 * Depends on: DialogIcons.gs (DIALOG_ICON_CSS, dialogIcon_, dialogSummary_),
 *             Menu.gs (escapeHtml_)
 */

const LEVEL_SHEETS = [
  { name: 'Items',              displayLabel: 'Items' },
  { name: 'Item Option Names',  displayLabel: 'Item Option Names' },
  { name: 'Item Option Values', displayLabel: 'Item Option Values' },
  { name: 'Item Groupings',     displayLabel: 'Item Groupings' }
];

const LEVEL_HEADER_ROW  = 1;
const LEVEL_GATE_COL    = 1;            // Column A on every sheet
const LEVEL_HEADER_TEXT = 'Level';

const LEVEL_RANK = { basic: 1, normal: 2, everything: 3 };

// ---- Menu entry points -----------------------------------------------------

function levelViewBasic()      { applyLevelTier_(1, 'Basic'); }
function levelViewNormal()     { applyLevelTier_(2, 'Normal'); }
function levelViewEverything() { applyLevelTier_(3, 'Everything'); }

// ---- Tier apply ------------------------------------------------------------

function applyLevelTier_(selectedRank, selectedLabel) {
  const ss = SpreadsheetApp.getActive();
  const reports = [];

  LEVEL_SHEETS.forEach(function (def) {
    const report = { label: def.displayLabel, error: null,
                     shown: 0, hidden: 0, skipped: 0,
                     activeBlank: 0, unknown: 0 };
    reports.push(report);

    const sheet = ss.getSheetByName(def.name);
    if (!sheet) { report.error = 'Sheet not found'; return; }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const firstDataRow = LEVEL_HEADER_ROW + 1;
    if (lastRow < firstDataRow || lastCol < 1) return;

    const header = sheet.getRange(LEVEL_HEADER_ROW, 1, 1, lastCol).getValues()[0];
    let levelCol = 0;
    for (let c = 0; c < header.length; c++) {
      if (String(header[c]).trim().toLowerCase() === LEVEL_HEADER_TEXT.toLowerCase()) {
        levelCol = c + 1; break;
      }
    }
    if (!levelCol) { report.error = '"Level" header not found'; return; }

    const numRows   = lastRow - firstDataRow + 1;
    const gateVals  = sheet.getRange(firstDataRow, LEVEL_GATE_COL, numRows, 1).getValues();
    const levelVals = sheet.getRange(firstDataRow, levelCol,       numRows, 1).getValues();

    const newGate  = [];
    const showList = [];
    const hideList = [];

    for (let i = 0; i < numRows; i++) {
      const sheetRow = firstDataRow + i;
      const gate = gateVals[i][0];
      const rawLevel = String(levelVals[i][0] == null ? '' : levelVals[i][0]).trim();

      if (rawLevel === '') {                 // blank Level -> skip, preserve
        newGate.push([gate]);
        report.skipped++;
        if (gate === true) report.activeBlank++;
        continue;
      }
      const rank = LEVEL_RANK[rawLevel.toLowerCase()];
      if (rank === undefined) {              // unrecognized Level -> skip, flag
        newGate.push([gate]);
        report.skipped++;
        report.unknown++;
        if (gate === true) report.activeBlank++;
        continue;
      }
      if (rank <= selectedRank) {
        newGate.push([true]);  showList.push(sheetRow); report.shown++;
      } else {
        newGate.push([false]); hideList.push(sheetRow); report.hidden++;
      }
    }

    sheet.getRange(firstDataRow, LEVEL_GATE_COL, numRows, 1).setValues(newGate);

    if (selectedRank >= 3) {
      // Everything: blanket-unhide the whole sheet (full visual reset).
      sheet.showRows(1, sheet.getMaxRows());
    } else {
      applyRowRuns_(sheet, showList, true);
      applyRowRuns_(sheet, hideList, false);
    }
  });

  showLevelReport_('Level View — ' + selectedLabel, reports, selectedLabel);
}

// ---- Hide Excluded (standalone) --------------------------------------------

function hideDisabledRows() {
  const ss = SpreadsheetApp.getActive();
  const reports = [];

  LEVEL_SHEETS.forEach(function (def) {
    const report = { label: def.displayLabel, error: null, hidden: 0 };
    reports.push(report);

    const sheet = ss.getSheetByName(def.name);
    if (!sheet) { report.error = 'Sheet not found'; return; }

    const lastRow = sheet.getLastRow();
    const firstDataRow = LEVEL_HEADER_ROW + 1;
    if (lastRow < firstDataRow) return;

    const numRows  = lastRow - firstDataRow + 1;
    const gateVals = sheet.getRange(firstDataRow, LEVEL_GATE_COL, numRows, 1).getValues();
    const hideList = [];
    for (let i = 0; i < numRows; i++) {
      if (gateVals[i][0] === false) hideList.push(firstDataRow + i);  // A=FALSE only
    }
    applyRowRuns_(sheet, hideList, false);
    report.hidden = hideList.length;
  });

  showLevelReport_('Hide All Excluded Rows', reports, null);
}

// ---- Helpers ---------------------------------------------------------------

// Collapse a sorted list of 1-based rows into contiguous runs and show/hide.
function applyRowRuns_(sheet, rows, show) {
  if (!rows.length) return;
  rows.sort(function (a, b) { return a - b; });
  let start = rows[0], prev = rows[0];
  for (let i = 1; i <= rows.length; i++) {
    const r = rows[i];
    if (r === prev + 1) { prev = r; continue; }
    const count = prev - start + 1;
    if (show) sheet.showRows(start, count); else sheet.hideRows(start, count);
    start = r; prev = r;
  }
}

// House-style HTML result modal, consistent with all Pricebook Tools dialogs.
function showLevelReport_(title, reports, selectedLabel) {
  const isTier = !!selectedLabel;

  let totalShown = 0, totalHidden = 0, totalSkipped = 0, errorCount = 0;
  const warnings = [];

  reports.forEach(function (r) {
    if (r.error) { errorCount++; return; }
    totalShown   += (r.shown   || 0);
    totalHidden  += (r.hidden  || 0);
    totalSkipped += (r.skipped || 0);
    if (r.activeBlank) warnings.push(r.label + ': ' + r.activeBlank +
        ' active row' + (r.activeBlank === 1 ? '' : 's') + ' with a blank Level');
    if (r.unknown) warnings.push(r.label + ': ' + r.unknown +
        ' row' + (r.unknown === 1 ? '' : 's') + ' with an unrecognized Level value');
  });

  const hasIssues = warnings.length > 0 || errorCount > 0;
  const summaryState = hasIssues ? 'warn' : 'ok';

  let summaryText;
  if (isTier) {
    summaryText = selectedLabel + ' applied — ' + totalShown + ' shown, ' +
      totalHidden + ' hidden' +
      (totalSkipped ? ', ' + totalSkipped + ' untagged/skipped' : '') +
      ' across ' + reports.length + ' sheets.';
  } else {
    summaryText = totalHidden + ' excluded row' + (totalHidden === 1 ? '' : 's') +
      ' hidden across ' + reports.length + ' sheets.';
  }

  const headRow = isTier
    ? '<th>Sheet</th><th class="num">Shown</th><th class="num">Hidden</th><th class="num">Skipped</th>'
    : '<th>Sheet</th><th class="num">Hidden</th>';
  const dataCols = isTier ? 3 : 1;

  let rows = '';
  reports.forEach(function (r) {
    rows += '<tr><td>' + escapeHtml_(r.label) + '</td>';
    if (r.error) {
      rows += '<td class="err-cell" colspan="' + dataCols + '">' + escapeHtml_(r.error) + '</td>';
    } else if (isTier) {
      rows += '<td class="num">' + (r.shown || 0) + '</td>' +
              '<td class="num">' + (r.hidden || 0) + '</td>' +
              '<td class="num">' + (r.skipped || 0) + '</td>';
    } else {
      rows += '<td class="num">' + (r.hidden || 0) + '</td>';
    }
    rows += '</tr>';
  });

  let warnBlock = '';
  if (warnings.length) {
    let items = '';
    warnings.forEach(function (w) { items += '<li>' + escapeHtml_(w) + '</li>'; });
    warnBlock = '<div class="group"><div class="group-header">' + dialogIcon_('warning') +
                'Review</div><ul class="warn-list">' + items + '</ul></div>';
  }

  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    DIALOG_ICON_CSS +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;display:flex;align-items:center;gap:8px;}' +
    '.summary.ok{background:#E8F5E9;border:1px solid #81C784;}' +
    '.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.group{margin-bottom:18px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;display:flex;align-items:center;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}' +
    'th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}' +
    'td{padding:5px 10px;border-bottom:1px solid #ECF0F3;}' +
    'th.num,td.num{text-align:right;width:64px;}' +
    '.err-cell{color:#C62828;font-style:italic;}' +
    '.warn-list{font-size:12px;color:#8A6D3B;margin:6px 0 0;padding-left:18px;line-height:1.6;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-primary{background:#0B5394;color:white;}' +
    '</style></head><body>' +
    dialogSummary_(summaryState, escapeHtml_(summaryText)) +
    '<div class="group"><div class="group-header">' + dialogIcon_(isTier ? 'layers' : 'visibility_off') +
    escapeHtml_(title) + '</div>' +
    '<table><thead><tr>' + headRow + '</tr></thead><tbody>' + rows + '</tbody></table></div>' +
    warnBlock +
    '<div class="actions"><button id="close-btn" class="btn btn-primary">Close</button></div>' +
    '<script>document.getElementById("close-btn").onclick=function(){google.script.host.close();};</script>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(500).setHeight(460), title);
}

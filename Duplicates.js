/**
 * FieldPulse Pricebook 2.0 — Duplicate Detection
 *
 * Real-time onEdit alerts, full-audit menu action, and the shared CF formula
 * builder. The trigger installer and CF rule installer live in Authorize.gs.
 */

const DUPLICATE_SCOPES = [
  { key: 'items', sheetName: 'Items', column: 5, columnLetter: 'E', columnLabel: 'Item', headerRow: 1, mode: 'simple', legacyFormulas: [] },
  { key: 'optionNames', sheetName: 'Item Option Names', column: 2, columnLetter: 'B', columnLabel: 'Option Name', headerRow: 1, mode: 'simple', legacyFormulas: ['=AND($A2<>"",COUNTIF($A$2:$A,$A2)>1)'] },
  { key: 'optionValues', sheetName: 'Item Option Values', column: 6, columnLetter: 'F', columnLabel: 'Option Selection',
    pairColumn: 5, pairColumnLetter: 'E', pairColumnLabel: 'Option Name', headerRow: 1, mode: 'paired', legacyFormulas: [] }
];

const DUP_HIGHLIGHT_BG    = '#FF6666';
const DUP_HIGHLIGHT_FG    = '#FFFFFF';
const DUP_TRIGGER_HANDLER = 'onEditDuplicateCheck';

function onEditDuplicateCheck(e) {
  if (!e || !e.range) return;
  try {
    const range = e.range; const sheet = range.getSheet(); const sheetName = sheet.getName(); const ss = sheet.getParent();
    const applicableScopes = DUPLICATE_SCOPES.filter(function (s) { return s.sheetName === sheetName; });
    if (applicableScopes.length === 0) return;
    const editStartRow = range.getRow(); const editEndRow = range.getLastRow();
    const editStartCol = range.getColumn(); const editEndCol = range.getLastColumn();
    const scopeResults = [];
    applicableScopes.forEach(function (scope) {
      const colsOfInterest = [scope.column];
      if (scope.mode === 'paired') colsOfInterest.push(scope.pairColumn);
      const touched = colsOfInterest.some(function (c) { return c >= editStartCol && c <= editEndCol; });
      if (!touched) return;
      const firstDataRow = scope.headerRow + 1;
      if (editEndRow < firstDataRow) return;
      const allDupes = findDuplicatesInScope_(ss, scope);
      if (allDupes.length === 0) return;
      const relevant = allDupes.filter(function (group) {
        return group.rows.some(function (r) { return r.row >= editStartRow && r.row <= editEndRow; });
      });
      if (relevant.length > 0) scopeResults.push({ scope: scope, groups: relevant });
    });
    if (scopeResults.length === 0) return;
    showDuplicatesModal_(scopeResults, 'realtime');
  } catch (err) { console.error('onEditDuplicateCheck failed: ' + err.message); }
}

function checkAllDuplicates() {
  const ss = SpreadsheetApp.getActive();
  const scopeResults = DUPLICATE_SCOPES.map(function (scope) {
    const sheet = ss.getSheetByName(scope.sheetName);
    if (!sheet) return { scope: scope, groups: [], error: 'Sheet "' + scope.sheetName + '" not found' };
    return { scope: scope, groups: findDuplicatesInScope_(ss, scope) };
  });
  showDuplicatesModal_(scopeResults, 'audit');
}

function buildDuplicateCfFormula_(scope) {
  const c = scope.columnLetter;
  if (scope.mode === 'paired') {
    const p = scope.pairColumnLetter;
    return '=AND($' + c + '2<>"",COUNTIFS($' + p + '$2:$' + p + ',$' + p + '2,$' + c + '$2:$' + c + ',$' + c + '2)>1)';
  }
  return '=AND($' + c + '2<>"",COUNTIF($' + c + '$2:$' + c + ',$' + c + '2)>1)';
}

function findDuplicatesInScope_(ss, scope) {
  const sheet = ss.getSheetByName(scope.sheetName);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  const firstDataRow = scope.headerRow + 1;
  if (lastRow < firstDataRow) return [];
  const numRows = lastRow - firstDataRow + 1;
  const mainValues = sheet.getRange(firstDataRow, scope.column, numRows, 1).getValues();
  let pairValues = null;
  if (scope.mode === 'paired') pairValues = sheet.getRange(firstDataRow, scope.pairColumn, numRows, 1).getValues();
  const groups = {};
  for (let i = 0; i < numRows; i++) {
    const mainRaw = mainValues[i][0];
    const mainNorm = normalizeForDupCheck_(mainRaw);
    if (mainNorm === '') continue;
    let key, pairRaw = null;
    if (scope.mode === 'paired') {
      pairRaw = pairValues[i][0];
      const pairNorm = normalizeForDupCheck_(pairRaw);
      if (pairNorm === '') continue;
      key = pairNorm + '|||' + mainNorm;
    } else key = mainNorm;
    if (!groups[key]) groups[key] = { value: String(mainRaw), pairValue: pairRaw !== null ? String(pairRaw) : null, rows: [] };
    groups[key].rows.push({ row: firstDataRow + i });
  }
  const dupes = [];
  for (const k in groups) {
    if (Object.prototype.hasOwnProperty.call(groups, k) && groups[k].rows.length > 1) dupes.push(groups[k]);
  }
  dupes.sort(function (a, b) { return a.rows[0].row - b.rows[0].row; });
  return dupes;
}

function normalizeForDupCheck_(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim().toLowerCase();
}

function showDuplicatesModal_(scopeResults, context) {
  let totalGroups = 0; let totalRows = 0;
  scopeResults.forEach(function (sr) { if (sr.groups) { totalGroups += sr.groups.length; sr.groups.forEach(function (g) { totalRows += g.rows.length; }); } });
  let body = '';
  scopeResults.forEach(function (sr) {
    body += '<div class="group"><div class="group-header">' + escapeHtml_(sr.scope.sheetName) + ' &middot; Column ' + sr.scope.columnLetter + ' (' + escapeHtml_(sr.scope.columnLabel) + ')</div>';
    if (sr.error) body += '<div class="error">' + escapeHtml_(sr.error) + '</div>';
    else if (!sr.groups || sr.groups.length === 0) body += '<div class="empty">✓ No duplicates</div>';
    else {
      body += '<div class="count">' + sr.groups.length + ' duplicate ' + (sr.groups.length === 1 ? 'group' : 'groups') + '</div>';
      body += '<table><thead><tr>';
      if (sr.scope.mode === 'paired') body += '<th>' + escapeHtml_(sr.scope.pairColumnLabel) + '</th>';
      body += '<th>Value</th><th>Rows</th></tr></thead><tbody>';
      sr.groups.forEach(function (g) {
        body += '<tr>';
        if (sr.scope.mode === 'paired') body += '<td>' + escapeHtml_(g.pairValue) + '</td>';
        body += '<td>' + escapeHtml_(g.value) + '</td>';
        body += '<td>' + g.rows.map(function (r) { return r.row; }).join(', ') + '</td></tr>';
      });
      body += '</tbody></table>';
    }
    body += '</div>';
  });
  const title = context === 'realtime' ? 'Duplicate Value Detected' : 'Duplicate Value Audit';
  let summaryClass, summaryText;
  if (totalGroups === 0) { summaryClass = 'ok'; summaryText = '✓ No duplicates found.'; }
  else if (context === 'realtime') { summaryClass = 'warn'; summaryText = '⚠ Your recent edit produced ' + totalGroups + ' duplicate ' + (totalGroups === 1 ? 'group' : 'groups') + ' (' + totalRows + ' rows).'; }
  else { summaryClass = 'warn'; summaryText = '⚠ ' + totalGroups + ' duplicate ' + (totalGroups === 1 ? 'group' : 'groups') + ' across ' + totalRows + ' rows.'; }
  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}.summary.ok{background:#E8F5E9;border:1px solid #81C784;}.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.group{margin-bottom:18px;}.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.count{color:#C62828;font-size:11px;margin:6px 0;}.empty{color:#2E7D32;font-size:12px;padding:4px 0;}.error{color:#C62828;font-size:12px;font-style:italic;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}td{padding:5px 10px;border-bottom:1px solid #ECF0F3;}' +
    '</style></head><body><div class="summary ' + summaryClass + '">' + summaryText + '</div>' + body + '</body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(580), title);
}

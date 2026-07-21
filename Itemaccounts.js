/**
 * FieldPulse Pricebook 2.0 — Check Item Accounts
 *
 * On the Items sheet, every INCLUDED item (col A = TRUE) must have both
 * account columns filled:
 *   Z  (26) — Income Account
 *   AA (27) — Purchase Account
 *
 * Values come from dropdowns, so validity is already enforced — this only
 * checks that both cells are non-blank. Excluded rows (A ≠ TRUE) and blank
 * rows are ignored. Read-only.
 *
 * Shared by the standalone menu item AND the Pre-Export Health Check
 * (auditItemAccounts_ is called from both).
 */

const ITEM_ACCT_SHEET       = 'Items';
const ITEM_ACCT_FIRST_ROW   = 3;    // header row 1, hidden formula row 2
const ITEM_ACCT_INCLUDE_COL = 1;    // A — include flag
const ITEM_ACCT_NAME_COL    = 5;    // E — Item
const ITEM_ACCT_INCOME_COL  = 26;   // Z — Income Account
const ITEM_ACCT_PURCH_COL   = 27;   // AA — Purchase Account
const ITEM_ACCT_READ_WIDTH  = 27;   // read A..AA in one shot
const ITEM_ACCT_MAX_LISTED  = 50;

function checkItemAccounts() {
  const ss = SpreadsheetApp.getActive();
  let audit;
  try {
    audit = auditItemAccounts_(ss);
  } catch (err) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Check Item Accounts — Error', err.message, ui.ButtonSet.OK);
    return;
  }
  showItemAccountsModal_(audit);
}

function auditItemAccounts_(ss) {
  const report = { error: null, checked: 0, ok: 0, bad: [] };

  const sheet = ss.getSheetByName(ITEM_ACCT_SHEET);
  if (!sheet) { report.error = 'Sheet "' + ITEM_ACCT_SHEET + '" not found'; return report; }

  const lastRow = sheet.getLastRow();
  if (lastRow < ITEM_ACCT_FIRST_ROW) return report;

  const numRows = lastRow - ITEM_ACCT_FIRST_ROW + 1;
  const data = sheet.getRange(ITEM_ACCT_FIRST_ROW, 1, numRows, ITEM_ACCT_READ_WIDTH).getValues();

  for (var i = 0; i < numRows; i++) {
    var row = data[i];
    var flag = row[ITEM_ACCT_INCLUDE_COL - 1];
    var included = (flag === true) || (String(flag).trim().toUpperCase() === 'TRUE');
    if (!included) continue;   // only included rows

    report.checked++;

    var name     = String(row[ITEM_ACCT_NAME_COL   - 1] == null ? '' : row[ITEM_ACCT_NAME_COL   - 1]).trim();
    var income   = String(row[ITEM_ACCT_INCOME_COL - 1] == null ? '' : row[ITEM_ACCT_INCOME_COL - 1]).trim();
    var purchase = String(row[ITEM_ACCT_PURCH_COL  - 1] == null ? '' : row[ITEM_ACCT_PURCH_COL  - 1]).trim();

    var missing = [];
    if (income === '')   missing.push('Income');
    if (purchase === '') missing.push('Purchase');

    if (missing.length) {
      report.bad.push({ row: ITEM_ACCT_FIRST_ROW + i, name: name || '(unnamed)', missing: missing.join(' + ') });
    } else {
      report.ok++;
    }
  }

  return report;
}

function showItemAccountsModal_(audit) {
  let summaryClass, summaryText;
  if (audit.error) {
    summaryClass = 'warn'; summaryText = '⚠ ' + audit.error;
  } else if (audit.bad.length === 0) {
    summaryClass = 'ok';
    summaryText = '✓ All ' + audit.checked + ' included item' + (audit.checked === 1 ? '' : 's') +
                  ' have an Income and Purchase account.';
  } else {
    summaryClass = 'warn';
    summaryText = '⚠ ' + audit.bad.length + ' included item' + (audit.bad.length === 1 ? '' : 's') +
                  ' missing an account.';
  }

  let body = '';
  if (!audit.error) {
    body += '<div class="group"><div class="group-header">Items <span class="col">(Z Income · AA Purchase)</span></div>';
    if (audit.bad.length === 0) {
      body += '<div class="empty">✓ All ' + audit.checked + ' included item' + (audit.checked === 1 ? '' : 's') + ' OK</div>';
    } else {
      body += '<div class="count">' + audit.bad.length + ' of ' + audit.checked + ' included item' +
              (audit.bad.length === 1 ? '' : 's') + ' need an account</div>';
      body += '<table><thead><tr><th>Row</th><th>Item</th><th>Missing</th></tr></thead><tbody>';
      audit.bad.slice(0, ITEM_ACCT_MAX_LISTED).forEach(function (b) {
        body += '<tr><td>' + b.row + '</td><td>' + escapeHtml_(b.name) + '</td><td>' + escapeHtml_(b.missing) + '</td></tr>';
      });
      body += '</tbody></table>';
      if (audit.bad.length > ITEM_ACCT_MAX_LISTED) {
        body += '<div class="more">…and ' + (audit.bad.length - ITEM_ACCT_MAX_LISTED) + ' more.</div>';
      }
    }
    body += '</div>';
  }

  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}' +
    '.summary.ok{background:#E8F5E9;border:1px solid #81C784;}' +
    '.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.group{margin-bottom:18px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.group-header .col{font-weight:400;color:#6B7C8C;font-size:11px;}' +
    '.count{color:#C62828;font-size:11px;margin:6px 0;}' +
    '.empty{color:#2E7D32;font-size:12px;padding:4px 0;}' +
    '.more{font-size:11px;color:#6B7C8C;font-style:italic;padding:6px 0;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}' +
    'th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}' +
    'td{padding:5px 10px;border-bottom:1px solid #ECF0F3;vertical-align:top;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-primary{background:#0B5394;color:white;}' +
    '</style></head><body>' +
    '<div class="summary ' + summaryClass + '">' + summaryText + '</div>' + body +
    '<div class="actions"><button id="close-btn" class="btn btn-primary">Close</button></div>' +
    '<script>document.getElementById("close-btn").onclick=function(){google.script.host.close();};</script>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(540), 'Check Item Accounts');
}

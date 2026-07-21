/**
 * FieldPulse Pricebook 2.0 — Check Defined Levels
 *
 * Verifies every REAL row on the four authoring sheets carries a valid
 * Pricebook Level tag: Basic, Normal, or Everything.
 *
 * A row is "real" when its name/identifier column is non-empty. Blank rows
 * (and the hidden formula row 2) are skipped entirely — the include flag in
 * column A is NOT used, because it can be set on otherwise-empty rows.
 *
 *   Sheet                Name column            Level column
 *   Items                E  (Item)              AZ
 *   Item Option Names    B  (Option Name)       M
 *   Item Option Values   F  (Option Selection)  R
 *   Item Groupings       F  (Item Name)         AT
 *
 * Two problems are reported per row:
 *   • Missing  — the Level cell is blank
 *   • Invalid  — the Level cell holds something other than the three values
 *
 * Read-only — never writes.
 */

const LEVEL_VALID = ['Basic', 'Normal', 'Everything'];
const LEVEL_FIRST_DATA_ROW = 3;   // header row 1, hidden formula row 2
const LEVEL_MAX_LISTED = 50;      // cap rows listed per sheet in the modal

const LEVEL_SPECS = [
  { name: 'Items',              keyCol: 5, keyLabel: 'Item',             levelCol: 52, levelLetter: 'AZ' },
  { name: 'Item Option Names',  keyCol: 2, keyLabel: 'Option Name',      levelCol: 13, levelLetter: 'M'  },
  { name: 'Item Option Values', keyCol: 6, keyLabel: 'Option Selection', levelCol: 18, levelLetter: 'R'  },
  { name: 'Item Groupings',     keyCol: 6, keyLabel: 'Item Name',        levelCol: 46, levelLetter: 'AT' }
];

function checkDefinedLevels() {
  const ss = SpreadsheetApp.getActive();
  let audit;
  try {
    audit = auditLevels_(ss);
  } catch (err) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Check Defined Levels — Error', err.message, ui.ButtonSet.OK);
    return;
  }
  showLevelsModal_(audit);
}

function auditLevels_(ss) {
  const result = { sheets: [], totalBad: 0, totalChecked: 0 };

  LEVEL_SPECS.forEach(function (spec) {
    const report = { name: spec.name, letter: spec.levelLetter, error: null, checked: 0, ok: 0, bad: [] };
    result.sheets.push(report);

    const sheet = ss.getSheetByName(spec.name);
    if (!sheet) { report.error = 'Sheet "' + spec.name + '" not found'; return; }

    const lastRow = sheet.getLastRow();
    if (lastRow < LEVEL_FIRST_DATA_ROW) return;

    const numRows = lastRow - LEVEL_FIRST_DATA_ROW + 1;
    const keys   = sheet.getRange(LEVEL_FIRST_DATA_ROW, spec.keyCol,   numRows, 1).getDisplayValues();
    const levels = sheet.getRange(LEVEL_FIRST_DATA_ROW, spec.levelCol, numRows, 1).getDisplayValues();

    for (var i = 0; i < numRows; i++) {
      var key = String(keys[i][0] == null ? '' : keys[i][0]).trim();
      if (key === '') continue;   // blank row — not a real entry, skip

      report.checked++;
      result.totalChecked++;

      var lvl = String(levels[i][0] == null ? '' : levels[i][0]).trim();
      if (lvl === '') {
        report.bad.push({ row: LEVEL_FIRST_DATA_ROW + i, name: key, reason: 'Missing', value: '' });
        result.totalBad++;
      } else if (LEVEL_VALID.indexOf(lvl) === -1) {
        report.bad.push({ row: LEVEL_FIRST_DATA_ROW + i, name: key, reason: 'Invalid', value: lvl });
        result.totalBad++;
      } else {
        report.ok++;
      }
    }
  });

  return result;
}

function showLevelsModal_(audit) {
  let summaryClass, summaryText;
  if (audit.totalBad === 0) {
    summaryClass = 'ok';
    summaryText = '✓ All ' + audit.totalChecked + ' defined row' + (audit.totalChecked === 1 ? '' : 's') +
                  ' have a valid Level.';
  } else {
    summaryClass = 'warn';
    summaryText = '⚠ ' + audit.totalBad + ' row' + (audit.totalBad === 1 ? '' : 's') +
                  ' missing or invalid Level.';
  }

  let body = '';
  audit.sheets.forEach(function (s) {
    body += '<div class="group"><div class="group-header">' + escapeHtml_(s.name) +
            ' <span class="col">(col ' + escapeHtml_(s.letter) + ')</span></div>';
    if (s.error) {
      body += '<div class="error">' + escapeHtml_(s.error) + '</div>';
    } else if (s.bad.length === 0) {
      body += '<div class="empty">✓ All ' + s.checked + ' defined row' + (s.checked === 1 ? '' : 's') + ' OK</div>';
    } else {
      body += '<div class="count">' + s.bad.length + ' of ' + s.checked + ' row' +
              (s.bad.length === 1 ? '' : 's') + ' need a Level</div>';
      body += '<table><thead><tr><th>Row</th><th>Name</th><th>Problem</th><th>Value</th></tr></thead><tbody>';
      s.bad.slice(0, LEVEL_MAX_LISTED).forEach(function (b) {
        var val = b.value === '' ? '<span class="muted">(blank)</span>' : escapeHtml_(b.value);
        body += '<tr><td>' + b.row + '</td><td>' + escapeHtml_(b.name) + '</td><td>' +
                escapeHtml_(b.reason) + '</td><td>' + val + '</td></tr>';
      });
      body += '</tbody></table>';
      if (s.bad.length > LEVEL_MAX_LISTED) {
        body += '<div class="more">…and ' + (s.bad.length - LEVEL_MAX_LISTED) + ' more.</div>';
      }
    }
    body += '</div>';
  });

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
    '.error{color:#C62828;font-size:12px;font-style:italic;}' +
    '.more{font-size:11px;color:#6B7C8C;font-style:italic;padding:6px 0;}' +
    '.muted{color:#6B7C8C;font-style:italic;}' +
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

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(560).setHeight(560), 'Check Defined Levels');
}

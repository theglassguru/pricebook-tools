/**
 * FieldPulse Pricebook 2.0 — Bulk Include / Exclude
 *
 * Sets column A on the active sheet to TRUE (Include All) or FALSE
 * (Exclude All). Only rows where the sheet's gate column is populated
 * are touched — blank rows are left alone so the sheet stays clean.
 *
 * Gate columns (presence marks the row as "real"):
 *   Items              → column E  (Item name)
 *   Item Option Names  → column B  (Option Name)
 *   Item Option Values → column E  (Option Name)
 *   Item Groupings     → column B  (Job Name)
 */

const BULK_TOGGLE_SPECS = {
  'Items':              { headerRow: 1, gateCol: 5 },
  'Item Option Names':  { headerRow: 1, gateCol: 2 },
  'Item Option Values': { headerRow: 1, gateCol: 5 },
  'Item Groupings':     { headerRow: 1, gateCol: 2 }
};

function includeAllOnActiveSheet() { bulkToggleActiveSheet_(true);  }
function excludeAllOnActiveSheet() { bulkToggleActiveSheet_(false); }

function bulkToggleActiveSheet_(value) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const sheetName = sheet.getName();
  const spec = BULK_TOGGLE_SPECS[sheetName];
  const label = value ? 'Include All' : 'Exclude All';

  if (!spec) {
    ui.alert(
      label,
      'This command works only on the Items, Item Option Names, Item Option Values, or Item Groupings sheets.\n\nYou are currently on "' + sheetName + '".',
      ui.ButtonSet.OK
    );
    return;
  }

  const lastRow = sheet.getLastRow();
  const firstDataRow = spec.headerRow + 1;
  if (lastRow < firstDataRow) {
    ui.alert(label, 'No data rows on "' + sheetName + '".', ui.ButtonSet.OK);
    return;
  }

  const numRows = lastRow - firstDataRow + 1;
  const gateValues = sheet.getRange(firstDataRow, spec.gateCol, numRows, 1).getValues();
  const currentA   = sheet.getRange(firstDataRow, 1, numRows, 1).getValues();

  let changed = 0;
  const newA = currentA.map(function (row, i) {
    const gate = gateValues[i][0];
    const hasGate = gate !== null && gate !== undefined && String(gate).trim() !== '';
    if (hasGate) {
      if (row[0] !== value) changed++;
      return [value];
    }
    return [row[0]];
  });

  sheet.getRange(firstDataRow, 1, numRows, 1).setValues(newA);
  SpreadsheetApp.flush();

  ui.alert(
    label,
    'Set ' + changed + ' row' + (changed === 1 ? '' : 's') + ' on "' + sheetName + '" to ' + (value ? 'TRUE' : 'FALSE') + '.',
    ui.ButtonSet.OK
  );
}

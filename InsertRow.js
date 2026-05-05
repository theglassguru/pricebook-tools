/**
 * FieldPulse Pricebook 2.0 — Insert Row Below Active
 *
 * Adds a new row directly below the active row on Items, Item Option
 * Values, or Item Groupings. Inherits category-defining columns from
 * the active row; formula columns are left untouched so each sheet's
 * ARRAYFORMULAs spill cleanly into the new row.
 *
 * Inherit specs (per Sean):
 *   Items              → A, B, C, D, F, G          → cursor parks on E
 *   Item Option Values → A, B, C, D, E, G          → cursor parks on F
 *   Item Groupings     → A, B, C, N–W              → cursor parks on F
 *
 * The active sheet is detected automatically. If the user is not on one
 * of the three supported sheets, the function alerts and exits.
 */

const INSERT_ROW_SPECS = {
  'Items': {
    headerRow:   1,
    inheritCols: [1, 2, 3, 4, 6, 7],     // A, B, C, D, F, G
    cursorCol:   5                        // E (Item)
  },
  'Item Option Values': {
    headerRow:   1,
    inheritCols: [1, 2, 3, 4, 5, 7],     // A, B, C, D, E, G
    cursorCol:   6                        // F (Option Selection)
  },
  'Item Groupings': {
    headerRow:   1,
    inheritCols: [1, 2, 3, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23],  // A, B, C, N–W
    cursorCol:   6                                                     // F (Item Name)
  }
};

function insertRowBelowActive() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const sheetName = sheet.getName();

  const spec = INSERT_ROW_SPECS[sheetName];
  if (!spec) {
    ui.alert(
      'Insert Row Below',
      'This command works only on the Items, Item Option Values, or Item Groupings sheets.\n\nYou are currently on "' + sheetName + '".',
      ui.ButtonSet.OK
    );
    return;
  }

  const activeRange = sheet.getActiveRange();
  if (!activeRange) {
    ui.alert('Insert Row Below', 'Click on a row first, then run this command.', ui.ButtonSet.OK);
    return;
  }

  const activeRow = activeRange.getRow();
  if (activeRow <= spec.headerRow) {
    ui.alert(
      'Insert Row Below',
      'Please click on a data row (row ' + (spec.headerRow + 1) + ' or below) first.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Snapshot the source row's values across all populated columns.
  const lastCol = sheet.getLastColumn();
  const sourceRow = sheet.getRange(activeRow, 1, 1, lastCol).getValues()[0];

  // Insert one blank row directly below the active row.
  sheet.insertRowsAfter(activeRow, 1);
  const newRow = activeRow + 1;
  SpreadsheetApp.flush();

  // Write inherited values, one column at a time. Writing to the formula
  // columns would block the ARRAYFORMULA spill at row 2 (#REF!), so we
  // touch only the columns in the inherit list.
  spec.inheritCols.forEach(function (colIdx) {
    if (colIdx > sourceRow.length) return;
    const value = sourceRow[colIdx - 1];
    sheet.getRange(newRow, colIdx).setValue(value);
  });

  // Park cursor on the first user-editable blank cell in the new row.
  sheet.setActiveRange(sheet.getRange(newRow, spec.cursorCol));
}

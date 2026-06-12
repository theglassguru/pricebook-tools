/**
 * Menu-driven entry — handles UI feedback.
 */
function updateNamedRangesSheet() {
  const result = updateNamedRangesSheet_();
  if (!result.ok) {
    SpreadsheetApp.getUi().alert(result.error);
  }
}

/**
 * Inner worker — returns {ok, note, error} for chaining into other flows
 * (e.g. Authorize & Activate). No UI side effects.
 */
function updateNamedRangesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Named Ranges');
  if (!sheet) {
    return { ok: false, error: 'Sheet "Named Ranges" not found' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 4).clearContent();
  }

  const namedRanges = ss.getNamedRanges()
    .slice()
    .sort((a, b) => a.getName().localeCompare(b.getName()));

  const rows = [];
  namedRanges.forEach(nr => {
    const r = nr.getRange();
    const sh = r.getSheet();
    const startCol = r.getColumn();
    const startRow = r.getRow();
    const numCols = r.getNumColumns();
    const numRows = r.getNumRows();
    const isSingleCell = numCols === 1 && numRows === 1;
    const startsOnRow2 = startRow === 2;

    if (isSingleCell || !startsOnRow2) {
      rows.push([nr.getName(), r.getA1Notation(), sh.getName(), '']);
      return;
    }

    const headers = sh.getRange(1, startCol, 1, numCols).getValues()[0];
    headers.forEach((h, i) => {
      const colLetter = sh.getRange(1, startCol + i).getA1Notation().replace(/\d+/, '');
      const label = `${i + 1} | ${colLetter} | ${h}`;
      if (i === 0) {
        rows.push([nr.getName(), r.getA1Notation(), sh.getName(), label]);
      } else {
        rows.push(['', '', '', label]);
      }
    });
  });

  if (rows.length) sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  return { ok: true, note: namedRanges.length + ' named range' + (namedRanges.length === 1 ? '' : 's') + ' refreshed' };
}

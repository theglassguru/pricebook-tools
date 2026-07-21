/**
 * FieldPulse Pricebook 2.0 — Link Option Names to Values
 *
 * On Item Option Names, turns each Option Name (col B) into a rich-text link
 * that jumps to that option's block of selections on Item Option Values
 * (col F). Mirrors the manual links Sean set on Glass Type / Thickness:
 *     #gid=<Item Option Values sheetId>&range=F<first>:F<last>
 *
 * A block is the contiguous run of rows on Item Option Values where the
 * Option Name (col E) equals the name. Audit-first: shows what it will link
 * and flags anything it cannot (name has no values, or its rows are not
 * contiguous — which would otherwise link across a neighbouring option).
 *
 * Rich-text links keep col B as plain text (not a formula), so lookups,
 * dropdowns and exports that read the Option Name are unaffected.
 */

const OPT_NAMES_SHEET   = 'Item Option Names';
const OPT_VALUES_SHEET  = 'Item Option Values';
const OPT_NAMES_COL     = 2;   // B — Option Name
const OPT_VALUES_KEYCOL = 5;   // E — Option Name (grouping key on the values sheet)
const OPT_VALUES_LINKCOL = 'F'; // F — Option Selection (what the link targets)
const OPT_FIRST_DATA_ROW = 3;  // header row 1, hidden formula row 2

function linkOptionNamesToValues() {
  const ss = SpreadsheetApp.getActive();
  let audit;
  try {
    audit = auditLinkOptions_(ss);
  } catch (err) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Link Option Names — Error', err.message, ui.ButtonSet.OK);
    return;
  }
  showLinkOptionsModal_(audit);
}

function auditLinkOptions_(ss) {
  const namesSheet  = ss.getSheetByName(OPT_NAMES_SHEET);
  const valuesSheet = ss.getSheetByName(OPT_VALUES_SHEET);
  if (!namesSheet)  throw new Error('Sheet "' + OPT_NAMES_SHEET + '" not found.');
  if (!valuesSheet) throw new Error('Sheet "' + OPT_VALUES_SHEET + '" not found.');

  const valuesGid = valuesSheet.getSheetId();

  // Read the Option Name key column (E) once; build first/last row per name.
  var keys = [];
  const vLast = valuesSheet.getLastRow();
  if (vLast >= OPT_FIRST_DATA_ROW) {
    keys = valuesSheet.getRange(OPT_FIRST_DATA_ROW, OPT_VALUES_KEYCOL, vLast - OPT_FIRST_DATA_ROW + 1, 1).getDisplayValues();
  }
  const span = {};   // name -> { min, max, count }
  for (var i = 0; i < keys.length; i++) {
    var nm = String(keys[i][0] == null ? '' : keys[i][0]).trim();
    if (!nm) continue;
    var row = OPT_FIRST_DATA_ROW + i;
    if (!span[nm]) span[nm] = { min: row, max: row, count: 1 };
    else { span[nm].max = row; span[nm].count++; }
  }

  // Walk the Option Names and plan a link per non-blank name cell.
  const audit = { valuesGid: valuesGid, toLink: [], skipped: [] };
  const nLast = namesSheet.getLastRow();
  if (nLast >= OPT_FIRST_DATA_ROW) {
    const names = namesSheet.getRange(OPT_FIRST_DATA_ROW, OPT_NAMES_COL, nLast - OPT_FIRST_DATA_ROW + 1, 1).getDisplayValues();
    for (var j = 0; j < names.length; j++) {
      var name = String(names[j][0] == null ? '' : names[j][0]).trim();
      if (!name) continue;
      var cellRow = OPT_FIRST_DATA_ROW + j;
      var b = span[name];
      if (!b) {
        audit.skipped.push({ row: cellRow, name: name, reason: 'No values found on ' + OPT_VALUES_SHEET });
        continue;
      }
      // Interleave check within [min, max]: blanks are OK, a DIFFERENT
      // option name inside the span means the values are scattered and can't
      // be one link.
      var interleaved = false;
      for (var k = b.min - OPT_FIRST_DATA_ROW; k <= b.max - OPT_FIRST_DATA_ROW; k++) {
        var v = String(keys[k][0] == null ? '' : keys[k][0]).trim();
        if (v !== '' && v !== name) { interleaved = true; break; }
      }
      if (interleaved) {
        audit.skipped.push({
          row: cellRow, name: name,
          reason: 'Scattered — other options fall between rows ' + b.min + '–' + b.max + '; regroup these values together, then re-run'
        });
      } else {
        var url = '#gid=' + valuesGid + '&range=' + OPT_VALUES_LINKCOL + b.min + ':' + OPT_VALUES_LINKCOL + b.max;
        audit.toLink.push({ row: cellRow, name: name, min: b.min, max: b.max, url: url });
      }
    }
  }
  return audit;
}

function applyOptionNameLinks() {
  const ss = SpreadsheetApp.getActive();
  const audit = auditLinkOptions_(ss);
  const namesSheet = ss.getSheetByName(OPT_NAMES_SHEET);
  const result = { linked: 0, errors: [] };

  audit.toLink.forEach(function (item) {
    try {
      const rich = SpreadsheetApp.newRichTextValue().setText(item.name).setLinkUrl(item.url).build();
      namesSheet.getRange(item.row, OPT_NAMES_COL).setRichTextValue(rich);
      result.linked++;
    } catch (err) {
      result.errors.push(item.name + ' (' + err.message + ')');
    }
  });
  SpreadsheetApp.flush();
  return result;
}

function showLinkOptionsModal_(audit) {
  const linkCount = audit.toLink.length;
  const skipCount = audit.skipped.length;

  let summaryClass, summaryText;
  if (linkCount === 0) {
    summaryClass = 'warn';
    summaryText = '⚠ No option names could be matched to a value block.';
  } else if (skipCount === 0) {
    summaryClass = 'info';
    summaryText = 'Ready to link ' + linkCount + ' option name' + (linkCount === 1 ? '' : 's') + '.';
  } else {
    summaryClass = 'warn';
    summaryText = 'Ready to link ' + linkCount + '; ' + skipCount + ' skipped (see below).';
  }

  let body = '<div class="group"><div class="group-header">Will link (' + linkCount + ')</div>';
  if (linkCount === 0) body += '<div class="empty">Nothing to link.</div>';
  else body += '<div class="ok-note">' + linkCount + ' option name' + (linkCount === 1 ? '' : 's') +
               ' matched to a contiguous value block and will be linked to Item Option Values.</div>';
  body += '</div>';

  if (skipCount > 0) {
    body += '<div class="group"><div class="group-header">Skipped (' + skipCount + ')</div>';
    body += '<table><thead><tr><th>Row</th><th>Option Name</th><th>Reason</th></tr></thead><tbody>';
    audit.skipped.forEach(function (s) {
      body += '<tr><td>' + s.row + '</td><td>' + escapeHtml_(s.name) + '</td><td>' + escapeHtml_(s.reason) + '</td></tr>';
    });
    body += '</tbody></table></div>';
  }

  const note = '<div class="note">Applies a rich-text link to each Option Name cell (col B). Re-run any time values are added or reordered. This overwrites existing links in those cells.</div>';
  const applyBtn = linkCount > 0 ? '<button id="apply-btn" class="btn btn-primary">Link ' + linkCount + '</button>' : '';

  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}' +
    '.summary.info{background:#E3F2FD;border:1px solid #64B5F6;}' +
    '.summary.ok{background:#E8F5E9;border:1px solid #81C784;}' +
    '.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.summary.err{background:#FDECEA;border:1px solid #E57373;}' +
    '.group{margin-bottom:16px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.ok-note{color:#2E7D32;font-size:12px;padding:2px 0;}.empty{color:#6B7C8C;font-size:12px;padding:2px 0;}' +
    '.note{font-size:11px;color:#6B7C8C;font-style:italic;padding:8px 0;margin-bottom:8px;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}' +
    'th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}' +
    'td{padding:5px 10px;border-bottom:1px solid #ECF0F3;vertical-align:top;}' +
    '.detail{color:#2E7D32;font-size:12px;padding:4px 0;white-space:pre-wrap;}.error{color:#C62828;font-size:12px;padding:4px 0;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-secondary{background:white;border-color:#D6DFE8;color:#1A2733;}.btn-primary{background:#0B5394;color:white;}' +
    '</style></head><body>' +
    '<div id="audit-view"><div class="summary ' + summaryClass + '">' + summaryText + '</div>' + body + note + '</div>' +
    '<div id="working-view" style="display:none;"><div class="summary info">Linking option names…</div></div>' +
    '<div id="done-view" style="display:none;"><div class="summary ok" id="done-summary"></div><div class="error" id="done-errors"></div></div>' +
    '<div id="error-view" style="display:none;"><div class="summary err">⚠ An error occurred.</div><div class="error" id="error-message"></div></div>' +
    '<div class="actions">' +
    '<button id="cancel-btn" class="btn btn-secondary">Cancel</button>' + applyBtn +
    '<button id="close-btn" class="btn btn-primary" style="display:none;">Close</button></div>' +
    '<script>(function(){' +
    'var cancel=document.getElementById("cancel-btn");var apply=document.getElementById("apply-btn");var close=document.getElementById("close-btn");' +
    'if(cancel)cancel.onclick=function(){google.script.host.close();};' +
    'if(close)close.onclick=function(){google.script.host.close();};' +
    'if(apply)apply.onclick=function(){' +
    'document.getElementById("audit-view").style.display="none";document.getElementById("working-view").style.display="block";' +
    'cancel.style.display="none";apply.style.display="none";' +
    'google.script.run.withSuccessHandler(onOk).withFailureHandler(onErr).applyOptionNameLinks();};' +
    'function onOk(r){document.getElementById("working-view").style.display="none";document.getElementById("done-view").style.display="block";' +
    'document.getElementById("done-summary").innerHTML="&#10003; Linked "+r.linked+" option name"+(r.linked===1?"":"s")+".";' +
    'document.getElementById("done-errors").textContent=r.errors.length?("Errors: "+r.errors.join("; ")):"";' +
    'close.style.display="inline-block";}' +
    'function onErr(e){document.getElementById("working-view").style.display="none";document.getElementById("error-view").style.display="block";' +
    'document.getElementById("error-message").textContent=(e&&e.message)?e.message:String(e);close.style.display="inline-block";}' +
    '})();</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(560), 'Link Option Names to Values');
}

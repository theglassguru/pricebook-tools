/**
 * FieldPulse Pricebook 2.0 — XLSX Export
 */

const XLSX_EXPORT_HEADER_ROWS = 2;

const XLSX_EXPORT_FILES = [
  { filename: 'dimensional_item_template.xlsx',
    sheets: [{ sourceName: 'dimensional_items', destName: 'Dimensional Item Import' }] },
  { filename: 'dimensional_item_option_template.xlsx',
    sheets: [
      { sourceName: 'dimensional_options', destName: 'Dimensional Options Import' },
      { sourceName: 'dim_options_values',  destName: 'Dimensional Option Values Imp' }
    ] },
  { filename: 'Pricebook_import_single_tab_with_categories.xlsx',
    sheets: [{ sourceName: 'pricebook_import_standard', destName: 'Standard' }] }
];

function exportPricebookFiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const plan = XLSX_EXPORT_FILES.map(function (file) {
    return { filename: file.filename, sheets: file.sheets.map(function (s) { return analyzeSourceSheet_(ss, s); }) };
  });
  showXlsxExportModal_(plan);
}

function analyzeSourceSheet_(ss, sheetSpec) {
  const sheet = ss.getSheetByName(sheetSpec.sourceName);
  if (!sheet) return { sourceName: sheetSpec.sourceName, destName: sheetSpec.destName, dataRows: 0, error: 'Source sheet not found' };
  const lastRow = sheet.getLastRow();
  return { sourceName: sheetSpec.sourceName, destName: sheetSpec.destName, dataRows: Math.max(0, lastRow - XLSX_EXPORT_HEADER_ROWS) };
}

function buildXlsxForIndex(index) {
  if (typeof index !== 'number' || index < 0 || index >= XLSX_EXPORT_FILES.length) throw new Error('Invalid file index: ' + index);
  return buildSingleXlsx_(SpreadsheetApp.getActiveSpreadsheet(), XLSX_EXPORT_FILES[index]);
}

function buildSingleXlsx_(ss, fileSpec) {
  const stamp = Utilities.getUuid().replace(/-/g, '').substring(0, 12);
  const tempName = '__pricebook_export_tmp_' + stamp + '_' + fileSpec.filename;
  const tempSs = SpreadsheetApp.create(tempName);
  const tempId = tempSs.getId();
  try {
    const defaultSheet = tempSs.getSheets()[0];
    for (let i = 0; i < fileSpec.sheets.length; i++) {
      const sheetSpec = fileSpec.sheets[i];
      const sourceSheet = ss.getSheetByName(sheetSpec.sourceName);
      if (!sourceSheet) throw new Error('Source sheet "' + sheetSpec.sourceName + '" not found.');
      let destSheet;
      if (i === 0) { defaultSheet.setName(sheetSpec.destName); destSheet = defaultSheet; }
      else destSheet = tempSs.insertSheet(sheetSpec.destName);
      const lastRow = sourceSheet.getLastRow();
      const lastCol = sourceSheet.getLastColumn();
      if (lastRow > 0 && lastCol > 0) {
        const values = sourceSheet.getRange(1, 1, lastRow, lastCol).getValues();
        destSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
        const maxRows = destSheet.getMaxRows();
        if (maxRows > values.length) destSheet.deleteRows(values.length + 1, maxRows - values.length);
        const maxCols = destSheet.getMaxColumns();
        if (maxCols > values[0].length) destSheet.deleteColumns(values[0].length + 1, maxCols - values[0].length);
      }
    }
    SpreadsheetApp.flush();
    const exportUrl = 'https://www.googleapis.com/drive/v3/files/' + tempId +
                      '/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const status = response.getResponseCode();
    if (status !== 200) throw new Error('xlsx export failed (HTTP ' + status + '): ' + response.getContentText().substring(0, 240));
    const bytes = response.getBlob().getBytes();
    const base64 = Utilities.base64Encode(bytes);
    return { filename: fileSpec.filename, base64: base64, sheetCount: fileSpec.sheets.length };
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); }
    catch (cleanupErr) { console.error('Failed to trash temp export file ' + tempId + ': ' + cleanupErr.message); }
  }
}

function showXlsxExportModal_(plan) {
  let totalRecords = 0; let missingCount = 0; let fileGroups = '';
  plan.forEach(function (file, fileIdx) {
    let sheetRows = '';
    file.sheets.forEach(function (s) {
      const isMissing = !!s.error;
      if (isMissing) missingCount++; else totalRecords += s.dataRows;
      const recordsCell = isMissing ? '—' : s.dataRows;
      const statusCell = isMissing ? '<span class="err-tag">' + escapeHtml_(s.error) + '</span>' : '✓';
      sheetRows += '<tr' + (isMissing ? ' class="missing"' : '') + '>' +
        '<td>' + escapeHtml_(s.destName) + '</td><td class="mono">' + escapeHtml_(s.sourceName) + '</td>' +
        '<td class="num">' + recordsCell + '</td><td class="status">' + statusCell + '</td></tr>';
    });
    fileGroups += '<div class="group"><div class="group-header">File ' + (fileIdx + 1) + ': ' + escapeHtml_(file.filename) + '</div>' +
      '<table><thead><tr><th>Destination Tab</th><th>Source Sheet</th><th class="num">Records</th><th class="status">Status</th></tr></thead><tbody>' + sheetRows + '</tbody></table></div>';
  });
  let summaryClass, summaryText;
  if (missingCount > 0) { summaryClass = 'err'; summaryText = '⚠ ' + missingCount + ' source sheet' + (missingCount === 1 ? '' : 's') + ' missing.'; }
  else { summaryClass = 'info'; summaryText = 'Ready to export ' + totalRecords + ' record' + (totalRecords === 1 ? '' : 's') + ' across ' + XLSX_EXPORT_FILES.length + ' xlsx files.'; }
  const exportDisabledAttr = missingCount > 0 ? ' disabled' : '';
  let progressItems = '';
  XLSX_EXPORT_FILES.forEach(function (f, idx) {
    progressItems += '<div class="progress-item pending" id="prog-' + idx + '"><span class="prog-icon">○</span> <span class="prog-name">' + escapeHtml_(f.filename) + '</span><span class="prog-status"></span></div>';
  });
  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}.summary.info{background:#E3F2FD;border:1px solid #64B5F6;}.summary.ok{background:#E8F5E9;border:1px solid #81C784;}.summary.err{background:#FDECEA;border:1px solid #E57373;}' +
    '.group{margin-bottom:18px;}.group-header{font-weight:600;color:#0B5394;font-size:12px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;word-break:break-all;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}th{background:#F5F8FB;text-align:left;padding:6px 8px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;font-size:11px;}th.num,td.num{text-align:right;width:70px;}th.status,td.status{text-align:center;width:50px;}td{padding:6px 8px;border-bottom:1px solid #ECF0F3;}td.mono{font-family:SFMono-Regular,Menlo,monospace;font-size:11px;color:#6B7C8C;}tr.missing td{color:#C62828;}.err-tag{color:#C62828;font-size:11px;font-style:italic;}' +
    '.progress-item{padding:8px 12px;border-bottom:1px solid #ECF0F3;font-size:12px;display:flex;align-items:center;gap:8px;}.progress-item .prog-icon{width:20px;text-align:center;font-size:14px;}.progress-item .prog-name{flex:1;word-break:break-all;}.progress-item .prog-status{font-size:11px;color:#6B7C8C;}.progress-item.pending{color:#6B7C8C;}.progress-item.active{color:#0B5394;font-weight:600;background:#F5F8FB;}.progress-item.done{color:#2E7D32;}.progress-item.error{color:#C62828;}' +
    '.err-msg{font-size:12px;color:#C62828;padding:8px 0;white-space:pre-wrap;}.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}.btn-secondary{background:white;border-color:#D6DFE8;color:#1A2733;}.btn-primary{background:#0B5394;color:white;}.btn:disabled{opacity:0.5;cursor:not-allowed;}' +
    '</style></head><body>' +
    '<div id="confirm-view"><div class="summary ' + summaryClass + '">' + summaryText + '</div>' + fileGroups + '</div>' +
    '<div id="progress-view" style="display:none;"><div class="summary info" id="progress-summary">Starting export…</div><div class="group"><div class="group-header">Progress</div>' + progressItems + '</div></div>' +
    '<div id="complete-view" style="display:none;"><div class="summary ok" id="complete-summary"></div><div class="group"><div class="group-header">Downloaded Files</div><div id="complete-list"></div></div></div>' +
    '<div id="error-view" style="display:none;"><div class="summary err">&#9888; An error occurred during export.</div><div class="err-msg" id="error-message"></div></div>' +
    '<div class="actions"><button id="cancel-btn" class="btn btn-secondary">Cancel</button><button id="export-btn" class="btn btn-primary"' + exportDisabledAttr + '>Export</button><button id="close-btn" class="btn btn-primary" style="display:none;">Close</button></div>' +
    '<script>var TOTAL=' + XLSX_EXPORT_FILES.length + ';var completed=[];' +
    'document.getElementById("cancel-btn").onclick=function(){google.script.host.close();};' +
    'document.getElementById("close-btn").onclick=function(){google.script.host.close();};' +
    'var exportBtn=document.getElementById("export-btn");if(!exportBtn.disabled){exportBtn.onclick=function(){startExport();};}' +
    'function startExport(){document.getElementById("confirm-view").style.display="none";document.getElementById("progress-view").style.display="block";document.getElementById("export-btn").style.display="none";document.getElementById("cancel-btn").disabled=true;exportNext(0);}' +
    'function exportNext(idx){if(idx>=TOTAL){finish();return;}setItemState(idx,"active","building…");document.getElementById("progress-summary").textContent="Building file "+(idx+1)+" of "+TOTAL+"…";google.script.run.withSuccessHandler(function(r){onFileBuilt(idx,r);}).withFailureHandler(function(e){onFileFailed(idx,e);}).buildXlsxForIndex(idx);}' +
    'function onFileBuilt(idx,r){triggerDownload(r.base64,r.filename);completed.push(r.filename);setItemState(idx,"done","downloaded");setTimeout(function(){exportNext(idx+1);},2500);}' +
    'function onFileFailed(idx,e){var msg=(e&&e.message)?e.message:String(e);setItemState(idx,"error","failed");document.getElementById("progress-view").style.display="none";document.getElementById("error-view").style.display="block";document.getElementById("error-message").textContent="File "+(idx+1)+" ("+XLSX_LABELS[idx]+"): "+msg;document.getElementById("cancel-btn").textContent="Close";document.getElementById("cancel-btn").disabled=false;}' +
    'function finish(){document.getElementById("progress-view").style.display="none";document.getElementById("complete-view").style.display="block";document.getElementById("complete-summary").innerHTML="&#10003; Exported "+completed.length+" xlsx file"+(completed.length===1?"":"s")+" to your Downloads folder.";var list=document.getElementById("complete-list");list.innerHTML="";for(var i=0;i<completed.length;i++){var d=document.createElement("div");d.className="progress-item done";d.innerHTML="<span class=\'prog-icon\'>✓</span><span class=\'prog-name\'>"+completed[i]+"</span>";list.appendChild(d);}document.getElementById("cancel-btn").style.display="none";document.getElementById("close-btn").style.display="inline-block";}' +
    'function setItemState(idx,state,statusText){var el=document.getElementById("prog-"+idx);if(!el)return;el.className="progress-item "+state;var icon=el.querySelector(".prog-icon");var status=el.querySelector(".prog-status");if(state==="active")icon.textContent="●";else if(state==="done")icon.textContent="✓";else if(state==="error")icon.textContent="✗";else icon.textContent="○";if(statusText)status.textContent=statusText;}' +
    'var XLSX_LABELS=' + JSON.stringify(XLSX_EXPORT_FILES.map(function(f){return f.filename;})) + ';' +
    'function triggerDownload(base64,filename){var url="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"+base64;var a=document.createElement("a");a.href=url;a.download=filename;document.body.appendChild(a);a.click();setTimeout(function(){document.body.removeChild(a);},50);}' +
    '</script></body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(580).setHeight(560), 'Export Pricebook Files');
}

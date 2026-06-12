/**
 * FieldPulse Pricebook 2.0 — Menu, Find Missing Markups,
 * Repair Calculated Columns, Gemini Sidebar
 */

function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Pricebook Tools')
    .addItem('🔐 Authorize & Activate Pricebook Tools', 'authorizeAndActivatePricebookTools')
    .addItem('☑️ Setup Guide', 'openGeminiSidebar')
    .addSeparator()
    .addSubMenu(ui.createMenu('🎚️ Pricebook Levels')
      .addItem('🔴 Basic Pricebook',      'levelViewBasic')
      .addItem('🟡 Normal Pricebook',     'levelViewNormal')
      .addItem('🟢 Everything Pricebook', 'levelViewEverything')
      .addItem('❌ Hide All Excluded Rows', 'hideDisabledRows'))
    .addSubMenu(ui.createMenu('💲 Markups')
      .addItem('🔍 Find Missing Markups', 'findMissingMarkups')
      .addItem('⚡ Run Markup Wizard',    'markupWizard'))
    .addItem('📦 Export Pricebook Files', 'exportPricebookFiles')
    .addSubMenu(ui.createMenu('🛠️ Utilities')
      .addItem('➕ Insert Row Below',           'insertRowBelowActive')
      .addItem('✅ Check for Duplicates',       'checkAllDuplicates')
      .addItem('🔧 Repair Calculated Columns',  'repairCalculatedColumns')
      .addItem('🏷️ Update Named Ranges',        'updateNamedRangesSheet'))
    .addToUi();
}

function findMissingMarkups() {
  const ss = SpreadsheetApp.getActive();
  const targets = [ITEMS_SHEET, OPTIONS_SHEET];
  const results = [];
  targets.forEach(function (def) {
    const sheet = ss.getSheetByName(def.name);
    if (!sheet) {
      results.push({ label: def.displayLabel, missing: [], error: 'Sheet "' + def.name + '" not found' });
      return;
    }
    const lastRow = sheet.getLastRow();
    const firstDataRow = def.headerRow + 1;
    if (lastRow < firstDataRow) {
      results.push({ label: def.displayLabel, missing: [] });
      return;
    }
    const numRows = lastRow - firstDataRow + 1;
    const costs   = sheet.getRange(firstDataRow, def.costCol,   numRows, 1).getValues();
    const markups = sheet.getRange(firstDataRow, def.markupCol, numRows, 1).getValues();
    const missing = [];
    for (let i = 0; i < numRows; i++) {
      const costRaw   = costs[i][0];
      const markupRaw = markups[i][0];
      const cost   = Number(costRaw);
      const markup = Number(markupRaw);
      const hasCost   = costRaw   !== '' && costRaw   !== null && !isNaN(cost)   && cost   > 0;
      const hasMarkup = markupRaw !== '' && markupRaw !== null && !isNaN(markup) && markup > 0;
      if (hasCost && !hasMarkup) missing.push({ row: firstDataRow + i, cost: cost });
    }
    results.push({ label: def.displayLabel, missing: missing });
  });
  showMissingMarkupsModal_(results);
}

function showMissingMarkupsModal_(results) {
  const totalMissing = results.reduce(function (s, r) { return s + (r.missing ? r.missing.length : 0); }, 0);
  let body = '';
  results.forEach(function (r) {
    body += '<div class="group"><div class="group-header">' + escapeHtml_(r.label) + '</div>';
    if (r.error) body += '<div class="error">' + escapeHtml_(r.error) + '</div>';
    else if (r.missing.length === 0) body += '<div class="empty">✓ No missing markups</div>';
    else {
      body += '<div class="count">' + r.missing.length + ' row' + (r.missing.length === 1 ? '' : 's') + ' with cost but no markup</div>';
      body += '<table><thead><tr><th>Row</th><th>Cost</th></tr></thead><tbody>';
      r.missing.forEach(function (m) { body += '<tr><td>' + m.row + '</td><td>$' + m.cost.toFixed(2) + '</td></tr>'; });
      body += '</tbody></table>';
    }
    body += '</div>';
  });
  const summaryClass = totalMissing > 0 ? 'warn' : 'ok';
  const summaryText  = totalMissing === 0 ? '✓ All items and option values with a cost have a markup.' :
    '⚠ ' + totalMissing + ' row' + (totalMissing === 1 ? '' : 's') + ' with cost but no markup.';
  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}' +
    '.summary.ok{background:#E8F5E9;border:1px solid #81C784;}' +
    '.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.group{margin-bottom:18px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.count{color:#C62828;font-size:11px;margin:6px 0;}.empty{color:#2E7D32;font-size:12px;padding:4px 0;}.error{color:#C62828;font-size:12px;font-style:italic;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}td{padding:5px 10px;border-bottom:1px solid #ECF0F3;}' +
    '</style></head><body><div class="summary ' + summaryClass + '">' + summaryText + '</div>' + body + '</body></html>';
  const output = HtmlService.createHtmlOutput(html).setWidth(460).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(output, 'Missing Markups Audit');
}

// =============================================================================
// REPAIR CALCULATED COLUMNS — Audit-first design with always-on sweep
//
// Click flow:
//   1. Audit row 2 of every registered column (bulk read per sheet)
//   2. Show modal listing errant columns with reasons (or "all healthy")
//   3. User clicks Repair → always sweep row 3 → maxRows on every registry
//      column to clear any orphan values below the anchor that could block
//      the ARRAYFORMULA spill, then re-audit and rewrite errant column
//      formulas at row 2.
//
// An ARRAYFORMULA column is errant if any of:
//   - Missing       → row 2 has no formula
//   - Mismatch      → row 2 formula doesn't match the registry after
//                     whitespace normalization
//   - Blocked spill → formula matches but row 2 displays #REF! (something
//                     below the anchor is blocking the array expansion)
//
// Sweep runs on every Repair invocation regardless of errant status.
// Repair (per errant column): clear anchor cell, flush, setFormula(row 2),
// flush.
// =============================================================================

function repairCalculatedColumns() {
  const ss = SpreadsheetApp.getActive();
  let audit;
  try {
    audit = auditCalculatedColumns_(ss);
  } catch (err) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Repair Calculated Columns — Error', err.message, ui.ButtonSet.OK);
    return;
  }
  showRepairAuditModal_(audit);
}

function auditCalculatedColumns_(ss) {
  const result = { sheets: [], totalErrant: 0 };

  CALC_SHEETS.forEach(function (sheetSpec) {
    const sheetReport = {
      label: sheetSpec.displayLabel,
      name: sheetSpec.name,
      error: null,
      totalColumns: Object.keys(sheetSpec.formulas).length,
      okCount: 0,
      errant: []
    };
    result.sheets.push(sheetReport);

    const sheet = ss.getSheetByName(sheetSpec.name);
    if (!sheet) {
      sheetReport.error = 'Sheet "' + sheetSpec.name + '" not found';
      return;
    }

    const lastCol = sheet.getLastColumn();
    const maxRows = sheet.getMaxRows();
    const firstDataRow = sheetSpec.headerRow + 1;

    if (maxRows < firstDataRow || lastCol < 1) {
      Object.keys(sheetSpec.formulas).forEach(function (letter) {
        sheetReport.errant.push({ letter: letter, reason: 'Missing' });
        result.totalErrant++;
      });
      return;
    }

    const formulasRow = sheet.getRange(firstDataRow, 1, 1, lastCol).getFormulas()[0];
    const displaysRow = sheet.getRange(firstDataRow, 1, 1, lastCol).getDisplayValues()[0];

    Object.keys(sheetSpec.formulas).forEach(function (letter) {
      const colIdx = colLetterToIndex_(letter);
      const arrIdx = colIdx - 1;

      let actualFormula = '';
      let displayValue = '';
      if (arrIdx < lastCol) {
        actualFormula = formulasRow[arrIdx] || '';
        displayValue = displaysRow[arrIdx] || '';
      }

      const expectedNorm = normalizeFormula_(sheetSpec.formulas[letter]);
      const actualNorm = normalizeFormula_(actualFormula);

      let reason = null;
      if (!actualNorm) {
        reason = 'Missing';
      } else if (actualNorm !== expectedNorm) {
        reason = 'Mismatch';
      } else if (displayValue && displayValue.indexOf('#REF!') !== -1) {
        reason = 'Blocked spill';
      }

      if (reason) {
        sheetReport.errant.push({ letter: letter, reason: reason });
        result.totalErrant++;
      } else {
        sheetReport.okCount++;
      }
    });
  });

  return result;
}

function normalizeFormula_(s) {
  if (!s) return '';
  return String(s).replace(/\s+/g, ' ').trim();
}

function repairErrantCalculatedColumns() {
  const ss = SpreadsheetApp.getActive();
  const audit = auditCalculatedColumns_(ss);

  const result = { sheets: [], totalRepaired: 0, totalSwept: 0 };

  CALC_SHEETS.forEach(function (sheetSpec) {
    const sheetReport = {
      label: sheetSpec.displayLabel,
      name: sheetSpec.name,
      swept: 0,
      repaired: [],
      errors: []
    };
    result.sheets.push(sheetReport);

    const sheetAudit = audit.sheets.filter(function (s) { return s.name === sheetSpec.name; })[0];
    if (sheetAudit && sheetAudit.error) {
      sheetReport.errors.push(sheetAudit.error);
      return;
    }

    const sheet = ss.getSheetByName(sheetSpec.name);
    if (!sheet) {
      sheetReport.errors.push('Sheet not found');
      return;
    }

    const maxRows = sheet.getMaxRows();
    const firstDataRow = sheetSpec.headerRow + 1;       // row 2 = ARRAYFORMULA anchor
    const firstSweepRow = firstDataRow + 1;             // row 3 = first row below anchor
    const sweepNumRows = maxRows - firstSweepRow + 1;

    // ---- SWEEP: clear row 3 → maxRows on every registry column ----
    // Runs unconditionally to remove any orphan static values that could
    // block the ARRAYFORMULA spill. Anchor row (row 2) is preserved.
    if (sweepNumRows > 0) {
      Object.keys(sheetSpec.formulas).forEach(function (letter) {
        try {
          const colIdx = colLetterToIndex_(letter);
          sheet.getRange(firstSweepRow, colIdx, sweepNumRows, 1).clearContent();
          sheetReport.swept++;
          result.totalSwept++;
        } catch (err) {
          sheetReport.errors.push('Sweep ' + letter + ' (' + err.message + ')');
        }
      });
      SpreadsheetApp.flush();
    }

    // ---- REPAIR: rewrite formulas for errant columns ----
    if (!sheetAudit || !sheetAudit.errant || sheetAudit.errant.length === 0) return;

    sheetAudit.errant.forEach(function (item) {
      const letter = item.letter;
      try {
        const colIdx = colLetterToIndex_(letter);
        sheet.getRange(firstDataRow, colIdx).clearContent();
        SpreadsheetApp.flush();
        sheet.getRange(firstDataRow, colIdx).setFormula(sheetSpec.formulas[letter]);
        SpreadsheetApp.flush();
        sheetReport.repaired.push(letter);
        result.totalRepaired++;
      } catch (err) {
        sheetReport.errors.push(letter + ' (' + err.message + ')');
      }
    });
  });

  return result;
}

function showRepairAuditModal_(audit) {
  const totalErrant = audit.totalErrant;

  let summaryClass, summaryText;
  if (totalErrant === 0) {
    summaryClass = 'ok';
    summaryText = '✓ All formulas healthy. Sweep ready to run.';
  } else {
    summaryClass = 'warn';
    summaryText = '⚠ ' + totalErrant + ' column' + (totalErrant === 1 ? '' : 's') + ' need repair.';
  }

  let body = '';
  audit.sheets.forEach(function (s) {
    body += '<div class="group"><div class="group-header">' + escapeHtml_(s.label) + '</div>';
    if (s.error) {
      body += '<div class="error">' + escapeHtml_(s.error) + '</div>';
    } else if (!s.errant || s.errant.length === 0) {
      body += '<div class="empty">✓ All ' + s.totalColumns + ' formulas OK</div>';
    } else {
      body += '<div class="count">' + s.errant.length + ' of ' + s.totalColumns + ' formula' +
              (s.errant.length === 1 ? '' : 's') + ' need repair (' + s.okCount + ' OK)</div>';
      body += '<table><thead><tr><th>Column</th><th>Reason</th></tr></thead><tbody>';
      s.errant.forEach(function (e) {
        body += '<tr><td>' + escapeHtml_(e.letter) + '</td><td>' + escapeHtml_(e.reason) + '</td></tr>';
      });
      body += '</tbody></table>';
    }
    body += '</div>';
  });

  const repairNote = '<div class="note">Repair will sweep all registry columns below row 2 and rewrite any errant formulas. This action cannot be undone.</div>';
  const buttonsHtml = '<button id="cancel-btn" class="btn btn-secondary">Cancel</button>' +
                      '<button id="repair-btn" class="btn btn-primary">Repair</button>';

  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}' +
    '.summary.ok{background:#E8F5E9;border:1px solid #81C784;}' +
    '.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.summary.info{background:#E3F2FD;border:1px solid #64B5F6;}' +
    '.summary.err{background:#FDECEA;border:1px solid #E57373;}' +
    '.group{margin-bottom:18px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.count{color:#C62828;font-size:11px;margin:6px 0;}' +
    '.empty{color:#2E7D32;font-size:12px;padding:4px 0;}' +
    '.error{color:#C62828;font-size:12px;font-style:italic;}' +
    '.note{font-size:11px;color:#6B7C8C;font-style:italic;padding:8px 0;margin-bottom:8px;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}' +
    'th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}' +
    'td{padding:5px 10px;border-bottom:1px solid #ECF0F3;}' +
    '.repaired{font-size:12px;color:#2E7D32;padding:4px 0;}' +
    '.swept{font-size:11px;color:#6B7C8C;padding:2px 0;}' +
    '.err-msg{font-size:12px;color:#C62828;padding:8px 0;white-space:pre-wrap;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-secondary{background:white;border-color:#D6DFE8;color:#1A2733;}' +
    '.btn-primary{background:#0B5394;color:white;}' +
    '.btn:disabled{opacity:0.6;cursor:not-allowed;}' +
    '</style></head><body>' +
    '<div id="audit-view">' +
    '<div class="summary ' + summaryClass + '">' + summaryText + '</div>' +
    body + repairNote + '</div>' +
    '<div id="repairing-view" style="display:none;">' +
    '<div class="summary info">Sweeping columns and repairing formulas. This may take a moment\u2026</div></div>' +
    '<div id="complete-view" style="display:none;">' +
    '<div class="summary ok" id="complete-summary"></div>' +
    '<div id="complete-body"></div></div>' +
    '<div id="error-view" style="display:none;">' +
    '<div class="summary err">&#9888; An error occurred during repair.</div>' +
    '<div class="err-msg" id="error-message"></div></div>' +
    '<div class="actions">' + buttonsHtml +
    '<button id="close-after-btn" class="btn btn-primary" style="display:none;">Close</button></div>' +
    '<script>(function(){' +
    'var cancelBtn=document.getElementById("cancel-btn");' +
    'var repairBtn=document.getElementById("repair-btn");' +
    'var closeAfterBtn=document.getElementById("close-after-btn");' +
    'if(cancelBtn)cancelBtn.onclick=function(){google.script.host.close();};' +
    'if(closeAfterBtn)closeAfterBtn.onclick=function(){google.script.host.close();};' +
    'if(repairBtn)repairBtn.onclick=function(){' +
    'document.getElementById("audit-view").style.display="none";' +
    'document.getElementById("repairing-view").style.display="block";' +
    'cancelBtn.style.display="none";repairBtn.style.display="none";' +
    'google.script.run.withSuccessHandler(onOk).withFailureHandler(onErr).repairErrantCalculatedColumns();};' +
    'function onOk(r){' +
    'document.getElementById("repairing-view").style.display="none";' +
    'document.getElementById("complete-view").style.display="block";' +
    'document.getElementById("complete-summary").innerHTML="&#10003; Swept "+r.totalSwept+" column"+(r.totalSwept===1?"":"s")+", repaired "+r.totalRepaired+" formula"+(r.totalRepaired===1?"":"s")+".";' +
    'var html="";r.sheets.forEach(function(s){' +
    'if(s.repaired.length===0&&s.errors.length===0&&s.swept===0)return;' +
    'html+="<div class=\\"group\\"><div class=\\"group-header\\">"+s.label+"</div>";' +
    'if(s.swept>0)html+="<div class=\\"swept\\">Swept "+s.swept+" column"+(s.swept===1?"":"s")+" below row 2</div>";' +
    'if(s.repaired.length>0)html+="<div class=\\"repaired\\">Repaired: "+s.repaired.join(", ")+"</div>";' +
    'if(s.errors.length>0)html+="<div class=\\"error\\">Errors: "+s.errors.join("; ")+"</div>";' +
    'html+="</div>";});' +
    'document.getElementById("complete-body").innerHTML=html;' +
    'closeAfterBtn.style.display="inline-block";}' +
    'function onErr(e){' +
    'document.getElementById("repairing-view").style.display="none";' +
    'document.getElementById("error-view").style.display="block";' +
    'document.getElementById("error-message").textContent=(e&&e.message)?e.message:String(e);' +
    'closeAfterBtn.style.display="inline-block";}' +
    '})();</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(560), 'Repair Calculated Columns');
}

function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function openGeminiSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('GeminiSidebar').setTitle('Pricebook Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

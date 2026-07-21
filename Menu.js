function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Pricebook Tools')
    // --- Setup & help ---
    .addItem('🔐 Activate Pricebook Tools', 'authorizeAndActivatePricebookTools')
    .addItem('❓ Setup Guide', 'openGeminiSidebar')
    .addSeparator()
    // --- View & price ---
    .addSubMenu(ui.createMenu('🎚️ Pricebook Levels')
      .addItem('🔴 Basic',      'levelViewBasic')
      .addItem('🟡 Normal',     'levelViewNormal')
      .addItem('🟢 Everything', 'levelViewEverything')
      .addItem('❌ Hide All Excluded Rows', 'hideDisabledRows'))
    .addSubMenu(ui.createMenu('💲 Markups')
      .addItem('🔍 Find Missing Markups', 'findMissingMarkups')
      .addItem('⚡ Run Markup Wizard',    'markupWizard'))
    .addSeparator()
    // --- Check & export ---
    .addItem('🩺 Pre-Export Health Check', 'preExportHealthCheck')
    .addItem('📦 Export Pricebook Files', 'exportPricebookFiles')
    .addSeparator()
    // --- Safety net ---
    .addSubMenu(ui.createMenu('🛠️ Utilities')
      .addItem('➕ Insert Row Below',           'insertRowBelowActive')
      .addItem('✅ Check for Duplicates',       'checkAllDuplicates')
      .addItem('🔎 Check Defined Levels',       'checkDefinedLevels')
      .addItem('🧾 Check Item Accounts',        'checkItemAccounts')
      .addItem('🔗 Link Option Names to Values', 'linkOptionNamesToValues')
      .addItem('🔧 Repair Calculated Columns',  'repairCalculatedColumns')
      .addItem('🪜 Repair Variables Tier Formulas', 'repairVariablesTierFormulas')
      .addItem('🏷️ Update Named Ranges (advanced)', 'updateNamedRangesSheet'))
    .addToUi();
}

// =============================================================================
// MISSING MARKUPS — audit worker + menu entry
// auditMissingMarkups_ is shared by the standalone menu item AND the
// Pre-Export Health Check, so both report identically.
// =============================================================================

function auditMissingMarkups_(ss) {
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
  return results;
}

function findMissingMarkups() {
  const ss = SpreadsheetApp.getActive();
  showMissingMarkupsModal_(auditMissingMarkups_(ss));
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
// PRE-EXPORT HEALTH CHECK — one-click readiness gate (v1)
//
// Aggregates the audits already defined in this file into a single verdict:
//   • Markups            — auditMissingMarkups_        (cost but no markup)
//   • Calculated Columns — auditCalculatedColumns_     (ARRAYFORMULA anchors)
//   • Tier Pricing       — auditVariablesTierFormulas_ (Variables tier tables)
// Read-only: it never writes. Each failing check names the tool that fixes it.
//
// DEFERRED to v2 (workers live in other files / need a spec decision):
//   • Duplicates (Check for Duplicates), export-source readiness (Export.js),
//     stray #REF!/#N/A error-cell scan, and blank required-field check.
// =============================================================================

function preExportHealthCheck() {
  const ss = SpreadsheetApp.getActive();
  const checks = [];

  // 1) Markups
  try {
    const mm = auditMissingMarkups_(ss);
    const miss = mm.reduce(function (s, r) { return s + (r.missing ? r.missing.length : 0); }, 0);
    const errs = mm.filter(function (r) { return r.error; }).map(function (r) { return r.error; });
    if (errs.length) checks.push({ label: 'Markups', ok: false, detail: errs.join('; ') });
    else checks.push({
      label: 'Markups',
      ok: miss === 0,
      detail: miss === 0 ? 'Every costed item and option value has a markup.'
                         : miss + ' row' + (miss === 1 ? '' : 's') + ' have a cost but no markup.',
      hint: '💲 Markups → Find Missing Markups'
    });
  } catch (e) {
    checks.push({ label: 'Markups', ok: false, detail: 'Check could not run: ' + e.message });
  }

  // 2) Calculated Columns
  try {
    const cc = auditCalculatedColumns_(ss);
    checks.push({
      label: 'Calculated Columns',
      ok: cc.totalErrant === 0,
      detail: cc.totalErrant === 0 ? 'All ARRAYFORMULA anchor columns are healthy.'
                                   : cc.totalErrant + ' column' + (cc.totalErrant === 1 ? '' : 's') + ' need repair.',
      hint: '🛠️ Utilities → Repair Calculated Columns'
    });
  } catch (e) {
    checks.push({ label: 'Calculated Columns', ok: false, detail: 'Check could not run: ' + e.message });
  }

  // 3) Tier Pricing
  try {
    const tf = auditVariablesTierFormulas_(ss);
    if (tf.error) {
      checks.push({ label: 'Tier Pricing', ok: false, detail: tf.error });
    } else {
      checks.push({
        label: 'Tier Pricing',
        ok: tf.errant.length === 0,
        detail: tf.errant.length === 0 ? 'All ' + tf.total + ' tier formulas are healthy.'
                                       : tf.errant.length + ' of ' + tf.total + ' tier formula' + (tf.errant.length === 1 ? '' : 's') + ' need repair.',
        hint: '🛠️ Utilities → Repair Variables Tier Formulas'
      });
    }
  } catch (e) {
    checks.push({ label: 'Tier Pricing', ok: false, detail: 'Check could not run: ' + e.message });
  }

  // 4) Item Accounts — every included item needs Income (Z) + Purchase (AA)
  try {
    const ia = auditItemAccounts_(ss);
    if (ia.error) {
      checks.push({ label: 'Item Accounts', ok: false, detail: ia.error });
    } else {
      checks.push({
        label: 'Item Accounts',
        ok: ia.bad.length === 0,
        detail: ia.bad.length === 0 ? 'Every included item has an Income and Purchase account.'
                                    : ia.bad.length + ' included item' + (ia.bad.length === 1 ? '' : 's') + ' missing an account.',
        hint: '🧾 Check Item Accounts'
      });
    }
  } catch (e) {
    checks.push({ label: 'Item Accounts', ok: false, detail: 'Check could not run: ' + e.message });
  }

  showHealthCheckModal_(checks);
}

function showHealthCheckModal_(checks) {
  const failing = checks.filter(function (c) { return !c.ok; }).length;

  let summaryClass, summaryText;
  if (failing === 0) {
    summaryClass = 'ok';
    summaryText = '✓ You’re ready to export — all ' + checks.length + ' checks passed.';
  } else {
    summaryClass = 'warn';
    summaryText = '⚠ ' + failing + ' of ' + checks.length + ' check' + (failing === 1 ? '' : 's') + ' need attention before export.';
  }

  let body = '';
  checks.forEach(function (c) {
    body += '<div class="group"><div class="group-header">' + escapeHtml_(c.label) + '</div>';
    if (c.ok) {
      body += '<div class="empty">✓ ' + escapeHtml_(c.detail) + '</div>';
    } else {
      body += '<div class="count">⚠ ' + escapeHtml_(c.detail) + '</div>';
      if (c.hint) body += '<div class="note">Fix with: ' + escapeHtml_(c.hint) + '</div>';
    }
    body += '</div>';
  });

  const footNote = failing === 0
    ? '<div class="note">All clear. Run 📦 Export Pricebook Files to finish.</div>'
    : '<div class="note">Clear the items above, then run this check again before exporting.</div>';

  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}' +
    '.summary.ok{background:#E8F5E9;border:1px solid #81C784;}' +
    '.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    '.group{margin-bottom:16px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.count{color:#C62828;font-size:12px;margin:6px 0;}.empty{color:#2E7D32;font-size:12px;padding:4px 0;}' +
    '.note{font-size:11px;color:#6B7C8C;font-style:italic;padding:4px 0;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-primary{background:#0B5394;color:white;}' +
    '</style></head><body>' +
    '<div class="summary ' + summaryClass + '">' + summaryText + '</div>' + body + footNote +
    '<div class="actions"><button id="close-btn" class="btn btn-primary">Close</button></div>' +
    '<script>document.getElementById("close-btn").onclick=function(){google.script.host.close();};</script>' +
    '</body></html>';

  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(480).setHeight(520), 'Pre-Export Health Check');
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
    '<div class="summary info">Sweeping columns and repairing formulas. This may take a moment…</div></div>' +
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

// =============================================================================
// REPAIR VARIABLES TIER FORMULAS — per-cell tier-cliff repair
//
// The Variables sheet holds two cost-tier tables: Glass & Mirrors (rows 12–21)
// and Hardware, Parts, Other (rows 26–35). Each tier row carries three
// generated formulas — G (Note text), I (Sell @ Ceiling), K (Cliff to Next
// Tier). K on the LAST row of each table is a static "—" (no next tier) and is
// intentionally NOT generated, so repair never writes those two cells.
//
// Canonical formulas come from VARIABLES_TIER_FORMULAS (Formulas.js), built by
// buildVariablesTierFormulas_() from VARIABLES_TIER_TABLES. They are per-cell
// scalar formulas, NOT ARRAYFORMULAs — there is no spill and no sweep. Repair
// audits each generated cell and rewrites any that are Missing or drifted
// (Mismatch). Click flow mirrors Repair Calculated Columns:
//   audit -> modal -> confirm -> rewrite errant cells -> completion summary.
// =============================================================================

function repairVariablesTierFormulas() {
  const ss = SpreadsheetApp.getActive();
  let audit;
  try {
    audit = auditVariablesTierFormulas_(ss);
  } catch (err) {
    const ui = SpreadsheetApp.getUi();
    ui.alert('Repair Variables Tier Formulas — Error', err.message, ui.ButtonSet.OK);
    return;
  }
  showVariablesTierAuditModal_(audit);
}

function tierCellRef_(key) {
  const s = String(key).toUpperCase();
  let i = 0;
  while (i < s.length && s[i] >= 'A' && s[i] <= 'Z') i++;
  const letter = s.substring(0, i);
  const row = parseInt(s.substring(i), 10);
  if (!letter || !(row > 0)) return null;
  return { letter: letter, row: row };
}

function auditVariablesTierFormulas_(ss) {
  const sheetName = 'Variables';
  const report = { sheetName: sheetName, error: null, total: 0, okCount: 0, errant: [] };

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    report.error = 'Sheet "' + sheetName + '" not found';
    return report;
  }

  const keys = Object.keys(VARIABLES_TIER_FORMULAS);
  report.total = keys.length;
  if (keys.length === 0) return report;

  let minRow = Infinity, maxRow = -Infinity, minCol = Infinity, maxCol = -Infinity;
  const cells = [];
  keys.forEach(function (key) {
    const ref = tierCellRef_(key);
    if (!ref) return;
    const colIdx = colLetterToIndex_(ref.letter);
    cells.push({ key: key, row: ref.row, colIdx: colIdx });
    if (ref.row < minRow) minRow = ref.row;
    if (ref.row > maxRow) maxRow = ref.row;
    if (colIdx < minCol) minCol = colIdx;
    if (colIdx > maxCol) maxCol = colIdx;
  });

  const liveFormulas = sheet
    .getRange(minRow, minCol, maxRow - minRow + 1, maxCol - minCol + 1)
    .getFormulas();

  cells.forEach(function (c) {
    const actualFormula = liveFormulas[c.row - minRow][c.colIdx - minCol] || '';
    const expectedNorm = normalizeFormula_(VARIABLES_TIER_FORMULAS[c.key]);
    const actualNorm = normalizeFormula_(actualFormula);
    let reason = null;
    if (!actualNorm) reason = 'Missing';
    else if (actualNorm !== expectedNorm) reason = 'Mismatch';
    if (reason) report.errant.push({ cell: c.key, reason: reason });
    else report.okCount++;
  });

  report.errant.sort(function (a, b) {
    const ra = tierCellRef_(a.cell), rb = tierCellRef_(b.cell);
    if (ra.row !== rb.row) return ra.row - rb.row;
    return colLetterToIndex_(ra.letter) - colLetterToIndex_(rb.letter);
  });

  return report;
}

function repairErrantVariablesTierFormulas() {
  const ss = SpreadsheetApp.getActive();
  const audit = auditVariablesTierFormulas_(ss);
  const result = { sheetName: audit.sheetName, repaired: [], errors: [] };

  if (audit.error) { result.errors.push(audit.error); return result; }

  const sheet = ss.getSheetByName(audit.sheetName);
  if (!sheet) { result.errors.push('Sheet not found'); return result; }

  audit.errant.forEach(function (item) {
    const ref = tierCellRef_(item.cell);
    try {
      sheet.getRange(ref.row, colLetterToIndex_(ref.letter))
           .setFormula(VARIABLES_TIER_FORMULAS[item.cell]);
      result.repaired.push(item.cell);
    } catch (err) {
      result.errors.push(item.cell + ' (' + err.message + ')');
    }
  });
  SpreadsheetApp.flush();
  return result;
}

function showVariablesTierAuditModal_(audit) {
  const hasError = !!audit.error;
  const totalErrant = hasError ? 0 : audit.errant.length;

  let summaryClass, summaryText;
  if (hasError) {
    summaryClass = 'err';
    summaryText = '⚠ ' + audit.error;
  } else if (totalErrant === 0) {
    summaryClass = 'ok';
    summaryText = '✓ All ' + audit.total + ' tier formulas healthy.';
  } else {
    summaryClass = 'warn';
    summaryText = '⚠ ' + totalErrant + ' of ' + audit.total + ' tier formula' +
                  (totalErrant === 1 ? '' : 's') + ' need repair.';
  }

  let body = '';
  if (!hasError) {
    body += '<div class="group"><div class="group-header">' + escapeHtml_(audit.sheetName) +
            ' — Cost Tier Tables</div>';
    if (totalErrant === 0) {
      body += '<div class="empty">✓ All ' + audit.total +
              ' formulas OK (Note / Sell @ Ceiling / Cliff)</div>';
    } else {
      body += '<div class="count">' + totalErrant + ' of ' + audit.total + ' formula' +
              (totalErrant === 1 ? '' : 's') + ' need repair (' + audit.okCount + ' OK)</div>';
      body += '<table><thead><tr><th>Cell</th><th>Reason</th></tr></thead><tbody>';
      audit.errant.forEach(function (e) {
        body += '<tr><td>' + escapeHtml_(e.cell) + '</td><td>' + escapeHtml_(e.reason) + '</td></tr>';
      });
      body += '</tbody></table>';
    }
    body += '</div>';
  }

  const repairNote = hasError ? '' :
    '<div class="note">Repair rewrites the generated Note (G), Sell @ Ceiling (I) and Cliff to Next ' +
    'Tier (K) formulas for any drifted cell. The static "—" in the last Cliff cell of each table is ' +
    'left untouched. This action cannot be undone.</div>';

  const repairBtnHtml = hasError ? '' :
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
    '.error{color:#C62828;font-size:12px;font-style:italic;padding:4px 0;}' +
    '.detail{color:#2E7D32;font-size:12px;padding:4px 0;white-space:pre-wrap;}' +
    '.note{font-size:11px;color:#6B7C8C;font-style:italic;padding:8px 0;margin-bottom:8px;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}' +
    'th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}' +
    'td{padding:5px 10px;border-bottom:1px solid #ECF0F3;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-secondary{background:white;border-color:#D6DFE8;color:#1A2733;}' +
    '.btn-primary{background:#0B5394;color:white;}' +
    '</style></head><body>' +
    '<div id="audit-view">' +
    '<div class="summary ' + summaryClass + '">' + summaryText + '</div>' +
    body + repairNote + '</div>' +
    '<div id="repairing-view" style="display:none;">' +
    '<div class="summary info">Repairing tier formulas…</div></div>' +
    '<div id="complete-view" style="display:none;">' +
    '<div class="summary ok" id="complete-summary"></div>' +
    '<div class="detail" id="complete-repaired"></div>' +
    '<div class="error" id="complete-errors"></div></div>' +
    '<div id="error-view" style="display:none;">' +
    '<div class="summary err">⚠ An error occurred during repair.</div>' +
    '<div class="error" id="error-message"></div></div>' +
    '<div class="actions">' +
    '<button id="cancel-btn" class="btn btn-secondary">Cancel</button>' + repairBtnHtml +
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
    'google.script.run.withSuccessHandler(onOk).withFailureHandler(onErr).repairErrantVariablesTierFormulas();};' +
    'function onOk(r){' +
    'document.getElementById("repairing-view").style.display="none";' +
    'document.getElementById("complete-view").style.display="block";' +
    'document.getElementById("complete-summary").innerHTML="&#10003; Repaired "+r.repaired.length+" tier formula"+(r.repaired.length===1?"":"s")+".";' +
    'document.getElementById("complete-repaired").textContent=r.repaired.length?("Repaired: "+r.repaired.join(", ")):"";' +
    'document.getElementById("complete-errors").textContent=r.errors.length?("Errors: "+r.errors.join("; ")):"";' +
    'closeAfterBtn.style.display="inline-block";}' +
    'function onErr(e){' +
    'document.getElementById("repairing-view").style.display="none";' +
    'document.getElementById("error-view").style.display="block";' +
    'document.getElementById("error-message").textContent=(e&&e.message)?e.message:String(e);' +
    'closeAfterBtn.style.display="inline-block";}' +
    '})();</script></body></html>';

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(540).setHeight(560),
    'Repair Variables Tier Formulas'
  );
}

// =============================================================================
// SETUP GUIDE SIDEBAR — narrow docked panel + wide floating view
//
// The GeminiSidebar HTML is now an Apps Script *template* (uses <?= mode ?>),
// so it must be served via createTemplateFromFile().evaluate(), NOT
// createHtmlOutputFromFile(). 'sidebar' = narrow docked; 'wide' = floating
// modeless dialog. The in-panel Expand/Dock buttons swap between the two.
// =============================================================================

function openGeminiSidebar() {
  const t = HtmlService.createTemplateFromFile('GeminiSidebar');
  t.mode = 'sidebar';
  const html = t.evaluate().setTitle('Pricebook Assistant');
  SpreadsheetApp.getUi().showSidebar(html);
}

function openGeminiWide() {
  const t = HtmlService.createTemplateFromFile('GeminiSidebar');
  t.mode = 'wide';
  const html = t.evaluate().setWidth(600).setHeight(680);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Pricebook Assistant');
}

function getPricebookFileName() {
  return SpreadsheetApp.getActiveSpreadsheet().getName();
}

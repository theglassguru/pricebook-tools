/**
 * FieldPulse Pricebook 2.0 — Authorize & Activate
 *
 * One-time per-user setup: installs the onEdit duplicate trigger, applies
 * duplicate-flag conditional formatting rules to the configured scopes, and
 * refreshes the Named Ranges sheet. Surfaces a single results modal so the
 * user sees exactly what was activated (or what failed).
 *
 * Depends on:
 *   - Duplicates.gs: DUPLICATE_SCOPES, DUP_TRIGGER_HANDLER, DUP_HIGHLIGHT_BG,
 *                    DUP_HIGHLIGHT_FG, buildDuplicateCfFormula_
 *   - Admin.gs:      updateNamedRangesSheet_
 *   - Menu.gs:       escapeHtml_
 */

function authorizeAndActivatePricebookTools() {
  const ss = SpreadsheetApp.getActive();
  const results = [];

  // Step 1: onEdit trigger for real-time duplicate alerts
  try {
    const existing = ScriptApp.getProjectTriggers().filter(function (t) {
      return t.getHandlerFunction() === DUP_TRIGGER_HANDLER && t.getEventType() === ScriptApp.EventType.ON_EDIT;
    });
    if (existing.length > 0) {
      results.push({ label: 'Real-time duplicate alerts', ok: true, note: 'already active for this user' });
    } else {
      ScriptApp.newTrigger(DUP_TRIGGER_HANDLER).forSpreadsheet(ss).onEdit().create();
      results.push({ label: 'Real-time duplicate alerts', ok: true, note: 'enabled for this user' });
    }
  } catch (err) {
    results.push({ label: 'Real-time duplicate alerts', ok: false, error: err.message });
  }

  // Step 2: Conditional formatting rules for duplicate flagging
  DUPLICATE_SCOPES.forEach(function (scope) {
    const sheet = ss.getSheetByName(scope.sheetName);
    if (!sheet) {
      results.push({ label: scope.sheetName + ' · ' + scope.columnLetter + ' highlighting', ok: false, error: 'Sheet not found' });
      return;
    }
    try {
      const rangeA1 = scope.columnLetter + '2:' + scope.columnLetter;
      const range = sheet.getRange(rangeA1);
      const formula = buildDuplicateCfFormula_(scope);
      const purgeList = [formula].concat(scope.legacyFormulas || []);
      const existing = sheet.getConditionalFormatRules();
      let purgedCount = 0;
      const keep = existing.filter(function (rule) {
        const bc = rule.getBooleanCondition();
        if (!bc) return true;
        const values = bc.getCriteriaValues();
        if (!values || values.length === 0) return true;
        if (purgeList.indexOf(values[0]) !== -1) { purgedCount++; return false; }
        return true;
      });
      const newRule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(formula).setBackground(DUP_HIGHLIGHT_BG).setFontColor(DUP_HIGHLIGHT_FG).setRanges([range]).build();
      keep.push(newRule);
      sheet.setConditionalFormatRules(keep);
      const legacyPurged = purgedCount - 1;
      let note = 'rule applied';
      if (legacyPurged > 0) note += ' (' + legacyPurged + ' legacy rule' + (legacyPurged === 1 ? '' : 's') + ' removed)';
      results.push({ label: scope.sheetName + ' · ' + scope.columnLetter + ' (' + scope.columnLabel + ') highlighting', ok: true, note: note });
    } catch (err) {
      results.push({ label: scope.sheetName + ' · ' + scope.columnLetter + ' highlighting', ok: false, error: err.message });
    }
  });

  // Step 3: Refresh Named Ranges sheet
  try {
    const nrResult = updateNamedRangesSheet_();
    if (nrResult.ok) {
      results.push({ label: 'Named Ranges sheet', ok: true, note: nrResult.note });
    } else {
      results.push({ label: 'Named Ranges sheet', ok: false, error: nrResult.error });
    }
  } catch (err) {
    results.push({ label: 'Named Ranges sheet', ok: false, error: err.message });
  }

  showAuthorizeActivateModal_(results);
}

function showAuthorizeActivateModal_(results) {
  let allOk = true;
  let body = '<table><thead><tr><th>Step</th><th>Status</th></tr></thead><tbody>';
  results.forEach(function (r) {
    if (r.ok) body += '<tr><td>' + escapeHtml_(r.label) + '</td><td class="ok">✓ ' + escapeHtml_(r.note || 'done') + '</td></tr>';
    else { allOk = false; body += '<tr><td>' + escapeHtml_(r.label) + '</td><td class="err">✗ ' + escapeHtml_(r.error) + '</td></tr>'; }
  });
  body += '</tbody></table>';
  const summaryClass = allOk ? 'ok' : 'warn';
  const summaryText = allOk ? '✓ Pricebook Tools authorized and active for this user.' : '⚠ Activation completed with errors.';
  const html = '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;}.summary.ok{background:#E8F5E9;border:1px solid #81C784;}.summary.warn{background:#FFF4E5;border:1px solid #FFB74D;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;}th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}td{padding:6px 10px;border-bottom:1px solid #ECF0F3;}td.ok{color:#2E7D32;}td.err{color:#C62828;}' +
    '</style></head><body><div class="summary ' + summaryClass + '">' + summaryText + '</div>' + body + '</body></html>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(540).setHeight(420), 'Authorize & Activate Pricebook Tools');
}

/**
 * FieldPulse Pricebook 2.0 — Markup Wizard + Shared Config
 * Container-bound Apps Script for the Dimensional Pricebook Setup sheet.
 *
 * This file holds:
 *   - Variables sheet config (two tier ranges: Glass & Mirrors, Hardware/Parts/Other)
 *   - Wizard-targeted sheet definitions (ITEMS_SHEET, OPTIONS_SHEET)
 *     used by the Markup Wizard and Find Missing Markups
 *   - Markup Wizard implementation (glass + hardware passes, Product type filter)
 *
 * The canonical ARRAYFORMULA registry for Repair Calculated Columns lives in
 * Formulas.gs. Menu wiring, Find Missing Markups, and Repair Calculated
 * Columns live in Menu.gs.
 *
 * Wizard modes:
 *   - 'fillMissing' (default): write tier markups only to eligible rows whose
 *     current Mark Up is empty or non-numeric. Rows with any positive Mark Up
 *     (manually customized or set by a prior wizard run) are preserved.
 *   - 'overwriteAll': write tier markups to every eligible row, overwriting
 *     any existing Mark Up.
 *
 * In both modes, rows with Mark Up == 1 are always preserved (flat-fee
 * passthroughs such as Service surcharges).
 */

// =============================================================================
// CONFIGURATION
// =============================================================================

const VARIABLES_SHEET_NAME    = 'Variables';
const GLASS_TIER_RANGE_A1     = 'C12:E21';   // 10 rows: Glass & Mirrors
const HARDWARE_TIER_RANGE_A1  = 'C26:E35';   // 10 rows: Hardware, Parts, Other Items

const GLASS_GROUP_LABEL    = 'Glass & Mirrors';
const HARDWARE_GROUP_LABEL = 'Hardware, Parts, Other Items';

const ITEMS_SHEET = {
  name:         'Items',
  displayLabel: 'Items Sheet',
  categoryCol:  4,
  typeCol:      6,
  costCol:      8,
  markupCol:    9,
  headerRow:    1
};

const OPTIONS_SHEET = {
  name:         'Item Option Values',
  displayLabel: 'Item Option Values Sheet',
  categoryCol:  4,
  typeCol:      null,   // IOV has no Type column; every row is inherently a Product
  costCol:      8,
  markupCol:    9,
  headerRow:    1
};

// =============================================================================
// WIZARD
// Applies markups to both glass-family and hardware/parts/other Product rows
// using their respective tier tables on the Variables sheet.
// =============================================================================

function markupWizard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let glassTiers, hardwareTiers, itemsAnalysis, optionsAnalysis;
  try {
    glassTiers    = readTierConfig_(ss, GLASS_TIER_RANGE_A1);
    hardwareTiers = readTierConfig_(ss, HARDWARE_TIER_RANGE_A1);
    if (glassTiers.length === 0 && hardwareTiers.length === 0) {
      ui.alert('Markup Wizard',
               'No valid tier rows were found in either ' +
               VARIABLES_SHEET_NAME + '!' + GLASS_TIER_RANGE_A1 + ' or ' +
               VARIABLES_SHEET_NAME + '!' + HARDWARE_TIER_RANGE_A1 + '.',
               ui.ButtonSet.OK);
      return;
    }
    itemsAnalysis   = analyzeSheet_(ss, ITEMS_SHEET,   glassTiers, hardwareTiers);
    optionsAnalysis = analyzeSheet_(ss, OPTIONS_SHEET, glassTiers, hardwareTiers);
  } catch (err) {
    ui.alert('Markup Wizard — Error', err.message, ui.ButtonSet.OK);
    return;
  }

  const totalEligible =
    itemsAnalysis.glass.eligibleCount    + itemsAnalysis.hardware.eligibleCount +
    optionsAnalysis.glass.eligibleCount  + optionsAnalysis.hardware.eligibleCount;
  if (totalEligible === 0) {
    ui.alert('Markup Wizard',
             'No eligible Product rows with a positive cost were found.',
             ui.ButtonSet.OK);
    return;
  }

  showWizardModal_(glassTiers, hardwareTiers, itemsAnalysis, optionsAnalysis);
}

function applyMarkupsFromWizard(mode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resolvedMode    = (mode === 'overwriteAll') ? 'overwriteAll' : 'fillMissing';
  const glassTiers      = readTierConfig_(ss, GLASS_TIER_RANGE_A1);
  const hardwareTiers   = readTierConfig_(ss, HARDWARE_TIER_RANGE_A1);
  const itemsAnalysis   = analyzeSheet_(ss, ITEMS_SHEET,   glassTiers, hardwareTiers);
  const optionsAnalysis = analyzeSheet_(ss, OPTIONS_SHEET, glassTiers, hardwareTiers);
  const itemsWritten    = applyMarkups_(ss, ITEMS_SHEET,   itemsAnalysis,   glassTiers, hardwareTiers, resolvedMode);
  const optionsWritten  = applyMarkups_(ss, OPTIONS_SHEET, optionsAnalysis, glassTiers, hardwareTiers, resolvedMode);
  return {
    mode:            resolvedMode,
    itemsGlass:      itemsWritten.glass,
    itemsHardware:   itemsWritten.hardware,
    itemsSkipped:    itemsWritten.skipped,
    optionsGlass:    optionsWritten.glass,
    optionsHardware: optionsWritten.hardware,
    optionsSkipped:  optionsWritten.skipped,
    itemsLabel:      ITEMS_SHEET.displayLabel,
    optionsLabel:    OPTIONS_SHEET.displayLabel,
    total: itemsWritten.glass + itemsWritten.hardware + optionsWritten.glass + optionsWritten.hardware,
    totalSkipped: itemsWritten.skipped + optionsWritten.skipped
  };
}

function showWizardModal_(glassTiers, hardwareTiers, itemsAnalysis, optionsAnalysis) {
  const totalEligible =
    itemsAnalysis.glass.eligibleCount    + itemsAnalysis.hardware.eligibleCount +
    optionsAnalysis.glass.eligibleCount  + optionsAnalysis.hardware.eligibleCount;
  const totalMissing =
    itemsAnalysis.glass.missingCount     + itemsAnalysis.hardware.missingCount +
    optionsAnalysis.glass.missingCount   + optionsAnalysis.hardware.missingCount;

  const glassTiersHtml    = buildTierTableHtml_(glassTiers,    [itemsAnalysis.glass,    optionsAnalysis.glass]);
  const hardwareTiersHtml = buildTierTableHtml_(hardwareTiers, [itemsAnalysis.hardware, optionsAnalysis.hardware]);

  const itemsBlock   = buildSheetBlock_(ITEMS_SHEET.displayLabel,   itemsAnalysis,   glassTiers, hardwareTiers);
  const optionsBlock = buildSheetBlock_(OPTIONS_SHEET.displayLabel, optionsAnalysis, glassTiers, hardwareTiers);

  const html =
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;background:#E3F2FD;border:1px solid #64B5F6;}' +
    '.summary.ok{background:#E8F5E9;border-color:#81C784;}' +
    '.summary.err{background:#FDECEA;border-color:#E57373;}' +
    '.mode{margin:0 0 16px;padding:10px 12px;background:#F5F8FB;border:1px solid #D6DFE8;border-radius:6px;}' +
    '.mode-header{font-weight:600;color:#0B5394;font-size:12px;margin-bottom:6px;}' +
    '.mode-option{display:block;padding:4px 0;cursor:pointer;font-size:12px;}' +
    '.mode-option input{margin-right:8px;vertical-align:middle;}' +
    '.mode-count{color:#6B7C8C;}' +
    '.mode-note{font-size:11px;color:#6B7C8C;margin-top:6px;font-style:italic;}' +
    '.group{margin-bottom:18px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.subgroup{margin:6px 0 10px;}' +
    '.subgroup-header{font-weight:600;color:#1A2733;font-size:12px;margin-bottom:3px;}' +
    '.eligible-count{font-size:12px;margin:4px 0;font-weight:500;}' +
    '.missing-count{font-size:11px;color:#6B7C8C;margin:2px 0 0;}' +
    '.tier-breakdown{font-size:11px;color:#6B7C8C;margin:2px 0 0 12px;line-height:1.55;}' +
    '.complete-line{font-size:12px;padding:4px 0;}' +
    '.skipped-line{font-size:11px;color:#6B7C8C;padding:2px 0 0;font-style:italic;}' +
    '.err-msg{font-size:12px;color:#C62828;padding:8px 0;white-space:pre-wrap;}' +
    '.empty-tiers{font-size:11px;color:#6B7C8C;font-style:italic;padding:4px 0;}' +
    'table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:8px;}' +
    'th{background:#F5F8FB;text-align:left;padding:6px 10px;border-bottom:1px solid #D6DFE8;font-weight:600;color:#0B5394;}' +
    'td{padding:5px 10px;border-bottom:1px solid #ECF0F3;}' +
    '.actions{position:fixed;bottom:0;left:0;right:0;background:white;border-top:1px solid #D6DFE8;padding:12px 16px;display:flex;justify-content:flex-end;gap:8px;}' +
    '.btn{padding:8px 16px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1px solid transparent;font-family:inherit;}' +
    '.btn-secondary{background:white;border-color:#D6DFE8;color:#1A2733;}' +
    '.btn-primary{background:#0B5394;color:white;}' +
    '.btn:disabled{opacity:0.6;cursor:not-allowed;}' +
    '</style></head><body>' +
    '<div id="confirm-view">' +
    '<div class="summary" id="summary-line"></div>' +
    '<div class="mode">' +
    '<div class="mode-header">Mode</div>' +
    '<label class="mode-option"><input type="radio" name="mode" value="fillMissing" checked>' +
    'Fill Missing Only <span class="mode-count">(' + totalMissing + ' row' + (totalMissing === 1 ? '' : 's') + ')</span></label>' +
    '<label class="mode-option"><input type="radio" name="mode" value="overwriteAll">' +
    'Overwrite All Markups <span class="mode-count">(' + totalEligible + ' row' + (totalEligible === 1 ? '' : 's') + ')</span></label>' +
    '<div class="mode-note">Rows with Mark Up = 1 are always preserved (flat-fee passthroughs).</div>' +
    '</div>' +
    '<div class="group"><div class="group-header">' + GLASS_GROUP_LABEL + ' — Tiers</div>' + glassTiersHtml + '</div>' +
    '<div class="group"><div class="group-header">' + HARDWARE_GROUP_LABEL + ' — Tiers</div>' + hardwareTiersHtml + '</div>' +
    itemsBlock + optionsBlock + '</div>' +
    '<div id="complete-view" style="display:none;">' +
    '<div class="summary ok" id="complete-summary"></div>' +
    '<div class="group"><div class="group-header">' + ITEMS_SHEET.displayLabel + '</div>' +
    '<div class="complete-line" id="items-result"></div>' +
    '<div class="skipped-line" id="items-skipped"></div></div>' +
    '<div class="group"><div class="group-header">' + OPTIONS_SHEET.displayLabel + '</div>' +
    '<div class="complete-line" id="options-result"></div>' +
    '<div class="skipped-line" id="options-skipped"></div></div></div>' +
    '<div id="error-view" style="display:none;">' +
    '<div class="summary err">&#9888; An error occurred while applying markups.</div>' +
    '<div class="err-msg" id="error-message"></div></div>' +
    '<div class="actions">' +
    '<button id="cancel-btn" class="btn btn-secondary">Cancel</button>' +
    '<button id="apply-btn" class="btn btn-primary">Apply Markups</button>' +
    '<button id="close-btn" class="btn btn-primary" style="display:none;">Close</button></div>' +
    '<script>' +
    'var fillCount=' + totalMissing + ',overwriteCount=' + totalEligible + ';' +
    'function updateSummary(){' +
    'var mode=document.querySelector("input[name=mode]:checked").value;' +
    'var n=(mode==="fillMissing")?fillCount:overwriteCount;' +
    'var w=(n===1)?"row":"rows";' +
    'var label=(mode==="fillMissing")?"fill missing markups on":"overwrite markups on";' +
    'document.getElementById("summary-line").textContent="Ready to "+label+" "+n+" Product "+w+".";' +
    'document.getElementById("apply-btn").disabled=(n===0);}' +
    'var radios=document.querySelectorAll("input[name=mode]");' +
    'for(var i=0;i<radios.length;i++)radios[i].onchange=updateSummary;' +
    'updateSummary();' +
    'document.getElementById("cancel-btn").onclick=function(){google.script.host.close();};' +
    'document.getElementById("close-btn").onclick=function(){google.script.host.close();};' +
    'document.getElementById("apply-btn").onclick=function(){' +
    'var mode=document.querySelector("input[name=mode]:checked").value;' +
    'var b=document.getElementById("apply-btn");b.disabled=true;b.textContent="Applying\u2026";' +
    'document.getElementById("cancel-btn").disabled=true;' +
    'google.script.run.withSuccessHandler(onOk).withFailureHandler(onErr).applyMarkupsFromWizard(mode);};' +
    'function onOk(r){document.getElementById("confirm-view").style.display="none";' +
    'document.getElementById("complete-view").style.display="block";' +
    'var modeLabel=(r.mode==="fillMissing")?"Filled missing markups on":"Overwrote markups on";' +
    'document.getElementById("complete-summary").innerHTML="&#10003; "+modeLabel+" "+r.total+" row"+(r.total===1?"":"s")+".";' +
    'var iTot=r.itemsGlass+r.itemsHardware,oTot=r.optionsGlass+r.optionsHardware;' +
    'document.getElementById("items-result").textContent=iTot+" row"+(iTot===1?"":"s")+" updated  ("+r.itemsGlass+" glass, "+r.itemsHardware+" hardware)";' +
    'document.getElementById("options-result").textContent=oTot+" row"+(oTot===1?"":"s")+" updated  ("+r.optionsGlass+" glass, "+r.optionsHardware+" hardware)";' +
    'if(r.itemsSkipped>0)document.getElementById("items-skipped").textContent=r.itemsSkipped+" row"+(r.itemsSkipped===1?"":"s")+" preserved (existing markup)";' +
    'if(r.optionsSkipped>0)document.getElementById("options-skipped").textContent=r.optionsSkipped+" row"+(r.optionsSkipped===1?"":"s")+" preserved (existing markup)";' +
    'document.getElementById("cancel-btn").style.display="none";' +
    'document.getElementById("apply-btn").style.display="none";' +
    'document.getElementById("close-btn").style.display="inline-block";}' +
    'function onErr(e){document.getElementById("confirm-view").style.display="none";' +
    'document.getElementById("error-view").style.display="block";' +
    'document.getElementById("error-message").textContent=(e&&e.message)?e.message:String(e);' +
    'document.getElementById("apply-btn").style.display="none";' +
    'document.getElementById("cancel-btn").textContent="Close";' +
    'document.getElementById("cancel-btn").disabled=false;}' +
    '</script></body></html>';

  const output = HtmlService.createHtmlOutput(html).setWidth(560).setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(output, 'Markup Wizard');
}

function buildTierTableHtml_(tierConfig, analysisGroups) {
  if (tierConfig.length === 0) {
    return '<div class="empty-tiers">No tiers configured.</div>';
  }
  let rows = '';
  for (let i = 0; i < tierConfig.length; i++) {
    const t = tierConfig[i];
    let totalRows = 0;
    for (let j = 0; j < analysisGroups.length; j++) {
      totalRows += (analysisGroups[j].tierCounts[i] || 0);
    }
    rows += '<tr><td>Tier ' + t.tierNumber + '</td><td>&le; $' + t.ceiling.toFixed(2) +
            '</td><td>' + t.markup.toFixed(2) + 'x</td><td>' + totalRows + '</td></tr>';
  }
  return '<table><thead><tr><th>Tier</th><th>Ceiling</th><th>Markup</th><th>Eligible Rows</th></tr></thead>' +
         '<tbody>' + rows + '</tbody></table>';
}

function buildSheetBlock_(label, analysis, glassTiers, hardwareTiers) {
  let html = '<div class="group"><div class="group-header">' + label + '</div>';
  html += buildSubgroupHtml_(GLASS_GROUP_LABEL,    analysis.glass,    glassTiers);
  html += buildSubgroupHtml_(HARDWARE_GROUP_LABEL, analysis.hardware, hardwareTiers);
  html += '</div>';
  return html;
}

function buildSubgroupHtml_(label, group, tierConfig) {
  const countWord = group.eligibleCount === 1 ? 'row' : 'rows';
  let html = '<div class="subgroup"><div class="subgroup-header">' + label + '</div>';
  html += '<div class="eligible-count">' + group.eligibleCount + ' eligible ' + countWord + '</div>';
  html += '<div class="missing-count">' + group.missingCount + ' missing markup</div>';
  let tierLines = '';
  for (let i = 0; i < tierConfig.length; i++) {
    if (group.tierCounts[i] > 0) {
      tierLines += 'Tier ' + tierConfig[i].tierNumber + ' (&le; $' + tierConfig[i].ceiling.toFixed(2) +
                   ', ' + tierConfig[i].markup.toFixed(2) + 'x): ' + group.tierCounts[i] + '<br>';
    }
  }
  if (tierLines) html += '<div class="tier-breakdown">' + tierLines + '</div>';
  html += '</div>';
  return html;
}

// =============================================================================
// CONFIG / CLASSIFICATION HELPERS
// =============================================================================

function readTierConfig_(ss, rangeA1) {
  const sheet = ss.getSheetByName(VARIABLES_SHEET_NAME);
  if (!sheet) throw new Error('Sheet "' + VARIABLES_SHEET_NAME + '" not found.');
  const values = sheet.getRange(rangeA1).getValues();
  const tiers = [];
  for (let i = 0; i < values.length; i++) {
    const ceiling = values[i][0];
    const markup  = values[i][2];
    if (typeof ceiling === 'number' && ceiling > 0 && typeof markup === 'number' && markup > 0) {
      tiers.push({ tierNumber: i + 1, ceiling: ceiling, markup: markup });
    }
  }
  tiers.sort(function(a, b) { return a.ceiling - b.ceiling; });
  return tiers;
}

/**
 * Classifies a row for markup-wizard purposes.
 *
 * Returns one of:
 *   'glass'    — row goes through the Glass & Mirrors tier table
 *   'hardware' — row goes through the Hardware/Parts/Other tier table
 *   'skip'     — row is excluded entirely
 *
 * Filter rules:
 *   - If skipTypeCheck is falsy (Items sheet): the Type column must equal
 *     "Product" (case-insensitive, trimmed). Service or empty yields 'skip'.
 *   - If skipTypeCheck is truthy (Item Option Values sheet): the type argument
 *     is ignored. IOV has no Type column because every row is inherently a
 *     product (an option value selection).
 *   - Glass-family categories (GLASS, IGU, MIRROR) without a HARDWARE/OTHER tag
 *     route to 'glass'. Mirrors stay in this pass per current convention.
 *   - Anything else with a non-empty category (HARDWARE-tagged rows, plus
 *     standalone SCREENS/DOORS/SKYLIGHTS/WINDOWS/etc.) routes to 'hardware'.
 */
function classifyForMarkup_(category, type, skipTypeCheck) {
  if (!skipTypeCheck) {
    if (!type) return 'skip';
    const trimmedType = String(type).trim().toUpperCase();
    if (trimmedType !== 'PRODUCT') return 'skip';
  }
  if (!category) return 'skip';
  const upper = String(category).toUpperCase();
  const isGlassFamily = upper.indexOf('GLASS')!==-1 || upper.indexOf('MIRROR')!==-1 || upper.indexOf('IGU')!==-1;
  const isHardware    = upper.indexOf('HARDWARE')!==-1 || upper.indexOf('OTHER')!==-1;
  if (isGlassFamily && !isHardware) return 'glass';
  return 'hardware';
}

function lookupTierIndex_(cost, tiers) {
  for (let i = 0; i < tiers.length; i++) {
    if (cost <= tiers[i].ceiling) return i;
  }
  return tiers.length - 1;
}

// =============================================================================
// SHEET ANALYSIS / APPLY
// =============================================================================

function makeAnalysisGroup_(n) {
  return {
    eligibleRows: [],
    eligibleCount: 0,
    missingCount: 0,
    tierCounts: new Array(n).fill(0),
    missingTierCounts: new Array(n).fill(0)
  };
}

function analyzeSheet_(ss, sheetDef, glassTiers, hardwareTiers) {
  const sheet = ss.getSheetByName(sheetDef.name);
  if (!sheet) throw new Error('Sheet "' + sheetDef.name + '" not found.');
  const lastRow = sheet.getLastRow();
  const result = {
    glass:    makeAnalysisGroup_(glassTiers.length),
    hardware: makeAnalysisGroup_(hardwareTiers.length)
  };
  if (lastRow <= sheetDef.headerRow) return result;

  const cols = [sheetDef.categoryCol, sheetDef.costCol, sheetDef.markupCol];
  if (sheetDef.typeCol) cols.push(sheetDef.typeCol);
  const minCol = Math.min.apply(null, cols);
  const maxCol = Math.max.apply(null, cols);
  const numRows = lastRow - sheetDef.headerRow;
  const values = sheet.getRange(sheetDef.headerRow + 1, minCol, numRows, maxCol - minCol + 1).getValues();

  for (let i = 0; i < values.length; i++) {
    const row      = values[i];
    const category = row[sheetDef.categoryCol - minCol];
    const type     = sheetDef.typeCol ? row[sheetDef.typeCol - minCol] : null;
    const cost     = row[sheetDef.costCol - minCol];
    const markup   = row[sheetDef.markupCol - minCol];
    if (typeof cost !== 'number' || cost <= 0) continue;
    // Flat-fee guard: Markup=1 means the row is an intentional pass-through
    // (Service surcharge, etc.). Always preserved, both modes.
    if (typeof markup === 'number' && markup === 1) continue;
    const klass = classifyForMarkup_(category, type, !sheetDef.typeCol);
    if (klass === 'skip') continue;
    const tiers = (klass === 'glass') ? glassTiers : hardwareTiers;
    if (tiers.length === 0) continue;
    const sheetRow = sheetDef.headerRow + 1 + i;
    const tierIdx  = lookupTierIndex_(cost, tiers);
    const hasExistingMarkup = (typeof markup === 'number' && markup > 0);
    result[klass].eligibleRows.push({
      sheetRow: sheetRow,
      cost: cost,
      hasExistingMarkup: hasExistingMarkup
    });
    result[klass].tierCounts[tierIdx]++;
    if (!hasExistingMarkup) {
      result[klass].missingTierCounts[tierIdx]++;
      result[klass].missingCount++;
    }
  }
  result.glass.eligibleCount    = result.glass.eligibleRows.length;
  result.hardware.eligibleCount = result.hardware.eligibleRows.length;
  return result;
}

function applyMarkups_(ss, sheetDef, analysis, glassTiers, hardwareTiers, mode) {
  const fillMissingOnly = (mode === 'fillMissing');
  const written = { glass: 0, hardware: 0, skipped: 0 };
  const targets = [];

  function gatherFromGroup(rows, tiers, klass) {
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      if (fillMissingOnly && r.hasExistingMarkup) {
        written.skipped++;
        continue;
      }
      targets.push({
        sheetRow: r.sheetRow,
        cost:     r.cost,
        tiers:    tiers,
        klass:    klass
      });
    }
  }

  gatherFromGroup(analysis.glass.eligibleRows,    glassTiers,    'glass');
  gatherFromGroup(analysis.hardware.eligibleRows, hardwareTiers, 'hardware');

  if (targets.length === 0) return written;

  const sheet = ss.getSheetByName(sheetDef.name);
  let minRow = Infinity, maxRow = -Infinity;
  for (let i = 0; i < targets.length; i++) {
    const r = targets[i].sheetRow;
    if (r < minRow) minRow = r;
    if (r > maxRow) maxRow = r;
  }
  const numRows  = maxRow - minRow + 1;
  const range    = sheet.getRange(minRow, sheetDef.markupCol, numRows, 1);
  const existing = range.getValues();
  const targetMap = {};
  for (let i = 0; i < targets.length; i++) {
    targetMap[targets[i].sheetRow] = targets[i];
  }
  for (let i = 0; i < numRows; i++) {
    const sheetRow = minRow + i;
    if (Object.prototype.hasOwnProperty.call(targetMap, sheetRow)) {
      const t = targetMap[sheetRow];
      const tierIdx = lookupTierIndex_(t.cost, t.tiers);
      existing[i][0] = t.tiers[tierIdx].markup;
      written[t.klass]++;
    }
  }
  range.setValues(existing);
  return written;
}

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

  const total =
    itemsAnalysis.glass.eligibleCount    + itemsAnalysis.hardware.eligibleCount +
    optionsAnalysis.glass.eligibleCount  + optionsAnalysis.hardware.eligibleCount;
  if (total === 0) {
    ui.alert('Markup Wizard',
             'No eligible Product rows with a positive cost were found.',
             ui.ButtonSet.OK);
    return;
  }

  showWizardModal_(glassTiers, hardwareTiers, itemsAnalysis, optionsAnalysis);
}

function applyMarkupsFromWizard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const glassTiers      = readTierConfig_(ss, GLASS_TIER_RANGE_A1);
  const hardwareTiers   = readTierConfig_(ss, HARDWARE_TIER_RANGE_A1);
  const itemsAnalysis   = analyzeSheet_(ss, ITEMS_SHEET,   glassTiers, hardwareTiers);
  const optionsAnalysis = analyzeSheet_(ss, OPTIONS_SHEET, glassTiers, hardwareTiers);
  const itemsWritten    = applyMarkups_(ss, ITEMS_SHEET,   itemsAnalysis,   glassTiers, hardwareTiers);
  const optionsWritten  = applyMarkups_(ss, OPTIONS_SHEET, optionsAnalysis, glassTiers, hardwareTiers);
  return {
    itemsGlass:      itemsWritten.glass,
    itemsHardware:   itemsWritten.hardware,
    optionsGlass:    optionsWritten.glass,
    optionsHardware: optionsWritten.hardware,
    itemsLabel:      ITEMS_SHEET.displayLabel,
    optionsLabel:    OPTIONS_SHEET.displayLabel,
    total: itemsWritten.glass + itemsWritten.hardware + optionsWritten.glass + optionsWritten.hardware
  };
}

function showWizardModal_(glassTiers, hardwareTiers, itemsAnalysis, optionsAnalysis) {
  const totalEligible =
    itemsAnalysis.glass.eligibleCount    + itemsAnalysis.hardware.eligibleCount +
    optionsAnalysis.glass.eligibleCount  + optionsAnalysis.hardware.eligibleCount;

  const glassTiersHtml    = buildTierTableHtml_(glassTiers,    [itemsAnalysis.glass,    optionsAnalysis.glass]);
  const hardwareTiersHtml = buildTierTableHtml_(hardwareTiers, [itemsAnalysis.hardware, optionsAnalysis.hardware]);

  const itemsBlock   = buildSheetBlock_(ITEMS_SHEET.displayLabel,   itemsAnalysis,   glassTiers, hardwareTiers);
  const optionsBlock = buildSheetBlock_(OPTIONS_SHEET.displayLabel, optionsAnalysis, glassTiers, hardwareTiers);

  const rowWord = totalEligible === 1 ? 'row' : 'rows';
  const summary = 'Ready to apply markups to ' + totalEligible + ' Product ' + rowWord +
                  ' across ' + ITEMS_SHEET.displayLabel + ' and ' + OPTIONS_SHEET.displayLabel + '.';

  const html =
    '<!DOCTYPE html><html><head><base target="_top"><style>' +
    'body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;margin:0;padding:16px 16px 72px;color:#1A2733;font-size:13px;}' +
    '.summary{border-radius:6px;padding:10px 12px;margin-bottom:14px;font-weight:600;background:#E3F2FD;border:1px solid #64B5F6;}' +
    '.summary.ok{background:#E8F5E9;border-color:#81C784;}' +
    '.summary.err{background:#FDECEA;border-color:#E57373;}' +
    '.group{margin-bottom:18px;}' +
    '.group-header{font-weight:600;color:#0B5394;font-size:13px;margin-bottom:6px;padding-bottom:4px;border-bottom:2px solid #0B5394;}' +
    '.subgroup{margin:6px 0 10px;}' +
    '.subgroup-header{font-weight:600;color:#1A2733;font-size:12px;margin-bottom:3px;}' +
    '.eligible-count{font-size:12px;margin:4px 0;font-weight:500;}' +
    '.tier-breakdown{font-size:11px;color:#6B7C8C;margin:2px 0 0 12px;line-height:1.55;}' +
    '.complete-line{font-size:12px;padding:4px 0;}' +
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
    '<div class="summary">' + summary + '</div>' +
    '<div class="group"><div class="group-header">' + GLASS_GROUP_LABEL + ' — Tiers</div>' + glassTiersHtml + '</div>' +
    '<div class="group"><div class="group-header">' + HARDWARE_GROUP_LABEL + ' — Tiers</div>' + hardwareTiersHtml + '</div>' +
    itemsBlock + optionsBlock + '</div>' +
    '<div id="complete-view" style="display:none;">' +
    '<div class="summary ok" id="complete-summary"></div>' +
    '<div class="group"><div class="group-header">' + ITEMS_SHEET.displayLabel + '</div>' +
    '<div class="complete-line" id="items-result"></div></div>' +
    '<div class="group"><div class="group-header">' + OPTIONS_SHEET.displayLabel + '</div>' +
    '<div class="complete-line" id="options-result"></div></div></div>' +
    '<div id="error-view" style="display:none;">' +
    '<div class="summary err">&#9888; An error occurred while applying markups.</div>' +
    '<div class="err-msg" id="error-message"></div></div>' +
    '<div class="actions">' +
    '<button id="cancel-btn" class="btn btn-secondary">Cancel</button>' +
    '<button id="apply-btn" class="btn btn-primary">Apply Markups</button>' +
    '<button id="close-btn" class="btn btn-primary" style="display:none;">Close</button></div>' +
    '<script>' +
    'document.getElementById("cancel-btn").onclick=function(){google.script.host.close();};' +
    'document.getElementById("close-btn").onclick=function(){google.script.host.close();};' +
    'document.getElementById("apply-btn").onclick=function(){' +
    'var b=document.getElementById("apply-btn");b.disabled=true;b.textContent="Applying\u2026";' +
    'document.getElementById("cancel-btn").disabled=true;' +
    'google.script.run.withSuccessHandler(onOk).withFailureHandler(onErr).applyMarkupsFromWizard();};' +
    'function onOk(r){document.getElementById("confirm-view").style.display="none";' +
    'document.getElementById("complete-view").style.display="block";' +
    'document.getElementById("complete-summary").innerHTML="&#10003; Updated "+r.total+" row"+(r.total===1?"":"s")+" successfully.";' +
    'var iTot=r.itemsGlass+r.itemsHardware,oTot=r.optionsGlass+r.optionsHardware;' +
    'document.getElementById("items-result").textContent=iTot+" row"+(iTot===1?"":"s")+" updated  ("+r.itemsGlass+" glass, "+r.itemsHardware+" hardware)";' +
    'document.getElementById("options-result").textContent=oTot+" row"+(oTot===1?"":"s")+" updated  ("+r.optionsGlass+" glass, "+r.optionsHardware+" hardware)";' +
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

  const output = HtmlService.createHtmlOutput(html).setWidth(540).setHeight(720);
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

function analyzeSheet_(ss, sheetDef, glassTiers, hardwareTiers) {
  const sheet = ss.getSheetByName(sheetDef.name);
  if (!sheet) throw new Error('Sheet "' + sheetDef.name + '" not found.');
  const lastRow = sheet.getLastRow();
  const result = {
    glass:    { eligibleRows: [], eligibleCount: 0, tierCounts: new Array(glassTiers.length).fill(0) },
    hardware: { eligibleRows: [], eligibleCount: 0, tierCounts: new Array(hardwareTiers.length).fill(0) }
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
    // (Service surcharge, etc.). Don't overwrite with a tier markup.
    if (typeof markup === 'number' && markup === 1) continue;
    const klass = classifyForMarkup_(category, type, !sheetDef.typeCol);
    if (klass === 'skip') continue;
    const tiers = (klass === 'glass') ? glassTiers : hardwareTiers;
    if (tiers.length === 0) continue;
    const sheetRow = sheetDef.headerRow + 1 + i;
    const tierIdx  = lookupTierIndex_(cost, tiers);
    result[klass].eligibleRows.push({ sheetRow: sheetRow, cost: cost });
    result[klass].tierCounts[tierIdx]++;
  }
  result.glass.eligibleCount    = result.glass.eligibleRows.length;
  result.hardware.eligibleCount = result.hardware.eligibleRows.length;
  return result;
}

function applyMarkups_(ss, sheetDef, analysis, glassTiers, hardwareTiers) {
  const written = { glass: 0, hardware: 0 };
  const allEligible = [];
  for (let i = 0; i < analysis.glass.eligibleRows.length; i++) {
    allEligible.push({
      sheetRow: analysis.glass.eligibleRows[i].sheetRow,
      cost:     analysis.glass.eligibleRows[i].cost,
      tiers:    glassTiers,
      klass:    'glass'
    });
  }
  for (let i = 0; i < analysis.hardware.eligibleRows.length; i++) {
    allEligible.push({
      sheetRow: analysis.hardware.eligibleRows[i].sheetRow,
      cost:     analysis.hardware.eligibleRows[i].cost,
      tiers:    hardwareTiers,
      klass:    'hardware'
    });
  }
  if (allEligible.length === 0) return written;

  const sheet = ss.getSheetByName(sheetDef.name);
  let minRow = Infinity, maxRow = -Infinity;
  for (let i = 0; i < allEligible.length; i++) {
    const r = allEligible[i].sheetRow;
    if (r < minRow) minRow = r;
    if (r > maxRow) maxRow = r;
  }
  const numRows  = maxRow - minRow + 1;
  const range    = sheet.getRange(minRow, sheetDef.markupCol, numRows, 1);
  const existing = range.getValues();
  const eligibleMap = {};
  for (let i = 0; i < allEligible.length; i++) {
    eligibleMap[allEligible[i].sheetRow] = allEligible[i];
  }
  for (let i = 0; i < numRows; i++) {
    const sheetRow = minRow + i;
    if (Object.prototype.hasOwnProperty.call(eligibleMap, sheetRow)) {
      const e = eligibleMap[sheetRow];
      const tierIdx = lookupTierIndex_(e.cost, e.tiers);
      existing[i][0] = e.tiers[tierIdx].markup;
      written[e.klass]++;
    }
  }
  range.setValues(existing);
  return written;
}

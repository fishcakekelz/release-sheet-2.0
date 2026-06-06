/**********************
  Status Dropdowns + Row Coloring

  Applies dropdown data validation and entire-row background coloring driven
  by the "Deployment Status" and "Validation Status" columns. Designed to run
  on either the ReleaseSheet or the Summary sheet — both are looked up by
  header name, so column letter doesn't matter.

  Refresh-safety: both columns are already in PROTECTED_COLUMNS in
  constants.gs and merged onto refreshed rows by mergeProtectedColumnsOntoRows_,
  so refreshes do NOT overwrite the values a PM has set.

  Conditional formatting precedence: Validation Status rules are added first
  so they win when both columns are populated (validation is the later step
  in the workflow). To change priority, swap the call order in
  applyStatusFormattingToSheet_.
**********************/

/* =================== Configuration =================== */

const DEPLOYMENT_STATUS_COLUMN = "Deployment Status";
const VALIDATION_STATUS_COLUMN = "Validation Status";

/**
 * Allowed values for the Deployment Status dropdown, in display order.
 * Each value also maps to the row background color (hex) applied when that
 * value is selected. Edit colors freely — re-running the setup function
 * picks up the new colors and replaces the matching rules in place.
 */
const DEPLOYMENT_STATUS_OPTIONS = {
  "Not Started":              "#f3f3f3", // light gray
  "Being Deployed":           "#fff2cc", // light yellow
  "Completed - successfully": "#d9ead3", // light green
  "Failed Deployment":        "#f4cccc", // light red
  "Failed - Rolled Back":     "#fce5cd", // light orange
  "Failed - Not Rolled Back": "#ea9999", // dark red
  "Cancelled":                "#cfe2f3", // light blue-gray
  "Rescheduled":              "#d9d2e9"  // light purple
};

const VALIDATION_STATUS_OPTIONS = {
  "Not Started":                  "#f3f3f3",
  "Being Validated":              "#fff2cc",
  "Cancelled":                    "#cfe2f3",
  "Completed - Successfully":     "#d9ead3",
  "Deferred":                     "#ffe599", // saturated yellow
  "Partial Success - New Defect": "#fce5cd",
  "Redeploy/Prod Val - Complete": "#b6d7a8", // medium green
  "Failed Deployment":            "#f4cccc",
  "Failed - Not Rolled Back":     "#ea9999",
  "Failed - Rolled Back":         "#fce5cd",
  "Rescheduled":                  "#d9d2e9",
  "Regression Only":              "#a4c2f4"  // medium blue
};

/* =================== Public entry points (menu / wired in) =================== */

/** Menu wrapper: apply dropdowns + row coloring to ReleaseSheet. */
function applyStatusFormattingToReleaseSheet() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(S_RELEASE);
  if (!sheet) {
    ui.alert("ReleaseSheet not found.");
    return;
  }
  applyStatusFormattingToSheet_(sheet, cfg.sheetHeaderRow);
  ui.alert(
    "Applied dropdowns + row coloring to ReleaseSheet for \"" +
      DEPLOYMENT_STATUS_COLUMN + "\" and \"" + VALIDATION_STATUS_COLUMN + "\"."
  );
}

/** Menu wrapper: apply dropdowns + row coloring to Summary. */
function applyStatusFormattingToSummary() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) {
    ui.alert("Summary sheet not found.");
    return;
  }
  applyStatusFormattingToSheet_(sheet, 1);
  ui.alert(
    "Applied dropdowns + row coloring to Summary for \"" +
      DEPLOYMENT_STATUS_COLUMN + "\" and \"" + VALIDATION_STATUS_COLUMN + "\"."
  );
}

/* =================== Core =================== */

/**
 * Applies status dropdowns + row-wide background coloring to a sheet.
 * Idempotent: re-running cleans up rules with the same formulas before
 * adding the new ones, so the user's other conditional-format rules are
 * preserved.
 *
 * Fast-path: by default, when this sheet already has our conditional-format
 * rules in place (detected by matching one of our status values inside a
 * custom-formula rule), we return immediately. This avoids the expensive
 * setConditionalFormatRules call on every Refresh. Pass `{ force: true }` to
 * rebuild the rules — useful when colors in *_OPTIONS were edited.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} headerRow 1-based; data starts at headerRow + 1
 * @param {{ force?: boolean }} [options]
 */
function applyStatusFormattingToSheet_(sheet, headerRow, options) {
  if (!sheet || !headerRow || headerRow < 1) return;
  options = options || {};

  const { map } = getHeaderIndexMapSafe(sheet, headerRow);
  const depCol = map[DEPLOYMENT_STATUS_COLUMN];
  const valCol = map[VALIDATION_STATUS_COLUMN];

  if (depCol === undefined && valCol === undefined) {
    Logger.log(
      'applyStatusFormattingToSheet_: neither "' + DEPLOYMENT_STATUS_COLUMN +
        '" nor "' + VALIDATION_STATUS_COLUMN + '" found on "' + sheet.getName() + '".'
    );
    return;
  }

  if (!options.force && hasStatusFormattingAlready_(sheet)) {
    return;
  }

  const dataFirstRow = headerRow + 1;
  const lastRow = Math.max(sheet.getMaxRows(), dataFirstRow);
  const lastCol = Math.max(sheet.getMaxColumns(), 1);

  // ---- 1. Dropdown data validation ----
  if (depCol !== undefined) {
    applyDropdown_(sheet, dataFirstRow, depCol + 1, lastRow, Object.keys(DEPLOYMENT_STATUS_OPTIONS));
  }
  if (valCol !== undefined) {
    applyDropdown_(sheet, dataFirstRow, valCol + 1, lastRow, Object.keys(VALIDATION_STATUS_OPTIONS));
  }

  // ---- 2. Conditional formatting (entire row) ----
  const rowRange = sheet.getRange(dataFirstRow, 1, lastRow - dataFirstRow + 1, lastCol);

  // Build the set of formulas we own, so we can drop any prior matching rules
  // before we re-add fresh ones (handles re-runs and color edits cleanly).
  const ourFormulas = new Set();
  const newRules = [];

  // Validation Status FIRST so its colors take precedence over Deployment Status.
  if (valCol !== undefined) {
    const valColLetter = colNumberToLetter_(valCol + 1);
    Object.keys(VALIDATION_STATUS_OPTIONS).forEach(value => {
      const formula = '=$' + valColLetter + dataFirstRow + '="' + value + '"';
      ourFormulas.add(formula);
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula)
          .setBackground(VALIDATION_STATUS_OPTIONS[value])
          .setRanges([rowRange])
          .build()
      );
    });
  }

  if (depCol !== undefined) {
    const depColLetter = colNumberToLetter_(depCol + 1);
    Object.keys(DEPLOYMENT_STATUS_OPTIONS).forEach(value => {
      const formula = '=$' + depColLetter + dataFirstRow + '="' + value + '"';
      ourFormulas.add(formula);
      newRules.push(
        SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(formula)
          .setBackground(DEPLOYMENT_STATUS_OPTIONS[value])
          .setRanges([rowRange])
          .build()
      );
    });
  }

  // Preserve other conditional format rules; drop only ones whose custom
  // formula matches one of ours (so re-runs don't duplicate rules).
  const existing = sheet.getConditionalFormatRules();
  const kept = existing.filter(rule => {
    const cond = rule.getBooleanCondition && rule.getBooleanCondition();
    if (!cond) return true; // gradient or unknown rule type — keep it
    const criteria = cond.getCriteriaType();
    if (criteria !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) return true;
    const args = cond.getCriteriaValues();
    if (!args || !args.length) return true;
    return !ourFormulas.has(args[0]);
  });

  // Order matters: ours first → status rules win for the row background.
  sheet.setConditionalFormatRules(newRules.concat(kept));
}

/**
 * Quick check: does this sheet already have our status-formatting rules?
 * We probe by looking for any custom-formula rule whose formula text mentions
 * one of our distinctive status values (e.g. "Being Deployed",
 * "Being Validated"). Cheap — a single getConditionalFormatRules call and
 * an in-memory scan.
 */
function hasStatusFormattingAlready_(sheet) {
  // Probe values are status options that only exist in OUR rules; a normal
  // user-authored CF rule is very unlikely to reference these strings.
  const probe = ['"Being Deployed"', '"Being Validated"', '"Redeploy/Prod Val - Complete"'];
  let rules;
  try {
    rules = sheet.getConditionalFormatRules();
  } catch (e) {
    return false;
  }
  for (let i = 0; i < rules.length; i++) {
    const rule = rules[i];
    const cond = rule.getBooleanCondition && rule.getBooleanCondition();
    if (!cond) continue;
    if (cond.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) continue;
    const args = cond.getCriteriaValues();
    if (!args || !args.length) continue;
    const formula = args[0] || "";
    for (let p = 0; p < probe.length; p++) {
      if (formula.indexOf(probe[p]) >= 0) return true;
    }
  }
  return false;
}

/* =================== Helpers =================== */

/**
 * Sets a "value in list" dropdown on a single column over a range of rows.
 * setAllowInvalid(false) means typing something not in the list shows an error.
 */
function applyDropdown_(sheet, firstRow, col1based, lastRow, allowedValues) {
  if (lastRow < firstRow) return;
  const range = sheet.getRange(firstRow, col1based, lastRow - firstRow + 1, 1);
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInList(allowedValues, true)
    .setAllowInvalid(false)
    .build();
  range.setDataValidation(validation);
}

/** 1 → "A", 27 → "AA", etc. */
function colNumberToLetter_(n) {
  let s = '';
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

/**********************
  Manual Backup Runbook

  Creates three tabs that, together, let any PM produce a Release Sheet
  using only their browser + Google Sheets — no Apps Script execution
  needed at the moment of use. Run setupBackupRunbook() once (while
  Apps Script is healthy) to install/refresh all three tabs.

  Tabs created:
    1. TEMPLATE         — pristine empty Release Sheet structure (hidden)
    2. BACKUP BUILD     — paste Jira CSV + SNOW CSV, formulas merge into Release Sheet shape
    3. BACKUP PROCEDURE — step-by-step runbook embedded directly in the workbook

  Why this exists: if Apps Script is broken, quota-exhausted, or
  otherwise unavailable on a release morning, a PM with browser +
  Sheets access can still produce a working Release Sheet by following
  the runbook in BACKUP PROCEDURE.

  External prerequisites the runbook references (one-time setup):
    - A saved Jira filter that runs the JQL produced by buildJqlForDate
    - A saved ServiceNow change_request list view with the columns
      fetchSnowAndBlockersInParallel_ requests
**********************/

/* =================== Tab name constants =================== */
const BACKUP_TEMPLATE_TAB = "TEMPLATE";
const BACKUP_BUILD_TAB = "BACKUP BUILD";
const BACKUP_PROCEDURE_TAB = "BACKUP PROCEDURE";

/* =================== BACKUP BUILD layout =================== */
// 200 data rows is enough for any realistic release (the largest weeks
// land at ~100; 200 leaves headroom and keeps formula recalculation cheap).
const BACKUP_BUILD_DATA_ROWS = 200;
// Section 1 (Jira CSV paste): columns A..Y (25 cols of headroom for
// custom fields). User pastes including the header row at row 1 — the
// pre-filled "expected" headers we put there get replaced by the user's
// actual headers, and the merge formulas use header-lookup so column
// order doesn't matter.
const BACKUP_BUILD_JIRA_FIRST_COL = 1;     // A
const BACKUP_BUILD_JIRA_LAST_COL = 25;     // Y
// Section 2 (SNOW CSV paste): columns AA..AY. One blank spacer column
// (Z) so paste boundaries are obvious.
const BACKUP_BUILD_SNOW_FIRST_COL = 27;    // AA
const BACKUP_BUILD_SNOW_LAST_COL = 51;     // AY
// Section 3 (output rows): starts at column BA. Width = number of
// canonical Release Sheet columns from STOP_RELEASE_SHEET_HEADER_ROW.
const BACKUP_BUILD_OUTPUT_FIRST_COL = 53;  // BA

/* =================== Suggested Jira / SNOW CSV headers =================== */
// These are pre-filled into the paste-area header row so PMs can see
// what the merge formulas expect. The formulas use header-lookup so the
// user's actual paste can have columns in any order — what matters is
// that the header TEXT matches one of the entries we look for.

const BACKUP_BUILD_JIRA_EXPECTED_HEADERS = [
  "Issue key",
  "Summary",
  "Issue Type",
  "Status",
  "Resolution",
  "Assignee",
  "Priority",
  "Custom field (Production Release Date)",
  "Custom field (ServiceNow Related Ticket)",
  "Custom field (Scrum Team)",
  "Custom field (Workstream)",
  "Custom field (Approvals)",
  "Custom field (Channel)",
  "Custom field (Acceptance Criteria)",
  "Description",
  "Custom field (Last Comment)",
  "Custom field (QA Artifacts Link)",
  "Custom field (UAT Artifacts Link)",
  "Custom field (Severity)",
  "", "", "", "", "", "" // spare columns
];

const BACKUP_BUILD_SNOW_EXPECTED_HEADERS = [
  "Number",
  "Start date",
  "End date",
  "Change owner",
  "Assigned to",
  "Change plan",
  "Backout plan",
  "Validation status",
  "Validation owner",
  "Validation plan",
  "Configuration item",
  "Type",
  "Opened",
  "Affected groups",
  "Affected locations",
  "Assignment group",
  "Impact description",
  "Description",
  "Department",
  "Division",
  "Company",
  "", "", "", "" // spare columns
];

/* =================== Output column source map =================== */
// For each Release Sheet column (in STOP_RELEASE_SHEET_HEADER_ROW
// order), describes how the BACKUP BUILD merge formula populates it.
//
// Types:
//   manual                    — leave blank; PM fills in directly (e.g. dropdowns)
//   jira-direct               — INDEX/MATCH on the Jira section by header text
//   jira-direct-with-fallback — try jiraHeader, fall back to fallbackHeader
//   jira-hyperlink            — Issue key wrapped as HYPERLINK to Jira browse URL
//   jira-cr-extract           — REGEXEXTRACT first CHG-prefix from a Jira CSV cell
//   snow-xlookup              — XLOOKUP into SNOW section by CR # for snowHeader
//   snow-xlookup-fallback     — try snowHeader, fall back to fallbackHeader
//   matrix-lookup             — VLOOKUP "Release & Approval Matrix" by Scrum Team
//                               (mirrors buildTechLeadLookupFromMatrix_)
const BACKUP_OUTPUT_SOURCES = {
  "Date":                       { type: "manual" },
  "Pre-Release Check Summary":  { type: "manual" }, // filled in by the Release Manager pre-release
  "JIRA Item":                  { type: "jira-hyperlink", jiraHeader: "Issue key" },
  "CR #":                       { type: "jira-cr-extract", jiraHeader: "Custom field (ServiceNow Related Ticket)" },
  "Start Time":                 { type: "snow-xlookup", snowHeader: "Start date" },
  "Validation Owner":           { type: "snow-xlookup", snowHeader: "Validation owner" },
  "Scrum Team":                 { type: "jira-direct", jiraHeader: "Custom field (Scrum Team)" },
  "Comments":                   { type: "manual" },
  "Approvals":                  { type: "jira-direct", jiraHeader: "Custom field (Approvals)" },
  "Dark Deployment":            { type: "manual" },
  "Late Addition":              { type: "manual" },
  "Acceptance Criteria":        { type: "jira-direct", jiraHeader: "Custom field (Acceptance Criteria)" },
  "Last Updated JIRA comment":  { type: "jira-direct", jiraHeader: "Custom field (Last Comment)" },
  "Workstream":                 { type: "jira-direct", jiraHeader: "Custom field (Workstream)" },
  "Channel":                    { type: "jira-direct", jiraHeader: "Custom field (Channel)" },
  "Jira Prod Release Date":     { type: "jira-direct", jiraHeader: "Custom field (Production Release Date)" },
  "Blocker Status":             { type: "manual" }, // issuelinks not in CSV export
  "Jira Status":                { type: "jira-direct", jiraHeader: "Status" },
  "Jira Resolution":            { type: "jira-direct", jiraHeader: "Resolution" },
  "Issue Type":                 { type: "jira-direct", jiraHeader: "Issue Type" },
  "Late Approval":              { type: "manual" },
  "JIRA Assignee":              { type: "jira-direct", jiraHeader: "Assignee" },
  "Related Issues":             { type: "manual" }, // issuelinks not in CSV export
  "Deployment Status":          { type: "manual" }, // dropdown
  "Validation Status":          { type: "manual" }, // dropdown
  "Release Type":               { type: "manual" },
  "Change Assignee":            { type: "snow-xlookup-fallback", snowHeader: "Change owner", fallbackHeader: "Assigned to" },
  "CHG Type":                   { type: "snow-xlookup", snowHeader: "Type" },
  "CTASK List":                 { type: "manual" }, // requires separate change_task export
  "Impact Description":         { type: "snow-xlookup", snowHeader: "Impact description" },
  "Change Plan":                { type: "snow-xlookup", snowHeader: "Change plan" },
  "Rollback Plan":              { type: "snow-xlookup", snowHeader: "Backout plan" },
  "JIRA Summary":               { type: "jira-direct", jiraHeader: "Summary" },
  "JIRA Description":           { type: "jira-direct", jiraHeader: "Description" },
  "Validation Plan":            { type: "snow-xlookup", snowHeader: "Validation plan" },
  "Deploy Start":               { type: "manual" }, // filled by Release Manager during the release
  "Deploy End":                 { type: "manual" }, // filled by Release Manager during the release
  "Impacted CI":                { type: "snow-xlookup", snowHeader: "Configuration item" },
  "Date Submitted":             { type: "snow-xlookup", snowHeader: "Opened" },
  "Affected Groups":            { type: "snow-xlookup", snowHeader: "Affected groups" },
  "Affected Locations":         { type: "snow-xlookup", snowHeader: "Affected locations" },
  "Assignment Group":           { type: "snow-xlookup", snowHeader: "Assignment group" },
  "Tech Lead / Sr Tech Lead":   { type: "matrix-lookup" },
  "Division":                   { type: "snow-xlookup-fallback", snowHeader: "Division", fallbackHeader: "Company" },
  "QA Artifacts Link":          { type: "jira-direct", jiraHeader: "Custom field (QA Artifacts Link)" },
  "UAT Artifacts Link":         { type: "jira-direct", jiraHeader: "Custom field (UAT Artifacts Link)" },
  "Severity":                   { type: "jira-direct-with-fallback", jiraHeader: "Custom field (Severity)", fallbackHeader: "Priority" },
  "Validation Start":           { type: "manual" },
  "Validation End":             { type: "manual" },
  "Deploy Duration":            { type: "manual" },
  "Release Duration":           { type: "manual" },
  "JIRA Approval Comment Sent": { type: "manual" }
};

/* =================== Public entry point =================== */

/**
 * Creates or refreshes the three backup tabs. Idempotent — safe to
 * re-run any time. Existing tabs with these names are wiped and
 * rewritten so a stale runbook can always be refreshed in one click.
 */
function setupBackupRunbook() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  try {
    Logger.log("[BackupRunbook] start");
    const t0 = Date.now();

    createOrResetTemplateTab_(ss);
    Logger.log("[BackupRunbook] TEMPLATE tab ready");

    createOrResetBackupBuildTab_(ss);
    Logger.log("[BackupRunbook] BACKUP BUILD tab ready");

    createOrResetBackupProcedureTab_(ss);
    Logger.log("[BackupRunbook] BACKUP PROCEDURE tab ready");

    SpreadsheetApp.flush();
    Logger.log("[BackupRunbook] complete in " + (Date.now() - t0) + "ms");

    ui.alert(
      "Backup runbook installed.\n\n" +
        "Tabs created/refreshed:\n" +
        "  • " + BACKUP_TEMPLATE_TAB + " (hidden)\n" +
        "  • " + BACKUP_BUILD_TAB + "\n" +
        "  • " + BACKUP_PROCEDURE_TAB + "\n\n" +
        "Open '" + BACKUP_PROCEDURE_TAB + "' for the step-by-step runbook " +
        "(Section A: Manual Create, Section B: Manual Close).\n\n" +
        "Don't forget the one-time external prep — Section C lists the " +
        "saved Jira filter + ServiceNow list view that need to be " +
        "created and recorded in the workbook before the backup is " +
        "usable on release morning."
    );
  } catch (e) {
    Logger.log("[BackupRunbook] FAILED: " + (e && e.message ? e.message : e));
    ui.alert("setupBackupRunbook failed: " + (e && e.message ? e.message : e));
    Logger.log(e);
  }
}

/* =================== Tab: TEMPLATE =================== */
/**
 * Creates (or wipes and rewrites) the hidden TEMPLATE tab. This is a
 * pristine, empty Release Sheet structure that PMs duplicate when the
 * live ReleaseSheet is corrupted or needs to be rebuilt from scratch.
 */
function createOrResetTemplateTab_(ss) {
  const sheet = getOrCreateSheet_(ss, BACKUP_TEMPLATE_TAB);
  resetSheet_(sheet);

  const headers = STOP_RELEASE_SHEET_HEADER_ROW;
  // Title banner (matches the live Release Sheet's row 1 layout: a
  // merged cell across all canonical columns, holding the release
  // date.)
  sheet.getRange(1, 1).setValue("<release date>");
  sheet
    .getRange(1, 1, 1, headers.length)
    .merge()
    .setBackground("#fff2cc")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  // Header row at row 3 (matches Config.sheetHeaderRow=3 convention; if
  // the live workbook uses a different header row, the PM can adjust
  // when duplicating).
  const headerRow = 3;
  sheet
    .getRange(headerRow, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setHorizontalAlignment("center");
  sheet.setFrozenRows(headerRow);

  // Reasonable starting column widths so headers are legible without
  // immediately needing manual resize.
  for (let c = 1; c <= headers.length; c++) {
    sheet.setColumnWidth(c, 160);
  }

  // Status formatting (dropdowns + entire-row CF rules) — mirrors the
  // live ReleaseSheet so a duplicated TEMPLATE feels identical.
  if (typeof applyStatusFormattingToSheet_ === "function") {
    try {
      applyStatusFormattingToSheet_(sheet, headerRow);
    } catch (e) {
      Logger.log(
        "[BackupRunbook] applyStatusFormattingToSheet_ on TEMPLATE failed: " + e
      );
    }
  }

  // Force the same data-row pixel height the live sheet uses, for the
  // first 100 data rows. PMs can extend later if a release exceeds
  // that.
  const rowHeightPx =
    typeof RELEASE_SHEET_DATA_ROW_HEIGHT_PX === "number" && RELEASE_SHEET_DATA_ROW_HEIGHT_PX > 0
      ? RELEASE_SHEET_DATA_ROW_HEIGHT_PX
      : 35;
  try {
    if (typeof sheet.setRowHeightsForced === "function") {
      sheet.setRowHeightsForced(headerRow + 1, 100, rowHeightPx);
    } else {
      sheet.setRowHeights(headerRow + 1, 100, rowHeightPx);
    }
  } catch (e) {
    Logger.log("[BackupRunbook] setRowHeights on TEMPLATE failed: " + e);
  }

  // Hide so the tab doesn't clutter the workbook for daily use; PMs
  // unhide via right-click on the tab bar -> Show all hidden sheets.
  try {
    sheet.hideSheet();
  } catch (e) {
    // hideSheet throws if it's the only visible sheet; ignore.
  }
}

/* =================== Tab: BACKUP BUILD =================== */
/**
 * Creates (or wipes and rewrites) the BACKUP BUILD tab. Three sections:
 *   1. Paste Jira CSV (columns A..Y)
 *   2. Paste SNOW CSV (columns AA..AY)
 *   3. Formula-driven output rows in Release Sheet shape (columns BA..)
 *
 * The output formulas use header-lookup (INDEX/MATCH against the paste
 * area's row 1) so the user's CSV columns can be in any order — what
 * matters is that the column header TEXT matches what BACKUP_OUTPUT_SOURCES
 * expects.
 */
function createOrResetBackupBuildTab_(ss) {
  const sheet = getOrCreateSheet_(ss, BACKUP_BUILD_TAB);
  resetSheet_(sheet);

  const dataRows = BACKUP_BUILD_DATA_ROWS;
  const releaseHeaders = STOP_RELEASE_SHEET_HEADER_ROW;

  // ---- Section 1: Jira paste area ----
  const jiraWidth = BACKUP_BUILD_JIRA_LAST_COL - BACKUP_BUILD_JIRA_FIRST_COL + 1;
  const jiraHeaderRow = padRowToWidth_(BACKUP_BUILD_JIRA_EXPECTED_HEADERS, jiraWidth);
  sheet
    .getRange(1, BACKUP_BUILD_JIRA_FIRST_COL, 1, jiraWidth)
    .setValues([jiraHeaderRow])
    .setFontWeight("bold")
    .setBackground("#cfe2f3")
    .setNote(
      "STEP 1: Paste your Jira CSV here, INCLUDING the header row " +
        "(it should overwrite this pre-filled row). Header text " +
        "matters; column order does not — the merge formulas in " +
        "Section 3 look up by header text."
    );

  // ---- Section 2: SNOW paste area ----
  const snowWidth = BACKUP_BUILD_SNOW_LAST_COL - BACKUP_BUILD_SNOW_FIRST_COL + 1;
  const snowHeaderRow = padRowToWidth_(BACKUP_BUILD_SNOW_EXPECTED_HEADERS, snowWidth);
  sheet
    .getRange(1, BACKUP_BUILD_SNOW_FIRST_COL, 1, snowWidth)
    .setValues([snowHeaderRow])
    .setFontWeight("bold")
    .setBackground("#fce5cd")
    .setNote(
      "STEP 2: Paste your ServiceNow change_request CSV here, " +
        "INCLUDING the header row (it should overwrite this " +
        "pre-filled row). The merge looks up by header text."
    );

  // ---- Section 3: Output (Release Sheet shape) ----
  sheet
    .getRange(1, BACKUP_BUILD_OUTPUT_FIRST_COL, 1, releaseHeaders.length)
    .setValues([releaseHeaders])
    .setFontWeight("bold")
    .setBackground("#d9ead3")
    .setNote(
      "STEP 3: After pasting Jira (Section 1) + SNOW (Section 2), the " +
        "rows below are auto-merged into the Release Sheet shape. " +
        "Select these data rows -> Copy -> in the live ReleaseSheet, " +
        "Paste Special -> Values only into your data rows."
    );

  // Build per-row formulas for every output column.
  const formulas = buildBackupBuildOutputFormulas_(dataRows, releaseHeaders);
  sheet
    .getRange(2, BACKUP_BUILD_OUTPUT_FIRST_COL, dataRows, releaseHeaders.length)
    .setFormulas(formulas);

  // Freeze top row so headers stay visible while scrolling.
  sheet.setFrozenRows(1);

  // Reasonable column widths in each section.
  for (let c = BACKUP_BUILD_JIRA_FIRST_COL; c <= BACKUP_BUILD_JIRA_LAST_COL; c++) {
    sheet.setColumnWidth(c, 180);
  }
  for (let c = BACKUP_BUILD_SNOW_FIRST_COL; c <= BACKUP_BUILD_SNOW_LAST_COL; c++) {
    sheet.setColumnWidth(c, 180);
  }
  for (let c = BACKUP_BUILD_OUTPUT_FIRST_COL; c < BACKUP_BUILD_OUTPUT_FIRST_COL + releaseHeaders.length; c++) {
    sheet.setColumnWidth(c, 160);
  }

  // Banner row above (insert a second header row at row 2 spanning
  // each section)
  // Actually: the simplest banner is via cell notes (already done) +
  // background colors. Skipping a second row keeps formula references
  // simple.
}

/**
 * For each output column in STOP_RELEASE_SHEET_HEADER_ROW, build a
 * formula string for every data row (1..dataRows). Returns a
 * dataRows x headers.length 2-D array suitable for setFormulas.
 */
function buildBackupBuildOutputFormulas_(dataRows, releaseHeaders) {
  // Range strings used inside generated formulas. Always fully
  // dollar-anchored ($A:$Y) so PMs can copy/paste output rows around
  // without breaking lookups.
  const jiraCols =
    "$" + colLetter_(BACKUP_BUILD_JIRA_FIRST_COL) + ":$" + colLetter_(BACKUP_BUILD_JIRA_LAST_COL);
  const jiraHeaderCols =
    "$" + colLetter_(BACKUP_BUILD_JIRA_FIRST_COL) + "$1:$" + colLetter_(BACKUP_BUILD_JIRA_LAST_COL) + "$1";

  const snowCrCol = colLetter_(BACKUP_BUILD_SNOW_FIRST_COL); // "Number" expected in 1st SNOW col
  const snowHeaderCols =
    "$" + colLetter_(BACKUP_BUILD_SNOW_FIRST_COL) + "$1:$" + colLetter_(BACKUP_BUILD_SNOW_LAST_COL) + "$1";
  const snowCrLookupRange =
    "$" + snowCrCol + "$2:$" + snowCrCol + "$" + (dataRows + 1);
  const snowDataRange =
    "$" + colLetter_(BACKUP_BUILD_SNOW_FIRST_COL) + "$2:$" +
    colLetter_(BACKUP_BUILD_SNOW_LAST_COL) + "$" + (dataRows + 1);

  // Output column letters keyed by header name (so SNOW lookups can
  // reference the CR # output cell at $<col><row>, and matrix lookups
  // can reference the Scrum Team output cell).
  const outputColLetterByHeader = {};
  releaseHeaders.forEach((h, i) => {
    outputColLetterByHeader[h] = colLetter_(BACKUP_BUILD_OUTPUT_FIRST_COL + i);
  });
  const crOutputCol = outputColLetterByHeader["CR #"];
  const scrumTeamOutputCol = outputColLetterByHeader["Scrum Team"];

  const matrix = "'Release & Approval Matrix'";

  const out = [];
  for (let r = 0; r < dataRows; r++) {
    const sheetRow = r + 2; // data starts at row 2
    const outputRow = [];
    for (let c = 0; c < releaseHeaders.length; c++) {
      const header = releaseHeaders[c];
      const src = BACKUP_OUTPUT_SOURCES[header] || { type: "manual" };
      let formula = "";

      switch (src.type) {
        case "manual":
          formula = ""; // intentionally blank for PM to fill in
          break;

        case "jira-direct": {
          const headerCellExpr = jiraIndexExpr_(jiraCols, jiraHeaderCols, sheetRow, src.jiraHeader);
          formula = '=IFERROR(' + headerCellExpr + ', "")';
          break;
        }

        case "jira-direct-with-fallback": {
          const primary = jiraIndexExpr_(jiraCols, jiraHeaderCols, sheetRow, src.jiraHeader);
          const fallback = jiraIndexExpr_(jiraCols, jiraHeaderCols, sheetRow, src.fallbackHeader);
          // LET so primary is evaluated once. If primary is blank or
          // errors (header not found in the user's CSV), use fallback.
          formula =
            '=LET(p, IFERROR(' + primary + ', ""), IF(p="", IFERROR(' + fallback + ', ""), p))';
          break;
        }

        case "jira-hyperlink": {
          const keyExpr = jiraIndexExpr_(jiraCols, jiraHeaderCols, sheetRow, src.jiraHeader);
          formula =
            '=LET(k, IFERROR(' + keyExpr + ', ""), ' +
            'IF(k="", "", HYPERLINK("https://pnmac.atlassian.net/browse/" & k, k)))';
          break;
        }

        case "jira-cr-extract": {
          const fieldExpr = jiraIndexExpr_(jiraCols, jiraHeaderCols, sheetRow, src.jiraHeader);
          // First CHG-prefix found in the multi-CR field (comma or
          // "##" delimited). Returns "" when no CHG match exists.
          formula =
            '=IFERROR(REGEXEXTRACT(IFERROR(' + fieldExpr + ', ""), "CHG\\d+"), "")';
          break;
        }

        case "snow-xlookup": {
          formula =
            "=" + snowXlookupBareExpr_(
              crOutputCol + sheetRow,
              snowCrLookupRange,
              snowDataRange,
              snowHeaderCols,
              src.snowHeader
            );
          break;
        }

        case "snow-xlookup-fallback": {
          const primary = snowXlookupBareExpr_(
            crOutputCol + sheetRow,
            snowCrLookupRange,
            snowDataRange,
            snowHeaderCols,
            src.snowHeader
          );
          const fallback = snowXlookupBareExpr_(
            crOutputCol + sheetRow,
            snowCrLookupRange,
            snowDataRange,
            snowHeaderCols,
            src.fallbackHeader
          );
          formula = '=LET(p, ' + primary + ', IF(p="", ' + fallback + ', p))';
          break;
        }

        case "matrix-lookup": {
          // Mirrors buildTechLeadLookupFromMatrix_ in Release-data.gs:
          //   col C of matrix = scrum team key; col F = lead;
          //   col G = sr lead. Combine as "lead/srlead" or just "lead"
          //   when sr lead is blank.
          const teamCell = "$" + scrumTeamOutputCol + sheetRow;
          formula =
            '=IFERROR(IF(' + teamCell + '="", "", ' +
            'IFERROR(' +
            'VLOOKUP(' + teamCell + ', ' + matrix + '!$C:$G, 4, FALSE) & ' +
            'IF(IFERROR(VLOOKUP(' + teamCell + ', ' + matrix + '!$C:$G, 5, FALSE), "")="", "", ' +
            '"/" & VLOOKUP(' + teamCell + ', ' + matrix + '!$C:$G, 5, FALSE))' +
            ', "")), "")';
          break;
        }

        default:
          formula = "";
      }

      outputRow.push(formula);
    }
    out.push(outputRow);
  }
  return out;
}

function jiraIndexExpr_(jiraCols, jiraHeaderCols, row, headerText) {
  const safeHeader = headerText.replace(/"/g, '""');
  return (
    'INDEX(' + jiraCols + ', ' + row +
    ', MATCH("' + safeHeader + '", ' + jiraHeaderCols + ', 0))'
  );
}

/**
 * Bare (no leading "=") SNOW XLOOKUP-by-CR expression. The caller
 * either prepends "=" directly (snow-xlookup) or composes inside a LET
 * (snow-xlookup-fallback). XLOOKUP's missing_value="" + match_mode=0
 * means a missing CR returns "" instead of #N/A; IFERROR catches the
 * INDEX/MATCH failure when the header isn't in the user's SNOW CSV.
 */
function snowXlookupBareExpr_(crCellRef, snowCrLookupRange, snowDataRange, snowHeaderCols, snowHeader) {
  const safeHeader = snowHeader.replace(/"/g, '""');
  return (
    'IFERROR(IF(' + crCellRef + '="", "", ' +
    'XLOOKUP(' + crCellRef + ', ' + snowCrLookupRange + ', ' +
    'INDEX(' + snowDataRange + ', 0, MATCH("' + safeHeader + '", ' + snowHeaderCols + ', 0)), ' +
    '"", 0)), "")'
  );
}

/* =================== Tab: BACKUP PROCEDURE =================== */
/**
 * Creates (or wipes and rewrites) the BACKUP PROCEDURE tab. Pure
 * documentation — every cell is text. Keeps the runbook in the same
 * workbook so it's never lost when a PM most needs it.
 */
function createOrResetBackupProcedureTab_(ss) {
  const sheet = getOrCreateSheet_(ss, BACKUP_PROCEDURE_TAB);
  resetSheet_(sheet);

  const cfg = (typeof getConfig === "function") ? safeGetConfig_() : {};
  const projects = cfg.projects ? String(cfg.projects) : "<paste cfg.projects here>";

  // Each entry is { text, style }. style is one of: title, h2, body,
  // pre, mono, callout. Rendered inline below.
  const lines = [
    { text: "MANUAL BACKUP — Release Sheet Runbook",                                                                style: "title" },
    { text: "Use this when 'Manually Create Release Sheet' or 'Close Sheet' (in the ITRM Functions menu) cannot be run.", style: "body" },
    { text: "All steps below are browser + Google Sheets only — no scripts required at the moment of use.",            style: "body" },
    { text: "",                                                                                                       style: "body" },

    { text: "Section A — Manual Create Release Sheet",                                                                style: "h2" },
    { text: "1. Open the saved Jira filter (Section C URL). Set 'Production Release Date' to the target date.",       style: "body" },
    { text: "   • If the target is a Friday in Pacific time, set the filter to F + Sa + Su (3-day window) — same window the script would use.", style: "body" },
    { text: "2. In Jira: Export -> Export Excel CSV (Current fields). Open the file in Excel/Sheets.",                style: "body" },
    { text: "3. Open the saved ServiceNow change_request list (Section C URL). Filter to last ~14 days OR to the CR numbers from the Jira CSV's 'ServiceNow Related Ticket' column.", style: "body" },
    { text: "4. In ServiceNow: List header -> Export -> CSV. Open the file.",                                          style: "body" },
    { text: "5. In this workbook, open the BACKUP BUILD tab.",                                                         style: "body" },
    { text: "   • Paste the Jira CSV (INCLUDING headers) starting at cell A1.",                                        style: "body" },
    { text: "   • Paste the SNOW CSV (INCLUDING headers) starting at cell AA1.",                                       style: "body" },
    { text: "6. Spot-check Section 3 (columns BA onward) — rows should populate automatically. If you see #N/A floods in a SNOW column, the CR # column may be missing CHG values, or a header name in your CSV doesn't match what the formulas expect (see BACKUP BUILD cell notes).", style: "body" },
    { text: "7. In the live ReleaseSheet tab: select all data rows below the header -> Edit -> Delete cells -> Shift up. (Or just clear the data rows.)", style: "body" },
    { text: "8. In BACKUP BUILD: select Section 3's data rows (BA through the last output column) -> Edit -> Copy.",   style: "body" },
    { text: "9. In ReleaseSheet: click the first cell below the header -> Edit -> Paste Special -> Values only.",      style: "body" },
    { text: "10. Set ReleaseSheet!A1 to the target release date (so refreshes / scheduled triggers see the right date).", style: "body" },
    { text: "11. Spot-check 2-3 rows against Jira/SNOW UIs to confirm the merge.",                                     style: "body" },
    { text: "",                                                                                                       style: "body" },

    { text: "Section B — Manual Close Sheet (archive + roll forward)",                                                style: "h2" },
    { text: "Replicates what the 'Close Sheet' / Stop Release menu does: archives the current Release Sheet to Summary, rolls A1 to the next release date, and re-pulls data for the new date.", style: "body" },
    { text: "1. In the live ReleaseSheet: select all data rows (below the header) -> Edit -> Copy.",                  style: "body" },
    { text: "2. Open the Summary tab. Click the first empty cell at the bottom of the data (Ctrl+Down on column A, then one row below).", style: "body" },
    { text: "3. Edit -> Paste Special -> Values only.",                                                                style: "body" },
    { text: "   • Summary's column order matches ReleaseSheet's, so values land correctly. If you see column drift, see COLUMN_MAPPING in apps-script/Summary.gs for the canonical mapping.", style: "body" },
    { text: "4. Back in ReleaseSheet: select all data rows -> Edit -> Delete cells -> Shift up (or just clear the cells).", style: "body" },
    { text: "5. Update ReleaseSheet!A1 to the next release date (typically the next calendar day; if next Friday, the F+Sa+Su window applies).", style: "body" },
    { text: "6. Run Section A above for the new date to repopulate the sheet.",                                       style: "body" },
    { text: "",                                                                                                       style: "body" },

    { text: "Section C — Saved URLs and one-time prep",                                                               style: "h2" },
    { text: "These are the external resources the runbook depends on. Set them up once and record the URLs here; they don't expire.", style: "body" },
    { text: "",                                                                                                       style: "body" },
    { text: "Saved Jira filter URL:",                                                                                 style: "body" },
    { text: "  <paste filter URL here, e.g. https://pnmac.atlassian.net/issues/?filter=NNNNN>",                       style: "callout" },
    { text: "",                                                                                                       style: "body" },
    { text: "Saved ServiceNow change_request list URL:",                                                              style: "body" },
    { text: "  <paste list URL here, e.g. https://pennymac.service-now.com/change_request_list.do?...>",              style: "callout" },
    { text: "",                                                                                                       style: "body" },
    { text: "Optional Google Doc with annotated screenshots:",                                                        style: "body" },
    { text: "  <paste Google Doc URL here>",                                                                          style: "callout" },
    { text: "",                                                                                                       style: "body" },
    { text: "JQL to save as a Jira filter (replace the date with a Jira-filter-controlled value, e.g. an EditableField):", style: "body" },
    { text: 'project in (' + projects + ') AND "Production Release Date" = "<DATE>"',                                 style: "pre" },
    { text: '  AND (project != "CAP" OR "Teams" = "AppDev Team")',                                                    style: "pre" },
    { text: 'ORDER BY key ASC',                                                                                       style: "pre" },
    { text: "",                                                                                                       style: "body" },
    { text: "ServiceNow change_request columns the merge formulas expect (set up the saved list view to include all of these):", style: "body" },
    { text: BACKUP_BUILD_SNOW_EXPECTED_HEADERS.filter(Boolean).join(", "),                                            style: "pre" },
    { text: "",                                                                                                       style: "body" },
    { text: "TEMPLATE tab (hidden): right-click the tab bar -> Show all hidden sheets to access. To use it, right-click TEMPLATE -> Duplicate -> rename the copy as needed.", style: "body" },
    { text: "",                                                                                                       style: "body" },

    { text: "Section D — When to use this backup, and known gaps",                                                    style: "h2" },
    { text: "Triggers (any of these is enough to switch to manual mode):",                                            style: "body" },
    { text: "  • 'Manually Create Release Sheet' throws an error twice in a row.",                                    style: "body" },
    { text: "  • Apps Script editor shows a yellow banner about exceeded daily quota.",                               style: "body" },
    { text: "  • Apps Script execution log shows repeated 'Service Spreadsheets timed out' even after retries.",     style: "body" },
    { text: "  • The script's developer is unavailable for >30 minutes on release morning.",                          style: "body" },
    { text: "  • Google Workspace status page reports a Sheets/Apps Script outage.",                                 style: "body" },
    { text: "",                                                                                                       style: "body" },
    { text: "Known gaps in manual mode (the next successful Refresh fills these in):",                                style: "body" },
    { text: "  • ChangeLog entries (NEW/MODIFIED/REMOVED) are not generated.",                                        style: "body" },
    { text: "  • Green-column protections (Approvals, Deployment Status, Validation Status, etc.) are not applied; rows are editable by anyone with edit access until the next successful Refresh.", style: "body" },
    { text: "  • Sub-second sort by SNOW start date within team is not preserved; rows are ordered by Scrum Team + CR # (PMs can re-sort manually if needed).", style: "body" },
    { text: "  • Blocker resolution lookup is unavailable from a CSV-only path; the Blockers / Blocker Status columns stay blank until a Refresh runs.", style: "body" },
    { text: "  • CTASK List requires a separate change_task export; backup leaves it blank.",                         style: "body" },
    { text: "",                                                                                                       style: "body" },
    { text: "Section E — Refreshing this runbook",                                                                    style: "h2" },
    { text: "If the Release Sheet column list changes, run Apps Script -> setupBackupRunbook() to regenerate all three tabs (TEMPLATE, BACKUP BUILD, BACKUP PROCEDURE). The function is idempotent — existing tabs are wiped and rewritten.", style: "body" }
  ];

  // Render rows. Column A holds the text; B left blank for cushion.
  const values = lines.map(l => [l.text]);
  sheet.getRange(1, 1, values.length, 1).setValues(values);
  sheet.setColumnWidth(1, 1100);

  // Apply per-row styling.
  for (let i = 0; i < lines.length; i++) {
    const r = i + 1;
    const cell = sheet.getRange(r, 1);
    switch (lines[i].style) {
      case "title":
        cell.setFontSize(18).setFontWeight("bold").setBackground("#cfe2f3");
        sheet.setRowHeight(r, 36);
        break;
      case "h2":
        cell.setFontSize(14).setFontWeight("bold").setBackground("#d9ead3");
        sheet.setRowHeight(r, 28);
        break;
      case "callout":
        cell.setFontFamily("Roboto Mono").setBackground("#fff2cc").setFontSize(11);
        break;
      case "pre":
        cell.setFontFamily("Roboto Mono").setBackground("#f3f3f3").setFontSize(11);
        break;
      case "mono":
        cell.setFontFamily("Roboto Mono").setFontSize(11);
        break;
      case "body":
      default:
        cell.setFontSize(11);
        break;
    }
  }

  sheet.setFrozenRows(1);
}

/* =================== Helpers =================== */

function getOrCreateSheet_(ss, name) {
  const existing = ss.getSheetByName(name);
  if (existing) {
    // Make sure the sheet is visible while we edit it; we'll re-hide
    // TEMPLATE at the end if appropriate.
    try { existing.showSheet(); } catch (e) { /* ignore */ }
    return existing;
  }
  return ss.insertSheet(name);
}

function resetSheet_(sheet) {
  // Clear values, formats, data validations, conditional formats, and
  // notes — everything. Used to make the setup function idempotent.
  try { sheet.clear(); } catch (e) { /* ignore */ }
  try { sheet.clearConditionalFormatRules(); } catch (e) { /* ignore */ }
  try { sheet.clearNotes(); } catch (e) { /* ignore */ }
  try {
    const lc = Math.max(sheet.getMaxColumns(), 1);
    const lr = Math.max(sheet.getMaxRows(), 1);
    sheet.getRange(1, 1, lr, lc).clearDataValidations();
  } catch (e) { /* ignore */ }
  // Drop any range protections we may have created previously.
  try {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
    protections.forEach(p => {
      try { p.remove(); } catch (_) { /* ignore */ }
    });
  } catch (e) { /* ignore */ }
}

function safeGetConfig_() {
  try {
    return getConfig();
  } catch (e) {
    return {};
  }
}

function padRowToWidth_(row, width) {
  const out = (row || []).slice();
  while (out.length < width) out.push("");
  return out.slice(0, width);
}

/** 1 -> "A", 27 -> "AA". Mirrors colNumberToLetter_ in StatusFormatting.gs. */
function colLetter_(n) {
  let s = "";
  let x = n;
  while (x > 0) {
    const rem = (x - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    x = Math.floor((x - 1) / 26);
  }
  return s;
}

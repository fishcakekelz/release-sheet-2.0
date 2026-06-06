/**********************
  Summary Sheet Functions
  - Archive Release Sheet to Summary
**********************/

/* =================== Configuration =================== */
const SUMMARY_SHEET_NAME = "Summary";
const RELEASE_SHEET_NAME = "ReleaseSheet";
const SUMMARY_SHEET_ROW_HEIGHT_PX = 21;

/* =================== Column Mapping: Release Sheet → Summary =================== */
// Maps Release Sheet column names to Summary column names
const COLUMN_MAPPING = {
  "Pre-Release Check Summary": "Pre-Release Check Summary",
  "JIRA Item": "JIRA Item",
  "CR #": "CR #",
  "Start Time": "Start Time",
  "Validation Owner": "Validation Owner",
  "Scrum Team": "Scrum Team",
  "Comments": "Notes/Comments",
  "Approvals": "Approvals",
  "Dark Deployment": "Dark Deployment",
  "Late Addition": "Late Add",
  "Acceptance Criteria": "Acceptance Criteria",
  "Last Updated JIRA comment": "Last Updated JIRA comment",
  "Workstream": "Workstream",
  "Channel": "Chnl",
  "Jira Prod Release Date": "Jira Prod Release Date",
  "Jira Status": "Jira Status",
  "Jira Resolution": "Jira Resolution",
  "Issue Type": "Issue Type",
  "JIRA Assignee": "Jira Assignee",
  "Related Issues": "Related Issues",
  "Deployment Status": "Deployment Status",
  "Validation Status": "Validation Status",
  "Release Type": "Release Type",
  "CHG Type": "CHG Type",
  "CTASK List": "CTASK details",
  "Impact Description": "CR Impact Description",
  "Change Plan": "Change Plan",
  "Rollback Plan": "Rollback Plan",
  "JIRA Summary": "JIRA Summary",
  "JIRA Description": "JIRA Description",
  "Validation Plan": "Validation Plan",
  "Deploy Start": "Deploy Start",
  "Deploy End": "Deploy End",
  "Impacted CI": "Impacted CI",
  "Date Submitted": "Date Submitted",
  "Affected Groups": "Affected Groups",
  "Affected Locations": "Affected Locations",
  "Assignment Group": "Assignment Group",
  "Tech Lead / Sr Tech Lead": "Tech Lead / Sr Tech Lead",
  "JIRA Approval Comment Sent": "JIRA Approval Comment Sent",
  "Change Assignee": "Change Owner",
  "Validation Start": "Validate Start",
  "Validation End": "Validate End",
  "Deploy Duration": "Deploy Duration",
  "Release Duration": "Release Duration",
  // Not copied to Summary (excluded columns)
  "Blocker Status": null,
  "Late Approval": null,
  "Division": null
};

/**
 * Forces Summary rows to 21px so data doesn't auto-expand row height.
 *
 * IMPORTANT: pass `firstRow` to limit the write to newly-appended rows. Rewriting
 * heights for every existing Summary row on every archive is the operation that
 * triggers "Service Spreadsheets timed out while accessing document" once
 * Summary has accumulated a few thousand rows. Existing rows already have the
 * forced height from when they were originally appended.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} [firstRow=1] First row to force-resize (1-based, inclusive).
 *   Default 1 keeps the legacy "do the whole sheet" behavior for callers that
 *   don't pass a value.
 */
function applySummarySheetRowHeights21_(sheet, firstRow) {
  if (!sheet) return;
  const h = SUMMARY_SHEET_ROW_HEIGHT_PX;
  const last = sheet.getLastRow();
  if (last < 1) return;

  const start = Math.max(1, Math.floor(Number(firstRow) || 1));
  if (start > last) return;
  const numRows = last - start + 1;

  if (typeof sheet.setRowHeightsForced === "function") {
    try {
      sheet.setRowHeightsForced(start, numRows, h);
      return;
    } catch (e) {
      // fall through
    }
  }
  if (typeof sheet.setRowHeights === "function") {
    try {
      sheet.setRowHeights(start, numRows, h);
      return;
    } catch (e) {
      // per-row
    }
  }
  for (let r = start; r <= last; r++) {
    if (typeof sheet.setRowHeightsForced === "function") {
      try {
        sheet.setRowHeightsForced(r, 1, h);
      } catch (e) {
        sheet.setRowHeight(r, h);
      }
    } else {
      sheet.setRowHeight(r, h);
    }
  }
}

/* =================== Helper: Get Header Index Map =================== */
function getHeaderMap(sheet, headerRow = 1) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return { map: {}, headers: [], duplicates: {} };
  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getValues()[0]
    .map(h => (h ? h.toString().trim() : ""));
  const map = {};
  const duplicates = {}; // Track duplicate column positions

  headers.forEach((h, i) => {
    if (h) {
      if (map[h] !== undefined) {
        if (!duplicates[h]) {
          duplicates[h] = [map[h]];
        }
        duplicates[h].push(i);
      } else {
        map[h] = i;
      }
    }
  });

  return { map, headers, duplicates };
}

/* =================== Internal: shared archive worker =================== */
/**
 * Reads the Config sheet (if present) and returns the configured
 * `sheetHeaderRow` for the Release sheet, defaulting to 1.
 *
 * Centralized so all three archive entry points stay in sync if the
 * Config schema changes.
 */
function getReleaseSheetHeaderRowFromConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (!configSheet) return 1;
  const configData = configSheet.getDataRange().getValues();
  let headerRow = 1;
  configData.forEach(row => {
    if (row[0] === "sheetHeaderRow") headerRow = Number(row[1]) || 1;
  });
  return headerRow;
}

/**
 * Shared archive worker used by all three public archive entry points
 * (archiveReleaseToSummary, archiveReleaseDataSilent, archiveAndClearRelease).
 *
 * Steps:
 *   1. Build header maps for both sheets via getHeaderMap.
 *   2. Read the Release data range with getValues() (fast — native types
 *      round-trip into Summary; let Summary's column formats render dates).
 *   3. Filter out completely empty rows.
 *   4. Build a `pairs` table once (releaseIdx -> summaryIdx [+ duplicate
 *      summary indices]) so the per-row mapping loop avoids re-iterating
 *      Object.entries(COLUMN_MAPPING) and skips per-cell map lookups.
 *   5. Append to Summary at lastRow + 1 with retry wrapping.
 *   6. Flush, then re-apply Summary row heights (only the new rows) and
 *      status formatting (fast-paths internally when rules already exist).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} releaseSheet
 * @param {GoogleAppsScript.Spreadsheet.Sheet} summarySheet
 * @param {number} headerRow Release sheet header row (1-based).
 * @returns {{ success: boolean, message: string, rowCount: number, appendRow: number }}
 */
function archiveReleaseDataInternal_(releaseSheet, summarySheet, headerRow) {
  if (!releaseSheet) {
    return { success: false, message: "Release Sheet not found.", rowCount: 0, appendRow: 0 };
  }
  if (!summarySheet) {
    return { success: false, message: "Summary sheet not found.", rowCount: 0, appendRow: 0 };
  }

  const releaseHeaders = getHeaderMap(releaseSheet, headerRow);
  const summaryHeaders = getHeaderMap(summarySheet, 1);

  const lastRow = releaseSheet.getLastRow();
  const lastCol = releaseSheet.getLastColumn();

  if (lastRow <= headerRow) {
    return { success: true, message: "No data to archive.", rowCount: 0, appendRow: 0 };
  }

  const numRows = lastRow - headerRow;
  // getValues (not getDisplayValues): native types round-trip into Summary
  // and Summary's per-column number/date formats render them. Wrapped in
  // the retry helper because a wide Release range read can hit transient
  // "Service timed out" on large sheets.
  const releaseData = withSpreadsheetRetry_(
    () => releaseSheet.getRange(headerRow + 1, 1, numRows, lastCol).getValues(),
    "archive:getValues"
  );

  const nonEmptyData = releaseData.filter(row => row.some(cell => cell !== "" && cell != null));

  if (nonEmptyData.length === 0) {
    return { success: true, message: "No data to archive.", rowCount: 0, appendRow: 0 };
  }

  const summaryColCount =
    summarySheet.getLastColumn() || Object.keys(summaryHeaders.map).length;

  // Precompute mapping pairs once. With ~46 entries in COLUMN_MAPPING and
  // 50-100 release rows, this drops 5k-10k Object.entries / map lookups
  // out of the hot path.
  const pairs = [];
  for (const [releaseCol, summaryCol] of Object.entries(COLUMN_MAPPING)) {
    if (summaryCol === null) continue;
    const releaseIdx = releaseHeaders.map[releaseCol];
    const summaryIdx = summaryHeaders.map[summaryCol];
    if (releaseIdx === undefined || summaryIdx === undefined) continue;
    pairs.push({
      releaseIdx,
      summaryIdx,
      dupIdxs: summaryHeaders.duplicates[summaryCol] || null
    });
  }

  const mappedData = nonEmptyData.map(releaseRow => {
    const summaryRow = new Array(summaryColCount).fill("");
    for (let i = 0; i < pairs.length; i++) {
      const p = pairs[i];
      const value = releaseRow[p.releaseIdx];
      summaryRow[p.summaryIdx] = value;
      if (p.dupIdxs) {
        for (let j = 0; j < p.dupIdxs.length; j++) {
          summaryRow[p.dupIdxs[j]] = value;
        }
      }
    }
    return summaryRow;
  });

  const summaryLastRow = summarySheet.getLastRow();
  const appendRow = summaryLastRow + 1;

  if (mappedData.length > 0 && mappedData[0].length > 0) {
    withSpreadsheetRetry_(
      () =>
        summarySheet
          .getRange(appendRow, 1, mappedData.length, mappedData[0].length)
          .setValues(mappedData),
      "archive:setValues"
    );
  }

  // Drain the queued setValues before doing more spreadsheet ops — the next
  // step (row-height + CF rules) is exactly where StopRelease was hitting
  // "Service Spreadsheets timed out while accessing document".
  SpreadsheetApp.flush();

  // Only force the row height on the rows we just appended. Rewriting heights
  // for every existing Summary row on every archive scaled linearly with
  // history and was the actual op causing the timeout.
  withSpreadsheetRetry_(
    () => applySummarySheetRowHeights21_(summarySheet, appendRow),
    "archive:setRowHeights(new rows only)"
  );

  // Drain the row-height batch before applyStatusFormattingToSheet_ touches
  // sheet-wide CF rules + data validations. Without this flush, CF rules has
  // to drain a fat queue itself and that's the boundary where archive&clear
  // was hitting "Service Spreadsheets timed out while accessing document".
  SpreadsheetApp.flush();

  withSpreadsheetRetry_(
    () => applyStatusFormattingToSheet_(summarySheet, 1),
    "archive:applyStatusFormatting"
  );

  return {
    success: true,
    message: "Archive complete.",
    rowCount: mappedData.length,
    appendRow
  };
}

/* =================== Archive Release Sheet to Summary =================== */
function archiveReleaseToSummary() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Archive to Summary",
    "This will copy all data from the Release Sheet to the Summary sheet.\n\nDo you want to continue?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert("Archive cancelled.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const releaseSheet = ss.getSheetByName(RELEASE_SHEET_NAME);
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  if (!releaseSheet) {
    ui.alert("Error: Release Sheet not found.");
    return;
  }

  if (!summarySheet) {
    ui.alert("Error: Summary sheet not found.");
    return;
  }

  const headerRow = getReleaseSheetHeaderRowFromConfig_();
  const lastRow = releaseSheet.getLastRow();

  if (lastRow <= headerRow) {
    ui.alert("No data to archive. Release Sheet is empty.");
    return;
  }

  const result = archiveReleaseDataInternal_(releaseSheet, summarySheet, headerRow);
  if (!result.success) {
    ui.alert("Archive failed: " + result.message);
    return;
  }
  if (result.rowCount === 0) {
    ui.alert("No data to archive. All rows are empty.");
    return;
  }

  const now = new Date();
  console.log(
    `Archived ${result.rowCount} rows to Summary sheet at row ${result.appendRow} on ${now}`
  );

  ui.alert(
    `Archive Complete!\n\n${result.rowCount} rows appended to Summary sheet (rows ${
      result.appendRow
    }-${result.appendRow + result.rowCount - 1}).\n\nColumns were mapped automatically.`
  );
}

/* =================== Silent Archive (for use by createReleaseSheet) =================== */
/**
 * Archives release sheet data to Summary without user prompts
 * @param {GoogleAppsScript.Spreadsheet.Sheet} releaseSheet - The release sheet
 * @param {number} headerRow - The header row number
 * @returns {{ success: boolean, message: string, rowCount: number }} result
 */
function archiveReleaseDataSilent(releaseSheet, headerRow) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  const result = archiveReleaseDataInternal_(releaseSheet, summarySheet, headerRow);
  if (result.success && result.rowCount > 0) {
    const now = new Date();
    console.log(
      `Silent archive: ${result.rowCount} rows to Summary at row ${result.appendRow} on ${now}`
    );
  }
  // Preserve the existing return shape (no `appendRow`) so existing callers
  // such as StopRelease that read `success / message / rowCount` are
  // unchanged.
  return {
    success: result.success,
    message: result.message,
    rowCount: result.rowCount
  };
}

/* =================== Clear Release Sheet After Archive =================== */
function clearReleaseSheetData() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Clear Release Sheet",
    "⚠️ WARNING: This will delete all data from the Release Sheet (keeping headers).\n\nMake sure you have archived first!\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert("Clear cancelled.");
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const releaseSheet = ss.getSheetByName(RELEASE_SHEET_NAME);

  if (!releaseSheet) {
    ui.alert("Error: Release Sheet not found.");
    return;
  }

  const configSheet = ss.getSheetByName("Config");
  let headerRow = 1;
  if (configSheet) {
    const configData = configSheet.getDataRange().getValues();
    configData.forEach(row => {
      if (row[0] === "sheetHeaderRow") headerRow = Number(row[1]) || 1;
    });
  }

  const lastRow = releaseSheet.getLastRow();
  const lastCol = releaseSheet.getLastColumn();

  if (lastRow > headerRow && lastCol > 0) {
    releaseSheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).clearContent();
    ui.alert("Release Sheet data cleared (headers preserved).");
  } else {
    ui.alert("Release Sheet is already empty.");
  }
}

/* =================== Archive and Clear (Combined) =================== */
function archiveAndClearRelease() {
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    "Archive & Clear Release Sheet",
    "This will:\n1. Copy all Release Sheet data to Summary (with column mapping)\n2. Clear the Release Sheet (keep headers)\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const releaseSheet = ss.getSheetByName(RELEASE_SHEET_NAME);
  const summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);

  if (!releaseSheet || !summarySheet) {
    ui.alert("Error: Required sheets not found.");
    return;
  }

  const headerRow = getReleaseSheetHeaderRowFromConfig_();

  const lastRow = releaseSheet.getLastRow();
  const lastCol = releaseSheet.getLastColumn();

  if (lastRow <= headerRow) {
    ui.alert("No data to archive.");
    return;
  }

  const result = archiveReleaseDataInternal_(releaseSheet, summarySheet, headerRow);
  if (!result.success) {
    ui.alert("Archive failed: " + result.message);
    return;
  }
  if (result.rowCount === 0) {
    ui.alert("No data to archive.");
    return;
  }

  // Drain everything the archive helper just queued (Summary writes,
  // row-heights, CF rules) before we touch a different sheet. Cross-sheet
  // writes that happen while the prior write batch is still pending are
  // exactly what triggered "Service Spreadsheets timed out while accessing
  // document" in this path; same shape as the ChangeLog timeout we hit in
  // refreshReleaseSheetInternal_.
  SpreadsheetApp.flush();

  // Re-read the Release range size right before clearing, in case anything
  // shifted during the archive, and wrap the clear in withSpreadsheetRetry_
  // so a transient timeout retries with another flush + small backoff
  // instead of failing the whole archive-and-clear.
  const releaseLastRow = releaseSheet.getLastRow();
  const releaseLastCol = releaseSheet.getLastColumn();
  if (releaseLastRow > headerRow && releaseLastCol > 0) {
    withSpreadsheetRetry_(
      () =>
        releaseSheet
          .getRange(headerRow + 1, 1, releaseLastRow - headerRow, releaseLastCol)
          .clearContent(),
      "archiveAndClear:clearContent"
    );
  }

  ui.alert(
    `Archive & Clear Complete!\n\n${
      result.rowCount
    } rows appended to Summary (rows ${result.appendRow}-${
      result.appendRow + result.rowCount - 1
    }).\nRelease Sheet has been cleared.`
  );
}

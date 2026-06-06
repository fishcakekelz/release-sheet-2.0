/**********************
  ReleaseSheet Backup populator

  Public entry: populateReleaseSheetBackup()

  Pulls the same Jira + ServiceNow data the main "Manually Create
  Release Sheet" flow uses, but writes it to a separate
  "ReleaseSheet Backup" tab instead of touching the live ReleaseSheet.

  Skips, by design, every post-write step that the main Create flow
  runs: row heights, font reset, conditional formatting + dropdowns,
  CR # blue highlights, green-column protections, ChangeLog, and the
  locked-release-date guard. The backup is meant to preserve data
  cheaply and resiliently — formatting / dropdowns / protections all
  exist on the live ReleaseSheet that PMs copy values back into.

  One-time user setup: right-click ReleaseSheet -> Duplicate, rename
  the copy to "ReleaseSheet Backup". The function detects a missing
  tab and shows a clear setup message instead of throwing.

  Why a separate function instead of a flag on createReleaseSheet:
  decoupled failure modes. If the main Create flow breaks because of
  a regression in formatting / protection / ChangeLog code, this
  backup runs none of that and continues to work.
**********************/

/* =================== Tab name =================== */
const S_RELEASE_BACKUP = "ReleaseSheet Backup";

/* =================== Public entry point =================== */

/**
 * Populates the "ReleaseSheet Backup" tab with Jira + SNOW data for
 * the release date in ReleaseSheet!A1. Idempotent — re-running clears
 * the data area and rewrites it.
 */
function populateReleaseSheetBackup() {
  const ui = SpreadsheetApp.getUi();
  const t0 = Date.now();
  const __mark = (label, since) =>
    Logger.log("[Backup] " + label + ": " + (Date.now() - since) + "ms");
  const __step = (label, fn) => {
    Logger.log("[Backup] starting: " + label);
    const t = Date.now();
    try {
      const r = fn();
      Logger.log("[Backup] done: " + label + " (" + (Date.now() - t) + "ms)");
      return r;
    } catch (e) {
      Logger.log(
        "[Backup] FAILED: " + label + " (" + (Date.now() - t) + "ms) — " +
          (e && e.message ? e.message : e)
      );
      throw e;
    }
  };

  Logger.log("[Backup] start");
  try {
    const cfg = getConfig();
    const ss = SpreadsheetApp.getActive();

    const backup = ss.getSheetByName(S_RELEASE_BACKUP);
    if (!backup) {
      ui.alert(
        "Tab '" + S_RELEASE_BACKUP + "' not found.\n\n" +
          "Set it up once: right-click '" + S_RELEASE + "' -> Duplicate, " +
          "rename the copy to '" + S_RELEASE_BACKUP + "', then re-run."
      );
      Logger.log("[Backup] aborted — '" + S_RELEASE_BACKUP + "' tab missing");
      return;
    }

    const releaseSheet = ss.getSheetByName(S_RELEASE);
    if (!releaseSheet) {
      ui.alert("'" + S_RELEASE + "' tab not found.");
      Logger.log("[Backup] aborted — '" + S_RELEASE + "' tab missing");
      return;
    }

    const releaseDateStr = getReleaseDateFromReleaseSheet();
    Logger.log("[Backup] release date from " + S_RELEASE + "!A1: " + releaseDateStr);

    // ==== Header map of the backup tab (same shape as ReleaseSheet) ====
    const { headers, map } = getHeaderIndexMapSafe(backup, cfg.sheetHeaderRow);
    if (!headers.length) {
      ui.alert(
        "'" + S_RELEASE_BACKUP + "' has no headers at row " +
          cfg.sheetHeaderRow + ".\n\n" +
          "Re-run after duplicating '" + S_RELEASE + "' so the backup tab " +
          "has the same header row."
      );
      Logger.log(
        "[Backup] aborted — backup tab has no headers at row " + cfg.sheetHeaderRow
      );
      return;
    }
    Logger.log("[Backup] backup headers: " + headers.length + " columns");

    // ==== Jira fetch ====
    const __tJira = Date.now();
    const issues = __step(
      "fetchJiraIssuesByJql + normaliseIssue",
      () => fetchJiraIssuesByJql(buildJqlForDate(releaseDateStr, cfg), cfg).map(normaliseIssue)
    );
    __mark("Jira (" + issues.length + " issues)", __tJira);

    // ==== SNOW + blocker resolution (single parallel batch) ====
    const allCrs = collectAllSnowCRNumbersFromIssues(issues);
    const blockerKeys = collectAllBlockerKeysFromIssues_(issues);
    const includeCTaskList = map["CTASK List"] !== undefined;
    const __tNet = Date.now();
    const { snowDataByCR, blockerLookup } = __step(
      "fetchSnowAndBlockersInParallel_ (" +
        allCrs.length + " CRs, " + blockerKeys.length + " blocker keys, " +
        "ctaskList=" + includeCTaskList + ")",
      () => fetchSnowAndBlockersInParallel_(allCrs, blockerKeys, cfg, includeCTaskList)
    );
    __mark("SNOW + blockers parallel", __tNet);
    applyBlockerResolutionLookup_(issues, blockerLookup);

    // ==== Expand to display rows + sort ====
    const snowWindowYmd = expandFridayWeekendJqlYmdsForY0_(releaseDateStr);
    let expanded = expandIssuesForDisplayRows(issues, snowDataByCR, snowWindowYmd);
    expanded = sortExpandedIssuesByScrumThenCr_(expanded, snowDataByCR);
    Logger.log(
      "[Backup] expanded display rows: " + expanded.length +
        " (multi-CR + sorted by ScrumTeam/CR)"
    );

    // ==== Build row matrix (Jira + SNOW + Tech Lead matrix all baked in) ====
    const techLeadLookup = buildTechLeadLookupFromMatrix_(ss);
    const __tBuild = Date.now();
    const rows = expanded.map(i =>
      buildRowFromIssue(i, headers, map, cfg, snowDataByCR, techLeadLookup)
    );
    __mark("Build new rows matrix (" + rows.length + " × " + headers.length + ")", __tBuild);

    // ==== Clear backup data area ====
    const lastRow = backup.getLastRow();
    if (lastRow > cfg.sheetHeaderRow) {
      __step(
        "clearContent backup data range (" +
          (lastRow - cfg.sheetHeaderRow) + " × " + headers.length + ")",
        () =>
          withSpreadsheetRetry_(
            () =>
              backup
                .getRange(cfg.sheetHeaderRow + 1, 1, lastRow - cfg.sheetHeaderRow, headers.length)
                .clearContent(),
            "backup:clearContent"
          )
      );
    }

    // ==== Write rows ====
    if (rows.length > 0) {
      __step(
        "setValues backup rows (" + rows.length + " × " + headers.length + ")",
        () =>
          withSpreadsheetRetry_(
            () =>
              backup
                .getRange(cfg.sheetHeaderRow + 1, 1, rows.length, headers.length)
                .setValues(rows),
            "backup:setValues"
          )
      );
    }

    __mark("TOTAL Backup populate", t0);
    Logger.log(
      "[Backup] complete: " + rows.length + " rows for " + releaseDateStr
    );

    ui.alert(
      "ReleaseSheet Backup populated.\n\n" +
        rows.length + " rows written for " + releaseDateStr + ".\n\n" +
        "Skipped (by design): row heights, font reset, conditional " +
        "formatting, dropdowns, green-column protections, ChangeLog, " +
        "locked-release-date guard.\n\n" +
        "If the live '" + S_RELEASE + "' is corrupted, copy values from " +
        "'" + S_RELEASE_BACKUP + "' back into it."
    );
  } catch (e) {
    Logger.log(
      "[Backup] FAILED at top level after " + (Date.now() - t0) + "ms: " +
        (e && e.message ? e.message : e)
    );
    ui.alert(
      "Populate ReleaseSheet Backup failed: " +
        (e && e.message ? e.message : String(e)) +
        "\n\nSee the Apps Script execution log for the [Backup] trace " +
        "showing which step failed."
    );
    Logger.log(e);
  }
}

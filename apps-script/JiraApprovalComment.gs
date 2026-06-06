/**********************
  Jira Approval → Comment

  When the "Approvals" column in the ReleaseSheet is edited to a value that
  looks like an approval (see APPROVAL_INDICATOR_REGEX in constants.gs), this
  script posts a standard comment to the corresponding Jira issue and stamps
  the "JIRA Approval Comment Sent" column so the row is not processed twice.

  Wiring:
    1. From the "ITRM Functions" menu → "Approval Comments" → "Install Trigger".
       (This must be done once per spreadsheet, by an editor with permission to
       call external APIs.) The installed trigger runs onApprovalEdit(e) below
       under the installer's account so UrlFetchApp is authorized.
    2. To process rows that were approved before the trigger was installed,
       use "Approval Comments" → "Post Comments For All Approved Rows".

  Dedup logic:
    - APPROVAL_COMMENT_SENT_COLUMN ("JIRA Approval Comment Sent") must be empty for the row
      to be processed. After a successful post the cell is set to a timestamp.
**********************/

/* =================== Installable trigger setup =================== */

/**
 * Registers an installable on-edit trigger so the handler runs with the
 * installer's auth scope (UrlFetchApp + Jira creds via Config sheet).
 */
function installApprovalEditTrigger() {
  const ui = SpreadsheetApp.getUi();
  removeApprovalEditTriggerSilent_();

  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("onApprovalEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("Approval edit trigger installed.");
  ui.alert(
    "Approval Comment Trigger Installed",
    "Edits to the \"" +
      APPROVAL_INDICATOR_COLUMN +
      "\" column on \"" +
      S_RELEASE +
      "\" will now post a comment to Jira and stamp the \"" +
      APPROVAL_COMMENT_SENT_COLUMN +
      "\" column.\n\n" +
      "Tip: only the user who installs this trigger can run it. If a different " +
      "PM owns approvals, ask them to run \"Install Trigger\" once.",
    ui.ButtonSet.OK
  );
}

function removeApprovalEditTrigger() {
  const removed = removeApprovalEditTriggerSilent_();
  SpreadsheetApp.getUi().alert(
    removed > 0
      ? "Removed " + removed + " approval-edit trigger(s)."
      : "No approval-edit triggers were installed."
  );
}

function removeApprovalEditTriggerSilent_() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  triggers.forEach(t => {
    if (t.getHandlerFunction() === "onApprovalEdit") {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  if (removed) Logger.log("Removed " + removed + " approval-edit trigger(s).");
  return removed;
}

/* =================== Edit trigger handler =================== */

/**
 * Installable on-edit trigger. Fires for every cell edit on the spreadsheet,
 * filters to the Approvals column on ReleaseSheet, and posts a Jira comment
 * if the new value looks like an approval and the row has not been stamped.
 *
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onApprovalEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    if (!sheet || sheet.getName() !== S_RELEASE) return;

    const cfg = getConfig();
    const headerRow = cfg.sheetHeaderRow;
    const editedRow = e.range.getRow();
    if (editedRow <= headerRow) return;

    const { map } = getHeaderIndexMapSafe(sheet, headerRow);
    const approvalCol = map[APPROVAL_INDICATOR_COLUMN];
    if (approvalCol === undefined) {
      Logger.log("onApprovalEdit: \"" + APPROVAL_INDICATOR_COLUMN + "\" column not found.");
      return;
    }

    // Only react to edits on the approvals column. e.range is 1-based.
    if (e.range.getColumn() !== approvalCol + 1) return;

    // For multi-cell pastes, walk every row in the edited range.
    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    const approverFromUser = (e.user && e.user.getEmail && e.user.getEmail()) || "";

    for (let r = 0; r < numRows; r++) {
      const rowNum = startRow + r;
      processApprovalRow_(sheet, rowNum, headerRow, map, cfg, approverFromUser);
    }
  } catch (err) {
    Logger.log("onApprovalEdit failed: " + err.message + "\n" + (err.stack || ""));
  }
}

/* =================== Manual: scan and post for all approved rows =================== */

/**
 * Walks every data row on the ReleaseSheet and posts a Jira comment for any
 * row where the Approvals column matches APPROVAL_INDICATOR_REGEX and the
 * "JIRA Approval Comment Sent" column is still blank. Useful for catching rows that were
 * approved before the trigger was installed.
 */
function processApprovedRowsBulk() {
  const ui = SpreadsheetApp.getUi();
  const cfg = getConfig();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(S_RELEASE);
  if (!sheet) {
    ui.alert("ReleaseSheet not found.");
    return;
  }
  const headerRow = cfg.sheetHeaderRow;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) {
    ui.alert("No data rows to process.");
    return;
  }

  const { map } = getHeaderIndexMapSafe(sheet, headerRow);
  const approvalCol = map[APPROVAL_INDICATOR_COLUMN];
  const sentCol = map[APPROVAL_COMMENT_SENT_COLUMN];
  const jiraCol = map["JIRA Item"];

  if (approvalCol === undefined || sentCol === undefined || jiraCol === undefined) {
    ui.alert(
      "Required columns missing on ReleaseSheet:\n" +
        " • " + APPROVAL_INDICATOR_COLUMN + "\n" +
        " • " + APPROVAL_COMMENT_SENT_COLUMN + "\n" +
        " • JIRA Item"
    );
    return;
  }

  const data = sheet
    .getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn())
    .getValues();

  // Identify candidate rows up-front so we can confirm before hitting Jira.
  const candidates = [];
  data.forEach((row, idx) => {
    const approval = (row[approvalCol] || "").toString();
    const sent = (row[sentCol] || "").toString().trim();
    const issueKey = (row[jiraCol] || "").toString().trim();
    if (!issueKey) return;
    if (sent) return;
    if (!APPROVAL_INDICATOR_REGEX.test(approval)) return;
    candidates.push({ rowNum: headerRow + 1 + idx, issueKey, approval });
  });

  if (!candidates.length) {
    ui.alert("No rows need a Jira approval comment.");
    return;
  }

  const confirm = ui.alert(
    "Post Jira approval comments?",
    "Found " + candidates.length + " row(s) with a name in \"" +
      APPROVAL_INDICATOR_COLUMN + "\" and no \"" +
      APPROVAL_COMMENT_SENT_COLUMN + "\" stamp:\n\n" +
      candidates.slice(0, 10).map(c => "• " + c.issueKey + "  — " + c.approval).join("\n") +
      (candidates.length > 10 ? "\n• …and " + (candidates.length - 10) + " more" : "") +
      "\n\nPost an \"Approved in Release Sheet by <name>\" comment to each Jira ticket?",
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const approverFromUser =
    (Session.getActiveUser() && Session.getActiveUser().getEmail()) || "";
  let posted = 0;
  let failed = 0;
  candidates.forEach(c => {
    const ok = processApprovalRow_(sheet, c.rowNum, headerRow, map, cfg, approverFromUser);
    if (ok) posted++;
    else failed++;
  });

  ui.alert(
    "Approval comments complete.\n\n" +
      "✅ Posted: " + posted + "\n" +
      (failed ? "❌ Failed: " + failed + " (see logs)\n" : "") +
      "Rows are stamped in the \"" + APPROVAL_COMMENT_SENT_COLUMN + "\" column."
  );
}

/* =================== Core: post one row's comment =================== */

/**
 * Posts the approval comment for a single row if it qualifies. Returns true if
 * a comment was successfully posted, false otherwise (skipped or errored).
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} rowNum 1-based sheet row
 * @param {number} headerRow
 * @param {Object<string, number>} map header → 0-based column index
 * @param {Object} cfg getConfig() result
 * @param {string} approverFromUser e.user.getEmail() (may be "")
 * @returns {boolean}
 */
function processApprovalRow_(sheet, rowNum, headerRow, map, cfg, approverFromUser) {
  const approvalCol = map[APPROVAL_INDICATOR_COLUMN];
  const sentCol = map[APPROVAL_COMMENT_SENT_COLUMN];
  const jiraCol = map["JIRA Item"];

  if (approvalCol === undefined || sentCol === undefined || jiraCol === undefined) {
    Logger.log(
      "processApprovalRow_: missing required columns. Approvals=" +
        approvalCol + " CommentSent=" + sentCol + " JIRA Item=" + jiraCol
    );
    return false;
  }
  if (rowNum <= headerRow) return false;

  // Read all three cells in one call.
  const lastCol = sheet.getLastColumn();
  const rowVals = sheet.getRange(rowNum, 1, 1, lastCol).getValues()[0];
  const approval = (rowVals[approvalCol] || "").toString();
  const alreadySent = (rowVals[sentCol] || "").toString().trim();
  const issueKey = (rowVals[jiraCol] || "").toString().trim();

  if (!issueKey) return false;
  if (alreadySent) return false; // already processed
  if (!APPROVAL_INDICATOR_REGEX.test(approval)) return false;

  const approverName = pickApproverName_(approval, approverFromUser);
  const tz = Session.getScriptTimeZone();
  const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss z");
  const commentText =
    "Approved in Release Sheet by " + approverName + " on " + timestamp + ".";

  try {
    postJiraApprovalComment_(issueKey, commentText, cfg);
  } catch (err) {
    Logger.log(
      "processApprovalRow_: failed to post comment for " + issueKey +
        " (row " + rowNum + "): " + err.message
    );
    return false;
  }

  // Stamp "JIRA Approval Comment Sent" so we don't re-fire. Use a short token + timestamp.
  sheet.getRange(rowNum, sentCol + 1).setValue("YES " + timestamp);
  Logger.log("Posted Jira approval comment for " + issueKey + " (row " + rowNum + ").");
  return true;
}

/**
 * Picks the most useful name for the approver. The PM / Sr Tech Lead types
 * their name into the Approvals cell, so that text is the source of truth.
 *  1. The text in the Approvals cell (the typed name).
 *  2. Otherwise the on-edit user's email (installable trigger only).
 *  3. Otherwise a constant fallback so the comment is never blank.
 */
function pickApproverName_(approvalCellValue, approverFromUser) {
  const cell = (approvalCellValue || "").toString().trim();
  if (cell) return cell;

  const fromUser = (approverFromUser || "").toString().trim();
  if (fromUser) return fromUser;

  return "the Release Sheet";
}

/* =================== Jira REST: POST comment (ADF) =================== */

/**
 * Posts a plain-text comment to /rest/api/3/issue/{issueKey}/comment.
 * Jira Cloud requires Atlassian Document Format (ADF) for the body.
 *
 * @param {string} issueKey e.g. "ABC-123"
 * @param {string} commentText
 * @param {Object} cfg getConfig() result (jiraBaseUrl, jiraEmail, jiraApiToken)
 */
function postJiraApprovalComment_(issueKey, commentText, cfg) {
  if (!issueKey) throw new Error("postJiraApprovalComment_: issueKey is required.");
  if (!cfg || !cfg.jiraBaseUrl) throw new Error("postJiraApprovalComment_: cfg.jiraBaseUrl missing.");

  const url =
    String(cfg.jiraBaseUrl).replace(/\/$/, "") +
    "/rest/api/3/issue/" + encodeURIComponent(issueKey) + "/comment";

  const payload = {
    body: {
      type: "doc",
      version: 1,
      content: [
        {
          type: "paragraph",
          content: [{ type: "text", text: commentText }]
        }
      ]
    }
  };

  const resp = UrlFetchApp.fetch(url, {
    method: "post",
    headers: getAuthHeader(cfg),
    contentType: "application/json",
    muteHttpExceptions: true,
    payload: JSON.stringify(payload)
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(
      "Jira POST /comment for " + issueKey + " failed: " + code + " " + resp.getContentText()
    );
  }
}

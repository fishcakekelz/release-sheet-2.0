//Pre-Release Check Report
const CONFIG = {
  SOURCE_SHEET_NAME: "ReleaseSheet",
  OUTPUT_SHEET_NAME: "Release Report",
  START_ROW: 2,

  JIRA_TICKET_COL: 6,   // F
  RESCHEDULED_COL: 20,  // Column T (Validation Status) used for Reschedule/Cancel exclusion

  ALL_CRS_SHEET_NAME: "All CRs for Today",
  ALL_CRS_START_ROW: 2,
  ALL_CRS_SCRUM_TEAM_COL: 32,
  ALL_CRS_CHG_COL: 3,
  ALL_CRS_BRACKET_COL: 5,
  ALL_CRS_START_TIME_COL: 8,

  ALL_CRS_LIGHT_YELLOW_BG: "#fff2cc",

  PASS_BG_COLOR: "#C6EFCE",
  DEFAULT_BG_COLOR: "#FFFFFF",
  DEFAULT_FONT_COLOR: "#000000",

  PRE_RELEASE_CHECK_SUMMARY_HEADER: "Pre-Release Check Summary",
  RELEASE_SHEET_JIRA_ITEM_HEADER: "JIRA Item",
  PASS_SUMMARY_TEXT: "PASS"
};

const COL = {
  SCRUM_TEAM: 5,
  CHANGE_TICKET: 10,
  SN_TICKET: 10,
  APPROVAL: 14,
  JIRA_CHANNEL: 28,
  DARK_DEPLOYMENT: 15,
  ISSUE_TYPE: 30,
  COL_I_DATE: 29,
  PROD_RELEASE_DT: 11,
  JIRA_STATUS: 8,
  RESOLUTION: 9,
  LINKED_TICKETS: 36,
  ASSIGNMENT_GROUP: 17,
  CHANGE_PLAN: 38,
  ROLLBACK_PLAN: 39,
  VALIDATION_OWNER: 13,
  VALIDATION_PLAN: 40,
  ACCEPTANCE_CRITERIA: 26,
  AFFECTED_GROUPS: 43,
  AFFECTED_LOCATIONS: 44,
  CTASKS: 19,
  QA_ARTIFACTS_LINK: 33,
  UAT_ARTIFACTS_LINK: 34,
  JIRA_DESCRIPTION: 25
};

/**
 * Menu entry point. Shows a confirmation prompt before running.
 * Also callable programmatically:
 *   - { silent: true }    skip the YES/NO prompt (used by Refresh)
 *   - { skipFlush: true } skip the internal SpreadsheetApp.flush() (caller flushed)
 *
 * Performance refactor:
 *   - 1 source-sheet read (header + all data rows) instead of 2.
 *   - 1 main loop (was 2 passes over the same data).
 *   - Pre-Release Check Summary write-back happens inline using the same
 *     read; no second header lookup, no second JIRA Item read.
 *   - Drops redundant clear-style setBackground/setFontColor calls before
 *     the per-row arrays.
 *   - Drops slow autoResizeColumns; column widths persist across runs.
 *   - Smart partial clear of Release Report keeps headers + freeze row, so
 *     we no longer re-set them on every run.
 *
 * @param {{silent?: boolean, skipFlush?: boolean}} [opts]
 * @returns {{
 *   failuresByJiraKey: Map<string,string>,
 *   passedJiraKeys: Set<string>,
 *   passCount: number,
 *   failCount: number,
 *   ran: boolean
 * }}
 */
function generateAuditReport(opts) {
  opts = opts || {};
  const silent = !!opts.silent;
  const skipFlush = !!opts.skipFlush;

  if (!silent) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      "Run Release Report",
      "Before running the Release Report, please ensure:\n" +
        "1) JIRA details are refreshed\n" +
        "2) SNOWPull details are refreshed\n\n" +
        "Do you want to proceed?",
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      ui.alert("Release Report cancelled. Please refresh JIRA and SNOWPull details, then run again.");
      return { failuresByJiraKey: new Map(), passedJiraKeys: new Set(), passCount: 0, failCount: 0, ran: false };
    }
  }

  const __t0 = Date.now();
  const __mark = (label, since) => Logger.log("[Audit] " + label + ": " + (Date.now() - since) + "ms");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  if (!source) throw new Error("Source sheet not found: " + CONFIG.SOURCE_SHEET_NAME);

  const lastRow = source.getLastRow();
  const lastCol = source.getLastColumn();

  const out = ensureSheet_(ss, CONFIG.OUTPUT_SHEET_NAME);

  // Capture existing notes BEFORE we touch the report sheet.
  const __tNotes = Date.now();
  const existingNotesByKey = captureExistingNotes_(out);
  __mark("captureExistingNotes_", __tNotes);

  // --- Smart clear of Release Report ---
  // Keeps headers + frozen row + column widths, so we don't re-set them.
  // Headers are only (re)written if A1 isn't already "Scrum team".
  const __tClear = Date.now();
  const reportLastRow = out.getLastRow();
  if (reportLastRow >= 2) {
    out.getRange(2, 1, reportLastRow - 1, 5).clearContent();
  }
  if (out.getRange(1, 1).getValue() !== "Scrum team") {
    out.getRange(1, 1, 1, 5).setValues([[
      "Scrum team",
      "Change Ticket #",
      "JIRA Ticket",
      "Audit Summary",
      "Notes"
    ]]);
    out.getRange(1, 1, 1, 5).setFontWeight("bold");
    out.setFrozenRows(1);
  }
  __mark("Release Report partial clear", __tClear);

  const failuresByJiraKey = new Map();
  const passedJiraKeys = new Set();

  if (lastRow < CONFIG.START_ROW) {
    out.getRange(2, 1).setValue("No data rows found on the ReleaseSheet.");
    return { failuresByJiraKey, passedJiraKeys, passCount: 0, failCount: 0, ran: true };
  }

  if (!skipFlush) SpreadsheetApp.flush();

  // ----- Single source read: headers + all data rows in one call -----
  const __tRead = Date.now();
  const allDisplayValues = source.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  __mark("source.getDisplayValues (1 call)", __tRead);

  const headers = (allDisplayValues[0] || []).map(h => (h || "").toString().trim());
  const values = allDisplayValues.slice(CONFIG.START_ROW - 1);
  const numRows = values.length;

  const headerMap = {};
  headers.forEach((h, i) => { if (h) headerMap[h] = i; });
  const summaryColIdx = headerMap[CONFIG.PRE_RELEASE_CHECK_SUMMARY_HEADER];

  if (summaryColIdx === undefined) {
    Logger.log("generateAuditReport: '" + CONFIG.PRE_RELEASE_CHECK_SUMMARY_HEADER +
      "' column not found on ReleaseSheet — write-back to that column will be skipped.");
  }

  const jiraTicketRange = source.getRange(CONFIG.START_ROW, CONFIG.JIRA_TICKET_COL, numRows, 1);

  // Preallocate background/font arrays (faster than Array.from with closures).
  const backgrounds = new Array(numRows);
  const fontColors = new Array(numRows);
  for (let i = 0; i < numRows; i++) {
    backgrounds[i] = [CONFIG.DEFAULT_BG_COLOR];
    fontColors[i] = [CONFIG.DEFAULT_FONT_COLOR];
  }

  // Per-row Pre-Release Check Summary values for the single setValues
  // write at the end. Defaults to current display value (= leave-as-is).
  const newSummaryColumn = (summaryColIdx !== undefined) ? new Array(numRows) : null;

  const failResults = [];
  let passCount = 0;
  let failCount = 0;

  const jiraSeen = new Map();
  const releaseSheetCRs = new Set();

  const __tLoop = Date.now();
  // ----- Single pass: build releaseSheetCRs + apply rules + build summary array -----
  for (let i = 0; i < numRows; i++) {
    const row = values[i];

    if (newSummaryColumn) {
      // Default: leave the existing cell value alone.
      newSummaryColumn[i] = [row[summaryColIdx]];
    }

    const jiraTicket = cleanCell_(row[CONFIG.JIRA_TICKET_COL - 1]);
    const statusTextM = cleanCell_(row[CONFIG.RESCHEDULED_COL - 1]);
    const isExcluded = /(Reschedul(e|ed|ing)?|Cancel(l?ed|ling)?)/i.test(statusTextM);

    // Track CR for "All CRs for Today" cross-check (matches old first-pass semantics).
    if (jiraTicket && !isExcluded) {
      const cr = cleanCell_(row[COL.CHANGE_TICKET - 1]);
      if (cr) releaseSheetCRs.add(normalizeKey_(cr));
    }

    if (!jiraTicket || isExcluded) continue;

    const scrumTeam = cleanCell_(row[COL.SCRUM_TEAM - 1]);
    const changeTicket = cleanCell_(row[COL.CHANGE_TICKET - 1]);

    const failures = [];
    const jiraKey = jiraTicket.trim();
    const jiraKeyLower = jiraKey.toLowerCase();

    const first = jiraSeen.get(jiraKeyLower);
    if (first) {
      failures.push("Duplicate JIRA ticket exists: " + first.original);
    } else {
      jiraSeen.set(jiraKeyLower, { original: jiraKey, firstRowIndex: i });
    }

    failures.push.apply(failures, evaluateRulesDisplayRow_(row));

    if (failures.length === 0) {
      backgrounds[i][0] = CONFIG.PASS_BG_COLOR;
      passCount++;
      if (jiraKeyLower) passedJiraKeys.add(jiraKeyLower);
      if (newSummaryColumn) newSummaryColumn[i] = [CONFIG.PASS_SUMMARY_TEXT];
    } else {
      const summaryText = failures.join("\n");
      failResults.push([scrumTeam, changeTicket, jiraTicket, summaryText, ""]);
      failCount++;
      if (jiraKeyLower) failuresByJiraKey.set(jiraKeyLower, summaryText);
      if (newSummaryColumn) newSummaryColumn[i] = [summaryText];
    }
  }
  __mark("Single-pass main loop (" + numRows + " rows)", __tLoop);

  // ----- "All CRs for Today" missing-CR rule (unchanged semantics) -----
  const __tAllCrs = Date.now();
  const allCrsSheet = ss.getSheetByName(CONFIG.ALL_CRS_SHEET_NAME);
  if (allCrsSheet) {
    const allLastRow = allCrsSheet.getLastRow();
    if (allLastRow >= CONFIG.ALL_CRS_START_ROW) {
      const allNumRows = allLastRow - CONFIG.ALL_CRS_START_ROW + 1;

      const allBtoE = allCrsSheet
        .getRange(CONFIG.ALL_CRS_START_ROW, CONFIG.ALL_CRS_SCRUM_TEAM_COL, allNumRows, 4)
        .getDisplayValues();

      const allStartTimes = allCrsSheet
        .getRange(CONFIG.ALL_CRS_START_ROW, CONFIG.ALL_CRS_START_TIME_COL, allNumRows, 1)
        .getValues();

      const chgBgs = allCrsSheet
        .getRange(CONFIG.ALL_CRS_START_ROW, CONFIG.ALL_CRS_CHG_COL, allNumRows, 1)
        .getBackgrounds();

      const yellowSet = new Set([
        CONFIG.ALL_CRS_LIGHT_YELLOW_BG,
        "#ffe599",
        "#ffd966",
        "#ffff00"
      ]);

      for (let r = 0; r < allNumRows; r++) {
        const scrumTeamFromAll = cleanCell_(allBtoE[r][0]);
        const chg = cleanCell_(allBtoE[r][1]);
        const bracketVal = cleanCell_(allBtoE[r][3]);
        if (!chg) continue;

        const bg = String(chgBgs[r][0] || "").toLowerCase();
        if (!yellowSet.has(bg)) continue;

        const timeInfo = coerceTimeInfo_(allStartTimes[r][0]);
        if (!timeInfo) continue;

        const onOrAfter3pm = (timeInfo.hours > 15) || (timeInfo.hours === 15 && timeInfo.minutes >= 0);
        if (!onOrAfter3pm) continue;

        const chgKeyLower = normalizeKey_(chg);
        if (!releaseSheetCRs.has(chgKeyLower)) {
          const bracket = bracketVal ? " [" + bracketVal + "]" : " [(blank)]";
          failResults.push([
            scrumTeamFromAll,
            chg,
            "",
            chg + bracket + " not on today's release sheet.",
            ""
          ]);
          failCount++;
        }
      }
    }
  }
  __mark("All CRs for Today scan", __tAllCrs);

  // ----- Apply highlights on JIRA Item col on ReleaseSheet (single batch each) -----
  const __tHighlight = Date.now();
  jiraTicketRange.setBackgrounds(backgrounds);
  jiraTicketRange.setFontColors(fontColors);
  __mark("setBackgrounds + setFontColors (JIRA Item col)", __tHighlight);

  // ----- Write Release Report rows + restore notes -----
  const __tWriteReport = Date.now();
  const PST_TZ = "America/Los_Angeles";
  const ranAtPST = Utilities.formatDate(new Date(), PST_TZ, "MM/dd/yyyy hh:mm:ss a");
  const ranByEmail = Session.getActiveUser().getEmail() || "Unknown user";

  if (failResults.length === 0) {
    out.getRange(2, 1).setValue(
      "All audited JIRA tickets passed. Highlighted " + passCount + " ticket(s) in green on Release Sheet."
    );
    const summaryRow = 4;
    out.getRange(summaryRow, 1).setValue("Summary: " + passCount + " passed (highlighted), 0 failed (listed).");
    out.getRange(summaryRow + 1, 1).setValue("Report last ran by " + ranByEmail + " at: " + ranAtPST + " PST");
  } else {
    out.getRange(2, 1, failResults.length, 5).setValues(failResults);
    // Single setWrap on cols D + E together (was 2 separate calls).
    out.getRange(2, 4, failResults.length, 2).setWrap(true);

    restoreNotes_(out, failResults.length, existingNotesByKey);

    const summaryRow = failResults.length + 3;
    out.getRange(summaryRow, 1).setValue(
      "Summary: " + passCount + " passed (highlighted), " + failCount + " failed (listed)."
    );
    out.getRange(summaryRow + 1, 1).setValue("Report last ran by " + ranByEmail + " at: " + ranAtPST + " PST");
  }
  __mark("Release Report write", __tWriteReport);

  // ----- Pre-Release Check Summary write-back (single setValues to ReleaseSheet) -----
  if (newSummaryColumn) {
    const __tWriteSummary = Date.now();
    source.getRange(CONFIG.START_ROW, summaryColIdx + 1, numRows, 1).setValues(newSummaryColumn);
    __mark("Pre-Release Check Summary setValues (" + numRows + " rows)", __tWriteSummary);
  }

  __mark("TOTAL Audit", __t0);
  return { failuresByJiraKey, passedJiraKeys, passCount, failCount, ran: true };
}

function evaluateRulesDisplayRow_(row) {
  const errors = [];

  const issueType = cleanCell_(row[COL.ISSUE_TYPE - 1]);
  const acceptanceCriteria = cleanCell_(row[COL.ACCEPTANCE_CRITERIA - 1]);
  const jiraDescriptionAT = cleanCell_(row[COL.JIRA_DESCRIPTION - 1]);
  const jiraChannel = cleanCell_(row[COL.JIRA_CHANNEL - 1]);
  const affectedGroups = cleanCell_(row[COL.AFFECTED_GROUPS - 1]);
  const darkDeployment = cleanCell_(row[COL.DARK_DEPLOYMENT - 1]);
  const jiraStatus = cleanCell_(row[COL.JIRA_STATUS - 1]);
  const resolution = cleanCell_(row[COL.RESOLUTION - 1]);
  const linkedTickets = cleanCell_(row[COL.LINKED_TICKETS - 1]);
  const validationOwner = cleanCell_(row[COL.VALIDATION_OWNER - 1]);
  const approval = cleanCell_(row[COL.APPROVAL - 1]);
  const rescheduledText = cleanCell_(row[CONFIG.RESCHEDULED_COL - 1]);
  const colI = cleanCell_(row[COL.COL_I_DATE - 1]);
  const colJ = cleanCell_(row[COL.PROD_RELEASE_DT - 1]);
  const assignmentGroup = cleanCell_(row[COL.ASSIGNMENT_GROUP - 1]);
  const changePlan = cleanCell_(row[COL.CHANGE_PLAN - 1]);
  const rollbackPlan = cleanCell_(row[COL.ROLLBACK_PLAN - 1]);
  const validationPlan = cleanCell_(row[COL.VALIDATION_PLAN - 1]);
  const affectedLocations = cleanCell_(row[COL.AFFECTED_LOCATIONS - 1]);
  const qaArtifactsLink = cleanCell_(row[COL.QA_ARTIFACTS_LINK - 1]);
  const uatArtifactsLink = cleanCell_(row[COL.UAT_ARTIFACTS_LINK - 1]);
  const ctasksText = cleanCell_(row[COL.CTASKS - 1]);
  const jiraTicket = cleanCell_(row[CONFIG.JIRA_TICKET_COL - 1]);

  if (isChannelAffectedGroupsMismatch_(jiraChannel, affectedGroups)) {
    const fVal = jiraChannel || "(blank)";
    const anVal = affectedGroups || "(blank)";
    errors.push(
      "Discrepancy between JIRA Channel [" + fVal + "] and CR Affected Groups [" + anVal + "]."
    );
  }

  if (equalsAny_(issueType, ["Story", "Task"]) && !acceptanceCriteria) {
    const descriptionMentionsAC = /Acceptance\s*Criteria/i.test(jiraDescriptionAT);
    if (!descriptionMentionsAC) {
      errors.push("Please verify CR validation plan;Missing Acceptance Criteria for Story/Task Tickets.");
    }
  }

  if (equalsAny_(issueType, ["Bug"])) {
    const hasAcceptanceCriteria = /Acceptance\s*Criteria/i.test(jiraDescriptionAT);
    const hasExpectedResult = /Expected\s*Result/i.test(jiraDescriptionAT);
    const hasExpectedOutcome = /Expected\s*Outcome/i.test(jiraDescriptionAT);
    const hasExpectedBehavior = /Expected\s*Behavior/i.test(jiraDescriptionAT);
    const hasDesiredResult =
      /Desired\s*Result/i.test(jiraDescriptionAT) ||
      /Expected\s*\/\s*Desired\s*Result/i.test(jiraDescriptionAT);
    const hasDesiredOutcome =
      /Desired\s*Outcome/i.test(jiraDescriptionAT) ||
      /Expected\s*\/\s*Desired\s*Outcome/i.test(jiraDescriptionAT);
    const hasExpectedMarker = /Expected\s*:/i.test(jiraDescriptionAT);
    const hasDesiredMarker = /Desired\s*:/i.test(jiraDescriptionAT);

    if (
      !hasAcceptanceCriteria &&
      !hasExpectedResult &&
      !hasExpectedBehavior &&
      !hasExpectedOutcome &&
      !hasDesiredResult &&
      !hasDesiredOutcome &&
      !hasExpectedMarker &&
      !hasDesiredMarker
    ) {
      errors.push("Please verify CR validation plan; Missing Acceptance Criteria/ Expected Result/ Expected Outcome/ Expected Behavior for Bug Ticket.");
    }
  }

  const isDarkDeployment = /^x$/i.test(darkDeployment);
  if (!isDarkDeployment) {
    if (!equalsAny_(jiraStatus, ["Done"]) && !equalsAny_(resolution, ["Ready for Production"])) {
      errors.push("JIRA not in Ready for Production.");
    }
  }

  if (linkedTickets) {
    errors.push(("Blocker tickets exist - " + linkedTickets).trim());
  }

  if (!validationOwner) errors.push("Missing Validation Owner.");

  const requiresPMApproval = /^(DCX-|LEADG-|POP-|CDM-|CE-|CP-|KCP-|KBF-|DLL-|AITECH-)/i.test(jiraTicket);
  if (requiresPMApproval && !approval) {
    errors.push("Missing Product Manager approval.");
  }

  const snTicketRaw = cleanCell_(row[COL.SN_TICKET - 1]);
  const snTicketIsMissing = !snTicketRaw || /CR\s*not\s*entered/i.test(snTicketRaw);

  if (!/Rescheduled/i.test(rescheduledText) && snTicketIsMissing) {
    errors.push("Missing ServiceNow Ticket # field.");
  }

  const iDateStr = normalizeMMDDYYYY_(colI);
  const jDateStr = extractDateFromJ_(colJ);
  if (iDateStr && jDateStr && iDateStr !== jDateStr) {
    errors.push("CR Start date (" + jDateStr + ") not same as JIRA Production Release Date (" + iDateStr + ").");
  }

  const timeInfo = extractTimeFromJ_(colJ);
  if (timeInfo && timeInfo.hours < 17) {
    errors.push("Mid-day CRs need EVP+ approval attached on the CR.");
  }

  if (/Operations Center/i.test(assignmentGroup)) {
    const hasTagDetails = /(gitlab|jenkins|tfs|tag|launchdarkly)/i.test(changePlan);
    const mentionsAttachment = /(attach|attached|attachment)/i.test(changePlan);
    const mentionsCTASK = /ctask/i.test(changePlan);

    if (!mentionsCTASK && !hasTagDetails && !mentionsAttachment) {
      errors.push("Missing deployment details in the Change plan for Operations assigned CR.");
    }
  }

  if (/^DB Operations$/i.test(assignmentGroup)) {
    const hasGitlabOrTag = /(gitlab|jenkins|tag)/i.test(changePlan);
    const mentionsCTASK = /ctask/i.test(changePlan);
    const mentionsScript = /script/i.test(changePlan);

    if (!hasGitlabOrTag && !mentionsCTASK && !mentionsScript) {
      errors.push("Missing deployment details in the Change plan for DB Operations assigned CR.");
    }
  }

  if (/^IT Releasement$/i.test(assignmentGroup)) {
    const mentionsCTASK = /ctask/i.test(changePlan);
    const mentionsConfig = /config/i.test(changePlan);

    if (!mentionsCTASK && !mentionsConfig) {
      errors.push("Missing deployment details in the Change plan for ITRM assigned CR.");
    }
  }

  const rollbackMentionsCTASK = /ctask/i.test(rollbackPlan);
  if (!rollbackMentionsCTASK && (!rollbackPlan || /^\s*TBD(E)?\s*$/i.test(rollbackPlan))) {
    errors.push("Missing Rollback plan details in the CR.");
  }

  if (validationPlan && /^\s*TBD(E)?\s*$/i.test(validationPlan)) {
    errors.push("Missing Validation plan details in the CR.");
  }

  if (!affectedGroups) errors.push("Missing Affected Groups in the CR.");
  if (/^All$/i.test(affectedGroups)) errors.push("Please confirm if Affected Groups = All is the correct value.");

  if (!affectedLocations) errors.push("Missing Affected Locations in the CR.");

  if (equalsAny_(jiraStatus, ["Done"]) && equalsAny_(resolution, ["Ready for Production"])) {
    if (!qaArtifactsLink && !uatArtifactsLink) {
      errors.push("Please verify JIRA comments for lower env testing details; Missing QA & UAT Artifacts link fields.");
    }
  }

  const hasCtasks = !!ctasksText && !/No\s*CTASKs\s*---/i.test(ctasksText);
  const ctasksMissingDetails = /---|--/.test(ctasksText);
  if (hasCtasks && ctasksMissingDetails) {
    errors.push("Missing Assignment Group or Assigned To or Due Dates on CTASKs.");
  }

  return errors;
}

function normalizeKey_(s) {
  return cleanCell_(s).replace(/\s+/g, "").toLowerCase();
}

function coerceTimeInfo_(v) {
  if (v === null || v === undefined || v === "") return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    return { hours: v.getHours(), minutes: v.getMinutes() };
  }
  return extractTimeFromH_(cleanCell_(v));
}

function extractTimeFromH_(timeStr) {
  if (!timeStr) return null;
  const s = String(timeStr).trim();

  let m = s.match(/(\d{1,2})(?::(\d{2}))?\s*([AaPp][Mm])\b/);
  if (m) {
    let hours = parseInt(m[1], 10);
    const minutes = parseInt(m[2] || "0", 10);
    const ampm = (m[3] || "").toUpperCase();
    if (ampm === "AM") {
      if (hours === 12) hours = 0;
    } else if (ampm === "PM") {
      if (hours !== 12) hours += 12;
    }
    return { hours, minutes };
  }

  m = s.match(/\b(\d{1,2})(?::(\d{2}))\b/);
  if (m) {
    const hours = parseInt(m[1], 10);
    const minutes = parseInt(m[2] || "0", 10);
    if (Number.isNaN(hours) || Number.isNaN(minutes)) return null;
    return { hours, minutes };
  }

  return null;
}

function isChannelAffectedGroupsMismatch_(jiraChannel, affectedGroups) {
  const channel = (jiraChannel || "").toString();
  const groups = (affectedGroups || "").toString();
  if (!channel.trim() || !groups.trim()) return false;

  const hasCDL = /CDL/i.test(channel);
  const hasPCG = /PCG/i.test(channel);
  const hasTPO = /TPO/i.test(channel);
  const hasRetail = /Retail/i.test(groups);
  const hasCorrespondent = /Correspondent/i.test(groups);
  const hasBroker = /Broker/i.test(groups);

  if (hasCDL && !hasRetail) return true;
  if (hasPCG && !hasCorrespondent) return true;
  if (hasTPO && !hasBroker) return true;

  return false;
}

function cleanCell_(v) {
  if (v === null || v === undefined) return "";

  let s = String(v)
    .replace(/[\u00A0]/g, " ")
    .replace(/[\u200B-\u200D\uFEFF]/g, "")
    .trim();

  if (!s) return "";

  s = s.replace(/^""+(?=[^"])/, "").trim();
  if (s.length >= 2 && s.startsWith('"') && s.endsWith('"')) {
    s = s.slice(1, -1).trim();
  }
  if (!s) return "";

  if (/^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(s)) return "";

  const upper = s.toUpperCase();
  if (upper === "N/A" || upper === "NA") return "";

  return s;
}

function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function equalsAny_(value, options) {
  const v = (value || "").trim().toLowerCase();
  return options.some(o => (o || "").trim().toLowerCase() === v);
}

function normalizeMMDDYYYY_(s) {
  if (!s) return "";
  const m = String(s).trim();
  if (/^\d{4}\-\d{2}\-\d{2}$/.test(m)) return m;
  const d = new Date(m);
  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-mm-dd");
}

function extractDateFromJ_(jStr) {
  if (!jStr || jStr.length < 10) return "";
  const datePart = jStr.substring(0, 10).trim();

  let d;
  if (/^\d{4}-\d{2}-\d{2}$/.test(datePart)) d = new Date(datePart + "T00:00:00");
  else d = new Date(datePart);

  if (isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function extractTimeFromJ_(jStr) {
  if (!jStr || jStr.length === 0) return null;
  const m = String(jStr).match(/(\d{1,2}):(\d{2})\s*([AaPp][Mm])?/);
  if (!m) return null;

  let hours = parseInt(m[1], 10);
  const minutes = parseInt(m[2], 10);
  const ampm = (m[3] || "").toUpperCase();
  if (Number.isNaN(hours) || Number.isNaN(minutes)) return null;

  if (ampm === "AM") {
    if (hours === 12) hours = 0;
  } else if (ampm === "PM") {
    if (hours !== 12) hours += 12;
  }

  return { hours, minutes, ampm };
}

function captureExistingNotes_(reportSheet) {
  const lastRow = reportSheet.getLastRow();
  if (lastRow < 2) return new Map();

  const lastCol = reportSheet.getLastColumn();
  if (lastCol < 5) return new Map();

  const data = reportSheet.getRange(2, 2, lastRow - 1, 4).getDisplayValues();

  const notesByKey = new Map();
  for (let i = 0; i < data.length; i++) {
    const changeTicket = cleanCell_(data[i][0]);
    const jiraTicket = cleanCell_(data[i][1]);
    const note = cleanCell_(data[i][3]);
    if (!note) continue;

    const key = makeCrJiraKey_(changeTicket, jiraTicket);
    if (!key) continue;

    if (!notesByKey.has(key)) notesByKey.set(key, note);
  }
  return notesByKey;
}

function restoreNotes_(reportSheet, numReportRows, notesByKey) {
  if (!numReportRows || numReportRows < 1) return;
  if (!notesByKey || notesByKey.size === 0) return;

  const bc = reportSheet.getRange(2, 2, numReportRows, 2).getDisplayValues();

  const notesToWrite = [];
  for (let i = 0; i < bc.length; i++) {
    const changeTicket = cleanCell_(bc[i][0]);
    const jiraTicket = cleanCell_(bc[i][1]);

    const key = makeCrJiraKey_(changeTicket, jiraTicket);
    const note = key ? (notesByKey.get(key) || "") : "";

    notesToWrite.push([note]);
  }

  reportSheet.getRange(2, 5, numReportRows, 1).setValues(notesToWrite);
}

function makeCrJiraKey_(changeTicket, jiraTicket) {
  const cr = cleanCell_(changeTicket);
  const jira = cleanCell_(jiraTicket);
  if (!cr && !jira) return "";
  return normalizeKey_(cr) + "||" + normalizeKey_(jira);
}

/**
 * Entry point used by Refresh Release Sheet's deferred tail.
 *
 * Runs the audit silently. The audit itself now writes the Pre-Release
 * Check Summary column back to ReleaseSheet inline as part of the same
 * single-pass loop, so no second sheet read or second write is needed.
 *
 * `skipFlush: true` because the refresh tail has already flushed.
 *
 * Wrapped in try/catch so an audit failure cannot break Refresh.
 */
function runAuditAndApplySummaryAfterRefresh_() {
  try {
    generateAuditReport({ silent: true, skipFlush: true });
  } catch (err) {
    Logger.log("runAuditAndApplySummaryAfterRefresh_ failed: " + (err && err.message ? err.message : err));
  }
}

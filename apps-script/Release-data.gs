/**********************
  Release Sheet Script
  - Create Release Sheet
  - Refresh Release Sheet (appends new, highlights modified/removed)
  - Logs edits in ChangeLog
  - Date locking to prevent accidental date changes
**********************/

/**
 * If y0 (yyyy-MM-dd) is a Friday in LA, returns that Fri + the following Sat & Sun; otherwise [y0] only.
 * Used for Jira "Production Release Date" JQL and for matching SNOW start_date to the same window.
 * Do not remove — buildJqlForDate and create/refresh depend on this.
 * @param {string|Date} y0
 * @returns {string[]}
 */
function expandFridayWeekendJqlYmdsForY0_(y0) {
  const la = "America/Los_Angeles";
  if (y0 instanceof Date && !isNaN(y0.getTime())) {
    y0 = Utilities.formatDate(y0, la, "yyyy-MM-dd");
  }
  if (!y0 || String(y0).indexOf("-") < 0) {
    if (y0) return [String(y0).trim()];
    return [""];
  }
  const ymd = String(y0).trim();
  const parsed = Utilities.parseDate(ymd + " 12:00:00", la, "yyyy-MM-dd HH:mm:ss");
  if (!parsed || isNaN(parsed.getTime())) return [ymd];
  const dow = Number(Utilities.formatDate(parsed, la, "u"));
  if (dow !== 5) return [ymd];
  return [0, 1, 2].map(n =>
    Utilities.formatDate(
      new Date(parsed.getTime() + n * 24 * 60 * 60 * 1000),
      la,
      "yyyy-MM-dd"
    )
  );
}

/* =================== onOpen menu =================== */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("ITRM Functions")
      .addItem("Manually Create Release Sheet", "createReleaseSheet")
      .addItem("Refresh Release Sheet", "refreshReleaseSheet")
      .addItem("CLOSE SHEET", "StopRelease")
      .addSeparator()
      .addItem('SNowPull', 'snowPull')
      .addItem('CTASKPull', 'pullCTASKTab')
      .addItem('Pull CR + CTASK Info', 'runDailyCRAndCTASKPull')
      .addSeparator()
      .addItem("Archive to Summary", "archiveReleaseToSummary")
      .addItem("Archive & Clear Release Sheet", "archiveAndClearRelease")
      .addItem("Stop Release", "StopRelease")
      .addSeparator()
      // runDailyReleaseCheck: see DailyReleaseCheck.gs (separate file; not duplicated here)
      .addItem('Daily Release Check', 'runDailyReleaseCheck')
      .addSeparator()
      .addItem('Send PCG Channel Email', 'testPCGChannelEmail')
      .addSeparator()
      .addSubMenu(ui.createMenu("Approval Comments")
        .addItem("Post Comments For All Approved Rows", "processApprovedRowsBulk")
        .addSeparator()
        .addItem("Install Trigger (auto-post on edit)", "installApprovalEditTrigger")
        .addItem("Remove Trigger", "removeApprovalEditTrigger"))
      .addSeparator()
      .addSubMenu(ui.createMenu("Scheduled Triggers")
        .addItem("Enable Daily Refresh (9 AM, 12 PM, 2 PM)", "createDailyTriggers")
        .addItem("Disable Daily Refresh", "removeDailyTriggersWithAlert")
        .addSeparator()
        .addItem("🧪 Start Test (every 5 min)", "createTestTrigger")
        .addItem("🧪 Stop Test", "removeTestTrigger"))
      .addSeparator()
      .addItem("Populate ReleaseSheet Backup", "populateReleaseSheetBackup")
      .addItem("Set up Manual Backup Runbook", "setupBackupRunbook")
      .addToUi();
    
  } catch (e) {
    // UI not available (e.g., running from script editor or trigger context)
    Logger.log('onOpen: UI not available - ' + e.message);
  }
}

/* =================== Utilities =================== */
function getConfig() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(S_CONFIG);
  if (!sheet) throw new Error("Config sheet not found.");
  const cfg = {};
  sheet.getDataRange().getValues().forEach(r => r[0] && (cfg[r[0]] = r[1]));
  cfg.sheetHeaderRow = Number(cfg.sheetHeaderRow || 1);
  // Optional speed-up for menu Refresh: set to true to skip protectManualColumns (re-locks green columns).
  const skip = cfg.refreshSkipProtectOnManualRefresh;
  cfg.refreshSkipProtectOnManualRefresh =
    skip === true || String(skip).toLowerCase() === "true";
  // Defer protections + row-height + font-reset to a follow-up trigger so the
  // user-visible "Refresh complete" alert pops as soon as the data is on the
  // sheet. Defaults to true; set Config.refreshDeferTail = false to revert to
  // running the tail inline. Headless (scheduled) refreshes also honor this.
  const defer = cfg.refreshDeferTail;
  cfg.refreshDeferTail =
    defer === undefined || defer === "" || defer === null
      ? true
      : !(defer === false || String(defer).toLowerCase() === "false");
  // Speed-up: only write the Master "debug dump" sheet when explicitly enabled.
  // The Master tab is purely diagnostic (Key/Summary/JSON of every issue) and
  // skipping it saves 1-3 seconds per Create for typical release sizes.
  const dump = cfg.writeMasterDebugDump;
  cfg.writeMasterDebugDump =
    dump === true || String(dump).toLowerCase() === "true";
  const defPx =
    typeof RELEASE_SHEET_DATA_ROW_HEIGHT_PX === "number" && RELEASE_SHEET_DATA_ROW_HEIGHT_PX > 0
      ? RELEASE_SHEET_DATA_ROW_HEIGHT_PX
      : 35;
  const rawPx = cfg.releaseDataRowHeight;
  const px = rawPx === "" || rawPx === null || rawPx === undefined ? defPx : Number(rawPx);
  cfg.releaseDataRowHeight = !isNaN(px) && px >= 8 && px <= 409 ? px : defPx;
  return cfg;
}

function getAuthHeader(cfg) {
  const raw = `${cfg.jiraEmail}:${cfg.jiraApiToken}`;
  return {
    Authorization: "Basic " + Utilities.base64Encode(raw),
    Accept: "application/json",
    "Content-Type": "application/json"
  };
}

function getReleaseDateFromReleaseSheet() {
  const cell = SpreadsheetApp.getActive().getSheetByName(S_RELEASE).getRange("A1").getValue();
  if (!cell) throw new Error("Release date missing in A1.");
  return Utilities.formatDate(parseReleaseDateValue(cell), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function parseReleaseDateValue(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const rawValue = String(value).trim();

  if (!rawValue) {
    throw new Error("Release date missing in A1.");
  }

  let match = rawValue.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match) {
    return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  }

  match = rawValue.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (match) {
    const month = Number(match[1]);
    const day = Number(match[2]);
    const year = match[3]
      ? (match[3].length === 2 ? 2000 + Number(match[3]) : Number(match[3]))
      : new Date().getFullYear();
    return new Date(year, month - 1, day);
  }

  throw new Error(`Unsupported release date format in A1: ${rawValue}`);
}

/* =================== Date Locking (uses Script Properties to avoid protection issues) =================== */

function getLockedReleaseDate() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty("lockedReleaseDate") || null;
}

function setLockedReleaseDate(dateStr) {
  const props = PropertiesService.getScriptProperties();
  props.setProperty("lockedReleaseDate", dateStr);
  Logger.log(`Locked release date set to: ${dateStr}`);
}

function clearLockedReleaseDate() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("lockedReleaseDate");
  Logger.log("Locked release date cleared.");
}

/**
 * Validates if the current date matches the locked date.
 * Returns true if OK to proceed, false if user cancelled.
 */
function validateDateChange(currentDate, isHeadless = false) {
  const lockedDate = getLockedReleaseDate();
  
  // If no locked date or dates match, proceed
  if (!lockedDate || lockedDate === currentDate) {
    return true;
  }
  
  // If headless (scheduled trigger), just log and proceed with the locked date
  if (isHeadless) {
    Logger.log(`Warning: Current date in A1 is ${currentDate}, but locked date is ${lockedDate}. Using locked date.`);
    return true;
  }
  
  // Show warning popup
  const ui = SpreadsheetApp.getUi();
 // Parse dates correctly to avoid timezone issues (YYYY-MM-DD format)
  const parseDate = (dateStr) => {
    const parts = dateStr.split("-");
    return new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  };
  
  const currentDateFormatted = Utilities.formatDate(parseDate(currentDate), Session.getScriptTimeZone(), "MMM d");
  const lockedDateFormatted = Utilities.formatDate(parseDate(lockedDate), Session.getScriptTimeZone(), "MMM d");
  
  const response = ui.alert(
    "⚠️ Use a different date?",
    `The date you just picked (${currentDate}) doesn't match the date currently on the sheet (${lockedDate}).\n\n` +
    `Would you like to proceed and update the sheet's locked date to ${currentDateFormatted}?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    // User confirmed, update the locked date
    setLockedReleaseDate(currentDate);
    return true;
  }
  
  // User cancelled, revert the date back to locked date
  const sheet = SpreadsheetApp.getActive().getSheetByName(S_RELEASE);
  sheet.getRange("A1").setValue(lockedDate);
  return false;
}

/* =================== Jira =================== */

/**
 * Jira REST /search `fields` must use API ids (e.g. customfield_10181), not script key names.
 * Accepts legacy "cf[10185]" from configs and normalizes to "customfield_10185".
 * @returns {string|null}
 */
function normalizeJiraFieldIdForApi_(value) {
  if (value === null || value === undefined) return null;
  const s = String(value).trim();
  if (!s) return null;
  const m = s.match(/^cf\[(\d+)\]$/i);
  if (m) return "customfield_" + m[1];
  return s;
}

/** Plain string for URL/text/ADF/single-select style Jira custom fields. */
function jiraFieldAsPlainString_(v) {
  if (v === null || v === undefined) return "";
  if (typeof v === "string" || typeof v === "number" || typeof v === "boolean") {
    return String(v).trim();
  }
  if (Array.isArray(v)) {
    return v
      .map(x => jiraFieldAsPlainString_(x))
      .filter(Boolean)
      .join(", ");
  }
  if (typeof v === "object") {
    if (v.content) return extractPlainDescription(v);
    if (v.value !== undefined && v.value !== null && String(v.value).trim() !== "") {
      return String(v.value).trim();
    }
    if (v.name !== undefined && v.name !== null) return String(v.name).trim();
    if (v.displayName) return String(v.displayName).trim();
  }
  return "";
}

/**
 * "Production Release Date" filter: one day, or IN (3 days) when A1/anchor is a Friday in LA
 */
function buildJqlForDate(dateStr, cfg) {
  const projectClause = cfg.projects ? `project in (${cfg.projects}) AND ` : "";
  // Use quoted field name - customfield_10220 syntax doesn't work in JQL search
  // Filter for AppDev Team only for CAP project (customfield_10254 = Teams)
  const teamsFilter = `(project != "CAP" OR "Teams" = "AppDev Team")`;
  const dates = expandFridayWeekendJqlYmdsForY0_(dateStr);
  const datePart =
    dates.length === 1
      ? `"Production Release Date" = "${dates[0]}"`
      : `"Production Release Date" in (${dates.map(d => `"${d}"`).join(", ")})`;
  return `${projectClause}${datePart} AND ${teamsFilter} ORDER BY key ASC`;
}

function fetchJiraIssuesByJql(jql, cfg) {
  const url = `${cfg.jiraBaseUrl}/rest/api/3/search/jql`;
  const allIssues = [];
  let nextPageToken = null;
  const maxPerPage = 100;  // Jira caps at 100 per request on new endpoint
  
  const fields = jiraFieldsForFetch_();

  // Paginate through all results using cursor-based pagination
  while (true) {
    const payload = {
      jql: jql,
      maxResults: maxPerPage,
      fields: fields
    };
    
    // Add nextPageToken if we have one (for pages after the first)
    if (nextPageToken) {
      payload.nextPageToken = nextPageToken;
    }

    const resp = UrlFetchApp.fetch(url, {
      method: "post",
      headers: getAuthHeader(cfg),
      contentType: "application/json",
      muteHttpExceptions: true,
      payload: JSON.stringify(payload)
    });

    if (resp.getResponseCode() >= 400) {
      throw new Error(`Jira API error ${resp.getResponseCode()}: ${resp.getContentText()}`);
    }

    const result = JSON.parse(resp.getContentText());
    const issues = result.issues || [];
    allIssues.push(...issues);
    
    Logger.log(`Fetched ${allIssues.length} issues so far...`);
    
    // Check if there's a next page
    nextPageToken = result.nextPageToken || null;
    if (!nextPageToken || issues.length === 0) {
      break;
    }
  }
  
  Logger.log(`Total issues fetched: ${allIssues.length}`);

  const kept = [];
  const excludedKeys = [];
  allIssues.forEach(issue => {
    if (isExcludedCapTicket_(issue)) {
      excludedKeys.push(issue.key);
    } else {
      kept.push(issue);
    }
  });
  if (excludedKeys.length) {
    const preview = excludedKeys.slice(0, 25).join(", ");
    const more = excludedKeys.length > 25 ? ` (+${excludedKeys.length - 25} more)` : "";
    Logger.log(
      `Excluded ${excludedKeys.length} ${CAP_EXCLUDE_PROJECT_KEY} ticket(s) ` +
        `matching ${CAP_EXCLUDED_TOKENS.join("/")} in Components or CMDB CI: ${preview}${more}`
    );
  }
  return kept;
}

/**
 * True when the issue is in the configured CAP project AND its Components or
 * CMDB CI field mentions one of CAP_EXCLUDED_TOKENS as a whole token
 * (case-insensitive, separator-aware). Non-CAP tickets are never excluded.
 *
 * Whole-token semantics: we tokenize each candidate value on whitespace and
 * common separators (commas, semicolons, pipes, slashes, parens, brackets),
 * lowercase each token, and compare against the lowercased token set. This
 * means:
 *   - "Luna"                 → excluded (token == "luna")
 *   - "Athena Service"       → excluded ("athena" is a standalone token)
 *   - "Luna, Athena Server"  → excluded (commas split into tokens)
 *   - "PROD-ATHENA-01"       → NOT excluded (hyphens don't split, single token)
 *   - "LunaPark"             → NOT excluded (no separator between "Luna" and "Park")
 *
 * Components are scanned per `name`. CMDB CI is normalized through
 * jiraFieldAsPlainString_ so we don't have to know whether the field is text,
 * single-select, or multi-select.
 */
function isExcludedCapTicket_(issue) {
  if (!issue || !issue.fields) return false;
  const projectKey =
    issue.fields.project && issue.fields.project.key
      ? String(issue.fields.project.key)
      : "";
  if (projectKey !== CAP_EXCLUDE_PROJECT_KEY) return false;
  if (!Array.isArray(CAP_EXCLUDED_TOKENS) || !CAP_EXCLUDED_TOKENS.length) return false;

  const restricted = {};
  CAP_EXCLUDED_TOKENS.forEach(t => {
    if (t) restricted[String(t).toLowerCase()] = true;
  });

  const candidates = [];
  const comps = issue.fields.components;
  if (Array.isArray(comps)) {
    comps.forEach(c => {
      if (c && c.name) candidates.push(String(c.name));
    });
  }
  if (typeof CAP_CMDB_CI_FIELD_ID === "string" && CAP_CMDB_CI_FIELD_ID) {
    const cmdbStr = jiraFieldAsPlainString_(issue.fields[CAP_CMDB_CI_FIELD_ID]);
    if (cmdbStr) candidates.push(cmdbStr);
  }

  for (let i = 0; i < candidates.length; i++) {
    const tokens = String(candidates[i])
      .toLowerCase()
      .split(/[\s,;|/()\[\]]+/)
      .filter(Boolean);
    for (let j = 0; j < tokens.length; j++) {
      if (restricted[tokens[j]]) return true;
    }
  }
  return false;
}

/***********************
 * Normalize Jira Issue
 ***********************/
function normaliseIssue(issue) {
  const f = issue.fields || {};

  const text = v =>
    Array.isArray(v)
      ? v.map(x => x.value || "").join(", ")
      : v?.value || v?.name || v || "";

  const lastComment =
    f.comment?.comments?.length
      ? extractPlainDescription(
          f.comment.comments.slice(-1)[0].body
        )
      : "";

  // ---- Issue links / blockers ----
  const JIRA_BASE_URL = "https://pnmac.atlassian.net/browse/"; // Base Jira URL
  
  const issuelinks = f.issuelinks || [];

  const blockerLinks = issuelinks.filter(link => {
    const inwardType = link.type?.inward?.toLowerCase() || "";
    return inwardType.includes("blocked by") && link.inwardIssue;
  });

  const blockers = blockerLinks.map(link => link.inwardIssue.key);

  // Structured per-blocker info. The Jira "issuelinks" stub typically only
  // carries summary/status/priority/issuetype on inwardIssue.fields — it does
  // NOT include resolution. We seed status from the stub and let
  // enrichBlockerInfosWithResolutions_() fill in resolution via a bulk fetch.
  const blockerInfos = blockerLinks.map(link => ({
    key: link.inwardIssue.key,
    status: link.inwardIssue.fields?.status?.name || "Unknown",
    resolution: link.inwardIssue.fields?.resolution?.name || ""
  }));

  const relatedIssues = issuelinks
    .map(link => link.outwardIssue?.key || link.inwardIssue?.key)
    .filter(Boolean);

  // ---- ServiceNow Related Ticket (customfield_10258): all CRs, expanded to rows later ----
  const rawCR = FIELD_MAP.crNumber ? text(f[FIELD_MAP.crNumber]) : "";
  const parsedCrs = parseRelatedTicketCRs(rawCR);

  return {
    key: issue.key,
    summary: f.summary || "",
    issueType: f.issuetype?.name || "",
    scrumTeam: FIELD_MAP.scrumTeam ? text(f[FIELD_MAP.scrumTeam]) : "",
    workstream: FIELD_MAP.workstream ? text(f[FIELD_MAP.workstream]) : "",
    prodReleaseDate: FIELD_MAP.prodReleaseDate ? text(f[FIELD_MAP.prodReleaseDate]) : "",

    allCrs: parsedCrs.crs,

    approvals: FIELD_MAP.approvals ? text(f[FIELD_MAP.approvals]) : "",
    channel: FIELD_MAP.channel ? text(f[FIELD_MAP.channel]) : "",
    status: f.status?.name || "",
    resolution: f.resolution?.name || "",
    assignee: f.assignee?.displayName || "",
    acceptanceCriteria: FIELD_MAP.acceptanceCriteria
      ? extractPlainDescription(f[FIELD_MAP.acceptanceCriteria])
      : "",
    description: extractPlainDescription(f.description),
    qaArtifactsLink: (() => {
      const id = normalizeJiraFieldIdForApi_(FIELD_MAP && FIELD_MAP.qaArtifactsLink);
      return id ? jiraFieldAsPlainString_(f[id]) : "";
    })(),
    uatArtifactsLink: (() => {
      const id = normalizeJiraFieldIdForApi_(FIELD_MAP && FIELD_MAP.uatArtifactsLink);
      return id ? jiraFieldAsPlainString_(f[id]) : "";
    })(),
    severity: (() => {
      const id = normalizeJiraFieldIdForApi_(FIELD_MAP && FIELD_MAP.severity);
      if (!id || id === "priority") return f.priority?.name || "";
      return jiraFieldAsPlainString_(f[id]) || f.priority?.name || "";
    })(),
    linkedIssues: issuelinks,
    relatedIssues,
    blockers,
    blockerInfos,
    lastComment
  };
}

/**
 * Bulk-fetches the current "Production Release Date" for the given Jira keys.
 * Used by refreshReleaseSheetInternal_ so appended REMOVED rows display the
 * ticket's NEW release date (i.e. where it moved to), not the stale value
 * from the sheet snapshot.
 *
 * Returns { [key]: "YYYY-MM-DD" | "" }. Missing keys (deleted tickets, fetch
 * errors) are simply absent — callers should leave the existing cell value
 * in that case.
 *
 * Mirrors fetchBlockerStatusAndResolutionByKeys_ for pagination, chunking,
 * and error-log + continue semantics.
 *
 * @param {string[]} keys Jira issue keys (deduped recommended).
 * @param {object} cfg getConfig() result (uses jiraBaseUrl + auth).
 * @returns {{[key: string]: string}}
 */
function fetchProdReleaseDateByKeys_(keys, cfg) {
  const out = {};
  const unique = Array.from(new Set((keys || []).filter(Boolean)));
  if (!unique.length || !cfg || !cfg.jiraBaseUrl) return out;
  const fieldId = normalizeJiraFieldIdForApi_(FIELD_MAP && FIELD_MAP.prodReleaseDate);
  if (!fieldId) return out;

  const url = `${cfg.jiraBaseUrl}/rest/api/3/search/jql`;
  const fields = [fieldId];
  const CHUNK = 50;
  for (let i = 0; i < unique.length; i += CHUNK) {
    const chunk = unique.slice(i, i + CHUNK);
    const jql = `key in (${chunk.map(k => `"${k}"`).join(", ")})`;
    let nextPageToken = null;
    do {
      const payload = { jql, maxResults: 100, fields };
      if (nextPageToken) payload.nextPageToken = nextPageToken;
      const resp = UrlFetchApp.fetch(url, {
        method: "post",
        headers: getAuthHeader(cfg),
        contentType: "application/json",
        muteHttpExceptions: true,
        payload: JSON.stringify(payload)
      });
      if (resp.getResponseCode() >= 400) {
        Logger.log(
          "fetchProdReleaseDateByKeys_: chunk " + i + " HTTP " +
            resp.getResponseCode() + " " + resp.getContentText().slice(0, 200)
        );
        break;
      }
      const result = JSON.parse(resp.getContentText());
      (result.issues || []).forEach(issue => {
        const f = issue.fields || {};
        out[issue.key] = jiraFieldAsPlainString_(f[fieldId]) || "";
      });
      nextPageToken = result.nextPageToken || null;
    } while (nextPageToken);
  }
  return out;
}

/**
 * Captures formula text (if any) or literal cell value for each cell in the
 * named columns across rows headerRow+1 .. sheet.getLastRow(). Designed to
 * round-trip through restoreProtectedColumns_ so user-managed VLOOKUP
 * formulas / manual edits survive a clear+write cycle (Refresh or Create).
 *
 * Each entry of the returned map is the formula string (starts with `=`) or
 * the literal value (string / number / Date). setValues() auto-detects the
 * difference on restore.
 *
 * @param {Sheet} sheet
 * @param {number} headerRow
 * @param {string[]} columnNames Header names to protect (matched in `map`).
 * @param {Object<string,number>} map header→0-based-column-index map.
 * @returns {{[colIdx: number]: any[]}}
 */
function snapshotProtectedColumns_(sheet, headerRow, columnNames, map) {
  const snapshot = {};
  if (!sheet || !headerRow || !Array.isArray(columnNames) || !columnNames.length) {
    return snapshot;
  }
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - headerRow;
  if (numRows <= 0) return snapshot;
  columnNames.forEach(name => {
    const ci = map && map[name];
    if (ci === undefined || ci === null) {
      Logger.log("snapshotProtectedColumns_: column not found, skipping: " + name);
      return;
    }
    const range = sheet.getRange(headerRow + 1, ci + 1, numRows, 1);
    const formulas = range.getFormulas();
    const values = range.getValues();
    snapshot[ci] = formulas.map((row, i) => row[0] || values[i][0]);
  });
  return snapshot;
}

/**
 * Writes a snapshot taken by snapshotProtectedColumns_ back into the same
 * column positions, capped at the sheet's current last row so we don't
 * extend the data area below the new write. setValues() auto-detects
 * formula strings (entries starting with `=`).
 *
 * @param {Sheet} sheet
 * @param {number} headerRow
 * @param {{[colIdx: number]: any[]}} snapshot Result of snapshotProtectedColumns_.
 */
function restoreProtectedColumns_(sheet, headerRow, snapshot) {
  if (!sheet || !snapshot) return;
  const colKeys = Object.keys(snapshot);
  if (!colKeys.length) return;
  const lastRow = sheet.getLastRow();
  const maxRows = lastRow - headerRow;
  if (maxRows <= 0) return;
  colKeys.forEach(ciStr => {
    const ci = Number(ciStr);
    const cells = snapshot[ci];
    if (!Array.isArray(cells) || !cells.length) return;
    const numRows = Math.min(cells.length, maxRows);
    if (numRows <= 0) return;
    const valuesToWrite = cells.slice(0, numRows).map(c => [c]);
    sheet.getRange(headerRow + 1, ci + 1, numRows, 1).setValues(valuesToWrite);
  });
}

/**
 * Bulk-fetches { status, resolution } for the given Jira issue keys.
 * The status from issuelinks is usually fresh enough, but resolution is not
 * present on the issuelink stub so we have to call /search ourselves.
 *
 * @param {string[]} keys Jira issue keys (deduped recommended).
 * @param {object} cfg getConfig() result (uses jiraBaseUrl + auth).
 * @returns {{[key: string]: {status: string, resolution: string}}}
 */
function fetchBlockerStatusAndResolutionByKeys_(keys, cfg) {
  const out = {};
  const unique = Array.from(new Set((keys || []).filter(Boolean)));
  if (!unique.length || !cfg || !cfg.jiraBaseUrl) return out;

  const url = `${cfg.jiraBaseUrl}/rest/api/3/search/jql`;
  const fields = ["status", "resolution"];
  // Keep each chunk's JQL well under URL/JQL length limits.
  const CHUNK = 50;
  for (let i = 0; i < unique.length; i += CHUNK) {
    const chunk = unique.slice(i, i + CHUNK);
    const jql = `key in (${chunk.map(k => `"${k}"`).join(", ")})`;
    let nextPageToken = null;
    do {
      const payload = { jql, maxResults: 100, fields };
      if (nextPageToken) payload.nextPageToken = nextPageToken;
      const resp = UrlFetchApp.fetch(url, {
        method: "post",
        headers: getAuthHeader(cfg),
        contentType: "application/json",
        muteHttpExceptions: true,
        payload: JSON.stringify(payload)
      });
      if (resp.getResponseCode() >= 400) {
        Logger.log(
          "fetchBlockerStatusAndResolutionByKeys_: chunk " + i + " HTTP " +
            resp.getResponseCode() + " " + resp.getContentText().slice(0, 200)
        );
        break;
      }
      const result = JSON.parse(resp.getContentText());
      (result.issues || []).forEach(issue => {
        const f = issue.fields || {};
        out[issue.key] = {
          status: f.status?.name || "",
          resolution: f.resolution?.name || ""
        };
      });
      nextPageToken = result.nextPageToken || null;
    } while (nextPageToken);
  }
  return out;
}

/**
 * Mutates `issues` in place: fills each issue.blockerInfos[*].resolution
 * (and refreshes status) using a single bulk fetch keyed off the unique set
 * of blocker keys across all issues.
 */
function enrichBlockerInfosWithResolutions_(issues, cfg) {
  if (!Array.isArray(issues) || !issues.length) return;
  const allKeys = [];
  issues.forEach(i => {
    (i.blockerInfos || []).forEach(b => {
      if (b && b.key) allKeys.push(b.key);
    });
  });
  if (!allKeys.length) return;
  const lookup = fetchBlockerStatusAndResolutionByKeys_(allKeys, cfg);
  issues.forEach(i => {
    (i.blockerInfos || []).forEach(b => {
      const hit = lookup[b.key];
      if (!hit) return;
      if (hit.status) b.status = hit.status;
      if (hit.resolution) b.resolution = hit.resolution;
    });
  });
}

/* =================== Helpers =================== */
const MAX_CELL_CHARS = 49000;

/**
 * Caps a free-text value for the ChangeLog. The ChangeLog is for human review
 * of "what moved on this Refresh"; long Jira description bodies are noise and
 * also bloat per-row write cost. Anything over `n` chars is truncated with a
 * trailing ellipsis. null/undefined become "".
 */
function truncateForChangelog_(s, n) {
  if (s == null) return "";
  const t = String(s);
  const cap = n || 200;
  return t.length > cap ? t.slice(0, cap - 1) + "\u2026" : t;
}

/**
 * Normalizes a cell value for comparison.
 * Handles Date objects, null/undefined, and trims whitespace.
 */
function normalizeForComparison(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    // Format date to YYYY-MM-DD to match Jira's format
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return val.toString().trim();
}

function extractPlainDescription(desc) {
  if (!desc?.content) return "";
  const walk = nodes =>
    nodes.map(n =>
      n.text || (n.content ? walk(n.content) : "")
    ).join("");
  return walk(desc.content);
}

/**
 * Retries a Spreadsheets-service call when Apps Script throws a transient
 * "Service timed out: Spreadsheets" error. Flushes the pending op queue
 * between attempts so the service can drain.
 */
function withSpreadsheetRetry_(fn, label) {
  const MAX_ATTEMPTS = 5;
  let lastErr;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      const msg = (e && e.message) ? e.message : String(e);
      // Match every flavor Apps Script reports for transient Sheets errors:
      //   - "Service Spreadsheets timed out while accessing document …"
      //     (literal "Service timed out" substring never matched — there's a
      //     word in between, which is why retries here used to silently no-op)
      //   - "Document <id> is missing (perhaps it was deleted, or you don't
      //     have read access?)" — despite the wording this is a transient
      //     backend error from Apps Script, NOT a real permission/delete
      //     issue; retrying after a flush + backoff almost always succeeds.
      //   - "Service invoked too many times", "Internal error",
      //     "try again later", "backend error" — other transient flavors.
      const transient =
        /timed out/i.test(msg) ||
        /Service invoked too many times/i.test(msg) ||
        /Internal error/i.test(msg) ||
        /try again later/i.test(msg) ||
        /backend error/i.test(msg) ||
        /is missing \(perhaps it was deleted/i.test(msg) ||
        /document .* is missing/i.test(msg);
      if (!transient || attempt === MAX_ATTEMPTS) throw e;
      Logger.log(
        "withSpreadsheetRetry_(" + (label || "op") + "): attempt " + attempt +
          " failed: " + msg + " — flushing and retrying"
      );
      try { SpreadsheetApp.flush(); } catch (_) {}
      // Slightly more aggressive backoff for the "Document … is missing"
      // flavor — it usually needs ~2-5s to clear, longer than a typical
      // setValues timeout. 1.5s × attempt^1.4 gives 1.5s, ~4s, ~7.5s, ~12s.
      const docMissing = /is missing/i.test(msg);
      const baseMs = docMissing ? 1500 : 1000;
      Utilities.sleep(Math.round(baseMs * Math.pow(attempt, docMissing ? 1.4 : 1)));
    }
  }
  throw lastErr;
}

function getHeaderIndexMapSafe(sheet, headerRow = 1) {
  return withSpreadsheetRetry_(() => {
    const lastCol = Math.max(sheet.getLastColumn(), 0);
    if (lastCol === 0) return { headers: [], map: {} };

    const rawHeaders = sheet
      .getRange(headerRow, 1, 1, lastCol)
      .getValues()[0];

    const headers = rawHeaders.map(h =>
      h === null || h === undefined ? "" : h.toString().trim()
    );

    const map = {};
    headers.forEach((h, i) => {
      if (h) map[h] = i;
    });

    return { headers, map };
  }, "getHeaderIndexMapSafe");
}

/**
 * Same spreadsheet tab as the sheet formula: 'Release & Approval Matrix' (case-insensitive fallback).
 */
function getReleaseApprovalMatrixSheet_(ss) {
  const preferred = "Release & Approval Matrix";
  let sh = ss.getSheetByName(preferred);
  if (sh) return sh;
  const want = preferred.toLowerCase();
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const n = sheets[i].getName().trim();
    if (n.toLowerCase() === want) return sheets[i];
  }
  Logger.log(
    "getReleaseApprovalMatrixSheet_: no tab \"" +
      preferred +
      "\". Tabs: " +
      sheets.map(s => "\"" + s.getName() + "\"").join(", ")
  );
  return null;
}

/**
 * Column AO / "Tech Lead / Sr Tech Lead" from tab "Release & Approval Matrix" (no constants.gs).
 * Mirrors:
 *   IF(F2="","", IFERROR( LET( r, MATCH(F2, 'Release & Approval Matrix'!$C$2:$C, 0), ... IF(g="", f, f&"/"&g)), ""))
 * First match in column C wins (same as MATCH row order).
 */
/**
 * Reads the Release & Approval Matrix sheet once and returns a
 * `{ scrumTeam: "lead/srlead" }` lookup map.
 *
 * Used by `buildRowFromIssue` so the Tech Lead column gets populated as
 * part of the main row matrix (no separate sheet read + sheet write
 * round-trip needed). Returns `null` if the matrix sheet is missing or
 * empty.
 */
function buildTechLeadLookupFromMatrix_(ss) {
  const MATRIX_FIRST_ROW = 2;
  const MATRIX_COL_KEY = 3; // C
  const MATRIX_COL_LEAD = 6; // F
  const MATRIX_COL_SR = 7; // G

  const matrix = getReleaseApprovalMatrixSheet_(ss);
  if (!matrix) return null;

  const mLast = matrix.getLastRow();
  if (mLast < MATRIX_FIRST_ROW) return {};

  const matrixRows = mLast - MATRIX_FIRST_ROW + 1;
  const matrixSpan = MATRIX_COL_SR - MATRIX_COL_KEY + 1; // 5 (C..G)
  const matrixVals = matrix.getRange(MATRIX_FIRST_ROW, MATRIX_COL_KEY, matrixRows, matrixSpan).getValues();
  const C_OFFSET = 0;
  const F_OFFSET = MATRIX_COL_LEAD - MATRIX_COL_KEY; // 3
  const G_OFFSET = MATRIX_COL_SR - MATRIX_COL_KEY;   // 4

  const lookup = {};
  for (let i = 0; i < matrixVals.length; i++) {
    const row = matrixVals[i];
    const k = (row[C_OFFSET] != null ? row[C_OFFSET] : "").toString().trim();
    if (!k || k in lookup) continue;
    const f = (row[F_OFFSET] != null ? row[F_OFFSET] : "").toString().trim();
    const g = (row[G_OFFSET] != null ? row[G_OFFSET] : "").toString().trim();
    lookup[k] = g === "" ? f : f + "/" + g;
  }
  return lookup;
}

function populateTechLeadsFromMatrix(sheet, headerRow, opts) {
  const RELEASE_COL_LOOKUP_FALLBACK = 6; // 1-based fallback if Scrum Team header missing
  const RELEASE_HEADER_OUT = "Tech Lead / Sr Tech Lead";

  // Reuse a precomputed header map when the caller already has one — this
  // avoids re-reading the Release sheet right after a heavy write batch,
  // which is the situation that produced "Service timed out: Spreadsheets".
  const map = (opts && opts.map) ? opts.map : getHeaderIndexMapSafe(sheet, headerRow).map;
  const outIdx = map[RELEASE_HEADER_OUT];
  if (outIdx === undefined) {
    Logger.log('populateTechLeadsFromMatrix: no "' + RELEASE_HEADER_OUT + '" column on Release sheet');
    return;
  }

  const lookup = buildTechLeadLookupFromMatrix_(sheet.getParent());
  if (!lookup) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return;

  // Same as protectManualColumns: getRange(row, col, numRows, numColumns) — NOT end row/column.
  const numDataRows = lastRow - headerRow;
  const lookupCol1 =
    map["Scrum Team"] !== undefined ? map["Scrum Team"] + 1 : RELEASE_COL_LOOKUP_FALLBACK;

  const releaseKeys = sheet.getRange(headerRow + 1, lookupCol1, numDataRows, 1).getValues();
  const out = releaseKeys.map(row => {
    const k = (row[0] != null ? row[0] : "").toString().trim();
    if (k === "") return [""];
    return [lookup[k] !== undefined ? lookup[k] : ""];
  });

  if (out.length !== numDataRows) {
    Logger.log(
      "populateTechLeadsFromMatrix: row count mismatch out=" + out.length + " vs data=" + numDataRows
    );
    return;
  }

  sheet.getRange(headerRow + 1, outIdx + 1, numDataRows, 1).setValues(out);
}

function buildJiraBrowseUrl(cfg, issueKey) {
  if (!cfg?.jiraBaseUrl || !issueKey) return "";
  return `${String(cfg.jiraBaseUrl).replace(/\/$/, "")}/browse/${issueKey}`;
}

/**
 * Builds a HYPERLINK formula string that renders as a clickable Jira link.
 * Falls back to the bare issue key when no Jira base URL is configured (so
 * the cell still shows something useful instead of an empty formula).
 *
 * Why a formula instead of setRichTextValues:
 *   getValues() on a HYPERLINK formula returns the *displayed* text (the
 *   key), so existing diff logic that reads `row[jiraItemCol]` keeps
 *   working unchanged. Writing the formula in the main setValues batch
 *   eliminates a separate setRichTextValues round-trip per refresh.
 */
function makeJiraItemHyperlinkFormula_(issueKey, cfg) {
  const key = issueKey == null ? "" : String(issueKey);
  if (!key) return "";
  const url = buildJiraBrowseUrl(cfg, key);
  if (!url) return key;
  // Escape any double-quotes in key (Jira keys never contain them, but be safe).
  const safeKey = key.replace(/"/g, '""');
  const safeUrl = url.replace(/"/g, '""');
  return `=HYPERLINK("${safeUrl}","${safeKey}")`;
}

function setJiraItemLinks(sheet, startRow, issueKeys, jiraItemColIndex, cfg) {
  if (jiraItemColIndex === undefined || !issueKeys.length) return;

  const richTextValues = issueKeys.map(issueKey => {
    const text = issueKey || "";
    const builder = SpreadsheetApp.newRichTextValue().setText(text);
    const url = buildJiraBrowseUrl(cfg, issueKey);

    if (url) {
      builder.setLinkUrl(url);
    }

    return [builder.build()];
  });

  sheet
    .getRange(startRow, jiraItemColIndex + 1, issueKeys.length, 1)
    .setRichTextValues(richTextValues);
}

function fetchSNOWRecordByCR(crNumber, cfg) {
  if (!crNumber) return null;

  const authHeader = "Basic " + Utilities.base64Encode(cfg.snUser + ":" + cfg.snPass);
  const url =
    `${cfg.snBaseUrl}/api/now/table/change_request` +
    `?sysparm_query=number=${encodeURIComponent(crNumber)}` +
    `&sysparm_limit=1` +
    `&sysparm_display_value=all`;

  const resp = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { Authorization: authHeader, Accept: "application/json" },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) return null;

  const json = JSON.parse(resp.getContentText());
  return (json.result && json.result[0]) || null;
}

function getSNOWAuthHeader(cfg) {
  return "Basic " + Utilities.base64Encode(cfg.snUser + ":" + cfg.snPass);
}

function getSNOWDisplayValue(field) {
  if (!field) return "";
  if (typeof field === "string") return field;
  return field.display_value || field.value || "";
}

/** First non-empty SNOW display string for a list of API field names on change_request. */
function snowDisplayFromRecord_(record, fieldNames) {
  if (!record || !fieldNames || !fieldNames.length) return "";
  for (let i = 0; i < fieldNames.length; i++) {
    const fld = record[fieldNames[i]];
    if (!fld) continue;
    const v = getSNOWDisplayValue(fld);
    if (v && String(v).trim()) return String(v).trim();
  }
  return "";
}

/** Jira fields array for search/jql (deduped); includes priority + every valid FIELD_MAP API id. */
function jiraFieldsForFetch_() {
  const out = [
    "key",
    "summary",
    "description",
    "issuetype",
    "comment",
    "status",
    "resolution",
    "assignee",
    "issuelinks",
    "priority",
    // Needed by isExcludedCapTicket_ — drops CAP tickets whose Components /
    // CMDB CI mention restricted tokens (see CAP_EXCLUDED_TOKENS in constants).
    "project",
    "components"
  ];
  const seen = {};
  out.forEach(id => {
    seen[id] = true;
  });
  if (typeof FIELD_MAP === "object" && FIELD_MAP) {
    Object.keys(FIELD_MAP).forEach(mapKey => {
      const id = normalizeJiraFieldIdForApi_(FIELD_MAP[mapKey]);
      if (!id) return;
      if (seen[id]) return;
      seen[id] = true;
      out.push(id);
    });
  }
  if (typeof CAP_CMDB_CI_FIELD_ID === "string" && CAP_CMDB_CI_FIELD_ID && !seen[CAP_CMDB_CI_FIELD_ID]) {
    seen[CAP_CMDB_CI_FIELD_ID] = true;
    out.push(CAP_CMDB_CI_FIELD_ID);
  }
  return out;
}

function chunkValues(values, chunkSize) {
  const chunks = [];
  for (let i = 0; i < values.length; i += chunkSize) {
    chunks.push(values.slice(i, i + chunkSize));
  }
  return chunks;
}

/**
 * Formats a single CTASK entry for the "CTASK List" sheet cell.
 *
 * Input: { number: "CTASK0001234", assignmentGroup: "Group A" }
 * Output: "CTASK0001234 (Group A)" if assignmentGroup is non-empty, else "CTASK0001234".
 * Multiple CTASKs are joined by ", " by the caller.
 */
function formatCtaskWithGroup_(task) {
  if (!task) return "";
  const num = (task.number || "").toString().trim();
  if (!num) return "";
  const grp = (task.assignmentGroup || "").toString().trim();
  return grp ? `${num} (${grp})` : num;
}

function fetchSNOWDataByCRNumbers(crNumbers, cfg, includeCTaskList) {
  if (!cfg?.snBaseUrl) return {};

  const uniqueCRs = [...new Set((crNumbers || []).filter(Boolean))];
  if (!uniqueCRs.length) return {};

  const authHeader = getSNOWAuthHeader(cfg);
  const chunkSize = 25;
  const baseSnowFields = [
    "number",
    "start_date",
    "u_change_owner",
    "assigned_to",
    "change_plan",
    "backout_plan",
    "u_validation_status",
    "u_validation_owner",
    "u_validation_plan",
    "u_notes_comments",
    "u_ctask_check",
    "end_date",
    "cmdb_ci",
    "type",
    "opened_at",
    "u_affected_groups",
    "u_affected_locations",
    "assignment_group",
    "u_impact_description",
    "impact_description",
    "description"
  ];
  const extraSnow =
    typeof SNOW_CHANGE_REQUEST_EXTRA_FIELDS !== "undefined" && SNOW_CHANGE_REQUEST_EXTRA_FIELDS
      ? SNOW_CHANGE_REQUEST_EXTRA_FIELDS
      : [];
  const seenSnow = {};
  const snowFieldList = [];
  baseSnowFields.concat(extraSnow).forEach(fn => {
    if (!fn || seenSnow[fn]) return;
    seenSnow[fn] = true;
    snowFieldList.push(fn);
  });
  const fields = snowFieldList.join(",");
  const snowDataByCR = {};

  const changeRequests = chunkValues(uniqueCRs, chunkSize).map(chunk => ({
    url:
      `${cfg.snBaseUrl}/api/now/table/change_request` +
      `?sysparm_query=${encodeURIComponent(`numberIN${chunk.join(",")}`)}` +
      `&sysparm_display_value=all` +
      `&sysparm_limit=${chunk.length}` +
      `&sysparm_fields=${fields}`,
    method: "get",
    headers: { Authorization: authHeader, Accept: "application/json" },
    muteHttpExceptions: true
  }));

  if (changeRequests.length) {
    const responses = UrlFetchApp.fetchAll(changeRequests);
    responses.forEach(resp => {
      if (resp.getResponseCode() !== 200) {
        Logger.log(`SNOW change_request fetch failed: ${resp.getResponseCode()} ${resp.getContentText()}`);
        return;
      }

      const json = JSON.parse(resp.getContentText());
      const records = Array.isArray(json.result) ? json.result : [];
      records.forEach(record => {
        const crNumber = getSNOWDisplayValue(record.number);
        if (!crNumber) return;
        snowDataByCR[crNumber] = {
          record,
          ctaskList: "No CTASKs found"
        };
      });
    });
  }

  if (!includeCTaskList) {
    return snowDataByCR;
  }

  const ctaskRequests = chunkValues(uniqueCRs, chunkSize).map(chunk => ({
    url:
      `${cfg.snBaseUrl}/api/now/table/change_task` +
      `?sysparm_query=${encodeURIComponent(`change_request.numberIN${chunk.join(",")}`)}` +
      `&sysparm_fields=change_request,number,assignment_group` +
      `&sysparm_display_value=all` +
      `&sysparm_limit=1000`,
    method: "get",
    headers: { Authorization: authHeader, Accept: "application/json" },
    muteHttpExceptions: true
  }));

  if (!ctaskRequests.length) {
    return snowDataByCR;
  }

  const ctaskNumbersByCR = {};
  const ctaskResponses = UrlFetchApp.fetchAll(ctaskRequests);
  ctaskResponses.forEach(resp => {
    if (resp.getResponseCode() !== 200) {
      Logger.log(`SNOW change_task fetch failed: ${resp.getResponseCode()} ${resp.getContentText()}`);
      return;
    }

    const json = JSON.parse(resp.getContentText());
    const tasks = Array.isArray(json.result) ? json.result : [];
    tasks.forEach(task => {
      const crNumber = getSNOWDisplayValue(task.change_request);
      const taskNumber = getSNOWDisplayValue(task.number);
      if (!crNumber || !taskNumber) return;
      const assignmentGroup = getSNOWDisplayValue(task.assignment_group);
      if (!ctaskNumbersByCR[crNumber]) {
        ctaskNumbersByCR[crNumber] = [];
      }
      ctaskNumbersByCR[crNumber].push({ number: taskNumber, assignmentGroup });
    });
  });

  uniqueCRs.forEach(crNumber => {
    if (!snowDataByCR[crNumber]) {
      snowDataByCR[crNumber] = { record: {}, ctaskList: "No CTASKs found" };
    }

    const tasks = ctaskNumbersByCR[crNumber] || [];
    snowDataByCR[crNumber].ctaskList = tasks.length
      ? tasks.map(formatCtaskWithGroup_).join(", ")
      : "No CTASKs found";
  });

  return snowDataByCR;
}

/**
 * Single UrlFetchApp.fetchAll batch covering all SNOW change_request /
 * change_task chunks AND all Jira blocker-resolution chunks in one shot.
 *
 * Why: change_request, change_task, and blocker-resolution lookups are all
 * independent of each other. Issuing them as one parallel batch makes the
 * effective wall time roughly max(SNOW_CR, SNOW_CT, BLOCKERS) instead of
 * SNOW_CR + SNOW_CT + BLOCKERS.
 *
 * @param {string[]} crNumbers     SNOW CR numbers (will be deduped).
 * @param {string[]} blockerKeys   Jira issue keys for blocker resolution
 *                                 lookup (will be deduped).
 * @param {object}   cfg
 * @param {boolean}  includeCTaskList  If true, also batches change_task
 *                                     queries and fills `ctaskList`.
 * @returns {{ snowDataByCR: Object, blockerLookup: Object }}
 */
function fetchSnowAndBlockersInParallel_(crNumbers, blockerKeys, cfg, includeCTaskList) {
  const snowDataByCR = {};
  const blockerLookup = {};

  const uniqueCRs = [...new Set((crNumbers || []).filter(Boolean))];
  const uniqueBlockerKeys = Array.from(new Set((blockerKeys || []).filter(Boolean)));

  /** @type {Array<{tag: string, request: object, chunkIndex?: number}>} */
  const specs = [];

  if (uniqueCRs.length && cfg && cfg.snBaseUrl) {
    const authHeader = getSNOWAuthHeader(cfg);
    const baseSnowFields = [
      "number", "start_date", "u_change_owner", "assigned_to",
      "change_plan", "backout_plan", "u_validation_status",
      "u_validation_owner", "u_validation_plan", "u_notes_comments",
      "u_ctask_check", "end_date", "cmdb_ci", "type", "opened_at",
      "u_affected_groups", "u_affected_locations", "assignment_group",
      "u_impact_description", "impact_description", "description"
    ];
    const extraSnow =
      typeof SNOW_CHANGE_REQUEST_EXTRA_FIELDS !== "undefined" && SNOW_CHANGE_REQUEST_EXTRA_FIELDS
        ? SNOW_CHANGE_REQUEST_EXTRA_FIELDS
        : [];
    const seenSnow = {};
    const snowFieldList = [];
    baseSnowFields.concat(extraSnow).forEach(fn => {
      if (!fn || seenSnow[fn]) return;
      seenSnow[fn] = true;
      snowFieldList.push(fn);
    });
    const fields = snowFieldList.join(",");
    const chunkSize = 25;

    chunkValues(uniqueCRs, chunkSize).forEach(chunk => {
      specs.push({
        tag: "snow_cr",
        request: {
          url:
            `${cfg.snBaseUrl}/api/now/table/change_request` +
            `?sysparm_query=${encodeURIComponent(`numberIN${chunk.join(",")}`)}` +
            `&sysparm_display_value=all` +
            `&sysparm_limit=${chunk.length}` +
            `&sysparm_fields=${fields}`,
          method: "get",
          headers: { Authorization: authHeader, Accept: "application/json" },
          muteHttpExceptions: true
        }
      });
    });

    if (includeCTaskList) {
      chunkValues(uniqueCRs, chunkSize).forEach(chunk => {
        specs.push({
          tag: "snow_ct",
          request: {
            url:
              `${cfg.snBaseUrl}/api/now/table/change_task` +
              `?sysparm_query=${encodeURIComponent(`change_request.numberIN${chunk.join(",")}`)}` +
              `&sysparm_fields=change_request,number,assignment_group` +
              `&sysparm_display_value=all` +
              `&sysparm_limit=1000`,
            method: "get",
            headers: { Authorization: authHeader, Accept: "application/json" },
            muteHttpExceptions: true
          }
        });
      });
    }
  }

  if (uniqueBlockerKeys.length && cfg && cfg.jiraBaseUrl) {
    const url = `${cfg.jiraBaseUrl}/rest/api/3/search/jql`;
    const fields = ["status", "resolution"];
    const CHUNK = 50;
    for (let i = 0; i < uniqueBlockerKeys.length; i += CHUNK) {
      const chunk = uniqueBlockerKeys.slice(i, i + CHUNK);
      const jql = `key in (${chunk.map(k => `"${k}"`).join(", ")})`;
      const payload = { jql, maxResults: 100, fields };
      specs.push({
        tag: "blocker",
        chunkIndex: i,
        request: {
          url,
          method: "post",
          headers: getAuthHeader(cfg),
          contentType: "application/json",
          muteHttpExceptions: true,
          payload: JSON.stringify(payload)
        }
      });
    }
  }

  if (!specs.length) return { snowDataByCR, blockerLookup };

  const responses = UrlFetchApp.fetchAll(specs.map(s => s.request));
  const ctaskNumbersByCR = {};

  responses.forEach((resp, i) => {
    const spec = specs[i];
    const code = resp.getResponseCode();

    if (spec.tag === "snow_cr") {
      if (code !== 200) {
        Logger.log(`SNOW change_request fetch failed: ${code} ${resp.getContentText()}`);
        return;
      }
      const json = JSON.parse(resp.getContentText());
      const records = Array.isArray(json.result) ? json.result : [];
      records.forEach(record => {
        const crNumber = getSNOWDisplayValue(record.number);
        if (!crNumber) return;
        snowDataByCR[crNumber] = {
          record,
          ctaskList: "No CTASKs found"
        };
      });
    } else if (spec.tag === "snow_ct") {
      if (code !== 200) {
        Logger.log(`SNOW change_task fetch failed: ${code} ${resp.getContentText()}`);
        return;
      }
      const json = JSON.parse(resp.getContentText());
      const tasks = Array.isArray(json.result) ? json.result : [];
      tasks.forEach(task => {
        const crNumber = getSNOWDisplayValue(task.change_request);
        const taskNumber = getSNOWDisplayValue(task.number);
        if (!crNumber || !taskNumber) return;
        const assignmentGroup = getSNOWDisplayValue(task.assignment_group);
        if (!ctaskNumbersByCR[crNumber]) ctaskNumbersByCR[crNumber] = [];
        ctaskNumbersByCR[crNumber].push({ number: taskNumber, assignmentGroup });
      });
    } else if (spec.tag === "blocker") {
      if (code >= 400) {
        Logger.log(
          "blocker resolution chunk " + spec.chunkIndex + " HTTP " + code + " " +
            resp.getContentText().slice(0, 200)
        );
        return;
      }
      const json = JSON.parse(resp.getContentText());
      (json.issues || []).forEach(issue => {
        const f = issue.fields || {};
        blockerLookup[issue.key] = {
          status: f.status?.name || "",
          resolution: f.resolution?.name || ""
        };
      });
    }
  });

  if (includeCTaskList) {
    uniqueCRs.forEach(crNumber => {
      if (!snowDataByCR[crNumber]) {
        snowDataByCR[crNumber] = { record: {}, ctaskList: "No CTASKs found" };
      }
      const tasks = ctaskNumbersByCR[crNumber] || [];
      snowDataByCR[crNumber].ctaskList = tasks.length
        ? tasks.map(formatCtaskWithGroup_).join(", ")
        : "No CTASKs found";
    });
  }

  return { snowDataByCR, blockerLookup };
}

/**
 * Applies a pre-fetched blocker resolution lookup onto issues. Mutates
 * issues[*].blockerInfos[*].{status, resolution} in place.
 */
function applyBlockerResolutionLookup_(issues, lookup) {
  if (!Array.isArray(issues) || !issues.length || !lookup) return;
  issues.forEach(i => {
    (i.blockerInfos || []).forEach(b => {
      const hit = lookup[b.key];
      if (!hit) return;
      if (hit.status) b.status = hit.status;
      if (hit.resolution) b.resolution = hit.resolution;
    });
  });
}

/**
 * Collects all unique blocker Jira keys across an issues array.
 */
function collectAllBlockerKeysFromIssues_(issues) {
  const out = [];
  if (!Array.isArray(issues)) return out;
  issues.forEach(i => {
    (i.blockerInfos || []).forEach(b => {
      if (b && b.key) out.push(b.key);
    });
  });
  return out;
}

/* =================== Column Protection =================== */

/**
 * Clears existing column protections created by this script
 */
function clearColumnProtections(sheet) {
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(p => {
    if (p.getDescription().startsWith("Protected Column:")) {
      p.remove();
    }
  });
}

/**
 * Protects the specified "green manual columns" so only designated editors can modify them.
 * The script can still write to these columns during refresh.
 */
function protectManualColumns(sheet, headerRow) {
  const { map } = getHeaderIndexMapSafe(sheet, headerRow);
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= headerRow) return;
  
  // Clear existing protections first to avoid duplicates
  clearColumnProtections(sheet);
  
  // Get the script owner (needed to ensure script can still edit)
  const me = Session.getEffectiveUser();
  
  PROTECTED_COLUMNS.forEach(colName => {
    const colIdx = map[colName];
    if (colIdx === undefined) {
      Logger.log(`Column "${colName}" not found in headers, skipping protection.`);
      return;
    }
    
    const colIndex = colIdx + 1; // Convert to 1-based
    const range = sheet.getRange(headerRow + 1, colIndex, lastRow - headerRow, 1);
    
    try {
      const protection = range.protect().setDescription(`Protected Column: ${colName}`);
      
      // Add the designated editors
      PROTECTED_COLUMN_EDITORS.forEach(email => {
        try {
          protection.addEditor(email);
        } catch (e) {
          Logger.log(`Could not add editor ${email}: ${e.message}`);
        }
      });
      
      // Add the script owner so the script can still write during refresh
      protection.addEditor(me);
      
      // Remove all other editors (keep only the ones we explicitly added)
      const editors = protection.getEditors();
      const allowedEditors = [me.getEmail().toLowerCase(), ...PROTECTED_COLUMN_EDITORS.map(e => e.toLowerCase())];
      
      editors.forEach(editor => {
        if (!allowedEditors.includes(editor.getEmail().toLowerCase())) {
          protection.removeEditor(editor);
        }
      });
      
      // Disable domain-wide edit access if enabled
      if (protection.canDomainEdit()) {
        protection.setDomainEdit(false);
      }
      
      Logger.log(`Protected column: ${colName}`);
    } catch (e) {
      Logger.log(`Error protecting column ${colName}: ${e.message}`);
    }
  });
}

/**
 * Builds a row array from a normalized issue.
 *
 * @param {object} issue
 * @param {string[]} headers
 * @param {Object<string, number>} map header-name -> 0-based column index
 * @param {object} cfg
 * @param {Object} [snowDataByCR]
 * @param {Object<string, string>} [techLeadLookup]  Optional scrumTeam -> "lead/srlead"
 *        lookup. When provided, the Tech Lead / Sr Tech Lead column is populated
 *        directly so the caller doesn't need to run populateTechLeadsFromMatrix
 *        as a separate post-write pass.
 */
function buildRowFromIssue(issue, headers, map, cfg, snowDataByCR, techLeadLookup) {
  const r = Array(headers.length).fill("");
  const s = (h, v) => map[h] !== undefined && (r[map[h]] = v ?? "");
  const crForSnow = issue.displayCr != null ? issue.displayCr : "";
  const snowData = (snowDataByCR && crForSnow && snowDataByCR[crForSnow]) || null;
  const record = snowData?.record || {};

  // --- Jira fields ---
  s("Scrum Team", issue.scrumTeam);
  s("Workstream", issue.workstream);
  s("CR #", crForSnow);

  s("Approvals", issue.approvals);
  s("JIRA Item", makeJiraItemHyperlinkFormula_(issue.key, cfg));
  s("Channel", issue.channel);
  s("Issue Type", issue.issueType);
  s("Jira Prod Release Date", issue.prodReleaseDate);
  s("Jira Status", issue.status);
  s("Jira Resolution", issue.resolution);
  s("JIRA Summary", issue.summary);
  s("JIRA Description", issue.description);
  s("Related Issues", (issue.relatedIssues || []).join(", "));

  // Blocker Status shows the status of each blocking ticket plus its Jira
  // Resolution when set (e.g., "DLL-123: Done (Fixed)" or "DLL-123: In Progress").
  // Resolution is populated by enrichBlockerInfosWithResolutions_ — falls back
  // to status-only if it wasn't run or the lookup didn't return a match.
  const blockerStatusLines = (issue.blockerInfos || []).map(b => {
    const tail = b.resolution ? ` (${b.resolution})` : "";
    return `${b.key}: ${b.status || "Unknown"}${tail}`;
  });
  s("Blocker Status", blockerStatusLines.join("\n"));
  
  s("Jira Assignee", issue.assignee);
  s("JIRA Assignee", issue.assignee);
  s("Acceptance Criteria", issue.acceptanceCriteria);
  s("Last Updated JIRA comment", issue.lastComment);
  s("QA Artifacts Link", issue.qaArtifactsLink);
  s("UAT Artifacts Link", issue.uatArtifactsLink);
  s("Severity", issue.severity);

  // --- Tech Lead lookup (from Release & Approval Matrix) ---
  // Bake-in saves a separate post-write read+write pass over the Tech Lead
  // column. Caller passes a precomputed scrumTeam -> "lead/srlead" map.
  if (techLeadLookup) {
    const teamKey = (issue.scrumTeam || "").toString().trim();
    if (teamKey && techLeadLookup[teamKey]) {
      s("Tech Lead / Sr Tech Lead", techLeadLookup[teamKey]);
    }
  }

  // --- ServiceNow fields ---
  if (crForSnow && cfg.snBaseUrl) {
    s("Start Time", record.start_date?.display_value || "");
    s("Change Assignee", record.u_change_owner?.display_value || record.assigned_to?.display_value || "");
    s("Change Plan", record.change_plan?.display_value || "");
    s("Rollback Plan", record.backout_plan?.display_value || "");
    s("Validation Status", record.u_validation_status?.display_value || "");
    s("Validation Owner", record.u_validation_owner?.display_value || "");
    s("Validation Plan", record.u_validation_plan?.display_value || "");
    s("SNOW Work Notes", record.u_notes_comments?.display_value || "");
    s("CTASK details", record.u_ctask_check || "");
    // "Deploy Start" / "Deploy End" are intentionally NOT populated from SNOW.
    // The Release Manager fills these in manually based on the actual deploy
    // window (SNOW start_date/end_date reflect the planned change window, not
    // the real deploy times). Leaving them blank on Refresh preserves any
    // manual entry and avoids overwriting truth with the SNOW estimate.
    s("Impacted CI", record.cmdb_ci?.display_value || "");
    s("CHG Type", record.type?.display_value || record.change_plan?.value || "");
    s("Date Submitted", record.opened_at?.display_value || "");
    s("Affected Groups", record.u_affected_groups?.display_value || "");
    s("Affected Locations", record.u_affected_locations?.display_value || "");
    s("Assignment Group", record.assignment_group?.display_value || "");
    const impactNames =
      typeof SNOW_IMPACT_DESCRIPTION_SOURCE_FIELDS !== "undefined" &&
      SNOW_IMPACT_DESCRIPTION_SOURCE_FIELDS.length
        ? SNOW_IMPACT_DESCRIPTION_SOURCE_FIELDS
        : ["u_impact_description", "impact_description"];
    s("Impact Description", snowDisplayFromRecord_(record, impactNames));

    // Division is intentionally NOT written here — the column is user-owned
    // (typically a VLOOKUP). REFRESH_PROTECTED_COLUMN_NAMES / CREATE_PROTECTED_COLUMN_NAMES
    // still keep it in the snapshot/restore list as a safety net in case any
    // future writer is added.

    if (map["CTASK List"] !== undefined) {
      s("CTASK List", snowData?.ctaskList || "No CTASKs found");
    }
  }

  return r;
}

function mergeProtectedColumnsOntoRows_(newRowsMatrix, map, expandedIssues, oldByKey) {
  expandedIssues.forEach((issue, i) => {
    const k = makeIssueCrRowKey(issue.key, issue.displayCr);
    const oldRow = oldByKey[k];
    if (!oldRow) return;
    const row = newRowsMatrix[i];
    PROTECTED_COLUMNS.forEach(name => {
      const ix = map[name];
      if (ix !== undefined) row[ix] = oldRow[ix];
    });
  });
}

/**
 * Sets pixel height (Format → Row height) for a contiguous run of rows.
 * Uses Sheet.setRowHeightsForced (preferred) so long/wrapped content does not re-expand rows.
 * Falls back to Sheet.setRowHeights, then per-row setRowHeightsForced / setRowHeight.
 */
function applyReleaseDataRowHeights_(sheet, firstRow, lastRow, heightPx) {
  if (!sheet || !heightPx || firstRow < 1 || lastRow < firstRow) return;
  const numRows = lastRow - firstRow + 1;
  if (typeof sheet.setRowHeightsForced === "function") {
    try {
      sheet.setRowHeightsForced(firstRow, numRows, heightPx);
      return;
    } catch (e) {
      // try fallbacks
    }
  }
  if (typeof sheet.setRowHeights === "function") {
    try {
      sheet.setRowHeights(firstRow, numRows, heightPx);
      return;
    } catch (e2) {
      // per-row
    }
  }
  for (let r = firstRow; r <= lastRow; r++) {
    if (typeof sheet.setRowHeightsForced === "function") {
      try {
        sheet.setRowHeightsForced(r, 1, heightPx);
      } catch (e3) {
        sheet.setRowHeight(r, heightPx);
      }
    } else {
      sheet.setRowHeight(r, heightPx);
    }
  }
}

/**
 * Applies the configured "Resize row" height to every used row in ReleaseSheet (row 1 through getLastRow).
 */
function applyReleaseSheetDataRowHeights_(sheet, cfg) {
  const h = Number(cfg && cfg.releaseDataRowHeight);
  if (!h || h < 8) return;
  const last = sheet.getLastRow();
  if (last < 1) return;
  applyReleaseDataRowHeights_(sheet, 1, last, h);
}

/**
 * Data rows (below the header) use regular weight and black text so template/defaults
 * (bold, colored) do not carry over after Create/Refresh.
 * @param {number} [numColumnsHint] e.g. headers.length — at least this many columns are formatted
 */
function applyDataRowsUnboldBlack_(sheet, headerRow, numColumnsHint) {
  if (!sheet || !headerRow || headerRow < 1) return;
  const lastRow = sheet.getLastRow();
  if (lastRow <= headerRow) return;
  const fromLastCol = sheet.getLastColumn();
  const w =
    numColumnsHint > 0 ? Math.max(fromLastCol, numColumnsHint) : Math.max(1, fromLastCol);
  if (w < 1) return;
  // Was: getRange(headerRow + 1, 1, lastRow, w) — that overshoots by `headerRow`
  // rows because the 3rd arg is row COUNT, not end row. Fixed to span only the
  // actual data rows.
  const numRows = lastRow - headerRow;
  const range = sheet.getRange(headerRow + 1, 1, numRows, w);
  range.setFontWeight("normal");
  range.setFontColor("#000000");
}

/**
 * One setBackgrounds on the CR # column (fast). Replaces per-row getRange().setBackground().
 */
function applyCrColumnBlueHighlights_(sheet, headerRow, expandedIssues, map) {
  const crCol = map["CR #"];
  if (crCol === undefined || !expandedIssues.length) return;
  const numRows = expandedIssues.length;
  const colBg = expandedIssues.map(issue => [
    issue.crCellHighlightBlue ? COLOR_MULTIPLE_CR : null
  ]);
  sheet.getRange(headerRow + 1, crCol + 1, numRows, 1).setBackgrounds(colBg);
}

/** Millis from SNOW start_date for this row's CR; missing CR/SNOW sorts last. */
function snowCrStartMillis_(issue, snowDataByCR) {
  const cr = (issue.displayCr || "").toString().trim();
  if (!cr || !snowDataByCR || !snowDataByCR[cr] || !snowDataByCR[cr].record) {
    return Number.MAX_SAFE_INTEGER;
  }
  const rec = snowDataByCR[cr].record;
  const raw = (rec.start_date && (rec.start_date.display_value || rec.start_date.value)) || "";
  const s = raw.toString().trim();
  if (!s) return Number.MAX_SAFE_INTEGER;
  const d = new Date(s);
  if (!isNaN(d.getTime())) return d.getTime();
  const fmts = ["yyyy-MM-dd HH:mm:ss", "MM/dd/yyyy HH:mm:ss", "yyyy-MM-dd", "MM/dd/yyyy"];
  for (let i = 0; i < fmts.length; i++) {
    try {
      const p = Utilities.parseDate(s, LA_TZ, fmts[i]);
      return p.getTime();
    } catch (e) {
      // try next format
    }
  }
  return Number.MAX_SAFE_INTEGER;
}

/** Empty CR # sorts after all real CHGs within the same Scrum Team. */
function crSortKey_(displayCr) {
  const c = (displayCr != null ? displayCr : "").toString().trim();
  if (!c) return "\uffff\uffff";
  return c.toUpperCase();
}

/** Empty Scrum Team sorts after named teams. */
function teamSortKey_(scrumTeam) {
  const t = (scrumTeam != null ? scrumTeam : "").toString().trim();
  if (!t) return "\uffff\uffff";
  return t.toLowerCase();
}

/**
 * Order rows: Scrum Team together, then CR # within team, then SNOW start, then JIRA key.
 */
function sortExpandedIssuesByScrumThenCr_(expanded, snowDataByCR) {
  if (!expanded || !expanded.length) return expanded || [];
  return expanded.slice().sort((a, b) => {
    const ta = teamSortKey_(a.scrumTeam);
    const tb = teamSortKey_(b.scrumTeam);
    if (ta !== tb) return ta < tb ? -1 : ta > tb ? 1 : 0;
    const ca = crSortKey_(a.displayCr);
    const cb = crSortKey_(b.displayCr);
    if (ca !== cb) return ca < cb ? -1 : ca > cb ? 1 : 0;
    const ma = snowCrStartMillis_(a, snowDataByCR);
    const mb = snowCrStartMillis_(b, snowDataByCR);
    if (ma !== mb) return ma - mb;
    return (a.key || "").toString().localeCompare((b.key || "").toString());
  });
}

/**
 * Populates data rows from Jira + SNOW (multi-CR expansion). Does not change A1.
 */
function populateReleaseDataRowsForIssues_(sheet, cfg, headerRow) {
  const dateObj = getReleaseDateFromReleaseSheet();
  const issues = fetchJiraIssuesByJql(buildJqlForDate(dateObj, cfg), cfg).map(normaliseIssue);
  const { headers, map } = getHeaderIndexMapSafe(sheet, headerRow);
  const snowWindowYmd = expandFridayWeekendJqlYmdsForY0_(dateObj);
  const allCrs = collectAllSnowCRNumbersFromIssues(issues);
  const blockerKeys = collectAllBlockerKeysFromIssues_(issues);
  const { snowDataByCR, blockerLookup } = fetchSnowAndBlockersInParallel_(
    allCrs,
    blockerKeys,
    cfg,
    map["CTASK List"] !== undefined
  );
  applyBlockerResolutionLookup_(issues, blockerLookup);
  let expanded = expandIssuesForDisplayRows(issues, snowDataByCR, snowWindowYmd);
  expanded = sortExpandedIssuesByScrumThenCr_(expanded, snowDataByCR);
  const techLeadLookup = buildTechLeadLookupFromMatrix_(sheet.getParent());
  const rows = expanded.map(i =>
    buildRowFromIssue(i, headers, map, cfg, snowDataByCR, techLeadLookup)
  );
  if (rows.length > 0) {
    withSpreadsheetRetry_(
      () => sheet.getRange(headerRow + 1, 1, rows.length, headers.length).setValues(rows),
      "populate:setValues mainRows"
    );
    withSpreadsheetRetry_(
      () => applyCrColumnBlueHighlights_(sheet, headerRow, expanded, map),
      "populate:applyCrColumnBlueHighlights"
    );
  }
  SpreadsheetApp.flush();
  withSpreadsheetRetry_(
    () => applyReleaseSheetDataRowHeights_(sheet, cfg),
    "populate:applyReleaseSheetDataRowHeights"
  );
  withSpreadsheetRetry_(
    () => applyDataRowsUnboldBlack_(sheet, headerRow, headers.length),
    "populate:applyDataRowsUnboldBlack"
  );
  // Drain row-heights + font-reset before the sheet-wide CF / data
  // validation pass; same boundary fix as createReleaseSheetInternal_.
  SpreadsheetApp.flush();
  withSpreadsheetRetry_(
    () => applyStatusFormattingToSheet_(sheet, headerRow),
    "populate:applyStatusFormatting"
  );
}

/* =================== Create Release Sheet =================== */

/**
 * Fetches Jira/SNOW and fills ReleaseSheet. Uses the date in A1 (via getReleaseDateFromReleaseSheet).
 * @param {{ fromStopRelease?: boolean }} [options] If fromStopRelease, skips the "existing data" archive dialog (fresh post-Stop tab) and skips the final success alert, returning a small result for Stop to show one combined message.
 * @returns {void|{ rowCount: number, releaseDateStr: string }} when fromStopRelease, returns the summary; otherwise void.
 */
function createReleaseSheetInternal_(options) {
  const __t0 = Date.now();
  const __mark = (label, since) => Logger.log(`[Create] ${label}: ${Date.now() - since}ms`);
  // __step: emits an entry log BEFORE running fn() and a duration log AFTER
  // it returns. If fn() throws or hangs, the last log line in the execution
  // log will be the "starting:" entry — that pinpoints which op was running
  // when the failure happened. Use this to wrap any op that we suspect can
  // throw "Service Spreadsheets timed out…".
  const __step = (label, fn) => {
    Logger.log(`[Create] starting: ${label}`);
    const t = Date.now();
    try {
      const r = fn();
      Logger.log(`[Create] done: ${label} (${Date.now() - t}ms)`);
      return r;
    } catch (e) {
      Logger.log(
        `[Create] FAILED: ${label} (${Date.now() - t}ms) — ${e && e.message ? e.message : e}`
      );
      throw e;
    }
  };
  options = options || {};
  const fromStop = !!options.fromStopRelease;
  Logger.log(`[Create] start (fromStopRelease=${fromStop})`);
  const cfg = getConfig();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(S_RELEASE);
  const master = ss.getSheetByName(S_MASTER);
  const ui = SpreadsheetApp.getUi();
  // Use the release date already in A1 for Jira + SNOW (set before call when from Stop).
  const releaseDateStr = getReleaseDateFromReleaseSheet();
  Logger.log(`[Create] release date from A1: ${releaseDateStr}`);

  // ==== Check if there's existing data and prompt for archival (not after a fresh Stop) ====
  if (!fromStop) {
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const hasExistingData = lastRow > cfg.sheetHeaderRow && lastCol > 0;

    if (hasExistingData) {
      // Cheap non-emptiness check: only read the JIRA Item column (or column A
      // as fallback). Reading the entire data range just to detect whether any
      // cell is non-empty was the second-largest avoidable cost in this path.
      const probeMap = getHeaderIndexMapSafe(sheet, cfg.sheetHeaderRow).map;
      const probeCol = (probeMap["JIRA Item"] !== undefined ? probeMap["JIRA Item"] : 0) + 1;
      const numProbeRows = lastRow - cfg.sheetHeaderRow;
      const probeVals = sheet.getRange(cfg.sheetHeaderRow + 1, probeCol, numProbeRows, 1).getValues();
      const hasContent = probeVals.some(r => r[0] !== "" && r[0] !== null && r[0] !== undefined);

      if (hasContent) {
        const archiveResponse = ui.alert(
          "⚠️ Existing Data Found",
          "The Release Sheet contains data from a previous release.\n\n" +
            "Has this data been archived to the Summary sheet?\n\n" +
            "• Click YES if already archived (will clear and create new sheet)\n" +
            "• Click NO to archive now before creating new sheet\n" +
            "• Click CANCEL to abort",
          ui.ButtonSet.YES_NO_CANCEL
        );

        if (archiveResponse === ui.Button.CANCEL) {
          ui.alert("Create Release Sheet cancelled.");
          return;
        }

        if (archiveResponse === ui.Button.NO) {
          const archiveNowResponse = ui.alert(
            "Archive to Summary",
            "Would you like to archive the current Release Sheet data to Summary now?\n\n" +
              "• Click YES to archive and then create new sheet\n" +
              "• Click NO to cancel (archive manually first)",
            ui.ButtonSet.YES_NO
          );

          if (archiveNowResponse === ui.Button.NO) {
            ui.alert(
              "Please archive the Release Sheet data to Summary first.\n\n" +
                "Use: Release Sheet menu → Archive to Summary"
            );
            return;
          }

          const archiveResult = archiveReleaseDataSilent(sheet, cfg.sheetHeaderRow);
          if (!archiveResult.success) {
            ui.alert("Archive failed: " + archiveResult.message);
            return;
          }

          ui.alert(
            `✅ Archived ${archiveResult.rowCount} rows to Summary sheet.\n\nNow creating new Release Sheet...`
          );
        }
      }
    }
  }

  // ==== Fetch Jira Issues ====
  const __tJira = Date.now();
  const issues = fetchJiraIssuesByJql(
    buildJqlForDate(releaseDateStr, cfg),
    cfg
  ).map(normaliseIssue);
  __mark("Jira fetch + normalize (" + issues.length + " issues)", __tJira);

  const snowWindowYmd = expandFridayWeekendJqlYmdsForY0_(releaseDateStr);
  const allCrs = collectAllSnowCRNumbersFromIssues(issues);
  const blockerKeys = collectAllBlockerKeysFromIssues_(issues);
  const { headers, map } = getHeaderIndexMapSafe(sheet, cfg.sheetHeaderRow);
  const __tNet = Date.now();
  const { snowDataByCR: snowDataByCRPre, blockerLookup } = fetchSnowAndBlockersInParallel_(
    allCrs,
    blockerKeys,
    cfg,
    map["CTASK List"] !== undefined
  );
  __mark(
    "SNOW + blockers parallel fetch (" +
      allCrs.length + " CRs, " +
      blockerKeys.length + " blocker keys)",
    __tNet
  );
  applyBlockerResolutionLookup_(issues, blockerLookup);
  let expandedPre = expandIssuesForDisplayRows(issues, snowDataByCRPre, snowWindowYmd);
  expandedPre = sortExpandedIssuesByScrumThenCr_(expandedPre, snowDataByCRPre);

  // ==== Clear master sheet for debugging/storage ====
  // Off by default — set Config.writeMasterDebugDump=true to re-enable. The
  // Master tab is purely diagnostic, and the JSON.stringify per issue +
  // separate-sheet write is the biggest avoidable cost in this function.
  if (cfg.writeMasterDebugDump && master) {
    __step("Master debug-dump write", () => {
      master.clear().getRange(1, 1, 1, 3).setValues([["Key", "Summary", "Raw"]]);
      if (expandedPre.length > 0) {
        master.getRange(2, 1, expandedPre.length, 3).setValues(
          expandedPre.map(i => [i.key, i.summary, JSON.stringify(i)])
        );
      }
    });
  }

  // ==== Clear release sheet except header (content AND ALL formatting) ====
  // .clear() is content + ALL formatting (background, fonts, borders), which
  // is a heavier op than .clearContent() and is the most common spot in
  // Create where "Service Spreadsheets timed out while accessing document"
  // surfaces. Wrap in withSpreadsheetRetry_ so a transient timeout retries
  // with another flush + small backoff instead of failing the whole Create.
  const currentLastRow = sheet.getLastRow();
  const currentLastCol = sheet.getLastColumn();
  Logger.log(
    `[Create] existing release range: ${currentLastRow} rows × ${currentLastCol} cols ` +
      `(headerRow=${cfg.sheetHeaderRow})`
  );

  // Snapshot user-managed columns (VLOOKUP formulas / manual edits) BEFORE
  // the clear so they survive the from-scratch rebuild. Restored after the
  // main-rows setValues below. See CREATE_PROTECTED_COLUMN_NAMES in
  // constants.gs. Note: .clear() below also wipes formatting (bg / font /
  // borders) on these columns — that's not restored here, only the
  // formula/value content is. Reapply via Sheets column-default formatting
  // if you want background colors preserved too.
  const __tCreateSnap = Date.now();
  const createProtectedSnapshot = snapshotProtectedColumns_(
    sheet, cfg.sheetHeaderRow, CREATE_PROTECTED_COLUMN_NAMES, map
  );
  Logger.log(
    `[Create] snapshotted ${Object.keys(createProtectedSnapshot).length} ` +
      `of ${(CREATE_PROTECTED_COLUMN_NAMES || []).length} protected col(s) in ` +
      `${Date.now() - __tCreateSnap}ms`
  );

  if (currentLastRow > cfg.sheetHeaderRow && currentLastCol > 0) {
    __step(
      `clearReleaseRange (${currentLastRow - cfg.sheetHeaderRow} × ${currentLastCol})`,
      () =>
        withSpreadsheetRetry_(
          () =>
            sheet
              .getRange(cfg.sheetHeaderRow + 1, 1, currentLastRow - cfg.sheetHeaderRow, currentLastCol)
              .clear(),
          "create:clearReleaseRange"
        )
    );
    // Drain the clear-format batch before the next setValues — large clears
    // queue a lot of internal state that the very next op has to flush, and
    // that's where the next-op timeout tends to surface.
    __step("flush after clearReleaseRange", () => SpreadsheetApp.flush());
  }

  const snowDataByCR = snowDataByCRPre;
  const expanded = expandedPre;

  // ==== Build rows ====
  const __tBuild = Date.now();
  const techLeadLookup = buildTechLeadLookupFromMatrix_(ss);
  const rows = expanded.map(i =>
    buildRowFromIssue(i, headers, map, cfg, snowDataByCR, techLeadLookup)
  );
  __mark("Build new rows matrix (" + rows.length + " rows × " + headers.length + " cols)", __tBuild);

  // ==== Write to sheet ====
  if (rows.length > 0) {
    __step(
      `setValues mainRows (${rows.length} × ${headers.length})`,
      () =>
        withSpreadsheetRetry_(
          () =>
            sheet
              .getRange(cfg.sheetHeaderRow + 1, 1, rows.length, headers.length)
              .setValues(rows),
          "create:setValues mainRows"
        )
    );
    __step("applyCrColumnBlueHighlights", () =>
      withSpreadsheetRetry_(
        () => applyCrColumnBlueHighlights_(sheet, cfg.sheetHeaderRow, expanded, map),
        "create:applyCrColumnBlueHighlights"
      )
    );
  }

  // Restore protected user-managed columns now that the data has been
  // written. Capped to the current last row by restoreProtectedColumns_
  // so we don't extend the data area past the freshly-written rows.
  __step("restoreProtectedColumns", () =>
    withSpreadsheetRetry_(
      () => restoreProtectedColumns_(sheet, cfg.sheetHeaderRow, createProtectedSnapshot),
      "create:restoreProtectedColumns"
    )
  );

  // ==== Lock the date (same calendar as A1 / releaseDateStr) ====
  setLockedReleaseDate(releaseDateStr);
  Logger.log(`[Create] locked release date set: ${releaseDateStr}`);

  // Drain any pending Spreadsheets ops from the large clear/write/highlight
  // batch above before further reads — otherwise the next read can throw
  // "Service timed out: Spreadsheets".
  __step("flush before post-write formatting", () => SpreadsheetApp.flush());
  const __tFmt = Date.now();
  __step("applyReleaseSheetDataRowHeights", () =>
    withSpreadsheetRetry_(
      () => applyReleaseSheetDataRowHeights_(sheet, cfg),
      "create:applyReleaseSheetDataRowHeights"
    )
  );
  __step("applyDataRowsUnboldBlack", () =>
    withSpreadsheetRetry_(
      () => applyDataRowsUnboldBlack_(sheet, cfg.sheetHeaderRow, headers.length),
      "create:applyDataRowsUnboldBlack"
    )
  );
  // Drain the row-height + font-reset batch before applyStatusFormattingToSheet_
  // touches sheet-wide CF rules + data validations. Without this flush, CF
  // rules has to drain a fat queue itself and that boundary is exactly where
  // we were hitting "Service Spreadsheets timed out…" in Create.
  __step("flush before applyStatusFormatting", () => SpreadsheetApp.flush());
  __step("applyStatusFormatting (sheet-wide CF + dropdowns)", () =>
    withSpreadsheetRetry_(
      () => applyStatusFormattingToSheet_(sheet, cfg.sheetHeaderRow),
      "create:applyStatusFormatting"
    )
  );
  __mark("Post-write formatting", __tFmt);
  __mark("TOTAL Create", __t0);

  const message =
    `Release Sheet created for ${releaseDateStr}.\n\n` +
    `Populated ${rows.length} rows (Jira + ServiceNow).\n\n` +
    `📅 Date locked to ${releaseDateStr} (from A1).\n\n` +
    `Note: Green columns will be locked after first Refresh.\n` +
    `Missing approval check runs automatically at 12 PM PST.`;

  if (fromStop) {
    return { rowCount: rows.length, releaseDateStr };
  }
  try {
    ui.alert(message);
  } catch (e) {
    Logger.log(message);
  }
}

function createReleaseSheet() {
  createReleaseSheetInternal_({ fromStopRelease: false });
}

/* =================== Stop Release (archive to Summary, roll A1 to next date) =================== */

/**
 * Parses a cell that may hold a release date; returns null if unusable (does not throw).
 */
function parseStopReleaseDateFromCell_(value) {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  try {
    return parseReleaseDateValue(value);
  } catch (e) {
    return null;
  }
}

/**
 * Release date for renaming the tab: first data row column D (4), else A1.
 */
function getStopReleaseRenameDate_(sheet, headerRow) {
  const fromD = sheet.getRange(headerRow + 1, 4).getValue();
  let dateObj = parseStopReleaseDateFromCell_(fromD);
  if (!dateObj || isNaN(dateObj.getTime())) {
    dateObj = parseStopReleaseDateFromCell_(sheet.getRange("A1").getValue());
  }
  if (!dateObj || isNaN(dateObj.getTime())) {
    throw new Error(
      "Could not read a release date from column D (first data row) or from A1 for the tab name."
    );
  }
  return dateObj;
}

/** Next calendar day (local JavaScript, midnight-normalized) for the “day after this release” prompt. */
function addOneCalendarDay_(d) {
  if (!d || isNaN(d.getTime())) return d;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + 1);
}

function stopReleaseTimeZone_() {
  return typeof LA_TZ === "string" ? LA_TZ : "America/Los_Angeles";
}

function formatDateForUserDisplay_(d) {
  if (!d || isNaN(d.getTime())) return "";
  return Utilities.formatDate(d, stopReleaseTimeZone_(), "M/d/yyyy");
}

/**
 * True if this calendar day is a Friday in Pacific (ties to expandFridayWeekendJqlYmdsForY0_ / F–Su window).
 */
function isFridayInPacificForReleaseDate_(d) {
  if (!d || isNaN(d.getTime())) return false;
  const tz = stopReleaseTimeZone_();
  const ymd = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  const parsed = Utilities.parseDate(ymd + " 12:00:00", tz, "yyyy-MM-dd HH:mm:ss");
  if (!parsed || isNaN(parsed.getTime())) return false;
  return Number(Utilities.formatDate(parsed, tz, "u")) === 5;
}

/**
 * Parses a typed date in the same formats as A1 (see parseReleaseDateValue).
 * @returns {Date|null}
 */
function parseUserTypedReleaseDate_(raw) {
  if (raw === null || raw === undefined) return null;
  const s = String(raw).trim();
  if (!s) return null;
  try {
    return parseReleaseDateValue(s);
  } catch (e) {
    return null;
  }
}

/**
 * Closes the current release cycle in place:
 *   1. Archives the current ReleaseSheet data to the Summary tab.
 *   2. Asks the user to confirm the next release date (defaults to current + 1 day).
 *   3. Updates ReleaseSheet!A1 to that date.
 *   4. Clears the old data and re-pulls Jira / ServiceNow for the new date,
 *      using the same load path as Create Release Sheet.
 *
 * Does NOT rename the tab or copy from TEMPLATE — the same ReleaseSheet tab is
 * reused. Depends on archiveReleaseDataSilent from Summary.gs.
 */
function StopRelease() {
  const ui = SpreadsheetApp.getUi();
  const __t0 = Date.now();
  const __mark = (label, since) => Logger.log(`[Stop] ${label}: ${Date.now() - since}ms`);
  // __step: emits "starting:" before fn() and "done:" / "FAILED:" after.
  // If the run dies, the last "starting:" line with no matching "done:" is
  // the exact op that hung — same pattern used in createReleaseSheetInternal_.
  const __step = (label, fn) => {
    Logger.log(`[Stop] starting: ${label}`);
    const t = Date.now();
    try {
      const r = fn();
      Logger.log(`[Stop] done: ${label} (${Date.now() - t}ms)`);
      return r;
    } catch (e) {
      Logger.log(
        `[Stop] FAILED: ${label} (${Date.now() - t}ms) — ${e && e.message ? e.message : e}`
      );
      throw e;
    }
  };
  Logger.log(`[Stop] start`);
  try {
    const cfg = getConfig();
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(S_RELEASE);
    if (!sheet) {
      Logger.log(`[Stop] release tab "${S_RELEASE}" not found — aborting`);
      ui.alert('Release sheet tab "' + S_RELEASE + '" was not found.');
      return;
    }
    if (typeof archiveReleaseDataSilent !== "function") {
      Logger.log(`[Stop] archiveReleaseDataSilent missing — aborting`);
      ui.alert(
        "archiveReleaseDataSilent is not defined in this project. Add your Summary/archive script so rows can be archived."
      );
      return;
    }

    const headerRow = cfg.sheetHeaderRow;
    const currentReleaseDate = getStopReleaseRenameDate_(sheet, headerRow);
    const nextCalDay = addOneCalendarDay_(currentReleaseDate);
    const currentStr = formatDateForUserDisplay_(currentReleaseDate);
    const nextStr = formatDateForUserDisplay_(nextCalDay);
    Logger.log(
      `[Stop] currentReleaseDate=${currentStr} suggestedNext=${nextStr} headerRow=${headerRow}`
    );

    const datePick = ui.alert(
      "Close Sheet — pick the next release date",
      "This will:\n" +
        "• Archive the current " + S_RELEASE + " data to the Summary sheet\n" +
        "• Update " + S_RELEASE + "!A1 to the next release date\n" +
        "• Clear the old data and re-pull Jira / ServiceNow for the new date\n\n" +
        "Current release date: " + currentStr + "\n\n" +
        "Use the next day (" + nextStr + ")?\n\n" +
        "• Yes — use " + nextStr + "\n" +
        "• No — I'll type a different date\n" +
        "• Cancel — do nothing",
      ui.ButtonSet.YES_NO_CANCEL
    );
    if (datePick === ui.Button.CANCEL) {
      Logger.log(`[Stop] user cancelled at date-pick prompt`);
      return;
    }

    let targetDate;
    if (datePick === ui.Button.YES) {
      targetDate = new Date(
        nextCalDay.getFullYear(),
        nextCalDay.getMonth(),
        nextCalDay.getDate()
      );
    } else {
      const promptR = ui.prompt(
        "New release date",
        "Enter the date for the new release (e.g. 4/30/2026 or 2026-04-30).",
        ui.ButtonSet.OK_CANCEL
      );
      if (promptR.getSelectedButton() !== ui.Button.OK) {
        Logger.log(`[Stop] user cancelled at custom-date prompt`);
        return;
      }
      const parsed = parseUserTypedReleaseDate_(promptR.getResponseText());
      if (!parsed || isNaN(parsed.getTime())) {
        Logger.log(
          `[Stop] unrecognized typed date: "${promptR.getResponseText()}" — aborting`
        );
        ui.alert("That date was not recognized. Use a format like 4/30/2026, 2026-04-30, or 04/30/2026.");
        return;
      }
      targetDate = new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
    }
    Logger.log(`[Stop] targetDate=${formatDateForUserDisplay_(targetDate)}`);

    if (isFridayInPacificForReleaseDate_(targetDate)) {
      Logger.log(`[Stop] target is a Friday — showing F/Sa/Su window note to user`);
      ui.alert(
        "Friday in Pacific (Los Angeles time)",
        "The new release is on a Friday in the Pacific time zone. Jira \"Production Release Date\" and ServiceNow matching will use a 3-day window: that Friday, Saturday, and Sunday (same as other Friday production releases).",
        ui.ButtonSet.OK
      );
    }

    const archiveResult = __step(
      "archiveReleaseDataSilent (current Release → Summary)",
      () => archiveReleaseDataSilent(sheet, headerRow)
    );
    if (!archiveResult || !archiveResult.success) {
      Logger.log(
        `[Stop] archive returned failure: ${archiveResult && archiveResult.message ? archiveResult.message : "unknown error"}`
      );
      ui.alert(
        "Archive failed: " +
          (archiveResult && archiveResult.message ? archiveResult.message : "unknown error")
      );
      return;
    }
    Logger.log(`[Stop] archived ${archiveResult.rowCount} row(s) to Summary`);

    // Drain any pending Summary-sheet writes (setValues, row heights, CF rules)
    // before kicking off createReleaseSheetInternal_, which itself does a large
    // clear/write batch. Without this flush both batches stack up against the
    // Spreadsheets service and trigger "Service Spreadsheets timed out while
    // accessing document".
    __step("flush after archive (drain Summary writes)", () => SpreadsheetApp.flush());

    // Roll A1 to the new release date and clear the locked-date guard so the
    // create-release flow can re-set it without prompting.
    __step(
      `set A1 to new release date (${formatDateForUserDisplay_(targetDate)})`,
      () => sheet.getRange(1, 1).setValue(targetDate)
    );
    __step("clearLockedReleaseDate", () => clearLockedReleaseDate());

    // Re-pull Jira / SNOW for the new date. fromStopRelease=true skips the
    // "existing data found" prompt (we just archived) and the success alert
    // (we'll show our own combined summary below).
    const createResult = __step(
      "createReleaseSheetInternal_(fromStopRelease=true)",
      () => createReleaseSheetInternal_({ fromStopRelease: true })
    );

    const rowsArchived =
      archiveResult.rowCount !== undefined && archiveResult.rowCount !== null
        ? String(archiveResult.rowCount)
        : "—";
    const nRows =
      createResult && typeof createResult.rowCount === "number" ? String(createResult.rowCount) : "0";
    const forDate = createResult && createResult.releaseDateStr ? createResult.releaseDateStr : "";
    __mark("TOTAL StopRelease", __t0);
    Logger.log(
      `[Stop] complete: archived=${rowsArchived} row(s), populated=${nRows} row(s) for ${forDate}`
    );

    ui.alert(
      "Close Sheet complete\n\n" +
        "• Archived " + rowsArchived + " row(s) to Summary\n" +
        "• " + S_RELEASE + "!A1 set to " + forDate + "\n" +
        "• Populated " + nRows + " Jira/ServiceNow row(s) for the new date"
    );
  } catch (e) {
    Logger.log(
      `[Stop] FAILED at top level after ${Date.now() - __t0}ms: ${e && e.message ? e.message : e}`
    );
    ui.alert("Close Sheet failed: " + e.message);
    Logger.log(e);
  }
}

/* =================== Refresh Release Sheet (rebuild rows; row key = JIRA Item + CR #) =================== */

/**
 * @param {boolean} headless - if true, skip UI date prompt and skip UI summary (scheduled path).
 * @returns {{ cancelled?: boolean, emptySheet?: boolean, dateObj: string, newCount: number, modifiedCount: number, removedCount: number, protectSkipped?: boolean }}
 */
function refreshReleaseSheetInternal_(headless) {
  const __t0 = Date.now();
  const __mark = (label, since) => Logger.log(`[Refresh] ${label}: ${Date.now() - since}ms`);
  const cfg = getConfig();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(S_RELEASE);
  const changeLog = ss.getSheetByName(S_CHANGELOG);

  if (!sheet) throw new Error("ReleaseSheet tab missing.");

  Logger.log(`[Refresh] start (headless=${!!headless})`);

  const dateObj = getReleaseDateFromReleaseSheet();

  if (!headless) {
    if (!validateDateChange(dateObj, false)) {
      return { cancelled: true, dateObj, newCount: 0, modifiedCount: 0, removedCount: 0 };
    }
  } else {
    validateDateChange(dateObj, true);
  }

  /** Menu refresh only: skip re-applying range protections when Config.refreshSkipProtectOnManualRefresh is true. */
  const skipProtect = !headless && cfg.refreshSkipProtectOnManualRefresh;

  const headerRow = cfg.sheetHeaderRow;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const { headers, map } = getHeaderIndexMapSafe(sheet, headerRow);
  const jiraItemCol = map["JIRA Item"];
  const crCol = map["CR #"];

  Logger.log(
    `[Refresh] sheet shape: lastRow=${lastRow} lastCol=${lastCol} ` +
      `headerRow=${headerRow} dataRows=${Math.max(0, lastRow - headerRow)} ` +
      `headers=${headers.length} for date ${dateObj}`
  );

  const __tJira = Date.now();
  const issues = fetchJiraIssuesByJql(buildJqlForDate(dateObj, cfg), cfg).map(normaliseIssue);
  __mark("Jira fetch + normalize (" + issues.length + " issues)", __tJira);

  const jiraMap = {};
  issues.forEach(i => {
    jiraMap[i.key] = i;
  });

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const changelogRows = [];

  if (lastRow <= headerRow) {
    populateReleaseDataRowsForIssues_(sheet, cfg, headerRow);
    if (!skipProtect) protectManualColumns(sheet, headerRow);
    if (!headless) {
      SpreadsheetApp.getUi().alert(
        "Release Sheet had no data rows — populated from Jira for " + dateObj + "."
      );
    }
    return {
      emptySheet: true,
      dateObj,
      newCount: Math.max(0, sheet.getLastRow() - headerRow),
      modifiedCount: 0,
      removedCount: 0,
      protectSkipped: skipProtect
    };
  }

  const snowWindowYmd = expandFridayWeekendJqlYmdsForY0_(dateObj);
  const allCrs = collectAllSnowCRNumbersFromIssues(issues);
  const blockerKeys = collectAllBlockerKeysFromIssues_(issues);
  const __tNet = Date.now();
  const { snowDataByCR, blockerLookup } = fetchSnowAndBlockersInParallel_(
    allCrs,
    blockerKeys,
    cfg,
    map["CTASK List"] !== undefined
  );
  __mark(
    "SNOW + blockers parallel fetch (" +
      allCrs.length + " CRs, " +
      blockerKeys.length + " blocker keys)",
    __tNet
  );
  applyBlockerResolutionLookup_(issues, blockerLookup);

  let expanded = expandIssuesForDisplayRows(issues, snowDataByCR, snowWindowYmd);
  expanded = sortExpandedIssuesByScrumThenCr_(expanded, snowDataByCR);

  const __tReadOld = Date.now();
  const oldData = withSpreadsheetRetry_(
    () => sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).getValues(),
    "refresh:readOldData"
  );
  __mark(
    "Read old data (" + (lastRow - headerRow) + " rows × " + lastCol + " cols)",
    __tReadOld
  );

  const __tIndex = Date.now();
  const oldByKey = {};
  const oldKeys = new Set();
  oldData.forEach(row => {
    const ji = (row[jiraItemCol] || "").toString().trim();
    const cr = crCol !== undefined ? (row[crCol] || "").toString().trim() : "";
    if (!ji) return;
    const k = makeIssueCrRowKey(ji, cr);
    oldByKey[k] = row;
    oldKeys.add(k);
  });
  __mark("Index old rows by key", __tIndex);

  const __tTLLookup = Date.now();
  const techLeadLookup = buildTechLeadLookupFromMatrix_(ss);
  __mark("Tech-lead matrix lookup", __tTLLookup);

  const __tBuild = Date.now();
  const newRowsMatrix = expanded.map(issue =>
    buildRowFromIssue(issue, headers, map, cfg, snowDataByCR, techLeadLookup)
  );
  mergeProtectedColumnsOntoRows_(newRowsMatrix, map, expanded, oldByKey);
  __mark("Build new rows matrix (" + newRowsMatrix.length + " rows)", __tBuild);

  const hw = headers.length;
  const padRow = row => {
    const o = [];
    for (let c = 0; c < hw; c++) {
      o.push(c < row.length ? row[c] : "");
    }
    return o;
  };

  const __tDiff = Date.now();
  const newKeyOrder = expanded.map(e => makeIssueCrRowKey(e.key, e.displayCr));
  const newKeySet = new Set(newKeyOrder);
  const newKeyToIndex = {};
  newKeyOrder.forEach((k, i) => {
    newKeyToIndex[k] = i;
  });

  newKeyOrder.forEach((k, idx) => {
    if (!oldKeys.has(k)) {
      const issue = expanded[idx];
      changelogRows.push([
        now,
        "NEW",
        issue.key,
        truncateForChangelog_(issue.summary || ""),
        `New logical row (${issue.displayCr || "no CR"})`
      ]);
    }
  });

  oldKeys.forEach(k => {
    if (newKeySet.has(k)) return;
    const oldRow = oldByKey[k];
    const ji = (oldRow[jiraItemCol] || "").toString().trim();
    if (ji && !jiraMap[ji]) {
      changelogRows.push([
        now,
        "REMOVED",
        ji,
        "",
        "Ticket no longer in Jira for this release date"
      ]);
    } else {
      changelogRows.push([
        now,
        "REMOVED",
        ji || k.split("||")[0] || k,
        "",
        "Logical row removed (CR / SNOW start date vs release date)"
      ]);
    }
  });

  const compareCols = [
    "Jira Status",
    "Jira Resolution",
    "Jira Assignee",
    "JIRA Assignee",
    "Approvals",
    "Jira Prod Release Date",
    "QA Artifacts Link",
    "UAT Artifacts Link",
    "Severity",
    "Impact Description"
  ];
  // Track which cells changed so we can paint them COLOR_MODIFIED after writing.
  // Each entry is { rowOffset, colIdx } where rowOffset is 0-based against newRowsMatrix
  // and colIdx is 0-based column index.
  const modifiedCells = [];
  oldKeys.forEach(k => {
    if (!newKeySet.has(k)) return;
    const idx = newKeyToIndex[k];
    const newRow = newRowsMatrix[idx];
    const oldRow = oldByKey[k];
    compareCols.forEach(colName => {
      if (PROTECTED_COLUMNS && PROTECTED_COLUMNS.includes(colName)) return;
      const ci = map[colName];
      if (ci === undefined) return;
      const ov = normalizeForComparison(oldRow[ci]);
      const nv = normalizeForComparison(newRow[ci]);
      if (ov !== nv) {
        changelogRows.push([
          now,
          "MODIFIED",
          expanded[idx].key,
          colName,
          truncateForChangelog_(`${ov} → ${nv}`)
        ]);
        modifiedCells.push({ rowOffset: idx, colIdx: ci });
      }
    });
  });

  const removedAppendRows = [];
  oldData.forEach(row => {
    const ji = (row[jiraItemCol] || "").toString().trim();
    if (ji && !jiraMap[ji]) removedAppendRows.push(padRow(row));
  });
  __mark(
    "Diff (NEW=" +
      changelogRows.filter(r => r[1] === "NEW").length +
      ", MODIFIED=" +
      modifiedCells.length +
      ", REMOVED=" +
      changelogRows.filter(r => r[1] === "REMOVED").length +
      ")",
    __tDiff
  );

  const mainRows = newRowsMatrix.map(padRow);
  const totalMain = mainRows.length;
  const totalAppend = removedAppendRows.length;
  const oldHeight = Math.max(0, lastRow - headerRow);
  const clearHeight = Math.max(oldHeight, totalMain + totalAppend, 1);
  const clearWidth = Math.max(lastCol, hw);

  // Snapshot user-managed columns (VLOOKUP formulas / manual edits) BEFORE the
  // clear. Restored after the data + REMOVED-rows write below. See
  // REFRESH_PROTECTED_COLUMN_NAMES in constants.gs.
  const __tSnap = Date.now();
  const protectedSnapshot = snapshotProtectedColumns_(
    sheet, headerRow, REFRESH_PROTECTED_COLUMN_NAMES, map
  );
  __mark(
    "Snapshot protected cols (" + Object.keys(protectedSnapshot).length +
      " of " + (REFRESH_PROTECTED_COLUMN_NAMES || []).length + ")",
    __tSnap
  );

  // Drain anything queued by the heavy read of oldData before issuing the
  // big clear — this is the same boundary fix used in
  // createReleaseSheetInternal_, where unflushed reads + a follow-up
  // clear/setValues was the most common spot for "Service Spreadsheets
  // timed out while accessing document".
  SpreadsheetApp.flush();

  const __tClear = Date.now();
  // clearContent() only — the leftover yellow/blue/red highlights from the
  // previous refresh are cleared as part of the setBackgrounds(bg) paint
  // below (the bg matrix is filled with null for non-highlighted cells, so
  // it doubles as a "reset" pass). Dropping the chained setBackground(null)
  // saves a full sheet-wide formatting round-trip — that op was the single
  // biggest contributor to the Refresh timeouts on large releases.
  withSpreadsheetRetry_(
    () => sheet.getRange(headerRow + 1, 1, clearHeight, clearWidth).clearContent(),
    "refresh:clearContent"
  );
  __mark("Clear data range (" + clearHeight + " × " + clearWidth + ")", __tClear);

  // Drain the clear before the next setValues — large clears queue a lot
  // of internal state that the very next op has to flush, and that's where
  // the next-op timeout tends to surface.
  SpreadsheetApp.flush();

  if (totalMain > 0) {
    const __tWriteMain = Date.now();
    withSpreadsheetRetry_(
      () => sheet.getRange(headerRow + 1, 1, totalMain, hw).setValues(mainRows),
      "refresh:setValues mainRows"
    );
    __mark("setValues main rows (" + totalMain + " × " + hw + ")", __tWriteMain);

    // JIRA Item hyperlinks are now embedded as HYPERLINK formulas inside
    // mainRows (see makeJiraItemHyperlinkFormula_), so a separate
    // setRichTextValues round-trip is no longer needed.

    // Always paint backgrounds across the full main-rows range so prior
    // highlights are cleared even when the new state has no modified or
    // blue cells. The bg matrix is mostly nulls (resets), with COLOR_MODIFIED
    // for changed cells and COLOR_MULTIPLE_CR overriding the CR # column
    // for blue-flagged rows.
    //
    // NOTE: rows whose Validation/Deployment Status drives a row-color rule
    // will visually override these direct cell backgrounds because
    // conditional formatting wins. The signals are most visible on rows
    // that haven't had a status set yet, which matches how refresh feedback
    // should work.
    const crColLocal = map["CR #"];
    const __tBg = Date.now();
    const bg = Array.from({ length: totalMain }, () => Array(hw).fill(null));
    modifiedCells.forEach(mc => {
      if (mc.rowOffset >= 0 && mc.rowOffset < totalMain && mc.colIdx >= 0 && mc.colIdx < hw) {
        bg[mc.rowOffset][mc.colIdx] = COLOR_MODIFIED;
      }
    });
    if (crColLocal !== undefined && crColLocal >= 0 && crColLocal < hw) {
      for (let i = 0; i < totalMain && i < expanded.length; i++) {
        if (expanded[i].crCellHighlightBlue) {
          bg[i][crColLocal] = COLOR_MULTIPLE_CR;
        }
      }
    }
    withSpreadsheetRetry_(
      () => sheet.getRange(headerRow + 1, 1, totalMain, hw).setBackgrounds(bg),
      "refresh:setBackgrounds mainRows"
    );
    __mark(
      "Paint backgrounds combined (modified=" + modifiedCells.length +
        ", " + totalMain + " × " + hw + ")",
      __tBg
    );
  }

  if (totalAppend > 0) {
    const __tAppend = Date.now();
    const start = headerRow + 1 + totalMain;

    // Refresh the "Jira Prod Release Date" cell on each REMOVED row from
    // Jira so the user can see where the ticket moved to instead of the
    // stale value from the sheet snapshot. Done BEFORE the JIRA Item
    // hyperlink wrap so the cell still holds the plain key when we extract
    // it.
    const prodDateCol = map["Jira Prod Release Date"];
    if (prodDateCol !== undefined && jiraItemCol !== undefined) {
      const removedKeys = removedAppendRows
        .map(r => (r[jiraItemCol] || "").toString().trim())
        .filter(Boolean);
      if (removedKeys.length) {
        const tFetch = Date.now();
        const lookup = fetchProdReleaseDateByKeys_(removedKeys, cfg);
        Logger.log(
          "Refreshed Prod Release Date for " + Object.keys(lookup).length +
            "/" + removedKeys.length + " REMOVED row(s) in " +
            (Date.now() - tFetch) + "ms"
        );
        removedAppendRows.forEach(row => {
          const ji = (row[jiraItemCol] || "").toString().trim();
          if (!ji) return;
          const fresh = lookup[ji];
          if (fresh !== undefined) row[prodDateCol] = fresh;
        });
      }
    }

    // oldData was read with getValues(), so the JIRA Item cell holds the
    // displayed key text. Re-wrap it as a HYPERLINK formula so the appended
    // REMOVED row stays clickable without a separate setRichTextValues call.
    if (jiraItemCol !== undefined) {
      removedAppendRows.forEach(row => {
        const ji = (row[jiraItemCol] || "").toString().trim();
        if (ji) row[jiraItemCol] = makeJiraItemHyperlinkFormula_(ji, cfg);
      });
    }
    withSpreadsheetRetry_(
      () => sheet.getRange(start, 1, totalAppend, hw).setValues(removedAppendRows),
      "refresh:setValues removedRows"
    );
    withSpreadsheetRetry_(
      () => sheet.getRange(start, 1, totalAppend, hw).setBackground(COLOR_REMOVED),
      "refresh:setBackground removedRows"
    );
    __mark("Append REMOVED rows (" + totalAppend + ")", __tAppend);
  }

  // Restore the snapshotted protected columns now that the data + REMOVED
  // rows have been written. Capped to the current last row so we don't
  // extend the data area past the new write.
  const __tRestore = Date.now();
  restoreProtectedColumns_(sheet, headerRow, protectedSnapshot);
  __mark("Restore protected cols", __tRestore);

  // Drain the queued Release-sheet write batch before crossing to a different
  // sheet (ChangeLog). Without this, `changeLog.getLastRow()` forces Apps
  // Script to flush the entire pending op queue against the Spreadsheets
  // service in one shot, which routinely exceeds the per-op timeout and
  // throws "Service timed out: Spreadsheets". Same pattern as Summary.gs
  // archiveReleaseDataSilent.
  const __tDrain = Date.now();
  SpreadsheetApp.flush();
  __mark("Pre-ChangeLog flush (drain Release-sheet writes)", __tDrain);

  if (changelogRows.length > 0 && changeLog) {
    const __tCl = Date.now();
    withSpreadsheetRetry_(
      () =>
        changeLog
          .getRange(changeLog.getLastRow() + 1, 1, changelogRows.length, changelogRows[0].length)
          .setValues(changelogRows),
      "refresh:changelog setValues"
    );
    __mark("ChangeLog append (" + changelogRows.length + " rows)", __tCl);
  }

  // Defer the slow tail (protections + row heights + font reset) to a
  // background trigger so the user-visible alert pops as soon as the data
  // and ChangeLog are on the sheet. Headless / scheduled runs honor the
  // same flag so they too return faster (the trigger still runs inline on
  // the Apps Script servers).
  const deferTail = !!cfg.refreshDeferTail;
  if (deferTail) {
    const __tSchedule = Date.now();
    scheduleRefreshTailTrigger_(skipProtect);
    __mark("Schedule deferred tail trigger", __tSchedule);
  } else {
    runRefreshTailInline_(sheet, headerRow, cfg, hw, skipProtect, __mark);
  }

  __mark("TOTAL Refresh" + (deferTail ? " (tail deferred)" : ""), __t0);
  const removedCount = changelogRows.filter(r => r[1] === "REMOVED").length;
  const modifiedCount = changelogRows.filter(r => r[1] === "MODIFIED").length;
  const newCount = changelogRows.filter(r => r[1] === "NEW").length;

  return {
    dateObj,
    newCount,
    modifiedCount,
    removedCount,
    protectSkipped: skipProtect,
    tailDeferred: deferTail
  };
}

/* =================== Deferred Refresh tail =================== */

/**
 * Trigger handler name. Kept stable so we can deduplicate pre-existing
 * triggers before scheduling a new one (avoids a thundering-herd of stale
 * triggers if a refresh fails mid-flight).
 */
const REFRESH_TAIL_TRIGGER_HANDLER = "__refreshTailHandler_";

/**
 * Property keys used to communicate state to the deferred tail handler.
 * (ScriptApp.newTrigger does not let us pass arguments, so we stash the
 * "skipProtect" decision in document properties.)
 */
const REFRESH_TAIL_PROP_SKIP_PROTECT = "refreshTail_skipProtect";

/**
 * Schedules a one-shot time-based trigger ~1.5 seconds out that runs the
 * Refresh tail (protections + row heights + font reset). Removes any
 * pre-existing tail trigger first so we don't accumulate orphans.
 */
function scheduleRefreshTailTrigger_(skipProtect) {
  try {
    removePreexistingRefreshTailTriggers_();
    PropertiesService.getDocumentProperties().setProperty(
      REFRESH_TAIL_PROP_SKIP_PROTECT,
      skipProtect ? "true" : "false"
    );
    ScriptApp.newTrigger(REFRESH_TAIL_TRIGGER_HANDLER)
      .timeBased()
      .after(1500)
      .create();
  } catch (e) {
    Logger.log("scheduleRefreshTailTrigger_ failed: " + (e && e.message ? e.message : e));
    // Fallback: run inline if scheduling failed (no async backup).
    try {
      const ss = SpreadsheetApp.getActive();
      const sheet = ss.getSheetByName(S_RELEASE);
      const cfg = getConfig();
      const headerRow = cfg.sheetHeaderRow;
      const hw = sheet.getLastColumn();
      runRefreshTailInline_(sheet, headerRow, cfg, hw, skipProtect, function () {});
    } catch (e2) {
      Logger.log("scheduleRefreshTailTrigger_ inline-fallback also failed: " + e2);
    }
  }
}

function removePreexistingRefreshTailTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === REFRESH_TAIL_TRIGGER_HANDLER) {
      try {
        ScriptApp.deleteTrigger(t);
      } catch (e) {
        Logger.log("removePreexistingRefreshTailTriggers_: " + e);
      }
    }
  });
}

/**
 * One-shot trigger handler. Runs the deferred Refresh tail and then deletes
 * its own trigger so it can't fire twice.
 */
function __refreshTailHandler_(e) {
  const __t0 = Date.now();
  const __mark = (label, since) => Logger.log(`[RefreshTail] ${label}: ${Date.now() - since}ms`);
  try {
    const props = PropertiesService.getDocumentProperties();
    const skipProtect = props.getProperty(REFRESH_TAIL_PROP_SKIP_PROTECT) === "true";
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(S_RELEASE);
    if (!sheet) return;
    const cfg = getConfig();
    const headerRow = cfg.sheetHeaderRow;
    const hw = sheet.getLastColumn();
    runRefreshTailInline_(sheet, headerRow, cfg, hw, skipProtect, __mark);
    __mark("TOTAL RefreshTail", __t0);
  } catch (err) {
    Logger.log("__refreshTailHandler_ error: " + (err && err.message ? err.message : err));
  } finally {
    // Remove this trigger (and any stragglers) so we don't fire again.
    if (e && e.triggerUid) {
      try {
        ScriptApp.getProjectTriggers().forEach(t => {
          if (t.getUniqueId() === e.triggerUid) ScriptApp.deleteTrigger(t);
        });
      } catch (e2) {
        Logger.log("__refreshTailHandler_ trigger cleanup error: " + e2);
      }
    } else {
      removePreexistingRefreshTailTriggers_();
    }
  }
}

/**
 * Runs the protections + row heights + font-reset tail synchronously.
 * Shared between the inline path (refreshDeferTail = false) and the
 * deferred trigger handler.
 */
function runRefreshTailInline_(sheet, headerRow, cfg, hw, skipProtect, __mark) {
  const __tProtect = Date.now();
  if (!skipProtect) {
    withSpreadsheetRetry_(
      () => protectManualColumns(sheet, headerRow),
      "refresh:protectManualColumns"
    );
  }
  __mark("protectManualColumns" + (skipProtect ? " (skipped)" : ""), __tProtect);

  const __tFlush = Date.now();
  SpreadsheetApp.flush();
  __mark("SpreadsheetApp.flush()", __tFlush);

  const __tHeights = Date.now();
  applyReleaseSheetDataRowHeights_(sheet, cfg);
  __mark("applyReleaseSheetDataRowHeights_", __tHeights);

  const __tUnbold = Date.now();
  applyDataRowsUnboldBlack_(sheet, headerRow, hw);
  __mark("applyDataRowsUnboldBlack_", __tUnbold);
  // applyStatusFormattingToSheet_ is intentionally NOT called here. The
  // dropdowns + row-color CF rules persist on the sheet across refreshes,
  // and re-applying them on every Refresh was a noticeable cost. They get
  // applied during Create (or the empty-sheet branch above via
  // populateReleaseDataRowsForIssues_). Run "Create Release Sheet" once if
  // you change the *_OPTIONS color maps and want them re-installed.

  // Pre-Release Check audit + Pre-Release Check Summary write-back. Lives in
  // the tail (not refreshReleaseSheetInternal_) so the user-visible refresh
  // alert pops as soon as data is on the sheet — the audit work happens
  // here, in the deferred trigger when refreshDeferTail is true (default).
  // Wrapped in try/catch so any audit failure cannot break refresh.
  const __tAudit = Date.now();
  try {
    if (typeof runAuditAndApplySummaryAfterRefresh_ === "function") {
      runAuditAndApplySummaryAfterRefresh_();
    } else {
      Logger.log("[RefreshTail] audit step skipped: runAuditAndApplySummaryAfterRefresh_ not defined.");
    }
  } catch (auditErr) {
    Logger.log("[RefreshTail] audit step failed: " + (auditErr && auditErr.message ? auditErr.message : auditErr));
  }
  __mark("Pre-Release Check audit + write-back", __tAudit);
}

function refreshReleaseSheet() {
  const r = refreshReleaseSheetInternal_(false);
  if (r.cancelled) return;

  if (r.emptySheet) return;

  const protectLine = r.protectSkipped
    ? "Column protection skipped (Config.refreshSkipProtectOnManualRefresh). Set it false and refresh once to re-lock green columns.\n\n"
    : r.tailDeferred
      ? "🔒 Manual column protections + row formatting will finish in the background in ~1–2 seconds.\n\n"
      : "🔒 Manual columns protected.\n\n";

  const message =
    `Refresh completed for ${r.dateObj}.\n\n` +
    `🟢 New logical rows: ${r.newCount}\n` +
    `🟡 Modified fields: ${r.modifiedCount}\n` +
    `🔴 Removed (changelog): ${r.removedCount}\n\n` +
    protectLine +
    `Note: Missing approval check runs automatically at 12 PM PST.\n\n` +
    `Color legend:\n` +
    `• 🟡 Yellow cell = field modified since last refresh (Jira Status, Resolution, Assignee, Prod Release Date, etc.)\n` +
    `• 🔵 Light blue = CR # (multiple ServiceNow Related Ticket values / SNOW date rules)\n` +
    `• 🔴 Red appended rows = ticket no longer on this release in Jira\n` +
    `• Row background (Deployment / Validation Status dropdowns) overrides the yellow modified-cell highlight when a status is set.`;

  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log(message);
  }
}


/**
 * DailyReleaseCheck.gs — separate from Release-data.gs
 * (Release-data only adds the ITRM menu item that calls runDailyReleaseCheck.)
 *
 * Daily view of ServiceNow CRs (today) vs CR #s on ReleaseSheet.
 * After writes, all rows (header + data) are forced to 21px via setRowHeightsForced, same idea as ReleaseSheet.
 */

const DAILY_RELEASE_CHECK_ROW_HEIGHT_PX = 21;
/** Aligned with Release sheet Friday/Production Release window (Fri+Sat+Sun) */
const DAILY_CHECK_PST_TZ = "America/Los_Angeles";

/**
 * Number of leading rows on the Daily Release Check tab that runDailyReleaseCheck
 * never touches (no writes, no clears, no row-height forcing). Use these rows for
 * manual notes / summary formulas — they survive every run.
 *
 * The header row sits directly below this block (DAILY_CHECK_HEADER_ROW) and data
 * rows start at DAILY_CHECK_DATA_START_ROW. To change how many rows are reserved,
 * edit DAILY_CHECK_PROTECTED_TOP_ROWS only — the other two derive from it.
 */
const DAILY_CHECK_PROTECTED_TOP_ROWS = 3;
const DAILY_CHECK_HEADER_ROW = DAILY_CHECK_PROTECTED_TOP_ROWS + 1;
const DAILY_CHECK_DATA_START_ROW = DAILY_CHECK_HEADER_ROW + 1;

/**
 * @param {string} ymd
 * @param {number} n
 * @param {string} tz
 * @returns {string} yyyy-MM-dd
 */
function addCalendarDaysYmdForDailyCheck_(ymd, n, tz) {
  const d = Utilities.parseDate(ymd + " 12:00:00", tz, "yyyy-MM-dd HH:mm:ss");
  d.setTime(d.getTime() + n * 24 * 60 * 60 * 1000);
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

/**
 * true when "today" in Pacific is a Friday (same DOW as expandFridayWeekendJqlYmdsForY0_ in Release-data).
 */
function isFridayTodayInDailyCheckPst_() {
  const ymd = Utilities.formatDate(new Date(), DAILY_CHECK_PST_TZ, "yyyy-MM-dd");
  const parsed = Utilities.parseDate(ymd + " 12:00:00", DAILY_CHECK_PST_TZ, "yyyy-MM-dd HH:mm:ss");
  if (!parsed || isNaN(parsed.getTime())) return false;
  return Number(Utilities.formatDate(parsed, DAILY_CHECK_PST_TZ, "u")) === 5;
}

/**
 * start_date in [Fri 00:00, Mon 00:00) in LA, encoded for Table API. Matches Fri+Sat+Sun when today is that Friday in LA.
 * @returns {{ query: string, fridayWeekend: boolean }}
 */
function getDailyCheckSnowStartDateQueryInfo_() {
  const notClosed = "state!=6";
  if (!isFridayTodayInDailyCheckPst_()) {
    return {
      fridayWeekend: false,
      query:
        "start_dateONToday@javascript:gs.beginningOfToday()@javascript:gs.endOfToday()^" + notClosed
    };
  }
  const friY = Utilities.formatDate(new Date(), DAILY_CHECK_PST_TZ, "yyyy-MM-dd");
  const monY = addCalendarDaysYmdForDailyCheck_(friY, 3, DAILY_CHECK_PST_TZ);
  const fri0 = Utilities.parseDate(friY + " 00:00:00", DAILY_CHECK_PST_TZ, "yyyy-MM-dd HH:mm:ss");
  const mon0 = Utilities.parseDate(monY + " 00:00:00", DAILY_CHECK_PST_TZ, "yyyy-MM-dd HH:mm:ss");
  const a = Utilities.formatDate(fri0, "UTC", "yyyy-MM-dd HH:mm:ss");
  const b = Utilities.formatDate(mon0, "UTC", "yyyy-MM-dd HH:mm:ss");
  // Lower inclusive Fri; upper exclusive Mon so all of Sunday is included (any time component).
  return {
    fridayWeekend: true,
    query: "start_date>=" + a + "^start_date<" + b + "^" + notClosed
  };
}

/**
 * Forces rows in [startRow..getLastRow()] to the same pixel height, preferring
 * "forced" so long text does not expand rows. Defaults to startRow=1 for
 * backwards compatibility; callers pass DAILY_CHECK_HEADER_ROW to leave any
 * user-customized row heights in the protected top rows untouched.
 */
function applyDailyCheckSheetRowHeights21_(sheet, startRow) {
  if (!sheet) return;
  const start = startRow && startRow > 0 ? startRow : 1;
  const last = sheet.getLastRow();
  if (last < start) return;
  const numRows = last - start + 1;
  const h = DAILY_RELEASE_CHECK_ROW_HEIGHT_PX;
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

function runDailyReleaseCheck() {
  const ss = SpreadsheetApp.getActive();
  const cfg = getConfig();
  const AFTER_HOUR = 17; // 5 PM PST

  /***********************
   * SHEETS
   ***********************/
  const releaseSheet = ss.getSheetByName(S_RELEASE);
  if (!releaseSheet) {
    throw new Error("ReleaseSheet not found (expected tab name: " + S_RELEASE + ").");
  }

  const outputSheet = ss.getSheetByName("Daily Release Check") || ss.insertSheet("Daily Release Check");

  /***********************
   * HELPERS
   ***********************/
  const unwrapSN = field => {
    if (!field) return "";
    if (typeof field === "string") return field;
    if (Array.isArray(field)) return field.map(f => unwrapSN(f)).join(", ");
    if (typeof field === "object") return field.display_value || field.value || "";
    return "";
  };

  const getSysIdFromRef = ref => {
    if (!ref?.link) return "";
    const match = ref.link.match(/\/([0-9a-f]{32})$/);
    return match ? match[1] : "";
  };

  const toPST = dateStr => {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    if (isNaN(d)) return "";
    return Utilities.formatDate(d, "PST", "yyyy-MM-dd HH:mm");
  };

  const truncate = (str, max = 250) => (str ? str.substring(0, max) : "");

  /***********************
   * READ CRs FROM RELEASE SHEET
   ***********************/
  const releaseHeaders = releaseSheet.getRange(1, 1, 1, releaseSheet.getLastColumn()).getValues()[0];

  const releaseCrCol = releaseHeaders.findIndex(h => String(h).trim().toUpperCase() === "CR #");
  if (releaseCrCol === -1) throw new Error('"CR #" column not found in ReleaseSheet');

  const lastDataRow = releaseSheet.getLastRow();
  const crValues =
    lastDataRow < 2
      ? []
      : releaseSheet
          .getRange(2, releaseCrCol + 1, lastDataRow, 1)
          .getValues()
          .flat();
  const releaseCRs = new Set(
    crValues
      .map(v => String(v).trim().toUpperCase())
      .filter(Boolean)
  );

  /***********************
   * FETCH TODAY'S CRs FROM SERVICENOW
   * (Fri in LA: Fri+Sat+Sun start dates; any other day: "today" only in SN user/session TZ for ONToday)
   ***********************/
  const { query: sysparmQuery, fridayWeekend: snowFriThroughSun } = getDailyCheckSnowStartDateQueryInfo_();

  const url =
    `${cfg.snBaseUrl}/api/now/table/change_request` +
    `?sysparm_query=${encodeURIComponent(sysparmQuery)}` +
    `&sysparm_fields=sys_id,number,type,short_description,u_impact_description,approval,state,start_date,end_date,u_business_justification,description,assignment_group,cmdb_ci,u_affected_groups,requested_by,u_requested_for,sys_created_on` +
    `&sysparm_display_value=true` +
    `&sysparm_limit=500`;

  Logger.log(
    "Fetching CRs from ServiceNow" +
      (snowFriThroughSun ? " (Fri+Sat+Sun, Pacific)" : " (today)") +
      ": " +
      url
  );

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    headers: {
      Authorization: "Basic " + Utilities.base64Encode(cfg.snUser + ":" + cfg.snPass),
      Accept: "application/json"
    },
    muteHttpExceptions: true
  });
  if (response.getResponseCode() >= 400) {
    throw new Error("ServiceNow: " + response.getResponseCode() + " " + response.getContentText());
  }

  const snowData = JSON.parse(response.getContentText()).result || [];
  Logger.log(`Fetched ${snowData.length} CRs from ServiceNow`);

  /***********************
   * FETCH USER DEPARTMENTS & DIVISIONS
   ***********************/
  const userSysIds = [
    ...new Set(snowData.map(cr => getSysIdFromRef(cr.u_requested_for)).filter(Boolean))
  ];

  Logger.log(`Found ${userSysIds.length} unique u_requested_for sys_ids: ${userSysIds.join(", ")}`);

  const userCache = {};
  const authH = { Authorization: "Basic " + Utilities.base64Encode(cfg.snUser + ":" + cfg.snPass) };

  userSysIds.forEach(sysId => {
    const userUrl =
      `${cfg.snBaseUrl}/api/now/table/sys_user` +
      `?sysparm_query=sys_id=${sysId}` +
      `&sysparm_fields=sys_id,department,division` +
      `&sysparm_display_value=true` +
      `&sysparm_limit=1`;

    Logger.log(`Fetching department & division for user sys_id: ${sysId}`);

    try {
      const resp = UrlFetchApp.fetch(userUrl, { method: "get", headers: { ...authH, Accept: "application/json" } });
      if (resp.getResponseCode() !== 200) {
        userCache[sysId] = { department: "", division: "" };
        return;
      }
      const result = JSON.parse(resp.getContentText()).result || [];
      if (result.length === 0) {
        userCache[sysId] = { department: "", division: "" };
        return;
      }
      const dept = (result[0].department && result[0].department.display_value) || "";
      const division = (result[0].division && result[0].division.display_value) || "";
      userCache[sysId] = { department: dept, division: division };
      Logger.log(
        `Cached for sys_id=${sysId}: Department="${dept}", Division="${division}"`
      );
    } catch (e) {
      Logger.log(`Error fetching department & division for ${sysId}: ${e}`);
      userCache[sysId] = { department: "", division: "" };
    }
  });

  const allDivisions = [...new Set(Object.values(userCache).map(u => u.division).filter(Boolean))];
  Logger.log(`Unique Divisions found: ${allDivisions.join(", ")}`);

  /***********************
   * CLEAR OUTPUT SHEET
   * Rows 1..DAILY_CHECK_PROTECTED_TOP_ROWS are user-managed and never touched.
   ***********************/
  const lastRowForClear = outputSheet.getLastRow();
  if (lastRowForClear >= DAILY_CHECK_DATA_START_ROW) {
    outputSheet
      .getRange(
        DAILY_CHECK_DATA_START_ROW,
        1,
        lastRowForClear - DAILY_CHECK_DATA_START_ROW + 1,
        outputSheet.getLastColumn()
      )
      .clearContent()
      .setBackground(null);
  }

  const maxC = outputSheet.getMaxColumns();
  const maxR = outputSheet.getMaxRows();
  // breakApart only the data area — any intentional merges in the protected
  // rows 1..DAILY_CHECK_PROTECTED_TOP_ROWS stay intact.
  if (maxR >= DAILY_CHECK_DATA_START_ROW) {
    outputSheet
      .getRange(
        DAILY_CHECK_DATA_START_ROW,
        1,
        maxR - DAILY_CHECK_DATA_START_ROW + 1,
        maxC
      )
      .breakApart();
  }
  // Force wrap=false on the header + data range so we don't override any
  // wrap=true the user set in rows 1..DAILY_CHECK_PROTECTED_TOP_ROWS.
  if (maxR >= DAILY_CHECK_HEADER_ROW) {
    outputSheet
      .getRange(
        DAILY_CHECK_HEADER_ROW,
        1,
        maxR - DAILY_CHECK_HEADER_ROW + 1,
        maxC
      )
      .setWrap(false);
  }

  /***********************
   * HEADERS
   ***********************/
  const headers = [
    "In Release Sheet",
    "Department",
    "CHG #",
    "Type",
    "Short Description",
    "Impact Description",
    "CR Approval-State",
    "Start Time (PST)",
    "End Time (PST)",
    "Business Justification",
    "Description",
    "Assignment Group",
    "Affected CI",
    "Affected Groups",
    "Requested By",
    "Requested For",
    "Created Time (PST)"
  ];

  outputSheet.getRange(DAILY_CHECK_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  outputSheet
    .getRange(DAILY_CHECK_HEADER_ROW, 1, 1, headers.length - 1)
    .setBackground("#1155CC")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold");
  outputSheet.setFrozenRows(DAILY_CHECK_HEADER_ROW);

  /***********************
   * PROCESS CRs
   ***********************/
  const before5 = [];
  const after5 = [];

  const targetDepartments = [
    "App Dev - Production",
    "App Dev - CET",
    "Communication Platforms"
  ];

  snowData.forEach(cr => {
    const snowCR = String(cr.number)
      .trim()
      .toUpperCase();
    const inRelease = releaseCRs.has(snowCR);

    const startRaw = unwrapSN(cr.start_date);
    const startDate = startRaw ? new Date(startRaw) : new Date(0);
    const isAfter5 = startDate.getHours() >= AFTER_HOUR;

    const state = unwrapSN(cr.state);
    if (state.toLowerCase() === "cancelled") return;

    const status = inRelease ? "✅" : "❌";

    const requestedForSysId = getSysIdFromRef(cr.u_requested_for);
    const userData = userCache[requestedForSysId] || {};
    const department = truncate(userData.department || "");

    const highlight =
      inRelease ||
      unwrapSN(cr.assignment_group).toUpperCase() === "IT RELEASE MANAGEMENT" ||
      targetDepartments.includes(department);

    const snowLink = `=HYPERLINK("${cfg.snBaseUrl}/nav_to.do?uri=change_request.do?sys_id=${cr.sys_id}","${snowCR}")`;

    Logger.log(
      `CR ${snowCR}: requested_for_sys_id=${requestedForSysId} → Department="${department}"`
    );

    const row = [
      status,
      department,
      snowLink,
      truncate(unwrapSN(cr.type)),
      truncate(unwrapSN(cr.short_description)),
      truncate(unwrapSN(cr.u_impact_description)),
      truncate(`${unwrapSN(cr.approval)} - ${state}`),
      toPST(startRaw),
      toPST(unwrapSN(cr.end_date)),
      truncate(unwrapSN(cr.u_business_justification)),
      truncate(unwrapSN(cr.description)),
      truncate(unwrapSN(cr.assignment_group)),
      truncate(unwrapSN(cr.cmdb_ci)),
      truncate(unwrapSN(cr.u_affected_groups)),
      truncate(unwrapSN(cr.requested_by)),
      truncate(unwrapSN(cr.u_requested_for)),
      toPST(unwrapSN(cr.sys_created_on))
    ];

    (isAfter5 ? after5 : before5).push({
      data: row,
      highlight,
      startDate
    });
  });

  before5.sort((a, b) => a.startDate - b.startDate);
  after5.sort((a, b) => a.startDate - b.startDate);

  /***********************
   * WRITE OUTPUT
   * Sheet.getRange(row, col, numRows, numColumns) — NOT (startRow, startCol, endRow, endCol).
   * One data row: getRange(rowNum, 1, 1, nCols) so the third param is numRows=1, not endRow=rowNum.
   ***********************/
  let rowNum = DAILY_CHECK_DATA_START_ROW;
  const writeRows = rows => {
    rows.forEach(r => {
      const nCols = r.data.length;
      const rowRange = outputSheet.getRange(rowNum, 1, 1, nCols);
      rowRange.setValues([r.data]);
      if (r.highlight) rowRange.setBackground("#fff2cc");
      rowNum++;
    });
  };

  writeRows(before5);

  if (after5.length) {
    // One row, merge B through last column (1 row, headers.length-1 wide starting at B)
    const mergeR = outputSheet.getRange(rowNum, 2, 1, headers.length - 1);
    mergeR
      .merge()
      .setValue("🌙 Changes with start date after 5 PM PST")
      .setBackground("#ddebf7")
      .setFontWeight("bold");
    rowNum++;
    writeRows(after5);
  }

  outputSheet.autoResizeColumns(1, headers.length);
  applyDailyCheckSheetRowHeights21_(outputSheet, DAILY_CHECK_HEADER_ROW);

  SpreadsheetApp.getUi().alert(
    "Daily Release Check complete: " +
      (before5.length + after5.length) +
      " CRs displayed" +
      (snowFriThroughSun ? " (Fri+Sat+Sun, Pacific). " : ". ") +
      "Row height: " +
      DAILY_RELEASE_CHECK_ROW_HEIGHT_PX +
      "px."
  );
}

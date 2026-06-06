/**
 * Multi-CR handling for ServiceNow Related Ticket (Jira customfield_10258).
 * Compares each CR's SNOW start_date calendar day to ReleaseSheet A1 in America/Los_Angeles.
 */

const LA_TZ = "America/Los_Angeles";

/**
 * Parses Jira "ServiceNow Related Ticket" text into unique CHG numbers.
 * @returns {{ crs: string[], multipleInField: boolean }}
 */
function parseRelatedTicketCRs(raw) {
  if (raw === null || raw === undefined) return { crs: [], multipleInField: false };
  const str = raw.toString().trim();
  if (!str) return { crs: [], multipleInField: false };

  const tokens = str.split(/[\s,;]+/).map(t => t.trim()).filter(Boolean);
  const nums = new Set();
  tokens.forEach(tok => {
    const digs = tok.match(/\d+/g);
    if (!digs) return;
    digs.forEach(d => {
      const n = Number(d);
      if (!isNaN(n) && n > 0) nums.add(n);
    });
  });

  if (nums.size === 0) return { crs: [], multipleInField: false };

  const crs = [...nums]
    .sort((a, b) => a - b)
    .map(n => `CHG${String(n).padStart(7, "0")}`);

  return { crs, multipleInField: crs.length > 1 };
}

/**
 * Calendar yyyy-MM-dd for the active release date in A1, interpreted in Pacific.
 */
function getReleaseCalendarDayPST() {
  const cell = SpreadsheetApp.getActive().getSheetByName(S_RELEASE).getRange("A1").getValue();
  if (!cell) throw new Error("Release date missing in A1.");

  if (cell instanceof Date) {
    return Utilities.formatDate(cell, LA_TZ, "yyyy-MM-dd");
  }

  const raw = String(cell).trim();
  const slash = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slash) {
    const parsed = Utilities.parseDate(raw, LA_TZ, "MM/dd/yyyy");
    return Utilities.formatDate(parsed, LA_TZ, "yyyy-MM-dd");
  }

  const ymd = raw.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (ymd) {
    const parsed = Utilities.parseDate(raw, LA_TZ, "yyyy-MM-dd");
    return Utilities.formatDate(parsed, LA_TZ, "yyyy-MM-dd");
  }

  const d = parseReleaseDateValue(cell);
  return Utilities.formatDate(d, LA_TZ, "yyyy-MM-dd");
}

/**
 * SNOW change_request.start_date as calendar yyyy-MM-dd in LA.
 */
function getSnowStartCalendarDayLA(record) {
  if (!record || !record.start_date) return "";
  const raw = (record.start_date.display_value || record.start_date.value || "").toString().trim();
  if (!raw) return "";

  const d = new Date(raw);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, LA_TZ, "yyyy-MM-dd");
  }

  const fmts = ["yyyy-MM-dd HH:mm:ss", "MM/dd/yyyy HH:mm:ss", "yyyy-MM-dd", "MM/dd/yyyy"];
  for (let i = 0; i < fmts.length; i++) {
    try {
      const p = Utilities.parseDate(raw, LA_TZ, fmts[i]);
      return Utilities.formatDate(p, LA_TZ, "yyyy-MM-dd");
    } catch (e) {
      // try next format
    }
  }
  return "";
}

/**
 * @param {string[]} allCrs
 * @param {Object} snowDataByCR map from fetchSNOWDataByCRNumbers
 * @param {string|string[]} releaseDayYmdOrWindow yyyy-MM-dd in LA, or an array of allowed days (Fri+Sat+Sun for Friday A1)
 * @returns {{ displayCrs: string[], highlightCrColumnBlue: boolean }}
 */
function classifyMultiCrDisplayRows(allCrs, snowDataByCR, releaseDayYmdOrWindow) {
  const allowList = Array.isArray(releaseDayYmdOrWindow)
    ? releaseDayYmdOrWindow
    : [releaseDayYmdOrWindow];
  const allow = new Set(allowList.filter(Boolean));

  const crDays = allCrs.map(cr => ({
    cr,
    day: getSnowStartCalendarDayLA(
      (snowDataByCR[cr] && snowDataByCR[cr].record) || null
    )
  }));
  const matching = crDays.filter(x => x.day && allow.has(x.day));

  if (!allCrs.length) {
    return { displayCrs: [""], highlightCrColumnBlue: false };
  }

  if (allCrs.length === 1) {
    const only = allCrs[0];
    const day = crDays[0].day;
    if (day && allow.has(day)) {
      return { displayCrs: [only], highlightCrColumnBlue: false };
    }
    return { displayCrs: [""], highlightCrColumnBlue: false };
  }

  if (matching.length === 0) {
    return { displayCrs: [""], highlightCrColumnBlue: true };
  }
  if (matching.length === 1) {
    return { displayCrs: [matching[0].cr], highlightCrColumnBlue: true };
  }
  return {
    displayCrs: matching.map(m => m.cr),
    highlightCrColumnBlue: true
  };
}

/**
 * One Jira issue -> one or more display rows (Rule 3 = duplicates).
 * @param {string|string[]} releaseDayYmdOrWindow same as classifyMultiCrDisplayRows
 */
function expandIssuesForDisplayRows(normalisedIssues, snowDataByCR, releaseDayYmdOrWindow) {
  const out = [];
  normalisedIssues.forEach(issue => {
    const cls = classifyMultiCrDisplayRows(issue.allCrs || [], snowDataByCR, releaseDayYmdOrWindow);
    cls.displayCrs.forEach(cr => {
      out.push(
        Object.assign({}, issue, {
          displayCr: cr,
          crCellHighlightBlue: cls.highlightCrColumnBlue
        })
      );
    });
  });
  return out;
}

function makeIssueCrRowKey(jiraItem, crCell) {
  return `${String(jiraItem || "").trim()}||${String(crCell || "").trim()}`;
}

function collectAllSnowCRNumbersFromIssues(issues) {
  const set = new Set();
  (issues || []).forEach(i => {
    (i.allCrs || []).forEach(cr => {
      if (cr) set.add(cr);
    });
  });
  return [...set];
}

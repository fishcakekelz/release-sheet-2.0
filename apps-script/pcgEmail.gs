/* =================== PCG / Non-Del Channel Email =================== */

/**
 * Channels that trigger this notification email.
 *
 * The matcher (channelTriggersRestrictedNotification_) normalizes the cell
 * value by uppercasing and stripping spaces and dashes, so the same logic
 * picks up all of these variants:
 *   PCG, pcg
 *   NonDel, Non-Del, nondel, non-del
 *   NonDel+, Non-Del+, nondel+, non-del+
 * It also matches when the cell contains multiple channels separated by
 * commas / pipes / etc., e.g. "Retail, PCG" or "Non-Del+ / PCG".
 */
const RESTRICTED_CHANNEL_TOKENS = ["PCG", "NONDEL"];

/**
 * Returns true when the Channel cell value mentions any restricted channel.
 *
 * We deliberately check substring (not exact match) so multi-channel cells
 * like "Retail, PCG" still trigger the notification. Stripping spaces and
 * dashes lets one rule cover NonDel / Non-Del / Non-Del+ / NonDel+.
 */
function channelTriggersRestrictedNotification_(channelValue) {
  if (channelValue === null || channelValue === undefined) return false;
  const norm = channelValue.toString().toUpperCase().replace(/[\s\-]/g, "");
  if (!norm) return false;
  return RESTRICTED_CHANNEL_TOKENS.some(t => norm.includes(t));
}

/**
 * Picks the Product Manager NAME for a given release-sheet row.
 * Strategy: Scrum Team match first (most specific) → Jira project key match
 * (project prefix from JIRA Item) → null if neither resolves.
 */
function resolvePmNameForRow_(row, cols) {
  const scrumTeam = cols.scrumTeam !== undefined
    ? (row[cols.scrumTeam] || "").toString().trim()
    : "";
  if (scrumTeam && SCRUM_TEAM_TO_PM[scrumTeam]) {
    return SCRUM_TEAM_TO_PM[scrumTeam];
  }

  const jiraItem = cols.jiraItem !== undefined
    ? (row[cols.jiraItem] || "").toString().trim()
    : "";
  if (jiraItem) {
    const m = jiraItem.match(/^([A-Z]+)-/);
    if (m && JIRA_PROJECT_TO_PM[m[1]]) {
      return JIRA_PROJECT_TO_PM[m[1]];
    }
  }

  return null;
}

/**
 * Converts a PM display name to a PNMAC email.
 *   1. PM_NAME_TO_EMAIL override map wins (use it for hyphens, shortened
 *      logins, and any name that doesn't follow the standard pattern).
 *   2. Auto-derive: "Firstname Lastname" → firstname.lastname@pnmac.com.
 *      Middle names are dropped; non-letter characters are stripped from
 *      each part. Returns null if we can't get at least two name parts.
 */
function pmEmailFromName_(name) {
  if (!name) return null;
  const trimmed = name.toString().trim();
  if (!trimmed) return null;

  if (typeof PM_NAME_TO_EMAIL !== "undefined" && PM_NAME_TO_EMAIL[trimmed]) {
    return PM_NAME_TO_EMAIL[trimmed];
  }

  const parts = trimmed
    .split(/\s+/)
    .map(p => p.toLowerCase().replace(/[^a-z]/g, ""))
    .filter(p => p.length > 0);
  if (parts.length < 2) return null;
  return parts[0] + "." + parts[parts.length - 1] + "@pnmac.com";
}

/**
 * Builds the final TO list: PCG_NOTIFY_EMAILS (always) plus the PM email
 * resolved per matching row. Returns:
 *   {
 *     recipients: string[],         // de-duped, lowercased emails
 *     resolved:  Array<{name,email,jiraItem}>,  // diagnostics
 *     unresolved: Array<{jiraItem, scrumTeam, projectKey}> // missed lookups
 *   }
 */
function buildRestrictedChannelRecipients_(rows, cols) {
  const recipients = new Set();
  (PCG_NOTIFY_EMAILS || []).forEach(e => recipients.add(e.toLowerCase()));

  const resolved = [];
  const unresolved = [];

  rows.forEach(row => {
    const jiraItem = cols.jiraItem !== undefined
      ? (row[cols.jiraItem] || "").toString().trim()
      : "";
    const scrumTeam = cols.scrumTeam !== undefined
      ? (row[cols.scrumTeam] || "").toString().trim()
      : "";

    const pmName = resolvePmNameForRow_(row, cols);
    if (!pmName) {
      const projectKey = jiraItem.match(/^([A-Z]+)-/);
      unresolved.push({
        jiraItem,
        scrumTeam,
        projectKey: projectKey ? projectKey[1] : ""
      });
      return;
    }

    const email = pmEmailFromName_(pmName);
    if (!email) {
      unresolved.push({ jiraItem, scrumTeam, pmName });
      return;
    }

    recipients.add(email.toLowerCase());
    resolved.push({ name: pmName, email, jiraItem });
  });

  return {
    recipients: Array.from(recipients),
    resolved,
    unresolved
  };
}

/**
 * TEST: Manually trigger the PCG/Non-Del channel email.
 */
function testPCGChannelEmail() {
  const cfg = getConfig();
  const sheet = SpreadsheetApp.getActive().getSheetByName(S_RELEASE);
  const releaseDate = getReleaseDateFromReleaseSheet();
  const count = checkPCGChannelAndNotify(sheet, cfg.sheetHeaderRow, releaseDate, cfg);
  Logger.log(`✅ Restricted-channel check complete - found ${count || 0} ticket(s)`);
  Logger.log(`📧 Email sent to: ${PCG_NOTIFY_EMAILS.join(", ")}`);
}

/**
 * Scans the release sheet for rows whose Channel matches one of the
 * restricted-channel tokens (PCG / NonDel / Non-Del / Non-Del+ / NonDel+)
 * and sends the nightly release notification email to PCG_NOTIFY_EMAILS.
 *
 * Behavior when nothing matches: no email is sent. A log line is written
 * ("No restricted-channel tickets found - email not sent") so you can see
 * the no-op in Apps Script Executions.
 *
 * Function name is preserved for backward compat with scheduledTriggers.gs.
 *
 * @param {Sheet}  sheet       - The release sheet
 * @param {number} headerRow   - Header row number
 * @param {string} releaseDate - Formatted release date string
 * @param {Object} cfg         - Configuration object
 * @returns {number} Number of matching rows found
 */
function checkPCGChannelAndNotify(sheet, headerRow, releaseDate, cfg) {
  const { map } = getHeaderIndexMapSafe(sheet, headerRow);
  const lastRow = sheet.getLastRow();

  if (lastRow <= headerRow) return 0;

  const channelCol = map["Channel"];
  if (channelCol === undefined) {
    Logger.log("Channel column not found - skipping restricted-channel check");
    return 0;
  }

  // Column indices for the email table (undefined columns are handled gracefully)
  const cols = {
    scrumTeam:        map["Scrum Team"],
    crNumber:         map["CR #"],
    approvals:        map["Approvals"],
    jiraItem:         map["JIRA Item"],
    description:      map["JIRA Description"],
    channel:          map["Channel"],
    darkDeployment:   map["Dark Deployment"],
    prodReleaseDate:  map["Jira Prod Release Date"],
    releaseTime:      map["Start Time"],
    affectedGroups:   map["Affected Groups"],
    businessImpact:   map["Impact Description"],
    deploymentStatus: map["Deployment Status"]
  };

  const dataRange = sheet.getRange(headerRow + 1, 1, lastRow - headerRow, sheet.getLastColumn());
  const data = dataRange.getValues();

  const matchedRows = [];
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const jiraItem = cols.jiraItem !== undefined ? (row[cols.jiraItem] || "").toString().trim() : "";
    if (!jiraItem) continue;

    if (channelTriggersRestrictedNotification_(row[channelCol])) {
      matchedRows.push(row);
    }
  }

  if (matchedRows.length > 0) {
    sendPCGChannelEmail(matchedRows, cols, releaseDate, cfg);
  } else {
    Logger.log("No restricted-channel tickets found - email not sent");
  }

  Logger.log(`Restricted-channel check: found ${matchedRows.length} ticket(s)`);
  return matchedRows.length;
}

/**
 * Builds and sends the formatted PCG/Non-Del notification email.
 *
 * Recipients live in PCG_NOTIFY_EMAILS in constants.gs (today: Manan and
 * Kelly). Edit that constant to change who receives this email.
 *
 * Function name is preserved for backward compat with scheduledTriggers.gs.
 *
 * @param {Array}  rows        - Array of row value arrays from the release sheet
 * @param {Object} cols        - Map of column key → column index
 * @param {string} releaseDate - Formatted release date string
 * @param {Object} cfg         - Configuration object (used for jiraBaseUrl)
 */
function sendPCGChannelEmail(rows, cols, releaseDate, cfg) {
  // Build the dynamic recipient list: PCG_NOTIFY_EMAILS + each ticket's PM.
  const built = buildRestrictedChannelRecipients_(rows, cols);
  const recipients = built.recipients;

  if (!recipients.length) {
    Logger.log("Restricted-channel email skipped: empty recipients list.");
    return;
  }

  Logger.log(`Restricted-channel recipients (${recipients.length}): ${recipients.join(", ")}`);
  if (built.resolved.length) {
    const summary = built.resolved
      .map(r => `${r.jiraItem} → ${r.name} <${r.email}>`)
      .join("\n  ");
    Logger.log(`PMs resolved per ticket:\n  ${summary}`);
  }
  if (built.unresolved.length) {
    Logger.log(
      "Could NOT resolve a PM for these rows (still emailed to PCG_NOTIFY_EMAILS):\n  " +
        built.unresolved
          .map(u => `${u.jiraItem || "?"}  team=${u.scrumTeam || "?"}  project=${u.projectKey || "?"}  pmName=${u.pmName || "?"}`)
          .join("\n  ")
    );
  }

  const ss = SpreadsheetApp.getActive();
  const sheetUrl = ss.getUrl();
  const greeting = recipients.length > 0
    ? recipients.map(email => {
        const name = email.split("@")[0].split(".")[0];
        return name.charAt(0).toUpperCase() + name.slice(1);
      }).join(", ")
    : "Team";

  // ---- Build HTML table ----
  const tableHeaders = [
    "Scrum Team",
    "CR #",
    "Approvals",
    "JIRA Item",
    "JIRA Description",
    "JIRA Channel",
    "Dark Deployment",
    "JIRA Prod Release Date",
    "Release Time (PST)",
    "ServiceNow Affected Groups",
    "Business Impact per ServiceNow CR",
    "Deployment Status"
  ];

  const colKeys = [
    "scrumTeam",
    "crNumber",
    "approvals",
    "jiraItem",
    "description",
    "channel",
    "darkDeployment",
    "prodReleaseDate",
    "releaseTime",
    "affectedGroups",
    "businessImpact",
    "deploymentStatus"
  ];

  const thStyle = 'style="background-color:#003366;color:#ffffff;padding:8px 10px;text-align:left;white-space:nowrap;"';
  const tableStyle = 'border="1" cellpadding="0" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px;width:100%;"';

  let tableHtml = `<table ${tableStyle}><thead><tr>`;
  tableHeaders.forEach(h => { tableHtml += `<th ${thStyle}>${h}</th>`; });
  tableHtml += `</tr></thead><tbody>`;

  rows.forEach((row, idx) => {
    const rowBg = idx % 2 === 0 ? "#ffffff" : "#f2f2f2";
    tableHtml += `<tr style="background-color:${rowBg};">`;
    colKeys.forEach(key => {
      const colIdx = cols[key];
      const val = colIdx !== undefined ? (row[colIdx] !== null && row[colIdx] !== undefined ? row[colIdx].toString() : "") : "";
      if (key === "jiraItem" && val && cfg && cfg.jiraBaseUrl) {
        tableHtml += `<td style="padding:6px 10px;"><a href="${cfg.jiraBaseUrl}/browse/${val}" style="color:#0066cc;">${val}</a></td>`;
      } else {
        tableHtml += `<td style="padding:6px 10px;">${val}</td>`;
      }
    });
    tableHtml += `</tr>`;
  });

  tableHtml += `</tbody></table>`;

  // ---- HTML body ----
  const htmlBody =
    `<p style="font-family:Arial,sans-serif;">Dear ${greeting},</p>` +
    `<p style="font-family:Arial,sans-serif;">The following JIRA tickets, marked with PCG, Non-Del, or Non-Del+ as an impacted channel, are scheduled for release tonight and will be discussed during today's release review meeting:</p>` +
    `<p style="font-family:Arial,sans-serif;">The deployment is scheduled to begin at 5:00 and 7:00 PM PST. The Product team anticipates no downtime or business user impact during this deployment.</p>` +
    `<p style="font-family:Arial,sans-serif;">Please use this link <a href="${sheetUrl}">Release Sheet</a> for the complete release content:</p>` +
    `<p style="font-family:Arial,sans-serif;">Kindly notify the ITRM team immediately if you foresee any issues or concerns.</p>` +
    `<br>${tableHtml}`;

  // ---- Plain text fallback ----
  const plainBody =
    `Dear ${greeting},\n\n` +
    `The following JIRA tickets, marked with PCG, Non-Del, or Non-Del+ as an impacted channel, are scheduled for release tonight and will be discussed during today's release review meeting:\n\n` +
    `The deployment is scheduled to begin at 5:00 and 7:00 PM PST. The Product team anticipates no downtime or business user impact during this deployment.\n\n` +
    `Please use this link for the complete release content: ${sheetUrl}\n\n` +
    `Kindly notify the ITRM team immediately if you foresee any issues or concerns.\n\n` +
    rows.map((row, i) => {
      const get = key => {
        const idx = cols[key];
        return idx !== undefined ? (row[idx] || "").toString() : "";
      };
      return `${i + 1}. ${get("jiraItem")} | ${get("scrumTeam")} | Channel: ${get("channel")} | ${get("description")}`;
    }).join("\n");

  try {
    MailApp.sendEmail({
      to: recipients.join(","),
      subject: `PCG / Non-Del Release Notification - ${releaseDate}`,
      body: plainBody,
      htmlBody: htmlBody
    });
    Logger.log(`Restricted-channel email sent to: ${recipients.join(", ")}`);
  } catch (e) {
    Logger.log(`Error sending restricted-channel email: ${e.message}`);
  }
}

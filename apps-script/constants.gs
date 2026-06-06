/* =================== Configuration Constants =================== */

/* =================== Field Map =================== */
/**
 * Values must be Jira REST field ids (customfield_##### or system names like "description").
 * Do NOT use script-style names here (e.g. "qaArtifactsLink") — those break /search `fields`.
 * "cf[#####]" is accepted in code and normalized to customfield_##### in Release-data.gs.
 */
const FIELD_MAP = {
  scrumTeam: "customfield_10254", // Teams field (multi-select)
  workstream: "customfield_10209",
  prodReleaseDate: "customfield_10220",
  crNumber: "customfield_10258", // CR #
  approvals: "customfield_10330",
  channel: "customfield_10253",
  acceptanceCriteria: "customfield_10052",
  linkedIssues: "issuelinks",
  status: "status",
  resolution: "resolution",
  assignee: "assignee",
  description: "description",
  lastComment: "customfield_10260",
  qaArtifactsLink: "customfield_10181",
  uatArtifactsLink: "customfield_10182",
  /** Use customfield_10185 — not cf[10185] (normalized in code, but prefer API id here). */
  severity: "customfield_10185"
};

/* =================== CAP Project Exclusion ===================
 * fetchJiraIssuesByJql() drops any CAP ticket whose Jira Components or CMDB CI
 * field contains one of CAP_EXCLUDED_TOKENS as a whole token (case-insensitive,
 * separator-aware so "PROD-ATHENA-01" does NOT match but "Athena Service" does).
 * Edit the tokens here to add/remove exclusions.
 */
const CAP_EXCLUDE_PROJECT_KEY = "CAP";
const CAP_EXCLUDED_TOKENS = ["Luna", "Athena"];
const CAP_CMDB_CI_FIELD_ID = "customfield_10177";

/* =================== Sheet names =================== */
const S_CONFIG = "Config";
const S_RELEASE = "ReleaseSheet";
const S_MASTER = "MasterPull";
const S_CHANGELOG = "ChangeLog";

/** Default pixel height for each row in ReleaseSheet (Format → Row height) after Create/Refresh. Config key releaseDataRowHeight overrides. */
const RELEASE_SHEET_DATA_ROW_HEIGHT_PX = 35;

/**
 * Columns whose existing cell content (typically VLOOKUP formulas / manual
 * edits) is preserved across the corresponding flow. Each list is matched
 * by header NAME, so reordering columns won't silently break protection.
 *
 * - REFRESH_PROTECTED_COLUMN_NAMES: refreshReleaseSheetInternal_ snapshots
 *   these columns before its clear, then restores them after writing data +
 *   the appended REMOVED rows.
 * - CREATE_PROTECTED_COLUMN_NAMES: createReleaseSheetInternal_ does the
 *   same around its from-scratch clear + write.
 *
 * Add/remove names to change protection. Names must match
 * STOP_RELEASE_SHEET_HEADER_ROW entries exactly.
 */
const REFRESH_PROTECTED_COLUMN_NAMES = [
  "Date",
  "Pre-Release Check Summary",
  "Division"
];
const CREATE_PROTECTED_COLUMN_NAMES = [
  "Pre-Release Check Summary",
  "Division"
];

/**
 * Canonical ReleaseSheet column headers (left → right) applied to the new tab when StopRelease runs
 * (after TEMPLATE is copied to ReleaseSheet). Single row, no duplicate names.
 * Config sheetHeaderRow is which row this is (usually 1).
 */
const STOP_RELEASE_SHEET_HEADER_ROW = [
  "Date",
  "Pre-Release Check Summary",
  "Division",
  "Workstream",
  "Scrum Team",
  "JIRA Item",
  "Jira Status",
  "Jira Resolution",
  "CR #",
  "Start Time",
  "Release Type",
  "Validation Owner",
  "Approvals",
  "Dark Deployment",
  "CHG Type",
  "Assignment Group",
  "Change Assignee",
  "CTASK List",
  "Deployment Status",
  "Validation Status",
  "Comments",
  "Late Addition",
  "Late Approval",
  "Acceptance Criteria",
  "Last Updated JIRA comment",
  "Channel",
  "Jira Prod Release Date",
  "Issue Type",
  "Severity",
  "JIRA Assignee",
  "QA Artifacts Link",
  "UAT Artifacts Link",
  "Related Issues",
  "Blocker Status",
  "Impact Description",
  "Change Plan",
  "Rollback Plan",
  "JIRA Summary",
  "JIRA Description",
  "Validation Plan",
  "Impacted CI",
  "Date Submitted",
  "Affected Groups",
  "Affected Locations",
  "Tech Lead / Sr Tech Lead",
  "Deploy Start",
  "Deploy End",
  "Validation Start",
  "Validation End",
  "Deploy Duration",
  "Release Duration",
  "JIRA Approval Comment Sent"
];

/* =================== ServiceNow (optional; Release-data merges these) =================== */
const SNOW_CHANGE_REQUEST_EXTRA_FIELDS = ["department", "u_division", "company"];
const SNOW_DIVISION_SOURCE_FIELDS = ["u_division", "company"];
const SNOW_DEPARTMENT_SOURCE_FIELDS = ["department"];
const SNOW_IMPACT_DESCRIPTION_SOURCE_FIELDS = [
  "u_impact_description",
  "impact_description",
  "description"
];

/* =================== Color constants =================== */
const COLOR_MODIFIED = "#ffff00"; // Yellow for modified tickets
const COLOR_REMOVED = "#ffcccc"; // Light red for removed tickets
const COLOR_NEW = "#ccffcc"; // Light green for new tickets
const COLOR_MISSING_APPROVAL = "#ffcc99"; // Orange for missing approvals
const COLOR_MULTIPLE_CR = "#cce5ff"; // Light blue for multiple CRs

/* =================== Scrum Team to Product Manager Mapping =================== */
const SCRUM_TEAM_TO_PM = {
  "MDM": "Matthew Baier",
  "Lead Data": "Matthew Baier",
  "CDP": "Matthew Baier",

  "CRM": "Jacob Goldstein",
  "Salesforce Marketing Cloud": "Jacob Goldstein",
  "Production Unified Comms": "Jacob Goldstein",

  "Customers Falcon": "Terry DuVarney",
  "Customers Raptor": "Terry DuVarney",
  "Dragon": "Terry DuVarney",
  "Kestrel": "Terry DuVarney",
  "Kestrel (fka: Marketing Sites)": "Terry DuVarney",
  "Marketing Sites": "Terry DuVarney",

  "MyHome": "Mayank Aggarwal",
  "COMM Analytics": "Mayank Aggarwal",

  "CE - Portals": "Niranjana Thiruvengadam",
  "TPIA": "Niranjana Thiruvengadam",
  "CDL-FF-Trendsetters": "Niranjana Thiruvengadam",
  "CDL-FF-Spartans": "Niranjana Thiruvengadam",

  "Core Data Engineering": "David Williams",
  "Core Services": "David Williams",
  "Pricing": "David Williams",
  "Document Nexus": "David Williams",

  "Fulfillment PCG": "Frank Moss",
  "Fulfillment TPO": "Frank Moss",
  "Fulfillment Engineering": "Frank Moss",
  "PCG/TPO Support (L2)": "Frank Moss",
  "DocGen": "Frank Moss",
  "E2E - Test Automation": "Frank Moss",

  "Escrow Administration": "Paula Moughton",
  "Core Specialty and HELOC": "Paula Moughton",
  "Loan Boarding & Servicing Transfers": "Paula Moughton",

  "Investor Reporting & Remittance": "Deanne Radonic",
  "Investor Allocations & Reconciliation": "Deanne Radonic",
  "Cash Management": "Deanne Radonic",
  "ACH Processing & PCA Recon": "Deanne Radonic",
  "Fees & Disbursements": "Deanne Radonic",
  "Default Reporting": "Deanne Radonic",

  "Modifications": "Raychel Cooksey",
  "SPOC & Asset Resolution": "Raychel Cooksey",
  "AP & Foreclosure &  Bankruptcy": "Raychel Cooksey",
  "Claims & Final Resolution": "Raychel Cooksey",
  "Collections & Campaign Builder": "Raychel Cooksey",
  "Loss Mitigation Decision Engine": "Raychel Cooksey",

  "Customer Communications": "Katha Kanwar",
  "Subservicing": "Katha Kanwar",
  "LAMP": "Katha Kanwar",
  "LAFA": "Katha Kanwar",
  "Core Components": "Katha Kanwar",

  "BRMS": "Darin Suh",
  "WorkQueue 2.0": "Darin Suh",
  "Digital Systems": "Darin Suh",
  "MacChat 2.0": "Darin Suh",
  "Data": "Darin Suh",
  "Foundation": "Darin Suh",

  "Servicing Production Support (L2)": "Brian Anderson",
  "Servicing Solution Design": "Brian Anderson",

  "Capital Markets - Gemini": "Srinivas Kanury",
  "Capital Markets - MMS / Ratesheets": "Srinivas Kanury",
  "Capital Markets - Vanna": "Srinivas Kanury",
  "Capital Markets - Mercury": "Srinivas Kanury",

  "Rims": "Irving Lopez",
  "DataHive": "Irving Lopez",
  "Tableau": "Irving Lopez",
  "Snowflake": "Irving Lopez",

  "Monthly patching": "Swapna Nadkarni",
  "PDM": "Ammar Alipo-on",
  "Telephony": "Gennadiy Karasev",
  "AI Accelerator": "Arthe Sampath",
  "Moda": "Tina Tararache",
  "MODA": "Tina Tararache",
  "PCH Hubs Reporting team": "Taylor Silvey"
};

/* =================== Jira Project Key to Product Manager Mapping =================== */
const JIRA_PROJECT_TO_PM = {
  LEADG: "Matthew Baier",
  CDM: "Jacob Goldstein",
  COM: "Mayank Aggarwal",
  DCX: "Terry DuVarney",
  POP: "Mayank Aggarwal",
  CE: "Niranjana Thiruvengadam",
  KCP: "David Williams",
  KBF: "Frank Moss",
  DLL: "Niranjana Thiruvengadam",
  SCS: "Paula Moughton",
  SFS: "Deanne Radonic",
  SDS: "Raychel Cooksey",
  SEF: "Katha Kanwar",
  SES: "Darin Suh",
  SSH: "Brian Anderson",
  CAP: "Srinivas Kanury",
  OM: "Swapna Nadkarni",
  EDMA: "Irving Lopez",
  TEL: "Gennadiy Karasev",
  AITECH: "Arthe Sampath",
  PF: "Tina Tararache"
};

/* =================== Email Configuration =================== */
const MISSING_APPROVAL_NOTIFY_EMAILS = ["manan.patel@pnmac.com", "kelly.mok@pnmac.com"];

const REFRESH_NOTIFY_EMAILS = ["manan.patel@pnmac.com", "kelly.mok@pnmac.com"];

const PCG_NOTIFY_EMAILS = ["manan.patel@pnmac.com", "kelly.mok@pnmac.com"];

/**
 * Override map for PM Name → email. The PCG / Non-Del email auto-derives
 * emails as firstname.lastname@pnmac.com from SCRUM_TEAM_TO_PM /
 * JIRA_PROJECT_TO_PM, which works for most PNMAC accounts but is wrong for
 * names with hyphens, suffixes, shortened logins, or non-standard formats.
 * Add explicit entries here for any PM whose real email differs from the
 * auto-derived guess. Keys must match the name string exactly as it appears
 * in SCRUM_TEAM_TO_PM / JIRA_PROJECT_TO_PM.
 *
 * Example:
 *   "Ammar Alipo-on": "ammar.alipo-on@pnmac.com",
 */
const PM_NAME_TO_EMAIL = {
  // Add overrides here.
};

/* =================== Protected Columns Configuration =================== */
const PROTECTED_COLUMNS = [
  "Approvals",
  "Dark Deployment",
  "Deployment Status",
  "Validation Status",
  "Validation Owner",
  "Late Addition",
  "Late Approval",
  "JIRA Approval Comment Sent"
];

/* =================== Jira Approval Comment Configuration =================== */
/**
 * The sheet column that holds the approval indicator. When this cell is edited
 * to a value matching APPROVAL_INDICATOR_REGEX, a comment is posted to Jira.
 */
const APPROVAL_INDICATOR_COLUMN = "Approvals";

/**
 * The sheet column where the script writes the timestamp after posting a Jira
 * comment so the same row is not processed twice.
 */
const APPROVAL_COMMENT_SENT_COLUMN = "JIRA Approval Comment Sent";

/**
 * Any non-empty value in the Approvals column counts as approval. The PM /
 * Sr Tech Lead types their name in this cell to approve, and that name is
 * used as the approver in the Jira comment. Whitespace-only cells are ignored.
 */
const APPROVAL_INDICATOR_REGEX = /\S/;

const PROTECTED_COLUMN_EDITORS = [
  "kelly.mok@pnmac.com",
  "manan.patel@pnmac.com",
  "mone.chen@pnmac.com",
  "swapna.nadkarni@pnmac.com",
  "mujahi.harper@pnmac.com",
  "jacob.railsback@pnmac.com",
  "alex.tirado@pnmac.com",
  "thirumaran.kirubanandam@pnmac.com",
  "v-rohan.mahajan@pnmac.com"
];

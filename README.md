# Release Sheet 2.0

Tooling that drives the ITRM weekly **Release Sheet** in Google Sheets, plus a small Python
package for offline workflows. The Apps Script side pulls release candidates from **Jira**
and corresponding change tickets from **ServiceNow** into a single, color-coded sheet that
PMs use to track each Friday/weekend production release.

> Repo: <https://github.com/fishcakekelz/release-sheet-2.0>

---

## What it does

- **Builds & refreshes** the `ReleaseSheet` tab for a given release date (defaults to the
  Friday + the following Sat/Sun in `America/Los_Angeles`).
- **Pulls Jira issues** via JQL (built from the date in `ReleaseSheet!A1`) and joins them
  to **ServiceNow CRs / CTASKs**.
- **Tracks changes** in a `ChangeLog` tab — new, modified, and removed rows are highlighted
  on refresh; PM-edited columns are preserved.
- **Posts back to Jira** automatically when an approval is recorded in the sheet
  (`JiraApprovalComment.gs`).
- **Sends notifications** for missing approvals and PCG / Non-Del channel releases
  (`pcgEmail.gs`).
- **Runs on a schedule** — 9 AM / 12 PM / 2 PM PT refresh + checks via installable
  time-based triggers (`scheduledTriggers.gs`).
- **Has a manual fallback** (`BackupRunbook.gs`) that generates `TEMPLATE`, `BACKUP BUILD`,
  and `BACKUP PROCEDURE` tabs so a PM can produce a Release Sheet from Jira/SNOW CSVs in
  the browser if Apps Script is unavailable.
- **Archives** completed releases to a `Summary` tab.

---

## Repository layout

```
.
├── apps-script/                Google Apps Script project (deployed via clasp)
│   ├── appsscript.json         Runtime manifest (V8, America/Los_Angeles)
│   ├── constants.gs            FIELD_MAP, sheet names, headers, protected columns
│   ├── Release-data.gs         Core Create / Refresh / Stop flow + ITRM Functions menu
│   ├── DailyReleaseCheck.gs    Daily SNOW-vs-Sheet reconciliation tab
│   ├── ReleaseSheetBackup.gs   Cheap, formatting-free backup of the live sheet
│   ├── RevisedReleaseReport.gs Pre-Release Check Report generation
│   ├── JiraApprovalComment.gs  On-edit trigger → posts approval comment to Jira
│   ├── StatusFormatting.gs     Deployment / Validation Status dropdowns + row coloring
│   ├── BackupRunbook.gs        Manual fallback tabs (TEMPLATE / BACKUP BUILD / PROCEDURE)
│   ├── Summary.gs              Archive Release Sheet → Summary
│   ├── multiCrRelease.gs       Multi-CHG handling for the SNOW Related Ticket field
│   ├── pcgEmail.gs             PCG / Non-Del channel notification email
│   └── scheduledTriggers.gs    Headless refresh + daily / test trigger management
├── src/releasesheet2/          Python package (scaffolding, CLI stub)
│   ├── __init__.py
│   └── cli.py                  `releasesheet2` entry point
├── pyproject.toml              Python package metadata (depends on openpyxl)
├── .clasp.json.example         Template for the clasp config (copy → .clasp.json)
├── .claspignore                Files clasp should not push to Apps Script
└── .gitignore
```

---

## The `ITRM Functions` menu

Opening the spreadsheet runs `onOpen()` in `Release-data.gs`, which builds this menu:

| Item | What it does |
| --- | --- |
| Manually Create Release Sheet | Builds the sheet from scratch for the date in `A1`. |
| Refresh Release Sheet | Re-pulls Jira + SNOW, highlights new/modified/removed rows. |
| CLOSE SHEET / Stop Release | Copies `TEMPLATE` onto `ReleaseSheet` and resets it. |
| SNowPull / CTASKPull / Pull CR + CTASK Info | Manual ServiceNow pulls. |
| Archive to Summary | Append the current sheet's rows to the `Summary` tab. |
| Archive & Clear Release Sheet | Archive then stop. |
| Daily Release Check | Reconciles today's SNOW CRs against the sheet. |
| Send PCG Channel Email | Notifies PMs of PCG / Non-Del channel work. |

Additional submenus (Approval Comments, Scheduled Triggers, Backup, etc.) are added by the
other `.gs` files.

---

## Setup

### 1. Apps Script (primary)

This project uses [clasp](https://github.com/google/clasp) to push the `.gs` files to a
Google Apps Script project bound to the Release Sheet spreadsheet.

```bash
npm install -g @google/clasp
clasp login

cp .clasp.json.example .clasp.json
# edit .clasp.json and paste the scriptId of your Apps Script project
clasp push
```

`.claspignore` keeps `.git/`, `.venv/`, `src/`, and Python artifacts out of the Apps Script
project. Only the contents of `apps-script/` are deployed (per `rootDir`).

### 2. Spreadsheet configuration

The script reads a `Config` tab for runtime values (Jira base URL + credentials, header
row, SNOW saved-search ids, row heights, etc.). At a minimum you'll need to populate the
Jira connectivity rows so `fetchJiraIssuesByJql()` can authenticate.

Jira custom field IDs live in `FIELD_MAP` in `apps-script/constants.gs` — update them if
your Jira instance uses different `customfield_#####` ids.

### 3. Scheduled refresh (optional)

From the script editor (or via the ITRM menu wiring), run `createDailyTriggers()` once to
install time-based triggers at 9 AM / 12 PM / 2 PM PT. Use `removeDailyTriggers()` to
uninstall. `createTestTrigger()` runs every 5 minutes for development.

### 4. Approval-comment trigger (optional)

Run `installApprovalEditTrigger()` once per spreadsheet (as an editor with `UrlFetchApp`
auth) to enable automatic Jira comments when an approval is recorded in the sheet.

---

## Python package

A small `releasesheet2` package lives in `src/`. Today it's a stub with a CLI entry point;
intended for offline / batch work (e.g. parsing exported `.xlsx` backups via `openpyxl`).

```bash
python -m venv .venv
source .venv/bin/activate
pip install -e .

releasesheet2 --version
```

---

## Development notes

- **Timezone:** everything date-related runs in `America/Los_Angeles`. The release window
  is "Friday + the following Sat/Sun" — see `expandFridayWeekendJqlYmdsForY0_` in
  `Release-data.gs`.
- **Refresh safety:** columns in `REFRESH_PROTECTED_COLUMN_NAMES` /
  `CREATE_PROTECTED_COLUMN_NAMES` (`constants.gs`) are snapshotted before clears and
  restored after writes so manual edits / VLOOKUPs survive a refresh.
- **CAP project exclusion:** Jira tickets in project `CAP` are dropped when their
  Components or CMDB CI contain `CAP_EXCLUDED_TOKENS` (`Luna`, `Athena`) as whole tokens.
- **Headers are looked up by name**, never by column letter — reordering columns in the
  sheet won't silently break code, but renaming a header will.
- **Do not commit `.clasp.json`** (it contains the bound `scriptId`). The example template
  is checked in instead.

---

## License

_TBD — add a `LICENSE` file if you plan to share this externally._

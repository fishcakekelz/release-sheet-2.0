/**
 * Headless version of refreshReleaseSheet for scheduled triggers.
 * Runs without UI prompts and sends email notifications.
 */
function refreshReleaseSheetScheduled() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(S_RELEASE);
    if (!sheet) throw new Error("ReleaseSheet tab missing.");

    const dateObj = getReleaseDateFromReleaseSheet();
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    validateDateChange(dateObj, true);

    const cfg = getConfig();
    const headerRow = cfg.sheetHeaderRow;

    const r = refreshReleaseSheetInternal_(true);
    if (r.cancelled) return;

    if (r.emptySheet) {
      let missingApprovalCount = 0;
      const currentHour = new Date().getHours();
      if (currentHour === 12) {
        missingApprovalCount = checkMissingApprovalsAndNotify(sheet, headerRow, dateObj, cfg);
      }
      if (currentHour === 14) {
        checkPCGChannelAndNotify(sheet, headerRow, dateObj, cfg);
      }
      setLockedReleaseDate(dateObj);
      Logger.log(`Initial scheduled population: ${r.newCount} rows for ${dateObj}`);
      sendRefreshEmail(dateObj, now, r.newCount, 0, 0, missingApprovalCount);
      return;
    }

    let missingApprovalCount = 0;
    const currentHour = new Date().getHours();
    if (currentHour === 12 || currentHour === 14) {
      missingApprovalCount = checkMissingApprovalsAndNotify(sheet, headerRow, r.dateObj, cfg);
      Logger.log(
        `${currentHour === 12 ? "12 PM" : "2 PM"} check: Found ${missingApprovalCount} tickets missing approval`
      );
    }

    if (currentHour === 14) {
      checkPCGChannelAndNotify(sheet, headerRow, r.dateObj, cfg);
      Logger.log("2 PM: PCG channel check complete");
    }

    Logger.log(
      `Scheduled refresh completed. New: ${r.newCount}, Modified: ${r.modifiedCount}, Removed: ${r.removedCount}`
    );
    sendRefreshEmail(dateObj, now, r.newCount, r.modifiedCount, r.removedCount, missingApprovalCount);
  } catch (error) {
    Logger.log(`Scheduled refresh FAILED: ${error.message}`);
    const email = Session.getActiveUser().getEmail();
    if (email) {
      MailApp.sendEmail({
        to: email,
        subject: "❌ Release Sheet Refresh FAILED",
        body: `Scheduled refresh failed at ${new Date().toISOString()}.\n\n` +
              `Error: ${error.message}\n\n` +
              `Please check the script logs for more details.`
      });
    }
  }
}
/* =================== Scheduled Triggers =================== */

/**
 * TEST: Creates a trigger that runs every 5 minutes for testing.
 */
function createTestTrigger() {
  deleteTrigger("refreshReleaseSheetScheduled");
  
  ScriptApp.newTrigger("refreshReleaseSheetScheduled")
    .timeBased()
    .everyMinutes(5)
    .create();
  
  Logger.log("TEST trigger created - runs every 5 minutes!");
  SpreadsheetApp.getUi().alert("TEST trigger created!\n\nThe release sheet will refresh every 5 minutes.\n\n⚠️ Remember to run 'Stop Test' when done!");
}

/**
 * TEST: Removes the test trigger
 */
function removeTestTrigger() {
  deleteTrigger("refreshReleaseSheetScheduled");
  Logger.log("Test trigger removed.");
  SpreadsheetApp.getUi().alert("Test trigger removed.");
}

/**
 * Creates triggers to refresh the release sheet daily at 9 AM, 12 PM, and 2 PM PST.
 */
function createDailyTriggers() {
  removeDailyTriggers();
  
  // 9 AM PST
  ScriptApp.newTrigger("refreshReleaseSheetScheduled")
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
  
  // 12 PM PST
  ScriptApp.newTrigger("refreshReleaseSheetScheduled")
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .create();
  
  // 2 PM PST
  ScriptApp.newTrigger("refreshReleaseSheetScheduled")
    .timeBased()
    .everyDays(1)
    .atHour(14)
    .create();
  
  Logger.log("Daily triggers created for 9 AM, 12 PM, and 2 PM PST!");
  SpreadsheetApp.getUi().alert("Daily triggers created!\n\nThe release sheet will refresh automatically at:\n• 9 AM PST\n• 12 PM PST\n• 2 PM PST");
}

/**
 * Deletes a trigger by function name.
 */
function deleteTrigger(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
      Logger.log(`Deleted trigger for: ${functionName}`);
    }
  });
}

/**
 * Remove all daily triggers
 */
function removeDailyTriggers() {
  deleteTrigger("refreshReleaseSheetScheduled");
  Logger.log("All daily triggers removed.");
}

/**
 * Remove all daily triggers with UI confirmation
 */
function removeDailyTriggersWithAlert() {
  removeDailyTriggers();
  SpreadsheetApp.getUi().alert("All daily triggers have been removed.");
}
</user_query>
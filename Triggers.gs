// File: Triggers.gs
// Description: Contains functions for creating, verifying, and managing
// time-driven triggers for the project.

// --- Triggers ---
function createTimeDrivenTrigger(functionName = 'processJobApplicationEmails', hours = 1) {
  let exists = false;
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === functionName && t.getEventType() === ScriptApp.EventType.CLOCK) {
        exists = true;
      }
    });
    if (!exists) {
      ScriptApp.newTrigger(functionName).timeBased().everyHours(hours).create();
      Logger.log(`[INFO] TRIGGER: ${hours}-hourly trigger for "${functionName}" created successfully.`);
    } else {
      Logger.log(`[INFO] TRIGGER: ${hours}-hourly trigger for "${functionName}" already exists.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] TRIGGER: Failed to create or verify ${hours}-hourly trigger for "${functionName}": ${e.message} (Stack: ${e.stack})`);
    // Optionally, inform user via UI if context allows and it's critical, but for setup, log is primary
  }
  return !exists; // Returns true if a new trigger was created in this call
}

function createOrVerifyStaleRejectTrigger(functionName = 'markStaleApplicationsAsRejected', hour = 2) { // Default to 2 AM
  let exists = false;
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === functionName && t.getEventType() === ScriptApp.EventType.CLOCK) {
        exists = true;
      }
    });
    if (!exists) {
      ScriptApp.newTrigger(functionName).timeBased().everyDays(1).atHour(hour).inTimezone(Session.getScriptTimeZone()).create();
      Logger.log(`[INFO] TRIGGER: Daily trigger for "${functionName}" (around ${hour}:00 script timezone) created successfully.`);
    } else {
      Logger.log(`[INFO] TRIGGER: Daily trigger for "${functionName}" already exists.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] TRIGGER: Failed to create or verify daily trigger for "${functionName}": ${e.message} (Stack: ${e.stack})`);
  }
  return !exists;
}

// File: GmailUtils.gs
// Description: Contains functions for interacting with Gmail, primarily label creation and management.

// --- Helper: Get or Create Gmail Label ---
function getOrCreateLabel(labelName) {
  if (!labelName || typeof labelName !== 'string' || labelName.trim() === "") {
    Logger.log(`[GMAIL_UTIL ERROR] Invalid labelName provided to getOrCreateLabel: "${labelName}"`);
    return null;
  }
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(labelName);
  } catch (e) {
    Logger.log(`[GMAIL_UTIL ERROR] Error checking for label "${labelName}": ${e.message}`);
    return null;
  }

  if (!label) {
    if (DEBUG_MODE) Logger.log(`[GMAIL_UTIL DEBUG] Label "${labelName}" not found. Creating...`);
    try {
      label = GmailApp.createLabel(labelName);
      Logger.log(`[GMAIL_UTIL INFO] Successfully created label: "${labelName}"`);
    } catch (e) {
      Logger.log(`[GMAIL_UTIL ERROR] Failed to create label "${labelName}": ${e.message}\n${e.stack}`);
      return null;
    }
  } else {
    if (DEBUG_MODE) Logger.log(`[GMAIL_UTIL DEBUG] Label "${labelName}" already exists.`);
  }

  // NEW DEBUG LOG INSIDE getOrCreateLabel
  if (label) {
    Logger.log(`[GMAIL_UTIL DEBUG RETURN CHECK] Returning label for "${labelName}". typeof: ${typeof label}, constructor.name: ${label.constructor ? label.constructor.name : 'N/A'}`);
  } else {
    Logger.log(`[GMAIL_UTIL DEBUG RETURN CHECK] Returning NULL for "${labelName}".`);
  }
  return label;
}

// --- Helper: Apply Labels After Processing ---
function applyFinalLabels(threadOutcomes, processingLabel, processedLabelObj, manualReviewLabelObj) {
  const threadIdsToUpdate = Object.keys(threadOutcomes);
  if (threadIdsToUpdate.length === 0) { Logger.log("[INFO] LABEL_MGMT: No thread outcomes to process."); return; }
  Logger.log(`[INFO] LABEL_MGMT: Applying labels for ${threadIdsToUpdate.length} threads.`);
  let successfulLabelChanges = 0; let labelErrors = 0;
  if (!processingLabel || !processedLabelObj || !manualReviewLabelObj || typeof processingLabel.getName !== 'function' || typeof processedLabelObj.getName !== 'function' || typeof manualReviewLabelObj.getName !== 'function' ) { Logger.log(`[ERROR] LABEL_MGMT: Invalid label objects. Aborting.`); return; }
  const toProcessLabelName = processingLabel.getName();
  for (const threadId of threadIdsToUpdate) {
    const outcome = threadOutcomes[threadId]; const targetLabelToAdd = (outcome === 'manual') ? manualReviewLabelObj : processedLabelObj; const targetLabelNameToAdd = targetLabelToAdd.getName();
    try {
      const thread = GmailApp.getThreadById(threadId); if (!thread) { Logger.log(`[WARN] LABEL_MGMT: Thread ${threadId} not found. Skip.`); labelErrors++; continue; }
      const currentThreadLabels = thread.getLabels().map(l => l.getName()); let labelsActuallyChangedThisThread = false;
      if (currentThreadLabels.includes(toProcessLabelName)) { try { thread.removeLabel(processingLabel); if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Removed "${toProcessLabelName}" from ${threadId}`); labelsActuallyChangedThisThread = true; } catch (e) { Logger.log(`[WARN] LABEL_MGMT: Fail remove "${toProcessLabelName}" from ${threadId}: ${e}`); } }
      if (!currentThreadLabels.includes(targetLabelNameToAdd)) { try { thread.addLabel(targetLabelToAdd); Logger.log(`[INFO] LABEL_MGMT: Added "${targetLabelNameToAdd}" to ${threadId}`); labelsActuallyChangedThisThread = true; } catch (e) { Logger.log(`[ERROR] LABEL_MGMT: Fail add "${targetLabelNameToAdd}" to ${threadId}: ${e}`); labelErrors++; continue; } }
      else if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Thread ${threadId} already has "${targetLabelNameToAdd}".`);
      if (labelsActuallyChangedThisThread) { successfulLabelChanges++; Utilities.sleep(200 + Math.floor(Math.random() * 100)); } // Slightly longer pause for label changes
    } catch (e) { Logger.log(`[ERROR] LABEL_MGMT: General error for ${threadId}: ${e}`); labelErrors++; }
  }
  Logger.log(`[INFO] LABEL_MGMT: Finished. Success changes/verified: ${successfulLabelChanges}. Errors: ${labelErrors}.`);
}


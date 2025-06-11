/**
 * Project: Master Job Manager (formerly Automated Job Application Tracker & Dashboard)
 * Description: A comprehensive system for tracking job applications from lead generation to final status,
 *              integrating email parsing (job alerts & application updates), spreadsheet management,
 *              and dashboard analytics.
 * Author: Francis John LiButti (Originals), AI Integration & Refinements by Assistant,
 *         Integrated system design with Assistant
 * Current Version: v5.0 (Milestone: "Integrated Leads & Tracker System")
 * Key Changes in v5.0:
 *   - Integrated "Job Leads Tracker" functionality as a new module.
 *   - Major architectural refactor into multiple .gs files for improved organization.
 *   - Unified configuration and shared API key management.
 *   - Master setup function for full system initialization.
 *   - Enhanced Gmail filter creation for both modules.
 *   - Robust label ID fetching using Advanced Gmail Service.
 * Original Job Application Tracker was v4.5.
 * Date: [Current Date of this milestone, e.g., May 21, 2025]
 */

// --- Initial Setup Function ---
function initialSetup_LabelsAndSheet() {
  Logger.log(`\n==== STARTING INITIAL SETUP (LABELS, SHEETS, TRIGGERS, DASHBOARD, HELPER) ====`);
  let messages = [];
  let overallSuccess = true; // Assuming this is your main success flag
  let dummyDataWasAdded = false;

  // --- 1. Gmail Label Verification/Creation AND Get trackerToProcessLabelId ---
  Logger.log("[INFO] SETUP (Tracker): Verifying/Creating Labels and getting 'To Process' ID...");
  let trackerLabelsOk = true;
  let trackerToProcessLabelId = null;
  const trackerToProcessLabelNameConst = TRACKER_GMAIL_LABEL_TO_PROCESS; // From Config.gs

  try {
    // Ensure all parent and sibling labels are created
    getOrCreateLabel(MASTER_GMAIL_LABEL_PARENT); Utilities.sleep(200); // From Config.gs
    getOrCreateLabel(TRACKER_GMAIL_LABEL_PARENT); Utilities.sleep(200); // From Config.gs
    const toProcessLabelObject = getOrCreateLabel(trackerToProcessLabelNameConst);
    Utilities.sleep(100);
    getOrCreateLabel(TRACKER_GMAIL_LABEL_PROCESSED); Utilities.sleep(100); // From Config.gs
    getOrCreateLabel(TRACKER_GMAIL_LABEL_MANUAL_REVIEW); Utilities.sleep(100); // From Config.gs
    Logger.log("[INFO] SETUP (Tracker): Standard getOrCreateLabel calls completed for tracker labels.");

    if (toProcessLabelObject) {
        Logger.log(`[DEBUG] SETUP (Tracker): toProcessLabelObject (for "${trackerToProcessLabelNameConst}") is NOT null. Name: "${toProcessLabelObject.getName ? toProcessLabelObject.getName() : 'getName not function'}", Constructor: ${toProcessLabelObject.constructor ? toProcessLabelObject.constructor.name : 'N/A'}`);
    } else {
        Logger.log(`[DEBUG] SETUP (Tracker): toProcessLabelObject (for "${trackerToProcessLabelNameConst}") IS NULL after getOrCreateLabel.`);
    }

    Utilities.sleep(1000); // Pause before fetching ID via Advanced Service

    Logger.log(`[INFO] SETUP (Tracker): Attempting to get Label ID for "${trackerToProcessLabelNameConst}" using Advanced Gmail Service for filter setup.`);
    try {
        const advancedGmailService = Gmail; // Gmail Advanced Service
        if (!advancedGmailService || !advancedGmailService.Users || !advancedGmailService.Users.Labels) {
            throw new Error("Gmail API Advanced Service (Gmail) is not available or not properly enabled for tracker filter setup.");
        }
        const labelsListResponse = advancedGmailService.Users.Labels.list('me');
        if (labelsListResponse.labels && labelsListResponse.labels.length > 0) {
            const targetLabelInfo = labelsListResponse.labels.find(l => l.name === trackerToProcessLabelNameConst);
            if (targetLabelInfo && targetLabelInfo.id) {
                trackerToProcessLabelId = targetLabelInfo.id;
                Logger.log(`[INFO] SETUP (Tracker): Successfully retrieved Label ID via Advanced Service: "${trackerToProcessLabelId}" for label "${trackerToProcessLabelNameConst}".`);
            } else {
                Logger.log(`[WARN] SETUP (Tracker): Label "${trackerToProcessLabelNameConst}" not found in list from Advanced Gmail Service. TargetLabelInfo: ${JSON.stringify(targetLabelInfo)}`);
                trackerLabelsOk = false;
            }
        } else {
            Logger.log(`[WARN] SETUP (Tracker): No labels returned by Advanced Gmail Service list. Response: ${JSON.stringify(labelsListResponse)}`);
            trackerLabelsOk = false;
        }
    } catch (e) {
        Logger.log(`[ERROR] SETUP (Tracker): Error using Advanced Gmail Service to get label ID for "${trackerToProcessLabelNameConst}": ${e.message}\nStack: ${e.stack}`);
        trackerLabelsOk = false;
    }

    if (!trackerToProcessLabelId) {
        Logger.log(`[ERROR] SETUP (Tracker): CRITICAL - Could not obtain ID for "${trackerToProcessLabelNameConst}". New filter for application updates cannot be created.`);
        trackerLabelsOk = false;
    }

  } catch (e) {
      Logger.log(`[ERROR] SETUP (Tracker): A critical outer error occurred during label creation/ID fetching for Job Application Tracker: ${e.toString()}\nStack: ${e.stack}`);
      trackerLabelsOk = false;
  }

  if (trackerLabelsOk) {
    messages.push("Tracker Labels & 'To Process' Label ID for filter: OK.");
  } else {
    messages.push("Tracker Labels & 'To Process' Label ID for filter: FAILED (Could not get 'To Process' ID).");
    overallSuccess = false; // Labels are critical
  }

  // --- NEW: Create Gmail Filter for Application Updates (Job Application Tracker Module) ---
  if (trackerToProcessLabelId) { // Only proceed if we got the label ID
    Logger.log(`[INFO] SETUP (Tracker): Proceeding to create/verify Gmail filter for application updates, using Label ID: "${trackerToProcessLabelId}" and Query: "${TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES}"...`);
    let filterCreationAttempted = false;
    try {
        const gmailApiServiceForFilter = Gmail; // Gmail Advanced Service
        let filterExists = false;

        if (!gmailApiServiceForFilter || !gmailApiServiceForFilter.Users || !gmailApiServiceForFilter.Users.Settings || !gmailApiServiceForFilter.Users.Settings.Filters) {
            throw new Error("Gmail API Advanced Service (Gmail) is not available or not properly enabled for TRACKER filter creation.");
        }

        const existingFiltersResponse = gmailApiServiceForFilter.Users.Settings.Filters.list('me');
        // ADD THIS DEBUG LINE:
        Logger.log(`[DEBUG] SETUP (Tracker): Raw existingFiltersResponse from Gmail API for tracker: ${JSON.stringify(existingFiltersResponse)}`);

        let existingFiltersList = []; // Default to an empty array

        if (existingFiltersResponse && existingFiltersResponse.filter && Array.isArray(existingFiltersResponse.filter)) {
            existingFiltersList = existingFiltersResponse.filter;
            Logger.log(`[DEBUG] SETUP (Tracker): Successfully extracted ${existingFiltersList.length} filters from API response for tracker.`);
        } else if (existingFiltersResponse && typeof existingFiltersResponse === 'object' && !existingFiltersResponse.hasOwnProperty('filter')) {
            Logger.log(`[INFO] SETUP (Tracker): No 'filter' property found in existingFiltersResponse for tracker (user likely has no filters).`);
        } else if (!existingFiltersResponse) {
            Logger.log(`[WARN] SETUP (Tracker): existingFiltersResponse from Gmail API for tracker was null or undefined.`);
        } else {
             Logger.log(`[WARN] SETUP (Tracker): existingFiltersResponse.filter was not an array or structure was unexpected for tracker. Response: ${JSON.stringify(existingFiltersResponse)}`);
        }

        if (existingFiltersList.length > 0) {
            for (const filterItem of existingFiltersList) {
                if (filterItem.criteria && filterItem.criteria.query === TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES && // From Config.gs
                    filterItem.action && filterItem.action.addLabelIds && filterItem.action.addLabelIds.includes(trackerToProcessLabelId)) {
                    filterExists = true;
                    break;
                }
            }
        }

        if (!filterExists) {
            filterCreationAttempted = true;
            const filterResource = {
                criteria: { query: TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES }, // From Config.gs
                action: { addLabelIds: [trackerToProcessLabelId], removeLabelIds: ['INBOX'] }
            };
            Logger.log(`[DEBUG] SETUP (Tracker): Attempting to create filter for tracker with resource: ${JSON.stringify(filterResource)}`);
            const createdFilterResponse = gmailApiServiceForFilter.Users.Settings.Filters.create(filterResource, 'me');

            if (createdFilterResponse && createdFilterResponse.id) {
                Logger.log(`[INFO] SETUP (Tracker): Gmail filter for application updates CREATED successfully. Filter ID: ${createdFilterResponse.id}`);
                messages.push("Tracker App Update Filter: CREATED.");
            } else {
                Logger.log(`[ERROR] SETUP (Tracker): Gmail filter creation call for application updates did NOT return a valid filter ID. Response: ${JSON.stringify(createdFilterResponse)}`);
                messages.push("Tracker App Update Filter: FAILED (Creation API call issue).");
                overallSuccess = false;
            }
        } else {
            Logger.log(`[INFO] SETUP (Tracker): Gmail filter for application updates (Query: "${TRACKER_GMAIL_FILTER_QUERY_APP_UPDATES}", Label ID: "${trackerToProcessLabelId}") ALREADY EXISTS.`);
            messages.push("Tracker App Update Filter: Exists.");
        }
    } catch (e) {
        if (e.message && e.message.toLowerCase().includes("filter already exists")) {
            Logger.log(`[WARN] SETUP (Tracker): Gmail filter for application updates likely already exists (API reported error: ${e.message}).`);
            messages.push("Tracker App Update Filter: Exists (API reported).");
        } else {
            Logger.log(`[CRITICAL ERROR] SETUP (Tracker): Error during creation/checking of Gmail filter for application updates: ${e.toString()}\nStack: ${e.stack}. Attempted: ${filterCreationAttempted}`);
            messages.push(`Tracker App Update Filter: FAILED CRITICALLY - ${e.message}.`); // This is where your error likely originates for the message
            overallSuccess = false;
        }
    }
  } else {
      Logger.log(`[WARN] SETUP (Tracker): Skipping creation of application update filter because Label ID for "${TRACKER_GMAIL_LABEL_TO_PROCESS}" was NOT obtained (trackerToProcessLabelId is falsy).`);
      messages.push("Tracker App Update Filter: SKIPPED (Label ID for 'To Process' was not obtained).");
      // Consider if skipping the filter creation is a critical failure:
      // overallSuccess = false;
  }
// --- END NEW FILTER CREATION (Tracker) ---


  // --- 2. Spreadsheet & "Applications" Data Sheet Setup ---
  Logger.log("[INFO] SETUP: Verifying/Creating Data Sheet & Tab ('Applications')...");
  const { spreadsheet: ss, sheet: dataSh } = getOrCreateSpreadsheetAndSheet(); // Assumes dataSh is APP_TRACKER_SHEET_TAB_NAME

  if (!ss || !dataSh) {
    messages.push(`Data Sheet/Tab ('${APP_TRACKER_SHEET_TAB_NAME}'): FAILED.`); // Using config constant for clarity
    overallSuccess = false;
  } else {
    messages.push(`Data Sheet/Tab ('${APP_TRACKER_SHEET_TAB_NAME}'): OK. Using spreadsheet "${ss.getName()}" and tab "${dataSh.getName()}".`);

    // --- 3. Add Dummy Data (if sheet is new/empty for initial chart creation) ---
    if (dataSh && dataSh.getLastRow() <= 1) {
      Logger.log(`[INFO] SETUP: "${APP_TRACKER_SHEET_TAB_NAME}" sheet is empty/header only. Adding temporary dummy data.`);
      try {
        const today = new Date();
        const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
        const twoWeeksAgo = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);

        let dummyRowsData = [
          [new Date(), twoWeeksAgo, "LinkedIn", "Alpha Inc.", "Engineer I", DEFAULT_STATUS, DEFAULT_STATUS, twoWeeksAgo, "Applied to Alpha", "", ""],
          [new Date(), weekAgo, "Indeed", "Beta LLC", "Analyst Pro", APPLICATION_VIEWED_STATUS, APPLICATION_VIEWED_STATUS, weekAgo, "Viewed at Beta", "", ""],
          [new Date(), today, "Wellfound", "Gamma Solutions", "Manager X", INTERVIEW_STATUS, INTERVIEW_STATUS, today, "Interview at Gamma", "", ""]
        ];

        dummyRowsData = dummyRowsData.map(row => {
            while (row.length < TOTAL_COLUMNS_IN_APP_SHEET) row.push(""); // From Config.gs
            return row.slice(0, TOTAL_COLUMNS_IN_APP_SHEET);             // From Config.gs
        });

        dataSh.getRange(2, 1, dummyRowsData.length, TOTAL_COLUMNS_IN_APP_SHEET).setValues(dummyRowsData);
        dummyDataWasAdded = true;
        Logger.log(`[INFO] SETUP: Added ${dummyRowsData.length} dummy rows to "${APP_TRACKER_SHEET_TAB_NAME}" sheet.`);
      } catch (e) {
        Logger.log(`[ERROR] SETUP: Failed to add dummy data to "${APP_TRACKER_SHEET_TAB_NAME}" sheet: ${e.toString()} \nStack: ${e.stack}`);
      }
    }
  }


  // --- 4. Dashboard & Helper Sheet Setup (only if 'ss' is valid) ---
  if (ss) {
    Logger.log("[INFO] SETUP: Verifying/Creating and Formatting Dashboard Sheet...");
    try {
      const dashboardSheet = getOrCreateDashboardSheet(ss); // Uses DASHBOARD_TAB_NAME from Config.gs
      if (dashboardSheet) {
        formatDashboardSheet(dashboardSheet);
        messages.push(`Dashboard Sheet ('${DASHBOARD_TAB_NAME}'): OK.`);

        Logger.log("[INFO] SETUP: Verifying/Creating Helper Sheet...");
        const helperSheet = getOrCreateHelperSheet(ss); // Uses HELPER_SHEET_NAME from Config.gs
        if (helperSheet) {
          messages.push(`Helper Sheet ('${HELPER_SHEET_NAME}'): OK.`);
        } else {
          messages.push(`Helper Sheet ('${HELPER_SHEET_NAME}'): FAILED.`);
          overallSuccess = false; // Helper sheet is important for dashboard
        }
      } else {
        messages.push(`Dashboard Sheet ('${DASHBOARD_TAB_NAME}'): FAILED.`);
        overallSuccess = false;
      }
    } catch (e) {
      Logger.log(`[ERROR] SETUP: Error during dashboard/helper setup: ${e.toString()} \nStack: ${e.stack}`);
      messages.push(`Dashboard/Helper Sheet: FAILED - ${e.message}.`);
      overallSuccess = false;
    }

    // --- 5. Update Dashboard Metrics (will use dummy data if it was added) ---
    // Only if dashboard and helper setup was perceived as successful from the above block.
    if (messages.some(msg => msg.includes(`Dashboard Sheet ('${DASHBOARD_TAB_NAME}'): OK.`)) &&
        messages.some(msg => msg.includes(`Helper Sheet ('${HELPER_SHEET_NAME}'): OK.`))) {
      try {
        Logger.log("[INFO] SETUP: Attempting initial dashboard metrics (chart data) update...");
        updateDashboardMetrics();
        messages.push("Dashboard Chart Data: Update attempted.");
      } catch (e) {
        Logger.log(`[ERROR] SETUP: Failed during updateDashboardMetrics: ${e.toString()} \nStack: ${e.stack}`);
        messages.push(`Dashboard Chart Data: FAILED - ${e.message}.`);
        // overallSuccess = false; // Might not be critical if it fails only here at setup
      }
    } else {
        messages.push("Dashboard Chart Data: Update SKIPPED due to issues with Dashboard/Helper sheet setup.");
    }


    // --- 6. Remove Dummy Data (if it was added AND dataSh is valid) ---
    if (dummyDataWasAdded && dataSh && typeof dataSh.deleteRows === 'function') {
      Logger.log("[INFO] SETUP: Removing temporary dummy data.");
      try {
        dataSh.deleteRows(2, 3); // We added 3 dummy rows
        Logger.log("[INFO] SETUP: Dummy data removed.");
      } catch (e) {
        Logger.log(`[ERROR] SETUP: Failed to remove dummy data: ${e.toString()} \nStack: ${e.stack}.`);
      }
    }
  } else {
    messages.push("Dashboard, Helper Sheets, and Metrics Update: SKIPPED (Spreadsheet object 'ss' was not available).");
    // If ss isn't available, dashboard/helper related setups were already marked as failed or skipped
    // so overallSuccess should already be false or reflect this state.
  }


  // --- 7. Trigger Verification/Creation ---
  Logger.log("[INFO] SETUP: Verifying/Creating Triggers...");
  if (createTimeDrivenTrigger('processJobApplicationEmails', 1)) { messages.push("Email Processor Trigger ('processJobApplicationEmails'): CREATED."); }
  else { messages.push("Email Processor Trigger ('processJobApplicationEmails'): Not newly created (check logs)."); }

  if (createOrVerifyStaleRejectTrigger('markStaleApplicationsAsRejected', 2)) { messages.push("Stale Reject Trigger ('markStaleApplicationsAsRejected'): CREATED."); }
  else { messages.push("Stale Reject Trigger ('markStaleApplicationsAsRejected'): Not newly created (check logs)."); }

  // --- 8. Final Summary and UI Alert ---
  const finalMsg = `Initial setup process ${overallSuccess ? "completed." : "encountered ISSUES."}\n\nSummary:\n- ${messages.join('\n- ')}`;
  Logger.log(`\n==== INITIAL SETUP ${overallSuccess ? "OK" : "ISSUES"} ====\n${finalMsg.replace(/\n- /g,'\n  - ')}`);
  try {
    if (SpreadsheetApp.getActiveSpreadsheet() && SpreadsheetApp.getUi()) {
        SpreadsheetApp.getUi().alert( `Setup ${overallSuccess?"Complete":"Issues"}`, finalMsg, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
        Logger.log("[INFO] UI Alert for initial setup skipped (no active spreadsheet UI or context).");
    }
  } catch(e) {
    Logger.log(`[INFO] UI Alert skipped: ${e.message}.`);
  }
  Logger.log("==== END INITIAL SETUP ====");
}

// --- Main Email Processing Function ---
function processJobApplicationEmails() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING PROCESS JOB EMAILS (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  const scriptProperties = PropertiesService.getScriptProperties(); // Use ScriptProperties for API Key
  const geminiApiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY);
  let useGemini = false;
  Logger.log(`[DEBUG] API_KEY_CHECK: Attempting key: "${GEMINI_API_KEY_PROPERTY}" using getScriptProperties()`);
  if (geminiApiKey) {
    Logger.log(`[DEBUG] API_KEY_CHECK: Retrieved: "${geminiApiKey.substring(0,10)}..." (len: ${geminiApiKey.length})`);
    if (geminiApiKey.trim() !== "" && geminiApiKey.startsWith("AIza") && geminiApiKey.length > 30) {
      useGemini = true; Logger.log("[INFO] PROCESS_EMAIL: Gemini API Key VALID. Will use Gemini.");
    } else { Logger.log(`[WARN] PROCESS_EMAIL: Gemini API Key found but seems INVALID. Using regex.`); }
  } else { Logger.log("[WARN] PROCESS_EMAIL: Gemini API Key NOT FOUND. Using regex."); }
  if (!useGemini) Logger.log("[INFO] PROCESS_EMAIL: Final decision: Using regex-based parser.");

  const { spreadsheet: ss, sheet: dataSheet } = getOrCreateSpreadsheetAndSheet(); 
  if (!ss || !dataSheet) { Logger.log(`[FATAL ERROR] PROCESS_EMAIL: Sheet/Tab fail. Aborting.`); return; }
  Logger.log(`[INFO] PROCESS_EMAIL: Sheet OK: "${ss.getName()}" / "${dataSheet.getName()}"`);

  let procLbl, processedLblObj, manualLblObj;
try { 
  // Use the correctly named constants from Config.gs
  procLbl = GmailApp.getUserLabelByName(TRACKER_GMAIL_LABEL_TO_PROCESS);       // <<<< CORRECTED
  processedLblObj = GmailApp.getUserLabelByName(TRACKER_GMAIL_LABEL_PROCESSED);  // <<<< CORRECTED
  manualLblObj = GmailApp.getUserLabelByName(TRACKER_GMAIL_LABEL_MANUAL_REVIEW); // <<<< CORRECTED
  
  if (!procLbl) throw new Error(`Label "${TRACKER_GMAIL_LABEL_TO_PROCESS}" not found.`);
  if (!processedLblObj) throw new Error(`Label "${TRACKER_GMAIL_LABEL_PROCESSED}" not found.`);
  if (!manualLblObj) throw new Error(`Label "${TRACKER_GMAIL_LABEL_MANUAL_REVIEW}" not found.`);
  
  if (DEBUG_MODE) Logger.log(`[DEBUG] PROCESS_EMAIL: Core Gmail labels for tracker verified (To Process, Processed, Manual Review).`);
} catch(e) {
  Logger.log(`[FATAL ERROR] PROCESS_EMAIL: Tracker labels missing or error fetching them! Error: ${e.message}`); 
  return; // Abort if labels can't be fetched
}

  const lastR = dataSheet.getLastRow(); 
  const existingDataCache = {}; 
  const processedEmailIds = new Set();
  if (lastR >= 2) {
    Logger.log(`[INFO] PRELOAD: Loading data from "${dataSheet.getName()}" (Rows 2 to ${lastR})...`);
    try {
      const colsToPreload = [COMPANY_COL, JOB_TITLE_COL, EMAIL_ID_COL, STATUS_COL, PEAK_STATUS_COL];
      const minCol = Math.min(...colsToPreload); const maxCol = Math.max(...colsToPreload);
      const numColsToRead = maxCol - minCol + 1;
      if (numColsToRead < 1 || minCol < 1) throw new Error("Invalid preload column calculation.");

      const preloadRange = dataSheet.getRange(2, minCol, lastR - 1, numColsToRead);
      const preloadValues = preloadRange.getValues();
      const coIdx = COMPANY_COL-minCol, tiIdx = JOB_TITLE_COL-minCol, idIdx = EMAIL_ID_COL-minCol, stIdx = STATUS_COL-minCol, pkIdx = PEAK_STATUS_COL-minCol;

      for (let i = 0; i < preloadValues.length; i++) {
        const rN = i + 2, rD = preloadValues[i];
        const eId = rD[idIdx]?.toString().trim()||"", oCo = rD[coIdx]?.toString().trim()||"", oTi = rD[tiIdx]?.toString().trim()||"", cS  = rD[stIdx]?.toString().trim()||"", cPkS = rD[pkIdx]?.toString().trim()||"";
        if(eId) processedEmailIds.add(eId);
        const cL = oCo.toLowerCase();
        if(cL && cL !== MANUAL_REVIEW_NEEDED.toLowerCase() && cL !== 'n/a'){
          if(!existingDataCache[cL]) existingDataCache[cL]=[];
          existingDataCache[cL].push({row:rN,emailId:eId,company:oCo,title:oTi,status:cS, peakStatus: cPkS});
        }
      }
      Logger.log(`[INFO] PRELOAD: Complete. Cached ${Object.keys(existingDataCache).length} co. ${processedEmailIds.size} IDs.`);
    } catch (e) { Logger.log(`[FATAL ERROR] Preload: ${e.toString()}\nStack:${e.stack}\nAbort.`); return; }
  } else { Logger.log(`[INFO] PRELOAD: Sheet empty or header only.`); }

  const THREAD_PROCESSING_LIMIT = 20; // Process up to 15 threads per run
  let threadsToProcess = [];
  try { threadsToProcess = procLbl.getThreads(0, THREAD_PROCESSING_LIMIT); } 
  catch (e) { Logger.log(`[ERROR] GATHER_THREADS: Failed for "${procLbl.getName()}": ${e}`); return; }

  const messagesToSort = []; let skippedCount = 0; let fetchErrorCount = 0;
  if (DEBUG_MODE) Logger.log(`[DEBUG] GATHER_THREADS: Found ${threadsToProcess.length} threads.`);
  for (const thread of threadsToProcess) {
    const tId = thread.getId();
    try {
      const mIT = thread.getMessages();
      for (const msg of mIT) {
        const mId = msg.getId();
        if (!processedEmailIds.has(mId)) { messagesToSort.push({ message: msg, date: msg.getDate(), threadId: tId }); } 
        else { skippedCount++; }
      }
    } catch (e) { Logger.log(`[ERROR] GATHER_MESSAGES: Thread ${tId}: ${e}`); fetchErrorCount++; }
  }
  Logger.log(`[INFO] GATHER_MESSAGES: New: ${messagesToSort.length}. Skipped: ${skippedCount}. Fetch errors: ${fetchErrorCount}.`);

  if (messagesToSort.length === 0) {
    Logger.log("[INFO] PROCESS_LOOP: No new messages.");
    try { updateDashboardMetrics(); } catch (e_dash) { Logger.log(`[ERROR] Dashboard update (no new msgs): ${e_dash.message}`); }
    Logger.log(`==== SCRIPT FINISHED (${new Date().toLocaleString()}) - No new messages ====`);
    return;
  }

  messagesToSort.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
  Logger.log(`[INFO] PROCESS_LOOP: Sorted ${messagesToSort.length} new messages.`);
  
  let threadProcessingOutcomes = {}; 
  let pTRC = 0, sUC = 0, nEC = 0, pEC = 0;

  for (let i = 0; i < messagesToSort.length; i++) {
    const elapsedTime = (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000;
    if (elapsedTime > 320) { Logger.log(`[WARN] Time limit nearing (${elapsedTime}s). Stopping loop.`); break; } // Slightly increased margin

    const entry = messagesToSort[i];
    const { message, date: emailDateObj, threadId } = entry;
    const emailDate = new Date(emailDateObj); 
    const msgId = message.getId();
    const pSTM = new Date();
    if(DEBUG_MODE)Logger.log(`\n--- Processing Msg ${i+1}/${messagesToSort.length} (ID: ${msgId}, Thread: ${threadId}) ---`);

    let companyName=MANUAL_REVIEW_NEEDED, jobTitle=MANUAL_REVIEW_NEEDED, applicationStatus=null; 
    let plainBodyText=null, requiresManualReview=false, sheetWriteOpSuccess=false;

    try {
      const emailSubject=message.getSubject()||"", senderEmail=message.getFrom()||"", emailPermaLink=`https://mail.google.com/mail/u/0/#inbox/${msgId}`, currentTimestamp=new Date();
      let detectedPlatform=DEFAULT_PLATFORM;
      try{const eAM=senderEmail.match(/<([^>]+)>/);if(eAM&&eAM[1]){const sD=eAM[1].split('@')[1]?.toLowerCase();if(sD){for(const k in PLATFORM_DOMAIN_KEYWORDS){if(sD.includes(k)){detectedPlatform=PLATFORM_DOMAIN_KEYWORDS[k];break;}}}}if(DEBUG_MODE)Logger.log(`[DEBUG] Platform: ${detectedPlatform}`);}catch(ePlat){Logger.log(`WARN: Plat Detect Err: ${ePlat}`);}
      try{plainBodyText=message.getPlainBody();}catch(eBody){Logger.log(`WARN: Get Body Fail Msg ${msgId}: ${eBody}`);plainBodyText="";}

      if (useGemini && plainBodyText && plainBodyText.trim() !== "") {
        const gRes = callGemini_forApplicationDetails(emailSubject, plainBodyText, geminiApiKey);
        if (gRes) { 
            companyName=gRes.company||MANUAL_REVIEW_NEEDED; jobTitle=gRes.title||MANUAL_REVIEW_NEEDED; applicationStatus=gRes.status;
            Logger.log(`[INFO] Gemini: C:"${companyName}", T:"${jobTitle}", S:"${applicationStatus}"`);
            if(!applicationStatus||applicationStatus===MANUAL_REVIEW_NEEDED||applicationStatus==="Update/Other"){const kS=parseBodyForStatus(plainBodyText); if(kS&&kS!==DEFAULT_STATUS){applicationStatus=kS;} else if(!applicationStatus&&kS===DEFAULT_STATUS){applicationStatus=DEFAULT_STATUS;}}
        } else { 
            Logger.log(`[WARN] Gemini fail Msg ${msgId}. Fallback regex.`);
            const rEx=extractCompanyAndTitle(message,detectedPlatform,emailSubject,plainBodyText);companyName=rEx.company;jobTitle=rEx.title;applicationStatus=parseBodyForStatus(plainBodyText);
        }
      } else { 
          const rEx=extractCompanyAndTitle(message,detectedPlatform,emailSubject,plainBodyText);companyName=rEx.company;jobTitle=rEx.title;applicationStatus=parseBodyForStatus(plainBodyText);
          if(DEBUG_MODE) Logger.log(`[DEBUG] Regex Parse: C:"${companyName}", T:"${jobTitle}", S (body scan):"${applicationStatus}"`);
      }
      
      requiresManualReview = (companyName === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED);
      const finalStatusForSheet = applicationStatus || DEFAULT_STATUS;
      const companyCacheKey = (companyName !== MANUAL_REVIEW_NEEDED) ? companyName.toLowerCase() : `_manual_review_placeholder_${msgId}`;
      let existingRowInfo = null; let targetSheetRow = -1;

      if (companyName !== MANUAL_REVIEW_NEEDED && existingDataCache[companyCacheKey]) {
          const pM = existingDataCache[companyCacheKey];
          if (jobTitle !== MANUAL_REVIEW_NEEDED) existingRowInfo = pM.find(e => e.title && e.title.toLowerCase() === jobTitle.toLowerCase());
          if (!existingRowInfo && pM.length > 0) existingRowInfo = pM.reduce((l,c)=>(c.row > l.row ? c : l), pM[0]);
          if (existingRowInfo) targetSheetRow = existingRowInfo.row;
      }

      if (targetSheetRow !== -1 && existingRowInfo) { // UPDATE EXISTING ROW
        const rangeToUpdate = dataSheet.getRange(targetSheetRow, 1, 1, TOTAL_COLUMNS_IN_APP_SHEET);
        const currentSheetValuesRow = rangeToUpdate.getValues()[0];
        let newSheetValues = [...currentSheetValuesRow];

        newSheetValues[PROCESSED_TIMESTAMP_COL - 1] = currentTimestamp;
        // EMAIL_DATE_COL: Update only if this email is newer than what's in sheet for this field specifically
        const existingEmailDate = currentSheetValuesRow[EMAIL_DATE_COL-1];
        if (!(existingEmailDate instanceof Date) || emailDate.getTime() > existingEmailDate.getTime()) {
            newSheetValues[EMAIL_DATE_COL - 1] = emailDate;
        }
        const existingLastUpdate = newSheetValues[LAST_UPDATE_DATE_COL-1];
        if(!(existingLastUpdate instanceof Date) || emailDate.getTime() > existingLastUpdate.getTime()){
            newSheetValues[LAST_UPDATE_DATE_COL-1] = emailDate;
        }
        newSheetValues[EMAIL_SUBJECT_COL-1]=emailSubject; newSheetValues[EMAIL_LINK_COL-1]=emailPermaLink; newSheetValues[EMAIL_ID_COL-1]=msgId; newSheetValues[PLATFORM_COL-1]=detectedPlatform;
        if(companyName!==MANUAL_REVIEW_NEEDED && (newSheetValues[COMPANY_COL-1]===MANUAL_REVIEW_NEEDED || companyName.toLowerCase()!==newSheetValues[COMPANY_COL-1]?.toLowerCase())) newSheetValues[COMPANY_COL-1]=companyName;
        if(jobTitle!==MANUAL_REVIEW_NEEDED && (newSheetValues[JOB_TITLE_COL-1]===MANUAL_REVIEW_NEEDED || jobTitle.toLowerCase()!==newSheetValues[JOB_TITLE_COL-1]?.toLowerCase())) newSheetValues[JOB_TITLE_COL-1]=jobTitle;
        
        const statusInSheetBeforeUpdate = currentSheetValuesRow[STATUS_COL-1]?.toString().trim() || DEFAULT_STATUS;
        let statusForThisUpdate = finalStatusForSheet; 
        // Refined Status Update Logic
        if (statusInSheetBeforeUpdate !== ACCEPTED_STATUS || statusForThisUpdate === ACCEPTED_STATUS) {
            const curRank = STATUS_HIERARCHY[statusInSheetBeforeUpdate] ?? 0;
            const newRank = STATUS_HIERARCHY[statusForThisUpdate] ?? 0;
            if (newRank >= curRank || statusForThisUpdate === REJECTED_STATUS || statusForThisUpdate === OFFER_STATUS ) {
                 newSheetValues[STATUS_COL - 1] = statusForThisUpdate;
            } else { Logger.log(`[DEBUG] Status not updated: "${statusForThisUpdate}" (rank ${newRank}) is not >= current "${statusInSheetBeforeUpdate}" (rank ${curRank}) and not final.`); }
        } else { Logger.log(`[DEBUG] Status is "${ACCEPTED_STATUS}", not changing to "${statusForThisUpdate}".`); }
        const statusAfterUpdate = newSheetValues[STATUS_COL - 1];

        // --- PEAK STATUS LOGIC for UPDATE ---
        let currentPeakFromSheet = existingRowInfo.peakStatus || currentSheetValuesRow[PEAK_STATUS_COL - 1]?.toString().trim();
        if (!currentPeakFromSheet || currentPeakFromSheet === MANUAL_REVIEW_NEEDED || currentPeakFromSheet === "") currentPeakFromSheet = DEFAULT_STATUS; 
        
        const currentPeakRank = STATUS_HIERARCHY[currentPeakFromSheet] ?? -2;
        const newStatusRankForPeak = STATUS_HIERARCHY[statusAfterUpdate] ?? -2; // Use the just-updated status for peak eval
        const excludedFromPeak = new Set([REJECTED_STATUS, ACCEPTED_STATUS, MANUAL_REVIEW_NEEDED, "Update/Other"]);

        let updatedPeakStatus = currentPeakFromSheet; 
        if (newStatusRankForPeak > currentPeakRank && !excludedFromPeak.has(statusAfterUpdate)) {
            updatedPeakStatus = statusAfterUpdate;
        } else if (currentPeakFromSheet === DEFAULT_STATUS && !excludedFromPeak.has(statusAfterUpdate) && STATUS_HIERARCHY[statusAfterUpdate] > STATUS_HIERARCHY[DEFAULT_STATUS]) {
            updatedPeakStatus = statusAfterUpdate; 
        }
        newSheetValues[PEAK_STATUS_COL - 1] = updatedPeakStatus;
        if(DEBUG_MODE) Logger.log(`[DEBUG] Peak Status Update: Row ${targetSheetRow}. Current Peak: "${currentPeakFromSheet}", New Current Status: "${statusAfterUpdate}", New Peak Set: "${updatedPeakStatus}"`);
        
        rangeToUpdate.setValues([newSheetValues]);
        Logger.log(`[INFO] SHEET WRITE: Updated Row ${targetSheetRow}. Status: "${statusAfterUpdate}", Peak: "${updatedPeakStatus}"`);
        sUC++; sheetWriteOpSuccess = true;
        const cacheKey = (newSheetValues[COMPANY_COL - 1] !== MANUAL_REVIEW_NEEDED) ? newSheetValues[COMPANY_COL - 1].toLowerCase() : companyCacheKey;
        if(existingDataCache[cacheKey]){existingDataCache[cacheKey]=existingDataCache[cacheKey].map(e=>e.row===targetSheetRow?{...e, status:statusAfterUpdate, peakStatus:updatedPeakStatus}:e);}

      } else { // APPEND NEW ROW
        const nRC = new Array(TOTAL_COLUMNS_IN_APP_SHEET).fill("");
        nRC[PROCESSED_TIMESTAMP_COL-1]=currentTimestamp; nRC[EMAIL_DATE_COL-1]=emailDate; nRC[PLATFORM_COL-1]=detectedPlatform; nRC[COMPANY_COL-1]=companyName; nRC[JOB_TITLE_COL-1]=jobTitle; nRC[STATUS_COL-1]=finalStatusForSheet; nRC[LAST_UPDATE_DATE_COL-1]=emailDate; nRC[EMAIL_SUBJECT_COL-1]=emailSubject; nRC[EMAIL_LINK_COL-1]=emailPermaLink; nRC[EMAIL_ID_COL-1]=msgId;

        const excludedFromPeakInit = new Set([REJECTED_STATUS, ACCEPTED_STATUS, MANUAL_REVIEW_NEEDED, "Update/Other"]);
        if (!excludedFromPeakInit.has(finalStatusForSheet)) { nRC[PEAK_STATUS_COL - 1] = finalStatusForSheet; } 
        else { nRC[PEAK_STATUS_COL - 1] = DEFAULT_STATUS; }
        if(DEBUG_MODE) Logger.log(`[DEBUG] Peak Status (New Row): Initial Peak set to "${nRC[PEAK_STATUS_COL - 1]}" for initial status "${finalStatusForSheet}"`);
        
        dataSheet.appendRow(nRC);
        const nSRN = dataSheet.getLastRow();
        Logger.log(`[INFO] SHEET WRITE: Appended Row ${nSRN}. Status: "${finalStatusForSheet}", Peak: "${nRC[PEAK_STATUS_COL - 1]}"`);
        nEC++; sheetWriteOpSuccess = true;
        const newEntryCacheKey = (nRC[COMPANY_COL-1] !== MANUAL_REVIEW_NEEDED) ? nRC[COMPANY_COL-1].toLowerCase() : companyCacheKey; // Use actual company or placeholder
        if(!existingDataCache[newEntryCacheKey]) existingDataCache[newEntryCacheKey]=[];
        existingDataCache[newEntryCacheKey].push({row:nSRN,emailId:msgId,company:nRC[COMPANY_COL-1],title:nRC[JOB_TITLE_COL-1],status:nRC[STATUS_COL-1], peakStatus:nRC[PEAK_STATUS_COL-1]});
      }

      // --- Corrected threadProcessingOutcomes logic ---
      if (sheetWriteOpSuccess) {
        pTRC++;
        processedEmailIds.add(msgId);
        let messageOutcome = (requiresManualReview || companyName === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED) ? 'manual' : 'done';
        
        // If thread is already marked 'manual' by a previous message in THIS run, it stays 'manual'.
        // Otherwise, set or update the outcome.
        if (threadProcessingOutcomes[threadId] !== 'manual') {
            threadProcessingOutcomes[threadId] = messageOutcome;
        }
        // If current message dictates 'manual', ensure thread outcome reflects that (it might be first message or override a 'done')
        if (messageOutcome === 'manual') {
             threadProcessingOutcomes[threadId] = 'manual';
        }
        // If threadProcessingOutcomes[threadId] was undefined, it's now set.
        if (DEBUG_MODE) Logger.log(`[DEBUG] Thread ${threadId} outcome for labeling set to: ${threadProcessingOutcomes[threadId]} (current message was: ${messageOutcome})`);
      } else {
        pEC++;
        threadProcessingOutcomes[threadId] = 'manual'; 
        Logger.log(`[ERROR] SHEET WRITE FAILED Msg ${msgId}. Thread ${threadId} auto-marked manual.`);
      }

    } catch (e) {
      Logger.log(`[FATAL ERROR] Processing Msg ${msgId} (Thread ${threadId}): ${e.message}\nStack: ${e.stack}`);
      threadProcessingOutcomes[threadId] = 'manual'; // Ensure thread is marked for manual review on error
      pEC++;
    }
    if(DEBUG_MODE){const pTTM=(new Date().getTime()-pSTM.getTime())/1000;Logger.log(`--- End Msg ${i+1}/${messagesToSort.length} --- Time: ${pTTM}s ---`);} 
    Utilities.sleep(200 + Math.floor(Math.random() * 150)); // Slightly reduced sleep
  }

  Logger.log(`\n[INFO] PROCESS_LOOP: Finished loop. Parsed: ${pTRC}, Sheet Updates: ${sUC}, New Entries: ${nEC}, Processing Errors: ${pEC}.`);
  if(DEBUG_MODE && Object.keys(threadProcessingOutcomes).length > 0) Logger.log(`[DEBUG] Final Thread Outcomes for Labeling: ${JSON.stringify(threadProcessingOutcomes)}`);
  else if (DEBUG_MODE) Logger.log(`[DEBUG] No thread outcomes recorded for labeling (threadProcessingOutcomes empty).`);

  applyFinalLabels(threadProcessingOutcomes, procLbl, processedLblObj, manualLblObj);
  
  try {
    Logger.log("[INFO] PROCESS_EMAIL: Attempting final dashboard metrics update...");
    updateDashboardMetrics();
  } catch (e_dash_final) {
    Logger.log(`[ERROR] PROCESS_EMAIL: Failed final dashboard update call: ${e_dash_final.message}`);
  }

  const SCRIPT_END_TIME = new Date();
  Logger.log(`\n==== SCRIPT FINISHED (${SCRIPT_END_TIME.toLocaleString()}) === Total Time: ${(SCRIPT_END_TIME.getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}

// --- Auto-Reject Stale Applications Function ---
function markStaleApplicationsAsRejected() {
  const SSTA = new Date();
  Logger.log(`\n==== AUTO-REJECT STALE START (${SSTA.toLocaleString()}) ====`);
  const { spreadsheet: ss, sheet: dataSheet } = getOrCreateSpreadsheetAndSheet(); // Ensure this uses APP_TRACKER_SHEET_TAB_NAME

  if (!ss || !dataSheet) {
    Logger.log(`[FATAL] AUTO-REJECT_STALE: No Sheet/Tab (expected "${APP_TRACKER_SHEET_TAB_NAME}"). Abort.`);
    return;
  }
  Logger.log(`[INFO] AUTO-REJECT_STALE: Using "${dataSheet.getName()}" in "${ss.getName()}".`);

  const dataRange = dataSheet.getDataRange();
  const sheetValues = dataRange.getValues();

  if (sheetValues.length <= 1) {
    Logger.log("[INFO] AUTO-REJECT_STALE: No data rows in sheet to process.");
    return;
  }

  const currentDate = new Date();
  const staleThresholdDate = new Date();
  staleThresholdDate.setDate(currentDate.getDate() - (WEEKS_THRESHOLD * 7)); // WEEKS_THRESHOLD from Config.gs
  Logger.log(`[INFO] AUTO-REJECT_STALE: Current Date: ${currentDate.toLocaleDateString()}, Stale if Last Update < ${staleThresholdDate.toLocaleDateString()} (Threshold: ${WEEKS_THRESHOLD} weeks)`);

  let updatedApplicationsCount = 0;
  let rowsProcessedForStaleCheck = 0;

  for (let i = 1; i < sheetValues.length; i++) { // Start from 1 to skip header row
    const currentRow = sheetValues[i];
    const currentRowNumber = i + 1; // Sheet row number for logging

    const currentStatus = currentRow[STATUS_COL - 1] ? currentRow[STATUS_COL - 1].toString().trim() : ""; // STATUS_COL from Config.gs
    const lastUpdateDateValue = currentRow[LAST_UPDATE_DATE_COL - 1]; // LAST_UPDATE_DATE_COL from Config.gs
    let lastUpdateDate;

    // Attempt to parse the lastUpdateDateValue
    if (lastUpdateDateValue instanceof Date && !isNaN(lastUpdateDateValue)) {
      lastUpdateDate = lastUpdateDateValue;
    } else if (lastUpdateDateValue && typeof lastUpdateDateValue === 'string' && lastUpdateDateValue.trim() !== "") {
      const parsedDate = new Date(lastUpdateDateValue);
      if (!isNaN(parsedDate)) {
        lastUpdateDate = parsedDate;
      } else {
        if (DEBUG_MODE) Logger.log(`[DEBUG] AUTO-REJECT_STALE: Row ${currentRowNumber} Skip - Invalid Last Update Date string: "${lastUpdateDateValue}"`);
        continue; // Skip if date string is invalid
      }
    } else {
      if (DEBUG_MODE) Logger.log(`[DEBUG] AUTO-REJECT_STALE: Row ${currentRowNumber} Skip - Missing or unparseable Last Update Date: "${lastUpdateDateValue}"`);
      continue; // Skip if no valid date
    }

    rowsProcessedForStaleCheck++;
    // Log details for every row being considered (before skip conditions)
    if (DEBUG_MODE) {
        Logger.log(`[DEBUG ROW ${currentRowNumber} CHECK] Status: "${currentStatus}", LastUpdate: ${lastUpdateDate.toLocaleDateString()}`);
    }


    // Condition 1: Skip if status is final or requires manual review or is empty
    if (FINAL_STATUSES_FOR_STALE_CHECK.has(currentStatus) || !currentStatus || currentStatus === MANUAL_REVIEW_NEEDED) { // FINAL_STATUSES_FOR_STALE_CHECK, MANUAL_REVIEW_NEEDED from Config.gs
      if (DEBUG_MODE) Logger.log(`[DEBUG] AUTO-REJECT_STALE: Row ${currentRowNumber} Skip - Status "${currentStatus}" is final, manual review, or empty.`);
      continue;
    }

    // Condition 2: Skip if last update date is not older than the threshold
    if (lastUpdateDate.getTime() >= staleThresholdDate.getTime()) { // Compare getTime() for precision
      if (DEBUG_MODE) Logger.log(`[DEBUG] AUTO-REJECT_STALE: Row ${currentRowNumber} Skip - Last Update Date ${lastUpdateDate.toLocaleDateString()} is NOT older than threshold ${staleThresholdDate.toLocaleDateString()}.`);
      continue;
    }

    // If both conditions above are passed, the application is stale and needs updating
    const oldPeakStatus = currentRow[PEAK_STATUS_COL - 1] ? currentRow[PEAK_STATUS_COL - 1].toString().trim() : 'Not Set'; // PEAK_STATUS_COL from Config.gs
    Logger.log(`[INFO] AUTO-REJECT_STALE: Row ${currentRowNumber} - MARKING STALE. LastUpd: ${lastUpdateDate.toLocaleDateString()}, OldStat: "${currentStatus}" -> NewStat: "${REJECTED_STATUS}". Peak was: "${oldPeakStatus}"`); // REJECTED_STATUS from Config.gs
    
    sheetValues[i][STATUS_COL - 1] = REJECTED_STATUS;
    sheetValues[i][LAST_UPDATE_DATE_COL - 1] = currentDate; // Update to current date
    sheetValues[i][PROCESSED_TIMESTAMP_COL - 1] = currentDate; // PROCESSED_TIMESTAMP_COL from Config.gs
    updatedApplicationsCount++;
  }

  Logger.log(`[INFO] AUTO-REJECT_STALE: Total rows processed for stale check (after date validation): ${rowsProcessedForStaleCheck}.`)
  if (updatedApplicationsCount > 0) {
    Logger.log(`[INFO] AUTO-REJECT_STALE: Found ${updatedApplicationsCount} stale applications to update. Writing changes to sheet...`);
    try {
      dataRange.setValues(sheetValues); // Write back the entire modified array
      Logger.log(`[INFO] AUTO-REJECT_STALE: Successfully updated ${updatedApplicationsCount} stale applications in the sheet.`);
    } catch (e) {
      Logger.log(`[ERROR] AUTO-REJECT_STALE: Sheet write operation failed: ${e.toString()}\nStack: ${e.stack}`);
    }
  } else {
    Logger.log("[INFO] AUTO-REJECT_STALE: No stale applications found meeting all criteria for update.");
  }
  const SETA = new Date();
  Logger.log(`==== AUTO-REJECT STALE END (${SETA.toLocaleString()}) ==== Time: ${(SETA.getTime() - SSTA.getTime()) / 1000}s ====`);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Job Manager Admin')
      .addItem('▶️ RUN FULL PROJECT SETUP', 'runFullProjectInitialSetup') // New Item
      .addSeparator()
      .addItem('Set Shared Gemini API Key', 'setSharedGeminiApiKey_UI')
      .addItem('Show All User Properties', 'showAllUserProperties')
      .addSeparator()
      .addItem('MANUAL: Set Hardcoded API Key (Temporary)', 'TEMPORARY_manualSetSharedGeminiApiKey')
      .addSeparator()
      .addItem('Setup: Job Tracker Module Only', 'initialSetup_LabelsAndSheet')
      .addItem('Setup: Job Leads Module Only', 'runInitialSetup_JobLeadsModule')
      .addToUi();
}

/**
 * Runs the complete initial setup for ALL modules of the Master Job Manager:
 * - Sets up the Job Application Tracker (sheets, labels, triggers).
 * - Sets up the Job Leads Tracker (sheet tab, labels, filter, trigger).
 * Designed to be run manually from the Apps Script editor or a custom menu once.
 */

function runFullProjectInitialSetup() {
  Logger.log("==== STARTING FULL PROJECT INITIAL SETUP (v2 Debug) ===="); // Added version for log clarity
  let overallSuccess = true;
  let setupMessages = [];

  // --- 1. Setup Job Application Tracker Module ---
  Logger.log("\n--- Attempting Setup for Job Application Tracker Module ---");
  try {
    initialSetup_LabelsAndSheet(); 
    Logger.log("--- COMPLETED CALL to initialSetup_LabelsAndSheet (Job Application Tracker Module) ---"); // NEW DEBUG LOG
    setupMessages.push("Job Application Tracker Module: Setup function executed.");
  } catch (e) {
    Logger.log(`ERROR during Job Application Tracker setup: ${e.toString()}\n${e.stack}`);
    setupMessages.push(`Job Application Tracker Module: FAILED - ${e.message}`);
    overallSuccess = false;
  }

  Logger.log("\n--- CHECKPOINT: Proceeding to Job Leads Tracker Module Setup? ---"); // NEW DEBUG LOG

  // --- 2. Setup Job Leads Tracker Module ---
  Logger.log("\n--- Attempting Setup for Job Leads Tracker Module ---");
  try {
    runInitialSetup_JobLeadsModule(); 
    Logger.log("--- COMPLETED CALL to runInitialSetup_JobLeadsModule (Job Leads Tracker Module) ---"); // NEW DEBUG LOG
    setupMessages.push("Job Leads Tracker Module: Setup function executed.");
  } catch (e) {
    Logger.log(`ERROR during Job Leads Tracker setup: ${e.toString()}\n${e.stack}`);
    setupMessages.push(`Job Leads Tracker Module: FAILED - ${e.message}`);
    overallSuccess = false;
  }

  Logger.log("\n--- CHECKPOINT: Proceeding to Final Summary? ---"); // NEW DEBUG LOG

  // --- 3. (Optional) Final UI Alert ---
  Logger.log("\n--- Finalizing Full Project Setup ---");
  const summaryMessage = `Full Project Initial Setup Summary:\n- ${setupMessages.join('\n- ')}\n\nOverall Status: ${overallSuccess ? "SUCCESSFUL (review logs for details)" : "ENCOUNTERED ISSUES (critical: review logs immediately!)"}`;
  Logger.log(summaryMessage);

  try {
    // Try to get UI. If running from editor without sheet context, this might not show.
    const ui = SpreadsheetApp.getUi();
    ui.alert(overallSuccess ? 'Full Setup Complete' : 'Full Setup Issues', summaryMessage, ui.ButtonSet.OK);
  } catch (e) {
    // Fallback for non-UI context (e.g., direct editor run)
    // Browser.msgBox can work here if you prefer a modal.
    // For now, relying on Logger.log for this case.
    Logger.log("Full setup UI alert skipped (no active spreadsheet UI or error). Summary logged above.");
  }

  Logger.log(`==== FULL PROJECT INITIAL SETUP ${overallSuccess ? "CONCLUDED" : "CONCLUDED WITH ISSUES"} (v2 Debug) ====\nReview all logs carefully.`);
}

// In Main.gs

/**
 * Adds a custom menu to the spreadsheet UI when the spreadsheet is opened.
 * This menu provides easy access to setup functions, processing tasks, and admin utilities.
 * @param {Object} e The event object (not typically used directly for onOpen menu creation).
 */
function onOpen(e) {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('⚙️ Master Job Manager'); // Top-level menu name

  // --- Full System Setup ---
  menu.addItem('▶️ RUN FULL PROJECT SETUP', 'runFullProjectInitialSetup');
  menu.addSeparator();

  // --- Individual Module Setups ---
  menu.addSubMenu(ui.createMenu('Module Setups')
      .addItem('Setup: Job Application Tracker', 'initialSetup_LabelsAndSheet')
      .addItem('Setup: Job Leads Tracker', 'runInitialSetup_JobLeadsModule'));
  menu.addSeparator();

  // --- Manual Processing Triggers ---
  menu.addSubMenu(ui.createMenu('Manual Processing')
      .addItem('📧 Process Application Emails', 'processJobApplicationEmails')
      .addItem('📬 Process Job Leads', 'processJobLeads')
      .addItem('🗑️ Mark Stale Applications', 'markStaleApplicationsAsRejected'));
  menu.addSeparator();

  // --- Configuration & Admin ---
  menu.addSubMenu(ui.createMenu('Admin & Config')
      .addItem('🔑 Set Shared Gemini API Key', 'setSharedGeminiApiKey_UI') // Assumes this is in AdminUtils.gs
      .addItem('🔍 Show All User Properties', 'showAllUserProperties')   // Assumes this is in AdminUtils.gs
      .addItem('🔩 TEMPORARY: Set Hardcoded API Key', 'TEMPORARY_manualSetSharedGeminiApiKey')); // Assumes this is in AdminUtils.gs
      // You could add items here to manually create/delete specific triggers if needed for advanced debugging.

  menu.addToUi();
}

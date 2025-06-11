// File: Leads_Main.gs
// Description: Contains the primary functions for the Job Leads Tracker module,
// including initial setup of the leads sheet/labels/filters and the
// ongoing processing of job lead emails.

/**
 * Sets up the Job Leads Tracker module:
 * - Ensures the "Potential Job Leads" sheet exists in the main spreadsheet and formats it.
 * - Creates necessary Gmail labels ("Master Job Manager/Job Application Potential/*").
 * - Creates a Gmail filter for "job alert" emails to apply the "NeedsProcess" label.
 * - Sets up a daily trigger for processing new job leads.
 * Designed to be run manually from the Apps Script editor once.
 */

// In Leads_Main.gs

/**
 * Sets up the Job Leads Tracker module:
 * - Ensures the "Potential Job Leads" sheet exists in the main spreadsheet and formats it.
 * - Creates necessary Gmail labels ("Master Job Manager/Job Application Potential/*").
 * - Creates a Gmail filter for "job alert" emails to apply the "NeedsProcess" label.
 * - Sets up a daily trigger for processing new job leads.
 * Designed to be run manually from the Apps Script editor once.
 */
function runInitialSetup_JobLeadsModule() {
  Logger.log("runInitialSetup_JobLeadsModule: User initiated from Apps Script editor. Proceeding with setup.");
  Logger.log(
      'This will:\n' +
      '1. Create/Verify the "Potential Job Leads" tab in your main spreadsheet (using Tab Name: "' + LEADS_SHEET_TAB_NAME + '").\n' + // From Config.gs
      '2. Style the "Potential Job Leads" tab with headers and formatting.\n' +
      '3. Create Gmail labels: "' + MASTER_GMAIL_LABEL_PARENT + '", "' + LEADS_GMAIL_LABEL_PARENT + '", "' + LEADS_GMAIL_LABEL_NEEDS_PROCESS + '", and "' + LEADS_GMAIL_LABEL_DONE_PROCESS + '".\n' + // From Config.gs
      '4. Create a Gmail filter for query: "' + LEADS_GMAIL_FILTER_QUERY + '" emails.\n' + // From Config.gs
      '5. Set up a daily trigger to automatically process new job leads (function: "processJobLeads").'
  );

  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch(e) { Logger.log("[LEADS_SETUP INFO] No Spreadsheet UI context available for alerts."); }

  // Define module-specific success flags if desired, or use a global one
  let leadsModuleSetupSuccess = true; // Example module-specific flag

  try {
    Logger.log('Starting runInitialSetup_JobLeadsModule core logic...');

    // --- Step 1: Get Main Spreadsheet & Leads Sheet Tab ---
    Logger.log(`[LEADS_SETUP INFO] Attempting to get or create the main spreadsheet...`);
    const { spreadsheet: mainSpreadsheet } = getOrCreateSpreadsheetAndSheet(); // From main SheetUtils.gs
    if (!mainSpreadsheet) {
      // This error will be caught by the main try-catch of runInitialSetup_JobLeadsModule
      throw new Error("CRITICAL (LEADS_SETUP): Could not get or create the main spreadsheet. Check logs for getOrCreateSpreadsheetAndSheet and Config.gs (FIXED_SPREADSHEET_ID or TARGET_SPREADSHEET_FILENAME).");
    }
    Logger.log(`[LEADS_SETUP INFO] MAIN SPREADSHEET successfully obtained: "${mainSpreadsheet.getName()}", ID: ${mainSpreadsheet.getId()}`);

    Logger.log(`[LEADS_SETUP INFO] Attempting to get or create the leads sheet tab: "${LEADS_SHEET_TAB_NAME}"`);
    let leadsSheet = mainSpreadsheet.getSheetByName(LEADS_SHEET_TAB_NAME); // LEADS_SHEET_TAB_NAME from Config.gs

    if (!leadsSheet) {
      Logger.log(`[LEADS_SETUP INFO] Leads sheet tab "${LEADS_SHEET_TAB_NAME}" not found. Attempting to create it...`);
      try {
        leadsSheet = mainSpreadsheet.insertSheet(LEADS_SHEET_TAB_NAME);
        if (leadsSheet) {
            Logger.log(`[LEADS_SETUP INFO] Successfully CREATED new tab: "${LEADS_SHEET_TAB_NAME}" in spreadsheet "${mainSpreadsheet.getName()}".`);
        } else {
            // This case should be rare if insertSheet doesn't throw an error but returns null.
            Logger.log(`[LEADS_SETUP ERROR] mainSpreadsheet.insertSheet("${LEADS_SHEET_TAB_NAME}") returned null or undefined. Sheet not created.`);
             // Error will be thrown by the check below.
        }
      } catch (e) {
        Logger.log(`[LEADS_SETUP ERROR] Error during mainSpreadsheet.insertSheet("${LEADS_SHEET_TAB_NAME}"): ${e.message}\nStack: ${e.stack}`);
        // Error will be caught by the check below and then the main try-catch.
      }
    } else {
      Logger.log(`[LEADS_SETUP INFO] Found EXISTING tab: "${LEADS_SHEET_TAB_NAME}" in spreadsheet "${mainSpreadsheet.getName()}".`);
    }

    // **Crucial check**: Ensure leadsSheet is now a valid sheet object
    if (!leadsSheet || typeof leadsSheet.getName !== 'function') {
        // This error will be caught by the main try-catch of runInitialSetup_JobLeadsModule
        throw new Error(`CRITICAL (LEADS_SETUP): Failed to obtain a valid sheet object for the leads sheet tab named "${LEADS_SHEET_TAB_NAME}" after attempting to find or create it.`);
    }
    Logger.log(`[LEADS_SETUP INFO] Successfully obtained leadsSheet object for tab: "${leadsSheet.getName()}". Proceeding with formatting.`);


    // --- Step 2: Setup Leads Sheet with Headers and Styling ---
    if (leadsSheet.getLastRow() === 0 || (leadsSheet.getLastRow() === 1 && leadsSheet.getLastColumn() <=1 && leadsSheet.getRange(1,1).isBlank())) {
        leadsSheet.clearContents();
        leadsSheet.clearFormats();
        Logger.log(`[LEADS_SETUP INFO] Cleared contents and formats for "${LEADS_SHEET_TAB_NAME}" as it appeared new/empty.`);
    } else {
        Logger.log(`[LEADS_SETUP INFO] "${LEADS_SHEET_TAB_NAME}" has existing content. Headers will be verified/added if missing. Styling will be applied.`);
    }

    if (leadsSheet.getLastRow() < 1 || leadsSheet.getRange(1,1).isBlank()) {
        leadsSheet.appendRow(LEADS_SHEET_HEADERS); // From Config.gs
        Logger.log(`[LEADS_SETUP INFO] Headers added to "${LEADS_SHEET_TAB_NAME}".`);
    } else {
        Logger.log(`[LEADS_SETUP INFO] Headers appear to exist in "${LEADS_SHEET_TAB_NAME}". Skipping appendRow.`);
    }

    const headerRange = leadsSheet.getRange(1, 1, 1, LEADS_SHEET_HEADERS.length); // From Config.gs
    headerRange.setFontWeight("bold").setHorizontalAlignment("center");
    leadsSheet.setFrozenRows(1);
    Logger.log("[LEADS_SETUP INFO] Header row styled and frozen for leads sheet.");

    const bandingRange = leadsSheet.getRange(1, 1, leadsSheet.getMaxRows(), LEADS_SHEET_HEADERS.length);
    try {
        const existingSheetBandings = leadsSheet.getBandings();
        for (let k = 0; k < existingSheetBandings.length; k++) existingSheetBandings[k].remove();
        bandingRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
        Logger.log("[LEADS_SETUP INFO] Applied LIGHT_GREY alternating row colors (banding) to leads sheet.");
    } catch (e) { Logger.log("[LEADS_SETUP WARN] Error applying banding to leads sheet: " + e.toString()); }

    const columnWidthsLeads = {
        "Date Added": 100, "Job Title": 220, "Company": 150, "Location": 130,
        "Source Email Subject": 220, "Link to Job Posting": 250, "Status": 80,
        "Source Email ID": 130, "Processed Timestamp": 100, "Notes": 250
    };
    for (let i = 0; i < LEADS_SHEET_HEADERS.length; i++) { // From Config.gs
        const headerName = LEADS_SHEET_HEADERS[i]; const columnIndex = i + 1;
        try {
            if (columnWidthsLeads[headerName]) {
                leadsSheet.setColumnWidth(columnIndex, columnWidthsLeads[headerName]);
            }
        }
        catch (e) { Logger.log(`[LEADS_SETUP WARN] Error setting column ${columnIndex} ("${headerName}") width for leads sheet: ${e.toString()}`); }
    }
    Logger.log("[LEADS_SETUP INFO] Set column widths for leads sheet.");

    const totalColumnsInLeadsSheet = leadsSheet.getMaxColumns();
    if (LEADS_SHEET_HEADERS.length < totalColumnsInLeadsSheet) { // From Config.gs
      try {
        leadsSheet.hideColumns(LEADS_SHEET_HEADERS.length + 1, totalColumnsInLeadsSheet - LEADS_SHEET_HEADERS.length);
        Logger.log(`[LEADS_SETUP INFO] Hid unused columns in leads sheet from column ${LEADS_SHEET_HEADERS.length + 1}.`);
      }
      catch (e) { Logger.log("[LEADS_SETUP WARN] Error hiding columns in leads sheet: " + e.toString()); }
    }
    Logger.log("[LEADS_SETUP INFO] Leads sheet styling applied.");


    // --- Step 3: Gmail Label and Filter Setup ---
    Logger.log(`[LEADS_SETUP INFO] Ensuring parent labels exist for leads module...`);
    getOrCreateLabel(MASTER_GMAIL_LABEL_PARENT); // From Config.gs
    Utilities.sleep(500);
    getOrCreateLabel(LEADS_GMAIL_LABEL_PARENT);  // From Config.gs
    Utilities.sleep(500);

    const needsProcessLabelNameConst = LEADS_GMAIL_LABEL_NEEDS_PROCESS; // From Config.gs
    const doneProcessLabelNameConst = LEADS_GMAIL_LABEL_DONE_PROCESS;   // From Config.gs

    const needsProcessLabelObject = getOrCreateLabel(needsProcessLabelNameConst);
    Utilities.sleep(500);
    const doneProcessLabelObject = getOrCreateLabel(doneProcessLabelNameConst);
    Utilities.sleep(500);
    Logger.log(`[LEADS_SETUP INFO] Called getOrCreateLabel for specific leads module labels ("${needsProcessLabelNameConst}", "${doneProcessLabelNameConst}").`);

    // --- Log inspection of the label objects directly after creation attempt ---
    if (needsProcessLabelObject) {
        Logger.log(`[DEBUG LEADS_SETUP] needsProcessLabelObject for "${needsProcessLabelNameConst}" is NOT null. Name: "${needsProcessLabelObject.getName ? needsProcessLabelObject.getName() : 'getName not function'}"`);
    } else {
        Logger.log(`[ERROR LEADS_SETUP] needsProcessLabelObject for "${needsProcessLabelNameConst}" IS NULL after getOrCreateLabel. Label might not have been created.`);
        leadsModuleSetupSuccess = false; // Label creation is critical
    }
    if (doneProcessLabelObject) {
        Logger.log(`[DEBUG LEADS_SETUP] doneProcessLabelObject for "${doneProcessLabelNameConst}" is NOT null. Name: "${doneProcessLabelObject.getName ? doneProcessLabelObject.getName() : 'getName not function'}"`);
    } else {
        Logger.log(`[ERROR LEADS_SETUP] doneProcessLabelObject for "${doneProcessLabelNameConst}" IS NULL after getOrCreateLabel. Label might not have been created.`);
        // leadsModuleSetupSuccess = false; // DoneProcess label also important
    }
    // --- End Log inspection ---

    let needsProcessLeadLabelId = null;

    if (needsProcessLabelObject) { // Only attempt to get ID if the label object seems to exist
        Logger.log(`[LEADS_SETUP INFO] Attempting to get Label ID for "${needsProcessLabelNameConst}" using Advanced Gmail Service.`);
        try {
            const advancedGmailService = Gmail; // Gmail Advanced Service
            if (!advancedGmailService || !advancedGmailService.Users || !advancedGmailService.Users.Labels) {
                throw new Error("Gmail API Advanced Service (Gmail) is not available or not properly enabled for LEADS filter setup.");
            }
            const labelsListResponse = advancedGmailService.Users.Labels.list('me');
            Logger.log(`[DEBUG LEADS_SETUP] Raw labelsListResponse from Gmail API (for leads 'NeedsProcess' ID lookup): ${JSON.stringify(labelsListResponse)}`);

            if (labelsListResponse.labels && labelsListResponse.labels.length > 0) {
                const targetLabelInfo = labelsListResponse.labels.find(l => l.name === needsProcessLabelNameConst);
                if (targetLabelInfo && targetLabelInfo.id) {
                    needsProcessLeadLabelId = targetLabelInfo.id;
                    Logger.log(`[LEADS_SETUP INFO] Successfully retrieved Label ID via Advanced Service for leads: "${needsProcessLeadLabelId}" for label "${needsProcessLabelNameConst}".`);
                } else {
                    Logger.log(`[LEADS_SETUP WARN] Label "${needsProcessLabelNameConst}" NOT FOUND in list from Advanced Gmail Service for leads. Expected it if getOrCreateLabel succeeded. Check logs. TargetLabelInfo: ${JSON.stringify(targetLabelInfo)}`);
                    leadsModuleSetupSuccess = false; // Cannot create filter without ID
                }
            } else {
                Logger.log(`[LEADS_SETUP WARN] No labels returned by Advanced Gmail Service list for user 'me' (leads 'NeedsProcess' ID lookup), or labels array is empty. Response: ${JSON.stringify(labelsListResponse)}`);
                leadsModuleSetupSuccess = false; // Cannot create filter without ID
            }
        } catch (e) {
            Logger.log(`[LEADS_SETUP ERROR] Error using Advanced Gmail Service to get label ID for "${needsProcessLabelNameConst}" (leads): ${e.message}\nStack: ${e.stack}`);
            leadsModuleSetupSuccess = false;
        }
    } else {
         Logger.log(`[LEADS_SETUP WARN] Skipping Advanced Gmail Service lookup for "${needsProcessLabelNameConst}" ID because the label object was not successfully created/retrieved by getOrCreateLabel earlier.`);
         leadsModuleSetupSuccess = false; // If NeedsProcess label wasn't created, ID fetch is pointless and setup fails
    }

    if (!needsProcessLeadLabelId) {
         const errorMsg = `CRITICAL (LEADS_SETUP): Could not obtain ID for Gmail label "${needsProcessLabelNameConst}". Filter creation for leads will fail.`;
         Logger.log(errorMsg); // Already logged details above
         // Throwing an error here will be caught by the main try-catch, setting leadsModuleSetupSuccess to false
         throw new Error(errorMsg);
    }
    Logger.log(`[LEADS_SETUP INFO] Using Gmail Label ID "${needsProcessLeadLabelId}" (obtained via Advanced Service) for "${needsProcessLabelNameConst}" for filter creation.`);

    // --- Filter Creation Logic for Leads ---
    Logger.log(`[LEADS_SETUP INFO] Proceeding to create/verify Gmail filter for leads...`);
    let leadsFilterCreationAttempted = false;
    try {
        const gmailApiServiceForLeadsFilter = Gmail;
        let leadsFilterExists = false;
        if (!gmailApiServiceForLeadsFilter || !gmailApiServiceForLeadsFilter.Users || !gmailApiServiceForLeadsFilter.Users.Settings || !gmailApiServiceForLeadsFilter.Users.Settings.Filters) {
            throw new Error("Gmail API Advanced Service (Gmail) is not available or not properly enabled for LEADS filter creation.");
        }
        const existingLeadsFiltersResponse = gmailApiServiceForLeadsFilter.Users.Settings.Filters.list('me');
        Logger.log(`[DEBUG LEADS_SETUP] Raw existingLeadsFiltersResponse from Gmail API (for leads filter check): ${JSON.stringify(existingLeadsFiltersResponse)}`);

        let existingLeadsFiltersList = [];
        if (existingLeadsFiltersResponse && existingLeadsFiltersResponse.filter && Array.isArray(existingLeadsFiltersResponse.filter)) {
            existingLeadsFiltersList = existingLeadsFiltersResponse.filter;
            Logger.log(`[DEBUG LEADS_SETUP] Successfully extracted ${existingLeadsFiltersList.length} filters from API response for leads.`);
        } else if (existingLeadsFiltersResponse && typeof existingLeadsFiltersResponse === 'object' && !existingLeadsFiltersResponse.hasOwnProperty('filter')) {
            Logger.log(`[INFO LEADS_SETUP] No 'filter' property found in existingLeadsFiltersResponse for leads (user likely has no filters).`);
        } else if (!existingLeadsFiltersResponse) {
             Logger.log(`[WARN LEADS_SETUP] existingLeadsFiltersResponse from Gmail API for leads was null or undefined.`);
        } else {
            Logger.log(`[WARN LEADS_SETUP] existingLeadsFiltersResponse.filter was not an array or structure was unexpected for leads. Response: ${JSON.stringify(existingLeadsFiltersResponse)}`);
        }

        if (existingLeadsFiltersList.length > 0) {
            for (const filterItem of existingLeadsFiltersList) {
                if (filterItem.criteria && filterItem.criteria.query === LEADS_GMAIL_FILTER_QUERY && // From Config.gs
                    filterItem.action && filterItem.action.addLabelIds && filterItem.action.addLabelIds.includes(needsProcessLeadLabelId)) {
                    leadsFilterExists = true;
                    break;
                }
            }
        }

        if (!leadsFilterExists) {
            leadsFilterCreationAttempted = true;
            const leadsFilterResource = {
                criteria: { query: LEADS_GMAIL_FILTER_QUERY }, // From Config.gs
                action: { addLabelIds: [needsProcessLeadLabelId], removeLabelIds: ['INBOX'] }
            };
            Logger.log(`[DEBUG LEADS_SETUP] Attempting to create filter for leads with resource: ${JSON.stringify(leadsFilterResource)}`);
            const createdLeadsFilterResponse = gmailApiServiceForLeadsFilter.Users.Settings.Filters.create(leadsFilterResource, 'me');
            if (createdLeadsFilterResponse && createdLeadsFilterResponse.id) {
                Logger.log(`[LEADS_SETUP INFO] Gmail filter for leads CREATED successfully. Filter ID: ${createdLeadsFilterResponse.id}`);
            } else {
                Logger.log(`[LEADS_SETUP ERROR] Gmail filter creation call for leads did NOT return a valid filter ID. Response: ${JSON.stringify(createdLeadsFilterResponse)}`);
                leadsModuleSetupSuccess = false;
            }
        } else {
            Logger.log(`[LEADS_SETUP INFO] Gmail filter for query "${LEADS_GMAIL_FILTER_QUERY}" and label ID "${needsProcessLeadLabelId}" (leads) ALREADY EXISTS.`);
        }
    } catch (e) {
        if (e.message && e.message.toLowerCase().includes("filter already exists")) {
            Logger.log(`[LEADS_SETUP WARN] Gmail filter (query: "${LEADS_GMAIL_FILTER_QUERY}", leads) likely already exists (API reported).`);
        } else {
            Logger.log(`[LEADS_SETUP ERROR] Error creating/checking Gmail filter for leads: ${e.toString()}\nStack: ${e.stack}. Attempted: ${leadsFilterCreationAttempted}`);
            leadsModuleSetupSuccess = false;
        }
    }
    Logger.log("[LEADS_SETUP INFO] Gmail label and filter setup (Step 3) for leads module completed checks.");

    // --- Step 4: Store Configuration (UserProperties) ---
    Logger.log(`[LEADS_SETUP INFO] Storing configuration to UserProperties...`);
    const userProps = PropertiesService.getUserProperties();

    if (needsProcessLeadLabelId) {
        userProps.setProperty(LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID, needsProcessLeadLabelId); // From Config.gs
        Logger.log(`[LEADS_SETUP INFO] Stored LEADS_USER_PROPERTY_NEEDS_PROCESS_LABEL_ID: ${needsProcessLeadLabelId}`);
    } else {
        Logger.log(`[LEADS_SETUP WARN] needsProcessLeadLabelId was not available to store in UserProperties. This is problematic.`);
        // leadsModuleSetupSuccess might already be false, but confirm.
        if (leadsModuleSetupSuccess) leadsModuleSetupSuccess = false;
    }

    const doneProcessLabelNameConstForProps = LEADS_GMAIL_LABEL_DONE_PROCESS; // From Config.gs
    let doneProcessLeadLabelId = null;
    if(doneProcessLabelObject) { // Only try to get ID if the 'Done' label object itself exists
        try {
            const advancedGmailService = Gmail;
            const labelsListResponse = advancedGmailService.Users.Labels.list('me');
            if (labelsListResponse.labels && labelsListResponse.labels.length > 0) {
                const targetLabelInfo = labelsListResponse.labels.find(l => l.name === doneProcessLabelNameConstForProps);
                if (targetLabelInfo && targetLabelInfo.id) {
                    doneProcessLeadLabelId = targetLabelInfo.id;
                    userProps.setProperty(LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID, doneProcessLeadLabelId); // From Config.gs
                    Logger.log(`[LEADS_SETUP INFO] Successfully retrieved and stored LEADS_USER_PROPERTY_DONE_PROCESS_LABEL_ID: ${doneProcessLeadLabelId} for "${doneProcessLabelNameConstForProps}".`);
                } else {
                    Logger.log(`[LEADS_SETUP WARN] Label "${doneProcessLabelNameConstForProps}" not found via Advanced Service for UserProperties storage (DoneProcess).`);
                }
            } else {
                Logger.log(`[LEADS_SETUP WARN] No labels returned by Advanced Gmail Service list (for DoneProcessLabel UserProperty).`);
            }
        } catch (e) {
            Logger.log(`[LEADS_SETUP ERROR] Error using Advanced Gmail Service for "${doneProcessLabelNameConstForProps}" ID for UserProperties: ${e.message}`);
        }
    } else {
        Logger.log(`[LEADS_SETUP WARN] Skipping Advanced Service ID fetch for "${doneProcessLabelNameConstForProps}" as its object wasn't obtained from getOrCreateLabel.`);
    }
    if (!doneProcessLeadLabelId && doneProcessLabelObject) { // If label was created but ID somehow not stored
         Logger.log(`[LEADS_SETUP WARN] doneProcessLeadLabelId for "${doneProcessLabelNameConstForProps}" could not be obtained or stored in UserProperties, though label object existed.`);
    }
    Logger.log('[LEADS_SETUP INFO] UserProperties storage attempt for Step 4 finished.');

    // --- Step 5: Create Time-Driven Trigger ---
    const triggerFunctionNameForLeads = 'processJobLeads';
    let triggerWasCreatedOrVerified = false;
    try {
        const existingTriggers = ScriptApp.getProjectTriggers();
        let triggerToModify = null;
        for (let i = 0; i < existingTriggers.length; i++) {
            if (existingTriggers[i].getHandlerFunction() === triggerFunctionNameForLeads) {
                triggerToModify = existingTriggers[i];
                break; // Found existing trigger
            }
        }

        if (triggerToModify) {
            // Optionally, delete and recreate, or just verify it's configured correctly
            // For simplicity here, let's delete and recreate to ensure config.
            // For a more advanced setup, you might check its current schedule.
            ScriptApp.deleteTrigger(triggerToModify);
            Logger.log(`[LEADS_SETUP INFO] Deleted existing trigger for ${triggerFunctionNameForLeads} to ensure correct configuration.`);
        }
        
        ScriptApp.newTrigger(triggerFunctionNameForLeads)
            .timeBased()
            .everyDays(1)
            .atHour(3) // Approx 3 AM in script's timezone
            .inTimezone(Session.getScriptTimeZone()) // Best practice to specify timezone
            .create();
        triggerWasCreatedOrVerified = true;
        Logger.log(`[LEADS_SETUP INFO] Successfully created/re-created a new daily trigger for ${triggerFunctionNameForLeads} to run around 3 AM.`);

    } catch (e) {
        Logger.log(`[LEADS_SETUP ERROR] Error creating/managing time-driven trigger for ${triggerFunctionNameForLeads}: ${e.toString()}`);
        leadsModuleSetupSuccess = false; // Trigger is important
    }

    // Final alert based on module's success
    if (leadsModuleSetupSuccess) {
        Logger.log('[LEADS_SETUP INFO] Initial setup for Job Leads Module appears to have completed successfully.');
        if (ui) ui.alert('Job Leads Module Setup Complete!', `The "Potential Job Leads" tab and associated Gmail configurations seem ready.\nA daily processing trigger for "${triggerFunctionNameForLeads}" has been ${triggerWasCreatedOrVerified ? 'created/verified' : 'checked (see logs)'}.`, ui.ButtonSet.OK);
    } else {
        Logger.log('[LEADS_SETUP ERROR] Initial setup for Job Leads Module encountered ISSUES. Please review logs carefully.');
        if (ui) ui.alert('Job Leads Module Setup Issues', 'The Job Leads Module setup encountered issues. Some features may not work. Please check Apps Script logs for details.', ui.ButtonSet.OK);
    }

  } catch (e) { // Catch critical errors from the main try block
    Logger.log(`CRITICAL OVERALL ERROR in Job Leads Module initial setup: ${e.toString()}\nStack: ${e.stack || 'No stack available'}`);
    if (ui) ui.alert('Critical Error During Leads Setup', `A critical error occurred: ${e.message || e}. Setup may be incomplete. Check Apps Script logs.`, ui.ButtonSet.OK);
    // Ensure this failure is also propagated if this function is part of a larger setup chain.
  }
}

/**
 * Processes emails labeled for job leads, extracts job information using Gemini,
 * and writes the data to the "Potential Job Leads" sheet.
 * Intended to be run by a time-driven trigger.
 */

function processJobLeads() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING JOB LEAD PROCESSING (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  const userProperties = PropertiesService.getUserProperties();
  // Use GEMINI_API_KEY_PROPERTY from Config.gs (shared with main tracker)
  const geminiApiKey = userProperties.getProperty(GEMINI_API_KEY_PROPERTY);

  // Get main spreadsheet ID (where "Potential Job Leads" tab resides)
  // This relies on the main tracker's config for spreadsheet access.
  const { spreadsheet: mainSpreadsheet } = getOrCreateSpreadsheetAndSheet(); // From main SheetUtils.gs

  if (!mainSpreadsheet) {
      Logger.log('[FATAL] Main spreadsheet not found for leads processing. Ensure FIXED_SPREADSHEET_ID or TARGET_SPREADSHEET_FILENAME in Config.gs is correct. Aborting.');
      return;
  }
  const targetSpreadsheetId = mainSpreadsheet.getId(); // Use the ID of the obtained main spreadsheet

  const needsProcessLabelName = LEADS_GMAIL_LABEL_NEEDS_PROCESS; // From Config.gs
  const doneProcessLabelName = LEADS_GMAIL_LABEL_DONE_PROCESS;   // From Config.gs

  if (!geminiApiKey) {
    Logger.log('[FATAL] Gemini API Key not found in UserProperties under key: "' + GEMINI_API_KEY_PROPERTY + '". Aborting job leads processing.');
    return;
  }
  if (!needsProcessLabelName || !doneProcessLabelName) {
    Logger.log('[FATAL] Gmail label configuration for leads is missing in Config.gs. Aborting job leads processing.');
    return;
  }
  Logger.log(`[INFO] Job Leads Config OK. API Key: ${geminiApiKey ? geminiApiKey.substring(0,5) + "..." : "NOT SET"}, Main SS ID: ${targetSpreadsheetId}`);

  // getSheetAndHeaderMapping_forLeads will be defined in Leads_SheetUtils.gs
  const { sheet: dataSheet, headerMap } = getSheetAndHeaderMapping_forLeads(targetSpreadsheetId, LEADS_SHEET_TAB_NAME); // LEADS_SHEET_TAB_NAME from Config.gs
  if (!dataSheet || !headerMap || Object.keys(headerMap).length === 0) {
    Logger.log(`[FATAL] Leads sheet "${LEADS_SHEET_TAB_NAME}" or its headers not found in spreadsheet ID ${targetSpreadsheetId}. Aborting job leads processing.`);
    return;
  }

  const needsProcessLabel = GmailApp.getUserLabelByName(needsProcessLabelName);
  const doneProcessLabel = GmailApp.getUserLabelByName(doneProcessLabelName);
  if (!needsProcessLabel) {
      Logger.log(`[FATAL] Gmail label "${needsProcessLabelName}" not found. Aborting job leads processing.`);
      return;
  }
  if (!doneProcessLabel) {
      // This is less critical for processing, but good for completion.
      Logger.log(`[WARN] Gmail label "${doneProcessLabelName}" not found. Processing will continue, but threads may not be re-labeled correctly.`);
  }

  // getProcessedEmailIdsFromSheet_forLeads will be in Leads_SheetUtils.gs

  const processedEmailIds = getProcessedEmailIdsFromSheet_forLeads(dataSheet, headerMap);
  Logger.log(`[INFO] Preloaded ${processedEmailIds.size} email IDs already processed for leads from sheet "${dataSheet.getName()}".`);

  const LEADS_THREAD_LIMIT = 10; // Consider moving to Config.gs
  const LEADS_MESSAGE_LIMIT_PER_RUN = 15; // Consider moving to Config.gs
  let messagesProcessedThisRun = 0;
  const threads = needsProcessLabel.getThreads(0, LEADS_THREAD_LIMIT);
  Logger.log(`[INFO] Found ${threads.length} threads in "${needsProcessLabelName}".`);

  for (const thread of threads) {
    if (messagesProcessedThisRun >= LEADS_MESSAGE_LIMIT_PER_RUN) {
      Logger.log(`[INFO] Leads message limit (${LEADS_MESSAGE_LIMIT_PER_RUN}) reached for this run.`);
      break;
    }
    if ((new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000 > 320) { // Approx 5m20s, adjust as needed
      Logger.log(`[WARN] Time limit approaching for leads processing. Stopping loop.`);
      break;
    }

    const messages = thread.getMessages();
    let threadContainedUnprocessedMessages = false;
    let allMessagesInThreadSuccessfullyProcessedThisRun = true; // Assume success until a message fails

    for (const message of messages) {
      if (messagesProcessedThisRun >= LEADS_MESSAGE_LIMIT_PER_RUN) break;
      const msgId = message.getId();

      if (processedEmailIds.has(msgId)) {
        // Logger.log(`[DEBUG] Msg ${msgId} in thread ${thread.getId()} already fully processed for leads. Skipping.`);
        continue;
      }
      threadContainedUnprocessedMessages = true; // Found at least one message not in our processed ID set

      Logger.log(`\n--- Processing Lead Msg ID: ${msgId}, Subject: "${message.getSubject()}" ---`);
      messagesProcessedThisRun++;
      let currentMessageProcessedSuccessfully = false; // Flag for current message's outcome

      try {
        let emailBody = message.getPlainBody();
        if (typeof emailBody !== 'string' || emailBody.trim() === "") {
          Logger.log(`[WARN] Msg ${msgId}: Invalid or empty body for leads. Skipping Gemini call for this message.`);
          currentMessageProcessedSuccessfully = true; // Successfully determined there's nothing to parse
          // No error entry needed here, it's just an empty email.
          continue;
        }

        // callGemini_forJobLeads and parseGeminiResponse_forJobLeads will be in GeminiService.gs
        const geminiApiResponse = callGemini_forJobLeads(emailBody, geminiApiKey);

        if (geminiApiResponse && geminiApiResponse.success) {
          const extractedJobsArray = parseGeminiResponse_forJobLeads(geminiApiResponse.data); // Pass raw API data

          if (extractedJobsArray && extractedJobsArray.length > 0) {
            Logger.log(`[INFO] Gemini extracted ${extractedJobsArray.length} potential job listings from msg ${msgId}.`);
            let atLeastOneGoodJobWrittenThisMessage = false;
            for (const jobData of extractedJobsArray) {
              if (jobData && jobData.jobTitle && jobData.jobTitle.toLowerCase() !== 'n/a' && jobData.jobTitle.toLowerCase() !== 'error') {
                // Prepare jobData for sheet writing
                jobData.dateAdded = new Date();
                jobData.sourceEmailSubject = message.getSubject();
                jobData.sourceEmailId = msgId;
                jobData.status = "New"; // Default status for new leads
                jobData.processedTimestamp = new Date();
                // writeJobDataToSheet_forLeads will be in Leads_SheetUtils.gs
                writeJobDataToSheet_forLeads(dataSheet, jobData, headerMap);
                atLeastOneGoodJobWrittenThisMessage = true;
              } else {
                Logger.log(`[INFO] A job object from msg ${msgId} was N/A or error. Skipping write for this specific item: ${JSON.stringify(jobData)}`);
              }
            }
            if (atLeastOneGoodJobWrittenThisMessage) {
                currentMessageProcessedSuccessfully = true;
            } else {
                // No good jobs were extracted, but Gemini call was successful.
                // Could be Gemini correctly found no jobs, or all were N/A.
                Logger.log(`[INFO] Msg ${msgId}: Gemini call successful, but no valid job listings written (all N/A or empty array).`);
                currentMessageProcessedSuccessfully = true; // Still consider message processed.
            }
          } else { // Gemini call success, but extractedJobsArray is null or empty
            Logger.log(`[INFO] Msg ${msgId}: Gemini call successful, but parsing yielded no job listings array or it was empty.`);
            // This could be a valid case where the email contains no jobs.
            currentMessageProcessedSuccessfully = true; // Consider the message processed.
          }
        } else { // Gemini API call failed
          Logger.log(`[ERROR] Gemini API call FAILED for msg ${msgId}. Details: ${geminiApiResponse ? geminiApiResponse.error : 'Response object was null'}`);
          // writeErrorEntryToSheet_forLeads will be in Leads_SheetUtils.gs
          writeErrorEntryToSheet_forLeads(dataSheet, message, "Gemini API Call/Parse Failed", geminiApiResponse ? geminiApiResponse.error : "Unknown API error", headerMap);
          allMessagesInThreadSuccessfullyProcessedThisRun = false; // Mark thread as having an issue
        }
      } catch (e) {
        Logger.log(`[FATAL SCRIPT ERROR] Processing msg ${msgId} for leads: ${e.toString()}\nStack: ${e.stack}`);
        // writeErrorEntryToSheet_forLeads will be in Leads_SheetUtils.gs
        writeErrorEntryToSheet_forLeads(dataSheet, message, "Script error during lead processing", e.toString(), headerMap);
        allMessagesInThreadSuccessfullyProcessedThisRun = false; // Mark thread as having an issue
      }

      if (currentMessageProcessedSuccessfully) {
        // Optionally add msgId to a temporary set for this run if you want to ensure labels are only changed if *all* new messages in a thread are fine.
        // For now, relying on allMessagesInThreadSuccessfullyProcessedThisRun
      }
      Utilities.sleep(1500 + Math.floor(Math.random() * 1000)); // Pause between messages
    } // End loop over messages in a thread

    // After processing all messages in a thread (or hitting limit):
    if (threadContainedUnprocessedMessages && allMessagesInThreadSuccessfullyProcessedThisRun) {
      if (doneProcessLabel) { // Check if doneProcessLabel was found
        thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
        Logger.log(`[INFO] Thread ID ${thread.getId()} successfully processed for leads and moved to "${doneProcessLabelName}".`);
        // Add all message IDs from this thread to our main processedEmailIds set to prevent re-processing in future runs.
        messages.forEach(m => processedEmailIds.add(m.getId()));
      } else {
        thread.removeLabel(needsProcessLabel); // Remove from "NeedsProcess" even if "DoneProcess" is missing
        Logger.log(`[WARN] Thread ID ${thread.getId()} processed for leads. Removed from "${needsProcessLabelName}", but "${doneProcessLabelName}" label was not found to apply.`);
      }
    } else if (threadContainedUnprocessedMessages) { // Some messages failed
      Logger.log(`[WARN] Thread ID ${thread.getId()} had processing issues with one or more messages. NOT moved from "${needsProcessLabelName}". Will be retried next run.`);
    } else if (!threadContainedUnprocessedMessages && messages.length > 0) { // Thread had messages, but all were already processed
        if (doneProcessLabel) {
            thread.removeLabel(needsProcessLabel).addLabel(doneProcessLabel);
            Logger.log(`[INFO] Thread ID ${thread.getId()} contained only previously processed lead messages. Ensured it is in "${doneProcessLabelName}".`);
        } else {
            thread.removeLabel(needsProcessLabel);
            Logger.log(`[INFO] Thread ID ${thread.getId()} contained only previously processed lead messages. Removed from "${needsProcessLabelName}".`);
        }
    } else { // Thread was empty or became empty
      Logger.log(`[INFO] Thread ID ${thread.getId()} appears empty or all its messages were skipped. Removing from "${needsProcessLabelName}".`);
      try { thread.removeLabel(needsProcessLabel); }
      catch(e) { Logger.log(`[DEBUG] Minor error removing label from (likely) already unlabelled/empty thread ${thread.getId()}: ${e}`);}
    }
    Utilities.sleep(500); // Pause between threads
  } // End loop over threads

  Logger.log(`\n==== JOB LEAD PROCESSING FINISHED (${new Date().toLocaleString()}) === Total Time: ${(new Date().getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}

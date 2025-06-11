// File: SheetUtils.gs
// Description: Contains utility functions for Google Sheets interaction,
// including sheet creation, formatting, and data access setup.

/**
 * Converts a 1-based column index to its letter representation (e.g., 1 -> A, 27 -> AA).
 * @param {number} column The 1-based column index.
 * @return {string} The column letter(s).
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Sets up specific formatting for the main "Applications" data sheet.
 * This includes headers, column widths, frozen rows, banding, etc.
 * It specifically checks if the provided sheet is the "Applications" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet object to format.
 */
function setupSheetFormatting(sheet) {
  // Robust check for valid sheet object at the very beginning
  if (!sheet || typeof sheet.getName !== 'function') {
    Logger.log(`[SHEET_UTIL ERROR] SETUP_FORMAT: Invalid sheet object passed. Parameter was: ${sheet}, Type: ${typeof sheet}`);
    return;
  }
  // Logger.log(`[SHEET_UTIL DEBUG] SETUP_FORMAT: Entered with sheet named: "${sheet.getName()}". Validating if it's the applications data sheet.`);

  // This formatting is ONLY for the APP_TRACKER_SHEET_TAB_NAME (e.g., "Applications")
  if (sheet.getName() !== APP_TRACKER_SHEET_TAB_NAME) { // Uses APP_TRACKER_SHEET_TAB_NAME from Config.gs
    if (DEBUG_MODE) Logger.log(`[SHEET_UTIL DEBUG] SETUP_FORMAT: Skipping specific application data sheet formatting for a non-matching tab: "${sheet.getName()}". Expected: "${APP_TRACKER_SHEET_TAB_NAME}".`);
    return;
  }
  Logger.log(`[SHEET_UTIL INFO] SETUP_FORMAT: Applying formatting to application data sheet: "${sheet.getName()}".`);

  // Only apply detailed header/column setup if the sheet is truly new/empty
  if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) {
    Logger.log(`[SHEET_UTIL INFO] SETUP_FORMAT: Sheet "${sheet.getName()}" is new/empty. Applying detailed formatting (headers, widths, etc.).`);
    
    // Define headers for the "Applications" sheet
    let headers = new Array(TOTAL_COLUMNS_IN_APP_SHEET).fill(''); // TOTAL_COLUMNS_IN_APP_SHEET from Config.gs
    headers[PROCESSED_TIMESTAMP_COL - 1] = "Processed Timestamp"; // Constants from Config.gs
    headers[EMAIL_DATE_COL - 1] = "Email Date";
    headers[PLATFORM_COL - 1] = "Platform";
    headers[COMPANY_COL - 1] = "Company Name";
    headers[JOB_TITLE_COL - 1] = "Job Title";
    headers[STATUS_COL - 1] = "Status";
    headers[PEAK_STATUS_COL - 1] = "Peak Status";
    headers[LAST_UPDATE_DATE_COL - 1] = "Last Update Email Date";
    headers[EMAIL_SUBJECT_COL - 1] = "Email Subject";
    headers[EMAIL_LINK_COL - 1] = "Email Link";
    headers[EMAIL_ID_COL - 1] = "Email ID";
    
    try { sheet.appendRow(headers); }
    catch (e) { Logger.log(`[SHEET_UTIL ERROR] SETUP_FORMAT: Failed to append header row to "${sheet.getName()}": ${e}`); return; }

    const headerRange = sheet.getRange(1, 1, 1, TOTAL_COLUMNS_IN_APP_SHEET);
    try {
        headerRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        sheet.setRowHeight(1, 40); // Set header row height
    } catch(e) { Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Error setting up header style for "${sheet.getName()}": ${e}`);}

    // Set column widths for "Applications" sheet
    try {
      sheet.setColumnWidth(PROCESSED_TIMESTAMP_COL, 160); sheet.setColumnWidth(EMAIL_DATE_COL, 120); sheet.setColumnWidth(PLATFORM_COL, 100); sheet.setColumnWidth(COMPANY_COL, 200); sheet.setColumnWidth(JOB_TITLE_COL, 250); sheet.setColumnWidth(STATUS_COL, 150);
      sheet.setColumnWidth(PEAK_STATUS_COL, 150); sheet.setColumnWidth(LAST_UPDATE_DATE_COL, 160); sheet.setColumnWidth(EMAIL_SUBJECT_COL, 300); sheet.setColumnWidth(EMAIL_LINK_COL, 100); sheet.setColumnWidth(EMAIL_ID_COL, 200);
    } catch (e) { Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Could not set column widths for "${sheet.getName()}": ${e}`); }

    // Default formatting for data rows
    const numDataRowsToFormat = sheet.getMaxRows() > 1 ? sheet.getMaxRows() - 1 : 1000; // Default to 1000 if getMaxRows is small
    if (numDataRowsToFormat > 0) {
        const allDataRange = sheet.getRange(2, 1, numDataRowsToFormat, TOTAL_COLUMNS_IN_APP_SHEET);
        try { allDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top'); }
        catch(e) { Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Error setting data range wrap/align for "${sheet.getName()}": ${e}`);}
        
        try { sheet.setRowHeightsForced(2, numDataRowsToFormat, 30); } // Use setRowHeightsForced for better consistency
        catch (e) { Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Could not set default row heights for "${sheet.getName()}": ${e}`); }
        
        if (EMAIL_LINK_COL > 0 && TOTAL_COLUMNS_IN_APP_SHEET >= EMAIL_LINK_COL) {
            try{ sheet.getRange(2, EMAIL_LINK_COL, numDataRowsToFormat, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP); }
            catch(e){ Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Email Link Col CLIP error for "${sheet.getName()}": ${e}`);}
        }
        // Apply banding
        try{
            const banding = sheet.getRange(1, 1, sheet.getMaxRows(), TOTAL_COLUMNS_IN_APP_SHEET) // Apply to whole sheet for future rows
                                 .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false); // Header true, Footer false
            banding.setHeaderRowColor("#D0E4F5") // Example light blue for header band
                   .setFirstRowColor("#FFFFFF") // White for first data band
                   .setSecondRowColor("#F3F3F3"); // Light grey for second data band
        }
        catch(e){ Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Banding error for "${sheet.getName()}": ${e}`);}
    }
    // Hide Peak Status Column and unused columns AFTER all other formatting for new sheet
    try { sheet.hideColumns(PEAK_STATUS_COL); Logger.log(`[SHEET_UTIL INFO] SETUP_FORMAT: Hid column ${PEAK_STATUS_COL} (Peak Status) for new sheet "${sheet.getName()}".`); }
    catch (e) { Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Could not hide Peak Status column for new sheet "${sheet.getName()}": ${e}`); }

    const maxSheetCols = sheet.getMaxColumns();
    if (maxSheetCols > TOTAL_COLUMNS_IN_APP_SHEET) {
        try { sheet.hideColumns(TOTAL_COLUMNS_IN_APP_SHEET + 1, maxSheetCols - TOTAL_COLUMNS_IN_APP_SHEET); }
        catch(e){ Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Hide unused columns error for new sheet "${sheet.getName()}": ${e}`); }
    }
    Logger.log(`[SHEET_UTIL INFO] SETUP_FORMAT: Detailed formatting attempt complete for new data sheet "${sheet.getName()}".`);
  } else { // Sheet already has content, just ensure critical formats
    if(DEBUG_MODE)Logger.log(`[SHEET_UTIL DEBUG] SETUP_FORMAT: Sheet "${sheet.getName()}" has content. Ensuring key formats.`);
    // Ensure Email Link column is CLIP
    if(EMAIL_LINK_COL > 0 && TOTAL_COLUMNS_IN_APP_SHEET >= EMAIL_LINK_COL && sheet.getLastRow() > 1){
        try{
            const emailLinkColRange = sheet.getRange(2, EMAIL_LINK_COL, sheet.getLastRow()-1, 1);
            // Check current strategy to avoid unnecessary calls
            if(emailLinkColRange.getWrapStrategies()[0][0] !== SpreadsheetApp.WrapStrategy.CLIP){
                emailLinkColRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
            }
        }
        catch(e){Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Error setting CLIP on existing email link column for "${sheet.getName()}": ${e}`);}
    }
    // Ensure Peak Status column is hidden
    if (PEAK_STATUS_COL > 0 && PEAK_STATUS_COL <= sheet.getMaxColumns()) {
        if (!sheet.isColumnHiddenByUser(PEAK_STATUS_COL)) {
          try { sheet.hideColumns(PEAK_STATUS_COL); Logger.log(`[SHEET_UTIL INFO] SETUP_FORMAT: Ensured Peak Status column is hidden on existing sheet "${sheet.getName()}".`); }
          catch(e) { Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Failed to hide Peak Status on existing sheet "${sheet.getName()}": ${e}`);}
        }
    }
  }
  // Frozen row should always be set for the "Applications" sheet
  try { if(sheet.getFrozenRows() < 1) sheet.setFrozenRows(1); }
  catch(e){ Logger.log(`[SHEET_UTIL WARN] SETUP_FORMAT: Could not set frozen rows on sheet "${sheet.getName()}": ${e}`); }
}


/**
 * Gets or creates the main project spreadsheet and the primary "Applications" data sheet tab.
 * It uses FIXED_SPREADSHEET_ID or TARGET_SPREADSHEET_FILENAME from Config.gs.
 * The returned 'sheet' property specifically refers to the "Applications" data sheet.
 * @return {{spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet | null, sheet: GoogleAppsScript.Spreadsheet.Sheet | null}}
 */
function getOrCreateSpreadsheetAndSheet() {
  let ss = null;
  let appTrackerSheet = null; // Specifically for the "Applications" sheet

  // --- Spreadsheet Access/Creation ---
  if (FIXED_SPREADSHEET_ID && FIXED_SPREADSHEET_ID.trim() !== "" && FIXED_SPREADSHEET_ID !== "YOUR_SPREADSHEET_ID_HERE") { // From Config.gs
    Logger.log(`[SHEET_UTIL INFO] SPREADSHEET: Attempting to open by Fixed ID: "${FIXED_SPREADSHEET_ID}"`); // More precise log
    try {
      ss = SpreadsheetApp.openById(FIXED_SPREADSHEET_ID);
      Logger.log(`[SHEET_UTIL INFO] SPREADSHEET: Successfully opened "${ss.getName()}" using Fixed ID.`);
    } catch (e) {
      Logger.log(`[SHEET_UTIL FATAL] SPREADSHEET: FIXED ID FAIL - Could not open spreadsheet with ID "${FIXED_SPREADSHEET_ID}". ${e.message}.`);
      return { spreadsheet: null, sheet: null }; // Critical failure
    }
  } else {
    Logger.log(`[SHEET_UTIL INFO] SPREADSHEET: Fixed ID not set or invalid. Attempting to find/create sheet by name: "${TARGET_SPREADSHEET_FILENAME}".`); // TARGET_SPREADSHEET_FILENAME from Config.gs
    try {
      const files = DriveApp.getFilesByName(TARGET_SPREADSHEET_FILENAME);
      if (files.hasNext()) {
        const file = files.next();
        ss = SpreadsheetApp.open(file);
        Logger.log(`[SHEET_UTIL INFO] SPREADSHEET: Found and opened existing spreadsheet by name: "${ss.getName()}" (ID: ${ss.getId()}).`);
        if (files.hasNext()) {
          Logger.log(`[SHEET_UTIL WARN] SPREADSHEET: Multiple files found with the name "${TARGET_SPREADSHEET_FILENAME}". Used the first one.`);
        }
      } else {
        Logger.log(`[SHEET_UTIL INFO] SPREADSHEET: No spreadsheet found by name "${TARGET_SPREADSHEET_FILENAME}". Attempting to create a new one.`);
        try {
          ss = SpreadsheetApp.create(TARGET_SPREADSHEET_FILENAME);
          Logger.log(`[SHEET_UTIL INFO] SPREADSHEET: Successfully created new spreadsheet: "${ss.getName()}" (ID: ${ss.getId()}).`);
        } catch (eCreate) {
          Logger.log(`[SHEET_UTIL FATAL] SPREADSHEET: CREATE FAIL - Failed to create new spreadsheet named "${TARGET_SPREADSHEET_FILENAME}". ${eCreate.message}.`);
          return { spreadsheet: null, sheet: null }; // Critical failure
        }
      }
    } catch (eDrive) {
      // This catch block is for errors during DriveApp.getFilesByName or SpreadsheetApp.open(file)
      Logger.log(`[SHEET_UTIL FATAL] SPREADSHEET: DRIVE/OPEN FAIL - Error interacting with Google Drive for "${TARGET_SPREADSHEET_FILENAME}". ${eDrive.message}. Ensure Drive API scope if necessary, or check filename.`);
      return { spreadsheet: null, sheet: null }; // Critical failure
    }
  }

  // --- "Applications" Sheet Tab (APP_TRACKER_SHEET_TAB_NAME) Access/Creation ---
  // This part is specific to the "Applications" sheet, not the "Leads" sheet
  if (ss) { // Only proceed if 'ss' was successfully obtained
    Logger.log(`[SHEET_UTIL INFO] SPREADSHEET is valid. Now checking for sheet tab: "${APP_TRACKER_SHEET_TAB_NAME}"`); // From Config.gs
    appTrackerSheet = ss.getSheetByName(APP_TRACKER_SHEET_TAB_NAME);
    if (!appTrackerSheet) {
      Logger.log(`[SHEET_UTIL INFO] TAB: Main application data sheet "${APP_TRACKER_SHEET_TAB_NAME}" not found in "${ss.getName()}". Creating...`);
      try {
        appTrackerSheet = ss.insertSheet(APP_TRACKER_SHEET_TAB_NAME);
        Logger.log(`[SHEET_UTIL INFO] TAB: Created main application data sheet "${APP_TRACKER_SHEET_TAB_NAME}".`);
        // Call setupSheetFormatting for the newly created "Applications" sheet
        setupSheetFormatting(appTrackerSheet); // From SheetUtils.gs (or wherever it is)
      } catch (eTabCreate) {
        Logger.log(`[SHEET_UTIL FATAL] TAB CREATE FAIL: Error creating main application data tab "${APP_TRACKER_SHEET_TAB_NAME}" in "${ss.getName()}". ${eTabCreate.message}.`);
        // Do not return 'ss' if its key sheet cannot be made.
        return { spreadsheet: ss, sheet: null }; // Return spreadsheet, but null for appTrackerSheet
      }
    } else {
      Logger.log(`[SHEET_UTIL INFO] TAB: Found existing main application data sheet "${APP_TRACKER_SHEET_TAB_NAME}".`);
      setupSheetFormatting(appTrackerSheet); // Ensure formatting for existing "Applications" sheet
    }

    // Attempt to delete the default "Sheet1" logic - Moved AFTER appTrackerSheet creation/verification
    // Ensure appTrackerSheet is valid before comparing its ID
    if (appTrackerSheet && typeof appTrackerSheet.getSheetId === 'function') {
        const defaultSheet = ss.getSheetByName('Sheet1');
        if (defaultSheet && defaultSheet.getSheetId() !== appTrackerSheet.getSheetId()) {
            let isKeySheet = false;
            const keySheetNames = [APP_TRACKER_SHEET_TAB_NAME, DASHBOARD_TAB_NAME, HELPER_SHEET_NAME, LEADS_SHEET_TAB_NAME]; // All key tabs from Config.gs
            for (const name of keySheetNames) {
                const sheetToCheck = ss.getSheetByName(name); // Use 'ss' here consistently
                if (sheetToCheck && sheetToCheck.getSheetId() === defaultSheet.getSheetId()) {
                    isKeySheet = true;
                    break;
                }
            }
            if (!isKeySheet && ss.getSheets().length > 1) {
                 try { ss.deleteSheet(defaultSheet); Logger.log(`[SHEET_UTIL INFO] TAB: Removed default 'Sheet1' as it was not a recognized key tab and not the only sheet.`); }
                 catch (eDeleteDefault) { Logger.log(`[SHEET_UTIL WARN] TAB: Failed to remove default 'Sheet1': ${eDeleteDefault.message}`); }
            }
        }
    } else if (ss && !appTrackerSheet) { // If ss is valid but appTrackerSheet is not
        Logger.log(`[SHEET_UTIL WARN] TAB: Could not verify/create the primary application data sheet "${APP_TRACKER_SHEET_TAB_NAME}". Default 'Sheet1' cleanup skipped.`);
    }

  } else { // This else corresponds to if (ss) for appTrackerSheet logic
    Logger.log(`[SHEET_UTIL FATAL] SPREADSHEET: Spreadsheet object ('ss') is null or invalid BEFORE attempting to get/create "${APP_TRACKER_SHEET_TAB_NAME}". Cannot proceed with application data sheet setup.`);
    // This case implies FIXED_SPREADSHEET_ID and TARGET_SPREADSHEET_FILENAME methods both failed.
    return { spreadsheet: null, sheet: null }; // Already handled earlier, but being explicit.
  }

  // The 'sheet' returned by this function is specifically the APP_TRACKER_SHEET_TAB_NAME sheet.
  return { spreadsheet: ss, sheet: appTrackerSheet };
}

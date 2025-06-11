// File: Leads_SheetUtils.gs
// Description: Contains utility functions specific to the "Potential Job Leads" sheet,
// such as header mapping, retrieving processed email IDs, and writing job lead data or errors...

/**
 * Retrieves a specific sheet by name from a given spreadsheet ID and maps its header names to column numbers.
 * @param {string} ssId The ID of the spreadsheet.
 * @param {string} sheetName The name of the sheet to retrieve.
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet | null, headerMap: Object}} An object containing the sheet and its header map.
 */

function getSheetAndHeaderMapping_forLeads(ssId, sheetName) {
  try {
    const ss = SpreadsheetApp.openById(ssId);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`[LEADS_SHEET_UTIL ERROR] Sheet "${sheetName}" not found in Spreadsheet ID "${ssId}".`);
      return { sheet: null, headerMap: {} };
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => {
      if (h && h.toString().trim() !== "") {
        headerMap[h.toString().trim()] = i + 1; // 1-based column index
      }
    });
    if (Object.keys(headerMap).length === 0) {
      Logger.log(`[LEADS_SHEET_UTIL WARN] No headers found in sheet "${sheetName}". An empty headerMap will be returned.`);
    } else {
      // Logger.log(`[LEADS_SHEET_UTIL INFO] Headers for "${sheetName}": ${JSON.stringify(headerMap)}`);
    }
    return { sheet: sheet, headerMap: headerMap };
  } catch (e) {
    Logger.log(`[LEADS_SHEET_UTIL ERROR] Error opening spreadsheet/sheet or getting headers. SS_ID: ${ssId}, SheetName: ${sheetName}. Error: ${e.toString()}`);
    return { sheet: null, headerMap: {} };
  }
}

/**
 * Retrieves a set of all unique email IDs from the "Source Email ID" column of the leads sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Potential Job Leads" sheet object.
 * @param {Object} headerMap An object mapping header names to column numbers for the sheet.
 * @return {Set<string>} A set of processed email IDs.
 */

function getProcessedEmailIdsFromSheet_forLeads(sheet, headerMap) {
  const ids = new Set();
  const emailIdColHeader = "Source Email ID"; // This should match a header in LEADS_SHEET_HEADERS

  if (!sheet) {
    Logger.log('[LEADS_SHEET_UTIL WARN] Cannot get processed IDs: sheet object is null.');
    return ids;
  }
  if (!headerMap || !headerMap[emailIdColHeader]) {
    Logger.log(`[LEADS_SHEET_UTIL WARN] Cannot get processed IDs: "${emailIdColHeader}" not found in headerMap for sheet "${sheet.getName()}". HeaderMap: ${JSON.stringify(headerMap)}`);
    return ids;
  }

  const emailIdColNum = headerMap[emailIdColHeader];
  const lastR = sheet.getLastRow();
  if (lastR < 2) { // No data rows
    return ids;
  }

  try {
    const rangeToRead = sheet.getRange(2, emailIdColNum, lastR - 1, 1);
    const emailIdValues = rangeToRead.getValues();
    emailIdValues.forEach(row => {
      if (row[0] && row[0].toString().trim() !== "") {
        ids.add(row[0].toString().trim());
      }
    });
  } catch (e) {
    Logger.log(`[LEADS_SHEET_UTIL ERROR] Error reading email IDs from column ${emailIdColNum} in sheet "${sheet.getName()}": ${e.toString()}`);
  }
  return ids;
}

/**
 * Appends a new row with job lead data to the "Potential Job Leads" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Potential Job Leads" sheet object.
 * @param {Object} jobData An object containing the job lead details.
 * @param {Object} headerMap An object mapping header names to column numbers for the sheet.
 */

function writeJobDataToSheet_forLeads(sheet, jobData, headerMap) {
  if (!sheet) { Logger.log("[LEADS_SHEET_UTIL ERROR] Cannot write job data: sheet object is null."); return; }
  if (!headerMap || Object.keys(headerMap).length === 0) { Logger.log("[LEADS_SHEET_UTIL ERROR] Cannot write job data: headerMap is invalid or empty."); return; }

  // Determine the number of columns based on the headerMap or LEADS_SHEET_HEADERS
  const numCols = LEADS_SHEET_HEADERS.length; // Use the definitive list of headers from Config.gs
  const newRow = new Array(numCols).fill("");

  // Helper to safely get data or default
  const getData = (propertyName, defaultValue = "") => jobData[propertyName] !== undefined && jobData[propertyName] !== null ? jobData[propertyName] : defaultValue;

  // Map jobData properties to columns based on headerMap and LEADS_SHEET_HEADERS
  // This ensures data is written to the correct column even if header order changes,
  // as long as LEADS_SHEET_HEADERS in Config.gs is the source of truth for column order.
  for (let i = 0; i < LEADS_SHEET_HEADERS.length; i++) {
      const headerName = LEADS_SHEET_HEADERS[i];
      if (headerMap[headerName]) { // Check if this header is expected and mapped
          switch (headerName) {
              case "Date Added":            newRow[i] = getData('dateAdded', new Date()); break;
              case "Job Title":             newRow[i] = getData('jobTitle', "N/A"); break;
              case "Company":               newRow[i] = getData('company', "N/A"); break;
              case "Location":              newRow[i] = getData('location', "N/A"); break;
              case "Source Email Subject":  newRow[i] = getData('sourceEmailSubject'); break;
              case "Link to Job Posting":   newRow[i] = getData('linkToJobPosting', "N/A"); break;
              case "Status":                newRow[i] = getData('status', "New"); break;
              case "Source Email ID":       newRow[i] = getData('sourceEmailId'); break;
              case "Processed Timestamp":   newRow[i] = getData('processedTimestamp', new Date()); break;
              case "Notes":                 newRow[i] = getData('notes'); break;
              // Add other cases if there are more headers in LEADS_SHEET_HEADERS
          }
      }
  }
  
  try {
    sheet.appendRow(newRow);
    // Logger.log(`[LEADS_SHEET_UTIL SUCCESS] Appended lead: "${jobData.jobTitle || 'N/A'}" to sheet "${sheet.getName()}".`);
  } catch (e) {
    Logger.log(`[LEADS_SHEET_UTIL ERROR] Failed to append row for lead "${jobData.jobTitle || 'N/A'}" to sheet "${sheet.getName()}": ${e.toString()}`);
  }
}

/**
 * Appends an error entry to the "Potential Job Leads" sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The "Potential Job Leads" sheet object.
 * @param {GoogleAppsScript.Gmail.GmailMessage} message The Gmail message that caused the error.
 * @param {string} errorType A short description of the error type.
 * @param {string} details Additional details about the error.
 * @param {Object} headerMap An object mapping header names to column numbers for the sheet.
 */

function writeErrorEntryToSheet_forLeads(sheet, message, errorType, details, headerMap) {
  if (!sheet) { Logger.log("[LEADS_SHEET_UTIL ERROR] Cannot write error entry: sheet object is null."); return; }
  if (!headerMap || Object.keys(headerMap).length === 0) { Logger.log("[LEADS_SHEET_UTIL ERROR] Cannot write error entry: headerMap is invalid or empty."); return; }
  
  const detailsString = typeof details === 'string' ? details.substring(0, 1000) : JSON.stringify(details).substring(0, 1000);
  Logger.log(`[LEADS_SHEET_UTIL INFO] Writing error entry for msg ${message ? message.getId() : 'Unknown Msg'}: ${errorType}. Details: ${detailsString}`);

  const errorJobData = {
      dateAdded: new Date(),
      jobTitle: "PROCESSING ERROR",
      company: errorType.substring(0, 250), // Keep company field relatively short
      location: "N/A",
      sourceEmailSubject: message ? message.getSubject() : "N/A",
      linkToJobPosting: "N/A",
      status: "Error",
      notes: `Type: ${errorType}. Details: ${detailsString}`, // Consolidate error info in Notes
      sourceEmailId: message ? message.getId() : "N/A",
      processedTimestamp: new Date()
  };
  // Use the same writeJobDataToSheet_forLeads function to write the error entry
  writeJobDataToSheet_forLeads(sheet, errorJobData, headerMap);
}

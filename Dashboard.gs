// File: Dashboard.gs
// Description: Manages the creation, formatting, and updating of the
// main dashboard sheet, its charts, and the associated helper data sheet.
// Includes logic for scorecard formulas and chart data preparation.

// --- Helper: Get or Create Dashboard Sheet and move to front ---
function getOrCreateDashboardSheet(spreadsheet) {
  let dashboardSheet = spreadsheet.getSheetByName(DASHBOARD_TAB_NAME); // From Config.gs
  if (!dashboardSheet) {
    dashboardSheet = spreadsheet.insertSheet(DASHBOARD_TAB_NAME, 0); // Insert at the very first position
    Logger.log(`[INFO] SETUP_DASH: Created new dashboard sheet "${DASHBOARD_TAB_NAME}" at the first position.`);

    // Optional: Clean up default "Sheet1" if it exists and isn't the data or dashboard sheet
    if (spreadsheet.getSheets().length > 2) {
        // --- CORRECTED LINE ---
        const appTrackerDataSheetName = APP_TRACKER_SHEET_TAB_NAME; // Use the correct constant from Config.gs
        // --- END CORRECTED LINE ---
        const defaultSheet = spreadsheet.getSheetByName('Sheet1');

        // Ensure the appTrackerDataSheetName is valid before using getSheetByName with it
        let dataSheetForComparison = null;
        if (appTrackerDataSheetName) {
            dataSheetForComparison = spreadsheet.getSheetByName(appTrackerDataSheetName);
        }

        if (defaultSheet && dataSheetForComparison && // Check if dataSheetForComparison was successfully retrieved
            defaultSheet.getSheetId() !== dataSheetForComparison.getSheetId() &&
            defaultSheet.getSheetId() !== dashboardSheet.getSheetId()) {
            try {
                spreadsheet.deleteSheet(defaultSheet);
                Logger.log(`[INFO] SETUP_DASH: Removed default 'Sheet1' after dashboard creation (it was not "${DASHBOARD_TAB_NAME}" or "${appTrackerDataSheetName}").`);
            } catch (eDeleteDefault) {
                Logger.log(`[WARN] SETUP_DASH: Failed to remove default 'Sheet1' after dashboard creation: ${eDeleteDefault.message}`);
            }
        } else if (defaultSheet && !dataSheetForComparison) {
            Logger.log(`[WARN] SETUP_DASH: Could not find the application data sheet named "${appTrackerDataSheetName}" for comparison when trying to delete default 'Sheet1'. 'Sheet1' was not deleted.`);
        }
    }
  } else {
    // If it exists, ensure it's the active sheet and then move it to the first position.
    spreadsheet.setActiveSheet(dashboardSheet);
    spreadsheet.moveActiveSheet(0); // Index 0 makes it the first sheet
    Logger.log(`[INFO] SETUP_DASH: Found existing dashboard sheet "${DASHBOARD_TAB_NAME}" and ensured it is the first tab.`);
  }
  return dashboardSheet;
}

// --- Helper: Format Dashboard Sheet (Initial Layout & Styling, plus Helper Sheet Formula Setup) ---
function formatDashboardSheet(dashboardSheet) {
  if (!dashboardSheet || typeof dashboardSheet.getName !== 'function') {
    Logger.log(`[DASHBOARD ERROR] FORMAT_DASH: Invalid dashboardSheet object provided. Cannot format.`);
    return;
  }
  Logger.log(`[DASHBOARD INFO] FORMAT_DASH: Starting initial formatting for dashboard sheet "${dashboardSheet.getName()}".`);

  dashboardSheet.clear();
  dashboardSheet.clearFormats();
  dashboardSheet.clearNotes();
  // Clear existing conditional formatting rules more safely
  try {
    const rules = dashboardSheet.getConditionalFormatRules();
    if (rules && rules.length > 0) {
        dashboardSheet.setConditionalFormatRules(rules.map(rule => null)); // This is a way to clear, or iterate and removeById
    }
  } catch (e) { Logger.log(`[DASHBOARD WARN] FORMAT_DASH: Could not clear all conditional format rules: ${e.message}`);}


  try {
    dashboardSheet.setHiddenGridlines(true);
    // Logger.log(`[DASHBOARD INFO] FORMAT_DASH: Gridlines hidden for "${dashboardSheet.getName()}".`);
  } catch (e) {
    Logger.log(`[DASHBOARD ERROR] FORMAT_DASH: Error hiding gridlines: ${e.toString()}`);
  }

  // --- Color Constants & Font Settings ---
  const TEAL_ACCENT_BG = "#26A69A", HEADER_TEXT_COLOR = "#FFFFFF",
        LIGHT_GREY_BG = "#F5F5F5", DARK_GREY_TEXT = "#424242",
        CARD_BORDER_COLOR = "#BDBDBD", VALUE_TEXT_COLOR = TEAL_ACCENT_BG,
        METRIC_FONT_SIZE = 15, METRIC_FONT_WEIGHT = "bold", LABEL_FONT_WEIGHT = "bold";
  const SECONDARY_CARD_BG = "#FFFDE7", SECONDARY_VALUE_COLOR = "#FF8F00";
  const ORANGE_CARD_BG = "#FFF3E0", ORANGE_VALUE_COLOR = "#EF6C00";
  const spacerColAWidth = 20;

  dashboardSheet.getRange("A1:M1").merge()
                .setValue("Job Application Dashboard").setBackground(TEAL_ACCENT_BG)
                .setFontColor(HEADER_TEXT_COLOR).setFontSize(18).setFontWeight("bold")
                .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dashboardSheet.setRowHeight(1, 45);
  dashboardSheet.setRowHeight(2, 10); // Spacer

  dashboardSheet.getRange("B3").setValue("Key Metrics Overview:")
                .setFontSize(14).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(3, 30);
  dashboardSheet.setRowHeight(4, 10); // Spacer

  // --- DEFINE Column Letters for Formulas (using APP_TRACKER_SHEET_TAB_NAME) ---
  const appSheetNameForFormula = `'${APP_TRACKER_SHEET_TAB_NAME}'`; // CORRECTED: Uses APP_TRACKER_SHEET_TAB_NAME from Config.gs
  const companyColLetter = columnToLetter(COMPANY_COL); // Constants from Config.gs
  const jobTitleColLetter = columnToLetter(JOB_TITLE_COL);
  const statusColLetter = columnToLetter(STATUS_COL);
  const peakStatusColLetter = columnToLetter(PEAK_STATUS_COL);
  const emailDateColLetter = columnToLetter(EMAIL_DATE_COL);
  const platformColLetter = columnToLetter(PLATFORM_COL);

  // --- Scorecard Setup on Dashboard Sheet (Formulas point to APP_TRACKER_SHEET_TAB_NAME) ---
  // Row 1 of Scorecards (Sheet Row 5)
  dashboardSheet.getRange("B5").setValue("Total Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("C5").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("E5").setValue("Peak Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("F5").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("H5").setValue("Interview Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("I5").setFormula(`=IFERROR(F5/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%");
  dashboardSheet.getRange("K5").setValue("Offer Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("L5").setFormula(`=IFERROR(F7/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%");
  dashboardSheet.setRowHeight(5, 40);
  dashboardSheet.setRowHeight(6, 10);

  // Row 2 of Scorecards (Sheet Row 7)
  dashboardSheet.getRange("B7").setValue("Active Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  let activeAppsFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>"&"", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${REJECTED_STATUS}", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${ACCEPTED_STATUS}"), 0)`;
  dashboardSheet.getRange("C7").setFormula(activeAppsFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("E7").setValue("Peak Offers").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("F7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${OFFER_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("H7").setValue("Current Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("I7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("K7").setValue("Current Assessments").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("L7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${ASSESSMENT_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.setRowHeight(7, 40);
  dashboardSheet.setRowHeight(8, 10);

  // Row 3 of Scorecards (Sheet Row 9)
  dashboardSheet.getRange("B9").setValue("Total Rejections").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("C9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("E9").setValue("Apps Viewed (Peak)").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("F9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${APPLICATION_VIEWED_STATUS}"),0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("H9").setValue("Manual Review").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  const compColManualFormula = `${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}="${MANUAL_REVIEW_NEEDED}"`;
  const titleColManualFormula = `${appSheetNameForFormula}!${jobTitleColLetter}2:${jobTitleColLetter}="${MANUAL_REVIEW_NEEDED}"`;
  const statusColManualFormula = `${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}="${MANUAL_REVIEW_NEEDED}"`;
  const finalManualReviewFormula = `=IFERROR(SUM(ARRAYFORMULA(SIGN((${compColManualFormula})+(${titleColManualFormula})+(${statusColManualFormula})))),0)`;
  dashboardSheet.getRange("I9").setFormula(finalManualReviewFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("K9").setValue("Direct Reject Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  const directRejectFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${DEFAULT_STATUS}",${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}")/C5, 0)`;
  dashboardSheet.getRange("L9").setFormula(directRejectFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%");
  dashboardSheet.setRowHeight(9, 40);
  dashboardSheet.setRowHeight(10, 15);

  const scorecardRangesToStyle = [
      "B5:C5", "E5:F5", "H5:I5", "K5:L5", "B7:C7", "E7:F7", "H7:I7", "K7:L7", "B9:C9", "E9:F9", "H9:I9", "K9:L9"
  ];
  scorecardRangesToStyle.forEach(rangeString => {
    const range = dashboardSheet.getRange(rangeString);
    range.setBackground(LIGHT_GREY_BG).setBorder(true, true, true, true, true, true, CARD_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
  });
  dashboardSheet.getRange("H9:I9").setBackground(SECONDARY_CARD_BG); // Manual Review Card
  dashboardSheet.getRange("K9:L9").setBackground(ORANGE_CARD_BG);    // Direct Reject Rate Card
  const primaryValueCells = ["C5", "F5", "I5", "L5", "C7", "F7", "I7", "L7", "C9", "F9"];
  primaryValueCells.forEach(cellAddress => { dashboardSheet.getRange(cellAddress).setFontColor(VALUE_TEXT_COLOR); });
  dashboardSheet.getRange("I9").setFontColor(SECONDARY_VALUE_COLOR); // Manual Review Value
  dashboardSheet.getRange("L9").setFontColor(ORANGE_VALUE_COLOR);    // Direct Reject Value
  const labelCellAddresses = ["B5", "E5", "H5", "K5", "B7", "E7", "H7", "K7", "B9", "E9", "H9", "K9"];
  labelCellAddresses.forEach(cellAddress => { dashboardSheet.getRange(cellAddress).setFontColor(DARK_GREY_TEXT); });

  const chartSectionTitleRow1 = 11;
  dashboardSheet.getRange("B" + chartSectionTitleRow1).setValue("Platform & Weekly Trends").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(chartSectionTitleRow1, 25);
  dashboardSheet.setRowHeight(chartSectionTitleRow1 + 1, 5);

  const chartSectionTitleRow2 = 28;
  dashboardSheet.getRange("B" + chartSectionTitleRow2).setValue("Application Funnel Analysis").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(chartSectionTitleRow2, 25);
  dashboardSheet.setRowHeight(chartSectionTitleRow2 + 1, 5);

  const labelWidth = 150; const valueWidth = 75; const spacerS = 15;
  dashboardSheet.setColumnWidth(1, spacerColAWidth);
  dashboardSheet.setColumnWidth(2, labelWidth); dashboardSheet.setColumnWidth(3, valueWidth); dashboardSheet.setColumnWidth(4, spacerS);
  dashboardSheet.setColumnWidth(5, labelWidth); dashboardSheet.setColumnWidth(6, valueWidth); dashboardSheet.setColumnWidth(7, spacerS);
  dashboardSheet.setColumnWidth(8, labelWidth); dashboardSheet.setColumnWidth(9, valueWidth); dashboardSheet.setColumnWidth(10, spacerS);
  dashboardSheet.setColumnWidth(11, labelWidth); dashboardSheet.setColumnWidth(12, valueWidth);
  dashboardSheet.setColumnWidth(13, spacerColAWidth);

  const ss = dashboardSheet.getParent();
  let helperSheet = ss.getSheetByName(HELPER_SHEET_NAME); // HELPER_SHEET_NAME from Config.gs
  if (!helperSheet) {
    helperSheet = getOrCreateHelperSheet(ss); // From Dashboard.gs (or SheetUtils.gs if moved)
    if (!helperSheet) {
        Logger.log(`[DASHBOARD ERROR] FORMAT_DASH_HELPER: Helper sheet "${HELPER_SHEET_NAME}" missing and could not be created. Cannot set formulas.`);
        return;
    }
  }
  if (helperSheet) {
    Logger.log(`[DASHBOARD INFO] FORMAT_DASH_HELPER: Setting up formulas in helper sheet "${helperSheet.getName()}".`);
    helperSheet.getRange("A:B").clearContent(); // Clear only relevant columns
    helperSheet.getRange("D:E").clearContent();
    helperSheet.getRange("G:H").clearContent();
    helperSheet.getRange("J:K").clearContent();
    Logger.log(`[DASHBOARD DEBUG] FORMAT_DASH_HELPER: Cleared A:B, D:E, G:H, J:K in helper sheet.`);

    helperSheet.getRange("A1").setValue("Platform"); helperSheet.getRange("B1").setValue("Count");
    const platformQueryFormula = `=IFERROR(QUERY(${appSheetNameForFormula}!${platformColLetter}2:${platformColLetter}, "SELECT ${platformColLetter}, COUNT(${platformColLetter}) WHERE ${platformColLetter} IS NOT NULL AND ${platformColLetter} <> '' GROUP BY ${platformColLetter} ORDER BY COUNT(${platformColLetter}) DESC LABEL ${platformColLetter} '', COUNT(${platformColLetter}) ''", 0), {"No Platforms",0})`;
    helperSheet.getRange("A2").setFormula(platformQueryFormula);

    helperSheet.getRange("J1").setValue("RAW_VALID_DATES_FOR_WEEKLY");
    const rawDatesFormula = `=IFERROR(FILTER(${appSheetNameForFormula}!${emailDateColLetter}2:${emailDateColLetter}, ISNUMBER(${appSheetNameForFormula}!${emailDateColLetter}2:${emailDateColLetter})), "")`;
    helperSheet.getRange("J2").setFormula(rawDatesFormula);
    helperSheet.getRange("J2:J").setNumberFormat("yyyy-mm-dd hh:mm:ss");

    helperSheet.getRange("K1").setValue("CALCULATED_WEEK_STARTS");
    const weekStartCalcFormula = `=ARRAYFORMULA(IF(ISBLANK(J2:J), "", DATE(YEAR(J2:J), MONTH(J2:J), DAY(J2:J) - WEEKDAY(J2:J, 2) + 1)))`; // Monday Start
    helperSheet.getRange("K2").setFormula(weekStartCalcFormula);
    helperSheet.getRange("K2:K").setNumberFormat("yyyy-mm-dd");

    helperSheet.getRange("D1").setValue("Week Starting");
    const uniqueWeeksFormula = `=IFERROR(SORT(UNIQUE(FILTER(K2:K, K2:K<>""))), {"No Data"})`;
    helperSheet.getRange("D2").setFormula(uniqueWeeksFormula);
    helperSheet.getRange("D2:D").setNumberFormat("yyyy-mm-dd");

    helperSheet.getRange("E1").setValue("Applications");
    const weeklyCountsFormula = `=ARRAYFORMULA(IF(D2:D="", "", COUNTIF(K2:K, D2:D)))`;
    helperSheet.getRange("E2").setFormula(weeklyCountsFormula);
    helperSheet.getRange("E2:E").setNumberFormat("0");

    helperSheet.getRange("G1").setValue("Stage"); helperSheet.getRange("H1").setValue("Count");
    const funnelStagesValues = [DEFAULT_STATUS, APPLICATION_VIEWED_STATUS, ASSESSMENT_STATUS, INTERVIEW_STATUS, OFFER_STATUS]; // From Config.gs
    helperSheet.getRange(2, 7, funnelStagesValues.length, 1).setValues(funnelStagesValues.map(stage => [stage]));
    helperSheet.getRange("H2").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}),0)`); // Total applications as first stage count
    for (let i = 1; i < funnelStagesValues.length; i++) {
      helperSheet.getRange(i + 2, 8).setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter}, G${i + 2}),0)`);
    }
    Logger.log(`[DASHBOARD INFO] FORMAT_DASH_HELPER: All helper formulas set in "${helperSheet.getName()}".`);
  }

  const lastUsedDataColumnOnDashboard = 13;
  const maxColsDashboard = dashboardSheet.getMaxColumns();
  if (maxColsDashboard > 0) { dashboardSheet.showColumns(1, maxColsDashboard); } // Show all first
  if (maxColsDashboard > lastUsedDataColumnOnDashboard) {
      dashboardSheet.hideColumns(lastUsedDataColumnOnDashboard + 1, maxColsDashboard - lastUsedDataColumnOnDashboard);
  }

  const lastUsedDataRowOnDashboard = 45; // Adjust if your dashboard content grows taller
  const maxRowsDashboard = dashboardSheet.getMaxRows();
  if (maxRowsDashboard > 1) { dashboardSheet.showRows(1, maxRowsDashboard); } // Show all first
  if (maxRowsDashboard > lastUsedDataRowOnDashboard) {
      dashboardSheet.hideRows(lastUsedDataRowOnDashboard + 1, maxRowsDashboard - lastUsedDataRowOnDashboard);
  }
  Logger.log(`[DASHBOARD INFO] FORMAT_DASH: Formatting concluded for Dashboard "${dashboardSheet.getName()}".`);
}

// --- Dashboard: Update Metrics and Chart Data (Helper Sheet is Formula-Driven) ---
function updateDashboardMetrics() {
  const SCRIPT_START_TIME_DASH = new Date();
  Logger.log(`\n==== STARTING DASHBOARD METRICS UPDATE (Helper Sheet is Formula-Driven) (${SCRIPT_START_TIME_DASH.toLocaleString()}) ====`);

  const { spreadsheet: ss } = getOrCreateSpreadsheetAndSheet();
  if (!ss) { 
    Logger.log(`[ERROR] UPDATE_DASH: Could not get spreadsheet. Aborting.`);
    return; 
  }

  const dashboardSheet = ss.getSheetByName(DASHBOARD_TAB_NAME); 
  const helperSheet = ss.getSheetByName(HELPER_SHEET_NAME); 

  if (!dashboardSheet) {
      Logger.log(`[WARN] UPDATE_DASH: Dashboard sheet missing. Cannot create/update charts.`);
      // If no dashboard sheet, can't do chart updates
      const SCRIPT_END_TIME_DASH_NO_DASH = new Date();
      Logger.log(`\n==== DASHBOARD METRICS UPDATE FINISHED (No Dashboard Sheet) (${SCRIPT_END_TIME_DASH_NO_DASH.toLocaleString()}) ====`);
      return;
  }
  if (!helperSheet) {
    Logger.log(`[ERROR] UPDATE_DASH: Helper sheet missing. Cannot verify chart data sources or create/update charts.`);
    // If no helper sheet, chart data sources are invalid
    const SCRIPT_END_TIME_DASH_NO_HELP = new Date();
    Logger.log(`\n==== DASHBOARD METRICS UPDATE FINISHED (No Helper Sheet) (${SCRIPT_END_TIME_DASH_NO_HELP.toLocaleString()}) ====`);
    return; 
  }

  Logger.log(`[INFO] UPDATE_DASH: All scorecard metrics AND chart helper data are formula-based.`);
  
  // This function's primary role now is to ensure the chart OBJECTS exist on the dashboard
  // and are correctly configured to point to the formula-driven helper sheet data.
  // The actual data aggregation happens via formulas in the helper sheet itself.

  // --- Call Chart Update Functions ---
  // These functions will check if charts exist. If not, they create them using data
  // from helperSheet (which is now formula-driven and should be up-to-date).
  // If charts exist, they mainly ensure ranges are still correct (though this often
  // isn't strictly needed if ranges are static like 'HelperSheet!A1:B10').
  if (dashboardSheet && helperSheet) { // Redundant check, but safe
     Logger.log(`[INFO] UPDATE_DASH: Ensuring chart objects are present and configured...`);
     try {
        updatePlatformDistributionChart(dashboardSheet, helperSheet);
        updateApplicationsOverTimeChart(dashboardSheet, helperSheet);
        updateApplicationFunnelChart(dashboardSheet, helperSheet);
        Logger.log(`[INFO] UPDATE_DASH: Chart object presence and configuration check complete.`);
     } catch (e) {
        Logger.log(`[ERROR] UPDATE_DASH: Error during chart update/creation calls: ${e.toString()} \nStack: ${e.stack}`);
     }
  } else {
      // This case should ideally be caught by earlier checks for dashboardSheet and helperSheet
      Logger.log(`[WARN] UPDATE_DASH: Skipping chart object updates as dashboardSheet or helperSheet is unexpectedly missing at this stage.`);
  }

  const SCRIPT_END_TIME_DASH = new Date();
  Logger.log(`\n==== DASHBOARD METRICS UPDATE FINISHED (${SCRIPT_END_TIME_DASH.toLocaleString()}) === Total Time: ${(SCRIPT_END_TIME_DASH.getTime() - SCRIPT_START_TIME_DASH.getTime())/1000}s ====`);
}

// --- Helper: Get or Create Helper Sheet ---
function getOrCreateHelperSheet(spreadsheet) {
  let helperSheet = spreadsheet.getSheetByName(HELPER_SHEET_NAME);
  if (!helperSheet) {
    helperSheet = spreadsheet.insertSheet(HELPER_SHEET_NAME);
    Logger.log(`[INFO] SETUP_HELPER: Created new helper sheet "${HELPER_SHEET_NAME}".`);
    try {
      helperSheet.hideSheet(); // Hide it by default
      Logger.log(`[INFO] SETUP_HELPER: Sheet "${HELPER_SHEET_NAME}" has been hidden.`);
    } catch (e) {
      Logger.log(`[WARN] SETUP_HELPER: Could not hide helper sheet "${HELPER_SHEET_NAME}": ${e}`);
    }
  } else {
    Logger.log(`[INFO] SETUP_HELPER: Found existing helper sheet "${HELPER_SHEET_NAME}". Ensuring it's hidden if not already.`);
    if (!helperSheet.isSheetHidden()) {
        try {
            helperSheet.hideSheet();
            Logger.log(`[INFO] SETUP_HELPER: Sheet "${HELPER_SHEET_NAME}" was visible and has now been hidden.`);
        } catch (e) {
            Logger.log(`[WARN] SETUP_HELPER: Could not hide existing helper sheet "${HELPER_SHEET_NAME}": ${e}`);
        }
    }
  }
  return helperSheet;
}

// --- Dashboard Chart: Update Platform Distribution Pie Chart ---
function updatePlatformDistributionChart(dashboardSheet, helperSheet) {
  Logger.log(`[INFO] CHART_PLATFORM: Attempting to create/update Platform Distribution chart.`);
  const CHART_TITLE = "Platform Distribution";
  let existingChart = null;

  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    if (charts[i].getOptions().get('title') === CHART_TITLE && charts[i].getContainerInfo().getAnchorColumn() === 2) {
      existingChart = charts[i];
      break;
    }
  }
  if(existingChart) Logger.log(`[DEBUG] CHART_PLATFORM: Found existing chart.`);
  else Logger.log(`[DEBUG] CHART_PLATFORM: No existing chart with title '${CHART_TITLE}'. Will create new.`);

  // This check is CRITICAL. Number of rows with actual data in column A of helper sheet.
  const lastPlatformRow = helperSheet.getRange("A1:A").getValues().filter(String).length; 
  let dataRange;

  // We need at least a header AND one row of data (lastPlatformRow >= 2) for a chart
  if (helperSheet.getRange("A1").getValue() === "Platform" && lastPlatformRow >= 2) {
      dataRange = helperSheet.getRange(`A1:B${lastPlatformRow}`);
      Logger.log(`[INFO] CHART_PLATFORM: Data range for chart set to ${HELPER_SHEET_NAME}!A1:B${lastPlatformRow}`);
  } else {
      Logger.log(`[WARN] CHART_PLATFORM: Not enough data or invalid header for platform chart (found ${lastPlatformRow} rows with content in A). Chart will not be created/updated.`);
      if (existingChart) { 
          try { dashboardSheet.removeChart(existingChart); Logger.log(`[INFO] CHART_PLATFORM: Removed existing chart due to insufficient data.`);}
          catch (e) { Logger.log(`[ERROR] CHART_PLATFORM: Could not remove chart: ${e}`); }
      }
      return; // EXIT HERE if no valid data
  }

  const chartSectionTitleRow1 = 11; 
  const anchorRow = chartSectionTitleRow1 + 2; // Should be 13
  const anchorCol = 2;  // Column B
  const chartWidth = 460; 
  const chartHeight = 280; 

    const optionsToSet = {
    title: CHART_TITLE,
    legend: { position: Charts.Position.RIGHT },
    pieHole: 0.4,
    width: chartWidth,
    height: chartHeight,
    sliceVisibilityThreshold: 0 // Add this line
  };

  try { // Wrap chart operations in try-catch
    if (existingChart) {
      Logger.log(`[DEBUG] CHART_PLATFORM: Modifying existing chart.`);
      let chartBuilder = existingChart.modify();
      chartBuilder = chartBuilder.clearRanges().addRange(dataRange).setChartType(Charts.ChartType.PIE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) chartBuilder = chartBuilder.setOption(key, optionsToSet[key]); }
      chartBuilder = chartBuilder.setPosition(anchorRow, anchorCol, 0, 0); 
      dashboardSheet.updateChart(chartBuilder.build());
      Logger.log(`[INFO] CHART_PLATFORM: Updated existing chart "${CHART_TITLE}".`);
    } else { 
      Logger.log(`[DEBUG] CHART_PLATFORM: Creating new chart.`);
      let newChartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) newChartBuilder = newChartBuilder.setOption(key, optionsToSet[key]); }
      newChartBuilder = newChartBuilder.addRange(dataRange).setPosition(anchorRow, anchorCol, 0, 0);
      dashboardSheet.insertChart(newChartBuilder.build());
      Logger.log(`[INFO] CHART_PLATFORM: Created new chart "${CHART_TITLE}".`);
    }
  } catch (e) {
    Logger.log(`[ERROR] CHART_PLATFORM: Failed during chart build/insert/update: ${e.message} ${e.stack}`);
  }
}

// --- Helper: Get Week Start Date (Monday) ---
function getWeekStartDate(inputDate) {
  const date = new Date(inputDate.getTime()); // Clone the date to avoid modifying the original
  const day = date.getDay(); // Sunday - 0, Monday - 1, ..., Saturday - 6
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
  return new Date(date.setDate(diff));
}

// --- Dashboard Chart: Update Applications Over Time (Weekly) Line Chart ---
function updateApplicationsOverTimeChart(dashboardSheet, helperSheet) {
  Logger.log(`[INFO] CHART_APPS_TIME: Attempting to create/update Applications Over Time chart.`);
  const CHART_TITLE = "Applications Over Time (Weekly)";
  let existingChart = null;

  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    // Check by title and a known anchor column (e.g., Col H where it's expected to start)
    if (charts[i].getOptions().get('title') === CHART_TITLE && charts[i].getContainerInfo().getAnchorColumn() === 8) { 
      existingChart = charts[i];
      break;
    }
  }
  if(existingChart) Logger.log(`[DEBUG] CHART_APPS_TIME: Found existing chart.`);
  else Logger.log(`[DEBUG] CHART_APPS_TIME: No existing chart with title '${CHART_TITLE}'. Will create new.`);

  // Determine the actual last row of data in the helper sheet for this chart (D:E)
  // Counts non-empty cells in column D (Week Starting)
  const lastWeeklyDataRowInHelper = helperSheet.getRange("D1:D").getValues().filter(String).length; 
  let dataRange;

  // Check if D1 has the correct header AND if there is at least one data row (header is row 1, so need >=2 for data)
  if (helperSheet.getRange("D1").getValue() === "Week Starting" && lastWeeklyDataRowInHelper >= 2) {
      dataRange = helperSheet.getRange(`D1:E${lastWeeklyDataRowInHelper}`); // e.g., D1:E5 if header + 4 data weeks
      Logger.log(`[INFO] CHART_APPS_TIME: Data range for chart set to ${helperSheet.getName()}!D1:E${lastWeeklyDataRowInHelper}`);
  } else {
      Logger.log(`[WARN] CHART_APPS_TIME: Not enough data or invalid header for weekly chart (Helper Col D has ${lastWeeklyDataRowInHelper} content rows including header). Chart will not be created/updated.`);
      if (existingChart) { 
        try { dashboardSheet.removeChart(existingChart); Logger.log(`[INFO] CHART_APPS_TIME: Removed existing chart due to insufficient/invalid data in helper sheet.`); }
        catch (e) { Logger.log(`[ERROR] CHART_APPS_TIME: Could not remove chart: ${e}`); }
      }
      return; // EXIT HERE if no valid data to plot
  }

  // Anchor Row values should align with your formatDashboardSheet
  const chartSectionTitleRow1 = 11; // As set in your latest formatDashboardSheet
  const anchorRow = chartSectionTitleRow1 + 2; // Should be 13
  const anchorCol = 8;  // Column H

  const chartWidth = 460; // As determined previously
  const chartHeight = 280; 

  const optionsToSet = {
    title: CHART_TITLE,
    hAxis: { title: 'Week Starting', textStyle: { fontSize: 10 }, format: 'M/d' }, // Short date format for axis
    vAxis: { title: 'Number of Applications', textStyle: { fontSize: 10 }, viewWindow: { min: 0 } },
    legend: { position: 'none' }, 
    colors: ['#26A69A'], 
    width: chartWidth,
    height: chartHeight,
    // pointSize: 5, // Optionally add points to the line chart
    // curveType: 'function' // For a smoothed line, if desired
  };

  try { 
    if (existingChart) { 
      Logger.log(`[DEBUG] CHART_APPS_TIME: Modifying existing chart.`);
      let chartBuilder = existingChart.modify();
      chartBuilder = chartBuilder.clearRanges().addRange(dataRange).setChartType(Charts.ChartType.LINE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) chartBuilder = chartBuilder.setOption(key, optionsToSet[key]); }
      chartBuilder = chartBuilder.setPosition(anchorRow, anchorCol, 0, 0);
      dashboardSheet.updateChart(chartBuilder.build());
      Logger.log(`[INFO] CHART_APPS_TIME: Updated existing chart "${CHART_TITLE}".`);
    } else { 
      Logger.log(`[DEBUG] CHART_APPS_TIME: Creating new chart.`);
      let newChartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.LINE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) newChartBuilder = newChartBuilder.setOption(key, optionsToSet[key]); }
      newChartBuilder = newChartBuilder.addRange(dataRange).setPosition(anchorRow, anchorCol, 0, 0);
      dashboardSheet.insertChart(newChartBuilder.build());
      Logger.log(`[INFO] CHART_APPS_TIME: Created new chart "${CHART_TITLE}".`);
    }
  } catch (e) {
      Logger.log(`[ERROR] CHART_APPS_TIME: Failed during chart build/insert/update: ${e.message} \nStack: ${e.stack}`);
  }
}

// --- Dashboard Chart: Update Application Funnel (Peak Stages) Column Chart ---
function updateApplicationFunnelChart(dashboardSheet, helperSheet) {
  Logger.log(`[INFO] CHART_FUNNEL: Attempting to create/update Application Funnel chart (using individual setOption).`);
  const CHART_TITLE = "Application Funnel (Peak Stages)";
  let existingChart = null;

  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    if (charts[i].getOptions().get('title') === CHART_TITLE && charts[i].getContainerInfo().getAnchorColumn() === 2) {
      existingChart = charts[i];
      break;
    }
  }
  if(existingChart) Logger.log(`[DEBUG] CHART_FUNNEL: Found existing chart.`);
  else Logger.log(`[DEBUG] CHART_FUNNEL: No existing chart found with title '${CHART_TITLE}'. Will create new.`);

  const lastFunnelDataRow = helperSheet.getRange("G:G").getValues().filter(String).length;
  let dataRange;
  if (helperSheet.getRange("G1").getValue() === "Stage" && lastFunnelDataRow >= 2) {
      dataRange = helperSheet.getRange(`G1:H${lastFunnelDataRow}`);
      Logger.log(`[INFO] CHART_FUNNEL: Data range for chart set to ${HELPER_SHEET_NAME}!G1:H${lastFunnelDataRow}`);
  } else {
      Logger.log(`[WARN] CHART_FUNNEL: No data/invalid header for funnel chart. Chart will not be created/updated.`);
      if (existingChart) { dashboardSheet.removeChart(existingChart); Logger.log(`[INFO] CHART_FUNNEL: Removed chart - no data.`);}
      return;
  }

  const chartSectionTitleRow2 = 28; 
  const anchorRow = chartSectionTitleRow2 + 2; // Should be 30
  const anchorCol = 2;  // Column B
  const chartWidth = 460; 
  const chartHeight = 280; 

  const optionsToSet = {
    title: CHART_TITLE,
    hAxis: { title: 'Application Stage', textStyle: { fontSize: 10 }, slantedText: true, slantedTextAngle: 30 },
    vAxis: { title: 'Number of Applications', textStyle: { fontSize: 10 }, viewWindow: { min: 0 } },
    legend: { position: 'none' }, // Or Charts.Position.NONE
    colors: ['#26A69A'], 
    bar: { groupWidth: '60%' }, 
    width: chartWidth,
    height: chartHeight,
  };
  Logger.log(`[DEBUG] CHART_FUNNEL: Options object for chart: ${JSON.stringify(optionsToSet)}`);

  if (existingChart) { 
    Logger.log(`[DEBUG] CHART_FUNNEL: Modifying existing chart.`);
    let chartBuilder = existingChart.modify();
    chartBuilder = chartBuilder.clearRanges();
    chartBuilder = chartBuilder.addRange(dataRange); // For modify, addRange can come before or after some setOptions
    chartBuilder = chartBuilder.setChartType(Charts.ChartType.COLUMN);
    
    // Loop to set options individually for modify
    for (const key in optionsToSet) {
      if (optionsToSet.hasOwnProperty(key)) {
        chartBuilder = chartBuilder.setOption(key, optionsToSet[key]);
      }
    }
    
    chartBuilder = chartBuilder.setPosition(anchorRow, anchorCol, 0, 0);
    const updatedChart = chartBuilder.build();
    dashboardSheet.updateChart(updatedChart);
    Logger.log(`[INFO] CHART_FUNNEL: Updated existing chart "${CHART_TITLE}".`);
  } else { // Creating a new chart
    Logger.log(`[DEBUG] CHART_FUNNEL: Creating new chart using individual setOption calls.`);
    let newChartBuilder = dashboardSheet.newChart();
    newChartBuilder = newChartBuilder.setChartType(Charts.ChartType.COLUMN);
    
    // Apply options individually
    for (const key in optionsToSet) {
      if (optionsToSet.hasOwnProperty(key)) {
        Logger.log(`[DEBUG] CHART_FUNNEL (New): Setting option ${key} = ${JSON.stringify(optionsToSet[key])}`);
        newChartBuilder = newChartBuilder.setOption(key, optionsToSet[key]);
      }
    }
    
    Logger.log(`[DEBUG] CHART_FUNNEL (New): After all individual setOption, type: ${typeof newChartBuilder}, has addRange: ${typeof newChartBuilder.addRange}`);
    newChartBuilder = newChartBuilder.addRange(dataRange);
    Logger.log(`[DEBUG] CHART_FUNNEL (New): After addRange, type: ${typeof newChartBuilder}, has setPosition: ${typeof newChartBuilder.setPosition}`);
    newChartBuilder = newChartBuilder.setPosition(anchorRow, anchorCol, 0, 0);
    const newChart = newChartBuilder.build();
    dashboardSheet.insertChart(newChart);
    Logger.log(`[INFO] CHART_FUNNEL: Created new chart "${CHART_TITLE}" using individual setOptions.`);
  }
}

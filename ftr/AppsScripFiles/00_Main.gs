/**
 * @file 00_Main.gs
 * @description Main router. Priority-based execution.
 * Updated to use DAVE_FOLDER_ID for both Prod and Test environments.
 */

function runDailyUpdate(forceProd = false, forceTest = false) {
  const startTime = new Date().getTime();
  const prodSheet = SpreadsheetApp.openById(FTR_SHEET_ID_PROD);
  const lookupSheet = prodSheet.getSheetByName(LOOKUP_SHEET_NAME);
  const isWeeklyRunNeeded = (lookupSheet.getRange(CW_CELL).getValue() !== lookupSheet.getRange(ACTUAL_CW_CELL).getValue());
  const dayOfWeek = new Date().getDay();

  let config;
  let triggerType = "";

  // 1. Check for Manual Overrides FIRST (Priority)
  if (forceProd) {
    triggerType = "MANUAL_PROD";
    Logger.log("--- Manually forcing PRODUCTION Weekly Workflow ---");
    config = {
      sheetId: FTR_SHEET_ID_PROD,
      slidesId: FTR_SLIDES_ID_PROD,
      context: "Production",
      importFolderId: DAVE_FOLDER_ID, // Master CSV Source
      isWeeklyRun: true
    };
  } else if (forceTest) {
    triggerType = "MANUAL_TEST";
    Logger.log("--- Manually forcing TESTING Daily Workflow ---");
    const ids = initializeTestingAndRelink();
    if (!ids) {
      workflowLog("Critical Failure: Could not initialize test environment.", "INIT_ERROR", startTime, prodSheet);
      return;
    }
    config = {
      sheetId: ids.sheetId,
      slidesId: ids.slidesId,
      context: "Testing",
      importFolderId: DAVE_FOLDER_ID, // FIX: Now uses the same master source as Production
      isWeeklyRun: false
    };
  }
  // 2. Automated logic (Secondary)
  else if (dayOfWeek === 4 && isWeeklyRunNeeded) {
    triggerType = "AUTO_PROD";
    Logger.log("--- Starting AUTOMATED PRODUCTION Weekly Workflow ---");
    config = {
      sheetId: FTR_SHEET_ID_PROD,
      slidesId: FTR_SLIDES_ID_PROD,
      context: "Production",
      importFolderId: DAVE_FOLDER_ID,
      isWeeklyRun: true
    };
  } else if ([1, 2, 3, 5].includes(dayOfWeek)) {
    triggerType = "AUTO_TEST";
    Logger.log("--- Starting AUTOMATED TESTING Daily Workflow ---");
    const ids = initializeTestingAndRelink();
    if (!ids) {
      workflowLog("Automated test setup failed.", "INIT_ERROR", startTime, prodSheet);
      return;
    }
    config = {
      sheetId: ids.sheetId,
      slidesId: ids.slidesId,
      context: "Testing",
      importFolderId: DAVE_FOLDER_ID, // FIX: Now uses the same master source as Production
      isWeeklyRun: false
    };
  } else {
    Logger.log("Not a scheduled run day or rollover not required. Exiting.");
    return;
  }

  // Set up target Spreadsheet object for logging
  const targetSS = SpreadsheetApp.openById(config.sheetId);

  // Clear logs only if starting a fresh weekly run or a manual test run (Step 0)
  clearWorkflowLogs(targetSS);

  // LOG START OF WORKFLOW TO SHEET
  workflowLog(`--- WORKFLOW STARTED: ${config.context} ---`, triggerType, startTime, targetSS);
  workflowLog(`Source Folder: ${config.importFolderId}`, "CONFIG_DETAILS", startTime, targetSS);
  workflowLog(`Target ID: ${config.sheetId} | Weekly: ${config.isWeeklyRun}`, "CONFIG_DETAILS", startTime, targetSS);

  // Set up the chain properties
  const props = PropertiesService.getScriptProperties();
  props.setProperty('STEP_INDEX', '0');
  props.setProperty('WF_CONFIG', JSON.stringify(config));
  props.setProperty('RELINK_CHART_INDEX', '0');

  // Start the relay race
  processChainedWorkflow();
}

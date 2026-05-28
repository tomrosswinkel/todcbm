 
/**
 * @file 04_Setup.gs
 * @description Logic for setting up the test environment and resumable chart relinking.
 */

/**
 * Initializes or resets the testing environment by creating fresh copies of the master files.
 * Enhanced with diagnostic logging to identify invalid IDs.
 * @returns {object|null} An object with new sheetId and slidesId, or null if failed.
 */
function initializeTestingAndRelink() {
  console.time("initializeTestingAndRelink");
  Logger.log('Initializing Test Environment...');

  try {
    // DIAGNOSTIC CHECK: Verify Folder IDs exist before proceeding
    let destinationFolder, backupFolder;
    try {
      destinationFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID_TEST);
    } catch (e) {
      throw new Error(`Invalid SOURCE_FOLDER_ID_TEST: ${SOURCE_FOLDER_ID_TEST}`);
    }

    try {
      backupFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID_TEST);
    } catch (e) {
      throw new Error(`Invalid BACKUP_FOLDER_ID_TEST: ${BACKUP_FOLDER_ID_TEST}`);
    }

    // Archive previous test files
    const oldSheet = destinationFolder.getFilesByName(FTR_SHEET_NAME_TEST);
    if (oldSheet.hasNext()) oldSheet.next().moveTo(backupFolder);
    const oldSlides = destinationFolder.getFilesByName(FTR_SLIDES_NAME_TEST);
    if (oldSlides.hasNext()) oldSlides.next().moveTo(backupFolder);

    // Copy production files - Verify Production IDs
    let prodSheet, prodSlides;
    try {
      prodSheet = DriveApp.getFileById(FTR_SHEET_ID_PROD);
    } catch (e) {
      throw new Error(`Invalid FTR_SHEET_ID_PROD: ${FTR_SHEET_ID_PROD}`);
    }

    try {
      prodSlides = DriveApp.getFileById(FTR_SLIDES_ID_PROD);
    } catch (e) {
      throw new Error(`Invalid FTR_SLIDES_ID_PROD: ${FTR_SLIDES_ID_PROD}`);
    }

    const testSheetFile = prodSheet.makeCopy(FTR_SHEET_NAME_TEST, destinationFolder);
    const testSlidesFile = prodSlides.makeCopy(FTR_SLIDES_NAME_TEST, destinationFolder);

    const ids = { sheetId: testSheetFile.getId(), slidesId: testSlidesFile.getId() };
    Logger.log(`Successfully created test environment. Sheet: ${ids.sheetId}`);

    console.timeEnd("initializeTestingAndRelink");
    return ids;

  } catch (e) {
    // Use the global reporter if available, otherwise standard log
    if (typeof reportFatalError === 'function') {
      reportFatalError(e.message, "INITIALIZE_TEST_ENV");
    } else {
      Logger.log(`⛔ Critical Setup Error: ${e.message}`);
    }
    return null;
  }
}

/**
 * Resumable Relinker. It processes charts one by one and saves progress to bypass time limits.
 * @param {string} slidesId
 * @param {string} sheetId
 * @param {string} prodSheetId
 * @param {number} startTime The true script start time from the engine.
 */
function relinkChartsChained(slidesId, sheetId, prodSheetId, startTime) {
  const props = PropertiesService.getScriptProperties();
  let currentIndex = parseInt(props.getProperty('RELINK_CHART_INDEX') || '0');

  const presentation = SlidesApp.openById(slidesId);
  const testSheet = SpreadsheetApp.openById(sheetId);

  // Cache charts from the test sheet to avoid heavy repeated lookups
  const allSourceCharts = {};
  testSheet.getSheets().forEach(sheet => {
    sheet.getCharts().forEach(c => allSourceCharts[c.getChartId()] = c);
  });

  const allSlides = presentation.getSlides();
  let globalCounter = 0;
  let relinkedInThisSession = 0;

  for (let s = 0; s < allSlides.length; s++) {
    const charts = allSlides[s].getSheetsCharts();
    for (let c = 0; c < charts.length; c++) {
      // If we already did this chart in a previous relay run, skip it
      if (globalCounter < currentIndex) {
        globalCounter++;
        continue;
      }

      const chart = charts[c];
      if (chart.getSpreadsheetId() === prodSheetId) {
        const sourceChart = allSourceCharts[chart.getChartId()];
        if (sourceChart) {
          const pos = { h: chart.getHeight(), w: chart.getWidth(), t: chart.getTop(), l: chart.getLeft() };
          chart.remove();
          allSlides[s].insertSheetsChart(sourceChart, pos.l, pos.t, pos.w, pos.h);
          relinkedInThisSession++;
        }
      }

      globalCounter++;

      // Check time limit after every chart
      if (new Date().getTime() - startTime > 260000) { // 4.3 minute threshold
        props.setProperty('RELINK_CHART_INDEX', globalCounter.toString());
        workflowLog(`Pausing chart relinking at index ${globalCounter}. Relinked ${relinkedInThisSession} charts this session.`, "RELINK_CHARTS", startTime);

        // CRITICAL FIX: Throw the pause error so the engine knows to create a trigger
        // and NOT increment the STEP_INDEX.
        throw new Error("PAUSE_REQUESTED");
      }
    }
  }

  // Reset for next time if we finish all charts
  workflowLog(`Successfully relinked all ${globalCounter} charts found.`, "RELINK_CHARTS", startTime);
  props.setProperty('RELINK_CHART_INDEX', '0');
}

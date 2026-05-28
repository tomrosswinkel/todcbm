 
/**
 * @file 03_SheetLogic.gs
 * @description Updated to support split Roll/Archive phases for performance stability.
 * Uses contentsOnly operations to prevent formula overhead during massive data moves.
 */

/**
 * PHASE 1: Updates the calendar week cell only.
 * Returns true if a week change was committed.
 */
function updateWeekAndCheck_(ss) {
  const lookupSheet = ss.getSheetByName(LOOKUP_SHEET_NAME);
  const actualcw = lookupSheet.getRange(ACTUAL_CW_CELL).getValue();
  const cwRange = lookupSheet.getRange(CW_CELL);
  const currentSheetCW = cwRange.getValue();

  if (actualcw !== currentSheetCW) {
    cwRange.setValue(actualcw);
    SpreadsheetApp.flush(); // Commit change to trigger ArrayFormula recalculations
    return true;
  }
  return false;
}

/**
 * PHASE 2: Archives the data and updates visuals.
 * This runs AFTER the calculation pause.
 */
function archiveWeeklyDataAndVisuals_(ss, config) {
  workflowLog("Starting archive and visual update phase.", "WEEKLY_ARCHIVE", null, ss);

  // 1. Move the data in the tables (Value-only move)
  rollWeeksAndPullCurrent(ss);

  // 2. Update Charts in the sheet
  updateCharts(ss);

  // 3. Refresh Slides
  updateSlides(config.slidesId);

  // 4. Send the Chat notification
  sendCardMessageToChat(config.sheetId, config.slidesId, config.context);

  workflowLog("Archive and Visuals complete.", "WEEKLY_ARCHIVE", null, ss);
}

/**
 * High-performance table-rolling logic.
 * Shifts historical data up and captures the "Current" state into the archive.
 */
function rollWeeksAndPullCurrent(ss) {
  console.time("rollWeeksAndPullCurrent");
  const lookupSheet = ss.getSheetByName(LOOKUP_SHEET_NAME);
  const fullcwLabel = lookupSheet.getRange(FULL_CW_CELL).getValue();
  const rowsInTable = lookupSheet.getRange(ROW_CELL).getValue();

  const tableRanges = [
    ["DC Room Burndown", 46, 2], ["Other Statistics", 4, 7], ["Other Statistics", 17, 7], ["Other Statistics", 85, 4], ["Other Statistics", 150, 6],
    ["Rack Burndown", 50, 8], ["Rack Burndown", 110, 8], ["Rack Burndown", 180, 6], ["Rack Burndown", 250, 8], ["Rack Burndown", 320, 8], ["Rack Burndown", 390, 8], ["Rack Burndown", 460, 8],
    ["PDU Burndown", 50, 2], ["PDU Burndown", 110, 2], ["PDU Burndown", 180, 6], ["Breaker Burndown", 50, 7], ["Breaker Burndown", 110, 7], ["Breaker Burndown", 173, 6], ["Breaker Burndown", 185, 7],
    ["Breaker Burndown", 250, 7], ["Breaker Burndown", 320, 7], ["Breaker Burndown", 390, 7], ["Breaker Burndown", 460, 7], ["Cabling Burndown", 50, 5], ["Cabling Burndown", 110, 5], ["Cabling Burndown", 180, 6]
  ];

  tableRanges.forEach(config => {
    const sheet = ss.getSheetByName(config[0]);
    if (!sheet) return;
    const dataStartRow = config[1] + 2;

    // 1. Shift historical data up (e.g. Move Rows 2-5 into 1-4)
    // We use copyTo with contentsOnly for maximum speed and to preserve target formatting.
    const sourceRange = sheet.getRange(dataStartRow + 1, 1, rowsInTable - 1, config[2]);
    const targetRange = sheet.getRange(dataStartRow, 1, rowsInTable - 1, config[2]);
    sourceRange.copyTo(targetRange, {contentsOnly: true});

    // 2. Set the new label (e.g. KW19) in the bottom-most historical slot
    const labelRow = dataStartRow + rowsInTable - 1;
    sheet.getRange(labelRow, 1).setValue(fullcwLabel);

    // 3. Pull "Current" formula results up into the new historical slot
    const currentSourceRow = labelRow + 1;
    if (config[2] > 1) {
      const currentValsRange = sheet.getRange(currentSourceRow, 2, 1, config[2] - 1);
      const historyTargetRange = sheet.getRange(labelRow, 2, 1, config[2] - 1);
      currentValsRange.copyTo(historyTargetRange, {contentsOnly: true});
    }
  });
  console.timeEnd("rollWeeksAndPullCurrent");
}

/**
 * Synchronizes assignments between RackData and WaveAssignment.
 */
function modifyWaveAssignment(ss) {
  console.time("modifyWaveAssignment");
  const waSheet = ss.getSheetByName("WaveAssignment");
  const rackDataSheet = ss.getSheetByName("RackData");

  const rackDataValues = rackDataSheet.getRange("AA2:AC" + rackDataSheet.getLastRow()).getValues();
  const sourceRackMap = new Map(rackDataValues.map(row => [row[0], row]));

  const waValues = waSheet.getDataRange().getValues();
  const waHeader = waValues.shift();
  const existingWaRacks = new Map(waValues.map(row => [row[0], row]));
  const finalWaData = [];
  const newRackIds = [];

  sourceRackMap.forEach((_v, rackId) => { if (!existingWaRacks.has(rackId)) newRackIds.push(rackId); });
  existingWaRacks.forEach((waRow, rackId) => { if (sourceRackMap.has(rackId)) finalWaData.push(waRow); });

  if (newRackIds.length > 0) {
    const templateRow = finalWaData.length > 0 ? finalWaData[0] : new Array(waHeader.length).fill('');
    const newRows = newRackIds.map(rackId => {
      const sourceRow = sourceRackMap.get(rackId);
      const row = [...templateRow];
      row[0] = sourceRow[0]; row[1] = sourceRow[1]; row[2] = sourceRow[2]; row[3] = "Not Assigned";
      return row;
    });
    finalWaData.unshift(...newRows);
  }

  waSheet.clearContents();
  waSheet.getRange(1, 1, 1, waHeader.length).setValues([waHeader]);
  if (finalWaData.length > 0) {
    waSheet.getRange(2, 1, finalWaData.length, waHeader.length).setValues(finalWaData);
  }
  console.timeEnd("modifyWaveAssignment");
}

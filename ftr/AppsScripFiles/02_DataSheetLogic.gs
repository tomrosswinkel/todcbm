 
/**
 * @file 02_DataSheetLogic.gs
 * @description Updated with batch-level logging for better visibility.
 */

/* function modifyRackData(ss, config, startTime) {
  const props = PropertiesService.getScriptProperties();
  const targetSheet = ss.getSheetByName('RackData');

  const archiveFolderId = (config.context === "Production") ? DAVEREPORTS_FOLDER_ID_PROD : DAVEREPORTS_FOLDER_ID_TEST;
  const archiveFolder = DriveApp.getFolderById(archiveFolderId);
  const today = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd");
  const fileName = `Dave Export ${today} Datacenter Rack Stats`;

  const files = archiveFolder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`Source not found: ${fileName}`);

  const sourceSheet = SpreadsheetApp.open(files.next()).getSheets()[0];
  const sourceValues = sourceSheet.getDataRange().getValues();
  const newRowCount = sourceValues.length;

  let lastProcessedDataRow = parseInt(props.getProperty('RACK_DATA_ROW_INDEX') || '0');
  const DATA_BATCH_SIZE = 5000;

  ensureRowBuffer_(targetSheet, newRowCount);

  if (lastProcessedDataRow < newRowCount) {
    for (let i = lastProcessedDataRow; i < newRowCount; i += DATA_BATCH_SIZE) {
      if (new Date().getTime() - startTime > 240000) {
        props.setProperty('RACK_DATA_ROW_INDEX', i.toString());
        throw new Error("PAUSE_REQUESTED");
      }
      let rowsToBatch = Math.min(DATA_BATCH_SIZE, newRowCount - i);
      let batchValues = sourceValues.slice(i, i + rowsToBatch);
      targetSheet.getRange(i + 1, 1, rowsToBatch, 14).setValues(batchValues);

      // PROGRESS LOG
      workflowLog(`RackData: Wrote rows ${i} to ${i + rowsToBatch}`, "UPDATE_RACK_DATA", startTime, ss);
    }
  }

  const currentLastRow = targetSheet.getLastRow();
  if (currentLastRow > newRowCount) {
    targetSheet.getRange(newRowCount + 1, 1, currentLastRow - newRowCount, 14).clearContent();
  }

  props.deleteProperty('RACK_DATA_ROW_INDEX');
}

function modifyPDUData(ss, config, startTime) {
  const props = PropertiesService.getScriptProperties();
  const targetSheet = ss.getSheetByName('PDUData');

  const archiveFolderId = (config.context === "Production") ? DAVEREPORTS_FOLDER_ID_PROD : DAVEREPORTS_FOLDER_ID_TEST;
  const archiveFolder = DriveApp.getFolderById(archiveFolderId);
  const today = Utilities.formatDate(new Date(), "UTC", "yyyy-MM-dd");
  const fileName = `Dave Export ${today} PDUs with power values`;

  const files = archiveFolder.getFilesByName(fileName);
  if (!files.hasNext()) throw new Error(`Source not found: ${fileName}`);

  const sourceSheet = SpreadsheetApp.open(files.next()).getSheets()[0];
  const sourceValues = sourceSheet.getDataRange().getValues();
  const newRowCount = sourceValues.length;

  let lastProcessedDataRow = parseInt(props.getProperty('PDU_DATA_ROW_INDEX') || '0');
  const DATA_BATCH_SIZE = 4000;

  ensureRowBuffer_(targetSheet, newRowCount);

  if (lastProcessedDataRow < newRowCount) {
    for (let i = lastProcessedDataRow; i < newRowCount; i += DATA_BATCH_SIZE) {
      if (new Date().getTime() - startTime > 220000) {
        props.setProperty('PDU_DATA_ROW_INDEX', i.toString());
        throw new Error("PAUSE_REQUESTED");
      }
      let rowsToBatch = Math.min(DATA_BATCH_SIZE, newRowCount - i);
      let batchValues = sourceValues.slice(i, i + rowsToBatch);
      targetSheet.getRange(i + 1, 1, rowsToBatch, 29).setValues(batchValues);

      // PROGRESS LOG
      workflowLog(`PDUData: Wrote rows ${i} to ${i + rowsToBatch}`, "UPDATE_PDU_DATA", startTime, ss);
    }
  }

  const currentLastRow = targetSheet.getLastRow();
  if (currentLastRow > newRowCount) {
    targetSheet.getRange(newRowCount + 1, 1, currentLastRow - newRowCount, 29).clearContent();
  }

  const uniqueRackIds = new Set(sourceValues.slice(1).map(row => row[0]));
  props.deleteProperty('PDU_DATA_ROW_INDEX');

  return { pduRowCount: newRowCount, uniqueRackCount: uniqueRackIds.size };
} */

function ensureRowBuffer_(sheet, requiredRows) {
  const currentMaxRows = sheet.getMaxRows();
  const BUFFER = 500;
  if (currentMaxRows < (requiredRows + BUFFER)) {
    const rowsToAdd = (requiredRows + BUFFER) - currentMaxRows;
    sheet.insertRowsAfter(currentMaxRows, rowsToAdd);
    Logger.log(`Buffer Check: Added ${rowsToAdd} rows to ${sheet.getName()}.`);
  }
}

function modifyCablingData(ss, newRowCount) {
  const sheet = ss.getSheetByName("CablingData");
  ensureRowBuffer_(sheet, newRowCount);
  const currentLastRow = sheet.getLastRow();
  if (currentLastRow > newRowCount) {
    sheet.getRange(newRowCount + 1, 1, currentLastRow - newRowCount, 1).clearContent();
  }
}

function cloneDataSheets(ss) {
  const targets = ["RackData", "PDUData", "CablingData"];
  targets.forEach(name => {
    const src = ss.getSheetByName(name);
    const dest = ss.getSheetByName("Previous" + name);
    if (src && dest) {
      const vals = src.getDataRange().getValues();
      dest.clearContents();
      dest.getRange(1, 1, vals.length, vals[0].length).setValues(vals);
    }
  });
}

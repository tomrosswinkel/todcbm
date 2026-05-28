/**
 * @file 01_FileProcessing.gs
 * @description Processes CSVs directly into the main FTR sheet.
 * Includes mid-batch checkpoints to prevent "Service Spreadsheets" timeouts.
 */

/**
 * Orchestrates the processing of CSVs and writes them directly to the FTR workbook.
 * @param {object} config The configuration object for the run.
 * @param {number} globalStartTime The timestamp from when the engine first started.
 */
function processAndLoadCSVs(config, globalStartTime) {
  const ss = SpreadsheetApp.openById(config.sheetId);
  const props = PropertiesService.getScriptProperties();
  // Use the engine's start time to ensure we respect the global 6/30 minute quota
  const startTime = globalStartTime || new Date().getTime();

  let importStage = parseInt(props.getProperty('IMPORT_STAGE') || '1');
  const importFolder = DriveApp.getFolderById(config.importFolderId);

  // --- STAGE 1: RACK DATA ---
  if (importStage === 1) {
    workflowLog("Starting Stage 1: Rack Data Processing", "IMPORT_AND_PREPARE", startTime, ss);
    const rackCsvFile = findLatestFile_(importFolder, "datacenter_rack_stats_report");
    if (!rackCsvFile) throw new Error("Could not find Rack CSV report.");

    const rackCsvData = Utilities.parseCsv(rackCsvFile.getBlob().getDataAsString(), ";");
    const processedRackData = processRackDataInMemory_(rackCsvData);

    workflowLog(`Rack CSV: ${processedRackData.length} rows detected. Writing to RackData...`, "IMPORT_AND_PREPARE", startTime, ss);

    // Write Rack data directly
    writeRackDataDirectly_(ss, processedRackData);

    props.setProperty('IMPORT_STAGE', '2');

    // Check time immediately after Rack processing (Needs ~3 mins for PDU start)
    if ((new Date().getTime() - startTime) / 1000 > 240) {
      workflowLog("Time low after Rack Import. Pausing for Relay.", "IMPORT_AND_PREPARE", startTime, ss);
      throw new Error("PAUSE_REQUESTED");
    }
    importStage = 2;
  }

  // --- STAGE 2: PDU DATA ---
  if (importStage === 2) {
    workflowLog("Starting Stage 2: PDU Data Processing", "IMPORT_AND_PREPARE", startTime, ss);
    const pduCsvFile = findLatestFile_(importFolder, "pdu_with_power_values_report");
    if (!pduCsvFile) throw new Error("Could not find PDU CSV report.");

    const pduCsvData = Utilities.parseCsv(pduCsvFile.getBlob().getDataAsString(), ";");
    const processedPduData = processPduDataInMemory_(pduCsvData);

    workflowLog(`PDU CSV: ${processedPduData.length} rows detected. Starting batched write...`, "IMPORT_AND_PREPARE", startTime, ss);

    // writePDUDataBatched will handle its own internal relay pauses if it takes too long
    const pduInfo = writePDUDataBatched_(ss, processedPduData, startTime);

    // Success - Clean up progress markers
    props.setProperty('UNIQUE_RACK_COUNT', pduInfo.uniqueRackCount.toString());
    props.deleteProperty('IMPORT_STAGE');
    props.deleteProperty('PDU_IMPORT_ROW_INDEX');

    workflowLog("CSV Data successfully imported.", "IMPORT_AND_PREPARE", startTime, ss);
  }
}

/**
 * Writes PDU data in chunks and checks the clock between batches to prevent Service Timeouts.
 * @private
 */
function writePDUDataBatched_(ss, values, startTime) {
  const targetSheet = ss.getSheetByName('PDUData');
  const props = PropertiesService.getScriptProperties();
  const totalRows = values.length;
  const numCols = values[0].length;

  // Resume from where we left off if a previous relay pause occurred
  let startRow = parseInt(props.getProperty('PDU_IMPORT_ROW_INDEX') || '0');
  const BATCH_SIZE = 3000;

  for (let i = startRow; i < totalRows; i += BATCH_SIZE) {
    // Check clock: If we've been running for more than 7.5 minutes total (450s)
    // We adjust this to 450s because triggers often allow more than 360s.
    // Fix recommended by Gemini: To ensure the PDU write stops exactly when the Engine wants it to, you should also update the time check from 450 down to 240.
    // This ensures that even heavy sub-steps respect your 4-minute safety threshold.
    if ((new Date().getTime() - startTime) / 1000 > 300) {
      props.setProperty('PDU_IMPORT_ROW_INDEX', i.toString());
      workflowLog(`Pausing PDU write at row ${i} of ${totalRows}.`, "IMPORT_AND_PREPARE", startTime, ss);
      throw new Error("PAUSE_REQUESTED");
    }

    let rowsInBatch = Math.min(BATCH_SIZE, totalRows - i);
    targetSheet.getRange(i + 1, 1, rowsInBatch, numCols).setValues(values.slice(i, i + rowsInBatch));
  }

  // Cleanup extra rows only when fully finished
  const currentLastRow = targetSheet.getLastRow();
  if (currentLastRow > totalRows) {
    targetSheet.getRange(totalRows + 1, 1, currentLastRow - totalRows, numCols).clearContent();
  }

  return { uniqueRackCount: new Set(values.slice(1).map(row => row[0])).size };
}

/**
 * Directly writes Rack data to the FTR workbook.
 * @private
 */
function writeRackDataDirectly_(ss, values) {
  const targetSheet = ss.getSheetByName('RackData');
  const newRowCount = values.length;
  const numCols = values[0].length;

  targetSheet.getRange(1, 1, newRowCount, numCols).setValues(values);

  // Clear extra rows
  const currentLastRow = targetSheet.getLastRow();
  if (currentLastRow > newRowCount) {
    targetSheet.getRange(newRowCount + 1, 1, currentLastRow - newRowCount, numCols).clearContent();
  }
}

/**
 * Finds the latest file based on the creation timestamp in Drive.
 * @private
 */
function findLatestFile_(folder, baseName) {
  const files = folder.getFiles();
  let latestFile = null;
  let latestTimestamp = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();

    if (fileName.startsWith(baseName) && fileName.endsWith(".csv")) {
      const createdTime = file.getDateCreated().getTime();
      if (createdTime > latestTimestamp) {
        latestFile = file;
        latestTimestamp = createdTime;
      }
    }
  }
  return latestFile;
}

/**
 * Transforms Rack CSV data: filters out "Virtual" and reorders columns.
 * @private
 */
function processRackDataInMemory_(rawData) {
  const header = rawData.shift();
  const processedData = rawData
    .filter(row => row[1] !== "Virtual")
    .map(row => [row[4], row[0], row[1], row[2], row[3], ...row.slice(5)]);

  processedData.sort((a, b) => a[1].localeCompare(b[1]) || a[2].localeCompare(b[2]) || a[3].localeCompare(b[3]) || a[4].localeCompare(b[4]));

  const newHeader = [header[4], header[0], header[1], header[2], header[3], ...header.slice(5)];
  return [newHeader, ...processedData];
}

/**
 * Transforms PDU CSV data: cleans "-" characters and reorders columns.
 * @private
 */
function processPduDataInMemory_(rawData) {
  const header = rawData.shift();
  const processedData = rawData.map(row => {
    const newRow = [row[2], row[0], row[1], ...row.slice(3)];
    return newRow.map(cell => (cell === "-" ? "" : cell));
  });

  processedData.sort((a, b) => a[2].localeCompare(b[2]));

  const newHeader = [header[2], header[0], header[1], ...header.slice(3)];
  return [newHeader, ...processedData];
}

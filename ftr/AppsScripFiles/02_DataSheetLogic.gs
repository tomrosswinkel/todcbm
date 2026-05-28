/**
 * @file 02_DataSheetLogic.gs
 */

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

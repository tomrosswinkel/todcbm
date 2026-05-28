/**
 * @file 99_Utilities.gs
 */

function generateMarkdownSample(sheetName, numRows = 6) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return Logger.log(`ERROR: ${sheetName} not found.`);
  const vals = sheet.getRange(1, 1, numRows, sheet.getLastColumn()).getValues();
  if (vals.length === 0) return;
  const sanitize = (c) => String(c).replace(/\|/g, '\\|').replace(/\n/g, ' ');
  let md = `\n\n--- ${sheetName} ---\n\n| ${vals[0].map(sanitize).join(' | ')} |\n| ${vals[0].map(() => ':---').join(' | ')} |\n`;
  vals.slice(1).forEach(r => md += `| ${r.map(sanitize).join(' | ')} |\n`);
  Logger.log(md);
}

function runAllSamplers() {
  ["RackData", "PreviousRackData", "PDUData", "PreviousPDUData", "CablingData", "LookupTables"].forEach(n => generateMarkdownSample(n, n === "LookupTables" ? 35 : 6));
}

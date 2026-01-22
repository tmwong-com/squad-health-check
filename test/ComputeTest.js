/**
 * Refill the crosscheck sheet when we delete then install a new compute sheet.
 * A dirty hack for testing purposes only.
 */

const _COLUMN_C = 3
const _COLUMN_D = 4
const _COLUMN_E = 5

const _SHEET_NAME = "Crosscheck sheet"
const _SURVEY_NAME = "Squad Health Check 2025-08-26"

function fillCrosscheckSheet(computeSheetName = COMPUTE_SHEET) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const crosscheckSheet = spreadsheet.getSheetByName(_SHEET_NAME)
  // Find the row in the compute sheet
  // with the static survey results.
  const computeSheet = spreadsheet.getSheetByName(computeSheetName)
  const computeRow = unwrap(computeSheet.createTextFinder(_SURVEY_NAME).findNext()).getRow()
  var formulae = []
  // Fill "Average OK"
  for (var i = 0; i < 22; i++) {
    formulae = formulae.concat([`=EQ(${INTEGERS_TO_COLUMNS[_COLUMN_C + i]}7,'${computeSheetName}'!${INTEGERS_TO_COLUMNS[_COLUMN_D + (i * 2)]}$${computeRow})`])
  }
  crosscheckSheet.getRange("C8:X8").setValues([formulae])
  // Fill "Standard deviation OK"
  formulae = []
  for (var i = 0; i < 22; i++) {
    formulae = formulae.concat([`=EQ(${INTEGERS_TO_COLUMNS[_COLUMN_C + i]}9,'${computeSheetName}'!${INTEGERS_TO_COLUMNS[_COLUMN_E + (i * 2)]}$${computeRow})`])
  }
  crosscheckSheet.getRange("C10:X10").setValues([formulae])
}

function runComputeSheetTests() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheetName = `Compute ${Utilities.getUuid()}`
  createComputeSheet(sheetName)
  triggerCompute(sheetName)
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(_SHEET_NAME))
  fillCrosscheckSheet(sheetName)
}

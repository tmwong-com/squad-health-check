/**
 * Refill the crosscheck sheet when we delete then install a new compute sheet.
 * A dirty hack for testing purposes only.
 */

const _COLUMN_C = 3
const _COLUMN_D = 4
const _COLUMN_E = 5

const _CROSSCHECK_SHEET = "Crosscheck sheet"

function fillCrosscheckSheet(computeSheetName = "Compute 1c93a3b2-63bd-424f-8c95-42f5e06292bc" /*COMPUTE_SHEET*/) {
  const crosscheckSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_CROSSCHECK_SHEET)
  var formulae = []
  // Fill "Average OK"
  for (var i = 0; i < 22; i++) {
    formulae = formulae.concat([`=EQ(${INTEGERS_TO_COLUMNS[_COLUMN_C + i]}7,'${computeSheetName}'!${INTEGERS_TO_COLUMNS[_COLUMN_D + (i * 2)]}$3)`])
  }
  crosscheckSheet.getRange("C8:X8").setValues([formulae])
  // Fill "Standard deviation OK"
  formulae = []
  for (var i = 0; i < 22; i++) {
    formulae = formulae.concat([`=EQ(${INTEGERS_TO_COLUMNS[_COLUMN_C + i]}9,'${computeSheetName}'!${INTEGERS_TO_COLUMNS[_COLUMN_E + (i * 2)]}$3)`])
  }
  crosscheckSheet.getRange("C10:X10").setValues([formulae])
}

function runComputeSheetTests() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheetName = `Compute ${Utilities.getUuid()}`
  createComputeSheet(sheetName)
  triggerCompute(sheetName)
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(_CROSSCHECK_SHEET))
  fillCrosscheckSheet(sheetName)
}

const _COLUMN_C = 3
const _COLUMN_D = 4
const _COLUMN_E = 5

const _CROSSCHECK_SHEET_NAME = "Crosscheck sheet"
const _FIXTURE_SURVEY_NAME = "Squad Health Check 2025-08-26"

/**
 * Create a new compute sheet
 * and test the formulas in the sheet.
 * This test is runnable from within the Apps Script editor,
 * and will delete the sheet
 * if all the unit tests pass.
 * @param {string} surveyTemplateSheetName The name of the survey template sheet to use
 *  when creating the compute sheet.
 */
function runComputeSheetTests(surveyTemplateSheetName = SURVEY_TEMPLATE_SHEET) {
  const computeSheetName = `Compute ${Utilities.getUuid()}`
  const computeSheet = createComputeSheet(computeSheetName, surveyTemplateSheetName)
  updateCompute(computeSheetName)
  test_computeAverageAndSdPerDimension(computeSheetName)
  Logger.log(`Deleting '${computeSheet.getName()}' sheet...`)
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(computeSheet)
}

/**
 * Test that the formulas in a compute sheet
 * correctly calculate the average and standard deviation
 * of survey responses
 * for each dimension,
 * using static survey responses
 * and expected values in the crosscheck sheet.
 * @param {string} computeSheetName The name of the compute sheet to test.
 * @throws {Error} If any computed value does not match the expected cross-check value.
 */
function test_computeAverageAndSdPerDimension(computeSheetName = COMPUTE_SHEET) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const crosscheckSheet = spreadsheet.getSheetByName(_CROSSCHECK_SHEET_NAME)
  // Find the row in the compute sheet
  // with the static survey results.
  const computeSheet = spreadsheet.getSheetByName(computeSheetName)
  const computeRow = unwrap(computeSheet.createTextFinder(_FIXTURE_SURVEY_NAME).findNext()).getRow()
  for (var i = 0; i < Object.keys(SURVEY_DIMENSIONS).length * getSurveySentimentsCount(); i++) {
    const dimension = computeSheet.getRange(`${INTEGERS_TO_COLUMNS[_COLUMN_D + (i * 2)]}1`).getValue()
    // Recall the shape of the compute sheet.
    // For a given survey name,
    // the name of the first dimension and sentiment
    // (by default, "Delivering value: Perception")
    // is cell D1,
    // the average score of the first dimension
    // is column D,
    // and the SD of the first dimension
    // is column E.
    // After that we march along two columns at a time.
    const computeValues = computeSheet.getRangeList(
      [
        `${INTEGERS_TO_COLUMNS[_COLUMN_D + (i * 2)]}${computeRow}`,
        `${INTEGERS_TO_COLUMNS[_COLUMN_E + (i * 2)]}${computeRow}`,
      ]
    ).getRanges()
    // Now the shape of the crosscheck sheet.
    // The crosscheck average
    // is cell C7,
    // and the crosscheck SD
    // is cell C8.
    // After that we march rightwards one column at a time.
    const xcheckValues = crosscheckSheet.getRangeList(
      [
        `${INTEGERS_TO_COLUMNS[_COLUMN_C + i]}7`,
        `${INTEGERS_TO_COLUMNS[_COLUMN_C + i]}8`,
      ]
    ).getRanges()
    Logger.log(`Checking ${dimension}...`)
    for (var j = 0; j < 2; j++) {
      const computeValue = computeValues[j].getValue().toFixed(2)
      const xcheckValue = xcheckValues[j].getValue().toFixed(2)
      if (computeValue != xcheckValue) {
        throw Error(`Computed value ${computeValue} not equal to cross-check value ${xcheckValue}`)
      }
    }
  }
}

/**
 * The primary purpose of these unit tests is to ensure that changes to the template sheet
 * and the script cell coordinate constants remain in sync
 * (e.g., that the DIMENSIONS_{COLUMN, ROW}_START constants point at the correct cell in the sheet).
 */

/**
 * Tests whether one or more elements of an array is blank.
 */
function arrayIsBlank(data) {
  return data.reduce((accumulator, currentValue) => accumulator || (currentValue.length == 0), false)
}

/**
 * Test that we generate aggregation formulae for all dimsentions surveyed.
 */
function test_computeFormulaeCoverAllDimensions(sheet) {
  const dimensionsCount = getSurveyDimensionsCount(sheet)
  const formulae = _createComputeFormulae(dimensionsCount)
  // One formula for count of respondees
  // plus for each dimension a pair of average and SD for each sentiment
  const expectedFormulaeLength = 1 + (dimensionsCount * getSurveySentimentsCount() * 2)
  return (expectedFormulaeLength == formulae.length)
}

/**
 * Test that the DIMENSIONS_{COLUMN, ROW}_START constants point to the start of the dimensions table.
 */
function test_dimensionsTableExists(sheet) {
  var data = sheet.getSheetValues(SURVEY_TEMPLATE_DIMENSIONS_ROW_START - 1, SURVEY_TEMPLATE_DIMENSIONS_COLUMN_START, 1, SURVEY_DIMENSIONS_HEADER.length)[0]
  return (
    data[0] == "Dimension" &&
    data[1] == "Good" &&
    data[2] == "Bad" &&
    data[3] == "Icon URL"
  )
}

/**
 * Test that we get a non-zero count of dimensions.
 */
function test_dimensionsTableHasDimensions(sheet) {
  var count = getSurveyDimensionsCount(sheet)
  return (typeof (count) == "number" && count > 0)
}

/**
 * Test that getting the name and a date from a survey name
 * actually returns a name and a date.
 */
function test_getSurveyResultSheetNameAndDate(_) {
  const nameAndDate = SQUAD_HEALTH_CHECK_SHEET_PREFIX + " 2025-01-01"
  const expected = [nameAndDate, "2025-01-01"]
  const got = _getNameAndDate(nameAndDate)
  return (expected.toString() == _getNameAndDate(SQUAD_HEALTH_CHECK_SHEET_PREFIX + " 2025-01-01").toString())
}


/**
 * Test that all of the dimension table elements are non-blank.
 */
function test_dimensionsTableNoBlanks(sheet) {
  var dimensionsCount = getSurveyDimensionsCount(sheet)
  try {
    var dimensions = sheet.getSheetValues(SURVEY_TEMPLATE_DIMENSIONS_ROW_START, SURVEY_TEMPLATE_DIMENSIONS_COLUMN_START, dimensionsCount, 4)
    return !dimensions.reduce((accumulator, currentValue) => accumulator || arrayIsBlank(currentValue), false)
  }
  catch {
    // If getSheetValues throws an exception,
    // one of the elements in the row is probably an image instead of a string
    // because of an indexing error during sheet creation.
    return false
  }
}

/**
 * Test that the end of the sheet corresponds to the end of the dimensions table.
 */
function test_sheetEndsAfterDimensions(sheet) {
  const expectedSheetLength =
    1 + // The description
    1 + // The dimensions table header
    getSurveyDimensionsCount(sheet) // The dimensions
  return (expectedSheetLength == sheet.getLastRow())
}

/**
 * Test that we get the expected survey name prefix.
 */
function test_surveyNamePrefix(_) {
  return SQUAD_HEALTH_CHECK_SHEET_PREFIX.toLowerCase() == "squad health check"
}

function test_validateDateStringInvalidDate(_) {
  return !validateDate("2005-12-1")
}

function test_validateDateStringValidDate(_) {
  return validateDate("2005-12-01")
}

const TESTS_TEMPLATE_SHEET = [
  test_computeFormulaeCoverAllDimensions,
  test_dimensionsTableExists,
  test_dimensionsTableHasDimensions,
  test_dimensionsTableNoBlanks,
  test_getSurveyResultSheetNameAndDate,
  test_sheetEndsAfterDimensions,
  test_surveyNamePrefix,
  test_validateDateStringInvalidDate,
  test_validateDateStringValidDate,
]

/**
 * Tests of static contents in the template Apps Script file.
 */

function tests_dimensionHeadersAndContentsConsistent(_) {
  return !Object.keys(SURVEY_DIMENSIONS).reduce(
    (accumulator, dimension) => accumulator || SURVEY_DIMENSIONS[dimension].length != SURVEY_DIMENSIONS_HEADER.length, false
  )
}

const TESTS_TEMPLATE_STATIC_CONTENTS = [
  tests_dimensionHeadersAndContentsConsistent,
]

function _runTests(tests, sheet = null) {
  var failed = 0
  for (const test of tests) {
    if (test(sheet)) {
      Logger.log("Test " + test.name + " passed.")
    } else {
      Logger.log("ERROR: Test " + test.name + " failed.")
      failed += 1
    }
  }
  return failed
}

/**
 * Run all tests on an existing survey template.
 * @param {SpreadsheetApp.Sheet} sheet A survey template sheet
 * @return {number} The number of unit tests that failed.
 */
function runSurveyTemplateSheetTests(sheet) {
  var failed = 0
  if (sheet) {
    failed = _runTests(TESTS_TEMPLATE_SHEET, sheet)
  } else {
    failed = 1
  }
  return failed
}

/**
 * Run all tests on the survey template static contents.
 * @return {number} The number of unit tests that failed.
 */
function runSurveyTemplateStaticContentsTests() {
  return _runTests(TESTS_TEMPLATE_STATIC_CONTENTS)
}

/**
 * Create a survey template and run all tests,
 * including the static content tests.
 * This test function is runnable from within the Apps Script editor,
 * and will delete the template
 * if all the unit tests pass.
 */
function runSurveyTemplateTestsInteractively() {
  var failed = runSurveyTemplateStaticContentsTests()
  const sheet = createSurveyTemplateSheet(`${SURVEY_TEMPLATE_SHEET} ${Utilities.getUuid()}`)
  failed += runSurveyTemplateSheetTests(sheet)
  if (failed != 0) {
    throw Error(`${failed} test(s) failed`)
  } else {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet)
  }
}

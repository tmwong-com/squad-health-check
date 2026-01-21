/**
 * The for the installed compute sheet.
 */
const COMPUTE_SHEET = "Compute"

/**
 * The column (and starting row within the column) to populate with the names of response sheets.
 */
const COMPUTE_SHEET_TRIGGER_CELL_NAME_COLUMN = "A"
const COMPUTE_SHEET_TRIGGER_CELL_DATE_COLUMN = "B"
const COMPUTE_SHEET_TRIGGER_CELL_ROW = 3

const _CS_COLUMN_A = 1
const _CS_COLUMN_REPONDENT_COUNT = 3
const _CS_COLUMN_SURVEY_DATE = 2
const _CS_ROW_DATA_START = 3

/**
 * Create a formula to call a Google Sheets statistic function
 * to aggregate string dimension sentiment responses.
 * Internally, a SWITCH subfunction converts string responses to numerical scores.
 * @param {string} statistic A statistic function name.
 * @param {number} column A column from the response table to aggregate.
 *   The column should point to the start of a sentiment responses (i.e., Perception, Trend) for a dimension.
 * @param {!Array<string>=} sentiments A array of string sentiment responses,
 *   ordered from most (at index 0) to least positive.
 * @return {string} A string that looks like
 *   =IF(NOT(OR(ISBLANK(A3),C3 = 0)), AVERAGE(IFERROR(SWITCH(INDIRECT(A3&"!C:C"), 'Survey template'!$G$4, 3, 'Survey template'!$H$4, 2, 'Survey template'!$I$4, 1))),)
 */
function _createFormula(statistic, column, sentiments) {
  return `=IF(NOT(OR(ISBLANK(A3),C3 = 0)), ${statistic}(IFERROR(SWITCH(INDIRECT(A3&"!${column}"), "${sentiments[0]}", 3, "${sentiments[1]}", 2, "${sentiments[2]}", 1))),)`
}

/**
 * Create a statistics formula pair
 * for a sentiment (e.g., "Perception")
 * @param {number} responseTableColumn A column in the survey response table
 *   corresponding to a sentiment for a dimension
 * @param {!Object<Sentiment>=} sentiment A sentiment to aggregate
 * @return {!Array<string>=} An array of formulae to compute the average and SD
 *   for the sentiment
 *     for a dimension represented in the response table column.
 */
function _createStatisticsFormulaPair(responseTableColumn, sentiments) {
  const column = `${INTEGERS_TO_COLUMNS[responseTableColumn]}:${INTEGERS_TO_COLUMNS[responseTableColumn]}`
  return [
    _createFormula("AVERAGE", column, sentiments),
    _createFormula("STDEV.P", column, sentiments)
  ]
}

/**
 * Create the formulae for a compute sheet to aggregate survey responses.
 * @param {number} The count of dimensions surveyed
 * @return {!Array<string>=} An array, where
 *   the first element contains a formula to count the number of respondants to a survey, and
 *   each succeesive set of four elements contains a pair of average and SD values
 *     for each sentiment
 *       for each dimension.
 */
function _createComputeFormulae(dimensionsCount) {
  unwrap(dimensionsCount)
  // First formula counts the number of respondents to a survey.
  var formulae = [`=IF(NOT(ISBLANK(A3)), COUNTIF(INDIRECT($A3&"!B:B"), "*@*"),)`]
  for (var i = 0; i < (dimensionsCount * Object.keys(SURVEY_SENTIMENTS).length);) {
    for (const sentiment in SURVEY_SENTIMENTS) {
      formulae = formulae.concat(_createStatisticsFormulaPair(_CS_COLUMN_REPONDENT_COUNT + i, SURVEY_SENTIMENTS[sentiment]))
      i++
    }
  }
  return formulae
}

/**
 * Populate a compute sheet with headers
 * and survey response processing formulae.
 */
function createComputeSheet(name = COMPUTE_SHEET) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const surveyTemplateSheet = unwrap(spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET))
  Logger.log("Creating compute formulae...")
  const dimensionsCount = getSurveyDimensionsCount(surveyTemplateSheet)
  const formulae = _createComputeFormulae(dimensionsCount)
  // Widen the sheet to accommodate:
  // 1. The survey name
  // 2. The date of the survey
  // 3. The number of respondents to the survey
  // 4. The number of formulae
  const computeSheetLastColumn = 3 + formulae.length
  Logger.log("Creating compute sheet...")
  const computeSheet = spreadsheet.insertSheet(name)
  if (computeSheetLastColumn > computeSheet.getMaxColumns()) {
    // Minus one to account for the first column _before_ the inserted columns.
    computeSheet.insertColumnsAfter(_CS_COLUMN_A, computeSheetLastColumn - computeSheet.getMaxColumns() - 1)
  }
  // It's just easier to set all the column widths the same,
  // then widen a handful as needed.
  computeSheet.setColumnWidths(_CS_COLUMN_A, computeSheet.getMaxColumns(), 40).setColumnWidth(_CS_COLUMN_A, 200).setColumnWidth(2, 100)
  // Set the banding for the whole compute sheet.
  // Remember that a new sheet has 1000 rows by default.
  computeSheet.getRange(`1:${computeSheet.getMaxRows()}`).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY)
  // Create the header rows
  Logger.log("Populating compute sheet headers...")
  var columnIndex = _CS_COLUMN_A
  computeSheet.getRange(1, columnIndex, 1, 3).setValues([["Survey name", "Date", "#"]])
  columnIndex += 3
  const dimensions = surveyTemplateSheet.getSheetValues(SURVEY_TEMPLATE_DIMENSIONS_ROW_START, SURVEY_TEMPLATE_DIMENSIONS_COLUMN_START, dimensionsCount, 1)
  dimensions.forEach(
    (d) => {
      for (const sentiment in SURVEY_SENTIMENTS) {
        // Merge the main header cells for a dimension into a cell for each sentiment,
        // leaving two subheader cells for average and standard deviation for each sentiment.
        // | SentimentA | SentimentB | ...
        // | Avg | SD   | Avg | SD   | ...
        // First the sentiment header...
        computeSheet
          .getRange(1, columnIndex, 1, 2)
          .mergeAcross()
          .setValue(d[0] + ": " + sentiment.toString())
          .setWrap(true)

        // ... then the statistics subheaders
        computeSheet
          .getRange(2, columnIndex, 1, 2)
          .setValues([["Avg", "SD"]])
        columnIndex += 2
      }
    }
  )
  // Populate the first data row with the formulae,
  // starting with the count of respondents,
  // followed by the averages and SD
  // for perception and trend
  // for each dimension.
  // While we're here, set up the formats.
  // The count of respondents is an integer (one hopes),
  // and the stats are to two decimal places.
  Logger.log("Populating compute sheet formulae...")
  const formulaeTemplateRange = computeSheet
    .getRange(3, _CS_COLUMN_REPONDENT_COUNT, 1, formulae.length)
    .setFormulas([formulae]).setNumberFormats([["0"].concat(Array.from({ length: formulae.length - 1 }, (_, i) => "0.00"))])
    .setHorizontalAlignment('right')
  computeSheet.getRange(_CS_ROW_DATA_START, _CS_COLUMN_SURVEY_DATE).setNumberFormat("yyyy-MM-dd")
  // Don't forget that the target fill ranges needs to include the source range.
  // This mix of 0-indexed and 1-indexed structures will be the death of me.
  const formulaeRange = computeSheet.getRange(_CS_ROW_DATA_START, _CS_COLUMN_REPONDENT_COUNT, computeSheet.getMaxRows() - 2, formulae.length)
  formulaeTemplateRange.autoFill(formulaeRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES)
  // Freeze the header rows and first three columns to make scrolling friendly,
  // and protect the sheet to prevent end users from shooting themselves in the foot.
  computeSheet.setFrozenColumns(_CS_COLUMN_REPONDENT_COUNT)
  computeSheet.setFrozenRows(2)
  computeSheet
    .protect()
    .setDescription(`Protect "${name}" against accidental modification`)
    .setWarningOnly(true)
}

/**
 * Trigger the compute sheet by filling a column
 * with all sheet names starting with the response sheet name prefix.
 * @customfunction
 * @param {string} name The name of the compute sheet
 */
function triggerCompute(name = COMPUTE_SHEET) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var namesAndDates = getNamesAndDates(spreadsheet).sort().reverse()
  // Insert the sheet names in the target range to trigger computes...
  var computeSheetTriggerRangeName = (
    name + "!" +
    COMPUTE_SHEET_TRIGGER_CELL_NAME_COLUMN + COMPUTE_SHEET_TRIGGER_CELL_ROW.toString() + ":" +
    COMPUTE_SHEET_TRIGGER_CELL_DATE_COLUMN + (COMPUTE_SHEET_TRIGGER_CELL_ROW + namesAndDates.length - 1).toString()
  )
  var computeSheetTriggerRange = spreadsheet.getRange(computeSheetTriggerRangeName)
  computeSheetTriggerRange.setValues(namesAndDates)
  // ... and clear the contents of any cells below the ranage.
  var clearRangeName = (
    name + "!" +
    COMPUTE_SHEET_TRIGGER_CELL_NAME_COLUMN + (COMPUTE_SHEET_TRIGGER_CELL_ROW + namesAndDates.length).toString() + ":" +
    COMPUTE_SHEET_TRIGGER_CELL_DATE_COLUMN
  )
  var clearRange = spreadsheet.getRange(clearRangeName)
  clearRange.clearContent()
}

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

/**
 * Arrangement of the compute sheet
 * First column: Name of survey response sheet (linked to the survey form)
 * Second column: Date of survey (derived from the name of the survey response sheet)
 * Third column: Number of survey respondents
 * Fourth column: First dimension in the survey
 *
 * Each dimension spans four columns:
 *
 *  | Dimension                 | ...
 *  | Sentiment A | Sentiment B | ...
 *  | Avg  | SD   | Avg  | SD   | ...
 *
 */

const _CS_COLUMN_A = 1
const _CS_COLUMN_DIMENSIONS_START = 4
const _CS_COLUMN_REPONDENT_COUNT = 3
const _CS_COLUMN_SURVEY_DATE = 2
const _CS_ROW_DATA_START = 3

const _LINES_PER_CHART = 4

/**
 * Create a formula to call a Google Sheets statistic function
 * to aggregate string dimension sentiment responses
 * suitable for insertion into the first non-header row
 * of the compute sheet.
 * Internally, a SWITCH subfunction converts string responses to numerical scores.
 * @param {string} statistic A statistic function name.
 * @param {number} column A column from the response table to aggregate.
 *   The column should point to the start of a sentiment responses (i.e., Perception, Trend) for a dimension.
 * @param {!Array<string>=} sentiments An array of string sentiment responses,
 *   ordered from most (at index 0) to least positive.
 * @return {string} An aggregation formula suitable for insertion into the compute sheet,
 *   with placeholders for the statistic, column, and sentiments replaced appropriately.
 *   The returned formulas use semicolons as argument separators
 *   and rely on the locale settings
 *   in the host Google Sheets spreadsheet
 *   to enforce conversion to the appropriate separator.
 */
function _createFormula(statistic, column, sentiments) {
  return `=IF(NOT(OR(ISBLANK(A3); C3 = 0)); ${statistic}(IFERROR(SWITCH(INDIRECT(A3&"!${column}"); "${sentiments[0]}"; 3; "${sentiments[1]}"; 2; "${sentiments[2]}"; 1))); )`
}

/**
 * Create a statistics formula pair
 * for a sentiment (e.g., "Perception")
 * @param {number} responseTableColumn A column in the survey response table
 *   corresponding to a sentiment for a dimension
 * @param {!Object<Sentiment>=} sentiment A sentiment to aggregate
 * @return {!Array<string>=} An array of formulas to compute the average and SD
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
 * Create the formulas for a compute sheet to aggregate survey responses.
 * @param {number} The count of dimensions surveyed
 * @return {!Array<string>=} An array, where
 *   the first element contains a formula to count the number of respondants to a survey, and
 *   each succeesive set of four elements contains a pair of average and SD values
 *     for each sentiment
 *       for each dimension.
 *   The returned formulas use semicolons as argument separators
 *   and rely on the locale settings
 *   in the host Google Sheets spreadsheet
 *   to enforce conversion to the appropriate separator.
 */
function _createComputeFormulas(dimensionsCount) {
  unwrap(dimensionsCount)
  // First formula counts the number of respondents to a survey.
  var formulas = [`=IF(NOT(ISBLANK(A3)); COUNTIF(INDIRECT($A3&"!B:B"); "*@*"); )`]
  for (var i = 0; i < (dimensionsCount * Object.keys(SURVEY_SENTIMENTS).length);) {
    for (const sentiment in SURVEY_SENTIMENTS) {
      formulas = formulas.concat(_createStatisticsFormulaPair(_CS_COLUMN_REPONDENT_COUNT + i, SURVEY_SENTIMENTS[sentiment]))
      i++
    }
  }
  return formulas
}

/**
 * Create a chart of sentiment results from a list of ranges.
 * @param {SpreadsheetApp.Sheet} sheet The sheet into which to embed the new chart
 * @param {string} title A title for the chart
 * @param {!Array<SpreadsheetApp.Range>=} A list of one x-axis label range
 *   and zero or more sentiment value ranges
 * @return {SpreadsheetApp.EmbeddedChart} A new line chart with one line for each range,
 *   embedded in the sheet.
 */
function createChartFromRangeList(sheet, title, ranges) {
  sheet.activate()
  // Don't forget that the first range is the x-axis labels.
  if (ranges.length > _LINES_PER_CHART + 1) {
    throw Error(`Too many ranges to plot on chart: Expected ${_LINES_PER_CHART}, got ${ranges.length}`)
  }
  const builder = sheet
    .newChart()
    .asLineChart()
    .setNumHeaders(1)
    .setOption("series.0.pointShape", "circle")
    .setOption("series.1.pointShape", "triangle")
    .setOption("series.2.pointShape", "square")
    .setOption("series.3.pointShape", "diamond")
    .setOption('treatLabelsAsText', true)
    .setPointStyle(Charts.PointStyle.HUGE)
    .setPosition(1, 1, 0, 0)
    .setRange(0, 3)
    .setTitle(title)
  for (const r in ranges) {
    builder.addRange(ranges[r])
  }
  const chart = builder.build()
  sheet.insertChart(chart)
  return (chart)
}

/**
 * Populate a compute sheet with headers
 * and survey response processing formulas.
 * @param {string} computeSheetName The name for the created compute sheet;
 *   by default, COMPUTE_SHEET.
 * @param {string} surveyTemplateSheetName The name of the survey template sheet;
 *   by default, SURVEY_TEMPLATE_SHEET.
 * @return {SpreadsheetApp.Sheet} The newly created compute sheet.
 */
function createComputeSheet(computeSheetName = COMPUTE_SHEET, surveyTemplateSheetName = SURVEY_TEMPLATE_SHEET) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const surveyTemplateSheet = unwrap(spreadsheet.getSheetByName(surveyTemplateSheetName))
  Logger.log("Creating compute formulas...")
  const dimensionsCount = getSurveyDimensionsCount(surveyTemplateSheet)
  const formulas = _createComputeFormulas(dimensionsCount)
  // Widen the sheet to accommodate:
  // 1. The survey name
  // 2. The date of the survey
  // 3. The number of respondents to the survey
  // 4. The number of formulas
  const computeSheetLastColumn = 3 + formulas.length
  Logger.log(`Creating '${computeSheetName}' sheet...`)
  const computeSheet = spreadsheet.insertSheet(computeSheetName)
  SpreadsheetApp.flush()
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
  // Populate the first data row with the formulas,
  // starting with the count of respondents,
  // followed by the averages and SD
  // for perception and trend
  // for each dimension.
  // While we're here, set up the formats.
  // The count of respondents is an integer (one hopes),
  // and the stats are to two decimal places.
  Logger.log("Populating compute sheet formulas...")
  // Set up the formulas for the first data row.
  const formulasTemplateRange = computeSheet
    .getRange(3, _CS_COLUMN_REPONDENT_COUNT, 1, formulas.length)
    .setFormulas([formulas]).setNumberFormats([["0"].concat(Array.from({ length: formulas.length - 1 }, (_, i) => "0.00"))])
    .setHorizontalAlignment('right')
  computeSheet.getRange(_CS_ROW_DATA_START, _CS_COLUMN_SURVEY_DATE).setNumberFormat("yyyy-MM-dd")
  // Don't forget that the target fill ranges needs to include the source range.
  // This mix of 0-indexed and 1-indexed structures will be the death of me.
  const formulasRange = computeSheet.getRange(_CS_ROW_DATA_START, _CS_COLUMN_REPONDENT_COUNT, computeSheet.getMaxRows() - 2, formulas.length)
  formulasTemplateRange.autoFill(formulasRange, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES)
  // Freeze the header rows and first three columns to make scrolling friendly,
  // and protect the sheet to prevent end users from shooting themselves in the foot.
  computeSheet.setFrozenColumns(_CS_COLUMN_REPONDENT_COUNT)
  computeSheet.setFrozenRows(_CS_ROW_DATA_START - 1)
  computeSheet
    .protect()
    .setDescription(`Protect "${computeSheetName}" against accidental modification`)
    .setWarningOnly(true)
  return computeSheet
}

/**
 * Create chart sheets.
 * Note that this function will _overwrite_ existing chart sheets.
 * @param {string} name The name of the source compute sheet;
 *   by default, COMPUTE_SHEET.
 */
function createChartSheets(computeSheetName = COMPUTE_SHEET) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const computeSheet = unwrap(spreadsheet.getSheetByName(computeSheetName))
  const surveyTemplateSheet = unwrap(spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET))
  const dimensionsCount = getSurveyDimensionsCount(surveyTemplateSheet);
  // For all charts,
  // the x-axis is the dates of the surveys.
  const xAxisLabels = `${INTEGERS_TO_COLUMNS[_CS_COLUMN_SURVEY_DATE]}1:${INTEGERS_TO_COLUMNS[_CS_COLUMN_SURVEY_DATE]}`
  // 0-indexed dimension;
  // 0 is the first in the dimensions table on the survey template sheet
  // Plot _LINES_PER_CHART dimensions per chart,
  // so increment starting dimension by _LINES_PER_CHART.
  for (var dimensionStart = 0; dimensionStart < dimensionsCount; dimensionStart += _LINES_PER_CHART) {
    // Ugly modular math to figure out how many dimensions left to plot.
    var dimensionsOnChart = (dimensionsCount - dimensionStart >= _LINES_PER_CHART) ? _LINES_PER_CHART : dimensionsCount % _LINES_PER_CHART
    // For each dimension,
    // and each sentiment,
    // we have an avg and an SD.
    var dimensionColumnWidth = getSurveySentimentsCount() * 2
    const sentiments = Object.keys(SURVEY_SENTIMENTS)
    for (const s in sentiments) {
      const title = `${sentiments[s]} ${1 + Math.floor(dimensionStart / _LINES_PER_CHART)}`
      const chartSheet = spreadsheet.getSheetByName(title)
      if (chartSheet) {
        Logger.log(`Deleting existing chart "${title}"...`)
        spreadsheet.deleteSheet(chartSheet)
      }
      // For this chunk of _LINES_PER_CHART dimensions,
      // and for this sentiment
      // which column do we start with on the compute sheet.
      var columnStart = _CS_COLUMN_DIMENSIONS_START + dimensionStart * dimensionColumnWidth + s * 2
      const rangeList = [xAxisLabels].concat(
        Array.from(
          // d count the internal dimension within the current chunk of _LINES_PER_CHART dimensions
          { length: dimensionsOnChart }, (_, d) =>
          `${INTEGERS_TO_COLUMNS[columnStart + d * dimensionColumnWidth]}1:${INTEGERS_TO_COLUMNS[columnStart + d * dimensionColumnWidth]}`
        )
      )
      const ranges = computeSheet.getRangeList(rangeList).getRanges()
      Logger.log(`Creating chart "${title}" from ranges ${rangeList}`)
      const chart = createChartFromRangeList(computeSheet, title, ranges)
      spreadsheet
        .moveChartToObjectSheet(chart)
        .setName(title)
    }
  }
}

/**
 * Get chart sheets.
 * @return {!Array<SpreadsheetApp.Sheet>=} An array of chart sheets.
 */
function getChartSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const surveyTemplateSheet = unwrap(spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET))
  const dimensionsCount = getSurveyDimensionsCount(surveyTemplateSheet);
  var chartSheets = []
  for (var dimensionStart = 0; dimensionStart < dimensionsCount; dimensionStart += _LINES_PER_CHART) {
    const sentiments = Object.keys(SURVEY_SENTIMENTS)
    for (const s in sentiments) {
      const title = `${sentiments[s]} ${1 + Math.floor(dimensionStart / _LINES_PER_CHART)}`
      const chartSheet = spreadsheet.getSheetByName(title)
      chartSheets = chartSheet ? chartSheets.concat([chartSheet]) : chartSheets
    }
  }
  return chartSheets
}

/**
 * Update the compute sheet by filling a column
 * with all sheet names starting with the response sheet name prefix.
 * @param {string} name The name of the compute sheet
 */
function updateCompute(name = COMPUTE_SHEET) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var namesAndDates = getNamesAndDates(spreadsheet).sort()
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

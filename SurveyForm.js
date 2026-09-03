/**
 * The format and regular expression for a survey date.
 */
const SURVEY_DATE_REGEXP = /\d{4}\-[0-1][0-9]\-[0-3][0-9]$/

/**
 * Unwrap a nullable value to raise an exception if the value is null.
 * @param {any} value The nullable value.
 * @return {any} The value if not null.
 */
function unwrap(value) {
  if (value == null) {
    throw new Error("Got unexpected null value")
  }
  return value
}

/**
 * 
 * Functions for processing survey names and dates.
 * 
 */

/**
 * Extract the name from a survey result sheet name.
 * @param {string} nameAndDate A survey result sheet name suffixed by a date in YYYY-MM-DD format.
 * @return {!Array<string, date>=} A list of the survey result sheet name and the date.
 */
function _getNameAndDate(nameAndDate) {
  const name = nameAndDate
  const date = name.match(SURVEY_DATE_REGEXP)
  if (date.length < 1) {
    throw Error("Got invalid survey result sheet name: Expected a name with a YYYY-MM-DD date, got " + name)
  }
  return [name, date[0]]
}

/**
 * Get the names and dates of all survey results sheets.
 * Note that we ignore case when comparing name prefixes!
 * @param {Spreadsheet} spreadsheet A spreadsheet.
 * @return {!Array<Array<string, date>>=} An array of lists of survey result sheet names and dates.
 */
function getNamesAndDates(spreadsheet) {
  var namesAndDates = new Array()
  var sheet = spreadsheet.getSheets();
  for (var i = 0; i < sheet.length; i++) {
    if (sheet[i].getName().toLowerCase().startsWith(SQUAD_HEALTH_CHECK_SHEET_PREFIX.toLowerCase())) {
      namesAndDates.push(_getNameAndDate(sheet[i].getName()))
    }
  }
  if (namesAndDates.length < 1) {
    namesAndDates.push(["", ""])
  }
  return namesAndDates
}

/**
 * Make a survey name from a date conforming to the survey date string format.
 * @param {String} date A string date.
 * @returns {String} A survey name with the date as a suffix,
 *  otherwise null if the date is not in the valid format or already exists.
 */
function makeName(dateString) {
  if (!validateDate(dateString)) {
    return null
  }
  const name = SQUAD_HEALTH_CHECK_SHEET_PREFIX + " " + dateString
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  // Ensure that we don't clash with an existing survey name
  const existingNames = getNamesAndDates(spreadsheet)
  for (const existingName of existingNames) {
    if (name == existingName[0]) {
      return null
    }
  }
  return name
}

/**
 * Validate that a date conforms to the survey date string format.
 * @param {String} date A string date.
 * @returns {Boolean} true if the date conforms to the survey date string format,
 *  otherwise false.
 */
function validateDate(dateString) {
  if (!SURVEY_DATE_REGEXP.test(dateString)) {
    return false; // Format does not match
  }
  return !isNaN((new Date(dateString)).getTime()); // Check if the date is valid
}

/**
 * 
 * Functions for generating a new Google Forms survey.
 * 
 */

/**
 * Get the image blob backing an in-cell icon image in the survey template sheet.
 * Google caches in-cell image data, so this works even after the original
 * source URL has gone stale.
 * @param {Sheet} sheet The survey template sheet.
 * @param {number} row The row of the dimension whose icon to read.
 * @return {Blob|null} The icon image blob, or null if the cell holds no image.
 */
function _getInCellIconBlob(sheet, row) {
  // Remember that the icon preview comes _after_ the dimension data.
  const iconColumn = SURVEY_TEMPLATE_DIMENSIONS_COLUMN_START + SURVEY_DIMENSIONS_HEADER.length
  const cellValue = sheet.getRange(row, iconColumn).getValue()
  if (cellValue != null && typeof cellValue.getContentUrl === "function") {
    return UrlFetchApp.fetch(cellValue.getContentUrl()).getBlob()
  }
  return null
}

/**
 * Add questions for a dimension to a survey form.
 * @param {Form} form The survey form to which to add the questions
 * @param {!Array<string>=} A four-element array containing
 *   the dimension description,
 *   the "good" environment description,
 *   the "bad" environment description,
 *   and a URL to a nifty icon
 * @param {boolean} includeIcons If true, include the nifty icons in the survey form.
 * @param {Sheet} templateSheet The survey template sheet holding an in-cell icon
 *   image to fall back on if the icon URL is stale.
 * @param {number} templateRow The row in the template sheet for this dimension.
 */
function _addDimension(form, data, includeIcons = true, templateSheet = null, templateRow = null) {
  const dimension = unwrap(data[0])
  const good = unwrap(data[1])
  const bad = unwrap(data[2])
  Logger.log("Adding dimension: '%s' Good: '%s' Bad: '%s'...", dimension, good, bad)
  const descriptions = {
    Perception: "Good: " + good + "\n" + "Bad: " + bad,
    Trend: SURVEY_TREND_DESCRIPTION
  }
  const dimensionItem = form.addImageItem()
    .setTitle(dimension)
  if (includeIcons) {
    // Icon loading is best-effort: a stale icon URL or unreadable cell image
    // must never prevent survey generation.
    var iconUrl = null
    try {
      var icon = null
      iconUrl = data[3]
      if (iconUrl) {
        try {
          icon = UrlFetchApp.fetch(iconUrl).getBlob()
        }
        catch (e) {
          Logger.log(`WARNING: Unable to load icon at "${iconUrl}", error: ${e}; falling back to in-cell icon image...`)
        }
      }
      if (icon == null && templateSheet != null && templateRow != null) {
        icon = _getInCellIconBlob(templateSheet, templateRow)
      }
      if (icon != null) {
        dimensionItem.setImage(icon)
      }
      else {
        Logger.log(`WARNING: No icon available for dimension "${dimension}", continuing without one...`)
      }
    }
    catch (e) {
      Logger.log(`WARNING: Unable to load icon at "${iconUrl}" or from the template sheet, error: ${e}; continuing without one...`)
    }
  }
  const sentiments = Object.keys(SURVEY_SENTIMENTS)
  for (var s in sentiments) {
    form.addMultipleChoiceItem()
      .setTitle(dimension + ": " + sentiments[s])
      .setHelpText(descriptions[sentiments[s]])
      .setChoiceValues(SURVEY_SENTIMENTS[sentiments[s]])
      .setRequired(true)
      .showOtherOption(false)
  }
}

/**
 * Add a destination response sheet to collect responses to a survey form.
 * @param {Spreadsheet} spreadsheet The destination GSheet in which to create the response sheet.
 * @param {Form} form The survey form.
 * @param {string} name The name of the survey.
 */
function _addDestination(spreadsheet, form, name) {
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId())
  // Flush required to force creation of the new survey responses sheet
  SpreadsheetApp.flush()
  // A new survey responses sheet is always the first sheet in a spreadsheet.
  spreadsheet
    .getSheets()[0]
    .setName(name).activate()
}

/**
 * Generate a new survey form and corresponding response sheet from a template,
 * and put the form in the root of the home GDrive.
 * The survey name is formed from a prefix and a suffix in the template
 * @customfunction
 */
function generateSurveyForm(name) {
  var name = unwrap(name)
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = unwrap(spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET))
  Logger.log("Creating survey form from template sheet " + sheet.getName())
  // Create the form...
  const description = getSurveyDescription(sheet)
  Logger.log("Creating form with name '%s'...", name)
  const form = FormApp.create(name, /* isPublished= */ false)
    .setAllowResponseEdits(true)
    .setCollectEmail(true)
    .setDescription(description)
    .setShowLinkToRespondAgain(false)
  // ... and add each dimension.
  // Remember that sheet rows and columns are 1-indexed,
  // but JavaScript array rows and columns are 0-indexed.
  for (var dimensionRow = SURVEY_TEMPLATE_DIMENSIONS_ROW_START; dimensionRow < SURVEY_TEMPLATE_DIMENSIONS_ROW_START + getSurveyDimensionsCount(sheet); dimensionRow++) {
    var data = sheet.getSheetValues(dimensionRow, SURVEY_TEMPLATE_DIMENSIONS_COLUMN_START, 1, SURVEY_DIMENSIONS_HEADER.length)[0]
    _addDimension(form, data, true, sheet, dimensionRow)
  }
  // Set the destination of survey responses to point back at the current spreadsheet.
  _addDestination(spreadsheet, form, name)
}

/**
 * Run a test to generate a Squad Health Check survey form.
 */
function runGenerateSurveyFormTest() {
  const name = `${SQUAD_HEALTH_CHECK_SHEET_PREFIX} 3008-01-01`
  generateSurveyForm(name)
}

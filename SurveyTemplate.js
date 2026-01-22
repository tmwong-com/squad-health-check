/**
 * Survey template static content.
 * 
 * Apologies for the long lines,
 * but it's programmatically cleaner
 * than trying to remove unnecessary newlines.
 */

/**
 * The name for the installed survey template sheet.
 */
const SURVEY_TEMPLATE_SHEET = "Survey template"

/**
 * A short description of the Squad Health Check.
 * Inserted at the top of every generated survey form.
 */
const SURVEY_DESCRIPTION = `
The Squad Health Check is a way for teams to gauge their perception of productivity, performance, and purpose. Originally developed at Spotify, the check enables teams to identify ways to improve their processes, individual skills, and overall work quality-of-life across eleven dimensions of team sentiment. Spotify's engineers have refined the check over the years, and we recommend reading their original 2014 blog post and follow-up 2023 blog post prior to using the tool. Here, we have adapted the questions from a distillation of the Spotify process by TeamRetro used under a Creative Commons Attribution-ShareAlike license.

Original blog post: https://engineering.atspotify.com/2014/09/squad-health-check-model
Follow-up post: https://engineering.atspotify.com/2023/03/getting-more-from-your-team-health-checks
TeamRetro Squad Health Check: https://www.teamretro.com/health-checks/squad-health-check`.trim()
const SURVEY_DESCRIPTION_COMMENT = "⬅ This description will appear at the top of every generated survey"

/**
 * The row and column of the cell
 * in a installed survey template sheet
 * that will contain the description.
 */
const SURVEY_TEMPLATE_DESCRIPTION_COLUMN = 1
const SURVEY_TEMPLATE_DESCRIPTION_ROW = 1

const SURVEY_DIMENSIONS_HEADER = Object.freeze([
  "Dimension",
  "Good",
  "Bad",
  "Icon URL",
])

/**
 * The dimensions used to quantify team productivity, performance, and purpose.
 * 1. Dimension name
 * 2. Description of "good"
 * 3. Description of "bad"
 * 4. URL to a friendly descriptive icon
 * 4. Dummy field (filled in with the actual icon)
 */
const SURVEY_DIMENSIONS = Object.freeze({
  DELIVERING: [
    "Delivering value",
    "We deliver great stuff! We’re proud of it and our stakeholders are really happy.",
    "We deliver crap. We feel ashamed to deliver it. Our stakeholders hate us.",
    "https://www.teamretro.com/wp-content/uploads/2024/02/diamond.png",
  ],
  EASE: [
    "Ease of release",
    "Releasing is simple, safe, painless and mostly automated.",
    "Releasing is risky, painful, lots of manual work and takes forever.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image11.png",
  ],
  FUN: [
    "Fun",
    "We love going to work and have great fun working together!",
    "Boooooooring…",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image2-150x150.png",
  ],
  HEALTH: [
    "Health of repository",
    "We’re proud of the quality of our repository of reusable artifacts: code, documentation, handbooks, etc. Code is easy to read and properly tested, documentation is up to date, etc.",
    "Our artifacts are a pile of dung and technical/documentation debt is raging out of control.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image13-150x150.png",
  ],
  LEARNING: [
    "Learning",
    "We’re learning lots of interesting stuff all the time!",
    "We never have time to learn anything.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image10-150x150.png",
  ],
  MISSION: [
    "Mission",
    "We know why we are here and we’re really excited about it!",
    "We have no idea why we are here. There’s no high lever picture or focus. Our so-called mission is completely unclear and uninspiring.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image15-150x150.png",
  ],
  PAWNS_OR_PLAYERS: [
    "Pawns or players",
    "We are in control of our own destiny! We decide what to build and how to build it.",
    "We are just pawns in a game of chess with no influence over what we build or how we build it.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image7-150x150.png",
  ],
  SPEED: [
    "Speed",
    "We get stuff done really quickly! No waiting and no delays.",
    "We never seem to get anything done. We keep getting stuck or interrupted. Tasks keep getting stuck on dependencies.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image9-150x150.png",
  ],
  SUITABLE_PROCESS: [
    "Suitable process",
    "Our way of working fits us perfectly!",
    "Our way of working sucks!",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image1-150x150.png",
  ],
  SUPPORT: [
    "Support",
    "We always get great support and help when we ask for it!",
    "We keep getting stuck because we can’t get the support and help that we ask for.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image8-150x150.png",
  ],
  TEAMWORK: [
    "Teamwork",
    "We are a totally gelled super-team with awesome collaboration!",
    "We are a bunch of individuals that neither know nor care about what the other people in the squad are doing.",
    "https://www.teamretro.com/wp-content/uploads/2019/08/image14-150x150.png",
  ],
})

/**
 * The starting row and starting column
 * in the installed survey template sheet
 * of the dimension descriptions,
 * good/bad narrative,
 * and icon URL.
 */
const SURVEY_TEMPLATE_DIMENSIONS_COLUMN_START = 1
const SURVEY_TEMPLATE_DIMENSIONS_ROW_START = 3

/**
 * The sentiments surveyed for each dimension.
 * - Perception: Where the team thinks a particular dimension stands (good or bad)
 * - Trend: Where the team thinks a dimension is headed (improving or deteriorating)
 */
const SURVEY_SENTIMENTS = Object.freeze({
  Perception: [
    "Good 🙂",
    "Neutral 😐",
    "Bad 🙁",
  ],
  Trend: [
    "Improving ↗️",
    "Stable ➡️",
    "Deteriorating ↘️",
  ],
})

/**
 * Create a survey template sheet using predefined static content
 * in the active spreadsheet.
 * @param {string} name The name of the new sheet
 * @return The new survey template sheet in the active spreadsheet.
 */
function createSurveyTemplateSheet(name = SURVEY_TEMPLATE_SHEET) {
  Logger.log(`Creating '${name}' sheet...`)
  const surveyTemplateSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name, 0).activate()
  SpreadsheetApp.flush()
  Logger.log(`Populating "${name}" headers...`)
  var rowIndex = 1
  surveyTemplateSheet
    .autoResizeRows(rowIndex, 1)
    .setColumnWidths(_CS_COLUMN_A, SURVEY_DIMENSIONS_HEADER.length, 200)
    .getRange(rowIndex, _CS_COLUMN_A, 1)
    .setValue(SURVEY_DESCRIPTION)
    .setVerticalAlignment("top")
    .setWrap(true)
  surveyTemplateSheet.getRange(rowIndex, _CS_COLUMN_A, 1, SURVEY_DIMENSIONS_HEADER.length)
    .merge()
  surveyTemplateSheet.getRange(rowIndex, _CS_COLUMN_A + SURVEY_DIMENSIONS_HEADER.length)
    .setValue(SURVEY_DESCRIPTION_COMMENT)
    .setWrap(true)
  rowIndex++
  surveyTemplateSheet
    .autoResizeRows(rowIndex, 1)
    .getRange(rowIndex, _CS_COLUMN_A, 1, SURVEY_DIMENSIONS_HEADER.length + 1)
    // Add one extra column for user-friendly icon image previews.
    .setValues([SURVEY_DIMENSIONS_HEADER.concat(["Icon"])])
    .setBackground("#cccccc") // Grey
    .setFontWeight("bold")
    .setVerticalAlignment("top")
  // Protect the sheet as a whole
  // and freeze the header rows.
  surveyTemplateSheet
    .protect()
    .setDescription(`Protect "${name}" against accidental modification`)
    .setWarningOnly(true)
  surveyTemplateSheet
    .setFrozenRows(rowIndex)
  rowIndex++
  for (const d in SURVEY_DIMENSIONS) {
    Logger.log(`Populating "${name}" dimension "${d}"...`)
    surveyTemplateSheet
      .setRowHeight(rowIndex, 100)
      .getRange(rowIndex, _CS_COLUMN_A, 1, SURVEY_DIMENSIONS_HEADER.length)
      .setValues([SURVEY_DIMENSIONS[d]])
      .setVerticalAlignment("top")
      .setWrap(true)
    try {
      const iconUrl = SURVEY_DIMENSIONS[d][3]
      const icon =
        SpreadsheetApp.newCellImage()
          .setSourceUrl(iconUrl)
          .build();
      surveyTemplateSheet
        // Remember that the icon preview comes _after_ the dimension data.
        .getRange(rowIndex, SURVEY_DIMENSIONS_HEADER.length + 1)
        .setValue(icon)
    }
    catch (e) {
      Logger.log(`WARNING: Unable to load icon at "${iconUrl}"...`)
    }
    rowIndex++
  }
  return surveyTemplateSheet
}

/**
 * Get the survey descrption from the survey template sheet.
 * @param {Sheet} spreadsheet The survey template sheet.
 * @return {string} The survey name prefix, typically is "Squad health check".
 */
function getSurveyDescription(sheet) {
  var sheet = unwrap(sheet)
  return sheet.getRange(SURVEY_TEMPLATE_DESCRIPTION_ROW, SURVEY_TEMPLATE_DESCRIPTION_COLUMN).getValue()
}

/**
 * Get the count of dimensions in the survey template sheet.
 * Assumes that the last non-empty row under the "Dimension" table in the template
 * is the last dimension question template.
 * If it is not, someone has broken the template format.
 * @param {Sheet} spreadsheet The survey template sheet.
 * @return {number} The count of dimensions defined in the template sheet.
 */
function getSurveyDimensionsCount(sheet) {
  var sheet = unwrap(sheet)
  return sheet.getLastRow() - SURVEY_TEMPLATE_DIMENSIONS_ROW_START + 1
}

/**
 * Get the count of surveyed sentiments.
 * @param {Sheet} spreadsheet The survey template sheet.
 *   Ignored since we hardcode the sentiments in the script.
 * @return {number} The count of sentiments defined.
 */
function getSurveySentimentsCount(sheet = null) {
  return Object.keys(SURVEY_SENTIMENTS).length
}

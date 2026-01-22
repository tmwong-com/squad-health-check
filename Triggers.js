const INSTALL_TEMPLATES_MENU_OPTION = "Install templates..."

const INSTALL_CHARTS_MENU_OPTION = "Install charts..."

const GENERATE_SURVEY_MENU_OPTION = "Generate survey form..."

/**
 * Trigger handler when a user installs a Workspace add-on.
 * https://developers.google.com/apps-script/guides/triggers#oninstalle
 */
function onInstall(event) {
  onOpen(event)
}

/**
 * Trigger handler when a user opens a GSheet.
 * https://developers.google.com/apps-script/guides/triggers#onopene
 */
function onOpen(event) {
  const ui = SpreadsheetApp.getUi()
  ui.createAddonMenu()
    .addItem(INSTALL_TEMPLATES_MENU_OPTION, showInstallTemplateSheets.name)
    .addItem(INSTALL_CHARTS_MENU_OPTION, showInstallCharts.name)
    .addItem(GENERATE_SURVEY_MENU_OPTION, showGenerateSurveyFormPrompt.name)
    .addToUi()
}

/**
 * Install the Squad Health Check template sheets for the user.
 */
function showInstallTemplateSheets() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  if (spreadsheet.getSheetByName(COMPUTE_SHEET) || spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET)) {
    ui.alert(
      `Cannot install Squad Health Check template sheets because '${SURVEY_TEMPLATE_SHEET}' and/or '${COMPUTE_SHEET}' sheets already exist. ` +
      `If you wish to reinstall the template sheets, delete or rename the existing sheets first.`
    )
    return
  }
  const button = ui.alert(
    "Install Squad Health Check sheets?",
    ui.ButtonSet.YES_NO,
  )
  switch (button) {
    case ui.Button.YES:
      createSurveyTemplateSheet()
      createComputeSheet()
      triggerCompute()
      break;
    default:
      break;
  }
}

/**
 * Install the chart sheets for the user.
 */
function showInstallCharts() {
  const ui = SpreadsheetApp.getUi();
  const chartSheets = getChartSheets()
  if (chartSheets.length) {
    const button = ui.alert(
      `Installing over existing chart sheets ${chartSheets.map((sheet) => `'${sheet.getName()}'`).join(", ")}. Continue?`,
      ui.ButtonSet.YES_NO,
    )
    switch (button) {
      case ui.Button.NO:
        return;
      default:
        break;
    }
  }
  createChartSheets()
}

/**
 * Prompt the user for a date stamp for a squad health check survey,
 * generate a Google Forms Squad Health Check survey,
 * and set up a survey response sheet.
 */
function showGenerateSurveyFormPrompt() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const ui = SpreadsheetApp.getUi();
  const surveyTemplateSheet = spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET)
  if (runSurveyTemplateSheetTests(surveyTemplateSheet) == 0 || !spreadsheet.getSheetByName(COMPUTE_SHEET)) {
    const response = ui.prompt(
      "Generate a new survey form",
      "Please enter the date in YYYY-MM-DD format for the Squad Health Check:",
      ui.ButtonSet.OK_CANCEL,
    )
    // Process the user's response.
    const button = response.getSelectedButton();
    const date = response.getResponseText();
    switch (button) {
      case ui.Button.OK:
        if (!validateDate(date)) {
          ui.alert("Invalid date '" + date + "'; expected YYYY-MM-DD.")
          break
        }
        const surveyName = makeName(date)
        if (surveyName == null) {
          ui.alert("Date '" + date + "' already in use.")
          break
        }
        generateSurveyForm(surveyName)
        triggerCompute()
        break
      default:
        break
    }
  } else {
    ui.alert(
      `Detected inconsistencies in the Squad Health Check '${SURVEY_TEMPLATE_SHEET}' and/or '${COMPUTE_SHEET}' sheets. ` +
      `Delete any existing copies of the '${SURVEY_TEMPLATE_SHEET}' and '${COMPUTE_SHEET}' sheets, ` +
      `run '${INSTALL_TEMPLATES_MENU_OPTION}', ` +
      `and then run '${GENERATE_SURVEY_MENU_OPTION}' again.`,
      ui.ButtonSet.OK,
    )
  }
}

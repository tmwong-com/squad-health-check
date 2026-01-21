const COPY_TEMPLATE_MENU_OPTION = "Install new template..."
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
    .addItem(COPY_TEMPLATE_MENU_OPTION, showCopyTemplateSpreadsheet.name)
    .addItem(GENERATE_SURVEY_MENU_OPTION, showGenerateSurveyFormPrompt.name)
    .addToUi()
}

/**
 * Create a copy of the Squad Health Check template spreadsheet for the user.
 */
function showCopyTemplateSpreadsheet() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  if (spreadsheet.getSheetByName(COMPUTE_SHEET) || spreadsheet.getSheetByName(SURVEY_TEMPLATE_SHEET)) {
    ui.alert(`ERROR:\nCannot install Squad Health Check templates:\n"${SURVEY_TEMPLATE_SHEET}" and/or "${COMPUTE_SHEET}" sheets already exist.`)
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
 * Prompt the user for a date stamp for a squad health check survey,
 * generate a Google Forms Squad Health Check survey,
 * and set up a survey response sheet.
 */
function showGenerateSurveyFormPrompt() {
  const ui = SpreadsheetApp.getUi();
  const surveyTemplateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SURVEY_TEMPLATE_SHEET)
  if (runSurveyTemplateSheetTests(surveyTemplateSheet) == 0) {
    const response = ui.prompt(
      "Generate a new Squad Health Check survey form",
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
      "First run '" + COPY_TEMPLATE_MENU_OPTION + "', " +
      "then open the new sheet named '" + SQUAD_HEALTH_CHECK_SHEET_PREFIX + "' in 'My Drive' " +
      "and run '" + GENERATE_SURVEY_MENU_OPTION + "'.",
      ui.ButtonSet.OK,
    )
  }
}

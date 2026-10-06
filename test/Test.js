/**
 * Run compute sheet and survey template tests.
 */
function runTests() {
  const surveyTemplateSheet = runSurveyTemplateTestsKeepSheet()
  runComputeSheetTests(surveyTemplateSheet.getName())
  Logger.log(`Deleting '${surveyTemplateSheet.getName()}' sheet...`)
  SpreadsheetApp.getActiveSpreadsheet().deleteSheet(surveyTemplateSheet)
}

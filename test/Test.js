/**
 * Run compute sheet and survey template tests.
 */
function runTests() {
  const surveyTemplateSheetName = runSurveyTemplateTests(true)
  runComputeSheetTests(surveyTemplateSheetName)
}

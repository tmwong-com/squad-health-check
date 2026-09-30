/**
 * Run compute sheet and survey template tests.
 * 
 * Note that this test will leave the crosscheck sheet
 * pointing at the new compute sheet.
 */
function runTests() {
  const surveyTemplateSheetName = runSurveyTemplateTests(true)
  runComputeSheetTests(surveyTemplateSheetName)
}

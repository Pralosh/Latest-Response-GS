// Function to extract the lastest response and append it to the "Latest Response Sheet"
function latestResponse() {
  // Open the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  
  // Get the last row with data
  var lastRow = sheet.getLastRow();
  
  // Get the range of the last row
  var lastRowRange = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
  
  // Get the values from the last row
  var lastRowValues = lastRowRange.getValues();

  // Set a target sheet
  var targetSheetName = 'Latest Response Sheet';
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(targetSheetName);
  
  // Get the next empty row in the destination sheet
  var targetSheetLastRow = targetSheet.getLastRow();
  var nextEmptyRow = targetSheetLastRow + 1;
  
  // Append the latest response to the destination sheet
  targetSheet.getRange(nextEmptyRow, 1, 1, lastRowValues[0].length).setValues(lastRowValues);

}

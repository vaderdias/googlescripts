// For Google Spreadsheets.
// Trims the spaces around the values in the selected column.

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
var sheet = spreadsheet.getActiveSheet();
var currentCell = sheet.getCurrentCell();
Logger.log("Current cell: " + currentCell);
var currentColumn = currentCell.getColumn() - 1;
Logger.log("Current column: " + currentColumn);
var oldValues = sheet.getDataRange().getValues();
var newValues = [];
oldValues.forEach(function(oldRow){
  var newRow = [...oldRow];
  newRow[currentColumn] = oldRow[currentColumn].trim();
  newValues.push(newRow);
});
sheet.clearContents();
sheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);

// For Google Spreadsheets.
// Explodes |-separated values in the selected column
// into multiple rows, copying the values from the other rows.
// After spliting the values, it also trims the spaces around
// each split value.

var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
var sheet = spreadsheet.getActiveSheet();
var currentCell = sheet.getCurrentCell();
Logger.log("Current cell: " + currentCell);
var currentColumn = currentCell.getColumn() - 1;
Logger.log("Current column: " + currentColumn);
var oldValues = sheet.getDataRange().getValues();
var newValues = [];
oldValues.forEach(function(oldRow){

  // separator
  var splitd = oldRow[currentColumn].split("|");
  splitd.forEach(function (item) {
    var newRow = [...oldRow];
    
    // trim
    newRow[currentColumn] = item.trim();
    newValues.push(newRow);
  });
});
sheet.clearContents();
sheet.getRange(1, 1, newValues.length, newValues[0].length).setValues(newValues);

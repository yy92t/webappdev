/** @OnlyCurrentDoc */

function pivotTable() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sourceData = spreadsheet.getRange('3:200'); // Define the source data range

  // Insert a new sheet and activate it
  var newSheet = spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1);
  newSheet.activate();
  newSheet.setHiddenGridlines(true);

  // Create a pivot table
  var pivotTable = newSheet.getRange('A2').createPivotTable(sourceData);

  // Add row groups
  pivotTable.addRowGroup(3);
  pivotTable.addRowGroup(2);

  // Add pivot values with SUM function
  pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM);

  // Set column widths and auto-resize
  newSheet.getRange('A:F').activate();
  newSheet.setColumnWidths(1, 4, 102);
  newSheet.autoResizeColumns(1, 4);
}
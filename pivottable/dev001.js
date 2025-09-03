/** @OnlyCurrentDoc */

function demo() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('3:95').activate();
  var sourceData = spreadsheet.getRange('3:95');
  spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1).activate();
  spreadsheet.getActiveSheet().setHiddenGridlines(true);
  var pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotGroup = pivotTable.addRowGroup(3);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotValue = pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue = pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup = pivotTable.addRowGroup(4);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(2, SpreadsheetApp.PivotTableSummarizeFunction.COUNTA);
  pivotValue = pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup = pivotTable.addRowGroup(4);
};

function demo2() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('3:95').activate();
  var sourceData = spreadsheet.getRange('3:95');
  spreadsheet.insertSheet(spreadsheet.getActiveSheet().getIndex() + 1).activate();
  spreadsheet.getActiveSheet().setHiddenGridlines(true);
  var pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotGroup = pivotTable.addRowGroup(3);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  var pivotValue = pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup = pivotTable.addRowGroup(2);
  pivotTable = spreadsheet.getRange('A1').createPivotTable(sourceData);
  pivotValue = pivotTable.addPivotValue(4, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotValue = pivotTable.addPivotValue(6, SpreadsheetApp.PivotTableSummarizeFunction.SUM);
  pivotGroup = pivotTable.addRowGroup(3);
  pivotGroup = pivotTable.addRowGroup(2);
  spreadsheet.getRange('A:D').activate();
  spreadsheet.getActiveSheet().setColumnWidths(1, 4, 102);
  spreadsheet.getActiveSheet().autoResizeColumns(1, 4);
};
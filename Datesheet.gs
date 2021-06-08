//Datesheet
function createDateSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // The name of the sheet containing the data you want to put in a table.
  var sheetName = "orders_export_1";

  var pivotTableParams = {};
  
  // The source indicates the range of data you want to put in the table.
  // optional arguments: startRowIndex, startColumnIndex, endRowIndex, endColumnIndex
  pivotTableParams.source = {
    sheetId: ss.getSheetByName(sheetName).getSheetId()
  };
  
  // Group rows, the 'sourceColumnOffset' corresponds to the column number in the source range
  // eg: 0 to group by the first column
  pivotTableParams.rows = [{
    sourceColumnOffset: 0,
    sortOrder: "ASCENDING",
    showTotals: true
  }];

  pivotTableParams.columns = [{
    sourceColumnOffset: 45,
    sortOrder: "ASCENDING"
  }];

  // Defines how a value in a pivot table should be calculated.
  pivotTableParams.values = [{
    summarizeFunction: "COUNTA",
    sourceColumnOffset: 0
  }];
    
  // Create a new sheet which will contain our Pivot Table
  var existingsheet = ss.getSheetByName('Datesheet');
  if (existingsheet) {
    ss.deleteSheet(existingsheet);
  }
  var pivotTableSheet = ss.insertSheet();
  pivotTableSheet.setName("Datesheet");
  var pivotTableSheetId = pivotTableSheet.getSheetId();
  
  // Add Pivot Table to new sheet
  // Meaning we send an 'updateCells' request to the Sheets API
  // Specifying via 'start' the sheet where we want to place our Pivot Table
  // And in 'rows' the parameters of our Pivot Table
  var request = {
    "updateCells": {
      "rows": {
        "values": [{
          "pivotTable": pivotTableParams
        }]
      },
      "start": {
        "sheetId": pivotTableSheetId
      },
      "fields": "pivotTable"
    }
  };

  Sheets.Spreadsheets.batchUpdate({'requests': [request]}, ss.getId());
}

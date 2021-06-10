function DateSheetNew() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  var itt = spreadsheet.getSheetByName('DateSheetNew');
  if (itt) {
    spreadsheet.deleteSheet(itt);
  }
  var newsheet = spreadsheet.insertSheet()
  newsheet.setName("DateSheetNew")
  var targetsheetName = "DateSheetNew"
  var sourcesheetName = "orders_export copy"
  var sourcesheet = spreadsheet.getSheetByName(sourcesheetName)
  var copy_to_range = sourcesheet.getRange("A1:A" + sourcesheet.getLastRow())
  var copy_to_range2 = sourcesheet.getRange("AT1:AT" + sourcesheet.getLastRow())
  var targetsheet = spreadsheet.getSheetByName(targetsheetName)
  var paste_to_range = targetsheet.getRange("A1:A")
  var paste_to_range2 = targetsheet.getRange("B1:B")
  copy_to_range.copyTo(paste_to_range)
  copy_to_range2.copyTo(paste_to_range2)

  var data = targetsheet.getDataRange().getValues()
  var newData = []
  for (var i in data) {
    var row = data[i];
    var duplicate = false;
    for (var j in newData) {
      if (row.join() == newData[j].join()) {
        duplicate = true;
      }
    }
    if (!duplicate) {
      newData.push(row);
    }
  }
  targetsheet.clearContents()
  targetsheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData)
  targetsheet.getRange("A1:B" + targetsheet.getLastRow()).createFilter()
  
}

function fillValuesDown() {
  var spreadsheet = SpreadsheetApp.getActiveSheet()
  var currentRange = spreadsheet.getRange("AT2:AT" + spreadsheet.getLastRow())
  var newRange = []
  var newFillValue
  currentRange.getValues().map(function(value) {
    if (value[0] !== '') {
      newFillValue = value[0]
      newRange.push([newFillValue])
    } else {
      newRange.push([newFillValue])
    }
  })
  currentRange.setValues(newRange)
  spreadsheet.createTextFinder('spice_delivery_date: ').replaceAllWith('')
}


// Google Apps Script to take a value in column AT and fill it down until a new value shows up in a cell for Google Sheets

function fillDateColumn() {

//To get the active spreadsheet 
  var spreadsheet = SpreadsheetApp.getActiveSheet()
  
//Active Range is the column range where I want to fill down values based on value that is given above 
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
  
//Setting values that we got
  currentRange.setValues(newRange)
  
//Changing string from spice_delivery_date to blank
  spreadsheet.createTextFinder('spice_delivery_date: ').replaceAllWith('')
}


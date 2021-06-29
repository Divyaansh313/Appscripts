function DeliveryPersonAssignment() {
  const id = '18b93ptrClKBgZVz--tSIccoUPGE4JCOKFPBE3TUoqjs'
  var spreadsheet = SpreadsheetApp.openById(id)
  var currentRange = spreadsheet.getRange("AH2:AH" + spreadsheet.getLastRow())
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
  var currentRange2 = spreadsheet.getRange("AO2:AO" + spreadsheet.getLastRow())
  var newRange2 = []
  var newFillValue2
  currentRange2.getValues().map(function(value) {
    if (value[0] !== '') {
      newFillValue2 = value[0]
      newRange2.push([newFillValue2])
    } else {
      newRange2.push([newFillValue2])
    }
  })
  currentRange2.setValues(newRange2)

    
  var currentRange3 = spreadsheet.getRange("AI2:AI" + spreadsheet.getLastRow())
  var newRange3 = []
  var newFillValue3
  currentRange3.getValues().map(function(value) {
    if (value[0] !== '') {
      newFillValue3 = value[0]
      newRange3.push([newFillValue3])
    } else {
      newRange3.push([newFillValue3])
    }
  })
  currentRange3.setValues(newRange3)

  var currentRange4 = spreadsheet.getRange("AJ2:AJ" + spreadsheet.getLastRow())
  var newRange4 = []
  var newFillValue4
  currentRange4.getValues().map(function(value) {
    if (value[0] !== '') {
      newFillValue4 = value[0]
      newRange4.push([newFillValue4])
    } else {
      newRange4.push([newFillValue4])
    }
  })
  currentRange4.setValues(newRange4)

  var currentRange5 = spreadsheet.getRange("C2:C" + spreadsheet.getLastRow())
  var newRange5 = []
  var newFillValue5
  currentRange5.getValues().map(function(value) {
    if (value[0] !== '') {
      newFillValue5 = value[0]
      newRange5.push([newFillValue5])
    } else {
      newRange5.push([newFillValue5])
    }
  })
  currentRange5.setValues(newRange5)

  var currentRange6 = spreadsheet.getRange("L2:L" + spreadsheet.getLastRow())
  var newRange6 = []
  var newFillValue6
  currentRange6.getValues().map(function(value) {
    if (value[0] !== '') {
      newFillValue6 = value[0]
      newRange6.push([newFillValue6])
    } else {
      newRange6.push([newFillValue6])
    }
  })
  currentRange6.setValues(newRange6)

  var itt = spreadsheet.getSheetByName('Delivery Person Sheet');
  if (itt) {
    spreadsheet.deleteSheet(itt);
  }
  var newsheet = spreadsheet.insertSheet()
  newsheet.setName("Delivery Person Sheet")
  var targetsheetName = "Delivery Person Sheet"
  var sourcesheetName = "orders_export copy"
  var sourcesheet = spreadsheet.getSheetByName(sourcesheetName)
  var copy_to_range = sourcesheet.getRange("A1:A" + sourcesheet.getLastRow())
  var copy_to_range2 = sourcesheet.getRange("AI1:AI" + sourcesheet.getLastRow())
  var copy_to_range3 = sourcesheet.getRange("AJ1:AJ" + sourcesheet.getLastRow())
  var copy_to_range4 = sourcesheet.getRange("AH1:AH" + sourcesheet.getLastRow())
  var copy_to_range5 = sourcesheet.getRange("AO1:AO" + sourcesheet.getLastRow())
  var copy_to_range6 = sourcesheet.getRange("C1:C" + sourcesheet.getLastRow())
  var copy_to_range7 = sourcesheet.getRange("L1:L" + sourcesheet.getLastRow())
  var targetsheet = spreadsheet.getSheetByName(targetsheetName)
  var paste_to_range = targetsheet.getRange("A1:A")
  var paste_to_range2 = targetsheet.getRange("B1:B")
  var paste_to_range3 = targetsheet.getRange("C1:C")
  var paste_to_range4 = targetsheet.getRange("D1:D")
  var paste_to_range5 = targetsheet.getRange("E1:E")
  var paste_to_range6 = targetsheet.getRange("G1:G")
  var paste_to_range7 = targetsheet.getRange("H1:H")

  copy_to_range.copyTo(paste_to_range)
  copy_to_range2.copyTo(paste_to_range2)
  copy_to_range3.copyTo(paste_to_range3)
  copy_to_range4.copyTo(paste_to_range4)
  copy_to_range5.copyTo(paste_to_range5)
  copy_to_range6.copyTo(paste_to_range6)
  copy_to_range7.copyTo(paste_to_range7)

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

  targetsheet.getRange(1,1).setValue('Order')
  targetsheet.getRange(1,2).setValue('Name')
  targetsheet.getRange(1,3).setValue('Address')
  targetsheet.getRange(1,4).setValue('Phone')
  targetsheet.getRange(1,5).setValue('Zip Code')
  targetsheet.getRange(1,6).setValue('Rider')
  targetsheet.getRange(1,7).setValue('Payment Status')
  targetsheet.getRange(1,8).setValue('Amount') 
  targetsheet.getRange(1,9).setValue('Region')

  targetsheet.deleteRow(2)
  var sourcesheetName2 = "PersonZip"
  var sourcesheet2 = spreadsheet.getSheetByName(sourcesheetName2)
  var search_range = sourcesheet2.getRange("A2:A" + sourcesheet2.getLastRow())
  var search_values = search_range.getValues()
  var sourcesheetName3 = "Delivery Person Sheet"
  var sourcesheet3 = spreadsheet.getSheetByName(sourcesheetName3)
  var compare_range = sourcesheet3.getRange("E2:E" + sourcesheet3.getLastRow())
  var compare_values = compare_range.getValues()
  var count_Sham = 0
  var count_anil = 0 
  var count_Devender = 0 
  var count_satnam1 = 0 
  var count_satnam2 = 0 
  var count_satnam3 = 0 
  var count_Rahul = 0 
  var count_Rajkumar = 0 
  var count_Sanjay = 0 
  var count_Mukesh = 0
  var count_Prakash = 0

  for (var k=0;k < search_values.length;k++){
    for (var l=0;l < compare_values.length;l++){
      if (search_values[k][0] == compare_values[l][0]){
        value = sourcesheet2.getRange('B' + (k+2)).getValue()
        var getCellValue =  sourcesheet3.getRange('F' + (l+2))
        if (value == "Sham" && count_Sham <= 10){
          count_Sham +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Raj Kumar" && count_Rajkumar <= 10){
          count_Rajkumar +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Sanjay" && count_Sanjay <= 10){
          count_Sanjay +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Mukesh" && count_Mukesh <= 10){
          count_Mukesh +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Prakash" && count_Prakash <= 10){
          count_Prakash +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Anil ji" && count_anil <= 15){
          count_anil +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Devender" && count_Devender <= 15){
          count_Devender +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Satnam ji 1" && count_satnam1 <= 15){
          count_satnam1 +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Satnam ji 2" && count_satnam2 <= 15){
          count_satnam2 +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Satnam ji 3" && count_satnam3 <= 15){
          count_satnam3 +=1  
          getCellValue.setValue(value)
        }
        else if (value == "Rahul" && count_Rahul <= 15){
          count_Rahul +=1  
          getCellValue.setValue(value)
        }
        else {
          getCellValue.setValue("Undefined")
        }
        var getCellValue2 = sourcesheet3.getRange('I' + (l+2))
        value2 = sourcesheet2.getRange('C' + (k+2)).getValue()
        getCellValue2.setValue(value2)
      }
    }
  }
  spreadsheet.createTextFinder('paid').replaceAllWith('PAID')
  spreadsheet.createTextFinder('pending').replaceAllWith('COD')

  var finalrange = spreadsheet.getRange("A2:I" + spreadsheet.getLastRow())
  finalrange.sort(6)
  finalrange.sort(9)

  var range_for_deletion = sourcesheet3.getRange("G2:G" + sourcesheet3.getLastRow())
  var values_for_deletion = range_for_deletion.getValues()
  for (var m=0; m < values_for_deletion.length; m++){
    if (values_for_deletion[m][0] == "PAID"){
      var getDeleteCellValue = sourcesheet3.getRange('H' + (m+2))
      getDeleteCellValue.clearContent()
    }
  }


}

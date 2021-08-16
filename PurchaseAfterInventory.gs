function createInventoryAfterPurchase(){
  const id = '18b93ptrClKBgZVz--tSIccoUPGE4JCOKFPBE3TUoqjs';
  var this_spreadsheet = SpreadsheetApp.openById(id);
  var targetsheet = this_spreadsheet.getSheetByName("purchase sheet after inventory");
  if (targetsheet){
    this_spreadsheet.deleteSheet(targetsheet);
  }
  var newsheet = this_spreadsheet.insertSheet();
  newsheet.setName("purchase sheet after inventory");

  var sheetFrom = this_spreadsheet.getSheetByName("Dated Purchase sheet");
  var sheetTo = this_spreadsheet.getSheetByName("Inventory sheet");

  var rangefromPurchaseSheet = sheetFrom.getRange("B2:B" + sheetFrom.getLastRow());
  var valuesfromPurchaseSheet = rangefromPurchaseSheet.getValues();

  var rangefromInventorySheet = sheetTo.getRange("B3:B" + sheetTo.getLastRow());
  var valuesfromInventorySheet = rangefromInventorySheet.getValues();
  var itemsfromPurchaseSheet=sheetFrom.getRange("A2:A"+sheetFrom.getLastRow());
  var itemValuesfromPurchaseSheet=itemsfromPurchaseSheet.getValues();
  var itemsfromInventorySheet=sheetTo.getRange("A3:A"+sheetTo.getLastRow());
  var itemValuesfromInventorySheet=itemsfromInventorySheet.getValues();

  var purchaseAfterInventorySheet = this_spreadsheet.getSheetByName("purchase sheet after inventory")

  for (var i=0; i < itemValuesfromPurchaseSheet.length; i++){
    for(var j=0;j<itemValuesfromInventorySheet.length;j++){
      if(itemValuesfromPurchaseSheet[i][0]==itemValuesfromInventorySheet[j][0]){
    var valuetoPaste = valuesfromPurchaseSheet[i][0] - valuesfromInventorySheet[j][0];
    var stringToPaste=valuetoPaste.toString();
    var getRangeTopaste = purchaseAfterInventorySheet.getRange("B" + (i+3));
    getRangeTopaste.setValue(stringToPaste);
  }
  }
  }
  var itemRangeFromInventoryaPurchase=purchaseAfterInventorySheet.getRange("B3:B"+purchaseAfterInventorySheet.getLastRow());
  var itemValuesFromInventoryaPurchase=itemRangeFromInventoryaPurchase.getValues();
  
  for(var k=0;k<itemValuesFromInventoryaPurchase.length;k++){
  if(itemValuesFromInventoryaPurchase[k][0]==''){
   var getCellValueRange=sheetFrom.getRange("B"+(k+2));
   var getCellValue=getCellValueRange.getValue();
   getRangeTopaste = purchaseAfterInventorySheet.getRange("B" + (k+3));
    getRangeTopaste.setValue(getCellValue);
  }
  }

  purchaseAfterInventorySheet.getRange(1,1).setValue('Lineitem name')
  purchaseAfterInventorySheet.getRange(1,2).setValue('Items to purchase after checking inventory')

  var copy_to_range_for_items = sheetFrom.getRange("A2:A" + sheetFrom.getLastRow());
  var paste_to_range_for_items = purchaseAfterInventorySheet.getRange("A3:A" + sheetTo.getLastRow());
  copy_to_range_for_items.copyTo(paste_to_range_for_items);

}

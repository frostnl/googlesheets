var gCurrentSpreadsheet = null;

function getCurrentSpreadsheet(){
  if (gCurrentSpreadsheet == null){
    var dataSheetId = '';
    
    if (dataSheetId == ''){
      gCurrentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    } else{
      gCurrentSpreadsheet = SpreadsheetApp.openById(dataSheetId);
    }
    
  } 
  
   return gCurrentSpreadsheet;
}

function getSheet(name,spreadsheet)
{
  if (spreadsheet == null) {
    var spreadsheet = getCurrentSpreadsheet();
  } else {
    var spreadsheet = SpreadsheetApp.openById(spreadsheet);
  }
    
    return spreadsheet.getSheetByName(name);
}


function monthDiff(d1, d2) {
    var months;
    months = (d2.getFullYear() - d1.getFullYear()) * 12;
    months -= d1.getMonth() + 1;
    months += d2.getMonth();
    return months <= 0 ? 0 : months;
}



function getRangeFromCellStart(rng) {
  var sheet = rng.getSheet(); 
  var col = 1;
  
  while (rng.offset(0,col+1).getValue() != '') {
      col++;
  } 
  
  return sheet.getRange(rng.getRow(),rng.getColumn(),sheet.getLastRow(),col); 
}



function OutputDates(fromDate, toDate, sheet, outputRange) {
  var DAY_MILLIS = 24 * 60 * 60 * 1000;
  var dates = [];
  var output = [];
  var date = new Date();
  date = fromDate
  
  if (toDate >= fromDate) {
    
    while (date <= toDate){
      dates = [];
      dates.push(date);
      output.push(dates);
      date =  new Date(date.getTime() + DAY_MILLIS);
    }
  
    var range = sheet.getRange(outputRange.getRow(), outputRange.getColumn(),output.length,1)
    
    range.setValues(output);
    
  }
  
  
}

/**
 * Return data folder id for saving csv files
 */
function getDataFolder(folderName){
  // find row
  var folderSheet = getCurrentSpreadsheet().getSheetByName('Folders')
  var startCell = folderSheet.getRange('B3')
  var rowOffset = 0;
  var bFound = false;
  
  
  var data = folderSheet.getRange("B3:D20").getValues();
  for (var i = 0; i <data.length; i++){
      if (data[i][0] == folderName){
        rowOffset = i;
        bFound = true;
      }
  }
  
  if (bFound) {  
    var folderId = data[rowOffset][2];
    var folder = DriveApp.getFolderById(folderId);
  }
  else{
    var folder = null;
  }
  
  return folder;
}



// RETURN a template structure
function getView(viewName) {

  var viewSheet = getCurrentSpreadsheet().getSheetByName('Views')
  var startCell = viewSheet.getRange('B3')
  var rowOffset = -1;
  
  var data = viewSheet.getRange("C3:C100").getValues();
  for (var i = 0; i <data.length; i++){
      if (data[i][0] == viewName){
        rowOffset = i;
      }
  }
  
  if (rowOffset >= 0) {
    var view = {
       Client: startCell.offset(rowOffset, 0).getValue(),
       viewName: startCell.offset(rowOffset, 1).getValue(),
       accountId: startCell.offset(rowOffset, 2).getValue(),
       propertyId: startCell.offset(rowOffset, 3).getValue(),
       viewId: startCell.offset(rowOffset, 4).getValue()
      }
  }
  else {
    var view = null;
  }
  
return view;
}








function testGetDataFolder() {
   Logger.log(getDataFolder()); 
}


//function getSheet(sheetName) {
//  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);   
//}


function removeFolderFileWithName(folder, fileName) {
  // check is a folder
  
  files = folder.getFilesByName(fileName);
  
  while (files.hasNext()) {
    var fileId = files.next().getId();
    Drive.Files.remove(fileId);
  }
  
}


function getRowVisibleArray(sheet) {

  var spreadsheetId = sheet.getParent().getId();
  var sheetId = sheet.getSheetId();
  
  // get last row that has content
  var endRow = sheet.getLastRow() + 1;
  
  // limit what's returned from the API
  var fields = "sheets(data(rowMetadata(hiddenByFilter)),properties/sheetId)";
  var sheets = Sheets.Spreadsheets.get(spreadsheetId, {fields: fields}).sheets;  
  
  var hiddenRows = [];  
  
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].properties.sheetId == sheetId) {
      var data = sheets[i].data;
      var rows = data[0].rowMetadata;
      for (var j = 0; j < rows.length; j++) {
        if (rows[j].hiddenByFilter){
          hiddenRows.push(j+1);
        }
      }
    }
  }
  
    var rowVisible = new Array(endRow);
    for (var i = 1; i <=endRow; i++) {
      rowVisible[i] = true;
    }
    for (var i = 0; i<hiddenRows.length; i++){
      hiddenIndex = hiddenRows[i];
      if (hiddenIndex >= 0 && hiddenIndex < endRow) {
        rowVisible[hiddenIndex] = false;
      }
    }   
  
  return rowVisible;  
} 

function isArray(obj) {
  if (typeof obj === "undefined") {
    return false
  }
  return obj.constructor == Array;
}

function createReport(data, reportDefinition){
  var outputArray = [];
  
  for (var i=0; i<data.length; i++){
    var item = data[i];
    var outputLine = [];
    for (var j=0; j<reportDefinition.length; j++){
      outputLine[j] = item[reportDefinition[j]];  
    }
    outputArray.push(outputLine);
  }
  
  return outputArray;
}

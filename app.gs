// change these values
var name = ""; // leave as empty string if you are getting value from a cell in ActiveSheet 
var nameRange = "<Cell range>"; // spell out the cell where you are reading the file name as, eg. "A1"
var folderId = "<Desination folder Id>"; // Google Id of destination folder, eg. "0B2_h6nTAN7gBU3ZLRFVkLmxVYkU"



/** 
* Main Function
*
*
*/

function exportFunction (rangeLastRow, rangeLastCol, rangeFirstRow, rangeFirstCol) {
   
    // initialize params 
    var folder = DriveApp.getFolderById(folderId);
    var sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
    var sourceSheet = sourceSpreadsheet.getActiveSheet();
    var sheetName = sourceSheet.getSheetName();
    var pdfName = name || sourceSheet.getRange(nameRange).getDisplayValue();
       
    // duplicate Spreadsheet 
    var tempSpreadsheet = DriveApp.getFileById(sourceSpreadsheet.getId()).makeCopy("tmp_for_pdf", folder);
    var tempSpreadsheet = SpreadsheetApp.open(tempSpreadsheet);
   
    // replace values, first param is spreadsheet file 
    replaceValues(tempSpreadsheet, sheetName, rangeLastRow, rangeLastCol, rangeFirstRow, rangeFirstCol); 
   
   
    // deletes the rest 
    var sheets = tempSpreadsheet.getSheets(); //return sheet array 
    for (var index in sheets) {
      if (sheets[index].getSheetName() !== sheetName) {
        tempSpreadsheet.deleteSheet(sheets[index]);
      }
    }

    saveToPdf(tempSpreadsheet, pdfName, folder); // first param must be a Spreadsheet file object
    
    DriveApp.getFileById(tempSpreadsheet.getId()).setTrashed(true); // delete tempSpreadsheet 
 
 }// close exportRange 



/** 
*Called functions
*
*/

  // replaces all values to DisplayValues 
  // theSpreadsheet is the spreadsheet file, sheetName is sheet level 
function replaceValues (theSpreadsheet, sheetName, rangeLastRow, rangeLastCol, rangeFirstRow, rangeFirstCol) {
  var theSheet = theSpreadsheet.getSheetByName(sheetName);
  var range = theSheet.getRange(rangeFirstRow, rangeFirstCol, rangeLastRow, rangeLastCol);
  var copyValues = range.getDisplayValues();
  
  range.clearContent();
  
  if (rangeLastRow < theSheet.getLastRow()){
    theSheet.getRange(rangeLastRow+1, 1, theSheet.getLastRow(), theSheet.getLastColumn()).clear();
  } 
  if(rangeLastCol < theSheet.getLastColumn()){
    theSheet.getRange(1, rangeLastCol+1, theSheet.getLastRow(), theSheet.getLastColumn()).clear();
  }
  if(rangeFirstRow > 1){
    theSheet.getRange(1, 1, rangeFirstRow-1, theSheet.getLastColumn()).clear();
  }
  if(rangeFirstCol > 1){
    theSheet.getRange(1, 1, theSheet.getLastRow(), rangeFirstCol-1).clear();
  }
  
  
  range.setValues(copyValues);
}
  
  // saveToPdf
  // note that theSheet is a file level param 
function saveToPdf(theSpreadsheet, pdfName, folder) {    
  var theBlob = theSpreadsheet.getBlob().getAs('application/pdf').setName(pdfName);
  var newPDF = folder.createFile(theBlob);
}



/**  
* UI Buttons
*
*/

// create a UI button, trigger onOpen 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Save to PDF')
  .addItem('Export Range', 'exportRange')
  .addItem('Export Fullsheet', 'exportSheet')
  .addToUi();
}

// UI functions
function exportRange() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeFirstRow = sourceSheet.getActiveRange().getRow();
  var rangeFirstCol = sourceSheet.getActiveRange().getColumn();
  var rangeLastRow = sourceSheet.getActiveRange().getLastRow();
  var rangeLastCol = sourceSheet.getActiveRange().getLastColumn();
  exportFunction(rangeLastRow, rangeLastCol, rangeFirstRow, rangeFirstCol);
}


function exportSheet() {
  var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangeLastRow = sourceSheet.getLastRow();
  var rangeLastCol = sourceSheet.getLastColumn();
  exportFunction(rangeLastRow, rangeLastCol, 1, 1);
}
// end UI 

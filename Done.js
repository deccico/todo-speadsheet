function moveSelectedRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  
  var targetSheetName = "Done"; 
  var targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("Target sheet '" + targetSheetName + "' not found!");
    return;
  }
  
  if (sourceSheet.getName() === targetSheetName) {
     SpreadsheetApp.getUi().alert("You are already on the destination sheet.");
     return;
  }
  
  var activeRange = sourceSheet.getActiveRange();
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  
  var numCols = sourceSheet.getLastColumn(); 
  if (numCols === 0) return; 
  
  // Get the data from the selected rows
  var dataRange = sourceSheet.getRange(startRow, 1, numRows, numCols);
  var dataToMove = dataRange.getValues();
  
  // Insert new blank rows directly under Row 1
  targetSheet.insertRowsAfter(1, numRows);
  
  // Paste the original data starting at Row 2
  targetSheet.getRange(2, 1, numRows, numCols).setValues(dataToMove);
  
  // NEW: Strip the bold formatting inherited from Row 1 across the whole inserted row
  targetSheet.getRange(2, 1, numRows, targetSheet.getMaxColumns()).setFontWeight("normal");
  
  // Add the current date and time to Column N (Column 14)
  var timestamp = new Date();
  targetSheet.getRange(2, 14, numRows, 1).setValue(timestamp);
  
  // Delete the original rows from the source sheet
  sourceSheet.deleteRows(startRow, numRows);
}

function moveToProjectsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  
  var targetSheetName = "Todo Projects"; 
  var targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert("Target sheet '" + targetSheetName + "' not found!");
    return;
  }
  
  if (sourceSheet.getName() === targetSheetName) {
     SpreadsheetApp.getUi().alert("You are already on the destination sheet.");
     return;
  }
  
  var activeRange = sourceSheet.getActiveRange();
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  
  var numCols = sourceSheet.getLastColumn(); 
  if (numCols === 0) return; 
  
  // Get the data from the selected rows
  var dataRange = sourceSheet.getRange(startRow, 1, numRows, numCols);
  var dataToMove = dataRange.getValues();
  
  // Insert new blank rows directly under Row 4
  targetSheet.insertRowsAfter(4, numRows);
  
  // Paste the original data starting at Row 2
  targetSheet.getRange(5, 1, numRows, numCols).setValues(dataToMove);
  
  // NEW: Strip the bold formatting inherited from Row 1 across the whole inserted row
  targetSheet.getRange(5, 1, numRows, targetSheet.getMaxColumns()).setFontWeight("normal");
    
  // Delete the original rows from the source sheet
  sourceSheet.deleteRows(startRow, numRows);
}

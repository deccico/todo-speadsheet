function moveSelectedRows() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getActiveSheet();
  
  var targetSheetName = "Done"; 
  var targetSheet = ss.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    ui.alert("Target sheet '" + targetSheetName + "' not found!");
    return;
  }
  
  if (sourceSheet.getName() === targetSheetName) {
     ui.alert("You are already on the destination sheet.");
     return;
  }
  
  // Prompt the user for total hours
  var response = ui.prompt(
    'Total Hours',
    'Please enter the total hours for the completed task(s):',
    ui.ButtonSet.OK_CANCEL
  );

  // Check if the user clicked OK
  if (response.getSelectedButton() == ui.Button.OK) {
    var hoursInput = response.getResponseText().trim();
    var hoursValue = parseFloat(hoursInput);
    
    // Validate that the input is actually a number
    if (isNaN(hoursValue) || hoursInput === "") {
      ui.alert('Invalid input', 'You must enter a numeric value for hours. Operation cancelled.', ui.ButtonSet.OK);
      return;
    }
  } else {
    // User clicked Cancel or closed the dialog box
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
  
  // Strip the bold formatting inherited from Row 1 across the whole inserted row
  targetSheet.getRange(2, 1, numRows, targetSheet.getMaxColumns()).setFontWeight("normal");
  
  // Add the current date and time to Column N (Column 14)
  var timestamp = new Date();
  targetSheet.getRange(2, 14, numRows, 1).setValue(timestamp);
  
  // Add the numeric hours value to Column O (Column 15)
  targetSheet.getRange(2, 15, numRows, 1).setValue(hoursValue);
  
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

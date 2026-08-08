function taskIsDone() {
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

  // Determine active row range before prompting
  var activeRange = sourceSheet.getActiveRange();
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  
  var numCols = sourceSheet.getLastColumn(); 
  if (numCols === 0) return; 

  // Read Column E (Column 5) values for the selection and calculate the default total
  var colEValues = sourceSheet.getRange(startRow, 5, numRows, 1).getValues();
  var defaultHours = 0;
  for (var i = 0; i < colEValues.length; i++) {
    var val = parseFloat(colEValues[i][0]);
    if (!isNaN(val)) {
      defaultHours += val;
    }
  }

  // Prompt user with the Column E default value indicated
  var response = ui.prompt(
    'Total Hours',
    'Please enter total hours for the task(s) (Press OK to accept default of ' + defaultHours + ' from Column E):',
    ui.ButtonSet.OK_CANCEL
  );

  var hoursValue;

  // Check if the user clicked OK
  if (response.getSelectedButton() == ui.Button.OK) {
    var hoursInput = response.getResponseText().trim();
    
    // If user leaves field blank, use default from Column E; otherwise validate their input
    if (hoursInput === "") {
      hoursValue = defaultHours;
    } else {
      hoursValue = parseFloat(hoursInput);
      if (isNaN(hoursValue)) {
        ui.alert('Invalid input', 'You must enter a numeric value for hours. Operation cancelled.', ui.ButtonSet.OK);
        return;
      }
    }
  } else {
    // User clicked Cancel or closed the dialog box
    return;
  }
  
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

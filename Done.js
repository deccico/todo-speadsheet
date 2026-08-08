/**
 * Strictly parses a single hours value. Unlike parseFloat, trailing garbage
 * is rejected ("3 hours", "1,5"), so a typo cannot silently record the wrong
 * number of hours.
 *
 * @param {*} value A cell value or typed text.
 * @return {?number} The parsed number, or null when not a valid number.
 */
function parseHoursNumber(value) {
  if (typeof value === "number") {
    return isNaN(value) ? null : value;
  }
  var trimmed = String(value).trim();
  if (trimmed === "") {
    return null;
  }
  var num = Number(trimmed);
  return isNaN(num) ? null : num;
}

/**
 * Maps the Column E cells of the selection to one hours estimate per row.
 *
 * @param {Array<Array>} columnEValues Values as returned by Range.getValues().
 * @return {Array<?number>} One estimate per row; null where the cell holds no number.
 */
function extractHoursEstimates(columnEValues) {
  var estimates = [];
  for (var i = 0; i < columnEValues.length; i++) {
    estimates.push(parseHoursNumber(columnEValues[i][0]));
  }
  return estimates;
}

/**
 * Total of the Column E estimates, rounded to 2 decimals — shown in the
 * prompt as the default. Null when the selection has no numeric estimate.
 *
 * @param {Array<Array>} columnEValues Values as returned by Range.getValues().
 * @return {?number} The default total hours, or null when unavailable.
 */
function computeDefaultHours(columnEValues) {
  var estimates = extractHoursEstimates(columnEValues);
  var total = 0;
  var found = false;
  for (var i = 0; i < estimates.length; i++) {
    if (estimates[i] !== null) {
      total += estimates[i];
      found = true;
    }
  }
  if (!found) {
    return null;
  }
  return Math.round(total * 100) / 100;
}

/**
 * Builds the prompt message, mentioning the Column E default when one exists.
 *
 * @param {?number} defaultHours Total from computeDefaultHours().
 * @return {string} The message for ui.prompt().
 */
function buildHoursPromptMessage(defaultHours) {
  if (defaultHours === null) {
    return 'Please enter the total hours for the completed task(s):';
  }
  return 'Please enter the total hours for the completed task(s), or leave ' +
    'blank to use each row\'s own Column E estimate (total: ' + defaultHours + '):';
}

/**
 * Turns the prompt response into a write plan for Column O of the Done sheet.
 * Blank input accepts the Column E estimates: each row receives its own
 * estimate (rows without one are left empty). A typed number is written to
 * every moved row, matching the script's historical behavior.
 *
 * @param {string} hoursInput Raw text the user typed in the prompt.
 * @param {Array<Array>} columnEValues Column E values of the selection.
 * @return {?{perRowHours: Array<?number>}|{sameHours: number}} The plan, or
 *     null when the input is invalid (non-numeric, or blank with no default).
 */
function resolveHoursPlan(hoursInput, columnEValues) {
  var trimmed = String(hoursInput).trim();
  if (trimmed === "") {
    if (computeDefaultHours(columnEValues) === null) {
      return null;
    }
    return { perRowHours: extractHoursEstimates(columnEValues) };
  }
  var typed = parseHoursNumber(trimmed);
  if (typed === null) {
    return null;
  }
  return { sameHours: typed };
}

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

  var activeRange = sourceSheet.getActiveRange();
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();

  var numCols = sourceSheet.getLastColumn();
  if (numCols === 0) return;

  // Offer a default taken from the estimates in Column E of the selection
  var colEValues = sourceSheet.getRange(startRow, 5, numRows, 1).getValues();
  var defaultHours = computeDefaultHours(colEValues);

  // Prompt the user for total hours
  var response = ui.prompt('Total Hours', buildHoursPromptMessage(defaultHours), ui.ButtonSet.OK_CANCEL);

  // User clicked Cancel or closed the dialog box
  if (response.getSelectedButton() != ui.Button.OK) {
    return;
  }

  var hoursPlan = resolveHoursPlan(response.getResponseText(), colEValues);
  if (hoursPlan === null) {
    ui.alert('Invalid input', 'You must enter a numeric value for hours. Operation cancelled.', ui.ButtonSet.OK);
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

  // Add the hours to Column O (Column 15): per-row estimates when the
  // default was accepted, otherwise the typed value on every row
  var hoursRange = targetSheet.getRange(2, 15, numRows, 1);
  if (hoursPlan.perRowHours) {
    var hoursColumn = [];
    for (var i = 0; i < hoursPlan.perRowHours.length; i++) {
      hoursColumn.push([hoursPlan.perRowHours[i] === null ? "" : hoursPlan.perRowHours[i]]);
    }
    hoursRange.setValues(hoursColumn);
  } else {
    hoursRange.setValue(hoursPlan.sameHours);
  }

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

  // Paste the original data starting at Row 5
  targetSheet.getRange(5, 1, numRows, numCols).setValues(dataToMove);

  // Strip the bold formatting inherited from the header across the whole inserted row
  targetSheet.getRange(5, 1, numRows, targetSheet.getMaxColumns()).setFontWeight("normal");

  // Delete the original rows from the source sheet
  sourceSheet.deleteRows(startRow, numRows);
}

// Allow the pure helpers to be unit-tested under Node; Apps Script has no
// module object, so this block is inert when deployed.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    parseHoursNumber: parseHoursNumber,
    extractHoursEstimates: extractHoursEstimates,
    computeDefaultHours: computeDefaultHours,
    buildHoursPromptMessage: buildHoursPromptMessage,
    resolveHoursPlan: resolveHoursPlan
  };
}

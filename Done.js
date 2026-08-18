// Column layout shared by the task sheets and the Done sheet
var ESTIMATE_COLUMN = 5;   // Column E: the hours estimate on task sheets
var TIMESTAMP_COLUMN = 14; // Column N: completion date/time on the Done sheet
var HOURS_COLUMN = 15;     // Column O: recorded hours on the Done sheet

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
 * Builds the prompt message, mentioning the Column E default when one
 * exists. Kept short so it is quick to read.
 *
 * @param {?number} defaultHours Total from computeDefaultHours().
 * @return {string} The message for ui.prompt().
 */
function buildHoursPromptMessage(defaultHours) {
  if (defaultHours === null) {
    return 'Enter the hours spent:';
  }
  return 'Enter the hours spent, or leave blank to use the Column E estimate ' +
    '(total: ' + defaultHours + '):';
}

/**
 * Turns the prompt response into a write plan for Column O of the Done sheet.
 * Blank input accepts the Column E estimates: each row receives its own
 * estimate (rows without one are left empty). A typed number is written to
 * every moved row.
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

/**
 * Expands a resolved hours plan to one value per row.
 *
 * @param {{perRowHours: Array<?number>}|{sameHours: number}} hoursPlan
 *     Plan from resolveHoursPlan().
 * @param {number} numRows Number of rows being moved.
 * @return {Array<?number>} The hours to record, one entry per row.
 */
function planToPerRowHours(hoursPlan, numRows) {
  if (hoursPlan.perRowHours) {
    return hoursPlan.perRowHours;
  }
  var hours = [];
  for (var i = 0; i < numRows; i++) {
    hours.push(hoursPlan.sameHours);
  }
  return hours;
}

/**
 * Builds the complete row values written to the Done sheet: the copied data
 * padded out to Column O, with the completion timestamp in Column N and the
 * hours in Column O. Producing everything up front lets the move use a
 * single setValues call instead of three separate writes, which is
 * noticeably quicker.
 *
 * @param {Array<Array>} dataRows The selected rows' values.
 * @param {*} timestamp The completion date/time.
 * @param {Array<?number>} perRowHours Hours per row; null leaves the cell empty.
 * @return {Array<Array>} The values for the Done sheet, one array per row.
 */
function buildDoneRows(dataRows, timestamp, perRowHours) {
  var width = Math.max(dataRows[0].length, HOURS_COLUMN);
  var out = [];
  for (var i = 0; i < dataRows.length; i++) {
    var row = dataRows[i].slice();
    while (row.length < width) {
      row.push("");
    }
    row[TIMESTAMP_COLUMN - 1] = timestamp;
    row[HOURS_COLUMN - 1] = perRowHours[i] === null ? "" : perRowHours[i];
    out.push(row);
  }
  return out;
}

/**
 * Inserts the given values directly under targetSheet's header block and
 * strips the bold formatting inherited from the row above. Shared by the
 * Done and Projects moves; the caller deletes the source rows.
 *
 * @param {number} headerRows Rows to insert after (1 for Done, 4 for Projects).
 * @param {Array<Array>} rowValues The values to write, one array per row.
 */
function insertRowsBelowHeader(targetSheet, headerRows, rowValues) {
  var numRows = rowValues.length;
  targetSheet.insertRowsAfter(headerRows, numRows);
  targetSheet.getRange(headerRows + 1, 1, numRows, rowValues[0].length).setValues(rowValues);
  targetSheet.getRange(headerRows + 1, 1, numRows, targetSheet.getMaxColumns()).setFontWeight("normal");
}

/**
 * Moves the selected row(s) to the Done sheet: prompts for the hours spent
 * (blank accepts the Column E estimates), inserts the rows at the top of
 * the Done sheet with the completion timestamp in Column N and the hours in
 * Column O, and deletes the originals. Fully native ui.prompt — no
 * HtmlService and no google.script.run, so it is fast and unaffected by
 * Google's multi-account PERMISSION_DENIED bug.
 */
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

  // One read serves both the Column E default and the later copy
  var dataToMove = sourceSheet.getRange(startRow, 1, numRows, numCols).getValues();
  var colEValues = [];
  for (var i = 0; i < dataToMove.length; i++) {
    colEValues.push([numCols >= ESTIMATE_COLUMN ? dataToMove[i][ESTIMATE_COLUMN - 1] : ""]);
  }
  var defaultHours = computeDefaultHours(colEValues);

  // Prompt the user for the hours (Enter accepts, blank uses the default)
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

  // Single write: data + timestamp (Column N) + hours (Column O) together
  var doneRows = buildDoneRows(dataToMove, new Date(), planToPerRowHours(hoursPlan, numRows));
  insertRowsBelowHeader(targetSheet, 1, doneRows);

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

  var dataToMove = sourceSheet.getRange(startRow, 1, numRows, numCols).getValues();
  insertRowsBelowHeader(targetSheet, 4, dataToMove);

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
    resolveHoursPlan: resolveHoursPlan,
    planToPerRowHours: planToPerRowHours,
    buildDoneRows: buildDoneRows
  };
}

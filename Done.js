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
 * Builds the dialog label, explaining the Column E suggestion when one
 * exists. The suggested number itself is prefilled in the input, so the
 * label does not repeat it.
 *
 * @param {?number} defaultHours Total from computeDefaultHours().
 * @return {string} The label for the hours dialog.
 */
function buildHoursPromptMessage(defaultHours) {
  if (defaultHours === null) {
    return 'Please enter the total hours for the completed task(s):';
  }
  return 'Please enter the total hours for the completed task(s). The ' +
    'suggested value is the Column E estimate total; accept it unchanged ' +
    'and each moved row keeps its own estimate:';
}

/**
 * Escapes text for safe embedding in HTML content or attribute values.
 *
 * @param {*} value The value to escape.
 * @return {string} The escaped HTML.
 */
function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

/**
 * Builds the self-contained HTML for the hours dialog. Apps Script's
 * ui.prompt cannot prefill its text box, so an HtmlService dialog is used
 * instead: the Column E suggestion is written into the input and
 * pre-selected, making OK/Enter accept it while typing replaces it. Built
 * from a string so the project stays a single pasteable file. Submitting
 * calls completeTaskMove(input, context) on the server via
 * google.script.run; a validation error is shown inside the dialog.
 *
 * @param {?number} defaultHours Total from computeDefaultHours().
 * @param {{sheetName: string, startRow: number, numRows: number}} context
 *     Snapshot of the selection, echoed back to completeTaskMove().
 * @return {string} The dialog HTML.
 */
function buildHoursDialogHtml(defaultHours, context) {
  // <-escape so a hostile sheet name cannot close the <script> block
  var contextJson = JSON.stringify(context).replace(/</g, "\\u003c");
  return '' +
    '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 13px;">' +
    '<p style="margin-top: 0;">' + escapeHtml(buildHoursPromptMessage(defaultHours)) + '</p>' +
    '<form id="hours-form">' +
    '<input type="text" id="hours" autofocus ' +
    'style="width: 100%; box-sizing: border-box; font-size: 13px; padding: 4px;" ' +
    'value="' + (defaultHours === null ? '' : escapeHtml(defaultHours)) + '">' +
    '<p id="error" style="color: #cc0000; min-height: 1.2em; margin: 8px 0;"></p>' +
    '<div style="text-align: right;">' +
    '<button type="button" id="cancel">Cancel</button> ' +
    '<button type="submit" id="ok">OK</button>' +
    '</div>' +
    '</form>' +
    '</div>' +
    '<script>' +
    'var context = ' + contextJson + ';' +
    'var input = document.getElementById("hours");' +
    'window.setTimeout(function() { input.focus(); input.select(); }, 0);' +
    'function setBusy(busy) {' +
    '  document.getElementById("ok").disabled = busy;' +
    '  document.getElementById("cancel").disabled = busy;' +
    '}' +
    'document.getElementById("cancel").onclick = function() { google.script.host.close(); };' +
    'document.getElementById("hours-form").onsubmit = function(event) {' +
    '  event.preventDefault();' +
    '  setBusy(true);' +
    '  document.getElementById("error").textContent = "";' +
    '  google.script.run' +
    '    .withSuccessHandler(function() { google.script.host.close(); })' +
    '    .withFailureHandler(function(err) {' +
    '      document.getElementById("error").textContent = (err && err.message) ? err.message : String(err);' +
    '      setBusy(false);' +
    '    })' +
    '    .completeTaskMove(input.value, context);' +
    '};' +
    '</script>';
}

/**
 * Turns the dialog input into a write plan for Column O of the Done sheet.
 * Accepting the suggestion — submitting the prefilled total unchanged, or
 * blank input — applies the Column E estimates: each row receives its own
 * estimate (rows without one are left empty). Any other number is written
 * to every moved row, matching the script's historical behavior.
 *
 * @param {string} hoursInput Raw text the user submitted in the dialog.
 * @param {Array<Array>} columnEValues Column E values of the selection.
 * @return {?{perRowHours: Array<?number>}|{sameHours: number}} The plan, or
 *     null when the input is invalid (non-numeric, or blank with no default).
 */
function resolveHoursPlan(hoursInput, columnEValues) {
  var trimmed = String(hoursInput).trim();
  var defaultHours = computeDefaultHours(columnEValues);
  if (trimmed === "") {
    if (defaultHours === null) {
      return null;
    }
    return { perRowHours: extractHoursEstimates(columnEValues) };
  }
  var typed = parseHoursNumber(trimmed);
  if (typed === null) {
    return null;
  }
  if (defaultHours !== null && typed === defaultHours) {
    return { perRowHours: extractHoursEstimates(columnEValues) };
  }
  return { sameHours: typed };
}

/**
 * Menu entry point: validates the selection and opens the hours dialog with
 * the Column E suggestion prefilled. showModalDialog does not block, so the
 * dialog calls completeTaskMove() to perform the actual move.
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

  if (sourceSheet.getLastColumn() === 0) return;

  // Offer a suggestion taken from the estimates in Column E of the selection
  var colEValues = sourceSheet.getRange(startRow, 5, numRows, 1).getValues();
  var defaultHours = computeDefaultHours(colEValues);

  var html = buildHoursDialogHtml(defaultHours, {
    sheetName: sourceSheet.getName(),
    startRow: startRow,
    numRows: numRows
  });
  ui.showModalDialog(
    HtmlService.createHtmlOutput(html).setWidth(360).setHeight(210),
    'Total Hours');
}

/**
 * Called by the hours dialog via google.script.run. Moves the snapshotted
 * selection to the Done sheet, writing the completion timestamp to Column N
 * and the resolved hours to Column O. Throws on invalid input so the dialog
 * shows the message and stays open for a correction.
 *
 * @param {string} hoursInput Raw text from the dialog input.
 * @param {{sheetName: string, startRow: number, numRows: number}} context
 *     Selection snapshot taken when the dialog was opened.
 */
function completeTaskMove(hoursInput, context) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(context.sheetName);
  var targetSheet = ss.getSheetByName("Done");

  if (!sourceSheet || !targetSheet) {
    throw new Error("The source or Done sheet no longer exists.");
  }

  var startRow = context.startRow;
  var numRows = context.numRows;
  if (!(startRow >= 1) || !(numRows >= 1)) {
    throw new Error("The selection is no longer valid. Please try again.");
  }

  var numCols = sourceSheet.getLastColumn();
  if (numCols === 0) return;

  var colEValues = sourceSheet.getRange(startRow, 5, numRows, 1).getValues();
  var hoursPlan = resolveHoursPlan(hoursInput, colEValues);
  if (hoursPlan === null) {
    throw new Error("You must enter a numeric value for hours.");
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
  // suggestion was accepted, otherwise the typed value on every row
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
    escapeHtml: escapeHtml,
    buildHoursDialogHtml: buildHoursDialogHtml,
    resolveHoursPlan: resolveHoursPlan
  };
}

// Column layout shared by the task sheets and the Done sheet
var ESTIMATE_COLUMN = 5;   // Column E: the hours estimate on task sheets
var TIMESTAMP_COLUMN = 14; // Column N: completion date/time on the Done sheet
var HOURS_COLUMN = 15;     // Column O: hours input on task sheets, recorded hours on Done

// The pending two-press move, remembered between button presses
var MOVE_MARKER_KEY = "pendingDoneMove";
var MOVE_MARKER_TTL_SECONDS = 300;

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
 * Maps one column of cells from the selection to one number per row.
 *
 * @param {Array<Array>} columnValues Values as returned by Range.getValues().
 * @return {Array<?number>} One value per row; null where the cell holds no number.
 */
function extractHoursEstimates(columnValues) {
  var estimates = [];
  for (var i = 0; i < columnValues.length; i++) {
    estimates.push(parseHoursNumber(columnValues[i][0]));
  }
  return estimates;
}

/**
 * Finds the first Column O cell that holds something non-blank that is not a
 * number ("4h", "1,5") — such a cell blocks the move until fixed.
 *
 * @param {Array<Array>} oValues Column O values of the rows to move.
 * @return {?{row: number, value: *}} Zero-based row and offending value, or
 *     null when every cell is numeric or blank.
 */
function findInvalidHours(oValues) {
  for (var i = 0; i < oValues.length; i++) {
    if (parseHoursNumber(oValues[i][0]) === null && String(oValues[i][0]).trim() !== "") {
      return { row: i, value: oValues[i][0] };
    }
  }
  return null;
}

/**
 * Decides what pressing the Done button should do for the selection, from
 * its Column E (estimates) and Column O (hours input) values:
 *
 * - "invalid": some Column O cell holds a non-number — nothing can happen
 *   until it is fixed.
 * - "move": every row already has numeric hours in Column O — move now.
 * - "stage": suggest hours first. stagedHours holds, per row, the existing
 *   Column O number, else the Column E estimate, else null; stagedCount is
 *   how many blank cells a suggestion exists for.
 *
 * @param {Array<Array>} eValues Column E values of the selection.
 * @param {Array<Array>} oValues Column O values of the selection.
 * @return {{action: string}} The plan, shaped as described above.
 */
function planStagedMove(eValues, oValues) {
  var invalid = findInvalidHours(oValues);
  if (invalid) {
    return { action: "invalid", row: invalid.row, value: invalid.value };
  }
  var estimates = extractHoursEstimates(eValues);
  var hours = extractHoursEstimates(oValues);
  var stagedHours = [];
  var stagedCount = 0;
  var allNumeric = true;
  for (var i = 0; i < hours.length; i++) {
    if (hours[i] !== null) {
      stagedHours.push(hours[i]);
      continue;
    }
    allNumeric = false;
    if (estimates[i] !== null) {
      stagedHours.push(estimates[i]);
      stagedCount++;
    } else {
      stagedHours.push(null);
    }
  }
  if (allNumeric) {
    return { action: "move", perRowHours: stagedHours };
  }
  return { action: "stage", stagedHours: stagedHours, stagedCount: stagedCount };
}

/**
 * Builds the toast shown after suggestions were written to Column O.
 *
 * @param {number} stagedCount How many cells received a suggestion.
 * @return {string} The toast message.
 */
function buildStageToast(stagedCount) {
  if (stagedCount === 0) {
    return "Enter the hours in Column O (now selected), then press Done again to move.";
  }
  return "Suggested hours from Column E are in Column O — adjust them if needed, then press Done again to move.";
}

/**
 * Builds the refusal toast for a non-numeric Column O cell.
 *
 * @param {number} sheetRow 1-based sheet row of the offending cell.
 * @param {*} value The offending cell value.
 * @return {string} The toast message.
 */
function buildInvalidHoursMessage(sheetRow, value) {
  return "Column O of row " + sheetRow + " contains '" + value +
    "', which is not a number. Fix it and press Done again.";
}

/**
 * Builds the confirmation toast shown once the rows were moved.
 *
 * @param {Array<?number>} perRowHours The hours recorded per moved row.
 * @return {string} The toast message.
 */
function buildMoveSuccessMessage(perRowHours) {
  var total = 0;
  var counted = 0;
  for (var i = 0; i < perRowHours.length; i++) {
    if (perRowHours[i] !== null) {
      total += perRowHours[i];
      counted++;
    }
  }
  var tasks = perRowHours.length === 1 ? "1 task" : perRowHours.length + " tasks";
  if (counted === 0) {
    return "Moved " + tasks + " to Done; no hours recorded.";
  }
  return "Moved " + tasks + " to Done; " + (Math.round(total * 100) / 100) + "h recorded.";
}

/**
 * Fingerprints the rows about to move by their Column A text (the task
 * detail), so a pending confirmation is dropped if rows shifted in between.
 *
 * @param {Array<Array>} colAValues Column A values of the rows.
 * @return {string} A stable fingerprint.
 */
function fingerprintRows(colAValues) {
  var texts = [];
  for (var i = 0; i < colAValues.length; i++) {
    texts.push(String(colAValues[i][0]));
  }
  return JSON.stringify(texts);
}

/**
 * Serializes the pending move for CacheService.
 *
 * @param {string} sheetName Sheet the staged rows are on.
 * @param {number} startRow First staged row.
 * @param {number} numRows Number of staged rows.
 * @param {Array<Array>} colAValues Column A values of the staged rows.
 * @return {string} The marker JSON.
 */
function buildMoveMarker(sheetName, startRow, numRows, colAValues) {
  return JSON.stringify({
    sheetName: sheetName,
    startRow: startRow,
    numRows: numRows,
    fingerprint: fingerprintRows(colAValues)
  });
}

/**
 * Parses a stored marker, rejecting anything malformed.
 *
 * @param {?string} markerJson Value read from CacheService, possibly null.
 * @return {?Object} The marker, or null when absent or invalid.
 */
function parseMoveMarker(markerJson) {
  if (!markerJson) {
    return null;
  }
  var marker;
  try {
    marker = JSON.parse(markerJson);
  } catch (e) {
    return null;
  }
  if (!marker || typeof marker.sheetName !== "string" ||
      typeof marker.fingerprint !== "string" ||
      !(marker.startRow >= 1) || !(marker.numRows >= 1)) {
    return null;
  }
  return marker;
}

/**
 * Whether the current button press confirms the pending move. The active row
 * may be one row BELOW the staged rows because pressing Enter after typing
 * in Column O advances the cursor; the fingerprint guarantees the staged
 * rows themselves did not shift.
 *
 * @param {Object} marker Marker from parseMoveMarker().
 * @param {string} sheetName The active sheet's name.
 * @param {number} activeRow The active range's first row.
 * @param {Array<Array>} colAValues Current Column A values at the marker rows.
 * @return {boolean} True when the press should complete the pending move.
 */
function markerMatches(marker, sheetName, activeRow, colAValues) {
  return marker.sheetName === sheetName &&
    activeRow >= marker.startRow &&
    activeRow <= marker.startRow + marker.numRows &&
    fingerprintRows(colAValues) === marker.fingerprint;
}

/**
 * Moves the selected row(s) to the Done sheet using a two-press flow with
 * the sheet itself as the hours input — no dialogs, no HtmlService and no
 * google.script.run (ui.prompt cannot prefill its text box, and HTML
 * dialogs break with PERMISSION_DENIED when several Google accounts are
 * signed in, so the suggestion is prefilled into Column O natively):
 *
 * 1st press: writes each selected row's Column E estimate into its blank
 * Column O cell, selects those cells and asks for confirmation via toast.
 * Existing Column O values are never overwritten.
 *
 * 2nd press: moves the rows using whatever Column O now holds. If every
 * selected row already had numeric hours in Column O, the first press moves
 * immediately.
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

  // A pending confirmation for these rows? Complete it — even if typing in
  // Column O moved the cursor one row down.
  var cache = CacheService.getUserCache();
  var marker = parseMoveMarker(cache.get(MOVE_MARKER_KEY));
  if (marker && marker.sheetName === sourceSheet.getName() &&
      marker.startRow + marker.numRows - 1 <= sourceSheet.getLastRow()) {
    var markerColA = sourceSheet.getRange(marker.startRow, 1, marker.numRows, 1).getValues();
    if (markerMatches(marker, sourceSheet.getName(), startRow, markerColA)) {
      confirmPendingMove(ss, sourceSheet, targetSheet, marker, cache);
      return;
    }
  }
  if (marker) {
    cache.remove(MOVE_MARKER_KEY);
  }

  var eValues = sourceSheet.getRange(startRow, ESTIMATE_COLUMN, numRows, 1).getValues();
  var oRange = sourceSheet.getRange(startRow, HOURS_COLUMN, numRows, 1);
  var oValues = oRange.getValues();
  var plan = planStagedMove(eValues, oValues);

  if (plan.action === "invalid") {
    ss.toast(buildInvalidHoursMessage(startRow + plan.row, plan.value), "Tasks NOT moved", 8);
    return;
  }

  if (plan.action === "move") {
    // Every row already has numeric hours in Column O — no confirmation needed
    moveRowsToDone(ss, sourceSheet, targetSheet, startRow, numRows, plan.perRowHours);
    return;
  }

  // Stage: suggest hours in the blank Column O cells and ask to confirm
  for (var i = 0; i < plan.stagedHours.length; i++) {
    if (String(oValues[i][0]).trim() === "" && plan.stagedHours[i] !== null) {
      oRange.getCell(i + 1, 1).setValue(plan.stagedHours[i]);
    }
  }
  oRange.activate();
  var colAValues = sourceSheet.getRange(startRow, 1, numRows, 1).getValues();
  cache.put(MOVE_MARKER_KEY,
    buildMoveMarker(sourceSheet.getName(), startRow, numRows, colAValues),
    MOVE_MARKER_TTL_SECONDS);
  ss.toast(buildStageToast(plan.stagedCount), "Confirm hours", 8);
}

/**
 * Completes a pending two-press move: records whatever Column O now holds.
 * Blank cells are accepted as "no hours" — the user saw them highlighted
 * and pressed again — but a non-numeric cell still refuses the move (and
 * keeps the confirmation pending so fixing it and pressing again works).
 */
function confirmPendingMove(ss, sourceSheet, targetSheet, marker, cache) {
  var oValues = sourceSheet.getRange(marker.startRow, HOURS_COLUMN, marker.numRows, 1).getValues();
  var invalid = findInvalidHours(oValues);
  if (invalid) {
    ss.toast(buildInvalidHoursMessage(marker.startRow + invalid.row, invalid.value), "Tasks NOT moved", 8);
    return;
  }
  cache.remove(MOVE_MARKER_KEY);
  moveRowsToDone(ss, sourceSheet, targetSheet, marker.startRow, marker.numRows,
    extractHoursEstimates(oValues));
}

/**
 * Copies rows into targetSheet directly under its header block and strips
 * the bold formatting inherited from the row above. Shared by the Done and
 * Projects moves; the caller deletes the source rows.
 *
 * @param {number} headerRows Rows to insert after (1 for Done, 4 for Projects).
 */
function copyRowsBelowHeader(sourceSheet, targetSheet, startRow, numRows, headerRows) {
  var numCols = sourceSheet.getLastColumn();
  var dataToMove = sourceSheet.getRange(startRow, 1, numRows, numCols).getValues();
  targetSheet.insertRowsAfter(headerRows, numRows);
  targetSheet.getRange(headerRows + 1, 1, numRows, numCols).setValues(dataToMove);
  targetSheet.getRange(headerRows + 1, 1, numRows, targetSheet.getMaxColumns()).setFontWeight("normal");
}

/**
 * Moves the rows to the Done sheet, writing the completion timestamp to
 * Column N and the given hours to Column O, then deletes the originals and
 * confirms with a toast.
 *
 * @param {Array<?number>} perRowHours Hours per row; null leaves the cell empty.
 */
function moveRowsToDone(ss, sourceSheet, targetSheet, startRow, numRows, perRowHours) {
  copyRowsBelowHeader(sourceSheet, targetSheet, startRow, numRows, 1);

  // Add the current date and time to Column N
  var timestamp = new Date();
  targetSheet.getRange(2, TIMESTAMP_COLUMN, numRows, 1).setValue(timestamp);

  // Add the hours to Column O
  var hoursColumn = [];
  for (var i = 0; i < perRowHours.length; i++) {
    hoursColumn.push([perRowHours[i] === null ? "" : perRowHours[i]]);
  }
  targetSheet.getRange(2, HOURS_COLUMN, numRows, 1).setValues(hoursColumn);

  // Delete the original rows from the source sheet
  sourceSheet.deleteRows(startRow, numRows);

  ss.toast(buildMoveSuccessMessage(perRowHours), "Done", 5);
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

  if (sourceSheet.getLastColumn() === 0) return;

  copyRowsBelowHeader(sourceSheet, targetSheet, startRow, numRows, 4);

  // Delete the original rows from the source sheet
  sourceSheet.deleteRows(startRow, numRows);
}

// Allow the pure helpers to be unit-tested under Node; Apps Script has no
// module object, so this block is inert when deployed.
if (typeof module !== "undefined" && module.exports) {
  module.exports = {
    parseHoursNumber: parseHoursNumber,
    extractHoursEstimates: extractHoursEstimates,
    findInvalidHours: findInvalidHours,
    planStagedMove: planStagedMove,
    buildStageToast: buildStageToast,
    buildInvalidHoursMessage: buildInvalidHoursMessage,
    buildMoveSuccessMessage: buildMoveSuccessMessage,
    fingerprintRows: fingerprintRows,
    buildMoveMarker: buildMoveMarker,
    parseMoveMarker: parseMoveMarker,
    markerMatches: markerMatches
  };
}

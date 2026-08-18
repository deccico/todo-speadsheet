const test = require("node:test");
const assert = require("node:assert/strict");

const {
  parseHoursNumber,
  extractHoursEstimates,
  findInvalidHours,
  planStagedMove,
  buildStageToast,
  buildInvalidHoursMessage,
  buildMoveSuccessMessage,
  fingerprintRows,
  buildMoveMarker,
  parseMoveMarker,
  markerMatches,
} = require("../Done.js");

test("parseHoursNumber accepts numbers and numeric text", () => {
  assert.equal(parseHoursNumber(4), 4);
  assert.equal(parseHoursNumber("1.5"), 1.5);
  assert.equal(parseHoursNumber(" 4.25 "), 4.25);
  assert.equal(parseHoursNumber("0"), 0);
});

test("parseHoursNumber is strict: trailing garbage is rejected", () => {
  assert.equal(parseHoursNumber("3 hours"), null);
  assert.equal(parseHoursNumber("7h"), null);
  assert.equal(parseHoursNumber("1,5"), null);
  assert.equal(parseHoursNumber("abc"), null);
});

test("parseHoursNumber rejects blanks and non-numeric cell types", () => {
  assert.equal(parseHoursNumber(""), null);
  assert.equal(parseHoursNumber("   "), null);
  assert.equal(parseHoursNumber(new Date(2026, 0, 1)), null);
  assert.equal(parseHoursNumber(true), null);
});

test("extractHoursEstimates maps each cell to a number or null", () => {
  assert.deepEqual(extractHoursEstimates([[2], [""], ["1.5"], ["abc"]]), [2, null, 1.5, null]);
});

test("findInvalidHours accepts numeric and blank cells", () => {
  assert.equal(findInvalidHours([[2], [""], ["1.5"], ["  "]]), null);
});

test("findInvalidHours reports the first non-numeric, non-blank cell", () => {
  assert.deepEqual(findInvalidHours([[2], ["4h"], ["abc"]]), { row: 1, value: "4h" });
});

test("planStagedMove moves when every row has numeric hours in Column O", () => {
  assert.deepEqual(planStagedMove([[1], [2]], [[3], ["4.5"]]),
    { action: "move", perRowHours: [3, 4.5] });
});

test("planStagedMove stages estimates into blank Column O cells", () => {
  assert.deepEqual(planStagedMove([[1], [2.5]], [[""], [""]]),
    { action: "stage", stagedHours: [1, 2.5], stagedCount: 2 });
});

test("planStagedMove keeps existing Column O values when staging", () => {
  assert.deepEqual(planStagedMove([[1], [2]], [[4], [""]]),
    { action: "stage", stagedHours: [4, 2], stagedCount: 1 });
});

test("planStagedMove leaves rows without an estimate as null", () => {
  assert.deepEqual(planStagedMove([[""], ["abc"]], [[""], [""]]),
    { action: "stage", stagedHours: [null, null], stagedCount: 0 });
});

test("planStagedMove treats whitespace-only Column O cells as blank", () => {
  assert.deepEqual(planStagedMove([[2]], [["   "]]),
    { action: "stage", stagedHours: [2], stagedCount: 1 });
});

test("planStagedMove refuses non-numeric Column O content", () => {
  assert.deepEqual(planStagedMove([[1], [2]], [["4h"], [""]]),
    { action: "invalid", row: 0, value: "4h" });
});

test("planStagedMove accepts an explicit zero in Column O", () => {
  assert.deepEqual(planStagedMove([[5]], [[0]]),
    { action: "move", perRowHours: [0] });
});

test("buildStageToast asks for confirmation when suggestions were written", () => {
  assert.match(buildStageToast(2), /adjust them if needed/);
  assert.match(buildStageToast(2), /press Done again/);
});

test("buildStageToast asks for manual input when nothing could be suggested", () => {
  assert.match(buildStageToast(0), /Enter the hours/);
  assert.match(buildStageToast(0), /press Done again/);
});

test("buildInvalidHoursMessage names the row and the offending value", () => {
  const message = buildInvalidHoursMessage(7, "4h");
  assert.match(message, /row 7/);
  assert.match(message, /'4h'/);
  assert.match(message, /not a number/);
});

test("buildMoveSuccessMessage totals the hours of the moved rows", () => {
  assert.equal(buildMoveSuccessMessage([3.5]), "Moved 1 task to Done; 3.5h recorded.");
  assert.equal(buildMoveSuccessMessage([1, 2.5]), "Moved 2 tasks to Done; 3.5h recorded.");
});

test("buildMoveSuccessMessage rounds away floating-point noise", () => {
  assert.equal(buildMoveSuccessMessage([0.1, 0.2]), "Moved 2 tasks to Done; 0.3h recorded.");
});

test("buildMoveSuccessMessage skips blank rows and says so when all are blank", () => {
  assert.equal(buildMoveSuccessMessage([2, null]), "Moved 2 tasks to Done; 2h recorded.");
  assert.equal(buildMoveSuccessMessage([null]), "Moved 1 task to Done; no hours recorded.");
});

test("fingerprintRows depends on the rows' Column A text", () => {
  assert.equal(fingerprintRows([["Fix login"], ["Ship v2"]]),
    fingerprintRows([["Fix login"], ["Ship v2"]]));
  assert.notEqual(fingerprintRows([["Fix login"]]), fingerprintRows([["Ship v2"]]));
});

test("buildMoveMarker survives a parse round-trip", () => {
  const marker = parseMoveMarker(buildMoveMarker("Tasks", 5, 2, [["a"], ["b"]]));
  assert.equal(marker.sheetName, "Tasks");
  assert.equal(marker.startRow, 5);
  assert.equal(marker.numRows, 2);
  assert.equal(marker.fingerprint, fingerprintRows([["a"], ["b"]]));
});

test("parseMoveMarker rejects null, garbage and malformed markers", () => {
  assert.equal(parseMoveMarker(null), null);
  assert.equal(parseMoveMarker(""), null);
  assert.equal(parseMoveMarker("not json"), null);
  assert.equal(parseMoveMarker("{}"), null);
  assert.equal(parseMoveMarker(JSON.stringify({ sheetName: "T", startRow: 0, numRows: 1, fingerprint: "[]" })), null);
});

test("markerMatches accepts the staged rows and one row below (Enter moved the cursor)", () => {
  const marker = parseMoveMarker(buildMoveMarker("Tasks", 5, 2, [["a"], ["b"]]));
  const colA = [["a"], ["b"]];
  assert.equal(markerMatches(marker, "Tasks", 5, colA), true);
  assert.equal(markerMatches(marker, "Tasks", 6, colA), true);
  assert.equal(markerMatches(marker, "Tasks", 7, colA), true);
});

test("markerMatches rejects rows outside the staged block's tolerance", () => {
  const marker = parseMoveMarker(buildMoveMarker("Tasks", 5, 2, [["a"], ["b"]]));
  const colA = [["a"], ["b"]];
  assert.equal(markerMatches(marker, "Tasks", 4, colA), false);
  assert.equal(markerMatches(marker, "Tasks", 8, colA), false);
});

test("markerMatches rejects another sheet or shifted rows", () => {
  const marker = parseMoveMarker(buildMoveMarker("Tasks", 5, 2, [["a"], ["b"]]));
  assert.equal(markerMatches(marker, "Other", 5, [["a"], ["b"]]), false);
  assert.equal(markerMatches(marker, "Tasks", 5, [["b"], ["c"]]), false);
});

const test = require("node:test");
const assert = require("node:assert/strict");

const {
  parseHoursNumber,
  extractHoursEstimates,
  computeDefaultHours,
  buildHoursPromptMessage,
  escapeHtml,
  buildHoursDialogHtml,
  resolveHoursPlan,
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

test("extractHoursEstimates maps each row's Column E cell to a number or null", () => {
  assert.deepEqual(extractHoursEstimates([[2], [""], ["1.5"], ["abc"]]), [2, null, 1.5, null]);
});

test("computeDefaultHours sums the numeric estimates of the selection", () => {
  assert.equal(computeDefaultHours([[2], [3.5]]), 5.5);
});

test("computeDefaultHours works for a single-row selection", () => {
  assert.equal(computeDefaultHours([[4]]), 4);
});

test("computeDefaultHours ignores blank and non-numeric cells", () => {
  assert.equal(computeDefaultHours([[2], [""], ["abc"]]), 2);
});

test("computeDefaultHours returns null when no cell is numeric", () => {
  assert.equal(computeDefaultHours([[""], ["estimate?"]]), null);
});

test("computeDefaultHours rounds away floating-point noise", () => {
  assert.equal(computeDefaultHours([[0.1], [0.2]]), 0.3);
});

test("buildHoursPromptMessage explains the suggestion when one exists", () => {
  const message = buildHoursPromptMessage(5.5);
  assert.match(message, /suggested value/);
  assert.match(message, /Column E/);
});

test("buildHoursPromptMessage omits the suggestion talk when there is none", () => {
  const message = buildHoursPromptMessage(null);
  assert.doesNotMatch(message, /suggest/i);
  assert.doesNotMatch(message, /Column E/);
});

test("escapeHtml escapes HTML metacharacters", () => {
  assert.equal(
    escapeHtml('<b>"a" & \'b\'</b>'),
    "&lt;b&gt;&quot;a&quot; &amp; &#39;b&#39;&lt;/b&gt;"
  );
});

const DIALOG_CONTEXT = { sheetName: "Tasks", startRow: 3, numRows: 2 };

test("buildHoursDialogHtml prefills the input with the suggested total", () => {
  const html = buildHoursDialogHtml(5.5, DIALOG_CONTEXT);
  assert.match(html, /value="5\.5"/);
});

test("buildHoursDialogHtml pre-selects the suggestion so typing replaces it", () => {
  const html = buildHoursDialogHtml(5.5, DIALOG_CONTEXT);
  assert.match(html, /input\.select\(\)/);
});

test("buildHoursDialogHtml leaves the input empty when there is no suggestion", () => {
  const html = buildHoursDialogHtml(null, DIALOG_CONTEXT);
  assert.match(html, /value=""/);
});

test("buildHoursDialogHtml embeds the label and the selection snapshot", () => {
  const html = buildHoursDialogHtml(4, DIALOG_CONTEXT);
  assert.match(html, /Column E/);
  assert.ok(html.includes('{"sheetName":"Tasks","startRow":3,"numRows":2}'));
  assert.match(html, /completeTaskMove\(/);
});

test("buildHoursDialogHtml neutralizes a hostile sheet name", () => {
  const hostile = '</script><script>alert(1)</script>';
  const html = buildHoursDialogHtml(4, { sheetName: hostile, startRow: 1, numRows: 1 });
  assert.ok(!html.includes(hostile));
  assert.ok(html.includes("\\u003c/script>"));
});

test("resolveHoursPlan: blank input accepts each row's own estimate", () => {
  assert.deepEqual(resolveHoursPlan("", [[1], [2.5]]), { perRowHours: [1, 2.5] });
});

test("resolveHoursPlan: rows without an estimate stay null in the plan", () => {
  assert.deepEqual(resolveHoursPlan("", [[1], [""]]), { perRowHours: [1, null] });
});

test("resolveHoursPlan: whitespace-only input counts as blank", () => {
  assert.deepEqual(resolveHoursPlan("   ", [[2]]), { perRowHours: [2] });
});

test("resolveHoursPlan: blank input is invalid when no estimate exists", () => {
  assert.equal(resolveHoursPlan("", [[""], ["abc"]]), null);
});

test("resolveHoursPlan: a typed number overrides the default for every row", () => {
  assert.deepEqual(resolveHoursPlan("7", [[1], [2]]), { sameHours: 7 });
});

test("resolveHoursPlan: submitting the suggested total unchanged keeps per-row estimates", () => {
  assert.deepEqual(resolveHoursPlan("3", [[1], [2]]), { perRowHours: [1, 2] });
});

test("resolveHoursPlan: the suggested total matches numerically, not textually", () => {
  assert.deepEqual(resolveHoursPlan("3.0", [[1], [2]]), { perRowHours: [1, 2] });
  assert.deepEqual(resolveHoursPlan(" 3 ", [[1], [2]]), { perRowHours: [1, 2] });
});

test("resolveHoursPlan: accepting the suggestion leaves estimate-less rows empty", () => {
  assert.deepEqual(resolveHoursPlan("2", [[2], [""]]), { perRowHours: [2, null] });
});

test("resolveHoursPlan: typed decimals and surrounding whitespace are accepted", () => {
  assert.deepEqual(resolveHoursPlan(" 4.25 ", [[""]]), { sameHours: 4.25 });
});

test("resolveHoursPlan: an explicit zero is accepted", () => {
  assert.deepEqual(resolveHoursPlan("0", [[5]]), { sameHours: 0 });
});

test("resolveHoursPlan: non-numeric input is invalid even with a default", () => {
  assert.equal(resolveHoursPlan("abc", [[5]]), null);
  assert.equal(resolveHoursPlan("3 hours", [[5]]), null);
  assert.equal(resolveHoursPlan("1,5", [[5]]), null);
});

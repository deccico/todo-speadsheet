const test = require("node:test");
const assert = require("node:assert/strict");

const {
  parseHoursNumber,
  extractHoursEstimates,
  computeDefaultHours,
  buildHoursPromptMessage,
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

test("buildHoursPromptMessage mentions the default when one exists", () => {
  assert.match(buildHoursPromptMessage(5.5), /leave blank/);
  assert.match(buildHoursPromptMessage(5.5), /total: 5\.5/);
});

test("buildHoursPromptMessage omits the default when there is none", () => {
  const message = buildHoursPromptMessage(null);
  assert.doesNotMatch(message, /blank/);
  assert.doesNotMatch(message, /default/);
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

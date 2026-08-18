# Todo Spreadsheet Scripts

Google Apps Script helpers for managing a todo list in Google Sheets. Selected
task rows can be moved to a **Done** sheet (recording a completion timestamp
and the hours spent) or to a **Todo Projects** sheet.

## Functions

### `taskIsDone()`

Moves the currently selected row(s) from the active sheet to the **Done**
sheet:

1. Opens a small dialog asking for the hours spent on the completed task(s).
   The input comes **prefilled with a suggestion** taken from **Column E**
   (the estimate column) — the total of the selection's numeric estimates —
   and the text is pre-selected, so pressing **OK**/**Enter** accepts it and
   typing a number replaces it. Accepting the suggestion unchanged (or
   clearing the input) gives each moved row its own Column E estimate (rows
   without a numeric estimate are left empty); any other number is written to
   every moved row.
2. Inserts the rows at the top of the **Done** sheet (starting at row 2) and
   clears any bold formatting inherited from the header row.
3. Writes the completion date/time to **Column N** and the hours to
   **Column O** of each moved row.
4. Deletes the original rows from the source sheet.

If Column E contains no numeric value anywhere in the selection, the input
starts empty and an explicit numeric input is required. Input is parsed
strictly: `3 hours` or `1,5` is rejected rather than silently recorded as `3`
or `1` — the dialog shows the error and stays open for a correction.

### `moveToProjectsSheet()`

Moves the currently selected row(s) from the active sheet to the
**Todo Projects** sheet, inserting them at row 5 (below that sheet's header
block) and clearing bold formatting. The original rows are deleted from the
source sheet.

## Spreadsheet layout assumptions

| Sheet | Purpose |
|---|---|
| (any task sheet) | Source of rows; Column E holds the hours estimate |
| `Done` | Completed tasks; Column N = completion timestamp, Column O = hours spent |
| `Todo Projects` | Project backlog; rows are inserted after row 4 |

## Installation

1. Open the spreadsheet in Google Sheets.
2. Go to **Extensions → Apps Script**.
3. Paste the contents of `Done.js` into the script editor and save.
4. Run the functions from the editor, bind them to buttons/drawings, or add
   them to a custom menu via an `onOpen` trigger.

Alternatively, manage the project from this repository with
[clasp](https://github.com/google/clasp).

## Development

Apps Script's `ui.prompt` cannot prefill its text box, so the hours prompt is
an `HtmlService` modal built from an inline string (`buildHoursDialogHtml`) —
no separate HTML file, keeping the project a single pasteable file. Because
`showModalDialog` does not block, `taskIsDone()` only opens the dialog; the
dialog submits the input back to `completeTaskMove(hoursInput, context)` via
`google.script.run`, which validates it and performs the move (invalid input
is shown inside the dialog, which stays open).

The prompt/default logic is factored into pure functions (`parseHoursNumber`,
`extractHoursEstimates`, `computeDefaultHours`, `buildHoursPromptMessage`,
`escapeHtml`, `buildHoursDialogHtml`, `resolveHoursPlan`) so it can be
unit-tested without the Apps Script runtime. With Node.js 18+ installed:

```bash
npm test
```

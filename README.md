# Todo Spreadsheet Scripts

Google Apps Script helpers for managing a todo list in Google Sheets. Selected
task rows can be moved to a **Done** sheet (recording a completion timestamp
and the hours spent) or to a **Todo Projects** sheet.

## Functions

### `taskIsDone()`

Moves the currently selected row(s) from the active sheet to the **Done**
sheet:

1. Prompts for the hours spent on the completed task(s), offering a default
   taken from **Column E** (the estimate column). Leave the input blank and
   press **OK** to accept it: each moved row then receives its own Column E
   estimate (rows without a numeric estimate are left empty), and the prompt
   shows the total of those estimates for reference. Type a number instead to
   override the default — that number is written to every moved row.
   (Apps Script's `ui.prompt` cannot prefill the text box, which is why the
   default is accepted via blank input.)
2. Inserts the rows at the top of the **Done** sheet (starting at row 2) and
   clears any bold formatting inherited from the header row.
3. Writes the completion date/time to **Column N** and the hours to
   **Column O** of each moved row.
4. Deletes the original rows from the source sheet.

If Column E contains no numeric value anywhere in the selection, no default is
offered and an explicit numeric input is required. Input is parsed strictly:
`3 hours` or `1,5` is rejected rather than silently recorded as `3` or `1`.

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

The prompt/default logic is factored into pure functions (`parseHoursNumber`,
`extractHoursEstimates`, `computeDefaultHours`, `buildHoursPromptMessage`,
`resolveHoursPlan`) so it can be unit-tested without the Apps Script runtime.
With Node.js 18+ installed:

```bash
npm test
```

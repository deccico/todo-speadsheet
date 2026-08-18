# Todo Spreadsheet Scripts

Google Apps Script helpers for managing a todo list in Google Sheets. Selected
task rows can be moved to a **Done** sheet (recording a completion timestamp
and the hours spent) or to a **Todo Projects** sheet.

## Functions

### `taskIsDone()`

Moves the currently selected row(s) from the active sheet to the **Done**
sheet using a two-press flow with no dialogs — the sheet itself is the hours
input:

1. **First press** — the suggestion is prefilled *in the sheet*: each
   selected row's **Column E** estimate is written into its **Column O**
   cell (only blank cells are filled; existing values and every other
   column are never touched), the cells are selected, and a toast asks you
   to confirm. Adjust any value directly in the cell — per-row overrides
   are possible.
2. **Second press** — the rows are moved: inserted at the top of the
   **Done** sheet (row 2) with bold formatting cleared, completion
   date/time written to **Column N**, and the Column O hours kept as the
   recorded hours. The original rows are deleted and a toast confirms the
   total.

Details of the flow:

- If you type the actual hours into Column O yourself before pressing, the
  rows move on the **first** press — no confirmation needed.
- The pending confirmation is remembered for 5 minutes, so it still
  completes if pressing Enter after typing in Column O moved the cursor one
  row down. A fingerprint of the tasks' Column A text guarantees rows that
  shifted in the meantime are never moved by mistake.
- A non-numeric Column O value (`4h`, `1,5`) refuses the move and points at
  the offending row; parsing is strict so a typo cannot silently record the
  wrong hours. Rows without an estimate are left blank for you to fill (or
  confirm as blank by pressing again).

### `moveToProjectsSheet()`

Moves the currently selected row(s) from the active sheet to the
**Todo Projects** sheet, inserting them at row 5 (below that sheet's header
block) and clearing bold formatting. The original rows are deleted from the
source sheet.

## Spreadsheet layout assumptions

| Sheet | Purpose |
|---|---|
| (any task sheet) | Source of rows; Column A holds the task detail, Column E the hours estimate, Column O the hours input for the move (columns M and after are otherwise empty) |
| `Done` | Completed tasks; Column N = completion timestamp, Column O = hours spent |
| `Todo Projects` | Project backlog; rows are inserted after row 4 |

## Design note: why the sheet is the input (no dialog)

The obvious design — a prompt with the suggestion prefilled in its text
box — is not reliably buildable in Apps Script:

- The native `ui.prompt` cannot prefill its text box at all.
- An `HtmlService` dialog can, but returning the value requires
  `google.script.run`, which suffers a long-standing Google bug: with more
  than one Google account signed into the browser (any browser), the call
  can run under the wrong account and fail with *"a server error occurred
  while reading from storage. Error code PERMISSION_DENIED"*. Such dialogs
  also take about a second to open (sandboxed iframe).

So the suggestion is prefilled into Column O natively, edited as a normal
cell, and read back on the second press. This uses no HTML, no iframes and
no `google.script.run`, making it immune to the multi-account bug and as
fast as Sheets itself.

## Installation

1. Open the spreadsheet in Google Sheets.
2. Go to **Extensions → Apps Script**.
3. Paste the contents of `Done.js` into the script editor and save.
4. Run the functions from the editor, bind them to buttons/drawings, or add
   them to a custom menu via an `onOpen` trigger.

Alternatively, manage the project from this repository with
[clasp](https://github.com/google/clasp).

## Development

The decision logic is factored into pure functions (`parseHoursNumber`,
`extractHoursEstimates`, `findInvalidHours`, `planStagedMove`,
`buildStageToast`, `buildInvalidHoursMessage`, `buildMoveSuccessMessage`,
`fingerprintRows`, `buildMoveMarker`, `parseMoveMarker`, `markerMatches`)
so it can be unit-tested without the Apps Script runtime. The pending
confirmation is stored per user in `CacheService` (no extra OAuth scope
needed). With Node.js 18+ installed:

```bash
npm test
```

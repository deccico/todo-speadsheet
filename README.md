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
   every moved row. OK validates the input inside the dialog and closes it
   immediately; the move then runs in the background and a toast
   (bottom-right) reports the result.
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

## Troubleshooting

**"We're sorry, a server error occurred while reading from storage. Error
code PERMISSION_DENIED"** — a long-standing Google bug, not a bug in this
script: when the browser profile is signed into more than one Google
account, `google.script.run` calls from an HTML dialog can run under the
wrong account and fail with this error. It affects every browser (the
sessions share one cookie jar) and there is no code-side fix. Ways around
it, in order of convenience:

- **Firefox**: install Mozilla's *Multi-Account Containers* extension and
  open the spreadsheet in a container where only the owning account is
  signed in — other accounts keep working in normal tabs.
- **Any browser**: make the owning account the *default* account — sign out
  of all Google accounts, sign in with the owning account first, then the
  others (the dialog binds to the first session).
- **Any browser**: use a private window, or a separate browser profile,
  with only the owning account signed in.

If the dialog still fails with only one account signed in, Firefox's
Enhanced Tracking Protection / Total Cookie Protection may be blocking the
dialog iframe's cookies: click the shield icon on the docs.google.com tab
and turn protection off for that site.

Its symptom with this script: the dialog closes on OK but the rows stay put
and no toast appears.

The dialog also takes a moment to open: `HtmlService` dialogs load inside a
sandboxed iframe, which costs roughly a second. That is inherent to Apps
Script (the native `ui.prompt` is instant but cannot prefill its text box),
and is why the dialog validates locally and closes immediately on OK rather
than staying open while the move runs.

## Development

Apps Script's `ui.prompt` cannot prefill its text box, so the hours prompt is
an `HtmlService` modal built from an inline string (`buildHoursDialogHtml`) —
no separate HTML file, keeping the project a single pasteable file. Because
`showModalDialog` does not block, `taskIsDone()` only opens the dialog. The
dialog validates the input locally (the source of `parseHoursNumber` is
injected into its script, so client and server apply the same rule), then
fires `completeTaskMove(hoursInput, context)` via `google.script.run` and
closes immediately; the move runs in the background and reports success or
failure through spreadsheet toasts.

The prompt/default logic is factored into pure functions (`parseHoursNumber`,
`extractHoursEstimates`, `computeDefaultHours`, `buildHoursPromptMessage`,
`escapeHtml`, `buildHoursDialogHtml`, `buildMoveSuccessMessage`,
`resolveHoursPlan`) so it can be unit-tested without the Apps Script
runtime. With Node.js 18+ installed:

```bash
npm test
```

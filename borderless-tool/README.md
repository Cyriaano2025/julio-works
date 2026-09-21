# borderless-tool

Writes the weekly Borderless report into the **JULIO** tab of the tracking
spreadsheet, in place, over the Google Sheets API. The file is never downloaded
or re-uploaded — only the listed cells are touched.

## One-time setup

```bash
pip install -r requirements.txt
```

Then put an OAuth **Desktop** client `credentials.json` next to the script. If
it is missing, the script prints step-by-step Google Cloud Console instructions
and stops. The first run opens a browser once; the login is cached in
`token.json`.

`credentials.json` and `token.json` are gitignored — never commit them.

## Weekly use

Either type `/borderless` in Claude Code followed by the JSON, or run it
directly:

```bash
python3 borderless_update.py input.json --dry-run   # preview
python3 borderless_update.py input.json             # write
```

### Input

```json
{"month": "SEPTEMBER", "label": "Week 3", "cells": {"25": "text", "26": "text"}}
```

* `label` — `Week 1`–`Week 4`, `Monthly Report`, `Variance Analysis` or `RAG`.
* `cells` — row number to text. Rows not listed are left untouched.

### Flags

| Flag | Effect |
| --- | --- |
| `--dry-run` | Print the plan, write nothing |
| `--overwrite` | Allow replacing cells that already have content |
| `--check-rows` | Re-run the row layout word check on a later run |
| `--refresh-labels` | Rewrite `row_labels.json` from the sheet as it is now |

## What it does

1. Reads `JULIO!A1:ZZ60`. Row 10 holds the month (merged, so it is carried
   forward across columns), row 11 holds the label. Both must match,
   case-insensitively, in exactly one column — otherwise it stops.
2. Builds a row-to-label map from columns B and C. The first run saves it as
   `row_labels.json`; later runs stop if a label on a row being written has
   changed.
3. On the first run, checks the expected wording of rows 25, 26, 28, 31, 32,
   33, 34, 38, 41, 42, 43 and 45.
4. Stops and shows the contents if any target cell is already filled, unless
   `--overwrite` is passed.
5. Writes only the listed cells with `values.batchUpdate`, `valueInputOption`
   `RAW`.
6. Applies text wrap to just those cells (`batchUpdate` / `repeatCell` /
   `wrapStrategy: WRAP`).
7. Reads the cells back and prints a cell / row label / value table.

Any check that fails exits non-zero having written nothing.

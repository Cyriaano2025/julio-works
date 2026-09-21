---
description: Write a weekly Borderless report into the JULIO Google Sheet
argument-hint: '{"month": "SEPTEMBER", "label": "Week 3", "cells": {"25": "..."}}'
allowed-tools: Bash(python3 borderless_update.py:*), Bash(cd /home/user/julio-works/borderless-tool && python3 borderless_update.py:*), Write
---

The user's report payload is everything after `/borderless`:

$ARGUMENTS

Do exactly this, in order, and nothing else:

1. Save the payload verbatim to `/home/user/julio-works/borderless-tool/input.json`.
   Do not reword, summarise, expand, shorten or "fix" any of the text — write the
   JSON through exactly as given. If it is not valid JSON, say so and stop.

2. Run:

   ```
   cd /home/user/julio-works/borderless-tool && python3 borderless_update.py input.json
   ```

3. If the script exits 0, show the user the read-back table it printed
   (cell, row label, value) and nothing else.

4. If the script stops for any reason — non-zero exit, a `STOPPED:` message,
   missing `credentials.json`, no matching column, changed row labels, a failed
   row layout check, or cells that already have content — show the user the
   script's own message explaining why, and **do nothing else**. Do not retry,
   do not pass `--overwrite`, do not edit the sheet by another route, and do not
   modify the payload to get past the check. Wait for the user to decide.

Notes:

- Pass `--dry-run` as well if the user asked to preview first.
- Only add `--overwrite` if the user explicitly asks for it in this message.
- The script writes only the listed cells over the Sheets API. Never download or
  re-upload the spreadsheet.

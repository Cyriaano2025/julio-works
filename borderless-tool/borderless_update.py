#!/usr/bin/env python3
"""Write a weekly report into the Borderless tracking sheet.

Reads a small JSON payload describing a month, a column label (e.g. "Week 3")
and a row-number -> text map, locates the matching column in the JULIO tab, and
writes only those cells in place via the Sheets API.

The spreadsheet is never downloaded or re-uploaded; every operation is a
targeted API call against the live file.

Usage:
    python3 borderless_update.py input.json [--dry-run] [--overwrite]
"""

import argparse
import json
import os
import sys

SPREADSHEET_ID = "1GL9pEfrcvVSilVL87z9S2UopU_YBLir1dTavQ6qVPj8"
TAB = "JULIO"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

ACCOUNT = "juliocyriaano@theechohouse.com"

# Where the sheet is laid out.
READ_RANGE = f"{TAB}!A1:ZZ60"
# The month header row and the week-label row directly beneath it are found
# at run time (they sit at 11/12 in the current template, 10/11 in an older
# one), by scanning this many rows from the top.
HEADER_SEARCH_ROWS = 20
LABEL_COLUMNS = (1, 2)  # 0-based: columns B and C hold the row descriptions

VALID_LABELS = [
    "Week 1", "Week 2", "Week 3", "Week 4",
    "Monthly Report", "Variance Analysis", "RAG",
]

HERE = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(HERE, "credentials.json")
TOKEN_FILE = os.path.join(HERE, "token.json")
ROW_LABELS_FILE = os.path.join(HERE, "row_labels.json")

# First-run sanity check: each row's description must contain this word.
EXPECTED_ROW_WORDS = {
    25: "Clients",
    26: "New business",
    28: "OPS",
    31: "Ongoing",
    32: "Top priorities",
    33: "P.O",
    34: "French West Africa",
    38: "addition",
    41: "Special",
    42: "Tidal Rave",
    43: "Partnerships",
    45: "Risk",
}


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

def stop(message, detail=None):
    """Abort loudly. Nothing has been written unless stated otherwise."""
    print(f"\nSTOPPED: {message}", file=sys.stderr)
    if detail:
        print(detail, file=sys.stderr)
    sys.exit(1)


def col_letter(index):
    """0-based column index -> A1 letter (0 -> A, 26 -> AA)."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def cell_at(row, col_index):
    return f"{col_letter(col_index)}{row}"


def get_cell(grid, row, col_index):
    """1-based row, 0-based column, from a ragged values.get grid."""
    if row - 1 >= len(grid):
        return ""
    line = grid[row - 1]
    if col_index >= len(line):
        return ""
    return (line[col_index] or "").strip()


def truncate(text, width=70):
    text = " ".join(str(text).split())
    return text if len(text) <= width else text[: width - 1] + "…"


def print_table(rows, headers):
    widths = [len(h) for h in headers]
    for row in rows:
        for i, value in enumerate(row):
            widths[i] = max(widths[i], len(str(value)))
    rule = "-+-".join("-" * w for w in widths)
    print(" | ".join(h.ljust(widths[i]) for i, h in enumerate(headers)))
    print(rule)
    for row in rows:
        print(" | ".join(str(v).ljust(widths[i]) for i, v in enumerate(row)))


# --------------------------------------------------------------------------
# auth
# --------------------------------------------------------------------------

CREDENTIALS_HOWTO = f"""
credentials.json was not found at:
    {CREDENTIALS_FILE}

Create an OAuth Desktop client and download it, signed in as {ACCOUNT}:

  1. Go to https://console.cloud.google.com/ and sign in as {ACCOUNT}.
  2. Top-left project picker -> "New Project". Name it e.g. "Borderless Tool"
     and click Create, then make sure it is the selected project.
  3. Left menu -> "APIs & Services" -> "Library". Search for
     "Google Sheets API", open it and click "Enable".
  4. Left menu -> "APIs & Services" -> "OAuth consent screen".
       - User type: "External" (or "Internal" if theechohouse.com is a
         Google Workspace org and it is offered), then Create.
       - App name: "Borderless Tool". User support email and Developer
         contact email: {ACCOUNT}. Save and Continue.
       - Scopes: Save and Continue (the scope is requested by the script).
       - Test users: click "Add users", add {ACCOUNT}, Save and Continue.
         (Skip this step if you chose "Internal".)
  5. Left menu -> "APIs & Services" -> "Credentials".
       - "Create Credentials" -> "OAuth client ID".
       - Application type: "Desktop app". Name: "Borderless Desktop".
       - Click Create, then "Download JSON" on the dialog (or the download
         icon next to the client in the list).
  6. Rename the downloaded file to exactly credentials.json and put it at:
         {CREDENTIALS_FILE}

Then re-run this script. A browser window will open once for you to grant
access; the resulting token is cached in token.json so you only log in once.
"""


def load_service():
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError:
        stop(
            "Google API libraries are missing.",
            "Install them with:\n"
            "    pip install -r requirements.txt\n"
            "(google-api-python-client google-auth-httplib2 google-auth-oauthlib)",
        )

    creds = None
    if os.path.exists(TOKEN_FILE):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        except ValueError:
            creds = None

    if creds and creds.valid:
        pass
    elif creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as exc:  # refresh token revoked or expired
            print(f"Cached token could not be refreshed ({exc}); logging in again.")
            creds = None
    else:
        creds = None

    if not creds:
        if not os.path.exists(CREDENTIALS_FILE):
            stop("credentials.json is missing, so I cannot sign in.", CREDENTIALS_HOWTO)
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        print(f"Opening a browser to sign in as {ACCOUNT} ...")
        creds = flow.run_local_server(port=0, prompt="consent")

    with open(TOKEN_FILE, "w") as handle:
        handle.write(creds.to_json())
    os.chmod(TOKEN_FILE, 0o600)

    return build("sheets", "v4", credentials=creds, cache_discovery=False)


# --------------------------------------------------------------------------
# steps
# --------------------------------------------------------------------------

def read_payload(path):
    if not os.path.exists(path):
        stop(f"Input file not found: {path}")
    try:
        with open(path) as handle:
            payload = json.load(handle)
    except json.JSONDecodeError as exc:
        stop(f"Input file is not valid JSON: {path}", str(exc))

    for key in ("month", "label", "cells"):
        if key not in payload:
            stop(f'Input is missing the "{key}" field.')

    month = str(payload["month"]).strip()
    label = str(payload["label"]).strip()

    matched = [v for v in VALID_LABELS if v.lower() == label.lower()]
    if not matched:
        stop(
            f'"{label}" is not a recognised label.',
            "Valid labels: " + ", ".join(VALID_LABELS),
        )
    label = matched[0]

    if not isinstance(payload["cells"], dict) or not payload["cells"]:
        stop('"cells" must be a non-empty object of row number -> text.')

    cells = {}
    for raw_row, text in payload["cells"].items():
        try:
            row = int(str(raw_row).strip())
        except ValueError:
            stop(f'"cells" key "{raw_row}" is not a row number.')
        if row < 1 or row > 60:
            stop(f"Row {row} is outside the sheet area I read (rows 1-60).")
        cells[row] = "" if text is None else str(text)

    return month, label, cells


def find_column(grid, month, label):
    """Locate the column matching month and label.

    The month header sits one row above the week labels, but which pair of
    rows that is has moved between versions of the template, so find it
    rather than trusting a fixed row number.
    """
    width = max((len(line) for line in grid), default=0)

    def months_along(row):
        """Row's values carried forward across merged cells."""
        out, carried = [], ""
        for col in range(width):
            value = get_cell(grid, row, col)
            if value:
                carried = value
            out.append(carried)
        return out

    matches = []   # (month_row, label_row, col)
    seen = []      # every month/label pair found, for diagnostics
    for month_row in range(1, min(len(grid), HEADER_SEARCH_ROWS)):
        label_row = month_row + 1
        months = months_along(month_row)
        for col in range(width):
            col_label = get_cell(grid, label_row, col)
            if not months[col] or not col_label:
                continue
            seen.append((month_row, label_row, col, months[col], col_label))
            if (months[col].lower() == month.lower()
                    and col_label.lower() == label.lower()):
                matches.append((month_row, label_row, col))

    if not matches:
        detail = [f"  row {m}/{l}  {col_letter(c)}: {mon} / {lab}"
                  for m, l, c, mon, lab in seen]
        stop(
            f'No column matches month "{month}" with label "{label}".',
            "Month/label pairs found in the sheet:\n"
            + ("\n".join(detail) or "  (none)"),
        )

    # Several row pairs can look plausible if the template repeats headers.
    rows_hit = {(m, l) for m, l, _ in matches}
    if len(rows_hit) > 1:
        where = ", ".join(f"rows {m}/{l}" for m, l in sorted(rows_hit))
        stop(
            f'"{month}" / "{label}" appears under more than one header row: {where}.',
            "I will not guess which one you meant.",
        )
    if len(matches) > 1:
        found = ", ".join(col_letter(c) for _, _, c in matches)
        stop(
            f'"{month}" / "{label}" matches more than one column: {found}.',
            "I will not guess which one you meant.",
        )
    return matches[0]


def build_row_labels(grid):
    labels = {}
    for row in range(1, len(grid) + 1):
        parts = [get_cell(grid, row, col) for col in LABEL_COLUMNS]
        text = " ".join(" ".join(p.split()) for p in parts if p).strip()
        if text:
            labels[str(row)] = text
    return labels


def check_row_labels(current, target_rows):
    """First run saves the map; later runs verify the rows we are writing."""
    if not os.path.exists(ROW_LABELS_FILE):
        with open(ROW_LABELS_FILE, "w") as handle:
            json.dump(current, handle, indent=2, sort_keys=True)
        print(f"First run: saved {len(current)} row labels to row_labels.json")
        return True

    with open(ROW_LABELS_FILE) as handle:
        saved = json.load(handle)

    changed = []
    for row in sorted(target_rows):
        key = str(row)
        was = saved.get(key, "")
        now = current.get(key, "")
        if was != now:
            changed.append((row, was, now))

    if changed:
        detail = ["The sheet's row labels no longer match row_labels.json:", ""]
        for row, was, now in changed:
            detail.append(f"  Row {row}")
            detail.append(f"    saved: {was or '(nothing)'}")
            detail.append(f"    sheet: {now or '(nothing)'}")
        detail.append("")
        detail.append(
            "The rows may have been moved or renamed. Check the sheet, then "
            "delete or update row_labels.json once you are happy."
        )
        stop("Row labels have changed since the map was saved.", "\n".join(detail))

    print(f"Row labels for {len(target_rows)} target rows match row_labels.json")
    return False


def check_expected_words(row_labels):
    problems = []
    for row, word in sorted(EXPECTED_ROW_WORDS.items()):
        actual = row_labels.get(str(row), "")
        if word.lower() not in actual.lower():
            problems.append((row, word, actual or "(empty)"))

    if problems:
        detail = ["These rows do not contain the words I expected:", ""]
        for row, word, actual in problems:
            detail.append(f"  Row {row}: expected to contain \"{word}\"")
            detail.append(f"            found: {actual}")
        detail.append("")
        detail.append("The sheet layout may have shifted. Nothing was written.")
        stop("Row layout check failed.", "\n".join(detail))

    print(f"Row layout check passed for {len(EXPECTED_ROW_WORDS)} rows")


def check_occupied(grid, col_index, cells, overwrite):
    occupied = []
    for row in sorted(cells):
        existing = get_cell(grid, row, col_index)
        if existing:
            occupied.append((cell_at(row, col_index), row, existing))

    if not occupied:
        return

    lines = ["These target cells already have content:", ""]
    for ref, row, existing in occupied:
        lines.append(f"  {ref} (row {row}): {truncate(existing, 100)}")
    lines.append("")

    if overwrite:
        lines.append("--overwrite was passed, so they will be replaced.")
        print("\n".join(lines))
        return

    lines.append("Re-run with --overwrite to replace them.")
    stop(f"{len(occupied)} target cell(s) are not empty.", "\n".join(lines))


def get_sheet_id(service):
    meta = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title))",
    ).execute()
    for sheet in meta.get("sheets", []):
        if sheet["properties"]["title"] == TAB:
            return sheet["properties"]["sheetId"]
    stop(f'The spreadsheet has no tab named "{TAB}".')


def write_values(service, col_index, cells):
    data = [
        {
            "range": f"{TAB}!{cell_at(row, col_index)}",
            "values": [[cells[row]]],
        }
        for row in sorted(cells)
    ]
    service.spreadsheets().values().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"valueInputOption": "RAW", "data": data},
    ).execute()
    print(f"Wrote {len(data)} cell(s).")


def apply_wrap(service, sheet_id, col_index, rows):
    requests = [
        {
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row - 1,
                    "endRowIndex": row,
                    "startColumnIndex": col_index,
                    "endColumnIndex": col_index + 1,
                },
                "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP"}},
                "fields": "userEnteredFormat.wrapStrategy",
            }
        }
        for row in sorted(rows)
    ]
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests},
    ).execute()
    print(f"Applied text wrap to {len(requests)} cell(s).")


def read_back(service, col_index, rows, row_labels):
    ranges = [f"{TAB}!{cell_at(row, col_index)}" for row in sorted(rows)]
    result = service.spreadsheets().values().batchGet(
        spreadsheetId=SPREADSHEET_ID,
        ranges=ranges,
    ).execute()

    table = []
    for row, value_range in zip(sorted(rows), result.get("valueRanges", [])):
        values = value_range.get("values", [[""]])
        value = values[0][0] if values and values[0] else ""
        table.append([
            cell_at(row, col_index),
            truncate(row_labels.get(str(row), ""), 28),
            truncate(value, 80),
        ])

    print("\nRead back from the sheet:\n")
    print_table(table, ["Cell", "Row label", "Value"])


# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Write a weekly report into the Borderless JULIO sheet."
    )
    parser.add_argument("input", help="JSON file with month, label and cells")
    parser.add_argument("--dry-run", action="store_true",
                        help="print the plan without writing anything")
    parser.add_argument("--overwrite", action="store_true",
                        help="allow replacing cells that already have content")
    parser.add_argument("--check-rows", action="store_true",
                        help="re-run the row layout word check on a later run")
    parser.add_argument("--refresh-labels", action="store_true",
                        help="rewrite row_labels.json from the sheet as it is now")
    args = parser.parse_args()

    month, label, cells = read_payload(args.input)

    print(f"Spreadsheet : {SPREADSHEET_ID}")
    print(f"Tab         : {TAB}")
    print(f"Target      : {month} / {label}")
    print(f"Rows        : {', '.join(str(r) for r in sorted(cells))}")
    if args.dry_run:
        print("Mode        : DRY RUN (nothing will be written)")
    print()

    service = load_service()

    grid = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=READ_RANGE,
    ).execute().get("values", [])
    if not grid:
        stop(f"{READ_RANGE} came back empty.")

    # 1. locate the column
    month_row, label_row, col_index = find_column(grid, month, label)
    print(f"Matched column {col_letter(col_index)} "
          f"(row {month_row} = {month}, row {label_row} = {label})")

    # 2. row label map
    row_labels = build_row_labels(grid)
    if args.refresh_labels:
        with open(ROW_LABELS_FILE, "w") as handle:
            json.dump(row_labels, handle, indent=2, sort_keys=True)
        print("--refresh-labels: rewrote row_labels.json from the sheet.")
        first_run = True
    else:
        first_run = check_row_labels(row_labels, cells.keys())

    # 3. layout word check
    if first_run or args.check_rows:
        check_expected_words(row_labels)

    # 4. occupied cells
    check_occupied(grid, col_index, cells, args.overwrite)

    print("\nPlan:\n")
    plan = [
        [
            cell_at(row, col_index),
            truncate(row_labels.get(str(row), ""), 28),
            truncate(cells[row], 60),
        ]
        for row in sorted(cells)
    ]
    print_table(plan, ["Cell", "Row label", "Value to write"])

    if args.dry_run:
        print("\nDry run: nothing was written. "
              "Re-run without --dry-run to apply.")
        return

    # 5-6. write, then wrap
    print()
    sheet_id = get_sheet_id(service)
    write_values(service, col_index, cells)
    apply_wrap(service, sheet_id, col_index, cells.keys())

    # 7. read back
    read_back(service, col_index, cells.keys(), row_labels)


if __name__ == "__main__":
    main()

"""
sheets_writer.py
────────────────
Writes Program Floor analysis data directly to Google Sheets using the
Google Sheets API via a service-account credential.

Setup (one-time, in Speckle Automate → Function Settings → Secrets):
  GOOGLE_CREDENTIALS_JSON  — full service account JSON (paste as one line)
  GOOGLE_SHARE_EMAIL       — (optional) your Google account email to share the sheet with

No URL needs to be pasted per run.
"""

import json
import os
from typing import Optional

# Load .env relative to THIS file's directory (works locally and in Speckle Automate)
try:
    from dotenv import load_dotenv
    _env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(dotenv_path=_env_path, override=False)
except ImportError:
    pass  # python-dotenv not installed, environment vars must be set externally


# Column definitions — must match SpeckleSheetReceiver.gs
ANALYSIS_COLUMNS = [
    "Occupancy", "Level", "Program", "Area", "Status",
    "Area_OffPeak", "Area_Morning", "Area_Afternoon", "Area_Evening",
]

TIMING_COLUMNS = [
    "Occupancy", "Time Band", "Time Range (s)", "Example Clock", "Area (mm)", "Area (m²)",
]

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _get_client():
    """Build an authenticated gspread client.

    Credential resolution order:
      1. google_credentials.json in same directory as this script (deployed with code)
      2. GOOGLE_CREDENTIALS_FILE env var — path to a service account JSON file
      3. GOOGLE_CREDENTIALS_JSON env var — raw JSON string (Speckle Automate secrets)
    """
    import gspread
    from google.oauth2.service_account import Credentials

    # 1. Try google_credentials.json in same directory (deployed with code)
    script_dir = os.path.dirname(os.path.abspath(__file__))
    default_creds_file = os.path.join(script_dir, "google_credentials.json")
    if os.path.isfile(default_creds_file):
        creds = Credentials.from_service_account_file(default_creds_file, scopes=_SCOPES)
        return gspread.authorize(creds)

    # 2. Try file path from env var (local dev with downloaded JSON key)
    creds_file = os.getenv("GOOGLE_CREDENTIALS_FILE", "").strip()
    # If relative path, resolve it relative to this script's directory
    if creds_file and not os.path.isabs(creds_file):
        creds_file = os.path.join(script_dir, creds_file)
    if creds_file and os.path.isfile(creds_file):
        creds = Credentials.from_service_account_file(creds_file, scopes=_SCOPES)
        return gspread.authorize(creds)

    # 3. Try raw JSON string (Speckle Automate function variable)
    creds_raw = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
    if creds_raw:
        creds_dict = json.loads(creds_raw)
        creds = Credentials.from_service_account_info(creds_dict, scopes=_SCOPES)
        return gspread.authorize(creds)

    raise EnvironmentError(
        "No Google credentials found. Place google_credentials.json in the project root, "
        "or set GOOGLE_CREDENTIALS_FILE / GOOGLE_CREDENTIALS_JSON."
    )


def _get_or_create_spreadsheet(gc, title: str, spreadsheet_id: str | None = None):
    import gspread

    if spreadsheet_id:
        return gc.open_by_key(spreadsheet_id)

    try:
        return gc.open(title)
    except gspread.SpreadsheetNotFound:
        return gc.create(title)


def _write_sheet(spreadsheet, sheet_name: str, columns: list, rows: list, header_color: dict):
    """Write rows to a named worksheet, creating it if needed."""
    import gspread

    try:
        ws = spreadsheet.worksheet(sheet_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=500, cols=len(columns) + 2)

    if not rows:
        ws.update("A1", [["No data."]])
        return

    # Build 2-D list: header + data
    data_rows = [
        [row.get(col, "") if row else "" for col in columns]
        for row in rows
        if row  # skip empty separator dicts
    ]
    all_values = [columns] + data_rows
    ws.update("A1", all_values)

    # Format header row
    header_range = ws.range(1, 1, 1, len(columns))
    for cell in header_range:
        cell.value = columns[cell.col - 1]
    ws.format(
        f"A1:{chr(64 + len(columns))}1",
        {
            "backgroundColor": header_color,
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        },
    )
    ws.freeze(rows=1)


def _write_dynamic_sheet(spreadsheet, sheet_name: str, rows: list, header_color: dict):
    """Write arbitrary rows to a named worksheet using the row dict keys as columns."""
    import gspread

    try:
        ws = spreadsheet.worksheet(sheet_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        column_count = max(len(rows[0].keys()) if rows else 1, 1) + 2
        ws = spreadsheet.add_worksheet(title=sheet_name, rows=max(len(rows) + 10, 100), cols=column_count)

    if not rows:
        ws.update("A1", [["No data."]])
        return

    columns = list(rows[0].keys())
    values = [columns] + [[row.get(column, "") for column in columns] for row in rows]
    ws.update("A1", values)
    ws.format(
        f"A1:{chr(64 + len(columns))}1",
        {
            "backgroundColor": header_color,
            "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
        },
    )
    ws.freeze(rows=1)


def write_to_google_sheets(
    sheet_title: str,
    rows: list,
    timing_rows: list,
    spreadsheet_id: str | None = None,
) -> str:
    """
    Create or update a Google Spreadsheet with two sheets:
      Sheet 1 "Analysis"        — per occupancy/level/program data
      Sheet 2 "Occupancy Timing" — area at each time band
    Returns the spreadsheet URL.
    Raises EnvironmentError if credentials are not configured.
    """
    gc = _get_client()
    spreadsheet = _get_or_create_spreadsheet(gc, sheet_title, spreadsheet_id=spreadsheet_id)

    # Sheet 1 — Analysis
    _write_sheet(
        spreadsheet, "Analysis", ANALYSIS_COLUMNS, rows,
        {"red": 0.102, "green": 0.451, "blue": 0.910},  # #1a73e8
    )

    # Sheet 2 — Occupancy Timing (skip header-dict row)
    data_timing = timing_rows[1:] if len(timing_rows) > 1 else []
    if data_timing:
        _write_sheet(
            spreadsheet, "Occupancy Timing", TIMING_COLUMNS, data_timing,
            {"red": 0.059, "green": 0.616, "blue": 0.345},  # #0f9d58
        )

    # Move Analysis to front
    try:
        ws = spreadsheet.worksheet("Analysis")
        spreadsheet.reorder_worksheets([ws] + [
            s for s in spreadsheet.worksheets() if s.title != "Analysis"
        ])
    except Exception:
        pass

    # Share with user email if provided
    share_email = os.getenv("GOOGLE_SHARE_EMAIL", "").strip()
    if share_email:
        try:
            spreadsheet.share(share_email, perm_type="user", role="writer", notify=False)
        except Exception:
            pass

    # Make the sheet publicly viewable by anyone with the link
    try:
        spreadsheet.share("", perm_type="anyone", role="reader")
    except Exception:
        pass

    return spreadsheet.url


def write_collection_areas_to_google_sheets(
    sheet_title: str,
    sheets: dict[str, list[dict]],
    spreadsheet_id: str | None = None,
) -> str:
    """Create or update a spreadsheet with one worksheet per collection area table."""
    gc = _get_client()
    spreadsheet = _get_or_create_spreadsheet(gc, sheet_title, spreadsheet_id=spreadsheet_id)

    worksheet_order = []
    header_color = {"red": 0.180, "green": 0.490, "blue": 0.196}

    for sheet_name, rows in sheets.items():
        _write_dynamic_sheet(spreadsheet, sheet_name, rows, header_color)
        worksheet_order.append(spreadsheet.worksheet(sheet_name))

    if worksheet_order:
        try:
            spreadsheet.reorder_worksheets(worksheet_order)
        except Exception:
            pass

    share_email = os.getenv("GOOGLE_SHARE_EMAIL", "").strip()
    if share_email:
        try:
            spreadsheet.share(share_email, perm_type="user", role="writer", notify=False)
        except Exception:
            pass

    try:
        spreadsheet.share("", perm_type="anyone", role="reader")
    except Exception:
        pass

    return spreadsheet.url

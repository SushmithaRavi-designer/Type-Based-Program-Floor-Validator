"""
sheets_writer.py
────────────────
Writes collection area data to Google Sheets using a service-account credential.

Credential resolution order (first match wins):
  1. google_credentials.json  — file in same directory as this script
  2. GOOGLE_CREDENTIALS_FILE  — env var pointing to a JSON key file path
  3. GOOGLE_CREDENTIALS_JSON  — env var containing raw JSON string
  4. GOOGLE_CREDENTIALS_JSON_BASE64 — env var containing base64-encoded JSON

Set GOOGLE_SHARE_EMAIL to automatically share the sheet with your account.
Set GOOGLE_SPREADSHEET_ID to write into an existing sheet instead of creating one.
"""

from __future__ import annotations

import ast
import base64
import json
import os
import re
from typing import Optional

try:
    from dotenv import load_dotenv
    _env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    load_dotenv(dotenv_path=_env_path, override=False)
except ImportError:
    pass


_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─────────────────────────────────────────────────────────────────────────────
# Credential parsing
# ─────────────────────────────────────────────────────────────────────────────

def _parse_credentials_json(raw_value: str) -> dict:
    """
    Parse a service-account credential dict from any of these formats:
      - Plain JSON object string
      - Double-encoded JSON string  (json string inside a json string)
      - Python-literal dict string
      - Base64-encoded JSON (standard or URL-safe, with or without padding)

    Raises ValueError with a clear message if none of the formats match.
    """
    raw = (raw_value or "").strip().strip('"').strip("'")
    if not raw:
        raise ValueError(
            "GOOGLE_CREDENTIALS_JSON is empty. "
            "Paste your service-account JSON into the GitHub secret or env var."
        )

    # 1. Plain JSON
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, dict):
            return parsed
        if isinstance(parsed, str):
            inner = json.loads(parsed)
            if isinstance(inner, dict):
                return inner
    except json.JSONDecodeError:
        pass

    # 2. Python literal dict (e.g. copied from a Python repr)
    try:
        lit = ast.literal_eval(raw)
        if isinstance(lit, dict):
            return lit
    except Exception:
        pass

    # 3. Base64-encoded JSON (handles missing padding and URL-safe alphabet)
    compact = re.sub(r"\s+", "", raw).replace(" ", "+")
    padded  = compact + "=" * ((4 - len(compact) % 4) % 4)
    for decode_fn in (base64.b64decode, base64.urlsafe_b64decode):
        try:
            decoded = json.loads(decode_fn(padded).decode("utf-8"))
            if isinstance(decoded, dict):
                return decoded
        except Exception:
            pass

    raise ValueError(
        "Could not parse Google credentials. "
        "Make sure GOOGLE_CREDENTIALS_JSON contains the full service-account JSON "
        "(copy the entire .json file content). "
        f"Value starts with: {raw[:60]!r}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Client construction
# ─────────────────────────────────────────────────────────────────────────────

def _get_client():
    """Return an authenticated gspread client. Raises with a clear message on failure."""
    import gspread
    from google.oauth2.service_account import Credentials

    script_dir = os.path.dirname(os.path.abspath(__file__))

    # 1. google_credentials.json next to this script
    default_file = os.path.join(script_dir, "google_credentials.json")
    if os.path.isfile(default_file):
        creds = Credentials.from_service_account_file(default_file, scopes=_SCOPES)
        return gspread.authorize(creds)

    # 2. GOOGLE_CREDENTIALS_FILE env var (file path)
    creds_file = os.getenv("GOOGLE_CREDENTIALS_FILE", "").strip()
    if creds_file:
        if not os.path.isabs(creds_file):
            creds_file = os.path.join(script_dir, creds_file)
        if os.path.isfile(creds_file):
            creds = Credentials.from_service_account_file(creds_file, scopes=_SCOPES)
            return gspread.authorize(creds)
        raise FileNotFoundError(
            f"GOOGLE_CREDENTIALS_FILE points to '{creds_file}' but the file does not exist."
        )

    # 3. Raw JSON string(s) from env vars
    candidates = [
        ("GOOGLE_CREDENTIALS_JSON",        os.getenv("GOOGLE_CREDENTIALS_JSON",        "").strip()),
        ("GOOGLE_CREDENTIALS_JSON_BASE64",  os.getenv("GOOGLE_CREDENTIALS_JSON_BASE64", "").strip()),
    ]

    last_error: Exception | None = None
    for env_name, raw in candidates:
        if not raw:
            continue
        try:
            creds_dict = _parse_credentials_json(raw)
        except ValueError as ex:
            last_error = Exception(f"{env_name} parse error: {ex}")
            continue

        # JSON parsed OK — now try to build credentials and authorise
        try:
            creds = Credentials.from_service_account_info(creds_dict, scopes=_SCOPES)
        except Exception as ex:
            last_error = Exception(
                f"{env_name}: credentials JSON was parsed successfully "
                f"(client_email={creds_dict.get('client_email','?')}) "
                f"but Credentials.from_service_account_info failed: {ex}"
            )
            continue

        try:
            return gspread.authorize(creds)
        except Exception as ex:
            last_error = Exception(
                f"{env_name}: authorisation failed for "
                f"{creds_dict.get('client_email','?')}: {ex}"
            )
            continue

    if last_error:
        raise last_error

    raise EnvironmentError(
        "No Google credentials found. Options:\n"
        "  A) Place google_credentials.json in the project root.\n"
        "  B) Set GOOGLE_CREDENTIALS_FILE to the path of your JSON key file.\n"
        "  C) Set GOOGLE_CREDENTIALS_JSON to the full contents of your JSON key file.\n"
        "  D) Set GOOGLE_CREDENTIALS_JSON_BASE64 to a base64-encoded version of the JSON."
    )


# ─────────────────────────────────────────────────────────────────────────────
# Spreadsheet helpers
# ─────────────────────────────────────────────────────────────────────────────

def _extract_spreadsheet_id(raw: str | None) -> str | None:
    if not raw:
        return None
    raw = raw.strip()
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
    return match.group(1) if match else raw


def _get_or_create_spreadsheet(gc, title: str, spreadsheet_id: str | None = None):
    import gspread

    sid = _extract_spreadsheet_id(spreadsheet_id)
    if sid:
        try:
            return gc.open_by_key(sid)
        except Exception:
            raise ValueError(
                f"Spreadsheet '{sid}' not found. "
                "Check the ID/URL and make sure the service account has Editor access."
            )
    try:
        return gc.open(title)
    except gspread.SpreadsheetNotFound:
        return gc.create(title)


def _share_spreadsheet(spreadsheet) -> None:
    """Share with GOOGLE_SHARE_EMAIL (writer) and make publicly readable."""
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


# ─────────────────────────────────────────────────────────────────────────────
# Sheet writers
# ─────────────────────────────────────────────────────────────────────────────

PROGRAM_GS_COLORS = {
    "medical": {"red": 0.835, "green": 0.910, "blue": 0.824},     # light green
    "transit": {"red": 0.855, "green": 0.910, "blue": 0.988},     # light blue
    "retail": {"red": 1.0, "green": 0.902, "blue": 0.8},          # light orange
    "amenities": {"red": 0.882, "green": 0.835, "blue": 0.906},   # light purple
    "residential": {"red": 1.0, "green": 0.949, "blue": 0.8},     # light yellow
    "office": {"red": 0.973, "green": 0.804, "blue": 0.804},      # light red
    "parking": {"red": 0.961, "green": 0.961, "blue": 0.961},     # light grey
}

def _write_dynamic_sheet(
    spreadsheet,
    sheet_name: str,
    rows: list[dict],
    header_color: dict,
) -> None:
    """Write a list of dicts to a worksheet, using dict keys as column headers."""
    import gspread

    # Get or create worksheet
    try:
        ws = spreadsheet.worksheet(sheet_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        col_count = max(len(rows[0]) if rows else 1, 1) + 2
        ws = spreadsheet.add_worksheet(
            title=sheet_name,
            rows=max(len(rows) + 20, 100),
            cols=col_count,
        )

    if not rows:
        ws.update("A1", [["No data."]])
        return

    columns   = list(rows[0].keys())
    values    = [columns] + [[row.get(col, "") for col in columns] for row in rows]
    ws.update("A1", values)

    ws.format(
        f"A1:{chr(64 + len(columns))}1",
        {
            "backgroundColor": header_color,
            "textFormat": {
                "bold": True,
                "foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0},
            },
        },
    )
    ws.freeze(rows=1)

    # Apply specific background colors to 'Program' cells
    if "Program" in columns:
        prog_col_idx = columns.index("Program")
        requests = []
        for row_idx, row in enumerate(rows, 1): # 1-based data rows (header is 0)
            prog_val = str(row.get("Program", "")).lower().strip()
            color = PROGRAM_GS_COLORS.get(prog_val)
            if color:
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": ws.id,
                            "startRowIndex": row_idx,
                            "endRowIndex": row_idx + 1,
                            "startColumnIndex": prog_col_idx,
                            "endColumnIndex": prog_col_idx + 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": color
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })
        
        if requests:
            import gspread.utils
            spreadsheet.batch_update({"requests": requests})


# ─────────────────────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────────────────────

def write_collection_areas_to_google_sheets(
    sheet_title: str,
    sheets: dict[str, list[dict]],
    spreadsheet_id: str | None = None,
) -> str:
    """
    Write one worksheet per entry in `sheets` to a Google Spreadsheet.

    Args:
        sheet_title:    Name used when creating a new spreadsheet.
        sheets:         {sheet_name: [row_dict, ...]} mapping.
        spreadsheet_id: Optional existing spreadsheet ID or URL.
                        Falls back to GOOGLE_SPREADSHEET_ID env var.

    Returns:
        The URL of the spreadsheet.

    Raises:
        EnvironmentError: If no credentials are configured.
        ValueError:       If credentials are malformed or the sheet is not accessible.
    """
    # Resolve spreadsheet ID: argument > env var
    resolved_id = (
        _extract_spreadsheet_id(spreadsheet_id)
        or _extract_spreadsheet_id(os.getenv("GOOGLE_SPREADSHEET_ID", "").strip())
        or None
    )

    try:
        gc = _get_client()
    except Exception as ex:
        raise RuntimeError(f"Google Sheets auth failed: {ex}") from ex

    try:
        spreadsheet = _get_or_create_spreadsheet(gc, sheet_title, spreadsheet_id=resolved_id)
    except Exception as ex:
        raise RuntimeError(f"Could not open/create spreadsheet: {ex}") from ex

    header_color     = {"red": 0.180, "green": 0.490, "blue": 0.196}  # dark green
    worksheet_order  = []

    for name, rows in sheets.items():
        _write_dynamic_sheet(spreadsheet, name, rows, header_color)
        worksheet_order.append(spreadsheet.worksheet(name))

    if worksheet_order:
        try:
            spreadsheet.reorder_worksheets(worksheet_order)
        except Exception:
            pass

    _share_spreadsheet(spreadsheet)
    return spreadsheet.url
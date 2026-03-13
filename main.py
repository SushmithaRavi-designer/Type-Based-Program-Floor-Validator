"""
main.py — Speckle Automate Function
Exports collection area data to Excel and/or Google Sheets.

Credentials are read from environment variables set in the GitHub Action
or Speckle Automate runtime. The UI input fields are optional overrides.
"""

from __future__ import annotations

import os
import re
from enum import Enum
from collections import defaultdict

from pydantic import AliasChoices, ConfigDict, Field, SecretStr
from speckle_automate import AutomateBase, AutomationContext, execute_automate_function

from flatten import flatten_base
from extractor import get_param_value, extract_numeric_value


# ─────────────────────────────────────────────────────────────────────────────
# Enums
# ─────────────────────────────────────────────────────────────────────────────

class OutputFormat(str, Enum):
    EXCEL         = "excel"
    GOOGLE_SHEETS = "google_sheets"
    BOTH          = "both"


# ─────────────────────────────────────────────────────────────────────────────
# Input Schema
# ─────────────────────────────────────────────────────────────────────────────

class FunctionInputs(AutomateBase):
    model_config = ConfigDict(populate_by_name=True)

    output_format: OutputFormat = Field(
        default=OutputFormat.EXCEL,
        alias="outputFormat",
        validation_alias=AliasChoices("output_format", "outputFormat"),
        serialization_alias="outputFormat",
        title="Output Format",
        description="Select output destination: excel, google_sheets, or both.",
    )

    google_credentials_json: SecretStr = Field(
        default=SecretStr(""),
        alias="googleCredentialsJson",
        validation_alias=AliasChoices("google_credentials_json", "googleCredentialsJson"),
        serialization_alias="googleCredentialsJson",
        title="Google Credentials JSON (optional override)",
        description=(
            "Leave blank — credentials are read from the GOOGLE_CREDENTIALS_JSON "
            "environment variable. Only fill this to override for a single run."
        ),
    )

    google_share_email: str = Field(
        default="",
        alias="googleShareEmail",
        validation_alias=AliasChoices("google_share_email", "googleShareEmail"),
        serialization_alias="googleShareEmail",
        title="Google Share Email (optional override)",
        description=(
            "Leave blank — read from GOOGLE_SHARE_EMAIL env var. "
            "Only fill to override for a single run."
        ),
    )

    google_spreadsheet_id: str = Field(
        default="",
        alias="googleSpreadsheetId",
        validation_alias=AliasChoices("google_spreadsheet_id", "googleSpreadsheetId"),
        serialization_alias="googleSpreadsheetId",
        title="Google Spreadsheet ID (optional override)",
        description=(
            "Leave blank — read from GOOGLE_SPREADSHEET_ID env var. "
            "Accepts a bare ID or a full Google Sheets URL. "
            "Only fill to switch to a different sheet for a single run."
        ),
    )


# ─────────────────────────────────────────────────────────────────────────────
# Area Parsing
# ─────────────────────────────────────────────────────────────────────────────

def _parse_area_value(raw) -> float:
    """
    Parse an area value from string, number, or dict {"value": ..., "units": ...}.
    Auto-converts ft² → m² and mm² → m². Returns 0.0 on failure.
    """
    if raw is None:
        return 0.0

    # Dict with value + units
    if isinstance(raw, dict):
        val   = raw.get("value")
        units = str(raw.get("units", "")).lower()
        area  = _parse_area_value(val)
        if area > 0:
            if "ft" in units or "feet" in units:
                return round(area / 10.764, 2)
            if "mm" in units:
                return round(area / 1_000_000, 2)
        return area

    # Numeric
    if isinstance(raw, (int, float)):
        try:
            v = float(raw)
            return round(v, 2) if v > 0 else 0.0
        except (ValueError, TypeError):
            return 0.0

    # String — extract numeric part and check for inline units
    raw_str = str(raw).strip().lower()
    numeric = extract_numeric_value(raw_str)
    if numeric:
        try:
            v = float(numeric)
            if v > 0:
                if "ft" in raw_str or "feet" in raw_str:
                    return round(v / 10.764, 2)
                if "mm" in raw_str and "²" in raw_str:
                    return round(v / 1_000_000, 2)
                return round(v, 2)
        except (ValueError, TypeError):
            pass
    return 0.0


def _extract_area_from_properties(obj) -> float:
    """Read Area from obj.properties (dict or object), returning m²."""
    properties = getattr(obj, "properties", None)
    candidates = []

    if isinstance(properties, dict):
        # Primary path: properties > Parameters > Instance Parameters > Dimensions > Area
        try:
            candidates.append(
                properties["Parameters"]["Instance Parameters"]["Dimensions"]["Area"]
            )
        except (KeyError, TypeError):
            pass
        # Fallback paths
        candidates += [
            properties.get("Area"),
            properties.get("area"),
        ]
    elif properties is not None:
        candidates += [
            getattr(properties, "Area", None),
            getattr(properties, "area", None),
        ]

    for candidate in candidates:
        area = _parse_area_value(candidate)
        if area > 0:
            return area
    return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Level / Ratio Extraction
# ─────────────────────────────────────────────────────────────────────────────

def _extract_level(obj) -> str:
    properties = getattr(obj, "properties", None)
    if isinstance(properties, dict):
        for key in ("Level", "level"):
            v = properties.get(key)
            if v and str(v).strip():
                return str(v).strip()
    elif properties is not None:
        for key in ("Level", "level"):
            v = getattr(properties, key, None)
            if v and str(v).strip():
                return str(v).strip()
    return "Unknown"


def _parse_ratio(raw) -> float | None:
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return round(float(raw), 2)
    numeric = extract_numeric_value(str(raw))
    if numeric is None:
        return None
    try:
        return round(float(numeric), 2)
    except (ValueError, TypeError):
        return None


def _extract_occupancy_ratios(obj) -> dict[str, float | None]:
    properties = getattr(obj, "properties", None)

    def _read(*keys):
        if isinstance(properties, dict):
            for k in keys:
                if k in properties:
                    return properties[k]
        elif properties is not None:
            for k in keys:
                v = getattr(properties, k, None)
                if v is not None:
                    return v
        return None

    return {
        "Morning Occupancy Ratio":   _parse_ratio(_read("Morning Occupancy Ratio",   "morningOccupancyRatio")),
        "Afternoon Occupancy Ratio": _parse_ratio(_read("Afternoon Occupancy Ratio", "afternoonOccupancyRatio")),
        "Evening Occupancy Ratio":   _parse_ratio(_read("Evening Occupancy Ratio",   "eveningOccupancyRatio")),
        "Night Occupancy Ratio":     _parse_ratio(_read("Night Occupancy Ratio",     "nightOccupancyRatio")),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Collection Helpers
# ─────────────────────────────────────────────────────────────────────────────

_SKIP_COLLECTIONS = {"ROOTCOLLECTION", "ROOT COLLECTION", "GRASSHOPPER MODEL", "MODEL"}

_SHEET_NAME_MAP = {
    "PROGRAM BLOCKS":  "PROGRAM BLOCK",
    "PROGRAM BLOCK":   "PROGRAM BLOCK",
    "MORNING":         "MORNING OCCUPANCY",
    "AFTERNOON":       "AFTERNOON OCCUPANCY",
    "EVENING":         "EVENING OCCUPANCY",
    "NIGHT":           "NIGHT OCCUPANCY",
}

_COLLECTION_PRIORITY = {
    "PROGRAM BLOCKS": 0, "PROGRAM BLOCK": 0,
    "MORNING": 1, "AFTERNOON": 2, "EVENING": 3, "NIGHT": 4,
}


def _normalize(name: str) -> str:
    return " ".join(str(name or "").strip().upper().split())


def _sheet_name(collection_name: str, existing: set[str]) -> str:
    normalized = _normalize(collection_name)
    base = _SHEET_NAME_MAP.get(normalized, normalized or "AREA EXPORT")[:31]
    if base not in existing:
        return base
    i = 2
    while True:
        suffix    = f"_{i}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        if candidate not in existing:
            return candidate
        i += 1


def _iter_collections(obj):
    if obj is None:
        return
    speckle_type = str(getattr(obj, "speckle_type", ""))
    elements     = getattr(obj, "elements", getattr(obj, "@elements", None)) or []
    if "Collection" in speckle_type and elements:
        yield obj
    for child in elements:
        yield from _iter_collections(child)


def _get_export_collections(root) -> dict:
    collections = {}
    for col in _iter_collections(root):
        name       = getattr(col, "name", "")
        normalized = _normalize(name)
        if not normalized or normalized in _SKIP_COLLECTIONS:
            continue
        collections[name] = col

    if collections:
        return dict(sorted(collections.items(), key=lambda kv: (_COLLECTION_PRIORITY.get(_normalize(kv[0]), 99), kv[0])))

    return {getattr(root, "name", "Model") or "Model": root}


def _level_sort_key(level: str):
    match = re.search(r"(\d+)", str(level or ""))
    return (0, int(match.group(1)), str(level)) if match else (1, 0, str(level))


def _build_rows(collection_obj) -> list[dict]:
    rows = []
    for obj in flatten_base(collection_obj):
        if "Collection" in str(getattr(obj, "speckle_type", "")):
            continue
        area = _extract_area_from_properties(obj)
        if area <= 0:
            continue
        ratios = _extract_occupancy_ratios(obj)
        rows.append({
            "Level":                     _extract_level(obj),
            "Element Name":              getattr(obj, "name", "") or "",
            "Properties Area":           round(area, 2),
            "Morning Occupancy Ratio":   ratios["Morning Occupancy Ratio"]   if ratios["Morning Occupancy Ratio"]   is not None else "",
            "Afternoon Occupancy Ratio": ratios["Afternoon Occupancy Ratio"] if ratios["Afternoon Occupancy Ratio"] is not None else "",
            "Evening Occupancy Ratio":   ratios["Evening Occupancy Ratio"]   if ratios["Evening Occupancy Ratio"]   is not None else "",
            "Night Occupancy Ratio":     ratios["Night Occupancy Ratio"]     if ratios["Night Occupancy Ratio"]     is not None else "",
        })
    rows.sort(key=lambda r: (_level_sort_key(r.get("Level", "")), str(r.get("Element Name", ""))))
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# Credential / Config Resolution
# ─────────────────────────────────────────────────────────────────────────────

def _extract_spreadsheet_id(raw: str) -> str:
    """Accept a bare spreadsheet ID or a full Google Sheets URL."""
    if not raw:
        return ""
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
    return match.group(1) if match else raw


def _resolve_config(function_inputs: FunctionInputs) -> dict:
    """
    Merge environment variables (set once in CI/Automate) with optional
    per-run UI overrides from function_inputs.

    Priority: UI input field > environment variable.
    """
    # Credentials
    creds_env   = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
    creds_input = function_inputs.google_credentials_json.get_secret_value().strip()
    credentials = creds_input or creds_env  # input wins if provided

    # Share email
    email = (
        function_inputs.google_share_email.strip()
        or os.getenv("GOOGLE_SHARE_EMAIL", "").strip()
    )

    # Spreadsheet ID
    raw_id = (
        function_inputs.google_spreadsheet_id.strip()
        or os.getenv("GOOGLE_SPREADSHEET_ID", "").strip()
    )
    spreadsheet_id = _extract_spreadsheet_id(raw_id) or None

    # Push resolved values back so downstream helpers (sheets_writer etc.) can
    # read them from os.environ without needing extra arguments.
    if credentials:
        os.environ["GOOGLE_CREDENTIALS_JSON"] = credentials
    if email:
        os.environ["GOOGLE_SHARE_EMAIL"] = email
    if spreadsheet_id:
        os.environ["GOOGLE_SPREADSHEET_ID"] = spreadsheet_id

    return {
        "credentials":    credentials,
        "email":          email,
        "spreadsheet_id": spreadsheet_id,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Export Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _do_excel_export(automate_context: AutomationContext, sheet_rows: dict) -> str:
    from csv_exporter import rows_to_excel_multi_sheet
    excel_path = rows_to_excel_multi_sheet(sheet_rows)
    named_path = excel_path.replace(".xlsx", "_collection_areas.xlsx")
    try:
        os.rename(excel_path, named_path)
        excel_path = named_path
    except OSError:
        pass
    automate_context.store_file_result(excel_path)
    return "Excel workbook attached."


def _do_google_sheets_export(sheet_rows: dict, spreadsheet_id: str | None) -> str:
    from sheets_writer import write_collection_areas_to_google_sheets
    url = write_collection_areas_to_google_sheets(
        "Collection_Area_Export",
        sheet_rows,
        spreadsheet_id=spreadsheet_id,
    )
    return url


# ─────────────────────────────────────────────────────────────────────────────
# Main Automate Function
# ─────────────────────────────────────────────────────────────────────────────

def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    # ── Resolve credentials / config ─────────────────────────────────────────
    config = _resolve_config(function_inputs)

    # ── Receive version and build sheet rows ──────────────────────────────────
    root        = automate_context.receive_version()
    collections = _get_export_collections(root)

    sheet_rows      : dict[str, list[dict]] = {}
    collection_counts: dict[str, int]        = {}
    collection_areas : dict[str, float]      = {}
    total_area       = 0.0

    for col_name, col_obj in collections.items():
        rows = _build_rows(col_obj)
        if not rows:
            continue
        sname                     = _sheet_name(col_name, set(sheet_rows.keys()))
        sheet_rows[sname]         = rows
        collection_counts[sname]  = len(rows)
        area                      = sum(r.get("Properties Area", 0) for r in rows)
        collection_areas[sname]   = area
        total_area               += area

    if not sheet_rows:
        automate_context.mark_run_failed(
            "No objects with a numeric Area value inside properties were found. "
            "Ensure each collection contains geometry objects with properties.Area set."
        )
        return

    # ── Exports ───────────────────────────────────────────────────────────────
    export_parts : list[str] = []
    warnings     : list[str] = []
    google_url               = ""

    fmt = function_inputs.output_format

    if fmt in (OutputFormat.EXCEL, OutputFormat.BOTH):
        try:
            export_parts.append(_do_excel_export(automate_context, sheet_rows))
        except Exception as ex:
            warnings.append(f"Excel export failed: {ex}")

    if fmt in (OutputFormat.GOOGLE_SHEETS, OutputFormat.BOTH):
        try:
            google_url = _do_google_sheets_export(sheet_rows, config["spreadsheet_id"])
            export_parts.append(f"Google Sheets: {google_url}")
        except Exception as ex:
            msg = str(ex)
            if "quota" in msg.lower():
                msg += " Use an existing sheet via GOOGLE_SPREADSHEET_ID env var."
            if "404" in msg:
                msg += " Check the spreadsheet ID and ensure the service account has Editor access."
            warnings.append(f"Google Sheets export failed: {msg}")
            # Fall back to Excel if Sheets failed and we haven't already done it
            if fmt == OutputFormat.GOOGLE_SHEETS and not export_parts:
                try:
                    export_parts.append(_do_excel_export(automate_context, sheet_rows))
                    warnings.append("Fell back to Excel because Google Sheets export failed.")
                except Exception as ex2:
                    warnings.append(f"Excel fallback also failed: {ex2}")

    if not export_parts:
        automate_context.mark_run_failed(" | ".join(warnings) or "No output was generated.")
        return

    # ── Build success message ─────────────────────────────────────────────────
    _OCC_SHEETS = {"MORNING OCCUPANCY", "AFTERNOON OCCUPANCY", "EVENING OCCUPANCY", "NIGHT OCCUPANCY"}
    occ_total   = sum(v for k, v in collection_areas.items() if k in _OCC_SHEETS)

    def _ratio(sname: str) -> float:
        if occ_total <= 0:
            return 0.0
        return round(collection_areas.get(sname, 0.0) / occ_total * 100, 2)

    def _first_ratio(sname: str, col: str) -> float | None:
        for row in sheet_rows.get(sname, []):
            v = _parse_ratio(row.get(col))
            if v is not None:
                return v
        return None

    occ_lines = "\n".join([
        f"Morning Occupancy Ratio:   {_first_ratio('MORNING OCCUPANCY',   'Morning Occupancy Ratio')   or _ratio('MORNING OCCUPANCY')}%",
        f"Afternoon Occupancy Ratio: {_first_ratio('AFTERNOON OCCUPANCY', 'Afternoon Occupancy Ratio') or _ratio('AFTERNOON OCCUPANCY')}%",
        f"Evening Occupancy Ratio:   {_first_ratio('EVENING OCCUPANCY',   'Evening Occupancy Ratio')   or _ratio('EVENING OCCUPANCY')}%",
        f"Night Occupancy Ratio:     {_first_ratio('NIGHT OCCUPANCY',     'Night Occupancy Ratio')     or _ratio('NIGHT OCCUPANCY')}%",
    ])

    google_line  = f"Google Sheets: {google_url}\n" if google_url else ""
    warning_line = f"\nWarnings: {' | '.join(warnings)}" if warnings else ""
    details      = " | ".join(f"{n}: {c} rows" for n, c in collection_counts.items())

    automate_context.mark_run_success(
        f"Collection area export complete.\n"
        f"{google_line}"
        f"Sheets exported: {', '.join(sheet_rows.keys())}\n"
        f"Total rows: {sum(collection_counts.values())}\n"
        f"Total area: {round(total_area, 2)} m²\n"
        f"Output: {' | '.join(export_parts)}\n"
        f"Details: {details}\n"
        f"{occ_lines}"
        f"{warning_line}"
    )


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    pass


if __name__ == "__main__":
    execute_automate_function(automate_function, FunctionInputs)
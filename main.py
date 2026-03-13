"""This module contains the function's business logic."""

import json
import os
import tempfile
from collections import defaultdict
from enum import Enum
from math import pi

from pydantic import Field
from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)

from flatten import flatten_base
from parser import parse_type_name
from kpi import (
    shannon_diversity,
    mono_functional_check,
    check_zone_compatibility,
    vertical_stacking_continuity,
)
from csv_exporter import rows_to_excel, rows_to_csv, rows_to_excel_multi_sheet
from extractor import get_param_value, estimate_area_from_display, get_material_color, get_level_info, extract_numeric_value


def _has_parameter_data_in_properties(obj) -> bool:
    """Check if object has parameter data nested under properties"""
    properties = getattr(obj, "properties", None)
    if isinstance(properties, dict):
        if properties.get("parameters") is not None or properties.get("Parameters") is not None or properties.get("type_parameters") is not None:
            return True
        params = properties.get("parameters") or properties.get("Parameters")
        if isinstance(params, dict) and "Type Parameters" in params:
            return True
        return False

    if properties is not None:
        params = getattr(properties, "parameters", None) or getattr(properties, "Parameters", None)
        type_params = getattr(properties, "type_parameters", None)
        if params is not None or type_params is not None:
            return True
        if isinstance(params, dict) and "Type Parameters" in params:
            return True

    return False


def _object_has_parameter_data(obj) -> bool:
    params = getattr(obj, "parameters", None)
    type_params = getattr(obj, "type_parameters", None)
    if params is not None or type_params is not None:
        return True

    if _has_parameter_data_in_properties(obj):
        return True

    elems = getattr(obj, "elements", None)
    if elems:
        try:
            for child in elems:
                if _object_has_parameter_data(child):
                    return True
        except Exception:
            for name in dir(elems):
                if name.startswith("_"):
                    continue
                try:
                    child = getattr(elems, name)
                    if _object_has_parameter_data(child):
                        return True
                except Exception:
                    continue

    return False


def _find_processable_objects(all_objects):
    results = []
    for obj in all_objects:
        try:
            if _object_has_parameter_data(obj):
                results.append(obj)
                continue
        except Exception:
            pass

        if getattr(obj, "type", None) or getattr(obj, "family", None) or getattr(obj, "name", None):
            if getattr(obj, "properties", None) is not None:
                results.append(obj)

    return results


def _parse_thresholds(func_inputs) -> dict:
    return {"Retail": 60, "Office": 75, "Housing": 80, "Exhibition": 65}


# ─────────────────────────────────────────────────────────────────────────────
# Area Extraction — PRIMARY FIX
# Path: properties["Parameters"]["Instance Parameters"]["Dimensions"]["Area"]
# ─────────────────────────────────────────────────────────────────────────────

def _extract_area_from_dimensions(obj) -> float:
    """
    Extract Area from the exact Speckle path visible in the viewer:
    properties > Parameters > Instance Parameters > Dimensions > Area

    Falls back through several alternative paths if the primary path fails.
    Returns area as float in m² (or 0.0 if not found).
    """
    properties = getattr(obj, "properties", None)

    # ── PRIMARY PATH (matches Speckle viewer screenshot) ──────────────────────
    # properties["Parameters"]["Instance Parameters"]["Dimensions"]["Area"]
    if isinstance(properties, dict):
        params = properties.get("Parameters")  # Capital P as shown in viewer
        if isinstance(params, dict):
            instance_params = params.get("Instance Parameters")
            if isinstance(instance_params, dict):
                dimensions = instance_params.get("Dimensions")
                if isinstance(dimensions, dict):
                    area_raw = dimensions.get("Area")
                    area_val = _parse_area_value(area_raw)
                    if area_val > 0:
                        return area_val

    # ── FALLBACK 1: lowercase 'parameters' ────────────────────────────────────
    if isinstance(properties, dict):
        params = properties.get("parameters")
        if isinstance(params, dict):
            instance_params = params.get("Instance Parameters")
            if isinstance(instance_params, dict):
                dimensions = instance_params.get("Dimensions")
                if isinstance(dimensions, dict):
                    area_raw = dimensions.get("Area")
                    area_val = _parse_area_value(area_raw)
                    if area_val > 0:
                        return area_val

    # ── FALLBACK 2: Nested under Type Parameters > Instance Parameters ─────────
    if isinstance(properties, dict):
        params = properties.get("Parameters") or properties.get("parameters")
        if isinstance(params, dict):
            type_params = params.get("Type Parameters")
            if isinstance(type_params, dict):
                instance_params = type_params.get("Instance Parameters")
                if isinstance(instance_params, dict):
                    dimensions = instance_params.get("Dimensions")
                    if isinstance(dimensions, dict):
                        area_raw = dimensions.get("Area")
                        area_val = _parse_area_value(area_raw)
                        if area_val > 0:
                            return area_val

    # ── FALLBACK 3: Direct Instance Parameters at properties root ─────────────
    if isinstance(properties, dict):
        instance_params = properties.get("Instance Parameters")
        if isinstance(instance_params, dict):
            dimensions = instance_params.get("Dimensions")
            if isinstance(dimensions, dict):
                area_raw = dimensions.get("Area")
                area_val = _parse_area_value(area_raw)
                if area_val > 0:
                    return area_val

    # ── FALLBACK 4: get_param_value helper ────────────────────────────────────
    for param_name in ["Area", "Instance Area", "Computed Area", "Area m2", "AreaM2"]:
        area_raw = get_param_value(obj, param_name)
        if area_raw:
            area_val = _parse_area_value(area_raw)
            if area_val > 0:
                return area_val

    return 0.0


def _parse_area_value(raw) -> float:
    """Parse area from string, number, or dict {"value": ..., "units": ...}.
    
    Auto-detects units and converts to m²:
    - If units contain 'ft' or 'feet' → convert from ft² to m² (÷ 10.764)
    - If units contain 'mm' → convert from mm² to m² (÷ 1,000,000)
    - Otherwise assume m²
    """
    if raw is None:
        return 0.0
    
    # Handle dict with value and units
    if isinstance(raw, dict):
        val = raw.get("value")
        units = str(raw.get("units", "")).lower()
        area_value = _parse_area_value(val)  # Recursive parse value
        
        if area_value > 0:
            # Unit conversion
            if "ft" in units or "feet" in units:
                # Convert ft² to m²
                return round(area_value / 10.764, 2)
            elif "mm" in units and "m" not in units.replace("mm", ""):
                # Convert mm² to m² (only if it's mm, not m)
                return round(area_value / 1_000_000, 2)
        return area_value
    
    # Handle numeric value
    if isinstance(raw, (int, float)):
        try:
            return round(float(raw), 2) if float(raw) > 0 else 0.0
        except (ValueError, TypeError):
            return 0.0
    
    # Handle string - check for units in the string
    raw_str = str(raw).strip().lower()
    numeric = extract_numeric_value(raw_str)
    if numeric:
        try:
            area_value = float(numeric)
            if area_value > 0:
                # Check for units in string
                if "ft" in raw_str or "feet" in raw_str:
                    return round(area_value / 10.764, 2)
                elif "mm" in raw_str and "²" in raw_str:
                    return round(area_value / 1_000_000, 2)
                else:
                    return round(area_value, 2)
        except (ValueError, TypeError):
            pass
    return 0.0


# ─────────────────────────────────────────────────────────────────────────────
# Google Sheets — PUBLIC CSV IMPORT LINK FIX
# ─────────────────────────────────────────────────────────────────────────────

def _generate_google_sheets_import_link(occupancy_name: str, csv_url: str = None) -> str:
    """
    Generate a Google Sheets link that directly imports a CSV from a public URL.

    If csv_url is provided (a publicly accessible CSV), this creates a Sheets
    importData formula link:
      https://docs.google.com/spreadsheets/create
      (then user pastes =IMPORTDATA("csv_url") into cell A1)

    For Speckle Automate, the CSV is stored as an attachment. Since Speckle
    attachment URLs require auth, we generate a helper link that opens a new
    Sheet pre-titled, and include the IMPORTDATA formula the user can paste.

    Returns a tuple-like string: "SHEET_LINK | =IMPORTDATA formula"
    """
    safe_name = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in occupancy_name)
    sheet_title = f"Program_Floor_{safe_name}"

    if csv_url:
        # Direct import via Google Sheets importData URL trick
        import urllib.parse
        encoded_url = urllib.parse.quote(csv_url, safe='')
        # Google Sheets can open a CSV directly if it's publicly accessible
        direct_link = f"https://docs.google.com/spreadsheets/d/create?title={sheet_title}&url={encoded_url}"
        return direct_link
    else:
        # Fallback: open new sheet + provide IMPORTDATA formula hint
        import urllib.parse
        title_encoded = urllib.parse.quote(sheet_title, safe='')
        create_link = f"https://docs.google.com/spreadsheets/create?title={title_encoded}"
        return create_link


# ─────────────────────────────────────────────────────────────────────────────
# Enums
# ─────────────────────────────────────────────────────────────────────────────

class ColorExtractionMode(str, Enum):
    ENABLED = "enabled"
    DISABLED = "disabled"


class OutputFormat(str, Enum):
    EXCEL = "excel"
    GOOGLE_SHEETS = "google_sheets"


class ReportLevel(str, Enum):
    SUMMARY = "summary"
    DETAILED = "detailed"
    VERBOSE = "verbose"


class ThresholdMode(str, Enum):
    STRICT = "strict"
    PERMISSIVE = "permissive"
    CUSTOM = "custom"


# ─────────────────────────────────────────────────────────────────────────────
# Timing-based Area Calculation
# ─────────────────────────────────────────────────────────────────────────────

# Timing bands (seconds):
#   Off-Peak : Timing > 75600 OR Timing < 32400  → 700 mm (all occupancies)
#   Morning  : 32400 ≤ Timing < 43200
#   Afternoon: 43200 ≤ Timing < 61200
#   Evening  : 61200 ≤ Timing < 75600
# Source formula: if(or(Timing>75600,Timing<32400), 700mm, if(occupancy, ...))
# All values in mm; exported as m² (÷ 1,000,000)

TIMING_AREAS_MM: dict = {
    #              off_peak  morning  afternoon  evening
    "Medical":     { "off_peak": 700, "morning": 56000, "afternoon": 70000, "evening": 42000 },
    "Hotel":       { "off_peak": 700, "morning": 28000, "afternoon": 35000, "evening": 63000 },
    "Transit":     { "off_peak": 700, "morning": 63000, "afternoon": 70000, "evening": 49000 },
    "Entertainment":{"off_peak": 700, "morning": 14000, "afternoon": 28000, "evening": 70000 },
    "Corporate":   { "off_peak": 700, "morning": 59500, "afternoon": 70000, "evening": 21000 },
    "WorkAdmin":   { "off_peak": 700, "morning": 56000, "afternoon": 66500, "evening": 14000 },
    "SkyZone":     { "off_peak": 700, "morning": 28000, "afternoon": 49000, "evening": 70000 },
    "Voids":       { "off_peak": 14000,"morning": 14000, "afternoon": 14000, "evening": 14000},
}

# Human-readable time band labels and example times for Sheet 2
TIMING_BANDS = [
    {"key": "off_peak",  "label": "Off-Peak",  "time_range": ">75600s or <32400s", "example": "00:00 / 22:00"},
    {"key": "morning",   "label": "Morning",   "time_range": "32400–43200s",        "example": "09:00–12:00"},
    {"key": "afternoon", "label": "Afternoon", "time_range": "43200–61200s",        "example": "12:00–17:00"},
    {"key": "evening",   "label": "Evening",   "time_range": "61200–75600s",        "example": "17:00–21:00"},
]


def get_area_by_timing(occupancy: str, timing_seconds: float = None) -> dict:
    """Return area (m²) for each time band for the given occupancy.
    
    If timing_seconds is provided, also returns 'current' key with the active band value.
    All values converted from mm to m² (÷ 1,000,000).
    """
    areas_mm = TIMING_AREAS_MM.get(occupancy, TIMING_AREAS_MM["Voids"])

    result = {
        "off_peak":  round(areas_mm["off_peak"]  / 1_000_000, 6),
        "morning":   round(areas_mm["morning"]   / 1_000_000, 6),
        "afternoon": round(areas_mm["afternoon"] / 1_000_000, 6),
        "evening":   round(areas_mm["evening"]   / 1_000_000, 6),
    }

    # Optionally resolve the active band
    if timing_seconds is not None:
        if timing_seconds > 75600 or timing_seconds < 32400:
            result["current"] = result["off_peak"]
        elif 32400 <= timing_seconds < 43200:
            result["current"] = result["morning"]
        elif 43200 <= timing_seconds < 61200:
            result["current"] = result["afternoon"]
        else:
            result["current"] = result["evening"]

    return result


def build_timing_sheet_rows() -> list:
    """Build rows for the Occupancy Timing Sheet (Sheet 2).
    
    Produces a table: Occupancy × Time Band → Area (mm) + Area (m²)
    This directly mirrors the formula:
      if(or(Timing>75600,Timing<32400), 700mm, if(Medical, ...), ...)
    """
    rows = []

    # Header
    rows.append({
        "Occupancy":        "Occupancy",
        "Time Band":        "Time Band",
        "Time Range (s)":   "Time Range (s)",
        "Example Clock":    "Example Clock",
        "Area (mm)":        "Area (mm)",
        "Area (m²)":        "Area (m²)",
    })

    for occupancy, bands in TIMING_AREAS_MM.items():
        for band in TIMING_BANDS:
            key       = band["key"]
            area_mm   = bands[key]
            area_m2   = round(area_mm / 1_000_000, 6)
            rows.append({
                "Occupancy":      occupancy,
                "Time Band":      band["label"],
                "Time Range (s)": band["time_range"],
                "Example Clock":  band["example"],
                "Area (mm)":      area_mm,
                "Area (m²)":      area_m2,
            })
        # Blank separator between occupancies
        rows.append({})

    return rows


# ─────────────────────────────────────────────────────────────────────────────
# Collection Detection
# ─────────────────────────────────────────────────────────────────────────────

def _get_collections_from_root(root_object) -> dict:
    collections = {}
    elements = getattr(root_object, "elements", [])
    if elements:
        for item in elements:
            speckle_type = getattr(item, "speckle_type", "")
            if "Collection" in speckle_type:
                collection_name = getattr(item, "name", f"Collection_{len(collections)}")
                collections[collection_name] = item
    return collections


def _get_generic_models_from_object(obj) -> list:
    flattened = list(flatten_base(obj))
    generic_models = _find_processable_objects(flattened)
    generic_models_prefer = [
        o for o in generic_models
        if "Generic Model" in str(getattr(o, "category", ""))
    ]
    if generic_models_prefer:
        generic_models = generic_models_prefer
    return generic_models


AREA_SHEET_NAME_MAP = {
    "PROGRAM BLOCKS": "PROGRAM BLOCK",
    "PROGRAM BLOCK": "PROGRAM BLOCK",
    "MORNING": "MORNING OCCUPANCY",
    "AFTERNOON": "AFTERNOON OCCUPANCY",
    "EVENING": "EVENING OCCUPANCY",
    "NIGHT": "NIGHT OCCUPANCY",
}

AREA_COLLECTION_SKIP_NAMES = {
    "ROOTCOLLECTION",
    "ROOT COLLECTION",
    "GRASSHOPPER MODEL",
    "MODEL",
}


def _normalize_collection_name(name: str) -> str:
    return " ".join(str(name or "").strip().upper().split())


def _collection_sheet_name(collection_name: str, existing_names: set[str] | None = None) -> str:
    normalized = _normalize_collection_name(collection_name)
    base_name = AREA_SHEET_NAME_MAP.get(normalized, normalized or "AREA EXPORT")
    base_name = base_name[:31]

    if not existing_names or base_name not in existing_names:
        return base_name

    suffix_index = 2
    while True:
        suffix = f"_{suffix_index}"
        candidate = f"{base_name[:31 - len(suffix)]}{suffix}"
        if candidate not in existing_names:
            return candidate
        suffix_index += 1


def _iter_collection_nodes(obj):
    if obj is None:
        return

    speckle_type = str(getattr(obj, "speckle_type", ""))
    elements = getattr(obj, "elements", getattr(obj, "@elements", None)) or []

    if "Collection" in speckle_type and elements:
        yield obj

    for child in elements:
        yield from _iter_collection_nodes(child)


def _extract_area_from_properties(obj) -> float:
    properties = getattr(obj, "properties", None)
    candidates = []

    if isinstance(properties, dict):
        candidates.extend([properties.get("Area"), properties.get("area")])
    elif properties is not None:
        candidates.extend([getattr(properties, "Area", None), getattr(properties, "area", None)])

    for candidate in candidates:
        area_val = _parse_area_value(candidate)
        if area_val > 0:
            return area_val

    return 0.0


def _extract_level_from_properties(obj) -> str:
    properties = getattr(obj, "properties", None)

    if isinstance(properties, dict):
        for key in ("Level", "level"):
            value = properties.get(key)
            if value is not None and str(value).strip():
                return str(value).strip()
    elif properties is not None:
        for key in ("Level", "level"):
            value = getattr(properties, key, None)
            if value is not None and str(value).strip():
                return str(value).strip()

    return "Unknown"


def _collection_sort_key(collection_name: str):
    normalized = _normalize_collection_name(collection_name)
    priority = {
        "PROGRAM BLOCKS": 0,
        "PROGRAM BLOCK": 0,
        "MORNING": 1,
        "AFTERNOON": 2,
        "EVENING": 3,
        "NIGHT": 4,
    }
    return (priority.get(normalized, 99), normalized)


def _get_area_export_collections(root_object) -> dict:
    collections = {}

    for collection in _iter_collection_nodes(root_object):
        collection_name = getattr(collection, "name", "")
        normalized = _normalize_collection_name(collection_name)
        if not normalized or normalized in AREA_COLLECTION_SKIP_NAMES:
            continue
        collections[collection_name] = collection

    if collections:
        return {
            name: collections[name]
            for name in sorted(collections.keys(), key=_collection_sort_key)
        }

    fallback_name = getattr(root_object, "name", "Model") or "Model"
    return {fallback_name: root_object}


def _build_collection_area_rows(collection_obj) -> list[dict]:
    rows = []

    for obj in flatten_base(collection_obj):
        speckle_type = str(getattr(obj, "speckle_type", ""))
        if "Collection" in speckle_type:
            continue

        area = _extract_area_from_properties(obj)
        if area <= 0:
            continue

        level = _extract_level_from_properties(obj)

        row = {
            "Level": level,
            "Element Name": getattr(obj, "name", "") or "",
            "Properties Area": round(area, 2),
        }

        units = getattr(obj, "units", "") or ""
        if units:
            row["Units"] = units

        rows.append(row)

    rows.sort(key=lambda row: (str(row.get("Element Name", "")), str(row.get("Element Id", ""))))
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# Input Schema
# ─────────────────────────────────────────────────────────────────────────────

class FunctionInputs(AutomateBase):
    output_format: OutputFormat = Field(
        default=OutputFormat.EXCEL,
        title="Output Format",
        description="Select the output format for the analysis results",
    )
    color_extraction_mode: ColorExtractionMode = Field(
        default=ColorExtractionMode.ENABLED,
        title="Color Extraction Mode",
        description="Enable or disable material color extraction from Speckle objects",
    )
    threshold_mode: ThresholdMode = Field(
        default=ThresholdMode.CUSTOM,
        title="Threshold Mode",
        description="STRICT: Apply thresholds strictly | PERMISSIVE: Allow higher tolerance | CUSTOM: Use defined threshold matrix",
    )
    default_threshold: float = Field(
        default=70.0,
        title="Default Threshold (%)",
        description="Fallback threshold used if a program is not defined in the matrix (0-100)",
        ge=0.0,
        le=100.0,
    )
    report_level: ReportLevel = Field(
        default=ReportLevel.DETAILED,
        title="Report Detail Level",
        description="SUMMARY: Basic statistics only | DETAILED: Full analysis | VERBOSE: Include all metadata",
    )

# ─────────────────────────────────────────────────────────────────────────────
# Main Automate Function
# ─────────────────────────────────────────────────────────────────────────────

def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    version_root_object = automate_context.receive_version()
    collections = _get_area_export_collections(version_root_object)

    sheet_rows = {}
    collection_counts = {}
    total_area = 0.0

    for collection_name, collection_obj in collections.items():
        rows = _build_collection_area_rows(collection_obj)
        if not rows:
            continue

        sheet_name = _collection_sheet_name(collection_name, set(sheet_rows.keys()))
        sheet_rows[sheet_name] = rows
        collection_counts[sheet_name] = len(rows)
        total_area += sum(row.get("Properties Area", 0) for row in rows)

    if not sheet_rows:
        automate_context.mark_run_failed(
            "No objects with properties.Area were found in the received collections. "
            "Ensure each collection contains geometry objects with a numeric Area value inside properties."
        )
        return

    export_summary = ""

    if function_inputs.output_format == OutputFormat.GOOGLE_SHEETS:
        try:
            from sheets_writer import write_collection_areas_to_google_sheets

            spreadsheet_url = write_collection_areas_to_google_sheets(
                "Collection_Area_Export",
                sheet_rows,
            )
            export_summary = f"Google Sheets: {spreadsheet_url}"
        except EnvironmentError:
            excel_path = rows_to_excel_multi_sheet(sheet_rows)
            named_path = excel_path.replace(".xlsx", "_collection_areas.xlsx")
            try:
                os.rename(excel_path, named_path)
                excel_path = named_path
            except Exception:
                pass
            try:
                automate_context.store_file_result(excel_path)
            finally:
                if os.path.exists(excel_path):
                    os.unlink(excel_path)
            export_summary = "Google Sheets credentials were not available, so an Excel workbook was attached instead."
    else:
        excel_path = rows_to_excel_multi_sheet(sheet_rows)
        named_path = excel_path.replace(".xlsx", "_collection_areas.xlsx")
        try:
            os.rename(excel_path, named_path)
            excel_path = named_path
        except Exception:
            pass
        try:
            automate_context.store_file_result(excel_path)
        finally:
            if os.path.exists(excel_path):
                os.unlink(excel_path)
        export_summary = "Excel workbook attached."

    sheet_descriptions = [
        f"{sheet_name}: {row_count} area rows"
        for sheet_name, row_count in collection_counts.items()
    ]

    automate_context.mark_run_success(
        "Collection area export complete.\n"
        f"Sheets: {', '.join(sheet_rows.keys())}\n"
        f"Rows exported: {sum(collection_counts.values())}\n"
        f"Total properties area: {round(total_area, 2)} m²\n"
        f"{export_summary}\n"
        f"Details: {'; '.join(sheet_descriptions)}"
    )


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _zones_for_program(zone_data: dict, program: str) -> str:
    matched = sorted(z for z, progs in zone_data.items() if program in progs)
    return ", ".join(matched) if matched else "N/A"


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    pass


if __name__ == "__main__":
    execute_automate_function(automate_function, FunctionInputs)
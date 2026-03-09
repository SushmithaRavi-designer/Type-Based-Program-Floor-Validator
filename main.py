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
    """Parse area from string, number, or dict {"value": ..., "units": ...}."""
    if raw is None:
        return 0.0
    if isinstance(raw, (int, float)):
        try:
            return round(float(raw), 2) if float(raw) > 0 else 0.0
        except (ValueError, TypeError):
            return 0.0
    if isinstance(raw, dict):
        val = raw.get("value")
        return _parse_area_value(val)
    # String — strip units and parse
    numeric = extract_numeric_value(str(raw))
    if numeric:
        try:
            return round(float(numeric), 2) if float(numeric) > 0 else 0.0
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

def get_area_by_timing(occupancy: str, timing_seconds: float = None) -> dict:
    occupancy_areas = {
        "Medical":       {"off_peak": 700,   "morning": 56000, "afternoon": 70000, "evening": 42000},
        "Hotel":         {"off_peak": 700,   "morning": 28000, "afternoon": 35000, "evening": 63000},
        "Transit":       {"off_peak": 700,   "morning": 63000, "afternoon": 70000, "evening": 49000},
        "Entertainment": {"off_peak": 700,   "morning": 14000, "afternoon": 28000, "evening": 70000},
        "Corporate":     {"off_peak": 700,   "morning": 59500, "afternoon": 70000, "evening": 21000},
        "WorkAdmin":     {"off_peak": 700,   "morning": 56000, "afternoon": 66500, "evening": 14000},
        "SkyZone":       {"off_peak": 700,   "morning": 28000, "afternoon": 49000, "evening": 70000},
        "Voids":         {"off_peak": 14000, "morning": 14000, "afternoon": 14000, "evening": 14000},
    }

    if occupancy not in occupancy_areas:
        occupancy = "Voids"

    areas_mm = occupancy_areas[occupancy]
    return {
        "off_peak":  round(areas_mm["off_peak"]  / 1_000_000, 4),
        "morning":   round(areas_mm["morning"]   / 1_000_000, 4),
        "afternoon": round(areas_mm["afternoon"] / 1_000_000, 4),
        "evening":   round(areas_mm["evening"]   / 1_000_000, 4),
    }


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
    # ── 1. Receive version ────────────────────────────────────────────────────
    version_root_object = automate_context.receive_version()

    # ── 2. Detect collections ─────────────────────────────────────────────────
    collections = _get_collections_from_root(version_root_object)

    if not collections:
        all_objects = list(flatten_base(version_root_object))
        generic_models = _find_processable_objects(all_objects)
        generic_models_prefer = [
            obj for obj in generic_models
            if "Generic Model" in str(getattr(obj, "category", ""))
        ]
        if generic_models_prefer:
            generic_models = generic_models_prefer
        collections = {"Model": version_root_object}

    # ── 2.5. Extract objects per collection ───────────────────────────────────
    collections_models = {}
    for collection_name, collection_obj in collections.items():
        collection_generic_models = _get_generic_models_from_object(collection_obj)
        if collection_generic_models:
            collections_models[collection_name] = collection_generic_models

    if not collections_models:
        automate_context.mark_run_failed(
            "No processable elements found in any collection. "
            "Ensure your model contains Revit Generic Model elements with Parameters > Instance Parameters > Dimensions > Area."
        )
        return

    # ── 2.7. Debug: Inspect first element ─────────────────────────────────────
    debug_info = []
    total_models = sum(len(models) for models in collections_models.values())
    debug_info.append(f"Collections: {len(collections_models)}, Total processable objects: {total_models}")

    first_obj = None
    for models in collections_models.values():
        if models:
            first_obj = models[0]
            break

    if first_obj:
        debug_info.append(f"speckle_type: {getattr(first_obj, 'speckle_type', 'NOT FOUND')}")
        debug_info.append(f"category: {getattr(first_obj, 'category', 'NOT FOUND')}")
        debug_info.append(f"type: {getattr(first_obj, 'type', 'NOT FOUND')}")
        debug_info.append(f"family: {getattr(first_obj, 'family', 'NOT FOUND')}")

        properties = getattr(first_obj, "properties", None)
        if isinstance(properties, dict):
            debug_info.append(f"properties keys (top-level): {list(properties.keys())[:20]}")
            params = properties.get("Parameters") or properties.get("parameters")
            if isinstance(params, dict):
                debug_info.append(f"Parameters keys: {list(params.keys())}")
                inst_params = params.get("Instance Parameters")
                if isinstance(inst_params, dict):
                    debug_info.append(f"Instance Parameters keys: {list(inst_params.keys())}")
                    dims = inst_params.get("Dimensions")
                    if isinstance(dims, dict):
                        debug_info.append(f"Dimensions keys: {list(dims.keys())}")
                        area_val = dims.get("Area")
                        debug_info.append(f"Area raw value: {area_val}")
                    else:
                        debug_info.append("Dimensions NOT found under Instance Parameters")
                else:
                    debug_info.append("Instance Parameters NOT found under Parameters")
            else:
                debug_info.append("Parameters dict NOT found in properties")
        else:
            debug_info.append(f"properties is not a dict: {type(properties)}")

        # Test the new extraction function
        test_area = _extract_area_from_dimensions(first_obj)
        debug_info.append(f"_extract_area_from_dimensions result: {test_area} m²")

    # ── 3. Thresholds ─────────────────────────────────────────────────────────
    thresholds = _parse_thresholds(function_inputs)

    if function_inputs.threshold_mode == ThresholdMode.STRICT:
        thresholds = {prog: max(10, thresh - 10) for prog, thresh in thresholds.items()}
        default_threshold = max(10, function_inputs.default_threshold - 10)
    elif function_inputs.threshold_mode == ThresholdMode.PERMISSIVE:
        thresholds = {prog: min(95, thresh + 10) for prog, thresh in thresholds.items()}
        default_threshold = min(95, function_inputs.default_threshold + 10)
    else:
        default_threshold = function_inputs.default_threshold

    # ── 4. Extract program / zone / floor / area ───────────────────────────────
    collection_data = {}
    floor_data: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    zone_data:  dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    floor_data_by_occupancy: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    material_colors: dict[str, str] = {}
    element_metadata: dict = {}

    for collection_name, generic_models in collections_models.items():
        coll_floor_data = defaultdict(lambda: defaultdict(float))
        coll_zone_data = defaultdict(lambda: defaultdict(float))
        coll_floor_data_by_occupancy = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
        coll_material_colors = {}
        coll_element_metadata = {}

        for obj in generic_models:
            # Type Name → Program, Zone, Level
            raw_type = get_param_value(obj, "Type Name") or ""
            if not raw_type:
                for alt_name in ["Type Name", "Family Type", "Type", "Name", "name", "typeName"]:
                    raw_type = get_param_value(obj, alt_name)
                    if raw_type:
                        break
            if not raw_type:
                raw_type = getattr(obj, "Type Name", "") or getattr(obj, "type_name", "") or getattr(obj, "name", "") or ""

            program, zone_from_name, floor_from_name = parse_type_name(raw_type)

            if zone_from_name != "Unknown":
                zone = zone_from_name
            else:
                zone = get_param_value(obj, "Type") or get_param_value(obj, "Family") or get_param_value(obj, "Family Type") or "Unknown"
            if not zone:
                zone = "Unknown"

            level = floor_from_name if floor_from_name != "Unknown" else get_level_info(obj, "Level")
            if not level:
                level = "Unknown"

            # Material color
            material_color = None
            if function_inputs.color_extraction_mode == ColorExtractionMode.ENABLED:
                material_color = get_material_color(obj, "Material")
            if not material_color:
                material_color = "Not Found"

            # Occupancy from groupname
            occupancy = "Unknown"
            properties = getattr(obj, "properties", None)
            if isinstance(properties, dict):
                occupancy = properties.get("groupname") or "Unknown"
            elif properties is not None:
                occupancy = getattr(properties, "groupname", "Unknown") or "Unknown"
            if not occupancy or occupancy == "Unknown":
                occupancy = "Unknown"

            # ── AREA: Use new dedicated extraction function ──────────────────
            area = _extract_area_from_dimensions(obj)

            # Timing areas
            timing_areas = get_area_by_timing(occupancy)

            obj_id = getattr(obj, "id", None) or id(obj)
            metadata_entry = {
                "program": program,
                "zone": zone,
                "level": level,
                "occupancy": occupancy,
                "material_color": material_color,
                "area": area,
                "area_off_peak":  timing_areas["off_peak"],
                "area_morning":   timing_areas["morning"],
                "area_afternoon": timing_areas["afternoon"],
                "area_evening":   timing_areas["evening"],
                "speckle_type": getattr(obj, "speckle_type", "Unknown"),
                "collection": collection_name,
            }

            coll_element_metadata[obj_id] = metadata_entry
            element_metadata[obj_id] = metadata_entry

            if material_color and material_color != "Not Found":
                if program not in coll_material_colors:
                    coll_material_colors[program] = material_color

            coll_floor_data[level][program] += area
            coll_floor_data_by_occupancy[occupancy][level][program] += area
            coll_zone_data[zone][program] += area

            floor_data[level][program] += area
            floor_data_by_occupancy[occupancy][level][program] += area
            zone_data[zone][program] += area
            material_colors.update(coll_material_colors)

        collection_data[collection_name] = {
            "floor_data": coll_floor_data,
            "zone_data": coll_zone_data,
            "floor_data_by_occupancy": coll_floor_data_by_occupancy,
            "material_colors": coll_material_colors,
            "element_metadata": coll_element_metadata,
            "model_count": len(list(generic_models)) if generic_models else 0,
        }

    # ── 5. KPIs and CSV rows — single flat sheet ─────────────────────────────
    all_csv_rows: list = []   # single list for one sheet / one CSV
    issues = []
    zone_issue_objects = []

    stacking = vertical_stacking_continuity(dict(floor_data))

    objs_with_area = sum(1 for meta in element_metadata.values() if meta.get("area", 0) > 0)
    total_area_sum = sum(meta.get("area", 0) for meta in element_metadata.values())
    debug_info.append(f"Objects with area > 0: {objs_with_area} / {len(element_metadata)}")
    debug_info.append(f"Total area sum: {total_area_sum:.2f} m²")

    for occupancy in sorted(floor_data_by_occupancy.keys()):
        occupancy_floor_data = floor_data_by_occupancy[occupancy]

        for level, prog_areas in sorted(occupancy_floor_data.items()):
            total     = sum(prog_areas.values())
            diversity = shannon_diversity(prog_areas)
            is_mono, dominant, dom_pct, allowed = mono_functional_check(
                prog_areas, thresholds, function_inputs.default_threshold
            )

            if is_mono:
                level_status = f"MONO-FUNCTIONAL ({dom_pct:.1f}% > {allowed}%)"
                issues.append(
                    f"Level {level} exceeds mono-functional threshold "
                    f"({dominant} = {dom_pct:.1f}%, limit = {allowed}%)."
                )
            else:
                level_status = "OK"

            timing_areas = get_area_by_timing(occupancy)

            for program, area in sorted(prog_areas.items()):
                all_csv_rows.append({
                    "Occupancy":       occupancy,
                    "Level":           level,
                    "Program":         program,
                    "Area":            round(area, 2),
                    "Status":          level_status,
                    "Area_OffPeak":    timing_areas["off_peak"],
                    "Area_Morning":    timing_areas["morning"],
                    "Area_Afternoon":  timing_areas["afternoon"],
                    "Area_Evening":    timing_areas["evening"],
                })

    # ── 6. Zone compatibility ─────────────────────────────────────────────────
    zone_issues = check_zone_compatibility(dict(zone_data), thresholds, default_threshold)
    if zone_issues:
        issues.extend(zone_issues)
        flagged_zones = {
            issue.split("Zone ")[1].split(" ")[0]
            for issue in zone_issues
            if "Zone " in issue
        }
        zone_issue_objects += [
            elem_meta for elem_meta in element_metadata.values()
            if elem_meta.get("zone") in flagged_zones
        ]

    # ── 7. Speckle viewer pins ────────────────────────────────────────────────
    if zone_issue_objects:
        automate_context.attach_error_to_objects(
            category="Zone Program Mismatch",
            affected_objects=zone_issue_objects,
            message="This zone has an incompatible program allocation.",
        )

    # ── 8. Summary comment ───────────────────────────────────────────────────
    summary_lines = [
        f"✅ Analysed {len(element_metadata)} elements across {len(floor_data)} levels and {len(zone_data)} zones.",
        "",
        "── DEBUG: Area Extraction ──",
    ]
    summary_lines.extend(debug_info)

    objs_with_areas = [
        (meta["level"], meta["zone"], meta["area"])
        for meta in element_metadata.values() if meta.get("area", 0) > 0
    ]
    if objs_with_areas:
        summary_lines.append("")
        summary_lines.append("Sample objects with extracted areas (first 20):")
        for level, zone, area in objs_with_areas[:20]:
            summary_lines.append(f"  Level={level}, Zone={zone}, Area={area} m²")

    summary_lines.append("")

    if material_colors and function_inputs.report_level in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
        summary_lines.append("── Material Colors by Program ──")
        for prog in sorted(material_colors.keys()):
            summary_lines.append(f"  {prog}: #{material_colors[prog]}")
        summary_lines.append("")

    if issues:
        summary_lines.append("⚠️  Issues detected:")
        summary_lines += [f"  • {i}" for i in issues]
    else:
        summary_lines.append("✅ All levels pass program allocation thresholds.")

    if function_inputs.report_level in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
        summary_lines += ["", "── Level Summary ──"]
        for level, prog_areas in sorted(floor_data.items()):
            total    = sum(prog_areas.values())
            dominant = max(prog_areas, key=prog_areas.get)
            pct      = prog_areas[dominant] / total * 100 if total else 0
            summary_lines.append(f"  {level}: total={total:.1f} m²  |  dominant={dominant} ({pct:.1f}%)")

        summary_lines += ["", "── Zone Summary ──"]
        for zone, prog_areas in sorted(zone_data.items()):
            programs = ", ".join(sorted(prog_areas.keys()))
            summary_lines.append(f"  {zone}: [{programs}]")

    if function_inputs.report_level == ReportLevel.VERBOSE:
        summary_lines += ["", "── Configuration ──"]
        summary_lines += [
            f"  Threshold Mode: {function_inputs.threshold_mode.value}",
            f"  Color Extraction: {function_inputs.color_extraction_mode.value}",
            f"  Default Threshold: {function_inputs.default_threshold}%",
            f"  Report Level: {function_inputs.report_level.value}",
        ]

    # ── 10. Append overall summary rows to single sheet ───────────────────────
    total_area_all  = sum(meta.get("area", 0) for meta in element_metadata.values())
    ok_count        = sum(1 for row in all_csv_rows if row.get("Status") == "OK")
    mono_count      = sum(1 for row in all_csv_rows if "MONO-FUNCTIONAL" in row.get("Status", ""))

    all_csv_rows.append({})   # blank separator row
    all_csv_rows.append({"Occupancy": "SUMMARY", "Level": "AGGREGATION SUMMARY", "Program": "", "Area": "", "Status": ""})
    all_csv_rows.append({"Occupancy": "Total Area (m²)",         "Level": "", "Program": "", "Area": round(total_area_all, 2), "Status": ""})
    all_csv_rows.append({"Occupancy": "OK Entries",              "Level": "", "Program": "", "Area": ok_count,                "Status": ""})
    all_csv_rows.append({"Occupancy": "MONO-FUNCTIONAL Entries", "Level": "", "Program": "", "Area": mono_count,              "Status": ""})

    # ── 11. Export — single file (one sheet / one CSV) ───────────────────────
    import urllib.parse

    if not all_csv_rows:
        all_csv_rows = [{"Occupancy": "No data", "Level": "", "Program": "", "Area": 0, "Status": ""}]

    if function_inputs.output_format == OutputFormat.GOOGLE_SHEETS:
        export_format = "CSV (for Google Sheets)"

        csv_content_str = rows_to_csv(all_csv_rows)
        tmp_file = tempfile.NamedTemporaryFile(
            suffix=".csv", delete=False,
            prefix="program_floor_analysis_",
            mode='w', encoding='utf-8',
        )
        tmp_file.write(csv_content_str)
        tmp_file.close()

        try:
            automate_context.store_file_result(tmp_file.name)
        except Exception:
            pass
        finally:
            try:
                if os.path.exists(tmp_file.name):
                    os.unlink(tmp_file.name)
            except Exception:
                pass

    else:
        export_format = "Excel"
        # Single sheet named "Analysis"
        csv_content = rows_to_excel_multi_sheet({"Analysis": all_csv_rows})
        try:
            automate_context.store_file_result(csv_content)
        finally:
            if os.path.exists(csv_content):
                os.unlink(csv_content)

    if issues:
        automate_context.set_context_view()

    # ── 12. Mark success ──────────────────────────────────────────────────────
    total_area_val = sum(meta.get("area", 0) for meta in element_metadata.values())
    total_elements = len(element_metadata)
    unique_occupancies = sorted(floor_data_by_occupancy.keys())

    sheet_title = "Program_Floor_Analysis"
    title_encoded = urllib.parse.quote(sheet_title, safe='')
    gs_link = f"https://docs.google.com/spreadsheets/create?title={title_encoded}"

    success_msg = (
        f"✅ Program floor analysis complete ({export_format} format — single sheet)\n"
        f"Processed: {total_elements} elements | {len(floor_data)} levels | "
        f"{len(zone_data)} zones | {len(collection_data)} collection(s)\n"
        f"Total area: {total_area_val:.2f} m² | Occupancies: {', '.join(unique_occupancies)}\n\n"
        f"📊 GOOGLE SHEETS:\n"
        f"  🔗 Create new sheet: {gs_link}\n"
        f"  📋 Then: File > Import > Upload the downloaded CSV\n"
    )

    automate_context.mark_run_success(success_msg)


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
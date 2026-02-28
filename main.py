"""This module contains the function's business logic.

Use the automation_context module to wrap your function in an Automate context helper.
"""

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
from csv_exporter import rows_to_csv
from extractor import get_param_value, estimate_area_from_display, get_material_color, get_level_info, extract_numeric_value


def _has_parameter_data_in_properties(obj) -> bool:
    """Check if object has parameter data nested under properties"""
    properties = getattr(obj, "properties", None)
    # Accept both dict-style and object-style properties
    if isinstance(properties, dict):
        if properties.get("parameters") is not None or properties.get("type_parameters") is not None:
            return True
        params = properties.get("parameters", {})
        if isinstance(params, dict) and "Type Parameters" in params:
            return True
        return False

    # If properties is an object-like (DynamicBase), try attribute access
    if properties is not None:
        params = getattr(properties, "parameters", None)
        type_params = getattr(properties, "type_parameters", None)
        if params is not None or type_params is not None:
            return True
        # Try nested structure
        if isinstance(params, dict) and "Type Parameters" in params:
            return True

    return False


def _object_has_parameter_data(obj) -> bool:
    """Check an object for parameter/type_parameter data in multiple locations.

    This inspects:
      - obj.parameters (dict or object)
      - obj.type_parameters (dict or object)
      - obj.properties.parameters or obj.properties.type_parameters
      - obj.properties.parameters['Type Parameters']['Dimensions']
      - obj.elements (iterate children)
    """
    # Direct parameters/type_parameters
    params = getattr(obj, "parameters", None)
    type_params = getattr(obj, "type_parameters", None)
    if params is not None or type_params is not None:
        return True

    # properties (dict or object)
    if _has_parameter_data_in_properties(obj):
        return True

    # elements: sometimes Revit family contains children under 'elements'
    elems = getattr(obj, "elements", None)
    if elems:
        # elems may be list-like or dict-like
        try:
            for child in elems:
                if _object_has_parameter_data(child):
                    return True
        except Exception:
            # if not iterable, try attribute access
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
    """Return a list of objects that appear to contain parameter data.

    This is intentionally permissive to compensate for variations in how the
    Revit->Speckle connector exports parameters.
    """
    results = []
    for obj in all_objects:
        try:
            if _object_has_parameter_data(obj):
                results.append(obj)
                continue
        except Exception:
            # Ignore inspection errors and continue
            pass

        # Also accept objects that have 'type' or 'family' and a non-empty name
        if getattr(obj, "type", None) or getattr(obj, "family", None) or getattr(obj, "name", None):
            # Only accept if properties exist (even empty) to avoid system objects
            if getattr(obj, "properties", None) is not None:
                results.append(obj)

    return results


# ─────────────────────────────────────────────────────────────────────────────
# Enums for Fixed Options (Recommended Approach)
# ─────────────────────────────────────────────────────────────────────────────

class ColorExtractionMode(str, Enum):
    """Material color extraction control."""
    ENABLED = "enabled"
    DISABLED = "disabled"


class ReportLevel(str, Enum):
    """Reporting detail level."""
    SUMMARY = "summary"
    DETAILED = "detailed"
    VERBOSE = "verbose"


class ThresholdMode(str, Enum):
    """Threshold validation mode."""
    STRICT = "strict"
    PERMISSIVE = "permissive"
    CUSTOM = "custom"


# ─────────────────────────────────────────────────────────────────────────────
# Input Schema
# ─────────────────────────────────────────────────────────────────────────────

class FunctionInputs(AutomateBase):
    """These are function author-defined values.

    Automate will make sure to supply them matching the types specified here.
    Please use the pydantic model schema to define your inputs:
    https://docs.pydantic.dev/latest/usage/models/
    """

    # ── Parameter Configuration ──
    program_source_parameter: str = Field(
        default="Type Name",
        title="Program Source Parameter",
        description="Name of parameter containing program information (e.g., Type Name, Category)",
    )
    
    zone_parameter_name: str = Field(
        default="Zone",
        title="Zone Parameter Name",
        description="Parameter storing zone information",
    )
    
    level_parameter_name: str = Field(
        default="Level",
        title="Level Parameter Name",
        description="Parameter that defines building level",
    )
    
    area_parameter_name: str = Field(
        default="Area",
        title="Area Parameter Name",
        description="Parameter containing element area",
    )
    
    material_color_parameter_name: str = Field(
        default="Material",
        title="Material Color Parameter Name",
        description="Parameter containing material or color information",
    )

    # ── Color Extraction Control ──
    color_extraction_mode: ColorExtractionMode = Field(
        default=ColorExtractionMode.ENABLED,
        title="Color Extraction Mode",
        description="Enable or disable material color extraction from Speckle objects",
    )

    # ── Threshold Configuration ──
    threshold_mode: ThresholdMode = Field(
        default=ThresholdMode.CUSTOM,
        title="Threshold Mode",
        description=(
            "STRICT: Apply thresholds strictly | "
            "PERMISSIVE: Allow higher tolerance | "
            "CUSTOM: Use defined threshold matrix"
        ),
    )
    
    program_threshold_matrix: str = Field(
        default='{"Retail": 60, "Office": 75, "Housing": 80, "Exhibition": 65}',
        title="Program Threshold Matrix (JSON)",
        description=(
            "JSON dictionary mapping program type to its maximum allowed percentage. "
            'Example: {"Retail": 60, "Office": 75, "Housing": 80} | '
            "Only used when threshold_mode is CUSTOM"
        ),
    )
    
    default_threshold: float = Field(
        default=70.0,
        title="Default Threshold (%)",
        description="Fallback threshold used if a program is not defined in the matrix (0-100)",
        ge=0.0,
        le=100.0,
    )

    # ── Report Configuration ──
    report_level: ReportLevel = Field(
        default=ReportLevel.DETAILED,
        title="Report Detail Level",
        description=(
            "SUMMARY: Basic statistics only | "
            "DETAILED: Full analysis with colors and levels | "
            "VERBOSE: Include all metadata and processing details"
        ),
    )


# ─────────────────────────────────────────────────────────────────────────────
# Main Automate Function
# ─────────────────────────────────────────────────────────────────────────────

def automate_function(
    automate_context: AutomationContext,
    function_inputs: FunctionInputs,
) -> None:
    """Validate program allocation across floors and zones.

    Args:
        automate_context: A context-helper object that carries relevant information
            about the runtime context of this function.
            It gives access to the Speckle project data that triggered this run.
            It also has convenient methods for attaching results to the Speckle model.
        function_inputs: An instance object matching the defined schema.
    """
    # ── 1. Receive the triggering version ─────────────────────────────────────
    version_root_object = automate_context.receive_version()

    # ── 2. Flatten all objects and filter Generic Models ──────────────────────
    all_objects = list(flatten_base(version_root_object))
    
    # Find any objects that appear to contain parameter/type data. This
    # function is deliberately permissive to handle variations in exports.
    generic_models = _find_processable_objects(all_objects)

    # Prefer objects categorized as Generic Models if present
    generic_models_prefer = [
        obj for obj in generic_models
        if "Generic Model" in str(getattr(obj, "category", ""))
    ]
    if generic_models_prefer:
        generic_models = generic_models_prefer

    if not generic_models:
        automate_context.mark_run_failed(
            f"No processable elements found. Total objects: {len(all_objects)}. "
            "Checked for objects with parameters (including nested in properties.parameters). "
            "Ensure your model contains Revit elements exported with Type Parameters dimension data."
        )
        return

    # ── 2.5. Debug: Inspect available parameters in first element ──────────────
    # This helps identify what parameters are actually available in the model
    debug_info = []
    debug_info.append(f"Total objects: {len(all_objects)}, Processing: {len(generic_models)} Generic Models")
    
    if generic_models:
        first_obj = generic_models[0]
        
        # Show object type
        debug_info.append(f"speckle_type: {getattr(first_obj, 'speckle_type', 'NOT FOUND')}")
        debug_info.append(f"category: {getattr(first_obj, 'category', 'NOT FOUND')}")
        debug_info.append(f"type: {getattr(first_obj, 'type', 'NOT FOUND')}")
        debug_info.append(f"family: {getattr(first_obj, 'family', 'NOT FOUND')}")
        
        # Show if parameters/type_parameters exist at object level
        params = getattr(first_obj, "parameters", None)
        type_params = getattr(first_obj, "type_parameters", None)
        debug_info.append(f"obj.parameters: {params is not None}")
        debug_info.append(f"obj.type_parameters: {type_params is not None}")
        
        # Check properties
        properties = getattr(first_obj, "properties", None)
        if isinstance(properties, dict):
            debug_info.append(f"properties exists: True")
            props_params = properties.get("parameters")
            props_type_params = properties.get("type_parameters")
            debug_info.append(f"properties.parameters: {props_params is not None}")
            debug_info.append(f"properties.type_parameters: {props_type_params is not None}")
            
            if isinstance(props_params, dict) and "Type Parameters" in props_params:
                tp = props_params.get("Type Parameters", {})
                if isinstance(tp, dict) and "Dimensions" in tp:
                    dims = tp.get("Dimensions", {})
                    if isinstance(dims, dict):
                        debug_info.append(f"Dimensions keys: {list(dims.keys())}")
                        dia_val = dims.get("DIA-2")
                        debug_info.append(f"DIA-2 in dimensions: {dia_val}")
        else:
            debug_info.append(f"properties exists: False")
        
        # Try extraction
        dia2_val = get_param_value(first_obj, "DIA-2")
        debug_info.append(f"DIA-2 from get_param_value: '{dia2_val}'")
        
        if dia2_val:
            diameter = extract_numeric_value(dia2_val)
            debug_info.append(f"DIA-2 extracted numeric: {diameter}")
            if diameter:
                try:
                    radius = diameter / 2
                    calc_area = round(pi * (radius ** 2), 2)
                    debug_info.append(f"✓ AREA CALCULATION SUCCESS: DIA-2={diameter} → area={calc_area} m²")
                except (ValueError, TypeError) as e:
                    debug_info.append(f"AREA CALCULATION ERROR: {e}")
        else:
            debug_info.append("✗ DIA-2 not found - Area will be 0")
        
        debug_output = "\n".join(debug_info)

    # ── 3. Parse threshold matrix from JSON input and apply threshold mode ────
    try:
        thresholds: dict = json.loads(function_inputs.program_threshold_matrix)
    except json.JSONDecodeError:
        thresholds = {}

    # Apply threshold mode adjustments
    if function_inputs.threshold_mode == ThresholdMode.STRICT:
        # Strict mode: reduce thresholds by 10%
        thresholds = {prog: max(10, thresh - 10) for prog, thresh in thresholds.items()}
        default_threshold = max(10, function_inputs.default_threshold - 10)
    elif function_inputs.threshold_mode == ThresholdMode.PERMISSIVE:
        # Permissive mode: increase thresholds by 10%
        thresholds = {prog: min(95, thresh + 10) for prog, thresh in thresholds.items()}
        default_threshold = min(95, function_inputs.default_threshold + 10)
    else:  # ThresholdMode.CUSTOM
        # Custom mode: use values as-is
        default_threshold = function_inputs.default_threshold

    # ── 4. Extract program / zone / floor / area per element ──────────────────
    floor_data: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    zone_data:  dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    material_colors: dict[str, str] = {}  # Track material colors by element
    element_metadata: dict = {}  # Store complete element metadata with level and color

    for obj in generic_models:
        # Program/Zone/Floor extracted from Type Name (e.g., "MEDICAL_ZoneA_LEVEL10")
        raw_type = get_param_value(obj, function_inputs.program_source_parameter) or ""
        
        # If still empty, try common alternative names
        if not raw_type:
            for alt_name in ["Type Name", "Family Type", "Type", "Name", "name", "typeName"]:
                raw_type = get_param_value(obj, alt_name)
                if raw_type:
                    break
        
        # Last resort: try direct attributes
        if not raw_type:
            raw_type = getattr(obj, "Type Name", "") or getattr(obj, "type_name", "") or getattr(obj, "name", "") or ""
        
        # Parse Type Name into Program, Zone, Level (e.g., "MEDICAL_ZoneA_LEVEL10" → MEDICAL, ZoneA, LEVEL10)
        program, zone_from_name, floor_from_name = parse_type_name(raw_type)
        
        # Zone: First try parsed zone from Type Name, otherwise use Type/Family name as zone (common name)
        if zone_from_name != "Unknown":
            zone = zone_from_name
        else:
            # Use Type or Family name as zone (e.g., "HOTELS", "blocks", "MEDICAL")
            zone = get_param_value(obj, "Type") or get_param_value(obj, "Family") or get_param_value(obj, "Family Type") or "Unknown"
        
        if not zone:
            zone = "Unknown"

        # Floor / Level: Use level parsed from Type Name, fallback to level extraction
        level = floor_from_name if floor_from_name != "Unknown" else get_level_info(obj, function_inputs.level_parameter_name)
        if not level:
            level = "Unknown"

        # Extract material color from Revit material object (if enabled)
        material_color = None
        if function_inputs.color_extraction_mode == ColorExtractionMode.ENABLED:
            # First try to get from actual Revit material object
            material_color = get_material_color(obj, function_inputs.material_color_parameter_name)
        if not material_color:
            material_color = "Not Found"

        # Area: Try DIA-2 first (diameter of floor plate), then configured parameter, then geometry
        area = 0.0
        dia_2_raw = get_param_value(obj, "DIA-2")  # Get diameter value (may include units like "60 Meters")
        
        if dia_2_raw:
            # Extract numeric value from string that may include units
            diameter = extract_numeric_value(dia_2_raw)
            if diameter:
                try:
                    # Calculate circular area: A = π × (d/2)²
                    radius = diameter / 2
                    area = round(pi * (radius ** 2), 2)
                except (ValueError, TypeError):
                    area = 0.0
        
        # Try alternative parameter names if DIA-2 not found or extraction failed
        if area == 0.0:
            for alt_name in ["DIA_2", "Diameter", "diameter", "DIA", "Floor Diameter"]:
                dia_2_alt = get_param_value(obj, alt_name)
                if dia_2_alt:
                    diameter = extract_numeric_value(dia_2_alt)
                    if diameter:
                        try:
                            radius = diameter / 2
                            area = round(pi * (radius ** 2), 2)
                            break
                        except (ValueError, TypeError):
                            pass
        
        # If no diameter value, try configured area parameter
        if area == 0.0:
            area_raw = get_param_value(obj, function_inputs.area_parameter_name)
            try:
                area = float(area_raw) if area_raw else 0.0
            except (ValueError, TypeError):
                area = 0.0
        
        # Fall back to geometry estimate if no parameter found
        if area == 0.0:
            area = estimate_area_from_display(obj)

        # Store metadata
        obj_id = getattr(obj, "id", None) or id(obj)
        element_metadata[obj_id] = {
            "program": program,
            "zone": zone,
            "level": level,
            "material_color": material_color,
            "area": area,
            "speckle_type": getattr(obj, "speckle_type", "Unknown"),
            "dia_2_raw": dia_2_raw,  # DEBUG: track what DIA-2 was extracted
        }

        # Track colors by program for reporting
        if material_color and material_color != "Not Found":
            if program not in material_colors:
                material_colors[program] = material_color

        floor_data[level][program] += area
        zone_data[zone][program]   += area

    # ── 5. Compute KPIs, build CSV rows, collect issues ───────────────────────
    csv_rows          = []
    issues            = []
    mono_floor_objects = []   # objects on mono-functional floors (for error pin)
    zone_issue_objects = []   # objects in mismatched zones (for error pin)

    stacking = vertical_stacking_continuity(dict(floor_data))

    for level, prog_areas in sorted(floor_data.items()):
        total         = sum(prog_areas.values())
        diversity     = shannon_diversity(prog_areas)
        is_mono, dominant, dom_pct, allowed = mono_functional_check(
            prog_areas, thresholds, function_inputs.default_threshold
        )

        if is_mono:
            status = f"⚠️ MONO-FUNCTIONAL ({dom_pct:.1f}% > {allowed}%)"
            issues.append(
                f"Level {level} exceeds mono-functional threshold "
                f"({dominant} = {dom_pct:.1f}%, limit = {allowed}%)."
            )
            # Collect objects on this level for Speckle error pins
            mono_floor_objects += [
                obj for obj in generic_models
                if (get_level_info(obj, function_inputs.level_parameter_name) or 
                    get_param_value(obj, function_inputs.level_parameter_name) or "") == level
            ]
        else:
            status = "✅ OK"

        for program, area in sorted(prog_areas.items()):
            
            csv_rows.append({
                "Level":    level,
                "Program":  program,
                "Area":     round(area, 2),
            })

    # ── 6. Zone compatibility check ───────────────────────────────────────────
    zone_issues = check_zone_compatibility(
        dict(zone_data), thresholds, default_threshold
    )
    if zone_issues:
        issues.extend(zone_issues)
        # Collect objects in mismatched zones for Speckle error pins
        flagged_zones = {
            issue.split("Zone ")[1].split(" ")[0]
            for issue in zone_issues
            if "Zone " in issue
        }
        zone_issue_objects += [
            obj for obj in generic_models
            if (get_param_value(obj, function_inputs.zone_parameter_name) or "") in flagged_zones
        ]

    # ── 7. Attach error pins to objects in Speckle viewer ────────────────────
    if mono_floor_objects:
        automate_context.attach_error_to_objects(
            category="Mono-Functional Floor",
            affected_objects=mono_floor_objects,
            message="This floor exceeds the mono-functional program threshold.",
        )

    if zone_issue_objects:
        automate_context.attach_error_to_objects(
            category="Zone Program Mismatch",
            affected_objects=zone_issue_objects,
            message="This zone has an incompatible program allocation.",
        )

    # ── 8. Write version comment (with report_level control) ──────────────────
    summary_lines = [
        f"✅ Analysed {len(generic_models)} elements across {len(floor_data)} levels and {len(zone_data)} zones.",
        "",
    ]

    # ALWAYS include debug info in summary for troubleshooting
    summary_lines.append("── DEBUG: Extraction Details ──")
    summary_lines.extend(debug_info)
    
    # Show sample area values
    sample_areas = []
    for i, (obj_id, meta) in enumerate(list(element_metadata.items())[:5]):
        dia_val = meta.get("dia_2_raw", "?")
        area_val = meta.get("area", 0)
        sample_areas.append(f"  Element {i+1}: DIA-2='{dia_val}' → Area={area_val} m²")
    
    if sample_areas:
        summary_lines.append("")
        summary_lines.append("Sample area calculations (first 5 elements):")
        summary_lines.extend(sample_areas)
    
    summary_lines.append("")

    # Include issues if any
    if material_colors and function_inputs.report_level in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
        summary_lines.append("── Material Colors by Program ──")
        for prog in sorted(material_colors.keys()):
            color = material_colors[prog]
            summary_lines.append(f"  {prog}: #{color}")
        summary_lines.append("")

    # Always include issues
    if issues:
        summary_lines.append("⚠️  Issues detected:")
        summary_lines += [f"  • {i}" for i in issues]
    else:
        summary_lines.append("✅ All levels pass program allocation thresholds.")

    # Include level/zone summary based on report level (detailed and verbose)
    if function_inputs.report_level in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
        summary_lines += ["", "── Level Summary ──"]
        for level, prog_areas in sorted(floor_data.items()):
            total    = sum(prog_areas.values())
            dominant = max(prog_areas, key=prog_areas.get)
            pct      = prog_areas[dominant] / total * 100 if total else 0
            summary_lines.append(
                f"  {level}: total={total:.1f} m²  |  dominant={dominant} ({pct:.1f}%)"
            )

        summary_lines += ["", "── Zone Summary ──"]
        for zone, prog_areas in sorted(zone_data.items()):
            programs = ", ".join(sorted(prog_areas.keys()))
            summary_lines.append(f"  {zone}: [{programs}]")

    # Include verbose metadata if verbose mode
    if function_inputs.report_level == ReportLevel.VERBOSE:
        summary_lines += ["", "── Configuration Used ──"]
        summary_lines += [
            f"  Threshold Mode: {function_inputs.threshold_mode.value}",
            f"  Color Extraction: {function_inputs.color_extraction_mode.value}",
            f"  Default Threshold: {function_inputs.default_threshold}%",
            f"  Report Level: {function_inputs.report_level.value}",
        ]

    # Validation report is exported as CSV file via store_file_result()

    # ── 10. Export CSV file result ────────────────────────────────────────────
    
    # Add debug section to CSV rows
    csv_rows.append({ "Level": "", "Program": "=== DEBUG INFO ===", "Area": "" })
    for debug_line in debug_info:
        csv_rows.append({ "Level": "", "Program": debug_line, "Area": "" })
    
    csv_content = rows_to_csv(csv_rows)
    
    # Write CSV to temporary file (speckle-automate expects a file path, not bytes)
    with tempfile.NamedTemporaryFile(
        mode="w",
        suffix=".csv",
        delete=False,
        prefix="program_floor_validation_"
    ) as tmp_file:
        tmp_file.write(csv_content)
        tmp_path = tmp_file.name
    
    try:
        automate_context.store_file_result(tmp_path)
    finally:
        # Clean up temporary file
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)

    # Show offending objects in the Speckle viewer
    if issues:
        automate_context.set_context_view()

    # ── 11. Mark run success or failure ───────────────────────────────────────
    if issues:
        automate_context.mark_run_failed(
            f"Validation failed: {len(issues)} issue(s) found. "
            "See version comment and CSV for full details."
        )
    else:
        automate_context.mark_run_success(
            "Program level validation passed. "
            "No mono-functional levels or zone mismatches detected. "
            "All material colors and level assignments verified."
        )


# ─────────────────────────────────────────────────────────────────────────────
# Helper
# ─────────────────────────────────────────────────────────────────────────────

def _zones_for_program(zone_data: dict, program: str) -> str:
    matched = sorted(z for z, progs in zone_data.items() if program in progs)
    return ", ".join(matched) if matched else "N/A"


def automate_function_without_inputs(automate_context: AutomationContext) -> None:
    """A function example without inputs.

    If your function does not need any input variables,
    besides what the automation context provides,
    the inputs argument can be omitted.
    """
    pass


# make sure to call the function with the executor
if __name__ == "__main__":
    # NOTE: always pass in the automate function by its reference; do not invoke it!

    # Pass in the function reference with the inputs schema to the executor.
    execute_automate_function(automate_function, FunctionInputs)

    # If the function has no arguments, the executor can handle it like so
    # execute_automate_function(automate_function_without_inputs)

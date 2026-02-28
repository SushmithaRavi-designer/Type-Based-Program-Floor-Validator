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
from extractor import get_param_value, estimate_area_from_display, get_material_color, get_level_info


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
    
    # Try filtering for Generic Models first
    generic_models = [
        obj for obj in all_objects
        if "Generic Model" in str(getattr(obj, "category", ""))
    ]
    
    # If no Generic Models found, try other common categories
    if not generic_models:
        # Try filtering by type or speckle_type
        generic_models = [
            obj for obj in all_objects
            if getattr(obj, "speckle_type", "").lower() in ["revitgenericmodel", "revitfamilyinstance"]
            and hasattr(obj, "parameters")
        ]
    
    # If still no results, just take all objects with parameters
    if not generic_models:
        generic_models = [
            obj for obj in all_objects
            if hasattr(obj, "parameters") and not getattr(obj, "category", "").startswith("System")
        ]

    if not generic_models:
        automate_context.mark_run_failed(
            f"No processable elements found. Total objects: {len(all_objects)}. "
            "Checked for Generic Models, family instances, and parametric objects. "
            "Ensure your model contains Revit elements with parameters."
        )
        return

    # ── 2.5. Debug: Inspect available parameters in first element ──────────────
    # This helps identify what parameters are actually available in the model
    debug_info = []
    if generic_models:
        first_obj = generic_models[0]
        
        # Get direct attributes
        debug_info.append(f"First object type: {type(first_obj).__name__}")
        debug_info.append(f"Object class: {first_obj.__class__.__module__}.{first_obj.__class__.__name__}")
        
        # List all direct attributes (non-private)
        direct_attrs = [attr for attr in dir(first_obj) if not attr.startswith("_")]
        debug_info.append(f"Direct attributes: {', '.join(direct_attrs[:15])}")
        
        # Get parameters dict structure
        params = getattr(first_obj, "parameters", None)
        if params:
            debug_info.append(f"Parameters object type: {type(params).__name__}")
            if isinstance(params, dict):
                param_keys = list(params.keys())
                debug_info.append(f"Parameter keys: {param_keys[:10]}")
                # Show structure of first param
                if param_keys:
                    first_key = param_keys[0]
                    first_val = params[first_key]
                    debug_info.append(f"Example param '{first_key}': {first_val}")
            else:
                # DynamicBase-style parameters
                param_attrs = [k for k in dir(params) if not k.startswith("_")]
                debug_info.append(f"Parameter attributes: {param_attrs[:10]}")
        else:
            debug_info.append("No 'parameters' object found")
        
        # Check for common properties
        for prop_name in ["Type", "type", "name", "Type Name", "category", "speckle_type", "level", "levelName"]:
            val = getattr(first_obj, prop_name, None)
            if val:
                debug_info.append(f"  - {prop_name}: {val}")
        
        # Store debug info for later output
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
        
        # Zone: Use zone parsed from Type Name (underscore-separated), fallback to explicit parameter
        zone = zone_from_name if zone_from_name != "Unknown" else get_param_value(obj, function_inputs.zone_parameter_name)
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
        dia_2_raw = get_param_value(obj, "DIA-2")  # Get diameter value
        if dia_2_raw:
            try:
                diameter = float(dia_2_raw)
                # Calculate circular area: A = π × (d/2)²
                radius = diameter / 2
                area = round(pi * (radius ** 2), 2)
            except (ValueError, TypeError):
                area = 0.0
        
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
            pct = (area / total * 100) if total else 0
            program_color = material_colors.get(program, "Not Found")
            
            csv_rows.append({
                "Level":               level,
                "Zone":                _zones_for_program(zone_data, program),
                "Program":             program,
                "Material Color":      program_color,
                "Area":                round(area, 2),
                "%":                   round(pct, 2),
                "Dominant":            dominant,
                "Diversity Index (H)": diversity,
                "Vertical Continuity": stacking.get(program, 0.0),
                "Status":              status,
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
        f"✅ Analysed {len(generic_models)} Generic Model elements "
        f"across {len(floor_data)} levels and {len(zone_data)} zones.",
        "",
    ]

    # ALWAYS include debug info in summary for troubleshooting
    summary_lines.append("── DEBUG: Element Structure ──")
    summary_lines.extend(debug_info)
    summary_lines.append("")

    # Include material colors based on report level (detailed or verbose only)
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

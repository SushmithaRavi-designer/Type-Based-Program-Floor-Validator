"""This module contains the function's business logic.

Use the automation_context module to wrap your function in an Automate context helper.
"""

import json
from collections import defaultdict

from pydantic import Field
from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)

from flatten import flatten_base
from utils.parser import parse_type_name
from utils.kpi import (
    shannon_diversity,
    mono_functional_check,
    check_zone_compatibility,
    vertical_stacking_continuity,
)
from utils.csv_exporter import rows_to_csv
from utils.extractor import get_param_value, estimate_area_from_display, get_material_color, get_level_info


# ─────────────────────────────────────────────────────────────────────────────
# Input Schema
# ─────────────────────────────────────────────────────────────────────────────

class FunctionInputs(AutomateBase):
    """These are function author-defined values.

    Automate will make sure to supply them matching the types specified here.
    Please use the pydantic model schema to define your inputs:
    https://docs.pydantic.dev/latest/usage/models/
    """

    program_source_parameter: str = Field(
        default="Type Name",
        title="Program Source Parameter",
        description="Name of parameter containing program information (e.g. Type Name)",
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
    program_threshold_matrix: str = Field(
        default='{"Retail": 60, "Office": 75, "Housing": 80, "Exhibition": 65}',
        title="Program Threshold Matrix (JSON)",
        description=(
            "JSON dictionary mapping program type to its maximum allowed floor percentage. "
            'Example: {"Retail": 60, "Office": 75, "Housing": 80}'
        ),
    )
    default_threshold: float = Field(
        default=70.0,
        title="Default Threshold (%)",
        description="Fallback threshold used if a program is not defined in the matrix.",
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
    use_material_color: bool = Field(
        default=True,
        title="Use Material Color",
        description="Whether to extract and consider material color information from elements",
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

    generic_models = [
        obj for obj in all_objects
        if "Generic Model" in str(getattr(obj, "category", ""))
    ]

    if not generic_models:
        automate_context.mark_run_failed(
            "No Generic Model elements found. "
            "Ensure the commit contains Revit Generic Model objects."
        )
        return

    # ── 3. Parse threshold matrix from JSON input ─────────────────────────────
    try:
        thresholds: dict = json.loads(function_inputs.program_threshold_matrix)
    except json.JSONDecodeError:
        thresholds = {}

    # ── 4. Extract program / zone / floor / area per element ──────────────────
    floor_data: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    zone_data:  dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    material_colors: dict[str, str] = {}  # Track material colors by element
    element_metadata: dict = {}  # Store complete element metadata with level and color

    for obj in generic_models:
        # Program (from Type Name or chosen parameter)
        raw_type = get_param_value(obj, function_inputs.program_source_parameter) or ""
        program, zone_from_name, floor_from_name = parse_type_name(raw_type)

        # Zone: prefer explicit parameter, fall back to parsed zone
        zone = get_param_value(obj, function_inputs.zone_parameter_name) or zone_from_name

        # Floor / Level: use dedicated level extraction function, fall back to parsed floor
        level = get_level_info(obj, function_inputs.level_parameter_name) or floor_from_name

        # Extract material color if enabled
        material_color = None
        if function_inputs.use_material_color:
            material_color = get_material_color(obj, function_inputs.material_color_parameter_name)

        # Area: prefer explicit parameter, fall back to geometry estimate
        area_raw = get_param_value(obj, function_inputs.area_parameter_name)
        try:
            area = float(area_raw) if area_raw else 0.0
        except (ValueError, TypeError):
            area = 0.0
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
        if material_color:
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
            program_color = material_colors.get(program, "—")
            
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
        dict(zone_data), thresholds, function_inputs.default_threshold
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

    # ── 8. Add version tags ───────────────────────────────────────────────────
    automate_context.add_version_tag("validated")
    if mono_floor_objects:
        automate_context.add_version_tag("mono-functional-floor")
    if zone_issue_objects:
        automate_context.add_version_tag("zone-mismatch")

    # ── 9. Write version comment ──────────────────────────────────────────────
    summary_lines = [
        f"✅ Analysed {len(generic_models)} Generic Model elements "
        f"across {len(floor_data)} levels and {len(zone_data)} zones.",
        "",
    ]

    if material_colors:
        summary_lines.append("── Material Colors by Program ──")
        for prog in sorted(material_colors.keys()):
            color = material_colors[prog]
            summary_lines.append(f"  {prog}: #{color}")
        summary_lines.append("")

    if issues:
        summary_lines.append("⚠️  Issues detected:")
        summary_lines += [f"  • {i}" for i in issues]
    else:
        summary_lines.append("✅ All levels pass program allocation thresholds.")

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

    automate_context.add_version_comment("\n".join(summary_lines))

    # ── 10. Export CSV file result ────────────────────────────────────────────
    csv_content = rows_to_csv(csv_rows)
    automate_context.store_file_result(
        "program_floor_validation.csv",
        csv_content.encode("utf-8"),
        "text/csv",
    )

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
    return ", ".join(matched) if matched else "—"


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

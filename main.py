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
from csv_exporter import rows_to_excel, rows_to_csv, rows_to_excel_multi_sheet
from extractor import get_param_value, estimate_area_from_display, get_material_color, get_level_info, extract_numeric_value


def _has_parameter_data_in_properties(obj) -> bool:
    """Check if object has parameter data nested under properties"""
    properties = getattr(obj, "properties", None)
    # Accept both dict-style and object-style properties
    if isinstance(properties, dict):
        # Try both 'parameters' and 'Parameters' (case variations)
        if properties.get("parameters") is not None or properties.get("Parameters") is not None or properties.get("type_parameters") is not None:
            return True
        # Check for Type Parameters section (try both cases)
        params = properties.get("parameters") or properties.get("Parameters")
        if isinstance(params, dict) and "Type Parameters" in params:
            return True
        return False

    # If properties is an object-like (DynamicBase), try attribute access
    if properties is not None:
        params = getattr(properties, "parameters", None) or getattr(properties, "Parameters", None)
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


# compatibility helper --------------------------------------------------

def _parse_thresholds(func_inputs) -> dict:
    """Return the default threshold matrix.
    
    A hardcoded default threshold matrix is used for all program types.
    """
    return {"Retail": 60, "Office": 75, "Housing": 80, "Exhibition": 65}



# ─────────────────────────────────────────────────────────────────────────────
# Enums for Fixed Options (Recommended Approach)
# ─────────────────────────────────────────────────────────────────────────────

class ColorExtractionMode(str, Enum):
    """Material color extraction control."""
    ENABLED = "enabled"
    DISABLED = "disabled"


class OutputFormat(str, Enum):
    """Export output format selection."""
    EXCEL = "excel"
    GOOGLE_SHEETS = "google_sheets"


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
# Timing-based Area Calculation
# ─────────────────────────────────────────────────────────────────────────────

def get_area_by_timing(occupancy: str, timing_seconds: float = None) -> dict:
    """Calculate area values for each time period based on occupancy.
    
    Time periods (in seconds):
    - Off-peak: Timing > 75600s OR Timing < 32400s
    - Morning (AM): 32400s ≤ Timing < 43200s
    - Afternoon (PM): 43200s ≤ Timing < 61200s
    - Evening (EV): 61200s ≤ Timing < 75600s
    
    Returns dict with area values for: off_peak, morning, afternoon, evening (in mm, converted to m²)
    1 m² = 1,000,000 mm²
    """
    
    # Define area values (in mm) for each occupancy type and time period
    occupancy_areas = {
        "Medical": {
            "off_peak": 700,
            "morning": 56000,
            "afternoon": 70000,
            "evening": 42000
        },
        "Hotel": {
            "off_peak": 700,
            "morning": 28000,
            "afternoon": 35000,
            "evening": 63000
        },
        "Transit": {
            "off_peak": 700,
            "morning": 63000,
            "afternoon": 70000,
            "evening": 49000
        },
        "Entertainment": {
            "off_peak": 700,
            "morning": 14000,
            "afternoon": 28000,
            "evening": 70000
        },
        "Corporate": {
            "off_peak": 700,
            "morning": 59500,
            "afternoon": 70000,
            "evening": 21000
        },
        "WorkAdmin": {
            "off_peak": 700,
            "morning": 56000,
            "afternoon": 66500,
            "evening": 14000
        },
        "SkyZone": {
            "off_peak": 700,
            "morning": 28000,
            "afternoon": 49000,
            "evening": 70000
        },
        "Voids": {
            "off_peak": 14000,
            "morning": 14000,
            "afternoon": 14000,
            "evening": 14000
        }
    }
    
    # Get the area values for this occupancy, default to Voids if not found
    if occupancy not in occupancy_areas:
        occupancy = "Voids"
    
    areas_mm = occupancy_areas[occupancy]
    
    # Convert from mm to m² (1 m² = 1,000,000 mm²)
    return {
        "off_peak": round(areas_mm["off_peak"] / 1_000_000, 4),
        "morning": round(areas_mm["morning"] / 1_000_000, 4),
        "afternoon": round(areas_mm["afternoon"] / 1_000_000, 4),
        "evening": round(areas_mm["evening"] / 1_000_000, 4),
    }


# ─────────────────────────────────────────────────────────────────────────────
# Collection Detection Helpers
# ─────────────────────────────────────────────────────────────────────────────

def _get_collections_from_root(root_object) -> dict:
    """Extract all collections from the root object.
    
    Returns:
        dict: {collection_name: collection_object}
    """
    collections = {}
    
    # Try to access elements attribute
    elements = getattr(root_object, "elements", [])
    if elements:
        for item in elements:
            # Check if item is a collection (has speckle_type attribute containing "Collection")
            speckle_type = getattr(item, "speckle_type", "")
            if "Collection" in speckle_type:
                collection_name = getattr(item, "name", f"Collection_{len(collections)}")
                collections[collection_name] = item
    
    return collections


def _get_generic_models_from_object(obj) -> list:
    """Extract generic models (processable objects) from a single object or collection.
    
    Args:
        obj: Could be root object, collection, or any Speckle object
        
    Returns:
        list: Flattened list of processable generic models
    """
    flattened = list(flatten_base(obj))
    generic_models = _find_processable_objects(flattened)
    
    # Prefer objects categorized as Generic Models if present
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
    """These are function author-defined values.

    Automate will make sure to supply them matching the types specified here.
    Please use the pydantic model schema to define your inputs:
    https://docs.pydantic.dev/latest/usage/models/
    """

    # ── Output Configuration ──
    output_format: OutputFormat = Field(
        default=OutputFormat.EXCEL,
        title="Output Format",
        description="Select the output format for the analysis results",
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

    # ── 2. Detect collections in the model ────────────────────────────────────
    collections = _get_collections_from_root(version_root_object)
    
    # If collections exist, process each collection separately
    if collections:
        for collection_name, collection_obj in collections.items():
            generic_models = _get_generic_models_from_object(collection_obj)
            if generic_models:
                automate_context.log(f"Processing collection: {collection_name} with {len(generic_models)} elements")
            else:
                automate_context.log(f"Skipping collection: {collection_name} (no processable elements)")
    else:
        # Fallback: If no collections, process all objects from root
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
        
        # Wrap in a single-collection format for uniform processing
        collections = {"Model": version_root_object}

    # ── 2.5. Extract objects from each collection for processing ──────────────
    # Build a dict of collection_name -> list of generic_models
    collections_models = {}
    for collection_name, collection_obj in collections.items():
        collection_generic_models = _get_generic_models_from_object(collection_obj)
        if collection_generic_models:
            collections_models[collection_name] = collection_generic_models

    if not collections_models:
        automate_context.mark_run_failed(
            "No processable elements found in any collection. "
            "Ensure your model contains Revit elements exported with Type Parameters dimension data."
        )
        return

    # ── 2.7. Debug: Inspect available parameters in first element ──────────────
    # This helps identify what parameters are actually available in the model
    debug_info = []
    total_models = sum(len(models) for models in collections_models.values())
    debug_info.append(f"Collections detected: {len(collections_models)}, Total processable models: {total_models}")
    
    # Get first object from first collection for inspection
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
        
        # Show if parameters/type_parameters exist at object level
        params = getattr(first_obj, "parameters", None)
        type_params = getattr(first_obj, "type_parameters", None)
        debug_info.append(f"obj.parameters: {params is not None}")
        debug_info.append(f"obj.type_parameters: {type_params is not None}")
        
        # Check properties
        properties = getattr(first_obj, "properties", None)
        if isinstance(properties, dict):
            debug_info.append(f"properties exists: True (dict)")
            debug_info.append(f"properties keys: {list(properties.keys())[:20]}")  # Show first 20 keys
            
            # Try both 'parameters' and 'Parameters' (case variations)
            props_params = properties.get("parameters") or properties.get("Parameters")
            props_type_params = properties.get("type_parameters")
            debug_info.append(f"properties['Parameters']: {props_params is not None}")
            debug_info.append(f"properties['type_parameters']: {props_type_params is not None}")
            
            if isinstance(props_params, dict):
                debug_info.append(f"  Parameters type: dict")
                debug_info.append(f"  Parameters keys: {list(props_params.keys())}")
                if "Type Parameters" in props_params:
                    tp = props_params.get("Type Parameters", {})
                    if isinstance(tp, dict):
                        debug_info.append(f"  Type Parameters keys: {list(tp.keys())}")
                        if "Instance Parameters" in tp:
                            inst_params = tp.get("Instance Parameters", {})
                            if isinstance(inst_params, dict):
                                debug_info.append(f"    Instance Parameters keys: {list(inst_params.keys())}")
                                area_val = inst_params.get("Area")
                                debug_info.append(f"    Area value: {area_val}")
            
            # Try to find Area in any nested location
            for key in properties.keys():
                if "area" in key.lower():
                    debug_info.append(f"  Found key with area: '{key}'")
        elif properties is not None:
            debug_info.append(f"properties exists: True (object, not dict)")
            try:
                props_attrs = [a for a in dir(properties) if not a.startswith('_')]
                debug_info.append(f"  properties attributes (first 20): {props_attrs[:20]}")
            except:
                debug_info.append(f"  Could not enumerate properties attributes")
        else:
            debug_info.append(f"properties exists: False")
        
        # Try extraction of Area parameter
        area_param_val = get_param_value(first_obj, "Area")
        debug_info.append(f"Area from get_param_value: '{area_param_val}'")
        
        if area_param_val:
            area_extracted = extract_numeric_value(area_param_val)
            debug_info.append(f"Area extracted numeric: {area_extracted}")
            if area_extracted:
                try:
                    area_final = round(float(area_extracted), 2)
                    debug_info.append(f"SUCCESS: Area={area_final} m²")
                except (ValueError, TypeError) as e:
                    debug_info.append(f"AREA EXTRACTION ERROR: {e}")
        else:
            debug_info.append("NOTICE: Area parameter not found - Area will be 0")
        
        debug_output = "\n".join(debug_info)

    # ── 3. Parse threshold matrix from JSON input and apply threshold mode ────
    # We provide a small compatibility helper because older automation
    # environments sometimes sent the parameters wrapped in a `function`
    # object.  The log message shown by the user reported a NameError when
    # trying to access `function.inputs` directly, so we avoid referring to
    # `function` at all and instead use this helper to load the JSON safely.
    thresholds = _parse_thresholds(function_inputs)

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

    # ── 4. Extract program / zone / floor / area per element (per collection) ──
    # Initialize data tracking for each collection
    collection_data = {}  # {collection_name: {data_dict}}
    
    # Global aggregations across all collections
    floor_data: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    zone_data:  dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))
    floor_data_by_occupancy: dict[str, dict[str, dict[str, float]]] = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    material_colors: dict[str, str] = {}
    element_metadata: dict = {}
    
    # Process each collection separately
    for collection_name, generic_models in collections_models.items():
        # Initialize data structures for this collection
        coll_floor_data = defaultdict(lambda: defaultdict(float))
        coll_zone_data = defaultdict(lambda: defaultdict(float))
        coll_floor_data_by_occupancy = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
        coll_material_colors = {}
        coll_element_metadata = {}
        
        for obj in generic_models:
        # Program/Zone/Floor extracted from Type Name (e.g., "MEDICAL_ZoneA_LEVEL10")
        raw_type = get_param_value(obj, "Type Name") or ""
        
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
        level = floor_from_name if floor_from_name != "Unknown" else get_level_info(obj, "Level")
        if not level:
            level = "Unknown"

        # Extract material color from Revit material object (if enabled)
        material_color = None
        if function_inputs.color_extraction_mode == ColorExtractionMode.ENABLED:
            # First try to get from actual Revit material object
            material_color = get_material_color(obj, "Material")
        if not material_color:
            material_color = "Not Found"

        # Extract occupancy/group name from properties
        occupancy = "Unknown"
        properties = getattr(obj, "properties", None)
        if isinstance(properties, dict):
            occupancy = properties.get("groupname") or "Unknown"
        elif properties is not None:
            occupancy = getattr(properties, "groupname", "Unknown") or "Unknown"
        
        if not occupancy or occupancy == "Unknown":
            occupancy = "Unknown"

        # Area: Extract from Instance Parameters → Area
        # Path: obj.properties["Parameters"]["Type Parameters"]["Instance Parameters"]["Area"]
        area = 0.0
        area_raw = None
        
        properties = getattr(obj, "properties", None)
        if isinstance(properties, dict):
            params = properties.get("Parameters")  # Note: capital P
            if isinstance(params, dict):
                type_params = params.get("Type Parameters")
                if isinstance(type_params, dict):
                    instance_params = type_params.get("Instance Parameters")
                    if isinstance(instance_params, dict):
                        area_raw = instance_params.get("Area")
                        
                        if area_raw is not None:
                            # Area may be a string like "60" or a dict like {"value": 60, "units": "m²"}
                            if isinstance(area_raw, dict):
                                area_numeric = area_raw.get("value")
                            else:
                                area_numeric = extract_numeric_value(str(area_raw))
                            
                            if area_numeric and area_numeric > 0:
                                try:
                                    area = round(float(area_numeric), 2)
                                except (ValueError, TypeError):
                                    area = 0.0
        
        # Fallback: Try to extract Area from get_param_value if not found above
        if area == 0.0:
            area_fallback = get_param_value(obj, "Area")
            if area_fallback:
                area_numeric = extract_numeric_value(str(area_fallback))
                if area_numeric and area_numeric > 0:
                    try:
                        area = round(float(area_numeric), 2)
                        area_raw = area_fallback
                    except (ValueError, TypeError):
                        area = 0.0
        
        # Second fallback: Try alternative parameter names
        if area == 0.0:
            for alt_area_name in ["Instance Area", "Computed Area", "Area m2", "AreaM2"]:
                area_fallback = get_param_value(obj, alt_area_name)
                if area_fallback:
                    area_numeric = extract_numeric_value(str(area_fallback))
                    if area_numeric and area_numeric > 0:
                        try:
                            area = round(float(area_numeric), 2)
                            area_raw = area_fallback
                            break
                        except (ValueError, TypeError):
                            pass

        # Calculate timing-based areas for this occupancy
        timing_areas = get_area_by_timing(occupancy)

        # Store metadata
        obj_id = getattr(obj, "id", None) or id(obj)
        metadata_entry = {
            "program": program,
            "zone": zone,
            "level": level,
            "occupancy": occupancy,
            "material_color": material_color,
            "area": area,
            "area_off_peak": timing_areas["off_peak"],
            "area_morning": timing_areas["morning"],
            "area_afternoon": timing_areas["afternoon"],
            "area_evening": timing_areas["evening"],
            "speckle_type": getattr(obj, "speckle_type", "Unknown"),
            "area_raw": area_raw,  # DEBUG: track what Area was extracted
            "collection": collection_name,  # Track which collection this element belongs to
        }
        
        coll_element_metadata[obj_id] = metadata_entry
        element_metadata[obj_id] = metadata_entry

        # Track colors by program for reporting
        if material_color and material_color != "Not Found":
            if program not in coll_material_colors:
                coll_material_colors[program] = material_color

        coll_floor_data[level][program] += area
        coll_floor_data_by_occupancy[occupancy][level][program] += area
        coll_zone_data[zone][program]   += area
        
        # Aggregate to global data
        floor_data[level][program] += area
        floor_data_by_occupancy[occupancy][level][program] += area
        zone_data[zone][program]   += area
        material_colors.update(coll_material_colors)
        
        # End: for obj in generic_models (per collection loop)
    
    # Store this collection's data
    collection_data[collection_name] = {
        "floor_data": coll_floor_data,
        "zone_data": coll_zone_data,
        "floor_data_by_occupancy": coll_floor_data_by_occupancy,
        "material_colors": coll_material_colors,
        "element_metadata": coll_element_metadata,
        "model_count": len(list(generic_models)) if generic_models else 0,
    }
    
    # End: for collection_name, generic_models in collections_models.items() (collection loop)

    # ── 5. Compute KPIs, build CSV rows per occupancy, collect issues ───────────────────────
    csv_rows_by_occupancy: dict[str, list] = defaultdict(list)  # {occupancy: [rows]}
    issues            = []
    mono_floor_objects = []   # objects on mono-functional floors (for error pin)
    zone_issue_objects = []   # objects in mismatched zones (for error pin)

    stacking = vertical_stacking_continuity(dict(floor_data))
    
    # DEBUG: Count objects with actual Area values and show distribution
    objs_with_area_param = sum(1 for meta in element_metadata.values() if meta.get("area_raw") is not None)
    objs_with_area = sum(1 for meta in element_metadata.values() if meta.get("area", 0) > 0)
    total_area_sum = sum(meta.get("area", 0) for meta in element_metadata.values())
    
    # Count unique Area values
    area_values = {}
    for meta in element_metadata.values():
        raw = meta.get("area_raw")
        if raw is not None:
            # Extract numeric value from dict if needed
            area_num = raw.get("value") if isinstance(raw, dict) else raw
            if area_num not in area_values:
                area_values[area_num] = 0
            area_values[area_num] += 1
    
    debug_info.append("")
    debug_info.append(f"AGGREGATION SUMMARY: {len(generic_models)} objects → {objs_with_area_param} with Area parameter → {objs_with_area} with valid area")
    debug_info.append(f"Total area across all objects: {total_area_sum:.2f} m²")
    debug_info.append(f"Unique Area values found: {len(area_values)}")
    for area_val, count in sorted(area_values.items()):
        debug_info.append(f"  Area={area_val}: {count} objects")

    # Build CSV rows separately for each occupancy
    for occupancy in sorted(floor_data_by_occupancy.keys()):
        occupancy_floor_data = floor_data_by_occupancy[occupancy]
        
        for level, prog_areas in sorted(occupancy_floor_data.items()):
            total         = sum(prog_areas.values())
            diversity     = shannon_diversity(prog_areas)
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

            for program, area in sorted(prog_areas.items()):
                
                csv_rows_by_occupancy[occupancy].append({
                    "Level":    level,
                    "Program":  program,
                    "Area":     round(area, 2),
                    "Status":   level_status,
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
            elem_meta for elem_meta in element_metadata.values()
            if elem_meta.get("zone") in flagged_zones
        ]

    # ── 7. Attach error pins to objects in Speckle viewer ────────────────────
    # Removed mono-functional floor error pins as requested
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
    
    # Show sample of per-object areas
    objs_with_areas = [(meta["level"], meta["zone"], meta.get("area_raw"), meta["area"]) 
                       for meta in element_metadata.values() if meta.get("area", 0) > 0]
    if objs_with_areas:
        summary_lines.append("")
        summary_lines.append("Sample objects with extracted areas (first 20):")
        for level, zone, area_raw, area in objs_with_areas[:20]:
            # Extract numeric from area_raw if dict
            area_param_val = area_raw.get("value") if isinstance(area_raw, dict) else area_raw if area_raw else "?"
            summary_lines.append(f"  Level={level}, Zone={zone}, Area Parameter={area_param_val}, Extracted Area={area} m²")
    
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

    # ── 10. Add aggregation summary to Excel ──────────────────────────────────
    
    # Calculate summary statistics
    total_area = sum(meta.get('area', 0) for meta in element_metadata.values())
    ok_count = sum(1 for row in csv_rows if row.get("Status") == "OK")
    mono_count = sum(1 for row in csv_rows if "MONO-FUNCTIONAL" in row.get("Status", ""))
    unique_levels = len(set(row.get("Level", "") for row in csv_rows if row.get("Level")))
    unique_programs = len(set(row.get("Program", "") for row in csv_rows if row.get("Program")))
    
    # Add summary separator and data
    csv_rows.append({"Level": "", "Program": "", "Area": "", "Status": ""})
    csv_rows.append({"Level": "SUMMARY", "Program": "AGGREGATION SUMMARY", "Area": "", "Status": ""})
    csv_rows.append({"Level": "Total Area (m2)", "Program": "", "Area": round(total_area, 2), "Status": ""})
    csv_rows.append({"Level": "Total Levels", "Program": "", "Area": unique_levels, "Status": ""})
    csv_rows.append({"Level": "Total Programs", "Program": "", "Area": unique_programs, "Status": ""})
    csv_rows.append({"Level": "OK Entries", "Program": "", "Area": ok_count, "Status": ""})
    csv_rows.append({"Level": "MONO-FUNCTIONAL Entries", "Program": "", "Area": mono_count, "Status": ""})

    # ── 10.5. Add Program Block aggregation (total area per program by occupancy with timing breakdown) ───────────────────────────────────────────
    # Calculate total area for each program AND occupancy combination, plus timing-based areas
    program_occupancy_areas = defaultdict(lambda: defaultdict(float))  # {program: {occupancy: area}}
    
    for meta in element_metadata.values():
        program = meta.get("program", "Unknown")
        occupancy = meta.get("occupancy", "Unknown")
        area = meta.get("area", 0)
        program_occupancy_areas[program][occupancy] += area
    
    # Add program block section to CSV with timing-based columns
    csv_rows.append({"Level": "", "Program": "", "Area": "", "Area_OffPeak": "", "Area_Morning": "", "Area_Afternoon": "", "Area_Evening": ""})
    csv_rows.append({"Level": "OCCUPANCY TIMING BREAKDOWN", "Program": "OCCUPANCY GROUP", "Area": "Total Area", "Area_OffPeak": "Off-Peak (mm)", "Area_Morning": "Morning (mm)", "Area_Afternoon": "Afternoon (mm)", "Area_Evening": "Evening (mm)"})
    
    for occupancy in sorted(set(meta.get("occupancy", "Unknown") for meta in element_metadata.values())):
        # Get timing-based areas for this occupancy
        timing_areas = get_area_by_timing(occupancy)
        
        # Calculate total area for this occupancy across all programs
        total_occupancy_area = sum(
            meta.get("area", 0) 
            for meta in element_metadata.values() 
            if meta.get("occupancy", "Unknown") == occupancy
        )
        
        csv_rows.append({
            "Level": occupancy,
            "Program": "AREAS BY TIMING",
            "Area": round(total_occupancy_area, 2),
            "Area_OffPeak": timing_areas["off_peak"],
            "Area_Morning": timing_areas["morning"],
            "Area_Afternoon": timing_areas["afternoon"],
            "Area_Evening": timing_areas["evening"],
        })

    # ── 11. Export file result based on selected output format ──────────────────────────────────────────
    # Use pre-built occupancy-separated rows
    unique_occupancies = sorted(csv_rows_by_occupancy.keys())
    
    if function_inputs.output_format == OutputFormat.GOOGLE_SHEETS:
        # Export as separate CSV files for each family
        export_format = "CSV (for Google Sheets)"
        csv_files = []
        
        for occupancy in unique_occupancies:
            occupancy_rows = csv_rows_by_occupancy[occupancy]
            
            # If no rows for this occupancy, create header rows at least
            if not occupancy_rows:
                occupancy_rows = [
                    {"Level": occupancy, "Program": "No data", "Area": 0, "Status": ""}
                ]
            
            # Create CSV content
            csv_content_str = rows_to_csv(occupancy_rows)
            
            # Create a properly named CSV file with occupancy in the filename
            safe_occupancy = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in occupancy)
            
            tmp_file = tempfile.NamedTemporaryFile(
                suffix=".csv",
                delete=False,
                prefix=f"program_floor_validation_{safe_occupancy}_",
                mode='w',
                encoding='utf-8',
                dir=None
            )
            tmp_file.write(csv_content_str)
            tmp_file.close()
            csv_files.append({
                'path': tmp_file.name,
                'occupancy': occupancy,
                'content': csv_content_str
            })
            
            # Store each occupancy CSV as a separate result
            try:
                automate_context.store_file_result(tmp_file.name)
            except Exception:
                pass
        
        # Clean up all temporary files (after storing)
        for csv_file_info in csv_files:
            try:
                csv_file_path = csv_file_info['path'] if isinstance(csv_file_info, dict) else csv_file_info
                if os.path.exists(csv_file_path):
                    os.unlink(csv_file_path)
            except Exception:
                pass
    
    else:
        # Export as Excel with multiple sheets (one per family)
        export_format = "Excel"
        excel_sheets = {}
        
        for occupancy in unique_occupancies:
            occupancy_rows = csv_rows_by_occupancy[occupancy]
            
            # If no rows for this occupancy, create header rows at least
            if not occupancy_rows:
                occupancy_rows = [
                    {"Level": occupancy, "Program": "No data", "Area": 0, "Status": ""}
                ]
            
            excel_sheets[occupancy] = occupancy_rows
        
        # Create multiple sheet Excel file
        csv_content = rows_to_excel_multi_sheet(excel_sheets)
        
        # Store the file result
        try:
            automate_context.store_file_result(csv_content)
        finally:
            # Clean up temporary file
            if os.path.exists(csv_content):
                os.unlink(csv_content)

    # Show offending objects in the Speckle viewer
    if issues:
        automate_context.set_context_view()

    # ── 12. Mark run success (always) ──────────────────────
    # Show only basic completion info in Speckle, details exported
    total_area_val = sum(meta.get('area', 0) for meta in element_metadata.values())
    family_count = len(unique_occupancies)
    total_elements = len(element_metadata)
    collection_summary = "\n\n📁 COLLECTIONS PROCESSED:\n" + "\n".join(
        f"  {coll_name}: {coll['model_count']} elements, Area: {sum(meta.get('area', 0) for meta in coll['element_metadata'].values()):.2f} m²"
        for coll_name, coll in collection_data.items()
    )
    
    # Always build Google Sheets links (even if Excel format is selected)
    google_sheets_links = "\n\n📊 GOOGLE SHEETS IMPORT LINKS (per family):\n"
    for unique_family in unique_occupancies:
        gs_link = _generate_google_sheets_import_link(unique_family)
        google_sheets_links += f"  ✓ {unique_family}: {gs_link}\n"
    google_sheets_links += "\nINSTRUCTIONS:\n1. Click the link above to create a new Google Sheet\n2. Go to File → Import → Upload the CSV file from downloads\n3. Select 'Replace spreadsheet' and confirm"
    
    success_msg = (
        f"Program floor analysis complete ({export_format} format with {family_count} families).\n"
        f"Processed {total_elements} elements across {len(floor_data)} levels and {len(zone_data)} zones from {len(collection_data)} collection(s).\n"
        f"Total area: {total_area_val:.2f} m². Exported {family_count} sheet(s)/file(s) for family groups: {', '.join(unique_occupancies)}"
        f"{collection_summary}"
        f"{google_sheets_links}"
    )
    
    automate_context.mark_run_success(success_msg)


# ─────────────────────────────────────────────────────────────────────────────
# Helper
# ─────────────────────────────────────────────────────────────────────────────

def _generate_google_sheets_import_link(occupancy_name: str) -> str:
    """
    Generate a Google Sheets import link for CSV data.
    Instructions: User opens the link to create a new sheet, then imports CSV via File > Import.
    """
    safe_name = "".join(c if c.isalnum() or c in ('-', '_') else '_' for c in occupancy_name)
    sheet_name = f"Program_Floor_{safe_name}"
    
    # Direct link to create new Google Sheet
    create_sheet_url = "https://docs.google.com/spreadsheets/create"
    
    # Return instruction with link
    return f"{create_sheet_url}?title={sheet_name}"


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

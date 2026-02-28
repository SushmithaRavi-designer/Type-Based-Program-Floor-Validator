"""
utils/extractor.py
──────────────────
Helpers for reading parameter values from Speckle / Revit objects
and for traversing the commit object tree.
"""

from typing import Optional


def get_param_value(obj, param_name: str) -> Optional[str]:
    """
    Try every common location a Revit parameter value might live in a
    Speckle object and return it as a string (or None if not found).

    Checks (in order):
      1. Direct attribute on obj (exact and case-insensitive)
      2. obj.type_parameters (Type Parameters in Revit)
      3. obj.parameters as a plain dict {key: value}
      4. obj.parameters as a nested dict {key: {"name": ..., "value": ...}}
      5. Parameter wrapped with units: {"value": 60, "unit": "Meters"}
      6. obj.parameters["Type Parameters"] nested dict
      7. obj.parameters as a Speckle DynamicBase whose child attributes
         are objects with .name / .value properties
      8. Alternative parameter names (replace hyphen with underscore, etc.)
    """
    if obj is None:
        return None

    # 1. Direct attribute (exact match first, then case-insensitive)
    val = getattr(obj, param_name, None)
    if val is not None and not _is_base_like(val):
        return str(val)
    
    # Try case-insensitive attribute matching
    for attr in _safe_dir(obj):
        if attr.lower() == param_name.lower() and not attr.startswith("_"):
            val = getattr(obj, attr, None)
            if val is not None and not _is_base_like(val):
                return str(val)

    # 2. Check type_parameters (Revit Type Parameters)
    type_params = getattr(obj, "type_parameters", None)
    if isinstance(type_params, dict):
        # Exact match
        if param_name in type_params:
            entry = type_params[param_name]
            if isinstance(entry, dict):
                val = entry.get("value")
                return str(val) if val is not None else None
            return str(entry) if entry is not None else None
        
        # Case-insensitive match
        for key, entry in type_params.items():
            if isinstance(key, str) and key.lower() == param_name.lower():
                if isinstance(entry, dict):
                    val = entry.get("value")
                    return str(val) if val is not None else None
                return str(entry) if entry is not None else None

    params = getattr(obj, "parameters", None)
    if params is None:
        return None

    # 3/4/5. Plain dict (exact and case-insensitive)
    if isinstance(params, dict):
        # Exact match first
        if param_name in params:
            entry = params[param_name]
            # Handle wrapped value with unit: {"value": 60, "unit": "Meters"}
            if isinstance(entry, dict):
                val = entry.get("value")
                if val is not None:
                    return str(val)
                # Also try nested name/value structure
                return str(entry.get("value", "")) or None
            return str(entry) if entry is not None else None
        
        # Case-insensitive match
        for key, entry in params.items():
            if isinstance(key, str) and key.lower() == param_name.lower():
                if isinstance(entry, dict):
                    val = entry.get("value")
                    if val is not None:
                        return str(val)
                    return str(entry.get("value", "")) or None
                return str(entry) if entry is not None else None
        
        # 6. Try alternative naming (replace hyphen with underscore)
        alt_param_name = param_name.replace("-", "_")
        if alt_param_name != param_name and alt_param_name in params:
            entry = params[alt_param_name]
            if isinstance(entry, dict):
                val = entry.get("value")
                if val is not None:
                    return str(val)
            return str(entry) if entry is not None else None
        
        # 7. Search by nested name key (original behavior)
        for entry in params.values():
            if isinstance(entry, dict):
                if entry.get("name", "").lower() == param_name.lower():
                    val = entry.get("value")
                    if val is not None:
                        return str(val)
        
        # 8. Check nested "Type Parameters" dict within parameters
        if "Type Parameters" in params:
            type_params_nested = params["Type Parameters"]
            if isinstance(type_params_nested, dict):
                # Exact match
                if param_name in type_params_nested:
                    entry = type_params_nested[param_name]
                    if isinstance(entry, dict):
                        val = entry.get("value")
                        return str(val) if val is not None else None
                    return str(entry) if entry is not None else None
                
                # Case-insensitive match
                for key, entry in type_params_nested.items():
                    if isinstance(key, str) and key.lower() == param_name.lower():
                        if isinstance(entry, dict):
                            val = entry.get("value")
                            return str(val) if val is not None else None
                        return str(entry) if entry is not None else None

    # 7. DynamicBase-style parameters object
    else:
        for key in _safe_dir(params):
            p = getattr(params, key, None)
            if p is None:
                continue
            name = getattr(p, "name", None)
            if name and name.lower() == param_name.lower():
                value = getattr(p, "value", None)
                return str(value) if value is not None else None

    return None



def extract_numeric_value(param_str: str) -> Optional[float]:
    """
    Extract numeric value from parameter string that may include units.
    Examples:
      "60 Meters" → 60.0
      "60" → 60.0
      "60.5" → 60.5
      "invalid" → None
    """
    if not param_str:
        return None
    
    # Remove common unit suffixes and extra whitespace
    cleaned = param_str.strip()
    
    # Try to extract the first numeric value (with optional decimal)
    import re
    match = re.match(r"(-?\d+\.?\d*)", cleaned)
    if match:
        try:
            return float(match.group(1))
        except ValueError:
            return None
    
    return None



def collect_objects(obj, results: list):
    """
    Recursively walk a Speckle object tree and accumulate every
    Base-like object into `results`.
    """
    results.append(obj)
    for key in _safe_dir(obj):
        val = getattr(obj, key, None)
        if _is_base_like(val):
            collect_objects(val, results)
        elif isinstance(val, list):
            for item in val:
                if _is_base_like(item):
                    collect_objects(item, results)


def get_material_color(obj, color_param_name: str = "Material") -> Optional[str]:
    """
    Extract material color information from a Speckle object.

    Attempts to find color from (in priority order):
      1. Material object with RGB/color properties (from Revit)
      2. Render material properties with diffuse/color
      3. Display/graphics override colors
      4. Direct color attributes
      5. Material name parameter

    Returns color as hex string (e.g., "FF0000") or None if not found.
    """
    if obj is None:
        return None

    # Priority 1: Try material object (Revit Material with RGB color)
    material = getattr(obj, "material", None)
    if material:
        # Try direct material color property
        material_color = getattr(material, "color", None)
        if material_color:
            return _normalize_hex_color(material_color)
        
        # Try material.diffuse (render material color)
        diffuse = getattr(material, "diffuse", None)
        if diffuse:
            return _normalize_hex_color(diffuse)

    # Priority 2: Try render material properties (Revit RenderMaterial)
    render_material = getattr(obj, "renderMaterial", None)
    if render_material:
        render_color = getattr(render_material, "diffuse", None)
        if render_color:
            return _normalize_hex_color(render_color)
        
        render_color = getattr(render_material, "color", None)
        if render_color:
            return _normalize_hex_color(render_color)

    # Priority 3: Try display color / graphics override color
    display_color = getattr(obj, "displayColor", None)
    if display_color:
        return _normalize_hex_color(display_color)

    # Priority 4: Try direct color attributes
    for color_attr in ["color", "Color", "materialColor", "fillColor"]:
        color_val = getattr(obj, color_attr, None)
        if color_val:
            return _normalize_hex_color(color_val)

    # Priority 5: Try color parameter directly (case-insensitive)
    for color_name in ["Color", "colour", "Colour", "COLOUR"]:
        color_val = get_param_value(obj, color_name)
        if color_val:
            return _normalize_hex_color(color_val)

    # Priority 6: Try material parameter with various names
    for material_name in [color_param_name, "Material", "Material Color", "MaterialColor", "Finish", "Surface"]:
        material_str = get_param_value(obj, material_name)
        if material_str:
            return _normalize_hex_color(material_str)

    return None


def get_level_info(obj, level_param_name: str = "Level") -> Optional[str]:
    """
    Extract level/floor information from a Speckle object.

    Checks (in order):
      1. Explicit level parameter (e.g., "Level", "Floor")
      2. Alternative names ("Floor", "Story", "Height", "Elevation")
      3. level.name attribute (Revit Level reference)
      4. levelName attribute
      5. Direct Level property
    """
    if obj is None:
        return None

    # Try explicit level parameter (exact and case-insensitive)
    level_val = get_param_value(obj, level_param_name)
    if level_val:
        return level_val.strip()

    # Try alternative level/floor parameter names
    alternative_names = ["Floor", "Story", "Storey", "Height", "Elevation", "Level Name", "LevelName"]
    for alt_name in alternative_names:
        if alt_name.lower() != level_param_name.lower():
            level_val = get_param_value(obj, alt_name)
            if level_val:
                return level_val.strip()

    # Try level object reference
    level = getattr(obj, "level", None)
    if level:
        level_name = getattr(level, "name", None)
        if level_name:
            return level_name.strip()

    # Try levelName direct property
    level_name = getattr(obj, "levelName", None)
    if level_name:
        return level_name.strip()
    
    # Try direct Level property (sometimes Revit stores as numeric/string)
    direct_level = getattr(obj, "Level", None)
    if direct_level:
        return str(direct_level).strip()

    return None


def estimate_area_from_display(obj) -> float:
    """
    Rough bounding-box XY area estimate derived from the mesh vertices
    stored in displayValue when no explicit area parameter is available.
    """
    try:
        display = getattr(obj, "displayValue", None)
        if display is None:
            return 0.0
        if isinstance(display, list):
            if not display:
                return 0.0
            display = display[0]
        vertices = getattr(display, "vertices", [])
        if not vertices or len(vertices) < 9:
            return 0.0
        xs = vertices[0::3]
        ys = vertices[1::3]
        dx = max(xs) - min(xs)
        dy = max(ys) - min(ys)
        return round(dx * dy, 2)
    except Exception:
        return 0.0


# ── Internal helpers ───────────────────────────────────────────────────────────

def _safe_dir(obj) -> list:
    try:
        return [k for k in dir(obj) if not k.startswith("_")]
    except Exception:
        return []


def _is_base_like(val) -> bool:
    """Return True if val looks like a Speckle Base object (not a primitive)."""
    if val is None or isinstance(val, (str, int, float, bool, bytes)):
        return False
    return hasattr(val, "__dict__") or hasattr(val, "id")


def _normalize_hex_color(color_input) -> Optional[str]:
    """
    Normalize various color formats to a 6-digit hex string.

    Handles:
      - Revit Color objects with R, G, B properties
      - Hex strings: "#FF0000", "FF0000", "0xFF0000"
      - RGB tuples/lists: (255, 0, 0), [255, 0, 0]
      - RGB dicts: {"r": 255, "g": 0, "b": 0}
      - Integer values: 16711680
      - Color names (basic mapping)

    Returns: "FF0000" format or None if conversion fails
    """
    if color_input is None:
        return None
    
    # Handle Revit Color object (has R, G, B properties)
    if hasattr(color_input, 'r') and hasattr(color_input, 'g') and hasattr(color_input, 'b'):
        try:
            r = int(getattr(color_input, 'r', 0))
            g = int(getattr(color_input, 'g', 0))
            b = int(getattr(color_input, 'b', 0))
            r = max(0, min(255, r))
            g = max(0, min(255, g))
            b = max(0, min(255, b))
            return f"{r:02X}{g:02X}{b:02X}"
        except (ValueError, TypeError):
            pass

    # Handle string hex colors
    if isinstance(color_input, str):
        color_input = color_input.strip()
        # Remove common prefixes
        if color_input.startswith("#"):
            color_input = color_input[1:]
        if color_input.startswith("0x") or color_input.startswith("0X"):
            color_input = color_input[2:]
        # Pad to 6 digits if needed
        if len(color_input) == 6 and all(c in "0123456789ABCDEFabcdef" for c in color_input):
            return color_input.upper()
        return None

    # Handle RGB tuple/list
    if isinstance(color_input, (list, tuple)) and len(color_input) >= 3:
        try:
            r = max(0, min(255, int(color_input[0])))
            g = max(0, min(255, int(color_input[1])))
            b = max(0, min(255, int(color_input[2])))
            return f"{r:02X}{g:02X}{b:02X}"
        except (ValueError, TypeError):
            return None

    # Handle RGB dict (case-insensitive key matching)
    if isinstance(color_input, dict):
        try:
            # Normalize dict keys to lowercase for matching
            lower_dict = {k.lower() if isinstance(k, str) else k: v 
                         for k, v in color_input.items()}
            r = int(lower_dict.get("r", lower_dict.get("red", 0)))
            g = int(lower_dict.get("g", lower_dict.get("green", 0)))
            b = int(lower_dict.get("b", lower_dict.get("blue", 0)))
            # Clamp to 0-255 range
            r = max(0, min(255, r))
            g = max(0, min(255, g))
            b = max(0, min(255, b))
            return f"{r:02X}{g:02X}{b:02X}"
        except (ValueError, TypeError):
            return None

    # Handle integer color value
    if isinstance(color_input, int):
        try:
            return f"{color_input & 0xFFFFFF:06X}"
        except Exception:
            return None

    return None

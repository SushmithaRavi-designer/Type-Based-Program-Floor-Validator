# Level and Material Color Extraction

## Overview

This enhancement adds comprehensive support for extracting and tracking **building levels** and **material colors** from Speckle Revit models. The system now intelligently recognizes multiple ways this information is stored and organized across the model.

## Features

### 1. **Level Information Extraction**

The `get_level_info()` function extracts building level information from Speckle objects using multiple strategies:

#### Extraction Strategy (in order):
1. **Explicit Level Parameter**: Checks the parameter specified by `level_parameter_name` (default: "Level")
2. **Level Object Reference**: Attempts to read from `obj.level.name`
3. **Direct levelName Property**: Falls back to `obj.levelName`

#### Example Speckle Properties:
```python
# Strategy 1: Parameter-based
{
  "parameters": {
    "Level": "L02"
  }
}

# Strategy 2: Level object reference
{
  "level": {
    "name": "Level 2"
  }
}

# Strategy 3: Direct property
{
  "levelName": "Ground Floor"
}
```

### 2. **Material Color Extraction**

The `get_material_color()` function extracts material color information using multiple strategies:

#### Extraction Strategy (in order):
1. **Direct Color Parameter**: Checks for a "Color" parameter
2. **Material Property**: Looks for `obj.material.color`
3. **Render Material**: Checks `obj.renderMaterial.diffuse` property
4. **Display Color**: Falls back to `obj.displayColor`
5. **Material Parameter**: Uses the parameter specified by `material_color_parameter_name`

#### Supported Color Formats:
- **Hex Strings**: `"FF0000"`, `"#FF0000"`, `"0xFF0000"`
- **RGB Tuples**: `(255, 0, 0)`
- **RGB Lists**: `[255, 0, 0]`
- **RGB Dictionaries**: `{"r": 255, "g": 0, "b": 0}`
- **Integer Values**: `16711680` (RGB as 24-bit int)

#### Normalized Output:
All colors are normalized to 6-digit uppercase hex format: `"FF0000"`

#### Example Speckle Properties:
```python
# Strategy 1: Direct parameter
{
  "parameters": {
    "Color": "0xFF0000"
  }
}

# Strategy 2-3: Material with color
{
  "material": {
    "color": [255, 0, 0]
  }
}

# Strategy 3: Render material
{
  "renderMaterial": {
    "diffuse": "#FF0000"
  }
}

# Strategy 4: Display color
{
  "displayColor": [255, 0, 0]
}
```

## Integration with Main Validation Function

### New Function Inputs

The `FunctionInputs` class now includes:

```python
level_parameter_name: str = Field(
    default="Level",
    title="Level Parameter Name",
    description="Parameter that defines building level",
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
```

### Data Collection

The validation function now tracks:

```python
element_metadata = {
    obj_id: {
        "program": program,
        "zone": zone,
        "level": level,                 # ← NEW: Building level
        "material_color": material_color,  # ← NEW: Material color
        "area": area,
        "speckle_type": speckle_type,
    }
}

material_colors = {
    "Retail": "FF0000",     # ← NEW: Colors mapped by program
    "Office": "0000FF",
}
```

## CSV Export Enhancement

The exported CSV now includes a **Material Color** column:

```
Level,Zone,Program,Material Color,Area,%,Dominant,Diversity Index (H),Vertical Continuity,Status
L02,ZoneA,Retail,FF0000,85.5,45.2,Retail,0.891,0.95,✅ OK
L02,ZoneA,Office,0000FF,94.3,50.1,Office,0.891,0.92,✅ OK
```

## Validation Improvements

### 1. **Level-Based Error Reporting**

Objects are now correctly grouped by their actual building level:
- Validation errors reference specific levels
- Error pins pinpoint objects on problematic levels
- Message formatting uses "Level" instead of "Floor"

### 2. **Material Color Summary**

The version comment now includes a material color summary:

```
── Material Colors by Program ──
  Retail: #FF0000
  Office: #0000FF
  Housing: #00FF00

── Level Summary ──
  L01: total=1200.5 m² | dominant=Retail (65.2%)
  L02: total=1150.3 m² | dominant=Office (72.1%)
```

## Usage Example

### Automate Function Inputs (UI)

```json
{
  "program_source_parameter": "Type Name",
  "zone_parameter_name": "Zone",
  "level_parameter_name": "Level",
  "material_color_parameter_name": "Material",
  "use_material_color": true,
  "area_parameter_name": "Area",
  "program_threshold_matrix": "{\"Retail\": 60, \"Office\": 75}",
  "default_threshold": 70.0
}
```

### Revit Element Properties

For the function to work optimally, ensure Revit elements have:
- A **Level** parameter or are placed on a level in the model
- A **Material** parameter or material assigned in the rendering properties
- An **Area** parameter or Revit geometry data

## Technical Details

### Helper Function: `_normalize_hex_color()`

Converts various color input formats to a standardized 6-digit hex string:

```python
_normalize_hex_color("FF0000")           → "FF0000"
_normalize_hex_color("#FF0000")          → "FF0000"
_normalize_hex_color((255, 0, 0))        → "FF0000"
_normalize_hex_color([255, 0, 0])        → "FF0000"
_normalize_hex_color({"r": 255, ...})    → "FF0000"
_normalize_hex_color(16711680)           → "FF0000"
```

### Error Handling

- **Missing levels**: Falls back through extraction strategies; defaults to "Unknown"
- **Invalid colors**: Returns `None` and doesn't interrupt processing
- **Malformed parameters**: Safely skipped with appropriate defaults
- **Null objects**: Functions return `None` immediately

## Backwards Compatibility

All enhancements are **backwards compatible**:
- Existing models without level/color parameters continue to work
- `use_material_color` can be disabled via settings
- CSV column order preserves existing columns
- Previous validation logic remains unchanged

## Debugging

To inspect extracted metadata, enable logging:

```python
# In element metadata tracking
for obj_id, metadata in element_metadata.items():
    print(f"{obj_id}: Level={metadata['level']}, Color={metadata['material_color']}")
```

## References

- **Speckle Object Model**: https://speckle.dev/
- **Revit Parameters**: Revit API Documentation
- **Color Standards**: CSS/Web color hex format (RRGGBB)

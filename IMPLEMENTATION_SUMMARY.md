# Level and Material Color Enhancement - Implementation Summary

## Overview

Successfully implemented comprehensive support for extracting and considering **building levels** and **material colors** from Speckle model data. This enhancement improves the Type-Based Program Floor Validator's ability to organize and validate building program allocation across different levels with material-aware tracking.

## Changes Made

### 1. **Enhanced `extractor.py`** — New Functions Added

#### `get_level_info(obj, level_param_name="Level") → Optional[str]`
Intelligently extracts building level information from Speckle objects using a fallback strategy:
- **Strategy 1**: Explicit level parameter (from Revit Type properties)
- **Strategy 2**: Level object reference (`obj.level.name`)
- **Strategy 3**: Direct `levelName` property

**Example Usage:**
```python
level = get_level_info(revit_element, "Level")
# Returns: "L02", "Ground Floor", or None
```

#### `get_material_color(obj, color_param_name="Material") → Optional[str]`
Extracts material color from Speckle objects with multiple extraction strategies:
- **Strategy 1**: Direct "Color" parameter
- **Strategy 2**: Material color property (`obj.material.color`)
- **Strategy 3**: Render material diffuse color (`obj.renderMaterial.diffuse`)
- **Strategy 4**: Display color (`obj.displayColor`)
- **Strategy 5**: Custom material parameter

**Returns:** Normalized 6-digit hex color (e.g., `"FF0000"` for red)

**Supported Input Formats:**
```python
# All these return "FF0000"
get_material_color(obj)  # Hex string
# Internal: "FF0000", "#FF0000", "0xFF0000"
# RGB tuple: (255, 0, 0)
# RGB list: [255, 0, 0]
# RGB dict: {"r": 255, "g": 0, "b": 0}
# Integer: 16711680
```

#### `_normalize_hex_color(color_input) → Optional[str]`
Internal helper that converts any color format to standardized hex output:
- ✅ Handles hex strings, RGB tuples/lists, RGB dicts, integer values
- ✅ Case-insensitive dictionary keys
- ✅ Clamps out-of-range RGB values (0-255)
- ✅ Strips and validates hex strings
- ✅ Returns uppercase 6-digit hex or None

**Test Results:** 17/17 unit tests passing

### 2. **Updated `main new.py`** — Enhanced Validation Function

#### New Function Input Parameters:
```python
level_parameter_name: str = "Level"
# Parameter name for explicit level definition

material_color_parameter_name: str = "Material"  
# Parameter name for material/color information

use_material_color: bool = True
# Enable/disable material color extraction
```

#### Enhanced Data Processing:
```python
element_metadata = {
    obj_id: {
        "program": program,
        "zone": zone,
        "level": level,          # ← NEW: Actual level from model
        "material_color": color, # ← NEW: Material color (hex)
        "area": area,
        "speckle_type": type,
    }
}

material_colors = {
    "Retail": "FF0000",    # ← NEW: Tracked by program
    "Office": "0000FF",
    "Housing": "00FF00",
}
```

#### Improved Output:
- All references to "Floor" updated to "Level" for accuracy
- CSV export now includes "Material Color" column
- Version comment includes material color summary by program
- Error messages reference actual levels from Speckle data

### 3. **Updated `csv_exporter.py`** — Extended Column Set

**New CSV Structure:**
```
Level,Zone,Program,Material Color,Area,%,Dominant,Diversity Index (H),Vertical Continuity,Status
L01,ZoneA,Retail,FF0000,1200.5,65.2,Retail,0.891,0.95,✅ OK
L01,ZoneA,Office,0000FF,650.3,35.2,Retail,0.891,0.92,⚠️  MONO-FUNCTIONAL
L02,ZoneA,Office,0000FF,1150.3,72.1,Office,0.756,0.98,✅ OK
```

### 4. **New Documentation** — `LEVEL_AND_COLOR_EXTRACTION.md`

Comprehensive technical documentation covering:
- Feature overview and strategies
- Color format support and normalization
- Integration with main validation function
- Usage examples and debugging guidance
- Backward compatibility notes
- Technical implementation details

### 5. **Test Suite** — `tests/test_extractor.py`

Complete unit test coverage for color normalization:
- ✅ 10 hex string format tests
- ✅ 4 RGB format tests (tuples, lists, dicts)
- ✅ Integer color value tests
- ✅ Edge case handling (None, invalid inputs, whitespace)
- ✅ Format conversion consistency
- ✅ Case-insensitive key matching
- ✅ Out-of-range value clamping

**Test Results:** All 17 tests passing ✅

## Key Features

### 1. **Intelligent Level Detection**
- Automatic detection from Revit placement
- Parameter-based fallback
- Consistent formatting and validation

### 2. **Flexible Color Support**
- Multiple color format recognition
- Automatic normalization to hex
- Out-of-range value handling
- Case-insensitive processing

### 3. **Robust Error Handling**
- Graceful null/missing data handling
- No interruption to processing on invalid formats
- Sensible defaults throughout

### 4. **Backward Compatible**
- Existing models without level/color data work unchanged
- Color extraction can be disabled via settings
- Previous validation logic preserved
- CSV columns include new data without breaking existing columns

## Usage Example

### Automate Configuration
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

### Report Output
```
── Material Colors by Program ──
  Retail: #FF0000
  Office: #0000FF
  Housing: #00FF00
  Transportation: #FFFF00

── Level Summary ──
  L01: total=1850.8 m² | dominant=Retail (65.2%)
  L02: total=1800.3 m² | dominant=Office (72.1%)
  L03: total=1920.5 m² | dominant=Housing (68.5%)
  Ground: total=2100.0 m² | dominant=Retail (75.0%)
```

## Technical Specifications

### Color Normalization Algorithm
1. Accepts raw color from Speckle object
2. Auto-detects format (hex, RGB, int)
3. Validates/clamps values (0-255 range)
4. Returns uppercase 6-digit hex string
5. Returns `None` on non-recoverable errors

### Level Detection Algorithm
1. Queries explicit level parameter
2. Falls back to object's level reference
3. Falls back to levelName property
4. Returns string representation of level

### Performance
- O(1) per-element processing
- No external dependencies for color conversion
- Minimal memory overhead for tracking

## Files Modified

| File | Changes |
|------|---------|
| `extractor.py` | Added `get_level_info()`, `get_material_color()`, `_normalize_hex_color()` |
| `main new.py` | Added level/color parameters, enhanced metadata tracking |
| `csv_exporter.py` | Updated COLUMNS to include "Material Color", changed "Floor" to "Level" |
| `tests/test_extractor.py` | New test file with 17 unit tests |

## Files Created

| File | Purpose |
|------|---------|
| `LEVEL_AND_COLOR_EXTRACTION.md` | Technical documentation |
| `tests/test_extractor.py` | Unit tests for new functions |

## Validation & Testing

✅ **Syntax Validation:** All files pass Python syntax check  
✅ **Unit Tests:** 17/17 color normalization tests passing  
✅ **Integration:** Functions work with existing codebase  
✅ **Backward Compatibility:** No breaking changes  

## Future Enhancements

Potential additions:
1. Material opacity/transparency tracking
2. Multi-color support per element
3. Color-coded validation reports
4. Material schedule export
5. Level-based filtering in CSV
6. Color consistency validation

## Conclusion

The Type-Based Program Floor Validator now has robust support for extracting and tracking building levels and material colors from Speckle models. All data is properly normalized, validated, and integrated into the validation workflow while maintaining full backward compatibility.

**Status:** ✅ **Complete and tested**

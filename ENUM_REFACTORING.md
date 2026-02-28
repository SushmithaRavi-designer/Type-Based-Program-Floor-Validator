# Function Inputs Refactoring — Using Enums (Best Practices)

## Overview

The Type-Based Program Floor Validator has been refactored to use **Python Enums** for function input parameters, following Speckle best practices. This makes the function more maintainable, provides better UI support in Speckle Automate, and reduces the potential for user errors.

## Why Enums?

### Before (String-Based)
```python
use_material_color: bool = Field(
    default=True,
    title="Use Material Color",
    description="Whether to extract material colors",
)
```

**Problems:**
- ❌ Boolean is unclear (what do True/False mean?)
- ❌ No dropdown in UI
- ❌ No type safety
- ❌ Easy for users to make mistakes

### After (Enum-Based)
```python
color_extraction_mode: ColorExtractionMode = Field(
    default=ColorExtractionMode.ENABLED,
    title="Color Extraction Mode",
    description="Enable or disable material color extraction",
)
```

**Benefits:**
- ✅ Clear, self-documenting options
- ✅ Dropdown menu in Speckle UI
- ✅ Type-safe values
- ✅ No invalid values possible
- ✅ Better IDE autocomplete

## Refactored Enums

### 1. ColorExtractionMode
Controls whether material colors are extracted from Speckle objects.

```python
class ColorExtractionMode(str, Enum):
    ENABLED = "enabled"      # Extract colors from all objects
    DISABLED = "disabled"     # Skip color extraction
```

**Usage in UI:**
```
┌─ Color Extraction Mode ────────────────┐
│ ▼ [Enabled                          ] │  ← Dropdown selector
│   □ Enabled                         │
│   □ Disabled                        │
└────────────────────────────────────────┘
```

**In Code:**
```python
if function_inputs.color_extraction_mode == ColorExtractionMode.ENABLED:
    material_color = get_material_color(obj, parameter_name)
```

### 2. ThresholdMode
Controls how validation thresholds are applied to program allocation checks.

```python
class ThresholdMode(str, Enum):
    STRICT = "strict"          # -10% from configured thresholds
    PERMISSIVE = "permissive"  # +10% from configured thresholds
    CUSTOM = "custom"          # Use thresholds as configured
```

**Behavior:**
| Mode | Example | Effect |
|------|---------|--------|
| STRICT | 60% threshold → 50% | Stricter validation |
| PERMISSIVE | 60% threshold → 70% | Relaxed validation |
| CUSTOM | Use "Program Threshold Matrix" value | As-is |

**Example:**
```python
if function_inputs.threshold_mode == ThresholdMode.STRICT:
    # Reduce all thresholds by 10%
    thresholds = {prog: max(10, thresh - 10) for prog, thresh in thresholds.items()}
elif function_inputs.threshold_mode == ThresholdMode.PERMISSIVE:
    # Increase all thresholds by 10%
    thresholds = {prog: min(95, thresh + 10) for prog, thresh in thresholds.items()}
else:  # ThresholdMode.CUSTOM
    # Use values exactly as provided in the matrix
    pass
```

### 3. ReportLevel
Controls the detail level of the version comment and analysis report.

```python
class ReportLevel(str, Enum):
    SUMMARY = "summary"      # Basic overview only
    DETAILED = "detailed"    # Full analysis with colors and levels
    VERBOSE = "verbose"      # All details + configuration metadata
```

**Report Content by Level:**

**SUMMARY:**
```
✅ Analysed 150 Generic Model elements across 5 levels and 3 zones.

⚠️ Issues detected:
  • Level L01 exceeds mono-functional threshold (Retail = 65.2%, limit = 60%).

✅ All other levels pass program allocation thresholds.
```

**DETAILED:** (includes color and level summaries)
```
✅ Analysed 150 Generic Model elements across 5 levels and 3 zones.

── Material Colors by Program ──
  Retail: #FF0000
  Office: #0000FF
  Housing: #00FF00

⚠️ Issues detected:
  • Level L01 exceeds mono-functional threshold...

── Level Summary ──
  L01: total=1200.5 m² | dominant=Retail (65.2%)
  L02: total=950.3 m² | dominant=Office (72.1%)

── Zone Summary ──
  ZoneA: [Retail, Office]
  ZoneB: [Office, Housing]
```

**VERBOSE:** (includes everything + configuration)
```
[All details from DETAILED]

── Configuration Used ──
  Threshold Mode: strict
  Color Extraction: enabled
  Default Threshold: 70.0%
  Report Level: verbose
```

**In Code:**
```python
if function_inputs.report_level in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
    summary_lines.append("── Level Summary ──")
    # Add level summary details

if function_inputs.report_level == ReportLevel.VERBOSE:
    summary_lines.append("── Configuration Used ──")
    # Add configuration details
```

## FunctionInputs Class Structure

### Parameter Organization

```python
class FunctionInputs(AutomateBase):
    # ── Parameter Configuration (flexible strings) ──
    program_source_parameter: str        # "Type Name", "Category", etc.
    zone_parameter_name: str             # "Zone", "ZoneID", etc.
    level_parameter_name: str            # "Level", "Floor", etc.
    area_parameter_name: str             # "Area", "GrossArea", etc.
    material_color_parameter_name: str   # "Material", "Color", etc.

    # ── Fixed Options (Enums) ──
    color_extraction_mode: ColorExtractionMode      # ENABLED/DISABLED
    threshold_mode: ThresholdMode                   # STRICT/PERMISSIVE/CUSTOM
    report_level: ReportLevel                       # SUMMARY/DETAILED/VERBOSE

    # ── Configuration Values (numeric/matrix) ──
    program_threshold_matrix: str        # JSON string
    default_threshold: float             # 0-100, ge=0, le=100
```

## Implementation Details

### Enum Value Handling

All enums inherit from `(str, Enum)` to ensure:
- ✅ String serialization for JSON
- ✅ Easy comparison: `mode == ColorExtractionMode.ENABLED`
- ✅ String conversion: `str(mode)` or `mode.value`

### Threshold Mode Adjustment Logic

```python
# Parse base thresholds
thresholds = json.loads(function_inputs.program_threshold_matrix)

# Apply mode adjustments
if function_inputs.threshold_mode == ThresholdMode.STRICT:
    thresholds = {prog: max(10, thresh - 10) for prog, thresh in thresholds.items()}
    default_threshold = max(10, function_inputs.default_threshold - 10)
elif function_inputs.threshold_mode == ThresholdMode.PERMISSIVE:
    thresholds = {prog: min(95, thresh + 10) for prog, thresh in thresholds.items()}
    default_threshold = min(95, function_inputs.default_threshold + 10)
else:  # ThresholdMode.CUSTOM
    default_threshold = function_inputs.default_threshold
```

**Safety Constraints:**
- STRICT mode: Minimum 10%, maximum 90%
- PERMISSIVE mode: Minimum 10%, maximum 95%
- CUSTOM: Uses configured value (0-100)

### Report Level Filtering

```python
# Material colors shown only in DETAILED or VERBOSE
if function_inputs.report_level in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
    # Include color summary

# Configuration shown only in VERBOSE
if function_inputs.report_level == ReportLevel.VERBOSE:
    # Include configuration details
```

## UI Appearance

In Speckle Automate, the refactored inputs appear as:

```
Program Source Parameter
[Type Name            ]  ← Text input (flexible)

Zone Parameter Name
[Zone                 ]  ← Text input (flexible)

Level Parameter Name
[Level                ]  ← Text input (flexible)

Color Extraction Mode
▼ [Enabled          ]  ← Dropdown (Enum)
  Enabled
  Disabled

Threshold Mode
▼ [Custom           ]  ← Dropdown (Enum)
  Strict
  Permissive
  Custom

Program Threshold Matrix
[{"Retail": 60...}    ]  ← JSON text input

Default Threshold
[70.0                 ]  ← Number input (0-100)

Report Level
▼ [Detailed         ]  ← Dropdown (Enum)
  Summary
  Detailed
  Verbose
```

## Backward Compatibility

**Migration Notes:**
- Old `use_material_color: bool` → New `color_extraction_mode: ColorExtractionMode`
  - `True` → `ColorExtractionMode.ENABLED`
  - `False` → `ColorExtractionMode.DISABLED`
- No changes to parameter name inputs (still flexible strings)
- No changes to threshold matrix format (still JSON string)

## Best Practices

✅ **Use Enums for:**
- Fixed set of predefined options
- Options that won't change frequently
- Values that affect behavior (on/off, mode selection)
- UI dropdown selections

✅ **Use Strings for:**
- Dynamic parameter names
- Open-ended configuration
- JSON/complex data structures

✅ **Use Numbers for:**
- Quantitative values (percentages, thresholds)
- Measurements with specific ranges

## Testing

The enum refactoring has been verified with:
- ✅ Syntax validation (AST parsing)
- ✅ Enum comparison operations
- ✅ Enum membership checks
- ✅ Threshold adjustment logic
- ✅ Report level filtering

Run tests:
```bash
python test_enums.py
```

## Examples

### Example 1: Basic Configuration
```json
{
  "color_extraction_mode": "enabled",
  "threshold_mode": "custom",
  "report_level": "detailed",
  "program_threshold_matrix": "{\"Retail\": 60, \"Office\": 75}",
  "default_threshold": 70.0
}
```

### Example 2: Strict Validation
```json
{
  "color_extraction_mode": "enabled",
  "threshold_mode": "strict",
  "report_level": "verbose",
  "program_threshold_matrix": "{\"Retail\": 60, \"Office\": 75, \"Housing\": 80}",
  "default_threshold": 75.0
}
```

### Example 3: Permissive with Summary Report
```json
{
  "color_extraction_mode": "disabled",
  "threshold_mode": "permissive",
  "report_level": "summary",
  "program_threshold_matrix": "{\"Retail\": 50}",
  "default_threshold": 60.0
}
```

## References

- [Speckle Function Inputs Documentation](https://speckle.dev/)
- [Python Enum Best Practices](https://docs.python.org/3/library/enum.html)
- [Pydantic Field Configuration](https://docs.pydantic.dev/latest/usage/models/)

---

**Status:** ✅ Implemented and tested  
**Date:** February 28, 2026

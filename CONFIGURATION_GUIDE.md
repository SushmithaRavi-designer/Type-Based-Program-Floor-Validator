# Speckle Model Analysis - Configuration Guide

## Function Input Configuration

When you run this Speckle Automate function, configure these parameters:

### Program Source Parameter
- **Default**: `Type Name`
- **Purpose**: Where to find program information (e.g., "Retail", "Office", "Housing")
- **Examples**: 
  - "Type Name" - from Revit family type names
  - "Program" - if you have a program parameter
  - "Use Type" - alternative parameter name
  
### Level Parameter Name
- **Default**: `Level`
- **Purpose**: Extract building level/floor information
- **Auto-Fallback**: If parameter not found, tries:
  - "Floor", "Story", "Storey", "Height", "Elevation"
  - Revit Level object references
  - Direct `level` property on geometry elements

### Zone Parameter Name
- **Default**: `Zone`
- **Purpose**: Extract spatial zone classification
- **Examples**: "Zone A", "Zone B", or regional identifiers
  
### Area Parameter Name
- **Default**: `Area`
- **Purpose**: Get element floor area
- **Auto-Fallback**: If missing, estimates from 3D bounding box geometry

### Material Color Parameter Name
- **Default**: `Material`
- **Purpose**: Extract material/color information
- **Auto-Fallback**: If parameter not found, searches for:
  - "Color", "Colour", "Material Color"
  - Direct color properties
  - Revit render material colors
  - Display colors (graphics overrides)

### Color Extraction Mode
- **Options**: 
  - `enabled` (default) - extract and show colors
  - `disabled` - skip color extraction
  
### Report Level
- **Options**:
  - `summary` - basic statistics only
  - `detailed` (default) - full analysis with colors and levels
  - `verbose` - include all metadata and processing details

### Threshold Mode
- **Options**:
  - `strict` - reduce thresholds by 10% (stricter checking)
  - `permissive` - increase thresholds by 10% (lenient checking)
  - `custom` (default) - use defined threshold matrix

### Threshold Configuration
- **Program Threshold Matrix** (JSON):
  ```json
  {
    "Retail": 60,
    "Office": 75,
    "Housing": 80,
    "Exhibition": 65
  }
  ```
  - Defines maximum % of dominant program allowed per floor
  - Only used when threshold mode is "custom"
  - Programs not in matrix use default threshold

- **Default Threshold**: `70.0` (70%)
  - Fallback for programs not in matrix
  - Range: 0-100%

## Expected CSV Output Columns

```
Level | Zone | Program | Material Color | Area | % | Dominant | Diversity Index (H) | Vertical Continuity | Status
------|------|---------|----------------|------|---|----------|-------------------|-------------------|--------
```

### Column Meanings

1. **Level** - Building floor/level where element is located
2. **Zone** - Spatial zone assignment
3. **Program** - Program type (use, occupation)
4. **Material Color** - Extracted color in hex format (e.g., "FF0000")
5. **Area** - Floor area in square meters
6. **%** - Percentage of total area on that level
7. **Dominant** - Program type with highest percentage
8. **Diversity Index (H)** - Shannon entropy (0=mono-functional, higher=more diversity)
9. **Vertical Continuity** - How consistently a program appears across floors
10. **Status** - Validation result (✅ OK or ⚠️ MONO-FUNCTIONAL)

## Troubleshooting: Not Seeing Your Data

### If all values show "Unknown":

1. **Check parameter names match your Revit model**:
   - Open your Revit model
   - Check Element Properties to see actual parameter names
   - Update configuration with exact parameter names

2. **Verify Generic Model elements are being sent**:
   - Ensure your commit includes Revit "Generic Model" family instances
   - Other element types (walls, doors, etc.) won't be processed

3. **Check parameter locations**:
   - Parameters might be on family type vs. instance
   - Function automatically searches both locations

4. **Use verbose report level for debugging**:
   - Set Report Level to "verbose"
   - Shows configuration used and processing details

### If Material Color shows "Not Found":

1. **Color might not be in standard location**:
   - Check Revit element properties for actual color parameter
   - May be named differently than "Material" or "Color"
   - Update Material Color Parameter Name field

2. **Try disabling color extraction first**:
   - Set Color Extraction Mode to "disabled"
   - This skips color detection to isolate the issue

3. **Verify Revit color is set**:
   - Generic Models must have color/material assignment
   - Function extracts from: parameter, material object, or graphics override

## Example Successful Configuration

For a typical Revit model with:
- Generic Models representing floor-area programs
- Type Names formatted as "Office-A" or "Retail-GF"
- Level parameters named "Level"
- No separate Zone parameter

**Use these settings**:
```
Program Source Parameter: Type Name
Level Parameter Name: Level
Zone Parameter Name: Zone
Area Parameter Name: Area
Material Color Parameter Name: Material
Color Extraction Mode: enabled
Report Level: detailed
Threshold Mode: custom
Program Threshold Matrix: {"Office": 75, "Retail": 60, "Housing": 80}
Default Threshold: 70
```

## Output Files

The function creates one file:
- **`program_floor_validation_*.csv`** - Validation report with all analysis results
  - Timestamped filename
  - Excel-compatible format
  - UTF-8 encoding for international characters

## What Gets Flagged as Issues

1. **Mono-Functional Floor**: Single program exceeds threshold %
   - Example: 80% Office + 20% Retail, threshold 75% → ISSUE
   - Elements flagged in Speckle viewer with red pins

2. **Zone Incompatibility**: Program doesn't match expected zones
   - Cross-checks program/zone compatibility
   - Custom thresholds can be defined per program

3. **Missing Data**: Elements without required parameters
   - Shown as "Unknown" in report
   - May indicate configuration issue or model gaps

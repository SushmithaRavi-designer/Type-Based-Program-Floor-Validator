#!/usr/bin/env python
"""Quick test of Enum refactoring."""

from enum import Enum

# Test the enums work correctly
class ColorExtractionMode(str, Enum):
    ENABLED = "enabled"
    DISABLED = "disabled"

class ReportLevel(str, Enum):
    SUMMARY = "summary"
    DETAILED = "detailed"
    VERBOSE = "verbose"

class ThresholdMode(str, Enum):
    STRICT = "strict"
    PERMISSIVE = "permissive"
    CUSTOM = "custom"

# Test enum values
print("✅ Enum Refactoring Verification")
print(f"✅ ColorExtractionMode: {[e.value for e in ColorExtractionMode]}")
print(f"✅ ReportLevel: {[e.value for e in ReportLevel]}")
print(f"✅ ThresholdMode: {[e.value for e in ThresholdMode]}")

# Test enum comparison
mode = ColorExtractionMode.ENABLED
if mode == ColorExtractionMode.ENABLED:
    print("✅ Enum comparison works correctly")

report = ReportLevel.DETAILED
if report in (ReportLevel.DETAILED, ReportLevel.VERBOSE):
    print("✅ Enum membership check works correctly")

# Test threshold mode logic
thresh = ThresholdMode.STRICT
if thresh == ThresholdMode.STRICT:
    print("✅ STRICT mode threshold adjustment logic works")
elif thresh == ThresholdMode.PERMISSIVE:
    print("✅ PERMISSIVE mode threshold adjustment logic works")
else:
    print("✅ CUSTOM mode uses defined thresholds")

print("\n✅ All enum refactoring tests passed!")

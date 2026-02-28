#!/usr/bin/env python3
"""Test the enhanced extraction features."""

from extractor import _normalize_hex_color
from parser import parse_type_name

# Test 1: Zone extraction from Type Name
print("=" * 60)
print("TEST 1: Zone Extraction from Type Name")
print("=" * 60)

test_cases = [
    ("MEDICAL_ZoneA_LEVEL10", ("MEDICAL", "ZoneA", "LEVEL10")),
    ("TRANS HQ_ZoneB_LEVEL11", ("TRANS HQ", "ZoneB", "LEVEL11")),
    ("OFFICE_L2", ("OFFICE", "Unknown", "L2")),
    ("Housing", ("Housing", "Unknown", "Unknown")),
    ("", ("Unknown", "Unknown", "Unknown")),
]

for input_type, expected in test_cases:
    result = parse_type_name(input_type)
    status = "✓" if result == expected else "✗"
    print(f"{status} parse_type_name('{input_type}')")
    print(f"   Expected: {expected}")
    print(f"   Got:      {result}")
    if result != expected:
        print(f"   FAILED!")
    print()

# Test 2: Color normalization with Revit Color objects
print("=" * 60)
print("TEST 2: Color Normalization (Revit Format)")
print("=" * 60)

class MockRevitColor:
    """Mock Revit Color object with R, G, B properties."""
    def __init__(self, r, g, b):
        self.r = r
        self.g = g
        self.b = b

color_tests = [
    (MockRevitColor(123, 255, 255), "7BFFFF", "RGB(123, 255, 255) - cyan"),
    (MockRevitColor(255, 0, 0), "FF0000", "RGB(255, 0, 0) - red"),
    (MockRevitColor(120, 120, 120), "787878", "RGB(120, 120, 120) - gray"),
    ("#FF0000", "FF0000", "Hex string #FF0000"),
    ("FF0000", "FF0000", "Hex string FF0000"),
    ((123, 255, 255), "7BFFFF", "RGB tuple (123, 255, 255)"),
    ([123, 255, 255], "7BFFFF", "RGB list [123, 255, 255]"),
]

for color_input, expected, desc in color_tests:
    result = _normalize_hex_color(color_input)
    status = "✓" if result == expected else "✗"
    print(f"{status} {desc}")
    print(f"   Expected: {expected}")
    print(f"   Got:      {result}")
    if result != expected:
        print(f"   FAILED!")
    print()

print("=" * 60)
print("All enhancement tests completed!")
print("=" * 60)

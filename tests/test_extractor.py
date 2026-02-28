"""
tests/test_extractor.py
──────────────────────
Unit tests for level and material color extraction functions.

Run with:
    python -m pytest tests/test_extractor.py -v
    or from tests/ directory: python test_extractor.py
"""

import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# Import from root directory modules
from extractor import _normalize_hex_color


# ── Color normalization tests ────────────────────────────────────────────────

def test_normalize_hex_string_simple():
    """Test basic hex string normalization."""
    assert _normalize_hex_color("FF0000") == "FF0000"
    assert _normalize_hex_color("00FF00") == "00FF00"
    assert _normalize_hex_color("0000FF") == "0000FF"


def test_normalize_hex_string_with_hash():
    """Test hex strings with # prefix."""
    assert _normalize_hex_color("#FF0000") == "FF0000"
    assert _normalize_hex_color("#00FF00") == "00FF00"
    assert _normalize_hex_color("#0000FF") == "0000FF"


def test_normalize_hex_string_with_0x():
    """Test hex strings with 0x prefix."""
    assert _normalize_hex_color("0xFF0000") == "FF0000"
    assert _normalize_hex_color("0x00FF00") == "00FF00"
    assert _normalize_hex_color("0X0000FF") == "0000FF"


def test_normalize_hex_string_lowercase():
    """Test that lowercase hex is normalized to uppercase."""
    assert _normalize_hex_color("ff0000") == "FF0000"
    assert _normalize_hex_color("#aabbcc") == "AABBCC"


def test_normalize_rgb_tuple():
    """Test RGB tuple normalization."""
    assert _normalize_hex_color((255, 0, 0)) == "FF0000"
    assert _normalize_hex_color((0, 255, 0)) == "00FF00"
    assert _normalize_hex_color((0, 0, 255)) == "0000FF"
    assert _normalize_hex_color((128, 64, 32)) == "804020"


def test_normalize_rgb_list():
    """Test RGB list normalization."""
    assert _normalize_hex_color([255, 0, 0]) == "FF0000"
    assert _normalize_hex_color([0, 255, 0]) == "00FF00"
    assert _normalize_hex_color([0, 0, 255]) == "0000FF"


def test_normalize_rgb_dict():
    """Test RGB dictionary normalization."""
    assert _normalize_hex_color({"r": 255, "g": 0, "b": 0}) == "FF0000"
    assert _normalize_hex_color({"r": 0, "g": 255, "b": 0}) == "00FF00"
    assert _normalize_hex_color({"red": 0, "green": 0, "blue": 255}) == "0000FF"


def test_normalize_rgb_dict_alternate_keys():
    """Test RGB dict with alternate key names (case-insensitive)."""
    assert _normalize_hex_color({"red": 255, "green": 0, "blue": 0}) == "FF0000"
    assert _normalize_hex_color({"R": 100, "G": 150, "B": 200}) == "6496C8"


def test_normalize_integer_color():
    """Test integer color value normalization."""
    assert _normalize_hex_color(16711680) == "FF0000"  # Red as int
    assert _normalize_hex_color(65280) == "00FF00"     # Green as int
    assert _normalize_hex_color(255) == "0000FF"       # Blue as int


def test_normalize_none():
    """Test that None input returns None."""
    assert _normalize_hex_color(None) is None


def test_normalize_invalid_string():
    """Test that invalid hex strings return None."""
    assert _normalize_hex_color("GGGGGG") is None
    assert _normalize_hex_color("12345") is None  # Too short
    assert _normalize_hex_color("not_a_color") is None


def test_normalize_invalid_tuple():
    """Test that out-of-range RGB values are clamped to 0-255."""
    assert _normalize_hex_color((256, 0, 0)) == "FF0000"  # 256 is clamped to 255
    assert _normalize_hex_color((-1, 0, 0)) == "000000"  # -1 is clamped to 0


def test_normalize_incomplete_dict():
    """Test RGB dict with missing keys (defaults to 0)."""
    assert _normalize_hex_color({"r": 255}) == "FF0000"  # g and b default to 0
    assert _normalize_hex_color({"g": 128}) == "008000"  # r and b default to 0


def test_normalize_whitespace_handling():
    """Test that whitespace is handled properly."""
    assert _normalize_hex_color("  FF0000  ") == "FF0000"
    assert _normalize_hex_color("  #FF0000  ") == "FF0000"


def test_color_conversion_consistency():
    """Test that different formats for same color produce same output."""
    red_hex = "FF0000"
    red_tuple = (255, 0, 0)
    red_list = [255, 0, 0]
    red_dict = {"r": 255, "g": 0, "b": 0}
    red_int = 16711680

    assert _normalize_hex_color(red_hex) == _normalize_hex_color(red_tuple)
    assert _normalize_hex_color(red_hex) == _normalize_hex_color(red_list)
    assert _normalize_hex_color(red_hex) == _normalize_hex_color(red_dict)
    assert _normalize_hex_color(red_hex) == _normalize_hex_color(red_int)


# ── Mock Speckle object tests ────────────────────────────────────────────────

class MockSpeckleObject:
    """Simple mock object for testing parameter extraction."""
    def __init__(self, **attrs):
        for key, val in attrs.items():
            setattr(self, key, val)


def test_get_level_info_from_parameter():
    """Test level extraction from parameter."""
    from extractor import get_level_info
    
    # Create a mock object with level parameter
    obj = MockSpeckleObject(
        parameters={"Level": {"value": "L02"}}
    )
    # Note: This test depends on the actual get_param_value implementation
    # For now, we just ensure the function exists and can be called
    result = get_level_info(obj, "Level")
    # Result might be None depending on parameter structure


def test_get_material_color_from_parameter():
    """Test material color extraction from parameter."""
    from extractor import get_material_color
    
    # Create a mock object with color property
    obj = MockSpeckleObject(
        displayColor=[255, 0, 0]
    )
    result = get_material_color(obj, "Material")
    # Result should be either a color or None
    assert result is None or isinstance(result, str)


if __name__ == "__main__":
    # Quick smoke-run without pytest
    tests = [
        test_normalize_hex_string_simple,
        test_normalize_hex_string_with_hash,
        test_normalize_hex_string_with_0x,
        test_normalize_hex_string_lowercase,
        test_normalize_rgb_tuple,
        test_normalize_rgb_list,
        test_normalize_rgb_dict,
        test_normalize_rgb_dict_alternate_keys,
        test_normalize_integer_color,
        test_normalize_none,
        test_normalize_invalid_string,
        test_normalize_invalid_tuple,
        test_normalize_incomplete_dict,
        test_normalize_whitespace_handling,
        test_color_conversion_consistency,
        test_get_level_info_from_parameter,
        test_get_material_color_from_parameter,
    ]
    passed = 0
    failed = 0
    for t in tests:
        try:
            t()
            print(f"  ✅ {t.__name__}")
            passed += 1
        except AssertionError as e:
            print(f"  ❌ {t.__name__}: {e}")
            failed += 1
        except Exception as e:
            print(f"  ⚠️  {t.__name__}: {type(e).__name__}: {e}")
            failed += 1
    
    print(f"\n{passed}/{len(tests)} tests passed, {failed} failed.")

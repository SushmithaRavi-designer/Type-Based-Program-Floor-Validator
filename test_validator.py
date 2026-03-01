"""
tests/test_validator.py
───────────────────────
Unit tests for 03_TypeBasedProgramFloorValidator.

Run with:
    python -m pytest tests/ -v
"""

import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from parser import parse_type_name, normalize_floor_label
from kpi import (
    shannon_diversity,
    mono_functional_check,
    check_zone_compatibility,
    vertical_stacking_continuity,
    floor_summary,
)
from csv_exporter import rows_to_csv, COLUMNS


# ── parser tests ───────────────────────────────────────────────────────────────

def test_parse_three_parts():
    assert parse_type_name("Retail_ZoneA_L2") == ("Retail", "ZoneA", "L2")

def test_parse_two_parts():
    assert parse_type_name("Office_L3") == ("Office", "Unknown", "L3")

def test_parse_one_part():
    assert parse_type_name("Housing") == ("Housing", "Unknown", "Unknown")

def test_parse_empty():
    assert parse_type_name("") == ("Unknown", "Unknown", "Unknown")

def test_normalize_floor():
    assert normalize_floor_label("Level 2") == "L02"
    assert normalize_floor_label("L5")      == "L05"
    assert normalize_floor_label("Ground")  == "Ground"


# ── kpi tests ──────────────────────────────────────────────────────────────────

def test_shannon_single_program():
    """One program → H = 0 (no diversity)."""
    assert shannon_diversity({"Retail": 100}) == 0.0

def test_shannon_equal_programs():
    """Two equal programs → H ≈ ln(2) ≈ 0.6931."""
    import math
    h = shannon_diversity({"Retail": 50, "Office": 50})
    assert abs(h - math.log(2)) < 0.001

def test_shannon_zero_total():
    assert shannon_diversity({"Retail": 0, "Office": 0}) == 0.0


def test_mono_functional_triggered():
    progs  = {"Retail": 850, "Office": 100, "Housing": 50}
    thresholds = {"Retail": 60}
    is_mono, dominant, pct, allowed = mono_functional_check(progs, thresholds)
    assert is_mono is True
    assert dominant == "Retail"
    assert pct > 60

def test_mono_functional_ok():
    progs = {"Retail": 400, "Office": 350, "Housing": 250}
    is_mono, _, _, _ = mono_functional_check(progs, {"Retail": 60})
    assert is_mono is False


def test_zone_compatibility_flag():
    zone_data = {"ZoneA": {"Retail": 900, "Office": 100}}
    issues = check_zone_compatibility(zone_data, {"Retail": 60})
    assert len(issues) == 1
    assert "ZoneA" in issues[0]

def test_zone_compatibility_ok():
    zone_data = {"ZoneA": {"Retail": 400, "Office": 600}}
    issues = check_zone_compatibility(zone_data, {"Retail": 60})
    assert issues == []


def test_vertical_stacking():
    floor_data = {
        "L1": {"Retail": 200, "Office": 100},
        "L2": {"Office": 150, "Housing": 80},
        "L3": {"Office": 200},
    }
    result = vertical_stacking_continuity(floor_data)
    # Office appears on all 3 floors → 1.0
    assert result["Office"] == 1.0
    # Retail only on L1 → 1/3
    assert abs(result["Retail"] - round(1/3, 3)) < 0.001
    # Housing only on L2 → 1/3
    assert abs(result["Housing"] - round(1/3, 3)) < 0.001


# ── csv exporter tests ─────────────────────────────────────────────────────────

def test_csv_columns():
    row = {
        "Floor": "L2", "Zone": "ZoneA", "Program": "Retail",
        "Area": 450.0, "%": 82.0, "Dominant": "Retail",
        "Diversity Index (H)": 0.5, "Vertical Continuity": 0.67,
        "Status": "MONO-FUNCTIONAL",
    }
    csv_text = rows_to_csv([row])
    first_line = csv_text.splitlines()[0]
    for col in COLUMNS:
        assert col in first_line

def test_csv_values():
    row = {
        "Floor": "L1", "Zone": "ZoneB", "Program": "Office",
        "Area": 300.0, "%": 45.0, "Dominant": "Office",
        "Diversity Index (H)": 1.1, "Vertical Continuity": 1.0,
        "Status": "OK",
    }
    csv_text = rows_to_csv([row])
    assert "L1" in csv_text
    assert "Office" in csv_text
    assert "300.0" in csv_text


# ── integration-style test ─────────────────────────────────────────────────────

def test_full_floor_summary():
    prog_areas = {"Retail": 700, "Office": 200, "Housing": 100}
    thresholds = {"Retail": 60}
    s = floor_summary("L2", prog_areas, thresholds)
    assert s["floor"] == "L2"
    assert s["total_area"] == 1000.0
    assert s["dominant_program"] == "Retail"
    assert s["is_mono_functional"] is True
    assert s["diversity_index"] > 0


def test_parse_thresholds_compatibility():
    """Ensure threshold parsing survives legacy "function" wrapper and bad JSON."""
    from main import _parse_thresholds

    class Legacy:
        def __init__(self, matrix):
            # object with nested inputs attribute
            self.inputs = type("X", (), {"program_threshold_matrix": matrix})

    legacy = Legacy('{"Retail": 50}')
    assert _parse_thresholds(legacy) == {"Retail": 50}

    # dict-style input also accepted
    assert _parse_thresholds({"program_threshold_matrix": '{"Office": 75}'}) == {"Office": 75}

    # invalid JSON results in empty dict rather than an exception
    assert _parse_thresholds({"program_threshold_matrix": "not json"}) == {}


if __name__ == "__main__":
    # Quick smoke-run without pytest
    tests = [
        test_parse_three_parts, test_parse_two_parts, test_parse_one_part,
        test_parse_empty, test_normalize_floor,
        test_shannon_single_program, test_shannon_equal_programs, test_shannon_zero_total,
        test_mono_functional_triggered, test_mono_functional_ok,
        test_zone_compatibility_flag, test_zone_compatibility_ok,
        test_vertical_stacking,
        test_csv_columns, test_csv_values,
        test_full_floor_summary,
    ]
    passed = 0
    for t in tests:
        try:
            t()
            print(f"  PASS {t.__name__}")
            passed += 1
        except Exception as e:
            print(f"  FAIL {t.__name__}: {e}")
    print(f"\n{passed}/{len(tests)} tests passed.")

"""
utils/parser.py
───────────────
Parse program, zone, and floor information from Revit Type Names.

Supported naming conventions
─────────────────────────────
  Program_Zone_Floor      →  "Retail_ZoneA_L2"
  Program_Floor           →  "Office_L3"
  Program                 →  "Housing"

The parser is intentionally flexible; unknown parts default to "Unknown".
"""

from typing import Tuple


def parse_type_name(type_name: str) -> Tuple[str, str, str]:
    """
    Split a Revit Type Name into (program, zone, floor).

    Parameters
    ----------
    type_name : str
        Raw value from the Type Name parameter, e.g. "Retail_ZoneA_L2".

    Returns
    -------
    (program, zone, floor) : Tuple[str, str, str]
        Each value is a non-empty string; falls back to "Unknown".

    Examples
    --------
    >>> parse_type_name("Retail_ZoneA_L2")
    ('Retail', 'ZoneA', 'L2')

    >>> parse_type_name("Office_L3")
    ('Office', 'Unknown', 'L3')

    >>> parse_type_name("Housing")
    ('Housing', 'Unknown', 'Unknown')

    >>> parse_type_name("")
    ('Unknown', 'Unknown', 'Unknown')
    """
    if not type_name or not type_name.strip():
        return ("Unknown", "Unknown", "Unknown")

    parts = [p.strip() for p in type_name.strip().split("_") if p.strip()]

    if len(parts) >= 3:
        return (parts[0], parts[1], parts[2])
    elif len(parts) == 2:
        return (parts[0], "Unknown", parts[1])
    elif len(parts) == 1:
        return (parts[0], "Unknown", "Unknown")
    else:
        return ("Unknown", "Unknown", "Unknown")


def normalize_floor_label(floor: str) -> str:
    """
    Normalise a floor label to a sortable string.

    Examples
    --------
    >>> normalize_floor_label("Level 2")
    'L02'
    >>> normalize_floor_label("L2")
    'L02'
    >>> normalize_floor_label("Ground")
    'Ground'
    """
    import re

    floor = floor.strip()
    # Extract trailing number if present
    match = re.search(r"(\d+)$", floor)
    if match:
        number = int(match.group(1))
        return f"L{number:02d}"
    return floor

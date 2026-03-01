"""
utils/csv_exporter.py
─────────────────────
Build and export the program validation results to CSV format.
"""

import csv
import io
from typing import List, Dict


COLUMNS = [
    "Level",
    "Program",
    "Area",
    "Status",
]


def rows_to_csv(rows: List[Dict]) -> str:
    """
    Convert a list of row dicts to a UTF-8 CSV string.
    Uses QUOTE_MINIMAL to ensure proper escaping and Excel compatibility.

    Parameters
    ----------
    rows : list of dicts with keys matching COLUMNS

    Returns
    -------
    str  Complete CSV text (header + data rows).
    """
    buf = io.StringIO()
    # QUOTE_MINIMAL: quotes only when needed for special characters
    # dialect='excel': standard Windows/Mac/Linux Excel format
    writer = csv.DictWriter(
        buf, 
        fieldnames=COLUMNS, 
        extrasaction="ignore",
        quoting=csv.QUOTE_MINIMAL,
        lineterminator='\r\n'  # Windows-style line endings for Excel
    )
    writer.writeheader()
    writer.writerows(rows)
    return buf.getvalue()


def rows_to_floor_summary_csv(floor_data: dict, stacking: dict, thresholds: dict,
                               default_threshold: float) -> str:
    """
    Export a compact per-floor summary CSV (one row per floor).

    Columns: Floor | Total Area | Programs | Dominant | Dominant % | Diversity H | Stacking
    """
    from utils.kpi import floor_summary

    buf = io.StringIO()
    summary_cols = [
        "Floor", "Total Area (m²)", "# Programs", "Dominant Program",
        "Dominant %", "Allowed %", "Mono-Functional?", "Diversity Index (H)",
    ]
    writer = csv.DictWriter(buf, fieldnames=summary_cols)
    writer.writeheader()

    for floor, prog_areas in sorted(floor_data.items()):
        s = floor_summary(floor, prog_areas, thresholds, default_threshold)
        writer.writerow({
            "Floor":               s["floor"],
            "Total Area (m²)":     s["total_area"],
            "# Programs":          s["num_programs"],
            "Dominant Program":    s["dominant_program"],
            "Dominant %":          s["dominant_pct"],
            "Allowed %":           s["allowed_pct"],
            "Mono-Functional?":    "YES ⚠️" if s["is_mono_functional"] else "NO ✅",
            "Diversity Index (H)": s["diversity_index"],
        })
    return buf.getvalue()

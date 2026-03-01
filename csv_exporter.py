"""
utils/csv_exporter.py
─────────────────────
Build and export the program validation results to Excel format.
"""

import csv
import io
import tempfile
from typing import List, Dict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill


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


def rows_to_excel(rows: List[Dict], filename: str = None) -> str:
    """
    Convert a list of row dicts to a properly formatted Excel (.xlsx) file.
    
    Parameters
    ----------
    rows : list of dicts with keys matching COLUMNS
    filename : str, optional. If provided, writes to this file. Otherwise returns temp file path.
    
    Returns
    -------
    str  Path to the Excel file.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Program Floor Analysis"
    
    # Write headers
    for col_idx, header in enumerate(COLUMNS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        # Style header
        cell.font = Font(bold=True, color="FFFFFF", size=11)
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Write data rows
    for row_idx, row_data in enumerate(rows, 2):
        for col_idx, header in enumerate(COLUMNS, 1):
            value = row_data.get(header, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            
            # Center align numeric columns
            if header in ["Area"]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            # Center align Status column
            if header == "Status":
                cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Set column widths
    ws.column_dimensions['A'].width = 15  # Level
    ws.column_dimensions['B'].width = 20  # Program
    ws.column_dimensions['C'].width = 15  # Area
    ws.column_dimensions['D'].width = 35  # Status
    
    # Set row height for header
    ws.row_dimensions[1].height = 25
    
    # Save to file
    if filename is None:
        # Create a temporary file
        tmp_file = tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            delete=False,
            prefix="program_floor_validation_"
        )
        filename = tmp_file.name
        tmp_file.close()
    
    wb.save(filename)
    return filename


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

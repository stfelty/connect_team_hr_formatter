"""Excel formatting and labeling module.

Takes raw sheet data and produces a professionally formatted .xlsx file
with styled headers, column widths, borders, and a title label.
"""

import logging
import os
from datetime import datetime
from typing import List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# -- Style constants ----------------------------------------------------------

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color="1F4E79")
TITLE_FILL = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
TITLE_ALIGNMENT = Alignment(horizontal="center", vertical="center")

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)

DATA_FONT = Font(name="Calibri", size=11)
DATA_ALIGNMENT = Alignment(vertical="center", wrap_text=True)
EVEN_ROW_FILL = PatternFill(start_color="F2F7FB", end_color="F2F7FB", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

METADATA_FONT = Font(name="Calibri", size=9, italic=True, color="808080")


def format_excel(
    headers: List[str],
    rows: List[List[str]],
    output_dir: str,
    filename_prefix: str,
    title_label: str | None = None,
) -> str:
    """Create a formatted Excel workbook from the provided data.

    Args:
        headers: Column header names.
        rows: Data rows (list of lists).
        output_dir: Directory to write the output file.
        filename_prefix: Prefix for the output filename.
        title_label: Optional title label for the report. Defaults to the
            filename_prefix with spaces.

    Returns:
        The full path to the generated .xlsx file.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    os.makedirs(output_dir, exist_ok=True)

    if title_label is None:
        title_label = filename_prefix.replace("_", " ")

    num_cols = len(headers)

    wb = Workbook()
    ws = wb.active
    ws.title = "Report"

    # -- Title row (row 1) ----------------------------------------------------
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=num_cols)
    title_cell = ws.cell(row=1, column=1, value=title_label)
    title_cell.font = TITLE_FONT
    title_cell.fill = TITLE_FILL
    title_cell.alignment = TITLE_ALIGNMENT
    ws.row_dimensions[1].height = 36

    # -- Metadata row (row 2) -------------------------------------------------
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=num_cols)
    meta_cell = ws.cell(
        row=2, column=1,
        value=f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}  |  Rows: {len(rows)}",
    )
    meta_cell.font = METADATA_FONT
    meta_cell.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 20

    # -- Header row (row 3) ---------------------------------------------------
    header_row_idx = 3
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row_idx, column=col_idx, value=header)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER
    ws.row_dimensions[header_row_idx].height = 28

    # -- Data rows (starting row 4) -------------------------------------------
    data_start_row = 4
    for row_offset, row_data in enumerate(rows):
        row_idx = data_start_row + row_offset
        for col_idx in range(1, num_cols + 1):
            value = row_data[col_idx - 1] if col_idx - 1 < len(row_data) else ""
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = DATA_FONT
            cell.alignment = DATA_ALIGNMENT
            cell.border = THIN_BORDER
            # Alternate row shading
            if row_offset % 2 == 0:
                cell.fill = EVEN_ROW_FILL

    # -- Auto-fit column widths -----------------------------------------------
    for col_idx in range(1, num_cols + 1):
        max_length = len(str(headers[col_idx - 1]))
        for row_data in rows:
            if col_idx - 1 < len(row_data):
                max_length = max(max_length, len(str(row_data[col_idx - 1])))
        adjusted_width = min(max_length + 4, 50)
        ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

    # -- Freeze panes below header --------------------------------------------
    ws.freeze_panes = f"A{data_start_row}"

    # -- Auto-filter on header row --------------------------------------------
    ws.auto_filter.ref = f"A{header_row_idx}:{get_column_letter(num_cols)}{header_row_idx + len(rows)}"

    wb.save(filepath)
    logger.info("Excel file written to %s", filepath)
    return filepath

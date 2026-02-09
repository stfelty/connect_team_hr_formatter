"""Excel formatting and labeling module.

Produces a .xlsx file matching the exact HR hours summary format:
  - Row 1: date range (start date in A1, end date in B1)
  - Row 2: column headers (bold)
  - Row 3+: data rows with numeric hours formatted to 2 decimal places
  - Two sheet tabs: "Hours Summary Report" and the start date (MM.DD.YYYY)
"""

import logging
import os
from datetime import date, datetime
from typing import List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

logger = logging.getLogger(__name__)

# Column order in the output spreadsheet
OUTPUT_COLUMNS = [
    ("Employee Number", "employee_id"),
    ("Last Name",       "last_name"),
    ("First Name",      "first_name"),
    ("PayType Name",    "pay_type"),
    ("Regular Hours",   "regular_hours"),
    ("OT1 Hours",       "ot1_hours"),
    ("Paid Hours",      "paid_hours"),
    ("Unpaid Hours",    "unpaid_hours"),
]

NUMERIC_FIELDS = {"regular_hours", "ot1_hours", "paid_hours", "unpaid_hours"}

HEADER_FONT = Font(name="Calibri", size=11, bold=True)
DATA_FONT = Font(name="Calibri", size=11)
DATE_FONT = Font(name="Calibri", size=11, bold=True)


def _write_sheet(ws, start_date_str: str, end_date_str: str, summaries: List[dict]) -> None:
    """Write the standard hours summary layout to a worksheet."""

    # Row 1: Date range
    ws.cell(row=1, column=1, value=start_date_str).font = DATE_FONT
    ws.cell(row=1, column=2, value=end_date_str).font = DATE_FONT

    # Row 2: Headers
    for col_idx, (header_name, _) in enumerate(OUTPUT_COLUMNS, start=1):
        cell = ws.cell(row=2, column=col_idx, value=header_name)
        cell.font = HEADER_FONT

    # Row 3+: Data
    data_start_row = 3
    for row_offset, summary in enumerate(summaries):
        row_idx = data_start_row + row_offset
        for col_idx, (_, field_key) in enumerate(OUTPUT_COLUMNS, start=1):
            value = summary.get(field_key, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = DATA_FONT
            if field_key in NUMERIC_FIELDS:
                cell.number_format = "0.00"
                cell.alignment = Alignment(horizontal="right")

    # Auto-fit column widths
    for col_idx, (header_name, field_key) in enumerate(OUTPUT_COLUMNS, start=1):
        max_len = len(header_name)
        for summary in summaries:
            max_len = max(max_len, len(str(summary.get(field_key, ""))))
        ws.column_dimensions[get_column_letter(col_idx)].width = max_len + 3


def format_excel(
    summaries: List[dict],
    report_date: date,
    output_dir: str,
    filename_prefix: str,
    end_date: date | None = None,
) -> str:
    """Create the hours summary Excel workbook matching the required format.

    Args:
        summaries: List of employee summary dicts from data_processor.
        report_date: The report / pay-period start date.
        output_dir: Directory to write the output file.
        filename_prefix: Prefix for the output filename.
        end_date: Pay period end date. Defaults to report_date.

    Returns:
        The full path to the generated .xlsx file.
    """
    if end_date is None:
        end_date = report_date

    start_date_str = report_date.strftime("%m/%d/%Y")
    end_date_str = end_date.strftime("%m/%d/%Y")
    tab_date_label = report_date.strftime("%m.%d.%Y")

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{filename_prefix}_{timestamp}.xlsx"
    filepath = os.path.join(output_dir, filename)
    os.makedirs(output_dir, exist_ok=True)

    wb = Workbook()

    # Sheet 1: "Hours Summary Report"
    ws1 = wb.active
    ws1.title = "Hours Summary Report"
    _write_sheet(ws1, start_date_str, end_date_str, summaries)

    # Sheet 2: date tab (MM.DD.YYYY)
    ws2 = wb.create_sheet(title=tab_date_label)
    _write_sheet(ws2, start_date_str, end_date_str, summaries)

    wb.save(filepath)
    logger.info("Excel file written to %s", filepath)
    return filepath

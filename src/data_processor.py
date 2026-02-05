"""Data processing module.

Transforms raw ConnectTeam clock-event rows from the Google Sheet into
per-employee daily hour summaries for the Excel report.

Source columns (Google Sheet):
    A: ID            – employee number
    B: start         – human-readable start ("January 21 2026 10:55:00")
    C: Start Timestamp – unix epoch seconds
    D: end           – human-readable end   ("January 21 2026 19:06:01")
    E: End Timestamp – unix epoch seconds
    F: User Id       – employee number (same as ID)

Processing steps:
    1. Parse each row's start/end timestamps.
    2. Skip rows where start and end fall on different calendar dates
       (overnight shifts).
    3. Calculate shift duration in hours.
    4. Group shifts by (employee_id, date) and sum hours.
    5. Return summary rows ready for the Excel formatter.
"""

import logging
from collections import defaultdict
from datetime import datetime, date
from typing import List

logger = logging.getLogger(__name__)

# Possible formats for the human-readable date column
DATE_FORMATS = [
    "%B %d %Y %H:%M:%S",   # January 21 2026 10:55:00
    "%B %d %Y %H:%M",      # January 21 2026 10:55
    "%m/%d/%Y %H:%M:%S",   # 01/21/2026 10:55:00
    "%m/%d/%Y %H:%M",      # 01/21/2026 10:55
    "%Y-%m-%d %H:%M:%S",   # 2026-01-21 10:55:00
]


def _parse_datetime(value: str) -> datetime | None:
    """Try multiple date formats to parse a datetime string."""
    value = value.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def _parse_timestamp(value: str) -> datetime | None:
    """Parse a unix timestamp string to datetime."""
    try:
        ts = int(value.strip())
        return datetime.fromtimestamp(ts)
    except (ValueError, TypeError, OSError):
        return None


def _find_col(headers: List[str], *candidates: str) -> int | None:
    """Return the index of the first matching header (case-insensitive)."""
    lower = [h.strip().lower() for h in headers]
    for c in candidates:
        if c.lower() in lower:
            return lower.index(c.lower())
    return None


def process_clock_events(
    headers: List[str],
    rows: List[List[str]],
    target_date: date | None = None,
) -> tuple[date, List[dict]]:
    """Process raw clock events into per-employee daily hour summaries.

    Args:
        headers: Column headers from the Google Sheet.
        rows: Raw data rows from the Google Sheet.
        target_date: If provided, only include shifts on this date.
                     If None, uses the most recent date found in the data.

    Returns:
        A tuple of (report_date, summaries) where summaries is a list of dicts
        with keys: employee_id, regular_hours, ot1_hours, paid_hours,
        unpaid_hours, pay_type.
    """
    # -- Locate columns -------------------------------------------------------
    id_col = _find_col(headers, "id", "employee id", "employee number")
    start_col = _find_col(headers, "start")
    end_col = _find_col(headers, "end")
    start_ts_col = _find_col(headers, "start timestamp")
    end_ts_col = _find_col(headers, "end timestamp")
    user_id_col = _find_col(headers, "user id", "userid")

    # Prefer ID column, fall back to User Id
    emp_col = id_col if id_col is not None else user_id_col

    if emp_col is None:
        raise ValueError("Cannot find employee ID column in sheet headers: " + str(headers))

    # -- Parse each row -------------------------------------------------------
    shifts = []  # list of (employee_id, start_dt, end_dt, hours)
    skipped_overnight = 0
    skipped_parse = 0

    for row_idx, row in enumerate(rows):
        def cell(col_idx):
            if col_idx is not None and col_idx < len(row):
                return row[col_idx].strip()
            return ""

        employee_id = cell(emp_col)
        if not employee_id:
            continue

        # Parse start datetime (try human-readable first, then timestamp)
        start_dt = _parse_datetime(cell(start_col)) if start_col is not None else None
        if start_dt is None and start_ts_col is not None:
            start_dt = _parse_timestamp(cell(start_ts_col))

        # Parse end datetime
        end_dt = _parse_datetime(cell(end_col)) if end_col is not None else None
        if end_dt is None and end_ts_col is not None:
            end_dt = _parse_timestamp(cell(end_ts_col))

        if start_dt is None or end_dt is None:
            skipped_parse += 1
            logger.warning("Row %d: could not parse start/end times, skipping", row_idx + 2)
            continue

        # Filter out overnight shifts (start and end on different dates)
        if start_dt.date() != end_dt.date():
            skipped_overnight += 1
            logger.debug(
                "Row %d: overnight shift (%s -> %s), skipping",
                row_idx + 2, start_dt.date(), end_dt.date(),
            )
            continue

        # Calculate hours
        duration_seconds = (end_dt - start_dt).total_seconds()
        if duration_seconds <= 0:
            logger.warning("Row %d: non-positive duration, skipping", row_idx + 2)
            continue

        hours = round(duration_seconds / 3600, 2)
        shifts.append((employee_id, start_dt.date(), hours))

    logger.info(
        "Parsed %d valid shifts (%d overnight skipped, %d unparseable)",
        len(shifts), skipped_overnight, skipped_parse,
    )

    if not shifts:
        raise ValueError("No valid shifts found in the data")

    # -- Determine report date ------------------------------------------------
    all_dates = sorted(set(s[1] for s in shifts))

    if target_date:
        report_date = target_date
    else:
        # Default to the most recent date in the data
        report_date = all_dates[-1]

    logger.info("Report date: %s (dates in data: %s)", report_date, all_dates)

    # -- Filter to target date and group by employee --------------------------
    daily: dict[str, float] = defaultdict(float)
    for emp_id, shift_date, hours in shifts:
        if shift_date == report_date:
            daily[emp_id] += hours

    if not daily:
        raise ValueError(f"No shifts found for date {report_date}")

    # -- Build summary rows ---------------------------------------------------
    summaries = []
    for emp_id in sorted(daily.keys()):
        regular = round(daily[emp_id], 2)
        summaries.append({
            "employee_id": emp_id,
            "last_name": "",
            "first_name": "",
            "pay_type": "Work",
            "regular_hours": regular,
            "ot1_hours": 0.0,
            "paid_hours": regular,
            "unpaid_hours": 0.0,
        })

    logger.info("Built summaries for %d employees on %s", len(summaries), report_date)
    return report_date, summaries

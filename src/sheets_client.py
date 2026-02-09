"""Google Sheets data extraction module.

Authenticates via a service account and pulls all rows/columns
from the configured spreadsheet range.
"""

import logging
from typing import List

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

logger = logging.getLogger(__name__)

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]


def get_sheets_service(service_account_file: str):
    """Build and return an authenticated Google Sheets API service."""
    creds = Credentials.from_service_account_file(service_account_file, scopes=SCOPES)
    service = build("sheets", "v4", credentials=creds)
    return service.spreadsheets()


def fetch_sheet_data(
    service_account_file: str,
    spreadsheet_id: str,
    sheet_range: str,
) -> tuple[List[str], List[List[str]]]:
    """Fetch data from a Google Sheet.

    Args:
        service_account_file: Path to the Google service account JSON key file.
        spreadsheet_id: The ID of the Google Spreadsheet.
        sheet_range: The A1-notation range to read (e.g. "Sheet1" or "Sheet1!A1:Z100").

    Returns:
        A tuple of (headers, rows) where headers is the first row and rows
        is a list of subsequent rows.

    Raises:
        ValueError: If the sheet is empty or contains no data.
    """
    logger.info("Fetching data from spreadsheet %s range '%s'", spreadsheet_id, sheet_range)

    sheets = get_sheets_service(service_account_file)
    result = sheets.values().get(
        spreadsheetId=spreadsheet_id,
        range=sheet_range,
    ).execute()

    values = result.get("values", [])

    if not values:
        raise ValueError(f"No data found in spreadsheet {spreadsheet_id} range '{sheet_range}'")

    headers = values[0]
    rows = values[1:]

    logger.info("Fetched %d rows with %d columns", len(rows), len(headers))
    return headers, rows

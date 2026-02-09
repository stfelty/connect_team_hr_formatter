"""Connect Team HR Formatter — main pipeline.

Orchestrates the full workflow:
  1. Pull raw clock events from a Google Sheet (populated by Zapier)
  2. Process: parse timestamps, filter overnight shifts, sum hours per employee
  3. Format as a labeled Excel spreadsheet (exact HR hours summary format)
  4. Upload the file to Amazon S3
"""

import argparse
import logging
import sys
from datetime import datetime, date

from config import Config
from src.sheets_client import fetch_sheet_data
from src.data_processor import process_clock_events
from src.excel_formatter import format_excel
from src.s3_uploader import upload_to_s3

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def _parse_date(value: str) -> date:
    """Parse a date string in MM/DD/YYYY format."""
    return datetime.strptime(value, "%m/%d/%Y").date()


def run_pipeline(
    skip_upload: bool = False,
    target_date: date | None = None,
    end_date: date | None = None,
) -> None:
    """Execute the full Sheets -> process -> Excel -> S3 pipeline.

    Args:
        skip_upload: If True, generate the Excel file but do not upload to S3.
        target_date: The date to generate the report for. Defaults to the most
                     recent date found in the sheet data.
        end_date: Optional end date for the pay period header. Defaults to
                  the same as target_date (single-day report).
    """
    # -- Validate configuration ------------------------------------------------
    if not Config.SPREADSHEET_ID:
        logger.error("GOOGLE_SHEETS_SPREADSHEET_ID is not set. Check your .env file.")
        sys.exit(1)

    if not skip_upload and not Config.S3_BUCKET_NAME:
        logger.error("S3_BUCKET_NAME is not set. Check your .env file or use --skip-upload.")
        sys.exit(1)

    # -- Step 1: Fetch raw clock events from Google Sheets ---------------------
    logger.info("Step 1/4 — Fetching clock events from Google Sheets")
    headers, rows = fetch_sheet_data(
        service_account_file=Config.SERVICE_ACCOUNT_FILE,
        spreadsheet_id=Config.SPREADSHEET_ID,
        sheet_range=Config.SHEET_RANGE,
    )

    # -- Step 2: Process into per-employee daily summaries ---------------------
    logger.info("Step 2/4 — Processing clock events")
    report_date, summaries = process_clock_events(
        headers=headers,
        rows=rows,
        target_date=target_date,
    )

    if end_date is None:
        end_date = report_date

    # -- Step 3: Format as Excel -----------------------------------------------
    logger.info("Step 3/4 — Formatting Excel spreadsheet")
    filepath = format_excel(
        summaries=summaries,
        report_date=report_date,
        output_dir=Config.OUTPUT_DIR,
        filename_prefix=Config.OUTPUT_FILENAME_PREFIX,
        end_date=end_date,
    )
    logger.info("Excel file created: %s", filepath)

    # -- Step 4: Upload to S3 --------------------------------------------------
    if skip_upload:
        logger.info("Step 4/4 — Skipping S3 upload (--skip-upload flag set)")
    else:
        logger.info("Step 4/4 — Uploading to S3")
        s3_key = upload_to_s3(
            filepath=filepath,
            bucket_name=Config.S3_BUCKET_NAME,
            s3_prefix=Config.S3_PREFIX,
            aws_access_key_id=Config.AWS_ACCESS_KEY_ID,
            aws_secret_access_key=Config.AWS_SECRET_ACCESS_KEY,
            aws_region=Config.AWS_REGION,
        )
        logger.info("Uploaded to s3://%s/%s", Config.S3_BUCKET_NAME, s3_key)

    logger.info("Pipeline complete.")


def main():
    parser = argparse.ArgumentParser(
        description="Pull ConnectTeam clock events, summarise daily hours, "
                    "format as HR hours summary Excel, and upload to S3.",
    )
    parser.add_argument(
        "--skip-upload",
        action="store_true",
        help="Generate the Excel file locally without uploading to S3.",
    )
    parser.add_argument(
        "--date",
        type=str,
        default=None,
        help="Target report date in MM/DD/YYYY format. "
             "Defaults to the most recent date in the sheet.",
    )
    parser.add_argument(
        "--end-date",
        type=str,
        default=None,
        help="Pay period end date in MM/DD/YYYY (shown in cell B1). "
             "Defaults to --date value.",
    )
    args = parser.parse_args()

    target_date = _parse_date(args.date) if args.date else None
    end_date = _parse_date(args.end_date) if args.end_date else None

    run_pipeline(
        skip_upload=args.skip_upload,
        target_date=target_date,
        end_date=end_date,
    )


if __name__ == "__main__":
    main()

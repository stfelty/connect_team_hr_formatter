"""Connect Team HR Formatter — main pipeline.

Orchestrates the full workflow:
  1. Pull data from a Google Sheet
  2. Format it as a labeled Excel spreadsheet (exact HR hours summary format)
  3. Upload the file to Amazon S3
"""

import argparse
import logging
import sys

from config import Config
from src.sheets_client import fetch_sheet_data
from src.excel_formatter import format_excel
from src.s3_uploader import upload_to_s3

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)


def run_pipeline(
    skip_upload: bool = False,
    start_date: str | None = None,
    end_date: str | None = None,
) -> None:
    """Execute the full Sheets -> Excel -> S3 pipeline.

    Args:
        skip_upload: If True, generate the Excel file but do not upload to S3.
        start_date: Pay period start date (MM/DD/YYYY).
        end_date: Pay period end date (MM/DD/YYYY).
    """
    # -- Validate configuration ------------------------------------------------
    if not Config.SPREADSHEET_ID:
        logger.error("GOOGLE_SHEETS_SPREADSHEET_ID is not set. Check your .env file.")
        sys.exit(1)

    if not skip_upload and not Config.S3_BUCKET_NAME:
        logger.error("S3_BUCKET_NAME is not set. Check your .env file or use --skip-upload.")
        sys.exit(1)

    # -- Step 1: Fetch data from Google Sheets ---------------------------------
    logger.info("Step 1/3 — Fetching data from Google Sheets")
    headers, rows = fetch_sheet_data(
        service_account_file=Config.SERVICE_ACCOUNT_FILE,
        spreadsheet_id=Config.SPREADSHEET_ID,
        sheet_range=Config.SHEET_RANGE,
    )

    # -- Step 2: Format as Excel -----------------------------------------------
    logger.info("Step 2/3 — Formatting Excel spreadsheet")
    filepath = format_excel(
        headers=headers,
        rows=rows,
        output_dir=Config.OUTPUT_DIR,
        filename_prefix=Config.OUTPUT_FILENAME_PREFIX,
        start_date=start_date,
        end_date=end_date,
    )
    logger.info("Excel file created: %s", filepath)

    # -- Step 3: Upload to S3 --------------------------------------------------
    if skip_upload:
        logger.info("Step 3/3 — Skipping S3 upload (--skip-upload flag set)")
    else:
        logger.info("Step 3/3 — Uploading to S3")
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
        description="Pull Google Sheet data, format as HR hours summary Excel, and upload to S3.",
    )
    parser.add_argument(
        "--skip-upload",
        action="store_true",
        help="Generate the Excel file locally without uploading to S3.",
    )
    parser.add_argument(
        "--start-date",
        type=str,
        default=None,
        help="Pay period start date in MM/DD/YYYY format (shown in cell A1).",
    )
    parser.add_argument(
        "--end-date",
        type=str,
        default=None,
        help="Pay period end date in MM/DD/YYYY format (shown in cell B1).",
    )
    args = parser.parse_args()

    run_pipeline(
        skip_upload=args.skip_upload,
        start_date=args.start_date,
        end_date=args.end_date,
    )


if __name__ == "__main__":
    main()

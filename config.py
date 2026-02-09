"""Application configuration loaded from environment variables."""

import os
from dotenv import load_dotenv

load_dotenv()


class Config:
    """Central configuration for the Sheets-to-S3 pipeline."""

    # Google Sheets
    SPREADSHEET_ID = os.getenv("GOOGLE_SHEETS_SPREADSHEET_ID", "")
    SHEET_RANGE = os.getenv("GOOGLE_SHEETS_RANGE", "Sheet1")
    SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")

    # AWS S3
    AWS_ACCESS_KEY_ID = os.getenv("AWS_ACCESS_KEY_ID", "")
    AWS_SECRET_ACCESS_KEY = os.getenv("AWS_SECRET_ACCESS_KEY", "")
    AWS_REGION = os.getenv("AWS_REGION", "us-east-1")
    S3_BUCKET_NAME = os.getenv("S3_BUCKET_NAME", "")
    S3_PREFIX = os.getenv("S3_PREFIX", "hr-reports/")

    # Output
    OUTPUT_DIR = os.getenv("OUTPUT_DIR", "output")
    OUTPUT_FILENAME_PREFIX = os.getenv("OUTPUT_FILENAME_PREFIX", "HR_Report")

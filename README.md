# Connect Team HR Formatter

A pipeline that pulls data from a Google Sheet, formats it as a professionally styled Excel spreadsheet, and uploads it to Amazon S3.

## Pipeline Overview

```
Google Sheet  →  Formatted .xlsx  →  Amazon S3
```

1. **Extract** — Reads rows and columns from a Google Sheet via the Sheets API (service account auth)
2. **Format** — Creates a styled Excel file with title label, headers, alternating row colors, auto-fit columns, filters, and frozen panes
3. **Upload** — Pushes the `.xlsx` file to an S3 bucket with a configurable key prefix

## Setup

### 1. Install dependencies

```bash
python -m venv venv
source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Configure environment

Copy the example env file and fill in your values:

```bash
cp .env.example .env
```

| Variable | Description |
|---|---|
| `GOOGLE_SHEETS_SPREADSHEET_ID` | The ID from your Google Sheet URL |
| `GOOGLE_SHEETS_RANGE` | Sheet name or A1 range (default: `Sheet1`) |
| `GOOGLE_SERVICE_ACCOUNT_FILE` | Path to your service account JSON key |
| `AWS_ACCESS_KEY_ID` | AWS access key (or use default credential chain) |
| `AWS_SECRET_ACCESS_KEY` | AWS secret key |
| `AWS_REGION` | AWS region (default: `us-east-1`) |
| `S3_BUCKET_NAME` | Target S3 bucket name |
| `S3_PREFIX` | Key prefix / folder in the bucket (default: `hr-reports/`) |
| `OUTPUT_DIR` | Local output directory (default: `output`) |
| `OUTPUT_FILENAME_PREFIX` | Filename prefix for generated files (default: `HR_Report`) |

### 3. Google Sheets API access

1. Create a project in [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the **Google Sheets API**
3. Create a **Service Account** and download the JSON key file
4. Place the key file in the project root (it's gitignored) and set `GOOGLE_SERVICE_ACCOUNT_FILE` in `.env`
5. Share your Google Sheet with the service account email address (viewer access is sufficient)

## Usage

Run the full pipeline:

```bash
python main.py
```

Generate Excel locally without uploading to S3:

```bash
python main.py --skip-upload
```

Set a custom report title:

```bash
python main.py --title "Q1 2026 HR Report"
```

## Project Structure

```
connect_team_hr_formatter/
├── main.py                  # Pipeline orchestrator
├── config.py                # Environment-based configuration
├── requirements.txt         # Python dependencies
├── .env.example             # Environment variable template
├── .gitignore
├── src/
│   ├── __init__.py
│   ├── sheets_client.py     # Google Sheets data extraction
│   ├── excel_formatter.py   # Excel formatting and labeling
│   └── s3_uploader.py       # S3 upload
└── output/                  # Generated Excel files (gitignored)
```

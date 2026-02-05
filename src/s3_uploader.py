"""Amazon S3 upload module.

Handles uploading the formatted Excel file to a configured S3 bucket
with an optional key prefix.
"""

import logging
import os

import boto3
from botocore.exceptions import ClientError

logger = logging.getLogger(__name__)


def upload_to_s3(
    filepath: str,
    bucket_name: str,
    s3_prefix: str = "",
    aws_access_key_id: str | None = None,
    aws_secret_access_key: str | None = None,
    aws_region: str = "us-east-1",
) -> str:
    """Upload a file to an S3 bucket.

    Args:
        filepath: Local path to the file to upload.
        bucket_name: Name of the target S3 bucket.
        s3_prefix: Optional prefix (folder path) in the bucket.
        aws_access_key_id: AWS access key. If None, uses default credential chain.
        aws_secret_access_key: AWS secret key. If None, uses default credential chain.
        aws_region: AWS region for the S3 client.

    Returns:
        The full S3 key (path) of the uploaded object.

    Raises:
        FileNotFoundError: If the local file doesn't exist.
        ClientError: If the S3 upload fails.
    """
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"File not found: {filepath}")

    filename = os.path.basename(filepath)
    s3_key = f"{s3_prefix}{filename}" if s3_prefix else filename

    session_kwargs = {"region_name": aws_region}
    if aws_access_key_id and aws_secret_access_key:
        session_kwargs["aws_access_key_id"] = aws_access_key_id
        session_kwargs["aws_secret_access_key"] = aws_secret_access_key

    s3_client = boto3.client("s3", **session_kwargs)

    logger.info("Uploading %s to s3://%s/%s", filepath, bucket_name, s3_key)

    try:
        s3_client.upload_file(
            Filename=filepath,
            Bucket=bucket_name,
            Key=s3_key,
            ExtraArgs={"ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
        )
    except ClientError as exc:
        logger.error("S3 upload failed: %s", exc)
        raise

    logger.info("Upload complete: s3://%s/%s", bucket_name, s3_key)
    return s3_key

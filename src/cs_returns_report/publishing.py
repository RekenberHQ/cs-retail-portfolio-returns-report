"""Direct VPN publication to the app's S3-compatible store and PostgreSQL registry."""
from datetime import datetime
import hashlib
import logging
import os
from pathlib import Path
import re
from uuid import uuid4
from zoneinfo import ZoneInfo


REPORT_TYPE = "cs-retail-portfolio-returns"
CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def required(name):
    value = os.getenv(name)
    if not value:
        raise ValueError(f"Set {name} to publish reports")
    return value


def publish_report(path: Path, report_name: str, row_count: int):
    import boto3
    from botocore.config import Config
    import psycopg

    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", report_name):
        raise ValueError("Invalid report name")
    organization_id = int(os.getenv("REKENBER_REPORT_ORGANIZATION_ID", "2"))
    if organization_id < 1 or row_count < 0:
        raise ValueError("Invalid organization ID or row count")
    connection_string = required("REKENBER_APP_DB_URL")
    bucket = required("REKENBER_REPORT_S3_BUCKET")
    root = required("REKENBER_REPORT_S3_ROOT").strip("/")
    s3 = boto3.client(
        "s3",
        endpoint_url=os.getenv("REKENBER_REPORT_S3_ENDPOINT") or None,
        region_name=os.getenv("REKENBER_REPORT_S3_REGION", "us-east-1"),
        aws_access_key_id=required("REKENBER_REPORT_S3_ACCESS_KEY_ID"),
        aws_secret_access_key=required("REKENBER_REPORT_S3_SECRET_ACCESS_KEY"),
        config=Config(signature_version="s3v4", s3={"addressing_style": "path"}, retries={"max_attempts": 3}),
    )
    content = path.read_bytes()
    digest = hashlib.sha256(content).hexdigest()
    upload_date = datetime.now(ZoneInfo("Europe/Prague")).date().isoformat()
    filename = f"cs-retail-portfolio-returns-{upload_date}.xlsx"
    # Immutable revisions keep the previous download intact if upload/DB registration fails.
    key = f"{root}/reports/organizations/{organization_id}/{REPORT_TYPE}/{upload_date}/{uuid4().hex}/{filename}"
    with psycopg.connect(connection_string, connect_timeout=15) as connection:
        with connection.cursor() as cursor:
            cursor.execute("SET LOCAL statement_timeout = '120s'")
            cursor.execute("""
                SELECT o.id FROM public.organization o
                JOIN public.organization_settings s ON s.organization_id = o.id
                WHERE o.id = %s AND o.umbrella_organization_id IS NULL AND s.reports_enabled
            """, (organization_id,))
            if cursor.fetchone() is None:
                raise ValueError("Reports must be enabled for the target umbrella organization before publication")
            logging.info("Publishing %s to organization %d (%d bytes)", report_name, organization_id, len(content))
            s3.put_object(Bucket=bucket, Key=key, Body=content, ContentType=CONTENT_TYPE)
            # Only publish the pointer after storage succeeds. Unique key makes reruns update one row.
            cursor.execute("""
                INSERT INTO public.generated_report
                    (organization_id, report_type, report_name, title, file_name,
                     storage_bucket, storage_key, content_sha256, byte_size, row_count)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (organization_id, report_type, report_name) DO UPDATE SET
                    title = EXCLUDED.title, file_name = EXCLUDED.file_name,
                    storage_bucket = EXCLUDED.storage_bucket, storage_key = EXCLUDED.storage_key,
                    content_sha256 = EXCLUDED.content_sha256,
                    byte_size = EXCLUDED.byte_size, row_count = EXCLUDED.row_count,
                    generated_at = CURRENT_TIMESTAMP
                RETURNING id
            """, (organization_id, REPORT_TYPE, upload_date, f"Návratnost retailového portfolia – {upload_date}", filename,
                  bucket, key, digest, len(content), row_count))
            report_id = cursor.fetchone()[0]
    # Old revisions are intentionally retained; no deletion of an object a download may be using.
    logging.info("Published report ID %s", report_id)
    return {"id": report_id, "organization_id": organization_id, "storage_key": key, "sha256": digest}

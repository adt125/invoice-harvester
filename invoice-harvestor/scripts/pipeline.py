#!/usr/bin/env python3
"""
Unified invoice harvester pipeline: Download → Extract → Push

This orchestrates all three steps (download from Outlook, extract from PDFs,
push to Google Sheets) with a single configuration and execution flow.

Usage:
    python3 pipeline.py --folder "Invoices" --last-month --replace

Or as a library:
    from pipeline import run_pipeline, PipelineConfig
    config = PipelineConfig(folder="Invoices", ...)
    result = run_pipeline(config)
"""

from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Optional

# Import core functions from sibling scripts
from download_outlook_attachments import (
    download_attachments_core,
    DownloadResult,
    GraphError,
    previous_calendar_month,
)
from extract_invoice_items import (
    extract_items_core,
    ExtractResult,
    ExtractError,
)
from push_to_google_sheets import (
    push_to_sheets_core,
    PushResult,
    SheetsPushError,
)
from clean_up import cleanup_temp_files, CleanupResult

SCRIPT_DIR = Path(__file__).resolve().parent  # Points to 'scripts/'
SKILL_ROOT = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_ROOT / "assets"
TEMP_DIR = ASSETS_DIR / "temp"
DEFAULT_OUTPUT_DIR = TEMP_DIR

DEFAULT_ATTACHMENTS_DIR = TEMP_DIR
DEFAULT_CSV_FILE = TEMP_DIR / "invoice_items.csv"


@dataclass
class PipelineConfig:
    """Configuration for the invoice harvester pipeline."""

    folder: str
    outlook_client_id: str
    outlook_tenant_id: str = "common"
    google_service_account_file: Optional[str | Path] = None
    google_sheet_id: Optional[str] = None

    # Directory/file overrides
    attachments_dir: Path = field(default_factory=lambda: DEFAULT_ATTACHMENTS_DIR)
    csv_file: Path = field(default_factory=lambda: DEFAULT_CSV_FILE)

    # Execution flags
    last_month: bool = False
    replace: bool = False
    dry_run: bool = False

    def validate(self) -> None:
        """Validate required fields."""
        if not self.google_service_account_file:
            raise ValueError("Missing google_service_account_file")
        if not self.google_sheet_id:
            raise ValueError("Missing google_sheet_id")


@dataclass
class PipelineResult:
    """Result from running the full pipeline."""

    download: Optional[DownloadResult] = None
    extract: Optional[ExtractResult] = None
    push: Optional[PushResult] = None
    clean_up: Optional[CleanupResult] = None
    success: bool = False
    error: Optional[str] = None

    def summary(self) -> str:
        """Return a human-readable summary of the pipeline execution."""
        if not self.success:
            return f"Pipeline failed: {self.error}"
        return f"""Pipeline Complete:
- Downloaded: {self.download.count} file(s)
- Extracted: {self.extract.count} item(s) to {self.extract.csv_path}
- Pushed: {self.push.rows_updated} row(s) to {len(self.push.sheets_modified)} sheet(s)"""


def run_pipeline(config: PipelineConfig) -> PipelineResult:
    """
    Execute: download → extract → push

    Returns structured result with all three step outputs.
    Each step is optional; early exit if no data from previous step.
    """

    try:
        config.validate()
        # Step 1: Download
        print(f"[1/4] Downloading from Outlook folder '{config.folder}'...")

        # Calculate date range if needed
        from_date = None
        to_date = None

        download_result = download_attachments_core(
            folder=config.folder,
            client_id=config.outlook_client_id,
            tenant_id=config.outlook_tenant_id,
            output_dir=config.attachments_dir,
            from_date=from_date,
            to_date=to_date,
            last_month=config.last_month,
            dry_run=config.dry_run,
        )
        print(f"  ✓ Downloaded {download_result.count} file(s)")

        if config.dry_run or download_result.count == 0:
            print("  ✓ Skipping extraction (no files)")
            return PipelineResult(
                download=download_result,
                extract=ExtractResult(count=0, csv_path=config.csv_file, items=[]),
                push=PushResult(rows_updated=0, sheets_modified=[], replace=False),
                success=True,
            )

        # Step 2: Extract
        print(f"[2/4] Extracting invoices from {config.attachments_dir}...")
        extract_result = extract_items_core(
            input_dir=config.attachments_dir,
            output_file=config.csv_file,
        )
        print(f"  ✓ Extracted {extract_result.count} item(s) to {config.csv_file}")

        if extract_result.count == 0:
            print("  ✓ Skipping push (no items)")
            return PipelineResult(
                download=download_result,
                extract=extract_result,
                push=PushResult(rows_updated=0, sheets_modified=[], replace=False),
                success=True,
            )

        # Step 3: Push
        print(f"[3/4] Pushing to Google Sheets (replace={config.replace})...")
        push_result = push_to_sheets_core(
            csv_file=config.csv_file,
            service_account_file=config.google_service_account_file,
            sheet_id=config.google_sheet_id,
            replace=config.replace,
        )
        print(
            f"  ✓ Pushed {push_result.rows_updated} row(s) to {len(push_result.sheets_modified)} sheet(s)"
        )

        # Step 4: Cleanup
        print(f"[4/4] cleaning up temp file in {TEMP_DIR})...")
        cleanup_result = cleanup_temp_files()

        return PipelineResult(
            download=download_result,
            extract=extract_result,
            push=push_result,
            clean_up=cleanup_result,
            success=True,
        )

    except (GraphError, ExtractError, SheetsPushError, ValueError) as e:
        return PipelineResult(
            success=False,
            error=str(e),
        )
    except Exception as e:
        return PipelineResult(
            success=False,
            error=f"{type(e).__name__}: {str(e)}",
        )


def parse_args() -> argparse.Namespace:
    """Parse command-line arguments."""
    from dotenv import load_dotenv

    load_dotenv(dotenv_path=ASSETS_DIR / ".env")
    parser = argparse.ArgumentParser(
        description="Invoice Harvester Pipeline: Download → Extract → Push"
    )

    parser.add_argument(
        "--folder",
        required=False,
        default="Instamart",
        help='Outlook folder name (e.g., "Invoices" or "Inbox/Invoices")',
    )

    # Pipeline flags
    parser.add_argument(
        "--last-month",
        action="store_true",
        help="Filter to previous calendar month",
    )
    parser.add_argument(
        "--replace",
        action="store_true",
        help="Replace existing sheet data instead of appending",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Preview without making changes",
    )

    # Outlook config (from .env or CLI)
    parser.add_argument(
        "--client-id",
        default=os.getenv("OUTLOOK_CLIENT_ID"),
        help="Outlook client ID (or OUTLOOK_CLIENT_ID env var)",
    )
    parser.add_argument(
        "--tenant-id",
        default=os.getenv("OUTLOOK_TENANT_ID", "common"),
        help="Outlook tenant ID (default: common)",
    )

    # Google Sheets config (from .env or CLI)
    parser.add_argument(
        "--service-account",
        default=os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE"),
        help="Google service account JSON path (or GOOGLE_SERVICE_ACCOUNT_FILE env var)",
    )
    parser.add_argument(
        "--sheet-id",
        default=os.getenv("GOOGLE_SHEET_ID"),
        help="Google Sheet ID (or GOOGLE_SHEET_ID env var)",
    )

    # Optional directory overrides
    parser.add_argument(
        "--attachments-dir",
        default=str(DEFAULT_ATTACHMENTS_DIR),
        help=f"Directory for downloaded attachments (default: {DEFAULT_ATTACHMENTS_DIR})",
    )
    parser.add_argument(
        "--csv-file",
        default=str(DEFAULT_CSV_FILE),
        help=f"CSV output file (default: {DEFAULT_CSV_FILE})",
    )

    return parser.parse_args()


def main() -> int:
    """CLI entry point for the full pipeline."""
    args = parse_args()

    try:
        config = PipelineConfig(
            folder=args.folder,
            outlook_client_id=args.client_id,
            outlook_tenant_id=args.tenant_id,
            google_service_account_file=args.service_account,
            google_sheet_id=args.sheet_id,
            attachments_dir=Path(args.attachments_dir),
            csv_file=Path(args.csv_file),
            last_month=args.last_month,
            replace=args.replace,
            dry_run=args.dry_run,
        )
        result = run_pipeline(config)

        print()
        print(result.summary())
        print()

        return 0 if result.success else 1
    except Exception as e:
        print(f"Fatal error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""
Push the extracted invoice CSV to Google Sheets.

By default this groups CSV rows by invoice month and writes each month to a
separate sheet tab, for example "2026-05". Use --replace to clear each monthly
tab before writing.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from dotenv import load_dotenv, find_dotenv

load_dotenv(find_dotenv())

SCRIPT_DIR = Path(__file__).resolve().parent  # Points to 'scripts/'
SKILL_ROOT = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_ROOT / "assets"
TEMP_DIR = ASSETS_DIR / "temp"
DEFAULT_INPUT_FILE = TEMP_DIR / "invoice_items.csv"
DEFAULT_SHEET_COLUMNS = "A:E"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class SheetsPushError(RuntimeError):
    pass


@dataclass
class PushResult:
    """Result of pushing invoice items to Google Sheets."""

    rows_updated: int
    sheets_modified: list[str]
    replace: bool


def parse_args() -> argparse.Namespace:
    load_dotenv(ASSETS_DIR / ".env")

    parser = argparse.ArgumentParser(
        description="Push invoice_items.csv to a Google Sheet."
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT_FILE),
        help=f"CSV input path. Default: {DEFAULT_INPUT_FILE}",
    )
    parser.add_argument(
        "--service-account-file",
        default=os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE"),
        help="Google service account JSON path. Can also use GOOGLE_SERVICE_ACCOUNT_FILE in .env.",
    )
    parser.add_argument(
        "--sheet-id",
        default=os.environ.get("GOOGLE_SHEET_ID"),
        help="Google spreadsheet ID. Can also use GOOGLE_SHEET_ID in .env.",
    )
    parser.add_argument(
        "--range",
        default=os.environ.get("GOOGLE_SHEET_RANGE", DEFAULT_SHEET_COLUMNS),
        help=f"Target columns/range. Can also use GOOGLE_SHEET_RANGE in .env. Default: {DEFAULT_SHEET_COLUMNS}",
    )
    parser.add_argument(
        "--replace",
        action="store_true",
        help="Clear each monthly sheet range before writing CSV header and rows.",
    )
    parser.add_argument(
        "--include-header",
        action="store_true",
        help="Include the CSV header when appending. Headers are always included with --replace.",
    )
    return parser.parse_args()


def project_path(path_text: str) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return SCRIPT_DIR / path


def read_csv_rows(csv_path: Path) -> tuple[list[str], list[list[str]]]:
    if not csv_path.exists():
        raise SheetsPushError(f"CSV file does not exist: {csv_path}")

    with csv_path.open(newline="", encoding="utf-8") as csv_file:
        reader = csv.reader(csv_file)
        rows = list(reader)

    if not rows:
        raise SheetsPushError(f"CSV file is empty: {csv_path}")

    header = rows[0]
    data_rows = rows[1:]
    return header, data_rows


def month_key(row: list[str], header: list[str]) -> str:
    try:
        date_index = header.index("invoice_date")
    except ValueError as exc:
        raise SheetsPushError("CSV must contain an invoice_date column.") from exc

    if date_index >= len(row) or not row[date_index]:
        raise SheetsPushError(f"CSV row is missing invoice_date: {row}")

    value = row[date_index]
    for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m")
        except ValueError:
            pass
    raise SheetsPushError(f"Could not parse invoice_date {value!r}.")


def group_rows_by_month(
    header: list[str], data_rows: list[list[str]]
) -> "OrderedDict[str, list[list[str]]]":
    grouped: "OrderedDict[str, list[list[str]]]" = OrderedDict()
    for row in data_rows:
        key = month_key(row, header)
        grouped.setdefault(key, []).append(row)
    return grouped


def range_columns(target_range: str) -> str:
    columns = target_range.split("!", 1)[-1]
    if not re.fullmatch(r"[A-Za-z]+(?::[A-Za-z]+)?", columns):
        raise SheetsPushError(
            f"GOOGLE_SHEET_RANGE must be columns like A:E, not {target_range!r}."
        )
    return columns.upper()


def quote_sheet_name(sheet_name: str) -> str:
    return "'" + sheet_name.replace("'", "''") + "'"


def month_range(sheet_name: str, columns: str) -> str:
    return f"{quote_sheet_name(sheet_name)}!{columns}"


def sheets_service(service_account_file: Path):
    if not service_account_file.exists():
        raise SheetsPushError(
            f"Service account JSON file does not exist: {service_account_file}"
        )

    credentials = Credentials.from_service_account_file(
        service_account_file,
        scopes=SCOPES,
    )
    return build("sheets", "v4", credentials=credentials)


def existing_sheet_titles(service, sheet_id: str) -> set[str]:
    spreadsheet = (
        service.spreadsheets()
        .get(
            spreadsheetId=sheet_id,
            fields="sheets.properties.title",
        )
        .execute()
    )
    return {sheet["properties"]["title"] for sheet in spreadsheet.get("sheets", [])}


def ensure_sheets(service, sheet_id: str, sheet_names: list[str]) -> set[str]:
    existing_titles = existing_sheet_titles(service, sheet_id)
    created_titles = {
        sheet_name for sheet_name in sheet_names if sheet_name not in existing_titles
    }
    requests = [
        {"addSheet": {"properties": {"title": sheet_name}}}
        for sheet_name in created_titles
    ]
    if not requests:
        return created_titles

    service.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": requests},
    ).execute()
    return created_titles


def updated_row_count(result: dict) -> int:
    if "updatedRows" in result:
        return int(result.get("updatedRows", 0))
    return int(result.get("updates", {}).get("updatedRows", 0))


def push_rows(
    service,
    sheet_id: str,
    target_range: str,
    rows: list[list[str]],
    replace: bool,
) -> int:
    values_api = service.spreadsheets().values()
    body = {"values": rows}

    if replace:
        values_api.clear(spreadsheetId=sheet_id, range=target_range).execute()
        result = values_api.update(
            spreadsheetId=sheet_id,
            range=target_range,
            valueInputOption="USER_ENTERED",
            body=body,
        ).execute()
    else:
        result = values_api.append(
            spreadsheetId=sheet_id,
            range=target_range,
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body,
        ).execute()

    return updated_row_count(result)


def push_to_sheets_core(
    csv_file: str | Path = DEFAULT_INPUT_FILE,
    service_account_file: str | Path | None = None,
    sheet_id: str | None = None,
    range_columns_var: str = DEFAULT_SHEET_COLUMNS,
    replace: bool = False,
    include_header: bool = False,
) -> PushResult:
    """
    Core push function - uploads invoice items to Google Sheets.

    Returns a PushResult with rows updated, sheets modified, and replace flag.
    """
    if not service_account_file:
        raise SheetsPushError(
            "Missing service_account_file. Provide via parameter or GOOGLE_SERVICE_ACCOUNT_FILE env var."
        )
    if not sheet_id:
        raise SheetsPushError(
            "Missing sheet_id. Provide via parameter or GOOGLE_SHEET_ID env var."
        )

    csv_path = project_path(csv_file)
    service_account_path = project_path(service_account_file)

    header, data_rows = read_csv_rows(csv_path)
    if not data_rows:
        raise SheetsPushError(f"No data rows found in {csv_path}")

    grouped_rows = group_rows_by_month(header, data_rows)
    columns = range_columns(range_columns_var)
    service = sheets_service(service_account_path)
    created_sheets = ensure_sheets(service, sheet_id, list(grouped_rows))

    updated_rows = 0
    sheets_modified = []

    for sheet_name, month_rows in grouped_rows.items():
        rows = month_rows
        if replace or include_header or sheet_name in created_sheets:
            rows = [header] + month_rows
        target_range = month_range(sheet_name, columns)
        month_updated_rows = push_rows(
            service=service,
            sheet_id=sheet_id,
            target_range=target_range,
            rows=rows,
            replace=replace,
        )
        updated_rows += month_updated_rows
        sheets_modified.append(sheet_name)
        print(f"{sheet_name}: pushed {month_updated_rows} row(s)")

    return PushResult(
        rows_updated=updated_rows,
        sheets_modified=sheets_modified,
        replace=replace,
    )


def main() -> int:
    args = parse_args()

    try:
        result = push_to_sheets_core(
            csv_file=args.input,
            service_account_file=args.service_account_file,
            sheet_id=args.sheet_id,
            range_columns=args.range,
            replace=args.replace,
            include_header=args.include_header,
        )
    except SheetsPushError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    action = "Replaced" if result.replace else "Appended"
    print(
        f"{action} {result.rows_updated} total row(s) across {len(result.sheets_modified)} monthly sheet(s)."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

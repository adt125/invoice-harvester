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
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build


PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_FILE = PROJECT_DIR / "invoice_items.csv"
DEFAULT_SHEET_COLUMNS = "A:E"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]


class SheetsPushError(RuntimeError):
    pass


def parse_args() -> argparse.Namespace:
    load_dotenv(PROJECT_DIR / ".env")

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
    return PROJECT_DIR / path


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
        raise SheetsPushError(f"Service account JSON file does not exist: {service_account_file}")

    credentials = Credentials.from_service_account_file(
        service_account_file,
        scopes=SCOPES,
    )
    return build("sheets", "v4", credentials=credentials)


def existing_sheet_titles(service, sheet_id: str) -> set[str]:
    spreadsheet = service.spreadsheets().get(
        spreadsheetId=sheet_id,
        fields="sheets.properties.title",
    ).execute()
    return {
        sheet["properties"]["title"]
        for sheet in spreadsheet.get("sheets", [])
    }


def ensure_sheets(service, sheet_id: str, sheet_names: list[str]) -> set[str]:
    existing_titles = existing_sheet_titles(service, sheet_id)
    created_titles = {
        sheet_name
        for sheet_name in sheet_names
        if sheet_name not in existing_titles
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


def main() -> int:
    args = parse_args()

    try:
        if not args.service_account_file:
            raise SheetsPushError("Missing GOOGLE_SERVICE_ACCOUNT_FILE in .env or --service-account-file.")
        if not args.sheet_id:
            raise SheetsPushError("Missing GOOGLE_SHEET_ID in .env or --sheet-id.")

        csv_path = project_path(args.input)
        service_account_file = project_path(args.service_account_file)
        header, data_rows = read_csv_rows(csv_path)
        if not data_rows:
            raise SheetsPushError(f"No data rows found in {csv_path}")

        grouped_rows = group_rows_by_month(header, data_rows)
        columns = range_columns(args.range)
        service = sheets_service(service_account_file)
        created_sheets = ensure_sheets(service, args.sheet_id, list(grouped_rows))

        updated_rows = 0
        for sheet_name, month_rows in grouped_rows.items():
            rows = month_rows
            if args.replace or args.include_header or sheet_name in created_sheets:
                rows = [header] + month_rows
            target_range = month_range(sheet_name, columns)
            month_updated_rows = push_rows(
                service=service,
                sheet_id=args.sheet_id,
                target_range=target_range,
                rows=rows,
                replace=args.replace,
            )
            updated_rows += month_updated_rows
            print(f"{sheet_name}: pushed {month_updated_rows} row(s)")
    except SheetsPushError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    action = "Replaced" if args.replace else "Appended"
    print(f"{action} {updated_rows} total row(s) across {len(grouped_rows)} monthly sheet(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

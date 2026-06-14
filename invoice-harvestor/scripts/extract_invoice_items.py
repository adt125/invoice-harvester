#!/usr/bin/env python3
"""
Extract invoice line items from downloaded PDF attachments into a CSV file.

Expected invoice shape:
- A "Date of Invoice:" field near the top.
- A ruled item table with serial number, description, quantity, and total amount columns.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

SCRIPT_DIR = Path(__file__).resolve().parent  # Points to 'scripts/'
SKILL_ROOT = SCRIPT_DIR.parent
ASSETS_DIR = SKILL_ROOT / "assets"
TEMP_DIR = ASSETS_DIR / "temp"
DEFAULT_INPUT_DIR = TEMP_DIR
DEFAULT_OUTPUT_FILE = TEMP_DIR / "invoice_items.csv"


class ExtractError(RuntimeError):
    pass


@dataclass
class InvoiceItem:
    source_file: str
    invoice_date: str
    description: str
    quantity: str
    amount: str


@dataclass
class ExtractResult:
    """Result of extracting invoice items from PDFs."""

    count: int
    csv_path: Path
    items: list[InvoiceItem]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract invoice item date, description, quantity, and amount from PDF attachments."
    )
    parser.add_argument(
        "--input",
        default=str(DEFAULT_INPUT_DIR),
        help=f"PDF file or directory of PDFs. Default: {DEFAULT_INPUT_DIR}",
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_FILE),
        help=f"CSV output path. Default: {DEFAULT_OUTPUT_FILE}",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Search input directories recursively.",
    )
    return parser.parse_args()


def project_path(path_text: str) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return SCRIPT_DIR / path


def find_pdfs(input_path: Path, recursive: bool) -> list[Path]:
    if input_path.is_file():
        if input_path.suffix.lower() != ".pdf":
            raise ExtractError(f"Input file is not a PDF: {input_path}")
        return [input_path]

    if not input_path.exists():
        raise ExtractError(f"Input path does not exist: {input_path}")
    if not input_path.is_dir():
        raise ExtractError(f"Input path is neither a PDF nor a directory: {input_path}")

    pattern = "**/*.pdf" if recursive else "*.pdf"
    return sorted(input_path.glob(pattern))


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def clean_amount(value: str) -> str:
    value = clean_cell(value)
    value = value.replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", value)
    return match.group(0) if match else ""


def clean_quantity(value: str) -> str:
    value = clean_cell(value).replace(",", "")
    match = re.search(r"-?\d+(?:\.\d+)?", value)
    return match.group(0) if match else ""


def normalize_invoice_date(value: str) -> str:
    value = value.strip()
    for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y%m%d"):
        try:
            return datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            pass
    return value


def invoice_date_from_text(text: str) -> str:
    match = re.search(
        r"Date\s+of\s+Invoice\s*:?\s*(\d{1,2}[-/]\d{1,2}[-/]\d{4})",
        text,
        flags=re.IGNORECASE,
    )
    if match:
        return normalize_invoice_date(match.group(1))
    return ""


def invoice_date_from_filename(path: Path) -> str:
    match = re.match(r"(\d{8})_", path.name)
    if not match:
        return ""
    return normalize_invoice_date(match.group(1))


def is_line_item_row(row: list[str]) -> bool:
    if len(row) < 3:
        return False
    return bool(re.fullmatch(r"\d+\.?", row[0])) and bool(row[1])


def row_amount(row: list[str]) -> str:
    for cell in reversed(row):
        amount = clean_amount(cell)
        if amount:
            return amount
    return ""


def extract_items_from_tables(
    pdf_path: Path, invoice_date: str, tables: Iterable[list[list[Any]]]
) -> list[InvoiceItem]:
    items: list[InvoiceItem] = []
    for table in tables:
        for raw_row in table:
            row = [clean_cell(cell) for cell in raw_row]
            if not is_line_item_row(row):
                continue

            amount = row_amount(row)
            if not amount:
                continue

            items.append(
                InvoiceItem(
                    source_file=pdf_path.name,
                    invoice_date=invoice_date,
                    description=row[1],
                    quantity=clean_quantity(row[2]),
                    amount=amount,
                )
            )
    return items


def extract_items_from_pdf(pdf_path: Path) -> list[InvoiceItem]:
    import pdfplumber

    items: list[InvoiceItem] = []
    all_text: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            all_text.append(text)

        invoice_date = invoice_date_from_text("\n".join(all_text))
        if not invoice_date:
            invoice_date = invoice_date_from_filename(pdf_path)

        for page in pdf.pages:
            tables = page.extract_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_tolerance": 6,
                    "snap_tolerance": 4,
                    "join_tolerance": 4,
                }
            )
            items.extend(extract_items_from_tables(pdf_path, invoice_date, tables))

    return items


def write_csv(output_path: Path, items: list[InvoiceItem]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as output_file:
        writer = csv.DictWriter(
            output_file,
            fieldnames=[
                "source_file",
                "invoice_date",
                "description",
                "quantity",
                "amount",
            ],
        )
        writer.writeheader()
        for item in items:
            writer.writerow(
                {
                    "source_file": item.source_file,
                    "invoice_date": item.invoice_date,
                    "description": item.description,
                    "quantity": item.quantity,
                    "amount": item.amount,
                }
            )


def extract_items_core(
    input_dir: str | Path = DEFAULT_INPUT_DIR,
    output_file: str | Path = DEFAULT_OUTPUT_FILE,
    recursive: bool = False,
) -> ExtractResult:
    """
    Core extraction function - extracts invoice items from PDFs into CSV.

    Returns an ExtractResult with count, CSV path, and items list.
    """
    input_path = project_path(input_dir)
    output_path = project_path(output_file)

    try:
        pdfs = find_pdfs(input_path, recursive)
        if not pdfs:
            raise ExtractError(f"No PDF files found in {input_path}")

        items: list[InvoiceItem] = []
        for pdf_path in pdfs:
            extracted = extract_items_from_pdf(pdf_path)
            print(f"{pdf_path.name}: extracted {len(extracted)} item(s)")
            items.extend(extracted)

        write_csv(output_path, items)
        return ExtractResult(
            count=len(items),
            csv_path=output_path,
            items=items,
        )
    except ModuleNotFoundError as exc:
        if exc.name == "pdfplumber":
            raise ExtractError(
                "Missing dependency pdfplumber. Run: python3 -m pip install -r requirements.txt"
            ) from exc
        raise


def main() -> int:
    args = parse_args()
    try:
        result = extract_items_core(
            input_dir=args.input,
            output_file=args.output,
            recursive=args.recursive,
        )
    except ExtractError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print(f"Done. Wrote {result.count} item(s) to {result.csv_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

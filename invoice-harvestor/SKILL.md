---
name: invoice-harvestor
description: Download invoice PDF attachments from Outlook, extract line items into CSV, and push the results to Google Sheets. Use when agent needs to run or maintain the invoice harvester workflow, work with Outlook invoice attachments, parse invoice PDFs with date/description/quantity/amount fields, create invoice_items.csv, or upload invoice rows into month-based Google Sheets tabs or user asks to feed the monthly instamart expense report to sheets.
---

# Invoice Harvester

Use this skill to operate or maintain the bundled Outlook -> PDF extraction -> Google Sheets invoice pipeline.

## Resources

- `scripts/pipeline.py`: preferred end-to-end entry point.
- `scripts/download_outlook_attachments.py`: downloads file attachments from a Microsoft Outlook mail folder using Microsoft Graph device-code auth.
- `scripts/extract_invoice_items.py`: extracts invoice line items from PDF tables into CSV.
- `scripts/push_to_google_sheets.py`: uploads CSV rows to Google Sheets, grouped by invoice month.
- `scripts/clean_up.py`: removes `assets/temp` after a successful pipeline run.
- `scripts/requirements.txt`: Python dependencies.
- `assets/.env.example`: environment variable template.

## Setup

Install dependencies from the skill scripts directory:

```bash
python3 -m venv .venv &&
source .venv/bin/activate &&
python3 -m pip install -r invoice-harvestor/scripts/requirements.txt
```

## End-to-End Run

Prefer the pipeline for normal use:

```bash
python3 invoice-harvestor/scripts/pipeline.py --folder "Instamart" --last-month --replace
```

Useful flags:

- `--folder`: Outlook folder display name or path, such as `Instamart` or `Inbox/Invoices`.
- `--last-month`: filter Outlook messages to the previous calendar month.
- `--replace`: clear each monthly Google Sheet tab before writing.
- `--dry-run`: authenticate and preview downloads without writing attachments or pushing rows.
- `--attachments-dir`: override the downloaded PDF directory.
- `--csv-file`: override the generated CSV path.
- `--client-id`, `--tenant-id`, `--service-account`, `--sheet-id`: override environment values.

The pipeline writes temporary PDFs and `invoice_items.csv` under `assets/temp`, pushes data to Sheets, then removes `assets/temp`.

## Individual Steps

Download attachments only:

```bash
python3 invoice-harvestor/scripts/download_outlook_attachments.py --folder "Instamart" --last-month
```

Add filters when needed:

```bash
python3 invoice-harvestor/scripts/download_outlook_attachments.py --folder "Inbox/Invoices" --from-date 2026-05-01 --to-date 2026-05-31 --subject-contains invoice --limit 25
```

Extract invoice items from PDFs:

```bash
python3 invoice-harvestor/scripts/extract_invoice_items.py --input invoice-harvestor/assets/temp --output invoice-harvestor/assets/temp/invoice_items.csv
```

Push CSV rows to Google Sheets:

```bash
python3 invoice-harvestor/scripts/push_to_google_sheets.py --input invoice-harvestor/assets/temp/invoice_items.csv --replace
```

## Data Shape

The extractor expects invoice PDFs with:

- a `Date of Invoice:` field near the top, or an attachment filename prefixed with `YYYYMMDD_`;
- ruled tables that `pdfplumber` can detect using line-based table extraction;
- line-item rows where column 1 is a serial number, column 2 is the description, column 3 is quantity, and the rightmost numeric cell is amount.

The generated CSV columns are:

```text
source_file,invoice_date,description,quantity,amount
```

The Sheets uploader groups rows by `invoice_date` month and writes each group to a tab named `YYYY-MM`. Missing tabs are created automatically. In append mode, headers are included only for newly created sheets or when `--include-header` is used; in replace mode, headers are always written.

## Authentication Notes

Outlook uses Microsoft Graph delegated auth with device-code flow. The first run prints a code and login URL, then stores the token cache at:

```text
~/.outlook-attachment-downloader-token-cache.json
```

Google Sheets uses a service account. Ensure the target spreadsheet is shared with the service account email from the JSON credentials.

## Maintenance Notes

- Keep pipeline behavior centered on `scripts/pipeline.py`; use individual scripts for debugging or partial reruns.
- Treat paths relative to the script directory unless the code explicitly accepts absolute paths.
- Preserve the monthly sheet grouping contract when changing CSV or Sheets logic.
- Update this file whenever command flags, environment variables, output locations, or CSV columns change.

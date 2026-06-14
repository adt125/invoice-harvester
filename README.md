# Outlook Attachment Downloader

This script downloads file attachments from messages in a specific Outlook folder using Microsoft Graph.
By default, downloaded attachments are saved into the project-local `attachments/` folder.

## Setup

1. Create a Microsoft Entra app registration.
2. In the app registration, enable public client/device-code authentication for mobile and desktop flows.
3. Add delegated Microsoft Graph permissions:
   - `User.Read`
   - `Mail.Read`
4. Copy the app/client ID.
5. Create a `.env` file in this project:

```bash
cp .env.example .env
```

Then set:

```text
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_TENANT_ID=consumers
```

6. Install dependencies:

```bash
python3 -m pip install -r requirements.txt
```

## Usage

```bash
python3 download_outlook_attachments.py --folder "Invoices"
```

For a nested folder:

```bash
python3 download_outlook_attachments.py --folder "Inbox/Invoices"
```

Useful filters:

```bash
python3 download_outlook_attachments.py \
  --folder "Invoices" \
  --from-date 2026-05-01 \
  --to-date 2026-05-31 \
  --subject-contains "receipt"
```

Previous calendar month:

```bash
python3 download_outlook_attachments.py --folder "Invoices" --last-month
```

For example, if today is in June, `--last-month` uses May 1 through May 31.

Downloaded files are prefixed with the message received date in `yyyyMMdd` format:

```text
attachments/20260501_invoice.pdf
```

`--since YYYY-MM-DD` is still supported as an alias for `--from-date`.

To save somewhere else:

```bash
python3 download_outlook_attachments.py --folder "Invoices" --output downloads
```

Preview without downloading:

```bash
python3 download_outlook_attachments.py --folder "Invoices" --dry-run
```

## Extract Invoice Items

After PDFs are downloaded into `attachments/`, extract line items into CSV:

```bash
python3 extract_invoice_items.py
```

Default output:

```text
invoice_items.csv
```

The CSV contains:

```text
source_file,invoice_date,description,quantity,amount
```

You can also choose paths:

```bash
python3 extract_invoice_items.py --input attachments --output invoice_items.csv
```

## Push To Google Sheets

To send `invoice_items.csv` to Google Sheets:

1. Create a Google Cloud project.
2. Enable the Google Sheets API.
3. Create a service account.
4. Download the service account JSON key into this project, for example:

```text
google-service-account.json
```

5. Open your target Google Sheet and share it with the service account email.
   The email is inside the JSON file as `client_email`.
6. Add these values to `.env`:

```text
GOOGLE_SERVICE_ACCOUNT_FILE=google-service-account.json
GOOGLE_SHEET_ID=your-google-sheet-id
GOOGLE_SHEET_RANGE=A:E
```

Append CSV rows to monthly sheets:

```bash
python3 push_to_google_sheets.py
```

Rows are grouped by `invoice_date`. For example, rows dated `2026-05-01` through `2026-05-31` go to a sheet tab named:

```text
2026-05
```

Missing monthly sheet tabs are created automatically.

Replace each monthly sheet range with the latest CSV:

```bash
python3 push_to_google_sheets.py --replace
```

Full monthly workflow:

```bash
python3 download_outlook_attachments.py --folder "Invoices" --last-month
python3 extract_invoice_items.py
python3 push_to_google_sheets.py --replace
```

The first run prints a device-code login prompt. After login, the token cache is stored at:

```text
~/.outlook-attachment-downloader-token-cache.json
```

## Notes

- The script downloads only regular file attachments.
- It skips embedded item attachments such as attached emails.
- If a file already exists, the script adds a numeric suffix unless `--overwrite` is used.
- If your app registration is configured for Microsoft personal accounts only, set `OUTLOOK_TENANT_ID=consumers`.
- If your app registration is configured for work/school accounts, use `OUTLOOK_TENANT_ID=common`, `organizations`, or your tenant ID, depending on your Azure setup.
- You can set `OUTLOOK_CLIENT_ID` and `OUTLOOK_TENANT_ID` in `.env`; command-line flags override `.env`.

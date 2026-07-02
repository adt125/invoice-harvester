# Invoice Harvester

Download invoice attachments from Outlook, extract invoice line items into CSV,
tag each item with Gemini, and push the tagged rows to Google Sheets.

## Pipeline Flow

Running `scripts/pipeline.py` performs:

1. Download invoice PDF attachments from an Outlook folder.
2. Extract invoice items into `assets/temp/invoice_items.csv`.
3. Tag each row and write `assets/temp/tagged_expenses.csv`.
4. Push the tagged CSV to Google Sheets.
5. Clean up temporary files.

Item tags are cached in `assets/tag_cache.json`, so repeated descriptions in
future monthly runs do not call the LLM again.

## Setup

Create and activate a virtual environment:

```bash
cd invoice-harvestor
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Create `.env` from the example:

```bash
cp .env.example .env
```

Fill in these values in `.env`:

```bash
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_TENANT_ID=common

GOOGLE_SERVICE_ACCOUNT_FILE=google-service-account.json
GOOGLE_SHEET_ID=your-google-sheet-id
GOOGLE_SHEET_RANGE=A:F

GOOGLE_GENAI_USE_VERTEXAI=0
GOOGLE_API_KEY=your-gemini-api-key
```

The Google service account JSON path is resolved relative to `assets/` when it
is not absolute, so `GOOGLE_SERVICE_ACCOUNT_FILE=google-service-account.json`
expects `assets/google-service-account.json`.

## Run The Pipeline

Default run:

```bash
python3 scripts/pipeline.py
```

Common run for the previous calendar month:

```bash
python3 scripts/pipeline.py --last-month --replace
```

Use a specific Outlook folder:

```bash
python3 scripts/pipeline.py --folder "Instamart" --last-month --replace
```

Preview the Outlook download step without writing attachments:

```bash
python3 scripts/pipeline.py --dry-run
```

## Useful Options

```bash
python3 scripts/pipeline.py --folder "Instamart" --last-month --replace
```

`--replace` clears each monthly Google Sheet tab before writing. Without it,
rows are appended.

`--tagged-csv-file` controls the CSV that is pushed to Google Sheets after
tagging. The default is `assets/temp/tagged_expenses.csv`.

`--tag-cache-file` controls the persistent description-to-tag cache. The
default is `assets/tag_cache.json`.

For all flags:

```bash
python3 scripts/pipeline.py --help
```

# Invoice Harvester

Turn grocery-delivery invoices into an itemized, searchable monthly expense
ledger in Google Sheets.

## Why use it?

When you order groceries or household supplies through services such as
Instamart or Blinkit, the invoice often arrives as an email attachment. Keeping
track of every item—and working out which costs to include in a shared
household expense—can mean reopening orders one by one and copying totals into
a spreadsheet.

Invoice Harvester automates the repetitive part. It reads invoice attachments
from an Outlook folder, extracts the purchased items, adds an AI-generated
category to each item, and organizes the results in Google Sheets. You can then
filter or total the sheet by month and category when reviewing expenses with
roommates or flatmates.

The sheet gives you an item-level record to use for expense reviews; it does
not automatically decide who bought, used, or owes a share of each item.

## What it does

- Downloads file attachments from a selected Outlook mail folder using
  Microsoft Graph.
- Extracts invoice date, item description, quantity, and amount from supported
  PDF invoices.
- Uses Gemini to assign one category to each item. The default categories
  are vegetables, fruits, masala, yogurt (curd), milk, paneer, utility, snacks,
  groceries, dining, aata, rice, and unknown.
- Caches description-to-category results locally to avoid repeat Gemini calls
  for previously seen items.
- Writes itemized rows to a Google Sheets tab for each invoice month
  (`YYYY-MM`), with columns for source file, invoice date, description,
  quantity, amount, and tag.
- Appends rows by default, or replaces the data in each affected monthly tab
  when run with `--replace`.

## How the pipeline works

Running `scripts/pipeline.py` performs these steps:

1. Authenticates to Outlook and downloads attachments from the selected folder.
2. Finds supported PDFs and extracts their invoice line items to
   `assets/temp/invoice_items.csv`.
3. Tags each item with Gemini and writes
   `assets/temp/tagged_expenses.csv`.
4. Groups rows by invoice month and sends them to Google Sheets.
5. Removes temporary files after a run that reaches the Sheets upload step.

The persistent category cache is stored at `assets/tag_cache.json`.

## Requirements

- Python 3.10 or newer
- An Outlook/Microsoft Entra app registration with delegated `User.Read` and
  `Mail.Read` access and device-code sign-in enabled
- A Gemini API key
- A Google Cloud service account with the Google Sheets API enabled
- A Google spreadsheet shared with the service account's email address as an
  editor

## Setup

Create a virtual environment and install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Create your local configuration file:

```bash
cp .env.example .env
```

Edit `.env` and set your values:

```dotenv
OUTLOOK_CLIENT_ID=your-app-client-id
OUTLOOK_TENANT_ID=consumers

GOOGLE_SERVICE_ACCOUNT_FILE=google-service-account.json
GOOGLE_SHEET_ID=your-google-sheet-id

GOOGLE_GENAI_USE_VERTEXAI=0
GOOGLE_API_KEY=your-gemini-api-key
```

Place the Google service-account JSON file in `assets/` when using the example
relative path. Alternatively, set `GOOGLE_SERVICE_ACCOUNT_FILE` to an absolute
path. For `GOOGLE_SHEET_ID`, use the ID from the spreadsheet URL.

On the first Outlook sign-in, follow the device-code instructions printed in
the terminal. Keep `.env` and the service-account JSON private; never commit
credentials to the repository.

## Run

Process invoices from the default Outlook folder (`Instamart`):

```bash
python3 scripts/pipeline.py
```

Process the previous calendar month and replace the data in the corresponding
monthly sheet tab:

```bash
python3 scripts/pipeline.py --folder "Instamart" --last-month --replace
```

Use another folder or a nested Outlook folder:

```bash
python3 scripts/pipeline.py --folder "Inbox/Grocery invoices" --last-month --replace
```

Preview which attachments match without saving them or sending data to Sheets:

```bash
python3 scripts/pipeline.py --folder "Instamart" --last-month --dry-run
```

`--replace` clears the affected monthly tab data before writing the rows found
in the current run. Use it only when the selected folder and date range contain
the complete set of invoices you want to keep for those months. Without
`--replace`, each run appends rows, so repeating a run can create duplicates.

See all available options with:

```bash
python3 scripts/pipeline.py --help
```

## Customize categories

The default category list and tagging instructions are in
`scripts/item_llm_tagging.py`. Edit the prompt there to match your household's
categories—for example, add or rename tags such as `cleaning`, `shared`, or
`personal`. Previously tagged descriptions use the local cache, so remove or
update the corresponding entry in `assets/tag_cache.json` if you want an
existing item to be categorized again with the changed instructions.

## Notes and limitations

- PDF extraction is designed for invoices with a readable invoice date and a
  ruled item table. Other invoice layouts may need changes to the extractor in
  `scripts/extract_invoice_items.py`.
- The downloader retrieves file attachments; the extraction step processes
  supported PDFs. It does not read invoices from other email providers.
- AI-generated categories can be incorrect. Review the sheet and adjust tags
  when needed.
- Item descriptions are sent to Gemini for classification. Consider your
  privacy requirements before processing invoices.
- This project records and categorizes expenses; roommate ownership,
  contribution splits, and settlements are managed outside the pipeline.

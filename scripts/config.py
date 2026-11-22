"""Shared configuration for invoice harvester scripts."""

from __future__ import annotations

from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
ROOT = SCRIPT_DIR.parent
ASSETS_DIR = ROOT / "assets"
TEMP_DIR = ASSETS_DIR / "temp"

DEFAULT_ATTACHMENTS_DIR = TEMP_DIR
DEFAULT_INVOICE_ITEMS_CSV = TEMP_DIR / "invoice_items.csv"
DEFAULT_TAGGED_EXPENSES_CSV = TEMP_DIR / "tagged_expenses.csv"
DEFAULT_TAG_CACHE_FILE = ASSETS_DIR / "tag_cache.json"
DEFAULT_OUTLOOK_FOLDER = "Instamart"
DEFAULT_OUTLOOK_TENANT_ID = "common"
DEFAULT_GOOGLE_SHEET_COLUMNS = "A:K"


def load_project_env() -> None:
    """Load the project's .env file."""
    try:
        from dotenv import load_dotenv

        load_dotenv(dotenv_path=ROOT / ".env")
    except ModuleNotFoundError:
        return


def resolve_from_scripts(path_text: str | Path) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return SCRIPT_DIR / path


def resolve_from_assets(path_text: str | Path) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return ASSETS_DIR / path


def resolve_from_temp(path_text: str | Path) -> Path:
    path = Path(path_text)
    if path.is_absolute():
        return path
    return TEMP_DIR / path

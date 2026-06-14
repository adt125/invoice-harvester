#!/usr/bin/env python3
"""
Download file attachments from messages in an Outlook mail folder.

Authentication uses Microsoft Graph delegated auth with a device-code flow.
Create a Microsoft Entra app registration and pass its client ID with
--client-id or OUTLOOK_CLIENT_ID.
"""

from __future__ import annotations

import argparse
import base64
import json
import os
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import quote


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
DEFAULT_SCOPES = ["User.Read", "Mail.Read"]
PROJECT_DIR = Path(__file__).resolve().parent
DEFAULT_OUTPUT_DIR = PROJECT_DIR / "attachments"
TOKEN_CACHE_FILE = Path.home() / ".outlook-attachment-downloader-token-cache.json"


class GraphError(RuntimeError):
    pass


def graph_quote(value: str) -> str:
    return quote(value, safe="")


def parse_args() -> argparse.Namespace:
    from dotenv import load_dotenv

    load_dotenv(PROJECT_DIR / ".env")
    parser = argparse.ArgumentParser(
        description="Download Outlook attachments from a specific mail folder."
    )
    parser.add_argument(
        "--folder",
        required=True,
        help='Folder display name or path, for example "Invoices" or "Inbox/Invoices".',
    )
    parser.add_argument(
        "--output",
        default=str(DEFAULT_OUTPUT_DIR),
        help=f"Directory where attachments are saved. Default: {DEFAULT_OUTPUT_DIR}",
    )
    parser.add_argument(
        "--client-id",
        default=os.environ.get("OUTLOOK_CLIENT_ID"),
        help="Microsoft Entra app registration client ID. Can also use OUTLOOK_CLIENT_ID in .env.",
    )
    parser.add_argument(
        "--tenant-id",
        default=os.environ.get("OUTLOOK_TENANT_ID", "common"),
        help='Tenant ID, or "common" for work/school + personal Microsoft accounts. Can also use OUTLOOK_TENANT_ID in .env. Default: common',
    )
    parser.add_argument(
        "--since",
        dest="from_date",
        help="Alias for --from-date. Include messages received on/after this date, in YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--from-date",
        help="Include messages received on/after this date, in YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--to-date",
        help="Include messages received on/before this date, in YYYY-MM-DD format.",
    )
    parser.add_argument(
        "--last-month",
        action="store_true",
        help="Use the previous calendar month as the date range.",
    )
    parser.add_argument(
        "--subject-contains",
        help="Only include messages whose subject contains this text.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Maximum number of messages to inspect after filters are applied.",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite files with the same name. By default, a numeric suffix is added.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print what would be downloaded without writing files.",
    )
    return parser.parse_args()


def parse_date(value: str, option_name: str) -> date:
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError as exc:
        raise GraphError(f"{option_name} must use YYYY-MM-DD format, for example 2026-05-01.") from exc


def previous_calendar_month(today: date | None = None) -> tuple[date, date]:
    today = today or date.today()
    first_this_month = today.replace(day=1)
    last_previous_month = first_this_month - timedelta(days=1)
    first_previous_month = last_previous_month.replace(day=1)
    return first_previous_month, last_previous_month


def effective_date_range(args: argparse.Namespace) -> tuple[date | None, date | None]:
    if args.last_month and (args.from_date or args.to_date):
        raise GraphError("Use either --last-month or manual date flags, not both.")
    if args.last_month:
        return previous_calendar_month()

    from_date = parse_date(args.from_date, "--from-date") if args.from_date else None
    to_date = parse_date(args.to_date, "--to-date") if args.to_date else None
    if from_date and to_date and from_date > to_date:
        raise GraphError("--from-date cannot be later than --to-date.")
    return from_date, to_date


def received_date_prefix(received: str) -> str:
    if not received:
        return "unknown-date"
    try:
        return datetime.fromisoformat(received.replace("Z", "+00:00")).strftime("%Y%m%d")
    except ValueError:
        return received[:10].replace("-", "") or "unknown-date"


def load_token_cache() -> Any:
    import msal

    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())
    return cache


def save_token_cache(cache: Any) -> None:
    if cache.has_state_changed:
        TOKEN_CACHE_FILE.write_text(cache.serialize())
        TOKEN_CACHE_FILE.chmod(0o600)


def get_access_token(client_id: str, tenant_id: str) -> str:
    import msal

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    cache = load_token_cache()
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=cache,
    )

    result: dict[str, Any] | None = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(DEFAULT_SCOPES, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=DEFAULT_SCOPES)
        if "user_code" not in flow:
            raise GraphError(f"Could not create device-code flow: {flow}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    save_token_cache(cache)

    if "access_token" not in result:
        raise GraphError(
            "Authentication failed: "
            + json.dumps(
                {
                    "error": result.get("error"),
                    "error_description": result.get("error_description"),
                },
                indent=2,
            )
        )
    return result["access_token"]


def graph_get(token: str, url: str, params: dict[str, str] | None = None) -> dict[str, Any]:
    import requests

    headers = {"Authorization": f"Bearer {token}"}
    response = requests.get(url, headers=headers, params=params, timeout=60)
    if not response.ok:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise GraphError(f"Graph request failed ({response.status_code}): {detail}")
    return response.json()


def graph_pages(
    token: str, url: str, params: dict[str, str] | None = None
) -> Iterable[dict[str, Any]]:
    next_url: str | None = url
    next_params = params
    while next_url:
        page = graph_get(token, next_url, next_params)
        yield from page.get("value", [])
        next_url = page.get("@odata.nextLink")
        next_params = None


def find_child_folder(token: str, parent_url: str, name: str) -> dict[str, Any] | None:
    params = {
        "$top": "100",
        "$select": "id,displayName,parentFolderId",
    }
    for folder in graph_pages(token, parent_url, params):
        if folder.get("displayName", "").lower() == name.lower():
            return folder
    return None


def resolve_folder_id(token: str, folder_path: str) -> str:
    parts = [part.strip() for part in folder_path.split("/") if part.strip()]
    if not parts:
        raise GraphError("Folder path is empty.")

    first = find_child_folder(token, f"{GRAPH_ROOT}/me/mailFolders", parts[0])
    if not first:
        raise GraphError(f'Could not find top-level mail folder "{parts[0]}".')

    folder = first
    for child_name in parts[1:]:
        child = find_child_folder(
            token,
            f"{GRAPH_ROOT}/me/mailFolders/{graph_quote(folder['id'])}/childFolders",
            child_name,
        )
        if not child:
            raise GraphError(
                f'Could not find child folder "{child_name}" under "{folder["displayName"]}".'
            )
        folder = child

    return folder["id"]


def sanitize_filename(name: str) -> str:
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip()
    name = re.sub(r"\s+", " ", name)
    return name[:180] or "attachment"


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 1
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def build_message_filter(args: argparse.Namespace) -> str:
    filters = ["hasAttachments eq true"]
    from_date, to_date = effective_date_range(args)
    if from_date:
        filters.append(f"receivedDateTime ge {from_date.isoformat()}T00:00:00Z")
    if to_date:
        to_exclusive = to_date + timedelta(days=1)
        filters.append(f"receivedDateTime lt {to_exclusive.isoformat()}T00:00:00Z")
    return " and ".join(filters)


def message_matches(args: argparse.Namespace, message: dict[str, Any]) -> bool:
    if args.subject_contains:
        subject = message.get("subject") or ""
        return args.subject_contains.lower() in subject.lower()
    return True


def download_attachments(args: argparse.Namespace) -> int:
    if not args.client_id:
        raise GraphError("Missing --client-id or OUTLOOK_CLIENT_ID in .env.")

    token = get_access_token(args.client_id, args.tenant_id)
    folder_id = resolve_folder_id(token, args.folder)
    output_dir = Path(args.output)
    if not output_dir.is_absolute():
        output_dir = PROJECT_DIR / output_dir

    if not args.dry_run:
        output_dir.mkdir(parents=True, exist_ok=True)

    messages_url = f"{GRAPH_ROOT}/me/mailFolders/{graph_quote(folder_id)}/messages"
    message_params = {
        "$top": "50",
        "$select": "id,subject,receivedDateTime,from,hasAttachments",
        "$filter": build_message_filter(args),
    }

    downloaded = 0
    inspected = 0

    for message in graph_pages(token, messages_url, message_params):
        if not message_matches(args, message):
            continue

        inspected += 1
        if args.limit and inspected > args.limit:
            break

        subject = message.get("subject") or "(no subject)"
        received = message.get("receivedDateTime") or "unknown-date"
        attachments_url = f"{GRAPH_ROOT}/me/messages/{graph_quote(message['id'])}/attachments"

        for attachment in graph_pages(token, attachments_url, {"$top": "50"}):
            attachment_type = attachment.get("@odata.type", "")
            if attachment_type != "#microsoft.graph.fileAttachment":
                print(f"Skipping non-file attachment on {received}: {subject}")
                continue

            attachment_id = attachment["id"]
            detail_url = (
                f"{GRAPH_ROOT}/me/messages/{graph_quote(message['id'])}"
                f"/attachments/{graph_quote(attachment_id)}"
            )
            detail = graph_get(token, detail_url)
            content = detail.get("contentBytes")
            if not content:
                print(f"Skipping attachment without contentBytes: {detail.get('name')}")
                continue

            filename = sanitize_filename(
                f"{received_date_prefix(received)}_{detail.get('name') or 'attachment'}"
            )
            target = output_dir / filename
            if not args.overwrite:
                target = unique_path(target)

            print(f"{'Would save' if args.dry_run else 'Saving'} {target} from {received}: {subject}")
            if not args.dry_run:
                target.write_bytes(base64.b64decode(content))
            downloaded += 1

    return downloaded


def main() -> int:
    args = parse_args()
    try:
        count = download_attachments(args)
    except GraphError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    print(f"Done. {'Matched' if args.dry_run else 'Downloaded'} {count} file attachment(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

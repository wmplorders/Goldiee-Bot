from __future__ import annotations

import base64
import re
import time
from datetime import date, datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any

from .attachment_rules import select_invoice_excels, select_pdfs
from .purchase_processor import build_summary, default_output_filename
from .state import append_log, load_state, save_state


SCOPES = ["https://www.googleapis.com/auth/gmail.modify"]


def get_gmail_service(project_root: Path):
    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Gmail libraries are not installed. Run setup.bat first, then place "
            "OAuth credentials at data/gmail_credentials.json."
        ) from exc

    token_path = project_root / "data" / "token.json"
    credentials_path = project_root / "data" / "gmail_credentials.json"
    if not credentials_path.exists():
        raise FileNotFoundError(
            f"Missing Gmail OAuth credentials: {credentials_path}. "
            "Download an OAuth desktop-client JSON from Google Cloud and save it there."
        )

    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
            creds = flow.run_local_server(port=0)
        token_path.write_text(creds.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=creds)


def monitor_gmail(config: dict[str, Any], project_root: Path, send_to_whatsapp) -> None:
    service = get_gmail_service(project_root)
    label_id = _ensure_label(service, config["gmail"]["processed_label"])
    while True:
        process_new_messages(service, label_id, config, send_to_whatsapp)
        time.sleep(int(config.get("poll_seconds", 120)))


def process_new_messages(service, label_id: str, config: dict[str, Any], send_to_whatsapp) -> None:
    state_path = config["paths"]["state"]
    state = load_state(state_path)
    processed = set(state.get("processed_message_ids", []))
    skipped = set(state.get("skipped_message_ids", []))

    start_date = datetime.strptime(config["gmail"]["start_date"], "%Y-%m-%d").date()
    query = (
        f'from:{config["gmail"]["sender"]} '
        f'after:{start_date.strftime("%Y/%m/%d")} '
        f'-label:"{config["gmail"]["processed_label"]}"'
    )
    response = service.users().messages().list(userId="me", q=query, maxResults=25).execute()
    messages = response.get("messages", [])

    for message_ref in messages:
        message_id = message_ref["id"]
        if message_id in processed or message_id in skipped:
            continue

        message = service.users().messages().get(userId="me", id=message_id).execute()
        metadata = _message_metadata(message)
        email_date = metadata["date"].date()
        age_days = (date.today() - email_date).days

        if age_days >= int(config["gmail"]["max_email_age_days"]):
            note = f"Found email dated {email_date.isoformat()}, not processed because it is {age_days} days old."
            send_to_whatsapp(note, attachments=[])
            append_log(config["paths"]["processing_log"], {
                "status": "skipped_old_email",
                "source": "gmail",
                "message_id": message_id,
                "email_date": email_date.isoformat(),
                "sender": metadata["from"],
                "subject": metadata["subject"],
                "note": note,
            })
            skipped.add(message_id)
            continue

        attachment_paths = _download_attachments(service, message_id, message, Path(config["paths"]["downloads"]))
        pdfs = select_pdfs(attachment_paths)
        invoice_excels = select_invoice_excels(attachment_paths)

        outputs = []
        invoice_numbers = []
        new_items = set()
        warnings = set()
        for excel_path in invoice_excels:
            result = build_summary(excel_path, Path(config["paths"]["processed"]) / default_output_filename(excel_path))
            outputs.append(Path(result["output"]))
            if result["invoice_no"]:
                invoice_numbers.append(result["invoice_no"])
            new_items.update(result["new_items"])
            if result["item_master_warning"]:
                warnings.add(result["item_master_warning"])
            append_log(config["paths"]["processing_log"], {
                "status": "processed",
                "source": "gmail",
                "message_id": message_id,
                "email_date": email_date.isoformat(),
                "sender": metadata["from"],
                "subject": metadata["subject"],
                **result,
                "output_file": result["output"],
            })

        service.users().messages().modify(
            userId="me",
            id=message_id,
            body={"addLabelIds": [label_id]},
        ).execute()
        processed.add(message_id)

        attachments = pdfs + outputs
        if attachments:
            invoice_text = ", ".join(invoice_numbers) or metadata["subject"]
            whatsapp_message = f"Purchase summary created for invoice {invoice_text}."
            if new_items:
                whatsapp_message += "\n\nNew items found:\n- " + "\n- ".join(sorted(new_items))
            if warnings:
                whatsapp_message += "\n\nWarning: " + " ".join(sorted(warnings))
            send_to_whatsapp(
                whatsapp_message,
                attachments=attachments,
            )

    state["processed_message_ids"] = sorted(processed)
    state["skipped_message_ids"] = sorted(skipped)
    save_state(state_path, state)


def _message_metadata(message: dict[str, Any]) -> dict[str, Any]:
    headers = {h["name"].casefold(): h["value"] for h in message["payload"].get("headers", [])}
    parsed_date = parsedate_to_datetime(headers.get("date", ""))
    if parsed_date.tzinfo:
        parsed_date = parsed_date.astimezone().replace(tzinfo=None)
    return {
        "from": headers.get("from", ""),
        "subject": headers.get("subject", ""),
        "date": parsed_date,
    }


def _download_attachments(service, message_id: str, message: dict[str, Any], download_dir: Path) -> list[Path]:
    download_dir.mkdir(parents=True, exist_ok=True)
    paths: list[Path] = []
    for part in _walk_parts(message.get("payload", {})):
        filename = part.get("filename")
        body = part.get("body", {})
        attachment_id = body.get("attachmentId")
        if not filename or not attachment_id:
            continue
        attachment = service.users().messages().attachments().get(
            userId="me",
            messageId=message_id,
            id=attachment_id,
        ).execute()
        data = base64.urlsafe_b64decode(attachment["data"].encode("utf-8"))
        safe_name = re.sub(r'[<>:"/\\|?*]+', "-", filename)
        path = download_dir / f"{message_id}_{safe_name}"
        path.write_bytes(data)
        paths.append(path)
    return paths


def _walk_parts(part: dict[str, Any]):
    yield part
    for child in part.get("parts", []) or []:
        yield from _walk_parts(child)


def _ensure_label(service, label_name: str) -> str:
    labels = service.users().labels().list(userId="me").execute().get("labels", [])
    for label in labels:
        if label["name"] == label_name:
            return label["id"]
    created = service.users().labels().create(
        userId="me",
        body={"name": label_name, "labelListVisibility": "labelShow", "messageListVisibility": "show"},
    ).execute()
    return created["id"]

from __future__ import annotations

import csv
import json
from datetime import datetime
from pathlib import Path
from typing import Any


def load_state(path: str | Path) -> dict[str, Any]:
    path = Path(path)
    if not path.exists():
        return {
            "processed_message_ids": [],
            "skipped_message_ids": [],
            "processed_whatsapp_message_ids": [],
        }
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def save_state(path: str | Path, state: dict[str, Any]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        json.dump(state, handle, indent=2)


def append_log(path: str | Path, row: dict[str, Any]) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "processed_at",
        "status",
        "source",
        "message_id",
        "email_date",
        "sender",
        "subject",
        "invoice_no",
        "invoice_date",
        "input_rows",
        "output_rows",
        "consolidated_rows",
        "output_file",
        "note",
    ]
    exists = path.exists()
    with path.open("a", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        if not exists:
            writer.writeheader()
        payload = {key: row.get(key, "") for key in fieldnames}
        payload["processed_at"] = payload["processed_at"] or datetime.now().isoformat(timespec="seconds")
        writer.writerow(payload)

from __future__ import annotations

import hashlib
import html
import json
import re
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date, datetime
from pathlib import Path
from typing import Any

from .settings import PROJECT_ROOT
from .settings import load_config


CACHE_DIR = PROJECT_ROOT / "data" / "item_master_cache"
TALLY_URL = "http://localhost:9000"


def get_item_master(force_refresh: bool = False, company_name: str | None = None) -> dict[str, Any]:
    company_name = company_name or load_config().get("tally", {}).get("company_name", "")
    if not company_name:
        raise RuntimeError("No Tally company_name is configured.")
    try:
        items = fetch_item_names(company_name)
        today_cache = CACHE_DIR / _cache_filename(date.today().isoformat(), company_name, items)
        if today_cache.exists() and not force_refresh:
            return _read_cache(today_cache, source="today_cache")
        return _save_cache(items, company_name)
    except Exception as exc:
        latest = _latest_cache(company_name=company_name)
        if not latest:
            raise RuntimeError(
                f"Could not fetch Tally Stock Item Master for company {company_name!r} "
                f"and no matching cache exists: {exc}"
            ) from exc
        result = _read_cache(latest, source="older_cache")
        result["warning"] = f"Tally unavailable; using cached item master from {result['fetched_at']}."
        return result


def fetch_item_names(company_name: str) -> list[str]:
    request_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Goldiee Bot Stock Items</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>"
        f"<SVCurrentCompany>{html.escape(company_name)}</SVCurrentCompany></STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"Goldiee Bot Stock Items\">"
        "<TYPE>Stock Item</TYPE><FETCH>Name</FETCH>"
        "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    ).encode("utf-8")
    request = urllib.request.Request(TALLY_URL, data=request_xml, headers={"Content-Type": "text/xml"})
    response = urllib.request.urlopen(request, timeout=30).read().decode("utf-8", errors="ignore")
    response = re.sub(r"&#(?:[0-8]|1[0-9]|2[0-9]|3[0-1]);", "", response)
    root = ET.fromstring(response)
    names = {
        node.attrib.get("NAME", "").strip()
        for node in root.findall(".//STOCKITEM")
        if node.attrib.get("NAME", "").strip()
    }
    if not names:
        raise RuntimeError(f"Tally returned no Stock Item masters for company {company_name!r}.")
    return sorted(names)


def _fingerprint(items: list[str]) -> str:
    exact_names = "\n".join(sorted(items))
    return hashlib.sha256(exact_names.encode("utf-8")).hexdigest()[:16]


def _company_key(company_name: str) -> str:
    return hashlib.sha256(company_name.encode("utf-8")).hexdigest()[:10]


def _cache_filename(cache_date: str, company_name: str, items: list[str]) -> str:
    return f"{cache_date}-{_company_key(company_name)}-{_fingerprint(items)}.json"


def _save_cache(items: list[str], company_name: str) -> dict[str, Any]:
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    now = datetime.now()
    payload = {
        "fetched_at": now.isoformat(timespec="seconds"),
        "date": now.date().isoformat(),
        "company_name": company_name,
        "company_fingerprint": _fingerprint(items),
        "matching": "strict_case_sensitive",
        "item_count": len(items),
        "items": items,
    }
    path = CACHE_DIR / _cache_filename(payload["date"], company_name, items)
    path.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    payload["source"] = "tally"
    payload["cache_path"] = str(path)
    return payload


def _latest_cache(date_prefix: str | None = None, company_name: str | None = None) -> Path | None:
    if not CACHE_DIR.exists():
        return None
    pattern = f"{date_prefix}-*.json" if date_prefix else "*.json"
    paths = list(CACHE_DIR.glob(pattern))
    if company_name:
        paths = [
            path
            for path in paths
            if json.loads(path.read_text(encoding="utf-8")).get("company_name") == company_name
        ]
    return max(paths, key=lambda path: path.stat().st_mtime) if paths else None


def _read_cache(path: Path, source: str) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    payload["source"] = source
    payload["cache_path"] = str(path)
    return payload

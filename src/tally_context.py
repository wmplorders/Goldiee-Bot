from __future__ import annotations

import html
import re
import urllib.request
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from typing import Any

from .settings import load_config


DEFAULT_TALLY_URL = "http://localhost:9000"


class TallyContextError(RuntimeError):
    pass


@dataclass(frozen=True)
class TallyConfig:
    company_name: str
    http_url: str = DEFAULT_TALLY_URL
    data_path: str = ""
    company_folder: str = ""
    require_exact_company: bool = True
    close_other_companies: bool = False


def get_tally_config(config: dict[str, Any] | None = None) -> TallyConfig:
    config = config or load_config()
    tally = config.get("tally", {})
    company_name = str(tally.get("company_name", "")).strip()
    if not company_name:
        raise TallyContextError("No Tally company_name is configured.")
    return TallyConfig(
        company_name=company_name,
        http_url=str(tally.get("http_url") or DEFAULT_TALLY_URL).strip(),
        data_path=str(tally.get("data_path") or "").strip(),
        company_folder=str(tally.get("company_folder") or "").strip(),
        require_exact_company=bool(tally.get("require_exact_company", True)),
        close_other_companies=bool(tally.get("close_other_companies", False)),
    )


def tally_request(body: str | bytes, tally_config: TallyConfig | None = None, timeout: int = 30) -> str:
    tally_config = tally_config or get_tally_config()
    data = body if isinstance(body, bytes) else body.encode("utf-8")
    request = urllib.request.Request(
        tally_config.http_url,
        data=data,
        headers={"Content-Type": "text/xml"},
    )
    return urllib.request.urlopen(request, timeout=timeout).read().decode("utf-8", errors="ignore")


def clean_tally_xml(response: str, include_udf_namespace: bool = False) -> str:
    response = re.sub(r"&#(?:[0-8]|1[0-9]|2[0-9]|3[0-1]);", "", response)
    if include_udf_namespace:
        response = response.replace("<ENVELOPE>", '<ENVELOPE xmlns:UDF="TallyUDF">', 1)
    return response


def list_loaded_companies(tally_config: TallyConfig | None = None) -> list[str]:
    tally_config = tally_config or get_tally_config()
    request_xml = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Goldiee Bot Loaded Companies</ID></HEADER><BODY><DESC>"
        "<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT></STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"Goldiee Bot Loaded Companies\">"
        "<TYPE>Company</TYPE><FETCH>Name</FETCH>"
        "</COLLECTION></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    root = ET.fromstring(clean_tally_xml(tally_request(request_xml, tally_config, timeout=20)))
    names = sorted(
        {
            node.attrib.get("NAME", "").strip()
            for node in root.findall(".//COMPANY")
            if node.attrib.get("NAME", "").strip()
        }
    )
    return names


def ensure_configured_company(tally_config: TallyConfig | None = None) -> dict[str, Any]:
    tally_config = tally_config or get_tally_config()
    loaded = list_loaded_companies(tally_config)
    if tally_config.require_exact_company and tally_config.company_name not in loaded:
        location = ""
        if tally_config.data_path or tally_config.company_folder:
            location = (
                f" Configured data path: {tally_config.data_path or '(blank)'};"
                f" company folder: {tally_config.company_folder or '(blank)'}."
            )
        raise TallyContextError(
            f"Configured Tally company is not loaded: {tally_config.company_name!r}. "
            f"Loaded companies: {loaded or 'none'}.{location} "
            "Open/select the configured company before running this command."
        )
    other_companies = [name for name in loaded if name != tally_config.company_name]
    if tally_config.close_other_companies and other_companies:
        raise TallyContextError(
            "Other Tally companies are open, and automatic closing is not available through "
            f"Tally HTTP. Close these companies first: {other_companies}"
        )
    return {
        "company_name": tally_config.company_name,
        "loaded_companies": loaded,
        "other_companies": other_companies,
        "data_path": tally_config.data_path,
        "company_folder": tally_config.company_folder,
    }


def current_company_static_variables(tally_config: TallyConfig) -> str:
    return f"<SVCurrentCompany>{html.escape(tally_config.company_name)}</SVCurrentCompany>"

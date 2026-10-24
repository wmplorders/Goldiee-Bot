from __future__ import annotations

import html
import xml.etree.ElementTree as ET
from decimal import Decimal
from typing import Any

from .tally_context import (
    clean_tally_xml,
    current_company_static_variables,
    ensure_configured_company,
    get_tally_config,
    tally_request,
)


def _text(node: ET.Element, tag: str) -> str:
    child = node.find(tag)
    return (child.text or "").strip() if child is not None else ""


def _udf_value(node: ET.Element, local_name: str) -> str:
    for child in node.iter():
        if child.tag.split("}")[-1] == local_name and not child.tag.endswith(".LIST"):
            value = (child.text or "").strip()
            if value:
                return value
    return ""


def _formula_for_names(names: list[str]) -> str:
    clauses = []
    for name in names:
        if '"' in name:
            raise ValueError(f'Scheme item name contains an unsupported quote: {name!r}')
        clauses.append(f'$Name = "{name}"')
    return " OR ".join(clauses)


def _inventory_formula_for_names(names: list[str]) -> str:
    clauses = []
    for name in names:
        if '"' in name:
            raise ValueError(f'Scheme item name contains an unsupported quote: {name!r}')
        clauses.append(f'$StockItemName = "{name}"')
    return " OR ".join(clauses)


def resolve_scheme_items(item_names: list[str]) -> dict[str, dict[str, Any]]:
    names = sorted({str(name).strip() for name in item_names if str(name).strip()})
    if not names:
        return {}

    tally_config = get_tally_config()
    ensure_configured_company(tally_config)
    static = current_company_static_variables(tally_config)
    item_formula = html.escape(_formula_for_names(names))
    master_request = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Goldiee Scheme Item Masters</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{static}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"Goldiee Scheme Item Masters\">"
        "<TYPE>Stock Item</TYPE><FILTER>GoldieeSchemeMaster</FILTER>"
        "<FETCH>Name,BaseUnits,AdditionalUnits,Denominator,Conversion</FETCH>"
        "</COLLECTION>"
        f"<SYSTEM TYPE=\"Formulae\" NAME=\"GoldieeSchemeMaster\">{item_formula}</SYSTEM>"
        "</TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    master_root = ET.fromstring(clean_tally_xml(tally_request(master_request, tally_config, timeout=30)))
    result: dict[str, dict[str, Any]] = {}
    for node in master_root.findall(".//STOCKITEM"):
        name = node.attrib.get("NAME", "").strip()
        if not name:
            continue
        denominator = Decimal(_text(node, "DENOMINATOR") or "0")
        conversion = Decimal(_text(node, "CONVERSION") or "0")
        if denominator <= 0 or conversion <= 0:
            raise ValueError(f"Scheme item {name!r} has no valid Case-to-PCS conversion in Tally.")
        if _text(node, "BASEUNITS").casefold() != "pcs":
            raise ValueError(f"Scheme item {name!r} does not use PCS as its Tally base unit.")
        if _text(node, "ADDITIONALUNITS").casefold() != "case":
            raise ValueError(f"Scheme item {name!r} does not use Case as its Tally alternate unit.")
        result[name] = {
            "pcs_per_case": denominator / conversion,
            "combinations": set(),
        }

    missing = [name for name in names if name not in result]
    if missing:
        raise ValueError("Scheme item missing in Tally: " + ", ".join(missing))

    inventory_formula = html.escape(_inventory_formula_for_names(names))
    history_request = (
        "<ENVELOPE><HEADER><VERSION>1</VERSION><TALLYREQUEST>Export</TALLYREQUEST>"
        "<TYPE>Collection</TYPE><ID>Goldiee Scheme Item History</ID></HEADER><BODY><DESC>"
        f"<STATICVARIABLES><SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>{static}</STATICVARIABLES>"
        "<TDL><TDLMESSAGE><COLLECTION NAME=\"Goldiee Scheme Item History\">"
        "<TYPE>Voucher</TYPE><FILTER>GoldieeSchemeHasTarget</FILTER>"
        "<FETCH>Date,IsCancelled,IsOptional,AllInventoryEntries.*,BatchAllocations.*</FETCH>"
        "</COLLECTION>"
        f"<SYSTEM TYPE=\"Formulae\" NAME=\"GoldieeSchemeTargetInventory\">{inventory_formula}</SYSTEM>"
        "<SYSTEM TYPE=\"Formulae\" NAME=\"GoldieeSchemeHasTarget\">"
        "$$FilterCount:AllInventoryEntries:GoldieeSchemeTargetInventory &gt; 0"
        "</SYSTEM></TDLMESSAGE></TDL></DESC></BODY></ENVELOPE>"
    )
    history_root = ET.fromstring(
        clean_tally_xml(tally_request(history_request, tally_config, timeout=120), include_udf_namespace=True)
    )
    for voucher in history_root.findall(".//VOUCHER"):
        if _text(voucher, "ISCANCELLED").casefold() == "yes":
            continue
        if _text(voucher, "ISOPTIONAL").casefold() == "yes":
            continue
        for entry in voucher.findall("ALLINVENTORYENTRIES.LIST"):
            name = _text(entry, "STOCKITEMNAME")
            if name not in result:
                continue
            for batch in entry.findall("BATCHALLOCATIONS.LIST"):
                ptr = _text(batch, "BATCHNAME")
                mrp = _udf_value(batch, "UDF_AA_BATCH_MRP_RATE")
                sell_rate = _udf_value(batch, "UDF_AA_BATCH_SELL_RATE")
                if not ptr or not mrp or not sell_rate:
                    continue
                result[name]["combinations"].add((ptr, Decimal(mrp), Decimal(sell_rate)))

    for details in result.values():
        details["combinations"] = sorted(
            details["combinations"],
            key=lambda combo: (Decimal(combo[0]), combo[0], combo[1], combo[2]),
        )
    return result


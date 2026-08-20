from __future__ import annotations

import asyncio
import re
from dataclasses import dataclass, field
from pathlib import Path

from .attachment_rules import select_invoice_excels
from .purchase_processor import build_summary, default_output_filename
from .state import append_log


DELIVERY_VOUCHER_COMMANDS = {
    "delivery voucher",
    "delivery vouchers",
    "generate voucher",
    "generate vouchers",
    "generate delivery vouchers",
    "generate challan",
    "generate challans",
}


@dataclass
class CommandResult:
    message: str
    attachments: list[Path] = field(default_factory=list)


def recognized_command(text: str) -> bool:
    normalized = " ".join(text.casefold().split())
    return (
        normalized == "purchase summary"
        or normalized in DELIVERY_VOUCHER_COMMANDS
        or normalized == "tally check"
        or normalized == "stock status"
        or normalized.startswith("stock goldiee mail ")
        or normalized == "dms stock"
        or normalized.startswith("unloading slip ")
    )


def uses_excel_attachment(text: str) -> bool:
    normalized = " ".join(text.casefold().split())
    return normalized == "purchase summary" or normalized in DELIVERY_VOUCHER_COMMANDS


async def run_command(
    text: str,
    attachments: list[Path],
    config: dict,
    source: str = "whatsapp",
) -> CommandResult:
    normalized = " ".join(text.casefold().split())
    if normalized == "tally check":
        return await asyncio.to_thread(_tally_check, config)
    if normalized == "purchase summary":
        return await asyncio.to_thread(_purchase_summary, attachments, config, source)
    if normalized in DELIVERY_VOUCHER_COMMANDS:
        return await asyncio.to_thread(_delivery_vouchers, attachments, config, source)
    if normalized == "stock status":
        return await asyncio.to_thread(_stock_status, config, source)
    if normalized.startswith("stock goldiee mail "):
        return await asyncio.to_thread(_stock_goldiee_mail, text, config, source)
    if normalized == "dms stock":
        return await asyncio.to_thread(_dms_stock, config, source)
    if normalized.startswith("unloading slip "):
        date_text = re.sub(r"^unloading\s+slip\s+", "", text.strip(), flags=re.IGNORECASE).strip()
        return await asyncio.to_thread(_unloading_slip, date_text, config, source)
    return CommandResult("Command not recognized.")


def _delivery_vouchers(attachments: list[Path], config: dict, source: str) -> CommandResult:
    from .delivery_vouchers import build_delivery_vouchers, default_output_filename

    excels = [path for path in attachments if path.suffix.casefold() in {".xlsx", ".xlsm"}]
    if not excels:
        uploads = Path(config["paths"]["uploads"])
        uploads.mkdir(parents=True, exist_ok=True)
        excels = sorted(
            (
                path
                for path in uploads.iterdir()
                if path.is_file() and path.suffix.casefold() in {".xlsx", ".xlsm"}
            ),
            key=lambda path: path.stat().st_mtime,
        )
    if not excels:
        return CommandResult(
            "I could not find the transport-details Excel. Send it on WhatsApp or place it in the "
            "bot's `uploads` folder, then send `delivery vouchers`."
        )

    source_path = excels[-1]
    output = Path(config["paths"]["processed"]) / default_output_filename(source_path)
    result = build_delivery_vouchers(source_path, output)
    errors = result["errors"]
    append_log(config["paths"]["processing_log"], {
        "status": "delivery_vouchers",
        "source": source,
        "input_rows": result["generated"] + len(errors),
        "output_rows": result["generated"],
        "output_file": result["output"],
        "note": f"Generated: {result['generated']}; Row errors: {len(errors)}",
    })
    message = f"Created {result['generated']} delivery voucher(s) in one PDF."
    if errors:
        shown = errors[:20]
        message += "\n\nRows not generated:\n- " + "\n- ".join(shown)
        if len(errors) > len(shown):
            message += f"\n- ...and {len(errors) - len(shown)} more row error(s)."
    return CommandResult(message, [Path(result["output"])])


def _tally_check(config: dict) -> CommandResult:
    from .tally_context import ensure_configured_company, get_tally_config

    result = ensure_configured_company(get_tally_config(config))
    loaded = ", ".join(result["loaded_companies"]) or "none"
    others = ", ".join(result["other_companies"]) or "none"
    return CommandResult(
        "Tally company lock OK."
        f"\nConfigured company: {result['company_name']}"
        f"\nLoaded companies: {loaded}"
        f"\nOther companies: {others}"
    )


def _purchase_summary(attachments: list[Path], config: dict, source: str) -> CommandResult:
    invoice_excels = select_invoice_excels(attachments)
    if not invoice_excels:
        return CommandResult(
            "I could not find a purchase invoice Excel. Send the Excel, then send `purchase summary`."
        )

    outputs: list[Path] = []
    invoice_numbers: list[str] = []
    new_items: set[str] = set()
    warnings: set[str] = set()
    for source_path in invoice_excels:
        output = Path(config["paths"]["processed"]) / default_output_filename(source_path)
        result = build_summary(source_path, output)
        outputs.append(Path(result["output"]))
        if result["invoice_no"]:
            invoice_numbers.append(result["invoice_no"])
        new_items.update(result["new_items"])
        if result["item_master_warning"]:
            warnings.add(result["item_master_warning"])
        append_log(config["paths"]["processing_log"], {
            "status": "processed",
            "source": source,
            **result,
            "output_file": result["output"],
        })
    invoices = ", ".join(invoice_numbers) or f"{len(outputs)} invoice file(s)"
    message = f"Purchase summary created for invoice {invoices}."
    if new_items:
        message += "\n\nNew items found:\n- " + "\n- ".join(sorted(new_items))
    if warnings:
        message += "\n\nWarning: " + " ".join(sorted(warnings))
    return CommandResult(message, outputs)


def _stock_status(config: dict, source: str) -> CommandResult:
    from .tally_batch_stock import build, default_output_path

    result = build(default_output_path())
    output = Path(result["output"])
    append_log(config["paths"]["processing_log"], {
        "status": "stock_status",
        "source": source,
        "output_file": str(output),
        "note": f"Rows: {result['rows']}; MRP rows: {result['mrp_rows']}; Sell rows: {result['sell_rows']}",
    })
    return CommandResult(f"Stock status created with {result['rows']} rows.", [output])


def _stock_goldiee_mail(text: str, config: dict, source: str) -> CommandResult:
    from .stock_aging_report import build_stock_aging_report, parse_report_date

    match = re.match(
        r"^\s*stock\s+goldiee\s+mail\s+(.+?)\s+to\s+(.+?)\s*$",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return CommandResult("Use: `stock goldiee mail <fromdate> to <todate>`")
    from_date = parse_report_date(match.group(1))
    to_date = parse_report_date(match.group(2))
    result = build_stock_aging_report(config, from_date, to_date)
    output = Path(result["output"])
    append_log(config["paths"]["processing_log"], {
        "status": "stock_goldiee_mail",
        "source": source,
        "output_file": str(output),
        "note": (
            f"Rows: {result['rows']}; Missing codes: {len(result['missing_codes'])}; "
            f"Missing weights: {len(result['missing_weights'])}; Stock source: {result['stock_source']}"
        ),
    })
    message = (
        f"Goldiee stock aging report created for {result['from_date']} to {result['to_date']} "
        f"with {result['rows']} rows."
    )
    if result["missing_codes"]:
        message += "\n\nItems without Tally item code:\n- " + "\n- ".join(result["missing_codes"][:20])
    if result["missing_weights"]:
        message += f"\n\nWeights could not be parsed for {len(result['missing_weights'])} item(s)."
    return CommandResult(message, [output])


def _dms_stock(config: dict, source: str) -> CommandResult:
    from .dms_client import DmsAuthenticationRequired, build_inventory_workbook

    try:
        result = build_inventory_workbook(config)
    except DmsAuthenticationRequired as exc:
        return CommandResult(str(exc))
    output = Path(result["output"])
    append_log(config["paths"]["processing_log"], {
        "status": "dms_stock",
        "source": source,
        "output_file": str(output),
        "note": f"Rows: {result['rows']}; Columns: {result['columns']}",
    })
    return CommandResult(f"DMS stock created with {result['rows']} rows.", [output])


def _unloading_slip(date_text: str, config: dict, source: str) -> CommandResult:
    from .unloading_slip import build_unloading_slip, parse_user_date

    result = build_unloading_slip(parse_user_date(date_text))
    attachments = [Path(result["xlsx"]), Path(result["pdf"])]
    append_log(config["paths"]["processing_log"], {
        "status": "unloading_slip",
        "source": source,
        "output_file": result["xlsx"],
        "note": f"PDF: {result['pdf']}; Rows: {result['rows']}",
    })
    invoices = ", ".join(result["invoice_numbers"]) or "not found"
    return CommandResult(
        f"Unloading slip created for {date_text}. Invoices combined: {invoices}.",
        attachments,
    )

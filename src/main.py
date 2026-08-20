from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .attachment_rules import select_invoice_excels
from .purchase_processor import build_summary, default_output_filename
from .settings import PROJECT_ROOT, load_config
from .state import append_log


def process_files(paths: list[Path], config: dict) -> list[Path]:
    outputs: list[Path] = []
    for source in select_invoice_excels(paths):
        output = Path(config["paths"]["processed"]) / default_output_filename(source)
        result = build_summary(source, output)
        outputs.append(Path(result["output"]))
        append_log(config["paths"]["processing_log"], {
            "status": "processed",
            "source": "manual",
            **result,
            "output_file": result["output"],
        })
        print(f"Created {result['output']}")
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser(description="Goldiee purchase invoice bot")
    sub = parser.add_subparsers(dest="command", required=True)

    process = sub.add_parser("process", help="Process one or more invoice Excel files")
    process.add_argument("files", nargs="+", type=Path)

    gmail_once = sub.add_parser("gmail-once", help="Check Gmail once for new invoices")
    monitor = sub.add_parser("monitor", help="Run WhatsApp command listener and Gmail monitor")

    whatsapp_login = sub.add_parser("whatsapp-login", help="Open Edge for WhatsApp QR/login")
    sub.add_parser("tally-check", help="Verify the configured Tally company is loaded")
    item_master = sub.add_parser("item-master", help="Warm the complete Tally Stock Item Master cache")
    stock_status = sub.add_parser("stock-status", help="Fetch Tally batch/PTR/MRP stock status")
    stock_goldiee_mail = sub.add_parser("stock-goldiee-mail", help="Create Goldiee stock ageing mail report")
    stock_goldiee_mail.add_argument("from_date", help="Date like 01-07-2026 or 01/07/2026")
    stock_goldiee_mail.add_argument("to_date", help="Date like 08-07-2026 or 08/07/2026")
    dms_authorize = sub.add_parser("dms-authorize", help="Authorize the browserless DMS client")
    dms_stock = sub.add_parser("dms-stock", help="Fetch current DMS stock without visible UI")
    unloading_slip = sub.add_parser("unloading-slip", help="Create unloading slip for a purchase date")
    unloading_slip.add_argument("date", help="Date like 30-05-2026 or 30/05/2026")
    unloading_slip.add_argument("invoices", nargs="*", help="Optional invoice numbers to include")
    delivery_vouchers = sub.add_parser(
        "delivery-vouchers",
        help="Create one combined delivery-voucher PDF from a transport-details Excel",
    )
    delivery_vouchers.add_argument("excel", type=Path)
    delivery_vouchers.add_argument("--output", type=Path)

    args = parser.parse_args()
    config = load_config()

    if args.command == "process":
        process_files(args.files, config)
        return 0

    if args.command in {"gmail-once", "monitor"}:
        if args.command == "monitor":
            import asyncio

            from .bot_runtime import run_bot

            asyncio.run(run_bot(config))
            return 0

        from .gmail_monitor import get_gmail_service, monitor_gmail, process_new_messages
        from .whatsapp_edge import send_to_configured_group

        def send_to_whatsapp(message, attachments=()):
            send_to_configured_group(config, message, attachments)

        service = get_gmail_service(PROJECT_ROOT)
        # Label creation is inside monitor_gmail; reuse it by importing the private helper here.
        from .gmail_monitor import _ensure_label

        label_id = _ensure_label(service, config["gmail"]["processed_label"])
        process_new_messages(service, label_id, config, send_to_whatsapp)
        return 0

    if args.command == "whatsapp-login":
        from .whatsapp_edge import send_to_configured_group

        send_to_configured_group(config, "Goldiee Purchase Bot is connected.")
        return 0

    if args.command == "tally-check":
        from .tally_context import ensure_configured_company, get_tally_config

        tally_config = get_tally_config(config)
        result = ensure_configured_company(tally_config)
        print(f"Configured company: {result['company_name']}")
        print(f"Loaded companies: {', '.join(result['loaded_companies']) or 'none'}")
        print(f"Other companies: {', '.join(result['other_companies']) or 'none'}")
        if result["data_path"]:
            print(f"Configured data path: {result['data_path']}")
        if result["company_folder"]:
            print(f"Configured company folder: {result['company_folder']}")
        print("Tally company lock OK.")
        return 0

    if args.command == "item-master":
        from .tally_item_master import get_item_master

        result = get_item_master(force_refresh=True)
        print(f"Cached {result['item_count']} Stock Item masters.")
        print(f"Company fingerprint: {result['company_fingerprint']}")
        print(f"Cache: {result['cache_path']}")
        return 0

    if args.command == "stock-status":
        from .tally_batch_stock import build, default_output_path

        result = build(default_output_path())
        append_log(config["paths"]["processing_log"], {
            "status": "stock_status",
            "source": "manual",
            "output_file": result["output"],
            "note": f"Rows: {result['rows']}; MRP rows: {result['mrp_rows']}; Sell rows: {result['sell_rows']}",
        })
        print(f"Created {result['output']}")
        return 0

    if args.command == "stock-goldiee-mail":
        from .stock_aging_report import build_stock_aging_report, parse_report_date

        result = build_stock_aging_report(
            config,
            parse_report_date(args.from_date),
            parse_report_date(args.to_date),
        )
        append_log(config["paths"]["processing_log"], {
            "status": "stock_goldiee_mail",
            "source": "manual",
            "output_file": result["output"],
            "note": (
                f"Rows: {result['rows']}; Missing codes: {len(result['missing_codes'])}; "
                f"Missing weights: {len(result['missing_weights'])}; Stock source: {result['stock_source']}"
            ),
        })
        print(f"Created {result['output']}")
        print(f"Rows: {result['rows']}")
        print(f"Missing item codes: {len(result['missing_codes'])}")
        print(f"Missing weights: {len(result['missing_weights'])}")
        return 0

    if args.command == "dms-authorize":
        from .dms_client import authorize

        authorize(config)
        print("DMS authorization saved.")
        return 0

    if args.command == "dms-stock":
        from .dms_client import DmsAuthenticationRequired, build_inventory_workbook

        try:
            result = build_inventory_workbook(config)
        except DmsAuthenticationRequired as exc:
            print(exc)
            return 2
        print(f"Created {result['output']} with {result['rows']} rows.")
        return 0

    if args.command == "unloading-slip":
        from .unloading_slip import build_unloading_slip, parse_user_date

        result = build_unloading_slip(parse_user_date(args.date), args.invoices)
        append_log(config["paths"]["processing_log"], {
            "status": "unloading_slip",
            "source": "manual",
            "output_file": result["xlsx"],
            "note": f"PDF: {result['pdf']}; Rows: {result['rows']}",
        })
        print(f"Created {result['xlsx']}")
        print(f"Created {result['pdf']}")
        return 0

    if args.command == "delivery-vouchers":
        from .delivery_vouchers import build_delivery_vouchers, default_output_filename

        output = args.output or Path(config["paths"]["processed"]) / default_output_filename(args.excel)
        result = build_delivery_vouchers(args.excel, output)
        append_log(config["paths"]["processing_log"], {
            "status": "delivery_vouchers",
            "source": "manual",
            "input_rows": result["generated"] + len(result["errors"]),
            "output_rows": result["generated"],
            "output_file": result["output"],
            "note": f"Generated: {result['generated']}; Row errors: {len(result['errors'])}",
        })
        print(f"Created {result['output']}")
        print(f"Generated vouchers: {result['generated']}")
        for error in result["errors"]:
            print(error)
        return 0

    return 1


if __name__ == "__main__":
    sys.exit(main())

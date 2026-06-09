from __future__ import annotations

import argparse
import re
from collections import OrderedDict
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo


REQUIRED_HEADERS = {
    "item": "ITEM",
    "ptr": "PTR",
    "mrp": "MRP",
    "case": "Case",
    "qty": "QTY.",
    "billing_rate": "Billing Rate",
    "amount": "AMOUNT",
    "discount": "DISCOUNT",
    "invoice_no": "Invoice No.",
    "date": "Date",
}


def _clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def _clean_filename_part(value: Any) -> str:
    text = str(value or "").strip()
    text = text.replace("/", "-").replace("\\", "-")
    text = re.sub(r'[<>:"|?*]+', "-", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" .-")


def _as_float(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    return float(value)


def _find_header_row_and_columns(ws) -> tuple[int, dict[str, int]]:
    wanted = {_clean_header(label): key for key, label in REQUIRED_HEADERS.items()}
    for row in range(1, min(ws.max_row or 1, 20) + 1):
        found: dict[str, int] = {}
        for col in range(1, (ws.max_column or 1) + 1):
            header = _clean_header(ws.cell(row, col).value)
            if header in wanted:
                found[wanted[header]] = col
        if set(REQUIRED_HEADERS) <= set(found):
            return row, found
    missing = ", ".join(REQUIRED_HEADERS.values())
    raise ValueError(f"Could not find required invoice headers: {missing}")


def _is_item_row(item: str) -> bool:
    normalized = item.strip().casefold()
    return bool(normalized) and normalized != "total"


def build_summary(source_path: Path, output_path: Path | None = None) -> dict[str, Any]:
    source_path = Path(source_path)
    if not source_path.exists():
        raise FileNotFoundError(source_path)

    raw_values_wb = load_workbook(source_path, data_only=True)
    raw_values_ws = raw_values_wb[raw_values_wb.sheetnames[0]]
    header_row, columns = _find_header_row_and_columns(raw_values_ws)

    groups: OrderedDict[tuple[str, float], dict[str, Any]] = OrderedDict()
    input_rows = 0
    first_invoice_no = ""
    first_invoice_date = ""

    for row in range(header_row + 1, (raw_values_ws.max_row or header_row) + 1):
        item = str(raw_values_ws.cell(row, columns["item"]).value or "").strip()
        if not _is_item_row(item):
            continue

        input_rows += 1
        billing_rate = _as_float(raw_values_ws.cell(row, columns["billing_rate"]).value)
        key = (item, billing_rate)

        invoice_no = str(raw_values_ws.cell(row, columns["invoice_no"]).value or "").strip()
        invoice_date = str(raw_values_ws.cell(row, columns["date"]).value or "").strip()
        first_invoice_no = first_invoice_no or invoice_no
        first_invoice_date = first_invoice_date or invoice_date

        if key not in groups:
            groups[key] = {
                "ptr": _as_float(raw_values_ws.cell(row, columns["ptr"]).value),
                "mrp": _as_float(raw_values_ws.cell(row, columns["mrp"]).value),
                "item": item,
                "case": 0.0,
                "qty": 0.0,
                "billing_rate": billing_rate,
                "discount": 0.0,
                "amount": 0.0,
                "invoice_no": invoice_no,
                "date": invoice_date,
            }

        group = groups[key]
        group["case"] += _as_float(raw_values_ws.cell(row, columns["case"]).value)
        group["qty"] += _as_float(raw_values_ws.cell(row, columns["qty"]).value)
        group["discount"] += _as_float(raw_values_ws.cell(row, columns["discount"]).value)
        group["amount"] += _as_float(raw_values_ws.cell(row, columns["amount"]).value)

    records = sorted(groups.values(), key=lambda record: record["item"].casefold())

    if output_path is None:
        output_path = source_path.with_name(_filename_from_invoice(source_path, first_invoice_no, first_invoice_date))
    output_path = Path(output_path)

    from .tally_item_master import get_item_master

    item_master = get_item_master()
    tally_items = set(item_master["items"])
    new_items = sorted({record["item"] for record in records if record["item"] not in tally_items})

    output_wb = load_workbook(source_path)
    raw_ws = output_wb[output_wb.sheetnames[0]]
    raw_ws.title = "Raw Data"
    if "Purchase Summary" in output_wb.sheetnames:
        del output_wb["Purchase Summary"]
    output_ws = output_wb.create_sheet("Purchase Summary")

    headers = [
        "No.",
        "PTR",
        "MRP",
        "Sell Rate",
        "ITEM",
        "Case",
        "QTY",
        "Billing Rate",
        "Discount %",
        "Invoice No.",
        "Date",
        "Status",
    ]
    output_ws.append(headers)

    for serial, record in enumerate(records, start=1):
        multiplier = 1.025 if record["item"].casefold().startswith("tea") else 1.03
        sell_rate = round(record["billing_rate"] * multiplier, 2)
        discount_percent = (
            round((record["discount"] / record["amount"]) * 100, 2) / 100
            if record["amount"]
            else 0
        )
        output_ws.append(
            [
                serial,
                record["ptr"],
                record["mrp"],
                sell_rate,
                record["item"],
                record["case"],
                record["qty"],
                record["billing_rate"],
                discount_percent,
                record["invoice_no"],
                record["date"],
                "NEW ITEM" if record["item"] in new_items else "",
            ]
        )

    _format_summary_sheet(output_ws, new_items)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_wb.save(output_path)

    return {
        "source": str(source_path),
        "output": str(output_path),
        "invoice_no": first_invoice_no,
        "invoice_date": first_invoice_date,
        "input_rows": input_rows,
        "output_rows": len(records),
        "consolidated_rows": input_rows - len(records),
        "total_case": sum(record["case"] for record in records),
        "total_qty": sum(record["qty"] for record in records),
        "created_at": datetime.now().isoformat(timespec="seconds"),
        "new_items": new_items,
        "item_master_source": item_master["source"],
        "item_master_fetched_at": item_master["fetched_at"],
        "item_master_company_fingerprint": item_master["company_fingerprint"],
        "item_master_warning": item_master.get("warning", ""),
    }


def default_output_filename(source_path: Path) -> str:
    source_path = Path(source_path)
    values_wb = load_workbook(source_path, data_only=True, read_only=False)
    values_ws = values_wb[values_wb.sheetnames[0]]
    header_row, columns = _find_header_row_and_columns(values_ws)
    for row in range(header_row + 1, (values_ws.max_row or header_row) + 1):
        item = str(values_ws.cell(row, columns["item"]).value or "").strip()
        if _is_item_row(item):
            invoice_no = str(values_ws.cell(row, columns["invoice_no"]).value or "").strip()
            invoice_date = str(values_ws.cell(row, columns["date"]).value or "").strip()
            return _filename_from_invoice(source_path, invoice_no, invoice_date)
    return f"{_clean_filename_part(source_path.stem)} Summary.xlsx"


def _filename_from_invoice(source_path: Path, invoice_no: str, invoice_date: str) -> str:
    safe_invoice = _clean_filename_part(invoice_no) or source_path.stem
    safe_date = _clean_filename_part(invoice_date)
    return f"{safe_invoice} {safe_date} Summary.xlsx".strip()


def _format_summary_sheet(ws, new_items: list[str]) -> None:
    max_row = ws.max_row
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row in range(2, max_row + 1):
        ws.cell(row, 1).alignment = Alignment(horizontal="center")
        ws.cell(row, 10).alignment = Alignment(horizontal="center")
        ws.cell(row, 11).alignment = Alignment(horizontal="center")
        for column in (2, 3, 4, 8):
            ws.cell(row, column).number_format = "0.00"
        ws.cell(row, 6).number_format = "General"
        ws.cell(row, 7).number_format = "0"
        ws.cell(row, 9).number_format = "0.00%"
        if ws.cell(row, 5).value in new_items:
            yellow = PatternFill("solid", fgColor="FFF2CC")
            for column in range(1, 13):
                ws.cell(row, column).fill = yellow
            ws.cell(row, 12).font = Font(bold=True, color="9C6500")

    widths = {
        "A": 7,
        "B": 10,
        "C": 10,
        "D": 12,
        "E": 48,
        "F": 12,
        "G": 12,
        "H": 14,
        "I": 14,
        "J": 20,
        "K": 14,
        "L": 14,
    }
    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 30

    if max_row > 1:
        table = Table(displayName="PurchaseSummaryTable", ref=f"A1:L{max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws.add_table(table)


def main() -> None:
    parser = argparse.ArgumentParser(description="Build Goldiee purchase summary workbook.")
    parser.add_argument("source", type=Path)
    parser.add_argument("output", type=Path, nargs="?")
    args = parser.parse_args()
    result = build_summary(args.source, args.output)
    for key, value in result.items():
        print(f"{key}: {value}")


if __name__ == "__main__":
    main()

"""Final compact Goenka Tradeventures delivery-voucher PDF template."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal
from typing import Any

from PIL import Image
from reportlab.lib.colors import black
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


@dataclass(frozen=True)
class VoucherData:
    voucher_no: str
    voucher_date: date | datetime
    invoice_no: str
    invoice_party: str
    station: str
    vehicle_no: str
    eway_no: str
    total_cases: Decimal
    per_case_freight: Decimal
    total_freight: Decimal
    invoice_value: Decimal
    invoice_value_words: str
    remarks: str = ""


def register_fonts() -> None:
    if "Arial" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("Arial", r"C:\Windows\Fonts\arial.ttf"))
    if "Arial-Bold" not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont("Arial-Bold", r"C:\Windows\Fonts\arialbd.ttf"))


def prepare_stamp(path: str, opacity: float = 0.78) -> Image.Image:
    image = Image.open(path).convert("RGBA")
    image.putalpha(image.getchannel("A").point(lambda value: int(value * opacity)))
    bbox = image.getbbox()
    return image.crop(bbox) if bbox else image


def _text(c, x, y, text, size=8, bold=False, align="left"):
    font = "Arial-Bold" if bold else "Arial"
    c.setFont(font, size)
    if align == "right":
        c.drawRightString(x, y, str(text))
    elif align == "center":
        c.drawCentredString(x, y, str(text))
    else:
        c.drawString(x, y, str(text))


def _fit_text(c, x, y, text, max_width, size, minimum, bold=False, align="left"):
    font = "Arial-Bold" if bold else "Arial"
    fitted = size
    while fitted > minimum and pdfmetrics.stringWidth(str(text), font, fitted) > max_width:
        fitted -= 0.25
    _text(c, x, y, text, fitted, bold, align)


def _wrap_text(text: str, font: str, size: float, max_width: float, max_lines: int = 2):
    words = str(text).split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        if not current or pdfmetrics.stringWidth(candidate, font, size) <= max_width:
            current = candidate
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        last = lines[-1]
        while last and pdfmetrics.stringWidth(last + "...", font, size) > max_width:
            last = last[:-1]
        lines[-1] = last.rstrip() + "..."
    return lines


def _rule(c, x1, x2, y, width=0.55):
    c.setLineWidth(width)
    c.line(x1, y, x2, y)


def _vertical_rule(c, x, y1, y2, width=0.45):
    c.setLineWidth(width)
    c.line(x, y1, x, y2)


def _label(c, x, y, text):
    _text(c, x, y, str(text).upper(), 6.5, True)


def _date_text(value: date | datetime) -> str:
    return value.strftime("%d/%m/%Y")


def _indian_number(value: Decimal, decimals=True) -> str:
    value = Decimal(str(value))
    sign = "-" if value < 0 else ""
    value = abs(value)
    whole = int(value)
    fraction = int((value - whole).quantize(Decimal("0.01")) * 100)
    digits = str(whole)
    if len(digits) > 3:
        grouped = digits[-3:]
        rest = digits[:-3]
        while rest:
            grouped = rest[-2:] + "," + grouped
            rest = rest[:-2]
    else:
        grouped = digits
    return f"{sign}{grouped}.{fraction:02d}" if decimals and fraction else f"{sign}{grouped}"


def draw_voucher(c, data: VoucherData, stamp: Image.Image) -> None:
    register_fonts()
    page_w, page_h = A4
    c.setFillColor(black)
    c.setStrokeColor(black)
    left = 13 * mm
    right = page_w - 13 * mm
    top = page_h - 12 * mm

    _text(c, left, top, "GOENKA TRADEVENTURES", 15, True)
    _text(c, left, top - 7 * mm, "PO Barckachha, Near Rice Mill, Muhkochwa,", 7.8)
    _text(c, left, top - 11.5 * mm, "Mirzapur - 231001 (U.P.)", 7.8)
    _text(c, left, top - 16 * mm, "8887931633 / 9919102011", 7.8)
    _text(c, left, top - 20.5 * mm, "GSTIN 09AKEPG9753A1Z8", 7.8)

    _text(c, right, top + 0.5 * mm, "DELIVERY VOUCHER", 13.5, True, "right")
    box_w, box_h = 55 * mm, 25 * mm
    box_x, box_y = right - box_w, top - 30.5 * mm
    c.setLineWidth(0.65)
    c.rect(box_x, box_y, box_w, box_h)
    _text(c, box_x + box_w / 2, box_y + 19.3 * mm, "VOUCHER NO.", 6.8, align="center")
    _fit_text(c, box_x + box_w / 2, box_y + 10.6 * mm, data.voucher_no, box_w - 14 * mm, 17, 12, True, "center")
    _rule(c, box_x + 7 * mm, box_x + box_w - 7 * mm, box_y + 7.5 * mm, 0.45)
    _text(c, box_x + box_w / 2, box_y + 3 * mm, f"VOUCHER DATE  {_date_text(data.voucher_date)}", 6.9, align="center")

    section_y = top - 33 * mm
    _label(c, left, section_y, "Consignee")
    _rule(c, left, left + 8 * mm, section_y - 2.8 * mm, 0.8)
    party_lines = _wrap_text(data.invoice_party, "Arial-Bold", 8.5, 91 * mm, 2)
    for idx, line in enumerate(party_lines):
        _text(c, left, section_y - (10.5 + idx * 4.2) * mm, line, 8.5, True)
    _text(c, left, section_y - 22.5 * mm, f"Consignee Address   {data.station}", 7.8)

    ship_x = 114 * mm
    _label(c, ship_x, section_y, "Shipment")
    _rule(c, ship_x, ship_x + 8 * mm, section_y - 2.8 * mm, 0.8)
    shipment = [
        ("DESPATCH FROM", "MIRZAPUR"),
        ("DESPATCH TO", data.station),
        ("DESTINATION / STATION", data.station),
    ]
    for idx, (key, val) in enumerate(shipment):
        row_y = section_y - (10 + idx * 7) * mm
        _text(c, ship_x, row_y, key, 7.1)
        _fit_text(c, right - 36 * mm, row_y, val, 36 * mm, 7.8, 6.5)

    meta_top = top - 61 * mm
    meta_bottom = meta_top - 16 * mm
    _rule(c, left, right, meta_top, 0.65)
    _rule(c, left, right, meta_bottom, 0.65)
    widths = [43, 47, 42, 50]
    metadata = [
        ("INVOICE NO.", data.invoice_no),
        ("INVOICE DATE", _date_text(data.voucher_date)),
        ("VEHICLE NO.", data.vehicle_no),
        ("E-WAY BILL NO.", data.eway_no),
    ]
    x = left
    for idx, ((key, val), width_mm) in enumerate(zip(metadata, widths)):
        width = width_mm * mm
        _text(c, x + width / 2, meta_top - 5.8 * mm, key, 7, True, "center")
        _fit_text(c, x + width / 2, meta_top - 12.2 * mm, val, width - 5 * mm, 9.2, 6.5, True, "center")
        x += width
        if idx < len(metadata) - 1:
            _vertical_rule(c, x, meta_bottom + 3 * mm, meta_top - 3 * mm)

    stat_top = meta_bottom - 3 * mm
    stat_bottom = stat_top - 17 * mm
    stat_width = 29 * mm
    statistics = [
        ("TOTAL CASES", _indian_number(data.total_cases, False)),
        ("FREIGHT PER CASE", f"₹{_indian_number(data.per_case_freight)}"),
        ("TOTAL FREIGHT", f"₹{_indian_number(data.total_freight)}"),
    ]
    x = left
    for key, val in statistics:
        _text(c, x + stat_width / 2, stat_top - 3.5 * mm, key, 5.5, align="center")
        _rule(c, x + stat_width / 2 - 3.5 * mm, x + stat_width / 2 + 3.5 * mm, stat_top - 6.4 * mm, 0.5)
        _fit_text(c, x + stat_width / 2, stat_top - 13.5 * mm, val, stat_width - 4 * mm, 9, 7, align="center")
        x += stat_width
        _vertical_rule(c, x, stat_bottom + 2 * mm, stat_top - 1 * mm)

    invoice_x = x + 7 * mm
    _label(c, invoice_x, stat_top - 0.5 * mm, "Total Invoice Value")
    _rule(c, invoice_x, invoice_x + 8 * mm, stat_top - 3.7 * mm, 0.8)
    invoice_amount = f"₹{_indian_number(data.invoice_value)}"
    _fit_text(c, invoice_x, stat_top - 10.5 * mm, invoice_amount, right - invoice_x - 3 * mm, 12, 9, True)
    _rule(c, invoice_x, right - 3 * mm, stat_top - 13.8 * mm, 0.45)
    _fit_text(c, invoice_x, stat_top - 17.2 * mm, data.invoice_value_words, right - invoice_x - 3 * mm, 6.5, 5.2)

    lower_y = top - 103 * mm
    _label(c, left, lower_y, "Remarks")
    _rule(c, left, left + 8 * mm, lower_y - 2.8 * mm, 0.8)
    if data.remarks:
        _fit_text(c, left, lower_y - 7.5 * mm, data.remarks, 115 * mm, 7.2, 5.5)

    sign_center_x = right - 23 * mm
    _text(c, sign_center_x, lower_y, "For GOENKA TRADEVENTURES", 6.5, True, "center")
    sign_line_y = lower_y - 9 * mm
    _rule(c, sign_center_x - 23 * mm, sign_center_x + 23 * mm, sign_line_y, 0.55)
    _text(c, sign_center_x, sign_line_y - 5.2 * mm, "Authorized Signatory", 6.8, align="center")
    stamp_size = 18 * mm
    c.drawImage(
        ImageReader(stamp),
        sign_center_x - stamp_size / 2,
        sign_line_y - stamp_size / 2 + 2 * mm,
        stamp_size,
        stamp_size,
        preserveAspectRatio=True,
        mask="auto",
    )

    _text(c, page_w / 2, top - 122.5 * mm, "This is a computer generated document.", 6.5, align="center")

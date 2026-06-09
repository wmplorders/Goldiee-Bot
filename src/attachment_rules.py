from __future__ import annotations

from pathlib import Path


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}


def is_pending_order(path: str | Path) -> bool:
    return "pending" in Path(path).name.casefold()


def is_excel(path: str | Path) -> bool:
    return Path(path).suffix.casefold() in EXCEL_EXTENSIONS


def is_pdf(path: str | Path) -> bool:
    return Path(path).suffix.casefold() == ".pdf"


def select_invoice_excels(paths: list[str | Path]) -> list[Path]:
    excels = [Path(path) for path in paths if is_excel(path) and not is_pending_order(path)]
    invoice_named = [path for path in excels if "invoice" in path.name.casefold()]
    return invoice_named or excels


def select_pdfs(paths: list[str | Path]) -> list[Path]:
    return [Path(path) for path in paths if is_pdf(path)]

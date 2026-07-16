from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

from .models import SourceLine, SourceReport
from .spreadsheet import SheetData, read_workbook
from .text_utils import (
    adjacent_value,
    collapse_spaces,
    excel_serial_to_date,
    find_column,
    normalize_delivery_way,
    normalize_document_no,
    normalize_part,
    safe_text,
    to_number,
)


PART_ALIASES = ("QAD PN", "Part No.", "Part No", "PartNo", "料号")
QUANTITY_ALIASES = ("Quantity", "Qty", "Qty.", "数量")


def _cell(row: list[Any], column: int | None) -> Any:
    if column is None or column < 0 or column >= len(row):
        return ""
    return row[column]


def _is_business_label(value: Any) -> bool:
    text = safe_text(value).lower()
    labels = (
        "bill to",
        "ship to",
        "serial",
        "page",
        "our ref",
        "customer",
        "delivery",
        "term",
        "telephone",
        "fax",
        "post code",
        "address",
        "payment",
    )
    return not text or any(label in text for label in labels)


def _extract_party_after_label(
    rows: list[list[Any]],
    label_pattern: str,
    max_rows: int = 35,
) -> str:
    for row_index, row in enumerate(rows[:max_rows]):
        for column_index, value in enumerate(row):
            if not re.search(label_pattern, safe_text(value), re.I):
                continue
            same_row = adjacent_value(row, column_index)
            if same_row and not _is_business_label(same_row):
                return collapse_spaces(same_row)
            for next_row in rows[row_index + 1 : min(len(rows), row_index + 6)]:
                for candidate_index in (column_index, column_index + 1):
                    candidate = safe_text(_cell(next_row, candidate_index))
                    if candidate and not _is_business_label(candidate):
                        return collapse_spaces(candidate)
    return ""


def _find_detail_sheet(
    sheets: list[SheetData],
) -> tuple[SheetData, int]:
    candidates: list[tuple[int, SheetData, int]] = []
    for sheet in sheets:
        for row_index, row in enumerate(sheet.rows):
            part_column = find_column(row, PART_ALIASES)
            quantity_column = find_column(row, QUANTITY_ALIASES)
            if part_column is None or quantity_column is None:
                continue
            score = 100
            if sheet.state == "visible":
                score += 10
            if find_column(row, ("Amount", "金额")) is not None:
                score += 5
            if find_column(row, ("Currency", "币种")) is not None:
                score += 3
            candidates.append((score, sheet, row_index))
    if not candidates:
        raise ValueError("源报表中没有找到同时包含 QAD PN/Part No. 和 Quantity/Qty 的明细表")
    candidates.sort(key=lambda item: item[0], reverse=True)
    _, sheet, header_index = candidates[0]
    return sheet, header_index


def _parse_metadata(
    source_path: Path,
    rows: list[list[Any]],
) -> tuple[str, date | None, str, str, str]:
    serial_no = ""
    pickup_date = None
    delivery_way = ""
    currency = ""

    for row in rows[:35]:
        for column_index, value in enumerate(row):
            text = safe_text(value)
            if re.search(r"serial\s*no", text, re.I):
                serial_no = (
                    normalize_document_no(adjacent_value(row, column_index))
                    or serial_no
                )
            if not serial_no:
                serial_no = normalize_document_no(text) or serial_no
            if re.search(r"pick\s*up\s*date|pickup\s*date", text, re.I):
                pickup_date = (
                    excel_serial_to_date(adjacent_value(row, column_index))
                    or pickup_date
                )
            if re.search(
                r"delivery\s*(way|mode|method)|transport",
                text,
                re.I,
            ):
                delivery_way = (
                    normalize_delivery_way(adjacent_value(row, column_index))
                    or delivery_way
                )
            if text.upper() in ("USD", "EUR", "CNY", "GBP", "JPY"):
                currency = text.upper()

    serial_no = (
        serial_no
        or normalize_document_no(source_path.stem)
        or source_path.stem
    )
    pickup_date = pickup_date or excel_serial_to_date(source_path.name)
    delivery_way = (
        delivery_way
        or normalize_delivery_way(source_path.name)
    )
    ship_to = _extract_party_after_label(rows, r"ship\s*to")
    return serial_no, pickup_date, ship_to, delivery_way, currency


def parse_source_report(path: str | Path) -> SourceReport:
    source_path = Path(path)
    sheets = read_workbook(source_path)
    detail_sheet, header_index = _find_detail_sheet(sheets)
    rows = detail_sheet.rows
    headers = rows[header_index]

    columns = {
        "item_no": find_column(headers, ("Item", "Item No.", "项号")),
        "qad_pn": find_column(headers, PART_ALIASES),
        "description_en": find_column(
            headers,
            ("Description", "English Description", "英文描述"),
        ),
        "quantity": find_column(headers, QUANTITY_ALIASES),
        "unit_price": find_column(
            headers,
            ("UP", "Unit Price", "单价"),
        ),
        "po_no": find_column(
            headers,
            ("PO_No", "PO No.", "PO No", "PO"),
        ),
        "amount": find_column(headers, ("Amount", "金额")),
        "currency": find_column(headers, ("Currency", "币种")),
    }

    serial_no, pickup_date, ship_to, delivery_way, report_currency = (
        _parse_metadata(source_path, rows)
    )

    lines: list[SourceLine] = []
    for row_index, row in enumerate(rows[header_index + 1 :], header_index + 2):
        qad_pn = normalize_part(_cell(row, columns["qad_pn"]))
        if not qad_pn or not re.search(r"\d", qad_pn):
            continue
        quantity = to_number(_cell(row, columns["quantity"]))
        unit_price = to_number(_cell(row, columns["unit_price"]))
        amount = to_number(_cell(row, columns["amount"]))
        if amount == 0 and quantity and unit_price:
            amount = round(quantity * unit_price, 6)
        if quantity <= 0 and amount <= 0:
            continue
        line_currency = (
            safe_text(_cell(row, columns["currency"])).upper()
            or report_currency
        )
        if line_currency:
            report_currency = line_currency
        lines.append(
            SourceLine(
                item_no=safe_text(_cell(row, columns["item_no"])),
                qad_pn=qad_pn,
                description_en=safe_text(
                    _cell(row, columns["description_en"])
                ),
                quantity=quantity,
                unit_price=unit_price,
                amount=round(amount, 6),
                currency=line_currency,
                po_no=safe_text(_cell(row, columns["po_no"])),
                source_file=source_path.name,
                source_sheet=detail_sheet.name,
                source_row=row_index,
            )
        )

    if not lines:
        raise ValueError(f"{source_path.name} 中没有数量或金额大于 0 的有效明细")

    return SourceReport(
        source_path=source_path,
        serial_no=serial_no,
        pickup_date=pickup_date,
        ship_to=ship_to,
        delivery_way=delivery_way,
        currency=report_currency,
        lines=tuple(lines),
    )

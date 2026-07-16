from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class SourceLine:
    item_no: str
    qad_pn: str
    description_en: str
    quantity: float
    unit_price: float
    amount: float
    currency: str
    po_no: str
    source_file: str
    source_sheet: str
    source_row: int


@dataclass(frozen=True)
class SourceReport:
    source_path: Path
    serial_no: str
    pickup_date: date | None
    ship_to: str
    delivery_way: str
    currency: str
    lines: tuple[SourceLine, ...]


@dataclass(frozen=True)
class RuleRecord:
    qad_pn: str
    goods_name: str
    hs_code: str
    brand: str = ""
    note: str = ""
    source_sheet: str = ""
    source_row: int = 0


@dataclass(frozen=True)
class MergeWarning:
    code: str
    message: str
    source_file: str = ""
    qad_pn: str = ""
    output_row: int | None = None


@dataclass
class MergePreview:
    template_sheet: str
    rule_sheet: str
    rows: list[dict[str, Any]]
    warnings: list[MergeWarning] = field(default_factory=list)

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def total_quantity(self) -> float:
        return sum(float(row.get("quantity") or 0) for row in self.rows)

    @property
    def total_amount(self) -> float:
        return round(sum(float(row.get("amount") or 0) for row in self.rows), 2)

    def summary(self) -> dict[str, Any]:
        by_serial: dict[str, dict[str, float | int]] = {}
        for row in self.rows:
            serial_no = str(row.get("serial_no") or "")
            current = by_serial.setdefault(
                serial_no,
                {"rows": 0, "quantity": 0.0, "amount": 0.0},
            )
            current["rows"] = int(current["rows"]) + 1
            current["quantity"] = float(current["quantity"]) + float(row.get("quantity") or 0)
            current["amount"] = float(current["amount"]) + float(row.get("amount") or 0)
        for current in by_serial.values():
            current["quantity"] = round(float(current["quantity"]), 6)
            current["amount"] = round(float(current["amount"]), 2)
        return {
            "rowCount": self.row_count,
            "totalQuantity": round(self.total_quantity, 6),
            "totalAmount": self.total_amount,
            "templateSheet": self.template_sheet,
            "ruleSheet": self.rule_sheet,
            "warnings": [
                {
                    "code": warning.code,
                    "message": warning.message,
                    "sourceFile": warning.source_file,
                    "qadPn": warning.qad_pn,
                    "outputRow": warning.output_row,
                }
                for warning in self.warnings
            ],
            "bySerial": by_serial,
        }


@dataclass
class MergeResult(MergePreview):
    output_path: Path = Path()

    def summary(self) -> dict[str, Any]:
        result = super().summary()
        result["outputPath"] = str(self.output_path)
        return result

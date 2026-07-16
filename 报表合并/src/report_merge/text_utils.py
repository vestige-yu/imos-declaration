from __future__ import annotations

import datetime as dt
import re
import unicodedata
from dataclasses import dataclass
from typing import Any, Iterable


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).replace("\u00a0", " ").strip()


def collapse_spaces(value: Any) -> str:
    return re.sub(r"\s+", " ", safe_text(value)).strip()


def normalize_header(value: Any) -> str:
    return re.sub(r"[\s_.:/\\-]+", "", safe_text(value).lower())


def normalize_part(value: Any) -> str:
    text = unicodedata.normalize("NFKC", safe_text(value)).upper()
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return text


@dataclass(frozen=True)
class PartNumberNormalization:
    literal: str
    canonical: str
    transformations: tuple[str, ...]


PART_RULE_LABELS = {
    "separator": "忽略分隔符",
    "outer_zero": "忽略首尾零",
    "version_suffix": "忽略版本后缀差异",
    "fullwidth": "全角转半角",
    "whitespace": "删除空白字符",
    "excel_decimal": "去除 Excel 数字 .0",
}


def analyze_part_number(value: Any) -> PartNumberNormalization:
    original = safe_text(value)
    transformations: list[str] = []
    text = unicodedata.normalize("NFKC", original)
    if text != original:
        transformations.append("fullwidth")
    if re.search(r"\s", text):
        transformations.append("whitespace")
    literal = re.sub(r"\s+", "", text).upper()
    if re.fullmatch(r"\d+\.0", literal):
        literal = literal[:-2]
        transformations.append("excel_decimal")

    working = literal
    version_match = re.fullmatch(
        r"(?P<base>.{4,})[-_/.](?P<version>\d{1,3}[A-Z0-9]*)",
        working,
    )
    if version_match:
        working = version_match.group("base")
        transformations.append("version_suffix")
    if re.search(r"[-_/.]", working):
        working = re.sub(r"[-_/.]+", "", working)
        transformations.append("separator")

    stripped = working.strip("0")
    if stripped != working:
        working = stripped or ("0" if working else "")
        transformations.append("outer_zero")

    return PartNumberNormalization(
        literal=literal,
        canonical=working,
        transformations=tuple(dict.fromkeys(transformations)),
    )


def canonical_part_number(value: Any) -> str:
    return analyze_part_number(value).canonical


def part_match_rule_labels(
    source_value: Any,
    rule_value: Any,
) -> tuple[str, ...]:
    source = analyze_part_number(source_value)
    rule = analyze_part_number(rule_value)
    codes = tuple(
        dict.fromkeys((*source.transformations, *rule.transformations))
    )
    return tuple(PART_RULE_LABELS[code] for code in codes)


def normalize_hs_code(value: Any) -> str:
    text = safe_text(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def to_number(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = safe_text(value).replace(",", "")
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def excel_serial_to_date(value: Any) -> dt.date | None:
    if isinstance(value, (int, float)):
        number = float(value)
        if 1 <= number <= 100000:
            return dt.date(1899, 12, 30) + dt.timedelta(days=int(number))
    text = safe_text(value)
    if not text:
        return None
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        return excel_serial_to_date(float(text))
    compact = re.sub(r"\s+", "", text)
    for date_format in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y%m%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%d/%m/%Y",
    ):
        try:
            return dt.datetime.strptime(compact, date_format).date()
        except ValueError:
            continue
    match = re.search(r"(20\d{6})", compact)
    if match:
        return dt.datetime.strptime(match.group(1), "%Y%m%d").date()
    return None


def find_column(headers: Iterable[Any], aliases: Iterable[str]) -> int | None:
    normalized_headers = [normalize_header(value) for value in headers]
    normalized_aliases = [normalize_header(value) for value in aliases]
    for alias in normalized_aliases:
        for index, header in enumerate(normalized_headers):
            if header == alias:
                return index
    for alias in normalized_aliases:
        for index, header in enumerate(normalized_headers):
            if alias and alias in header:
                return index
    return None


def adjacent_value(row: list[Any], index: int) -> str:
    text = safe_text(row[index]) if index < len(row) else ""
    if ":" in text:
        candidate = text.split(":", 1)[-1].strip()
        if candidate:
            return candidate
    for value in row[index + 1 :]:
        candidate = safe_text(value)
        if candidate:
            return candidate
    return ""


def normalize_document_no(value: Any) -> str:
    text = safe_text(value).upper().replace(" ", "")
    match = re.search(r"\b[A-Z]{1,4}\d{6,}(?:-\d+)?\b", text)
    return match.group(0) if match else ""


def normalize_delivery_way(value: Any) -> str:
    compact = re.sub(r"[\s/_-]+", "", safe_text(value).lower())
    if any(term in compact for term in ("air", "dhl", "fedex", "ups", "express")):
        return "Air"
    if any(term in compact for term in ("sea", "ocean", "fcl", "lcl")):
        return "Sea"
    if any(term in compact for term in ("road", "truck")):
        return "Road"
    if any(term in compact for term in ("rail", "train")):
        return "Rail"
    return collapse_spaces(value)

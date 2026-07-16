#!/usr/bin/env python3
import cgi
import copy
import html
import io
import json
import math
import os
import posixpath
import random
import re
import shutil
import sqlite3
import struct
import tempfile
import sys
import threading
import time
import traceback
import unicodedata
import urllib.parse
import uuid
import webbrowser
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from merge_bridge import service as merge_service
except ModuleNotFoundError:
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from merge_bridge import service as merge_service


APP_NAME = "SuriWorkDeclaration"
BASE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
STATIC_DIR = RESOURCE_DIR / "static"

DEFAULT_TEMPLATE = RESOURCE_DIR / "报关单 IMOS 空白模板.xlsx"
DEFAULT_RULES = RESOURCE_DIR / "报关单配置关系表.xlsx"
SOURCE_RULES = BASE_DIR.parent / "最新逻辑_请优先看" / "报关单配置关系表.xlsx"
LEGACY_RULES = RESOURCE_DIR / "2026+Daily+Export+List.xlsx"


def app_data_dir():
    override = os.environ.get("SURI_DATA_DIR")
    if override:
        return Path(override).expanduser()
    if os.name == "nt":
        root = os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local"
        return Path(root) / APP_NAME
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / APP_NAME
    return Path.home() / ".local" / "share" / APP_NAME


DATA_DIR = app_data_dir()
TEMPLATE_DIR = DATA_DIR / "templates"
RULES_DIR = DATA_DIR / "rules"
HISTORY_DIR = DATA_DIR / "history"
OUTPUT_DIR = DATA_DIR / "outputs"
DB_PATH = DATA_DIR / "app.db"

PUBLIC_TOKEN = os.environ.get("PUBLIC_TOKEN", "imos-demo")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "admin-demo")
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", str(25 * 1024 * 1024)))

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
EXCEL_COMPAT_NAMESPACES = {
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "xr6": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6",
    "xr10": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10",
    "x16r2": "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main",
}
ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing")
ET.register_namespace("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac")
ET.register_namespace("x16r2", "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
ET.register_namespace("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures")


def now_stamp():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def ensure_app_dirs():
    for directory in (DATA_DIR, TEMPLATE_DIR, RULES_DIR, HISTORY_DIR, OUTPUT_DIR):
        directory.mkdir(parents=True, exist_ok=True)


def active_template_path():
    uploaded = TEMPLATE_DIR / "template.xlsx"
    return uploaded if uploaded.exists() else DEFAULT_TEMPLATE


def active_rules_path():
    uploaded = RULES_DIR / "rules.xlsx"
    if uploaded.exists():
        return uploaded
    if DEFAULT_RULES.exists():
        return DEFAULT_RULES
    if SOURCE_RULES.exists():
        return SOURCE_RULES
    return LEGACY_RULES


def storage_meta_path():
    return DATA_DIR / "meta.json"


def declaration_config_status():
    ensure_app_dirs()
    meta_path = storage_meta_path()
    meta = json_loads(meta_path.read_text("utf-8"), {}) if meta_path.exists() else {}
    template_path = active_template_path()
    rules_path = active_rules_path()
    template_meta = meta.get("template") or {}
    rules_meta = meta.get("rules") or {}
    rule_count = 0
    conflict_count = 0
    if rules_path.exists():
        rules = load_rules(rules_path)
        rule_count = len([key for key in rules if not key.startswith("__")])
        conflict_count = len(rules.get("__part_conflicts", {}))
    return {
        "template": {
            "filename": (
                template_meta.get("filename")
                or (
                    "报关单空白模板.xlsx"
                    if template_path == DEFAULT_TEMPLATE
                    else template_path.name
                )
            ),
            "updatedAt": template_meta.get("updatedAt") or "内置默认",
            "source": "uploaded" if (TEMPLATE_DIR / "template.xlsx").exists() else "default",
        },
        "rules": {
            "filename": rules_meta.get("filename") or rules_path.name,
            "updatedAt": rules_meta.get("updatedAt") or "内置默认",
            "source": "uploaded" if (RULES_DIR / "rules.xlsx").exists() else "default",
            "recordCount": rule_count,
            "conflictCount": conflict_count,
        },
    }


def db_connect():
    ensure_app_dirs()
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with db_connect() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS history_records (
                id TEXT PRIMARY KEY,
                contract_no TEXT,
                created_at TEXT NOT NULL,
                invoice_name TEXT,
                packing_name TEXT,
                output_name TEXT,
                warnings_json TEXT NOT NULL,
                totals_json TEXT NOT NULL,
                file_paths_json TEXT NOT NULL,
                preview_json TEXT NOT NULL,
                recognized_json TEXT NOT NULL DEFAULT '[]'
            )
            """
        )


def json_dumps(value):
    return json.dumps(value, ensure_ascii=False, indent=2)


def json_loads(value, fallback):
    if not value:
        return fallback
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return fallback


def safe_filename(name, fallback):
    clean = Path(name or fallback).name
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", clean).strip(" .")
    return clean or fallback


def display_upload_name(path, fallback):
    name = Path(path).name or fallback
    return re.sub(r"^[0-9a-f]{32}-", "", name)


def copy_history_file(source, directory, prefix, original_name):
    suffix = Path(original_name or source).suffix.lower()
    destination = directory / f"{prefix}{suffix or '.xlsx'}"
    shutil.copy2(source, destination)
    return destination


def row_to_history(row, include_preview=False):
    file_paths = json_loads(row["file_paths_json"], {})
    result = {
        "id": row["id"],
        "contractNo": row["contract_no"] or "",
        "createdAt": row["created_at"],
        "invoiceName": row["invoice_name"] or "",
        "packingName": row["packing_name"] or "",
        "outputName": row["output_name"] or "",
        "warnings": json_loads(row["warnings_json"], []),
        "totals": json_loads(row["totals_json"], {}),
        "filePaths": file_paths,
    }
    if include_preview:
        result["preview"] = json_loads(row["preview_json"], {})
        result["recognizedFiles"] = json_loads(row["recognized_json"], [])
    return result


def normalize_history_datetime(value, end_of_day=False):
    text = safe_text(value).strip()
    if not text:
        return ""
    text = urllib.parse.unquote(text).replace("T", " ")
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        return f"{text} {'23:59:59' if end_of_day else '00:00:00'}"
    if re.match(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$", text):
        return f"{text}:59" if end_of_day else f"{text}:00"
    if re.match(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$", text):
        return text
    raise ValueError("历史查询时间格式不正确")


def create_history_record(preview, invoice_file, packing_file, classifications, record_id=None):
    init_db()
    record_id = record_id or uuid.uuid4().hex
    record_dir = HISTORY_DIR / record_id
    record_dir.mkdir(parents=True, exist_ok=True)
    invoice_name = safe_filename(display_upload_name(invoice_file, "invoice.xls"), "invoice.xls")
    packing_name = safe_filename(display_upload_name(packing_file, "packing.xls"), "packing.xls")
    invoice_dest = copy_history_file(invoice_file, record_dir, "invoice", invoice_name)
    packing_dest = copy_history_file(packing_file, record_dir, "packing", packing_name)
    preview_path = record_dir / "preview.json"
    preview_path.write_text(json_dumps(preview), "utf-8")
    file_paths = {
        "invoice": str(invoice_dest),
        "packing": str(packing_dest),
        "preview": str(preview_path),
        "output": "",
    }
    with db_connect() as conn:
        conn.execute(
            """
            INSERT INTO history_records (
                id, contract_no, created_at, invoice_name, packing_name,
                output_name, warnings_json, totals_json, file_paths_json,
                preview_json, recognized_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                record_id,
                preview.get("contractNo", ""),
                now_stamp(),
                invoice_name,
                packing_name,
                "",
                json_dumps(preview.get("warnings", [])),
                json_dumps(preview.get("totals", {})),
                json_dumps(file_paths),
                json_dumps(preview),
                json_dumps(classifications),
            ),
        )
    return record_id


def update_history_output(record_id, generated_path, output_name, preview):
    record_dir = HISTORY_DIR / record_id
    record_dir.mkdir(parents=True, exist_ok=True)
    output_path = record_dir / "output.xlsx"
    shutil.copy2(generated_path, output_path)
    with db_connect() as conn:
        row = conn.execute("SELECT file_paths_json FROM history_records WHERE id = ?", (record_id,)).fetchone()
        file_paths = json_loads(row["file_paths_json"], {}) if row else {}
        file_paths["output"] = str(output_path)
        conn.execute(
            """
            UPDATE history_records
               SET output_name = ?,
                   warnings_json = ?,
                   totals_json = ?,
                   file_paths_json = ?,
                   preview_json = ?
             WHERE id = ?
            """,
            (
                output_name,
                json_dumps(preview.get("warnings", [])),
                json_dumps(preview.get("totals", {})),
                json_dumps(file_paths),
                json_dumps(preview),
                record_id,
            ),
        )
    return output_path


def json_response(handler, status, payload):
    body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def undeclared_ignorable_prefixes(xml_bytes):
    text = xml_bytes.decode("utf-8", errors="ignore")
    declared = set(re.findall(r"\bxmlns:([A-Za-z_][\w.-]*)=", text))
    missing = set()
    for match in re.finditer(r"\bIgnorable=\"([^\"]+)\"", text):
        for prefix in match.group(1).split():
            if prefix and prefix not in declared:
                missing.add(prefix)
    return sorted(missing)


def ensure_excel_compat_namespace_declarations(xml_bytes):
    text = xml_bytes.decode("utf-8")
    required = set()
    for attr in ("Ignorable", "Requires"):
        for match in re.finditer(rf"\b{attr}=\"([^\"]+)\"", text):
            required.update(prefix for prefix in match.group(1).split() if prefix)
    declared = set(re.findall(r"\bxmlns:([A-Za-z_][\w.-]*)=", text))
    additions = [
        f' xmlns:{prefix}="{EXCEL_COMPAT_NAMESPACES[prefix]}"'
        for prefix in sorted(required - declared)
        if prefix in EXCEL_COMPAT_NAMESPACES
    ]
    if not additions:
        return xml_bytes

    root_start = text.find("<")
    if text.startswith("<?xml"):
        declaration_end = text.find("?>")
        root_start = text.find("<", declaration_end + 2)
    root_end = text.find(">", root_start)
    if root_start < 0 or root_end < 0:
        return xml_bytes
    text = text[:root_end] + "".join(additions) + text[root_end:]
    return text.encode("utf-8")


def serialize_excel_xml(root):
    xml_bytes = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    return ensure_excel_compat_namespace_declarations(xml_bytes)


def split_cell_ref(ref):
    match = re.match(r"\$?([A-Z]+)\$?(\d+)$", safe_text(ref).upper())
    if not match:
        return None
    return int(match.group(2)), column_to_number(match.group(1))


def worksheet_cell_bounds(root):
    max_row = 1
    max_col = 1
    for cell in root.findall(f".//{{{NS_MAIN}}}c"):
        pos = split_cell_ref(cell.attrib.get("r", ""))
        if not pos:
            continue
        row, col = pos
        max_row = max(max_row, row)
        max_col = max(max_col, col)
    return max_row, max_col


def worksheet_dimension_end(root):
    dimension = root.find(f"{{{NS_MAIN}}}dimension")
    if dimension is None:
        return None
    ref = dimension.attrib.get("ref", "")
    end_ref = ref.split(":")[-1]
    return split_cell_ref(end_ref)


def update_worksheet_dimension(root):
    max_row, max_col = worksheet_cell_bounds(root)
    dimension = root.find(f"{{{NS_MAIN}}}dimension")
    if dimension is None:
        dimension = ET.Element(f"{{{NS_MAIN}}}dimension")
        sheet_pr = root.find(f"{{{NS_MAIN}}}sheetPr")
        insert_at = list(root).index(sheet_pr) + 1 if sheet_pr is not None else 0
        root.insert(insert_at, dimension)
    dimension.attrib["ref"] = f"A1:{excel_col_name(max_col - 1)}{max_row}"


def validate_xlsx_file(path, require_dimensions=False):
    try:
        with zipfile.ZipFile(path, "r") as workbook:
            required = {
                "[Content_Types].xml",
                "xl/workbook.xml",
                "xl/_rels/workbook.xml.rels",
            }
            names = set(workbook.namelist())
            missing = required - names
            if missing:
                raise ValueError("缺少必要文件：" + "、".join(sorted(missing)))
            for name in workbook.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    content = workbook.read(name)
                    ET.fromstring(content)
                    missing_prefixes = undeclared_ignorable_prefixes(content) if name.endswith(".xml") else []
                    if missing_prefixes:
                        raise ValueError(f"{name} 存在未声明的 Excel 兼容性前缀：" + "、".join(missing_prefixes))
                    if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                        sheet_root = ET.fromstring(content)
                        end = worksheet_dimension_end(sheet_root)
                        max_row, max_col = worksheet_cell_bounds(sheet_root)
                        if require_dimensions and end is None:
                            raise ValueError(f"{name} 缺少工作表 dimension")
                        if require_dimensions and end is not None:
                            end_row, end_col = end
                            if end_row < max_row or end_col < max_col:
                                raise ValueError(f"{name} 工作表 dimension 未覆盖实际单元格")
    except zipfile.BadZipFile as exc:
        raise ValueError("生成文件不是有效的 Excel 工作簿") from exc
    except ET.ParseError as exc:
        raise ValueError("生成文件内部结构异常") from exc


def log_runtime_error(message):
    try:
        ensure_app_dirs()
        log_path = DATA_DIR / "error.log"
        with log_path.open("a", encoding="utf-8") as handle:
            handle.write(f"[{now_stamp()}] {message.rstrip()}\n\n")
    except Exception:
        pass


def safe_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def to_number(value, default=0.0):
    if value is None or value == "":
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return default
    try:
        return float(text)
    except ValueError:
        return default


def round2(value):
    return round(float(value or 0), 2)


def normalized_part_text(value):
    text = unicodedata.normalize("NFKC", safe_text(value)).upper()
    text = re.sub(r"\s+", "", text)
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text


def is_plausible_part_no(value):
    text = normalized_part_text(value)
    return bool(
        text
        and re.fullmatch(r"[A-Z0-9._/-]+", text)
        and re.search(r"\d", text)
    )


def part_match_keys(value):
    text = normalized_part_text(value)
    if not text:
        return []

    def canonicalize(candidate):
        compact = re.sub(r"[-_/.]+", "", candidate)
        if compact.isdigit():
            compact = compact.strip("0")
        return compact or ("0" if candidate else "")

    keys = []

    def add_key(candidate):
        key = canonicalize(candidate)
        if key and key not in keys:
            keys.append(key)

    add_key(text)
    version_match = re.fullmatch(r"(.+)[-_/.]([A-Z0-9]{1,4})", text)
    if (
        version_match
        and re.search(r"\d", version_match.group(1))
        and re.search(r"\d", version_match.group(2))
    ):
        add_key(version_match.group(1))
    return keys


def normalize_part(value):
    keys = part_match_keys(value)
    return keys[0] if keys else ""


def part_lookup(mapping, value, default=""):
    for key in part_match_keys(value):
        if key in mapping:
            return mapping[key]
    return default


def part_rule_lookup(rules, value):
    conflicts = rules.get("__part_conflicts", {})
    for key in part_match_keys(value):
        if key in conflicts:
            return {}, key, conflicts[key]
        if key in rules:
            return rules[key], key, None
    return {}, "", None


def normalize_transport_mode(value):
    text = safe_text(value)
    compact = re.sub(r"[\s/_-]+", "", text).lower()
    if any(term in compact for term in ("air", "dhl", "fedex", "ups", "express")):
        return "航空运输"
    if any(term in compact for term in ("sea", "ocean", "fcl", "lcl")):
        return "水路运输"
    return ""


def normalize_trade_term(value):
    text = safe_text(value).upper()
    match = re.search(r"\b(CPT|FOB|EXW)\b", text)
    if match:
        return match.group(1)
    return text.split()[0] if text.split() else ""


def normalize_hs(value):
    text = safe_text(value)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    return re.sub(r"\D", "", text)


def normalize_brand(value):
    text = safe_text(value)
    upper = text.upper().replace(" ", "")
    if not text or text == "0":
        return "无"
    if upper == "FOMOCO":
        return "FoMoCo"
    if upper == "FORD":
        return "Ford"
    if upper == "VOLVO":
        return "VOLVO"
    if upper == "JAGUARLANDROVER":
        return "JAGUAR LAND ROVER"
    if text in ("无", "N/A", "NA"):
        return "无"
    return text


def excel_serial_from_yyyymmdd(text):
    import datetime as _dt

    base = _dt.date(1899, 12, 30)
    for ymd in re.findall(r"20\d{6}", text or ""):
        year = int(ymd[:4])
        month = int(ymd[4:6])
        day = int(ymd[6:8])
        try:
            parsed = _dt.date(year, month, day)
        except ValueError:
            continue
        return (parsed - base).days
    return ""


def excel_serial_from_date_value(value):
    import datetime as _dt

    if isinstance(value, (int, float)):
        number = float(value)
        return number if 1 <= number <= 100000 else ""

    text = safe_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d+(?:\.\d+)?", text):
        number = float(text)
        return number if 1 <= number <= 100000 else ""
    compact = re.sub(r"\s+", "", text)
    for date_format in (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%d/%m/%Y",
    ):
        try:
            parsed = _dt.datetime.strptime(compact, date_format).date()
            return (parsed - _dt.date(1899, 12, 30)).days
        except ValueError:
            continue
    return excel_serial_from_yyyymmdd(compact)


def adjacent_value(row, idx):
    text = safe_text(row[idx]) if idx < len(row) else ""
    if ":" in text:
        candidate = text.split(":", 1)[-1].strip()
        if candidate:
            return candidate
    for value in row[idx + 1:]:
        candidate = safe_text(value)
        if candidate:
            return candidate
    return ""


def normalize_document_no(value):
    text = safe_text(value).upper().replace(" ", "")
    match = re.search(r"\b[A-Z]{1,4}\d{6,}(?:-\d+)?\b", text)
    return match.group(0) if match else ""


def is_business_label(value):
    text = safe_text(value).lower()
    labels = (
        "bill to", "ship to", "serial", "page", "our ref", "customer",
        "delivery", "term", "telephone", "fax", "post code", "address",
    )
    return not text or any(label in text for label in labels)


def extract_party_after_label(sheet, label_pattern, max_rows=30):
    for row_idx, row in enumerate(sheet[:max_rows]):
        for col_idx, value in enumerate(row):
            if re.search(label_pattern, safe_text(value), re.I):
                same_row = adjacent_value(row, col_idx)
                if same_row and not is_business_label(same_row):
                    return same_row
                for next_row in sheet[row_idx + 1: min(len(sheet), row_idx + 6)]:
                    for idx in (col_idx, col_idx + 1):
                        if idx < len(next_row):
                            candidate = safe_text(next_row[idx])
                            if candidate and not is_business_label(candidate):
                                return candidate
    return ""


def extract_block_after_label(sheet, label_pattern, max_rows=30, block_rows=6):
    values = []
    stop_patterns = (
        r"bill\s*to", r"ship\s*to", r"payment", r"delivery\s*term",
        r"delivery\s*way", r"customer\s*code", r"serial\s*no",
    )
    for row_idx, row in enumerate(sheet[:max_rows]):
        for col_idx, value in enumerate(row):
            if not re.search(label_pattern, safe_text(value), re.I):
                continue
            for offset, current_row in enumerate(sheet[row_idx:min(len(sheet), row_idx + block_rows + 1)]):
                if offset:
                    row_text = " ".join(safe_text(v) for v in current_row)
                    if any(re.search(pattern, row_text, re.I) for pattern in stop_patterns):
                        break
                start_col = col_idx + 1 if offset == 0 else col_idx
                for idx in range(start_col, len(current_row)):
                    text = safe_text(current_row[idx])
                    if text:
                        values.append(text)
            return values
    return values


COUNTRY_ALIASES = [
    ("UNITED STATES OF AMERICA", "美国"),
    ("UNITED STATES", "美国"),
    ("U.S.A", "美国"),
    ("USA", "美国"),
    ("AMERICA", "美国"),
    ("VIET NAM", "越南"),
    ("VIETNAM", "越南"),
    ("ITALIA", "意大利"),
    ("ITALY", "意大利"),
    ("CHINA", "中国"),
    ("GERMANY", "德国"),
    ("DEUTSCHLAND", "德国"),
    ("FRANCE", "法国"),
    ("UNITED KINGDOM", "英国"),
    ("GREAT BRITAIN", "英国"),
    ("BRITAIN", "英国"),
    ("UK", "英国"),
    ("SPAIN", "西班牙"),
    ("POLAND", "波兰"),
    ("MEXICO", "墨西哥"),
    ("INDIA", "印度"),
    ("JAPAN", "日本"),
    ("KOREA", "韩国"),
    ("SOUTH KOREA", "韩国"),
    ("THAILAND", "泰国"),
    ("MALAYSIA", "马来西亚"),
    ("INDONESIA", "印度尼西亚"),
    ("TURKEY", "土耳其"),
    ("BRAZIL", "巴西"),
    ("CANADA", "加拿大"),
    ("AUSTRALIA", "澳大利亚"),
    ("NETHERLANDS", "荷兰"),
    ("HOLLAND", "荷兰"),
    ("BELGIUM", "比利时"),
    ("SWEDEN", "瑞典"),
    ("CZECH", "捷克"),
]


def infer_country_from_text(context, rules, allow_china=True):
    text = safe_text(context)
    if not text:
        return ""
    lowered = text.lower()
    for raw, standard in rules.get("__standardization", {}).get("国家/地区", []):
        if not allow_china and safe_text(standard) == "中国":
            continue
        raw_text = safe_text(raw).lower()
        if raw_text and raw_text in lowered:
            return safe_text(standard)
    upper = text.upper()
    for alias, standard in COUNTRY_ALIASES:
        if not allow_china and standard == "中国":
            continue
        if re.search(rf"\b{re.escape(alias)}\b", upper):
            return standard
    return ""


class XlsWorkbook:
    """Small BIFF8 reader for the old .xls files used by this workflow."""

    END_OF_CHAIN = 0xFFFFFFFE
    FREE_SECTOR = 0xFFFFFFFF

    def __init__(self, data):
        self.data = data
        self.workbook_stream = self._read_workbook_stream()
        self.shared_strings = self._read_sst()
        self.sheet_metadata = self._read_sheet_metadata()
        self.sheets = self._read_sheets()

    @classmethod
    def load(cls, path):
        return cls(Path(path).read_bytes())

    def _u16(self, offset):
        return struct.unpack_from("<H", self.data, offset)[0]

    def _u32(self, offset):
        return struct.unpack_from("<I", self.data, offset)[0]

    def _sector(self, idx, sector_size):
        start = 512 + idx * sector_size
        return self.data[start:start + sector_size]

    def _read_workbook_stream(self):
        if self.data[:8] != bytes.fromhex("d0cf11e0a1b11ae1"):
            raise ValueError("不是有效的老式 .xls 文件")

        sector_size = 1 << self._u16(30)
        fat_sector_count = self._u32(44)
        first_dir_sector = self._u32(48)
        difat = []
        for i in range(109):
            value = self._u32(76 + i * 4)
            if value != self.FREE_SECTOR:
                difat.append(value)

        fat = []
        for sector_idx in difat[:fat_sector_count]:
            block = self._sector(sector_idx, sector_size)
            fat.extend(struct.unpack_from("<I", block, i)[0] for i in range(0, len(block), 4))

        def chain(start):
            out, seen, current = [], set(), start
            while (
                current not in (self.END_OF_CHAIN, self.FREE_SECTOR)
                and current not in seen
                and current < len(fat)
            ):
                seen.add(current)
                out.append(current)
                current = fat[current]
            return out

        directory = b"".join(self._sector(i, sector_size) for i in chain(first_dir_sector))
        workbook_start = workbook_size = None
        for offset in range(0, len(directory), 128):
            entry = directory[offset:offset + 128]
            if len(entry) < 128:
                continue
            name_len = struct.unpack_from("<H", entry, 64)[0]
            name = entry[:max(0, name_len - 2)].decode("utf-16le", "ignore")
            if name in ("Workbook", "Book"):
                workbook_start = struct.unpack_from("<I", entry, 116)[0]
                workbook_size = struct.unpack_from("<I", entry, 120)[0]
                break

        if workbook_start is None:
            raise ValueError("未找到 Workbook 数据流")
        stream = b"".join(self._sector(i, sector_size) for i in chain(workbook_start))
        return stream[:workbook_size]

    def _records(self):
        pos = 0
        data = self.workbook_stream
        while pos + 4 <= len(data):
            rec_type, size = struct.unpack_from("<HH", data, pos)
            payload = data[pos + 4:pos + 4 + size]
            yield rec_type, payload
            pos += 4 + size

    def _read_sst(self):
        segments = []
        collecting = False
        for rec_type, payload in self._records():
            if rec_type == 0x00FC:
                segments = [payload]
                collecting = True
            elif rec_type == 0x003C and collecting:
                segments.append(payload)
            elif collecting:
                break

        if not segments or len(segments[0]) < 8:
            return []

        class SegReader:
            def __init__(self, segs):
                self.segs = segs
                self.seg = 0
                self.pos = 8

            def _advance(self):
                self.seg += 1
                self.pos = 0
                return self.seg < len(self.segs)

            def read(self, count):
                out = bytearray()
                while count > 0 and self.seg < len(self.segs):
                    current = self.segs[self.seg]
                    if self.pos >= len(current):
                        if not self._advance():
                            break
                        continue
                    take = min(count, len(current) - self.pos)
                    out.extend(current[self.pos:self.pos + take])
                    self.pos += take
                    count -= take
                return bytes(out)

            def read_characters(self, count, is_16bit):
                parts = []
                remaining = count
                while remaining > 0 and self.seg < len(self.segs):
                    current = self.segs[self.seg]
                    if self.pos >= len(current):
                        if not self._advance():
                            break
                        if self.pos < len(self.segs[self.seg]):
                            flags = self.segs[self.seg][self.pos]
                            self.pos += 1
                            is_16bit = bool(flags & 0x01)
                        continue
                    width = 2 if is_16bit else 1
                    chars_here = min(remaining, (len(current) - self.pos) // width)
                    if chars_here <= 0:
                        self.pos = len(current)
                        continue
                    raw = current[self.pos:self.pos + chars_here * width]
                    self.pos += chars_here * width
                    remaining -= chars_here
                    parts.append(raw.decode("utf-16le" if is_16bit else "latin1", "ignore"))
                return "".join(parts)

        unique_count = struct.unpack_from("<I", segments[0], 4)[0]
        reader = SegReader(segments)
        strings = []
        for _ in range(unique_count):
            header = reader.read(3)
            if len(header) < 3:
                break
            char_count = struct.unpack_from("<H", header, 0)[0]
            flags = header[2]
            rich_runs = 0
            ext_size = 0
            if flags & 0x08:
                raw = reader.read(2)
                if len(raw) == 2:
                    rich_runs = struct.unpack("<H", raw)[0]
            if flags & 0x04:
                raw = reader.read(4)
                if len(raw) == 4:
                    ext_size = struct.unpack("<I", raw)[0]
            strings.append(reader.read_characters(char_count, bool(flags & 0x01)))
            if rich_runs:
                reader.read(rich_runs * 4)
            if ext_size:
                reader.read(ext_size)
        return strings

    @staticmethod
    def _decode_rk(raw):
        multiplied = raw & 0x01
        is_integer = raw & 0x02
        value_bits = raw & 0xFFFFFFFC
        if is_integer:
            if value_bits & 0x80000000:
                value_bits -= 0x100000000
            value = value_bits >> 2
        else:
            value = struct.unpack("<d", struct.pack("<Q", value_bits << 32))[0]
        return value / 100 if multiplied else value

    @staticmethod
    def _formula_result(raw):
        if len(raw) < 8:
            return None
        if raw[6:8] == b"\xFF\xFF":
            return None
        value = struct.unpack("<d", raw[:8])[0]
        if value != value:
            return None
        return value

    @staticmethod
    def _decode_label_text(payload):
        if len(payload) < 8:
            return ""
        char_count = struct.unpack_from("<H", payload, 6)[0]
        remaining = len(payload) - 8
        if remaining >= char_count and (remaining == char_count or payload[8] not in (0, 1, 4, 5, 8, 9, 12, 13)):
            return payload[8:8 + char_count].decode("latin1", "ignore")
        if remaining >= char_count * 2 and remaining == char_count * 2:
            return payload[8:8 + char_count * 2].decode("utf-16le", "ignore")
        if len(payload) < 9:
            return ""
        flags = payload[8]
        offset = 9
        width = 2 if flags & 0x01 else 1
        raw = payload[offset:offset + char_count * width]
        return raw.decode("utf-16le" if flags & 0x01 else "latin1", "ignore")

    def _read_sheet_metadata(self):
        sheets = []
        for rec_type, payload in self._records():
            if rec_type != 0x0085 or len(payload) < 8:
                continue
            name_length = payload[6]
            flags = payload[7]
            width = 2 if flags & 0x01 else 1
            raw = payload[8:8 + name_length * width]
            sheets.append({
                "name": raw.decode("utf-16le" if width == 2 else "latin1", "ignore"),
                "state": "visible" if payload[4] == 0 else "hidden",
            })
        return sheets

    def _read_sheets(self):
        sheets = []
        current = None
        pending_formula_cell = None
        for rec_type, payload in self._records():
            if rec_type == 0x0809 and len(payload) >= 4:
                stream_type = struct.unpack_from("<H", payload, 2)[0]
                if stream_type == 0x0010:
                    current = {}
                    pending_formula_cell = None
            elif rec_type == 0x000A:
                if current is not None:
                    sheets.append(current)
                    current = None
                pending_formula_cell = None
            elif current is None:
                continue
            elif rec_type == 0x0204 and len(payload) >= 9:
                row, col, _ = struct.unpack_from("<HHH", payload, 0)
                current[(row, col)] = self._decode_label_text(payload)
            elif rec_type == 0x00FD and len(payload) >= 10:
                row, col, _, idx = struct.unpack_from("<HHHI", payload, 0)
                current[(row, col)] = self.shared_strings[idx] if idx < len(self.shared_strings) else ""
            elif rec_type == 0x0203 and len(payload) >= 14:
                row, col, _ = struct.unpack_from("<HHH", payload, 0)
                current[(row, col)] = struct.unpack_from("<d", payload, 6)[0]
            elif rec_type == 0x027E and len(payload) >= 10:
                row, col, _ = struct.unpack_from("<HHH", payload, 0)
                current[(row, col)] = self._decode_rk(struct.unpack_from("<I", payload, 6)[0])
            elif rec_type == 0x00BD and len(payload) >= 6:
                row, first_col, last_col = struct.unpack_from("<HHH", payload, 0)
                offset = 6
                for col in range(first_col, last_col + 1):
                    if offset + 6 <= len(payload):
                        _, raw = struct.unpack_from("<HI", payload, offset)
                        current[(row, col)] = self._decode_rk(raw)
                    offset += 6
            elif rec_type == 0x0006 and len(payload) >= 14:
                row, col, _ = struct.unpack_from("<HHH", payload, 0)
                value = self._formula_result(payload[6:14])
                pending_formula_cell = (row, col)
                if value is not None:
                    current[(row, col)] = value
            elif rec_type == 0x0207 and pending_formula_cell and len(payload) >= 3:
                length = struct.unpack_from("<H", payload, 0)[0]
                flags = payload[2]
                raw = payload[3:3 + length * (2 if flags & 1 else 1)]
                current[pending_formula_cell] = raw.decode("utf-16le" if flags & 1 else "latin1", "ignore")
                pending_formula_cell = None
        return sheets


def column_to_number(col):
    result = 0
    for ch in col:
        result = result * 26 + ord(ch.upper()) - 64
    return result


def cell_ref(row, col):
    name = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        name = chr(65 + rem) + name
    return f"{name}{row + 1}"


class XlsxBook:
    def __init__(self, path):
        self.path = Path(path)
        self.zip = zipfile.ZipFile(self.path)
        self.shared_strings = self._shared_strings()
        self.sheet_states = {}
        self.sheet_map = self._sheet_paths()

    def close(self):
        self.zip.close()

    def _shared_strings(self):
        if "xl/sharedStrings.xml" not in self.zip.namelist():
            return []
        root = ET.fromstring(self.zip.read("xl/sharedStrings.xml"))
        return ["".join(t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t")) for si in root.findall(f"{{{NS_MAIN}}}si")]

    def _sheet_paths(self):
        workbook = ET.fromstring(self.zip.read("xl/workbook.xml"))
        rels = ET.fromstring(self.zip.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        result = {}
        for sheet in workbook.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = rel_map[rid].lstrip("/")
            name = sheet.attrib["name"]
            result[name] = "xl/" + target if not target.startswith("xl/") else target
            self.sheet_states[name] = sheet.attrib.get("state", "visible")
        return result

    def sheet_values(self, sheet_name):
        path = self.sheet_map[sheet_name]
        root = ET.fromstring(self.zip.read(path))
        values = {}
        for cell in root.findall(f".//{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r")
            if not ref:
                continue
            values[ref] = self._cell_value(cell)
        return values

    def _cell_value(self, cell):
        cell_type = cell.attrib.get("t")
        if cell_type == "inlineStr":
            inline = cell.find(f"{{{NS_MAIN}}}is")
            return "".join(t.text or "" for t in inline.findall(f".//{{{NS_MAIN}}}t")) if inline is not None else ""
        value = cell.find(f"{{{NS_MAIN}}}v")
        if value is None:
            return ""
        text = value.text or ""
        if cell_type == "s":
            idx = int(float(text)) if text else -1
            return self.shared_strings[idx] if 0 <= idx < len(self.shared_strings) else ""
        return text

    def rows(self, sheet_name):
        values = self.sheet_values(sheet_name)
        grid = {}
        max_row = max_col = 0
        for ref, value in values.items():
            match = re.match(r"([A-Z]+)(\d+)", ref)
            if not match:
                continue
            col = column_to_number(match.group(1)) - 1
            row = int(match.group(2)) - 1
            grid[(row, col)] = value
            max_row, max_col = max(max_row, row), max(max_col, col)
        return [[grid.get((r, c), "") for c in range(max_col + 1)] for r in range(max_row + 1)]


def read_spreadsheet_entries(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".xls":
        workbook = XlsWorkbook.load(path)
        return [
            {
                "name": (
                    workbook.sheet_metadata[idx]["name"]
                    if idx < len(workbook.sheet_metadata)
                    else f"Sheet {idx + 1}"
                ),
                "state": (
                    workbook.sheet_metadata[idx]["state"]
                    if idx < len(workbook.sheet_metadata)
                    else "visible"
                ),
                "cells": sheet,
            }
            for idx, sheet in enumerate(workbook.sheets)
        ]
    if suffix == ".xlsx":
        book = XlsxBook(path)
        try:
            return [
                {
                    "name": name,
                    "state": book.sheet_states.get(name, "visible"),
                    "cells": {
                        (r, c): value
                        for r, row in enumerate(book.rows(name))
                        for c, value in enumerate(row)
                        if safe_text(value)
                    },
                }
                for name in book.sheet_map
            ]
        finally:
            book.close()
    raise ValueError("仅支持 .xls 或 .xlsx 文件")


def read_spreadsheet(path):
    return [entry["cells"] for entry in read_spreadsheet_entries(path)]


def matrix_from_sheet(sheet):
    if not sheet:
        return []
    max_row = max(row for row, _ in sheet)
    max_col = max(col for _, col in sheet)
    return [[sheet.get((r, c), "") for c in range(max_col + 1)] for r in range(max_row + 1)]


def row_contains(row, *terms):
    text = " ".join(safe_text(v).lower() for v in row)
    return all(term.lower() in text for term in terms)


EXCEL_ERROR_VALUES = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NUM!", "#NAME?", "#NULL!", "N/A", "NA"}


def normalized_header(row):
    return [safe_text(v).lower().replace(" ", "").replace("_", "") for v in row]


def is_enabled_value(value):
    text = safe_text(value).strip().lower()
    return text not in ("否", "no", "n", "false", "0", "禁用", "停用")


def find_header_index_by_terms(rows, required_groups):
    for idx, row in enumerate(rows):
        compact = "".join(safe_text(v).lower().replace(" ", "").replace("_", "") for v in row)
        if all(any(term.lower().replace(" ", "").replace("_", "") in compact for term in group) for group in required_groups):
            return idx
    return None


def header_col(headers, *terms):
    normalized_terms = [term.lower().replace(" ", "").replace("_", "") for term in terms]
    compact_headers = [safe_text(v).lower().replace(" ", "").replace("_", "") for v in headers]
    for term in normalized_terms:
        for idx, value in enumerate(compact_headers):
            if term == value:
                return idx
    for term in normalized_terms:
        for idx, value in enumerate(compact_headers):
            if term in value:
                return idx
    return None


def standardize_config_value(rules, category, value, allow_contains=False):
    text = safe_text(value)
    if not text:
        return ""
    entries = rules.get("__standardization", {}).get(category, [])
    lowered = text.lower()
    for raw, standard in entries:
        if lowered == safe_text(raw).lower():
            return standard
    if allow_contains:
        for raw, standard in entries:
            raw_text = safe_text(raw).lower()
            if raw_text and raw_text in lowered:
                return standard
    return ""


def find_header_row(rows, required_any_groups):
    for idx, row in enumerate(rows):
        text = " ".join(safe_text(v).lower() for v in row)
        compact = text.replace(" ", "").replace("_", "")
        matched = True
        for group in required_any_groups:
            if not any(term.lower().replace(" ", "").replace("_", "") in compact for term in group):
                matched = False
                break
        if matched:
            return idx
    return None


def excel_col_name(col):
    name = ""
    col += 1
    while col:
        col, rem = divmod(col - 1, 26)
        name = chr(65 + rem) + name
    return name


def anomaly(file_kind, sheet_idx, row_idx, col_idx, field, value, reason):
    cell = f"{excel_col_name(col_idx)}{row_idx + 1}"
    text = safe_text(value)
    return {
        "file": file_kind,
        "sheet": f"Sheet {sheet_idx + 1}",
        "row": row_idx + 1,
        "column": excel_col_name(col_idx),
        "cell": cell,
        "field": field,
        "value": text,
        "message": f"{file_kind} {cell}（{field}）{reason}" + (f"，当前值：{text}" if text else ""),
    }


def numeric_value_state(value, require_positive=False):
    text = safe_text(value)
    if not text:
        return None, "为空"
    if text.upper() in EXCEL_ERROR_VALUES or text.startswith("#"):
        return None, "是 Excel 异常值"
    try:
        number = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None, "不是有效数字"
    if require_positive and number <= 0:
        return number, "必须大于 0"
    return number, ""


def text_value_state(value):
    text = safe_text(value)
    if not text:
        return "为空"
    if text.upper() in EXCEL_ERROR_VALUES or text.startswith("#"):
        return "是 Excel 异常值"
    return ""


def append_numeric_anomaly(anomalies, file_kind, sheet_idx, row_idx, row, col_idx, field, require_positive=True):
    if col_idx is None or col_idx >= len(row):
        anomalies.append(anomaly(file_kind, sheet_idx, row_idx, 0, field, "", "缺少字段列"))
        return
    _, reason = numeric_value_state(row[col_idx], require_positive=require_positive)
    if reason:
        anomalies.append(anomaly(file_kind, sheet_idx, row_idx, col_idx, field, row[col_idx], reason))


def append_text_anomaly(anomalies, file_kind, sheet_idx, row_idx, row, col_idx, field):
    if col_idx is None or col_idx >= len(row):
        anomalies.append(anomaly(file_kind, sheet_idx, row_idx, 0, field, "", "缺少字段列"))
        return
    reason = text_value_state(row[col_idx])
    if reason:
        anomalies.append(anomaly(file_kind, sheet_idx, row_idx, col_idx, field, row[col_idx], reason))


def parse_invoice(path):
    sheets = [matrix_from_sheet(s) for s in read_spreadsheet(path)]
    invoice_sheet = None
    header_index = None
    invoice_sheet_idx = 0
    for sheet_idx, sheet in enumerate(sheets):
        idx = find_header_row(sheet, [("qad pn", "part no"), ("qty", "quantity"), ("amount",)])
        if idx is not None:
            invoice_sheet = sheet
            invoice_sheet_idx = sheet_idx
            header_index = idx
            break
    if not invoice_sheet:
        raise ValueError("Invoice 中没有找到包含 QAD PN 和 Quantity/Qty 的明细表")

    header = normalized_header(invoice_sheet[header_index])

    def find_col(*names):
        normalized_names = [name.lower().replace(" ", "").replace("_", "") for name in names]
        for name in normalized_names:
            for idx, value in enumerate(header):
                if name == value:
                    return idx
        for name in normalized_names:
            for idx, value in enumerate(header):
                if name in value:
                    return idx
        return None

    def find_raw_col(*names):
        raw_header = [safe_text(v).lower() for v in invoice_sheet[header_index]]
        for name in names:
            for idx, value in enumerate(raw_header):
                if name.lower() == value:
                    return idx
            for idx, value in enumerate(raw_header):
                if name.lower() in value:
                    return idx
        return None

    cols = {
        "part": find_col("qad pn", "part no"),
        "description_en": find_col("description"),
        "quantity": find_col("qty", "quantity"),
        "unit_price": find_col("up"),
        "po": find_col("po_no"),
        "amount": find_col("amount"),
        "currency": find_col("currency"),
        "goods_name": find_raw_col("商品名称", "货物名称", "goods name"),
        "hs_code": find_col("hs code", "hscode", "商品编号"),
        "brand": find_raw_col("品牌", "brand"),
        "model": find_raw_col("车型", "model"),
    }
    description_cols = [i for i, value in enumerate(header) if "description" in value]
    if len(description_cols) > 1:
        cols["goods_name"] = description_cols[-1]

    contract = ""
    consignee = ""
    trade_term = "CPT"
    transport_mode = normalize_transport_mode(Path(path).name)
    currency = "USD"
    filename_date = excel_serial_from_yyyymmdd(Path(path).name)
    pickup_date = ""
    for row in invoice_sheet[:30]:
        for idx, value in enumerate(row):
            text = safe_text(value)
            if re.search(r"serial\s*no", text, re.I):
                contract = normalize_document_no(adjacent_value(row, idx)) or contract
            if not contract:
                contract = normalize_document_no(text) or contract
            if "CASCO Imos" in text:
                consignee = text.strip()
            if re.search(r"delivery\s*term", text, re.I):
                candidate = text.split(":", 1)[-1] if ":" in text else ""
                if not candidate and idx + 1 < len(row):
                    candidate = row[idx + 1]
                trade_term = normalize_trade_term(candidate) or trade_term
            if re.search(r"delivery\s*(way|mode|method)|transport", text, re.I):
                candidate = text.split(":", 1)[-1] if ":" in text else ""
                if not candidate and idx + 1 < len(row):
                    candidate = row[idx + 1]
                transport_mode = normalize_transport_mode(candidate) or transport_mode
            if re.search(r"pick\s*up\s*date|pickup\s*date", text, re.I):
                pickup_date = excel_serial_from_date_value(adjacent_value(row, idx)) or pickup_date
            if text in ("USD", "EUR", "CNY"):
                currency = text
    consignee = (
        extract_party_after_label(invoice_sheet, r"bill\s*to")
        or extract_party_after_label(invoice_sheet, r"ship\s*to")
        or consignee
    )
    bill_to_context = "\n".join(extract_block_after_label(invoice_sheet, r"bill\s*to"))
    ship_to_context = "\n".join(extract_block_after_label(invoice_sheet, r"ship\s*to"))

    items = []
    anomalies = []
    for row_idx, row in enumerate(invoice_sheet[header_index + 1:], header_index + 1):
        part = safe_text(row[cols["part"]]) if cols["part"] is not None and cols["part"] < len(row) else ""
        if not part or not re.search(r"\d", part):
            continue
        quantity = to_number(row[cols["quantity"]]) if cols["quantity"] is not None and cols["quantity"] < len(row) else 0
        unit_price = to_number(row[cols["unit_price"]]) if cols["unit_price"] is not None and cols["unit_price"] < len(row) else 0
        amount = to_number(row[cols["amount"]]) if cols["amount"] is not None and cols["amount"] < len(row) else 0
        is_applied_row = quantity > 0 or amount > 0
        if is_applied_row:
            append_text_anomaly(anomalies, "Invoice", invoice_sheet_idx, row_idx, row, cols["part"], "QAD PN")
            append_numeric_anomaly(anomalies, "Invoice", invoice_sheet_idx, row_idx, row, cols["quantity"], "Quantity", True)
            append_numeric_anomaly(anomalies, "Invoice", invoice_sheet_idx, row_idx, row, cols["amount"], "Amount", True)
            append_numeric_anomaly(anomalies, "Invoice", invoice_sheet_idx, row_idx, row, cols["unit_price"], "UP", True)
            append_text_anomaly(anomalies, "Invoice", invoice_sheet_idx, row_idx, row, cols["currency"], "Currency")
        if amount == 0 and quantity and unit_price:
            amount = quantity * unit_price
        if quantity <= 0 and amount <= 0:
            continue
        item_currency = safe_text(row[cols["currency"]]) if cols["currency"] is not None and cols["currency"] < len(row) else currency
        currency = item_currency or currency
        items.append({
            "partNo": part,
            "descriptionEn": safe_text(row[cols["description_en"]]) if cols["description_en"] is not None and cols["description_en"] < len(row) else "",
            "quantity": quantity,
            "unitPrice": unit_price,
            "amount": amount,
            "currency": item_currency or currency,
            "poNo": safe_text(row[cols["po"]]) if cols["po"] is not None and cols["po"] < len(row) else "",
            "goodsName": "",
            "hsCode": "",
            "brand": "",
            "model": safe_text(row[cols["model"]]) if cols["model"] is not None and cols["model"] < len(row) else "",
            "sourceRow": row_idx + 1,
            "sourceSheet": f"Sheet {invoice_sheet_idx + 1}",
        })

    return {
        "contractNo": contract or normalize_document_no(Path(path).stem) or Path(path).stem.split()[0],
        "consignee": consignee or "CASCO Imos Italia S.R.L.",
        "billToContext": bill_to_context,
        "shipToContext": ship_to_context,
        "countryContext": "\n".join(
            safe_text(value)
            for row in invoice_sheet[:30]
            for value in row
            if safe_text(value)
        ),
        "tradeTerm": trade_term or "CPT",
        "transportMode": transport_mode,
        "currency": currency or "USD",
        "exportDateSerial": pickup_date or filename_date,
        "items": items,
        "anomalies": anomalies,
    }


def parse_packing(path, expected_contract_no=""):
    entries = read_spreadsheet_entries(path)
    package_count = 0
    gross_weight = 0.0
    net_weight = 0.0
    part_weights = {}
    part_gross_weights = {}
    export_date = excel_serial_from_yyyymmdd(Path(path).name)
    anomalies = []
    parser_warnings = []
    expected_key = re.sub(r"\s+", "", safe_text(expected_contract_no)).upper()
    candidates = []

    for sheet_idx, entry in enumerate(entries):
        sheet = matrix_from_sheet(entry["cells"])
        header_idx = find_header_row(
            sheet,
            [
                ("qad pn", "part no", "fg pn", "fg_pn", "fgpn"),
                ("qty", "quantity"),
                ("n.w", "nw"),
                ("g.w", "gw"),
            ],
        )
        if header_idx is None:
            continue
        sheet_text = re.sub(
            r"\s+",
            "",
            " ".join(safe_text(value) for row in sheet for value in row),
        ).upper()
        candidates.append({
            "index": sheet_idx,
            "name": entry.get("name") or f"Sheet {sheet_idx + 1}",
            "state": entry.get("state", "visible"),
            "sheet": sheet,
            "header": header_idx,
            "contractMatch": bool(expected_key and expected_key in sheet_text),
        })

    if expected_key and any(candidate["contractMatch"] for candidate in candidates):
        candidates = [candidate for candidate in candidates if candidate["contractMatch"]]
    visible_candidates = [candidate for candidate in candidates if candidate["state"] == "visible"]
    if visible_candidates:
        candidates = visible_candidates

    selected = candidates[0] if candidates else None
    source_sheet = selected["name"] if selected else ""
    if selected:
        sheet_idx = selected["index"]
        sheet = selected["sheet"]
        header_idx = selected["header"]
        header = normalized_header(sheet[header_idx])

        def find_col(*terms):
            normalized_terms = [term.lower().replace(" ", "").replace("_", "") for term in terms]
            for term in terms:
                normalized_term = term.lower().replace(" ", "").replace("_", "")
                for col, value in enumerate(header):
                    if value == normalized_term:
                        return col
            for term in normalized_terms:
                for col, value in enumerate(header):
                    if term in value:
                        return col
            return None

        part_col = find_col("qad pn", "part no", "fg pn", "fg_pn", "fgpn")
        qty_col = find_col("qty", "quantity")
        net_col = find_col("n.w.(kg)", "n.w", "nw")
        gross_col = find_col("g.w.(kg)", "g.w", "gw")
        pallet_col = find_col("pallet item", "pallet no", "pallet")

        required_columns = {
            "料号": part_col,
            "数量": qty_col,
            "净重": net_col,
            "毛重": gross_col,
        }
        missing_columns = [name for name, col in required_columns.items() if col is None]
        if missing_columns:
            parser_warnings.append(
                f"Packing list 工作表 {source_sheet} 缺少必要列：{'、'.join(missing_columns)}，请复核"
            )
            selected = None

    if selected:
        pallets = set()
        summed_net_weight = 0.0
        summed_gross_weight = 0.0
        total_net_weight = 0.0
        total_gross_weight = 0.0
        for row_idx, row in enumerate(sheet[header_idx + 1:], header_idx + 1):
            label = " ".join(safe_text(v).lower() for v in row)
            if "total" in label:
                total_packages = (
                    int(to_number(row[pallet_col]))
                    if pallet_col is not None and pallet_col < len(row) and to_number(row[pallet_col])
                    else 0
                )
                if total_packages:
                    package_count = total_packages
                total_net = to_number(row[net_col]) if net_col < len(row) else 0
                total_gross = to_number(row[gross_col]) if gross_col < len(row) else 0
                if total_net:
                    total_net_weight = round2(total_net)
                if total_gross:
                    total_gross_weight = round2(total_gross)
                continue

            if max(part_col, qty_col, net_col, gross_col) >= len(row):
                continue
            part = safe_text(row[part_col])
            quantity = to_number(row[qty_col])
            if quantity < 1 or not part:
                continue
            append_text_anomaly(anomalies, "Packing list", sheet_idx, row_idx, row, part_col, "QAD PN")
            append_numeric_anomaly(anomalies, "Packing list", sheet_idx, row_idx, row, qty_col, "Quantity", True)
            append_numeric_anomaly(anomalies, "Packing list", sheet_idx, row_idx, row, net_col, "N.W.(KG)", True)
            append_numeric_anomaly(anomalies, "Packing list", sheet_idx, row_idx, row, gross_col, "G.W.(KG)", True)
            item_net_weight = round2(to_number(row[net_col]))
            item_gross_weight = round2(to_number(row[gross_col]))
            for part_key in part_match_keys(part):
                part_weights[part_key] = round2(part_weights.get(part_key, 0) + item_net_weight)
                part_gross_weights[part_key] = round2(part_gross_weights.get(part_key, 0) + item_gross_weight)
            summed_net_weight += item_net_weight
            summed_gross_weight += item_gross_weight
            if pallet_col is not None and pallet_col < len(row):
                for number in re.findall(r"\d+", safe_text(row[pallet_col])):
                    pallets.add(int(number))

        if not package_count and pallets:
            package_count = len(pallets)
        if not package_count:
            for row in sheet:
                for value in row:
                    match = re.search(
                        r"\b(\d+)\s*(pallet|plt|carton|ctn|托盘|箱)s?\b",
                        safe_text(value),
                        re.I,
                    )
                    if match:
                        package_count = max(package_count, int(match.group(1)))

        net_weight = total_net_weight or round2(summed_net_weight)
        gross_weight = total_gross_weight or round2(summed_gross_weight)
        if total_net_weight and summed_net_weight and abs(total_net_weight - round2(summed_net_weight)) > 0.5:
            parser_warnings.append(
                f"Packing list 工作表“{source_sheet}”明细净重 {round2(summed_net_weight)} "
                f"与合计净重 {total_net_weight} 不一致"
            )
        if total_gross_weight and summed_gross_weight and abs(total_gross_weight - round2(summed_gross_weight)) > 0.5:
            parser_warnings.append(
                f"Packing list 工作表“{source_sheet}”明细毛重 {round2(summed_gross_weight)} "
                f"与合计毛重 {total_gross_weight} 不一致"
            )
    else:
        parser_warnings.append("Packing list 未找到同时包含料号、数量、净重和毛重的有效明细表")

    return {
        "packageCount": int(package_count) if package_count else "",
        "grossWeight": round2(gross_weight) if gross_weight else "",
        "netWeight": round2(net_weight) if net_weight else "",
        "partWeights": part_weights,
        "partGrossWeights": part_gross_weights,
        "exportDateSerial": export_date,
        "anomalies": anomalies,
        "warnings": parser_warnings,
        "sourceSheet": source_sheet,
    }


def load_rules(path=None):
    path = Path(path) if path else active_rules_path()
    if not path.exists():
        return {}
    rules = {
        "__standardization": {},
        "__part_conflicts": {},
        "__part_sources": {},
    }
    book = XlsxBook(path)
    try:
        if "值标准化表" in book.sheet_map:
            rows = book.rows("值标准化表")
            if rows:
                headers = rows[0]
                type_col = header_col(headers, "类型")
                raw_col = header_col(headers, "原始值")
                standard_col = header_col(headers, "标准值")
                enabled_col = header_col(headers, "是否启用")
                for row in rows[1:]:
                    if enabled_col is not None and enabled_col < len(row) and not is_enabled_value(row[enabled_col]):
                        continue
                    if type_col is None or raw_col is None or standard_col is None:
                        continue
                    if max(type_col, raw_col, standard_col) >= len(row):
                        continue
                    category = safe_text(row[type_col])
                    raw_value = safe_text(row[raw_col])
                    standard_value = safe_text(row[standard_col])
                    if category and raw_value and standard_value:
                        rules["__standardization"].setdefault(category, []).append((raw_value, standard_value))

        product_sheet_names = [name for name in book.sheet_map if "商品主数据" in name]
        if product_sheet_names:
            sheet_names = product_sheet_names
        else:
            sheet_names = []
            fallback_sheet_names = []
            for sheet_name in book.sheet_map:
                candidate_rows = book.rows(sheet_name)
                if not candidate_rows:
                    continue
                candidate_header_idx = find_header_index_by_terms(
                    candidate_rows,
                    [("qad pn", "qad p/n", "料号/part no", "part no"), ("hs code", "商品编号")],
                )
                if candidate_header_idx is None:
                    continue
                candidate_headers = candidate_rows[candidate_header_idx]
                qad_col = header_col(candidate_headers, "QAD PN", "QAD P/N", "料号/Part No")
                description_col = header_col(
                    candidate_headers,
                    "中文商品名称",
                    "商品名称及规格型号",
                    "中文品名",
                    "Description",
                )
                brand_candidate_col = header_col(candidate_headers, "品牌", "Brand")
                if qad_col is not None and description_col is not None and brand_candidate_col is not None:
                    sheet_names.append(sheet_name)
                else:
                    fallback_sheet_names.append(sheet_name)
            if not sheet_names:
                sheet_names = fallback_sheet_names

        for sheet_name in sheet_names:
            rows = book.rows(sheet_name)
            if not rows:
                continue
            header_idx = find_header_index_by_terms(
                rows,
                [("qad pn", "qad p/n", "料号/part no", "part no", "料号"), ("hs code", "商品编号")],
            )
            if header_idx is None:
                continue
            headers = rows[header_idx]
            enabled_col = header_col(headers, "是否启用")
            source_sheet_col = header_col(headers, "来源Sheet", "来源 Sheet")
            primary_part_col = header_col(
                headers,
                "QAD PN",
                "QAD P/N",
                "料号/Part No",
                "Part No.",
                "Part No",
                "料号",
            )
            old_part_col = header_col(headers, "备用料号/Old Part No", "Old Part No.", "Old Part No")
            part_cols = [
                col
                for col in (primary_part_col, old_part_col)
                if col is not None
            ]
            hs_cols = [
                col
                for col in (
                    header_col(headers, "更新hs code", "最新hs code"),
                    header_col(headers, "HS Code", "商品编号"),
                )
                if col is not None
            ]
            desc_col = header_col(
                headers,
                "中文商品名称",
                "商品名称及规格型号",
                "中文品名",
                "Description",
            )
            brand_col = header_col(headers, "品牌", "Brand")
            if not part_cols or not hs_cols:
                continue
            for row in rows[header_idx + 1:]:
                if enabled_col is not None and enabled_col < len(row) and not is_enabled_value(row[enabled_col]):
                    continue
                record = {}
                for hs_col in hs_cols:
                    if hs_col < len(row) and normalize_hs(row[hs_col]):
                        record["hsCode"] = normalize_hs(row[hs_col])
                        break
                if desc_col is not None and desc_col < len(row) and safe_text(row[desc_col]):
                    record["goodsName"] = safe_text(row[desc_col])
                if brand_col is not None and brand_col < len(row) and safe_text(row[brand_col]):
                    record["brand"] = (
                        standardize_config_value(rules, "品牌", row[brand_col])
                        or normalize_brand(row[brand_col])
                    )
                for part_col in part_cols:
                    if part_col >= len(row):
                        continue
                    raw_part = safe_text(row[part_col])
                    if not is_plausible_part_no(raw_part):
                        continue
                    keys = part_match_keys(raw_part)
                    if (
                        source_sheet_col is not None
                        and source_sheet_col < len(row)
                        and "补充" in safe_text(row[source_sheet_col])
                    ):
                        keys = keys[:1]
                    if not keys:
                        continue
                    for key in keys:
                        sources = rules["__part_sources"].setdefault(key, [])
                        if raw_part not in sources:
                            sources.append(raw_part)
                        if key in rules["__part_conflicts"]:
                            rules["__part_conflicts"][key]["parts"] = list(sources)
                            continue
                        existing = rules.get(key)
                        if existing is None:
                            rules[key] = dict(record)
                            continue
                        conflict_fields = [
                            field
                            for field in ("hsCode", "goodsName", "brand")
                            if existing.get(field)
                            and record.get(field)
                            and existing[field] != record[field]
                        ]
                        if conflict_fields:
                            rules.pop(key, None)
                            rules["__part_conflicts"][key] = {
                                "parts": list(sources),
                                "fields": conflict_fields,
                            }
                            continue
                        for field, value in record.items():
                            if value and not existing.get(field):
                                existing[field] = value
    finally:
        book.close()
    return rules


def infer_trade_country(invoice, rules):
    return (
        infer_country_from_text(invoice.get("billToContext", ""), rules, allow_china=False)
        or infer_country_from_text(invoice.get("countryContext", ""), rules, allow_china=False)
    )


def infer_destination_country(invoice, rules):
    return (
        infer_country_from_text(invoice.get("shipToContext", ""), rules, allow_china=False)
        or infer_country_from_text(invoice.get("consignee", ""), rules, allow_china=False)
        or infer_trade_country(invoice, rules)
    )


def merge_preview(invoice, packing, rules):
    warnings = []
    warnings.extend(packing.get("warnings", []))
    source_anomalies = []
    source_anomalies.extend(invoice.get("anomalies", []))
    source_anomalies.extend(packing.get("anomalies", []))
    if not invoice["items"]:
        warnings.append("Invoice 未解析到商品明细")
    if not packing.get("packageCount"):
        warnings.append("Packing list 未识别到件数，请在预览中复核")
    if not packing.get("grossWeight"):
        warnings.append("Packing list 未识别到毛重，请在预览中复核")
    if not packing.get("netWeight"):
        warnings.append("Packing list 未识别到净重，请在预览中复核")

    enriched = []
    missing_rules = []
    rule_conflicts = []
    for item in invoice["items"]:
        rule, _, conflict = part_rule_lookup(rules, item["partNo"])
        merged = dict(item)
        merged["hsCode"] = rule.get("hsCode", "")
        merged["goodsName"] = rule.get("goodsName", "")
        merged["brand"] = normalize_brand(rule.get("brand", ""))
        merged["netWeight"] = part_lookup(packing.get("partWeights", {}), item["partNo"])
        merged["grossWeight"] = part_lookup(packing.get("partGrossWeights", {}), item["partNo"])
        if conflict:
            rule_conflicts.append(
                f"{item['partNo']}（冲突料号：{'、'.join(conflict.get('parts', []))}）"
            )
        if not merged["hsCode"] or not merged["goodsName"]:
            missing_rules.append(item["partNo"])
        enriched.append(merged)

    if rule_conflicts:
        warnings.append(
            "以下料号按新标准化规则匹配到多条不同商品主数据，已停止自动取值："
            + "；".join(rule_conflicts[:20])
        )
    if missing_rules:
        warnings.append("以下 Part No. 未完整匹配到 HS/商品名称规则：" + "、".join(missing_rules[:20]))

    trade_country = infer_trade_country(invoice, rules)
    destination_country = infer_destination_country(invoice, rules)
    if not trade_country and destination_country:
        trade_country = destination_country
    if not destination_country:
        warnings.append("未能根据 Invoice Ship to 识别运抵国/最终目的国，相关字段将留空")
    if not trade_country:
        warnings.append("未能根据 Invoice Bill to 识别贸易国，相关字段将留空")

    groups = {}
    for item in enriched:
        key = (
            item.get("hsCode") or "UNKNOWN",
            item.get("goodsName") or item.get("descriptionEn") or "未匹配商品名称",
            item.get("brand") or "无",
            item.get("currency") or invoice["currency"],
            destination_country,
        )
        group = groups.setdefault(key, {
            "hsCode": key[0],
            "goodsName": key[1],
            "brand": key[2],
            "currency": key[3],
            "destinationCountry": key[4],
            "quantity": 0.0,
            "amount": 0.0,
            "netWeight": 0.0,
            "parts": [],
            "sourceItems": [],
        })
        group["quantity"] += to_number(item["quantity"])
        group["amount"] += to_number(item["amount"])
        group["netWeight"] += to_number(item.get("netWeight"))
        group["parts"].append(item["partNo"])
        group["sourceItems"].append(item)

    commodity_lines = list(groups.values())
    for idx, row in enumerate(commodity_lines, 1):
        row["itemNo"] = idx
        row["quantity"] = round2(row["quantity"])
        row["amount"] = round2(row["amount"])
        row["netWeight"] = round2(row["netWeight"])

    if not any(row["netWeight"] for row in commodity_lines) and packing.get("netWeight") and commodity_lines:
        total_qty = sum(row["quantity"] for row in commodity_lines) or 1
        for row in commodity_lines:
            row["netWeight"] = round2(float(packing["netWeight"]) * row["quantity"] / total_qty)

    audit_samples = []
    for row in commodity_lines:
        for item in row.get("sourceItems", []):
            audit_samples.append({
                "itemNo": row["itemNo"],
                "hsCode": row["hsCode"],
                "goodsName": row["goodsName"],
            "brand": row["brand"],
            "qadPartNo": item.get("partNo", ""),
            "invoiceQuantity": round2(to_number(item.get("quantity"))),
                "packingNetWeight": round2(to_number(item.get("netWeight"))),
                "packingGrossWeight": round2(to_number(item.get("grossWeight"))),
                "unitPrice": round2(to_number(item.get("unitPrice"))),
                "invoiceAmount": round2(to_number(item.get("amount"))),
                "currency": item.get("currency") or invoice["currency"],
                "poNo": item.get("poNo", ""),
                "invoiceSourceSheet": item.get("sourceSheet", ""),
                "invoiceSourceRow": item.get("sourceRow", ""),
            })
    for row in commodity_lines:
        row.pop("sourceItems", None)

    total_quantity = round2(sum(row["quantity"] for row in commodity_lines))
    total_amount = round2(sum(row["amount"] for row in commodity_lines))
    total_net_weight = round2(sum(row["netWeight"] for row in commodity_lines)) or packing.get("netWeight", "")

    if packing.get("netWeight") and total_net_weight:
        diff = abs(float(packing["netWeight"]) - float(total_net_weight))
        if diff > 0.5:
            warnings.append(f"商品分摊净重 {total_net_weight} 与 Packing list 净重 {packing['netWeight']} 不一致")

    return {
        "contractNo": invoice["contractNo"],
        "consignee": invoice["consignee"],
        "tradeTerm": invoice["tradeTerm"],
        "transportMode": invoice.get("transportMode", ""),
        "currency": invoice["currency"],
        "declarationDateSerial": invoice.get("exportDateSerial") or packing.get("exportDateSerial") or "",
        "exportDateSerial": "",
        "packageKind": "再生木托",
        "packageCount": packing.get("packageCount") or "",
        "grossWeight": packing.get("grossWeight") or "",
        "netWeight": packing.get("netWeight") or total_net_weight,
        "originCountry": "中国",
        "tradeCountry": trade_country,
        "destinationCountry": destination_country,
        "domesticSource": "苏州工业园区",
        "commodityLines": commodity_lines,
        "totals": {
            "quantity": total_quantity,
            "amount": total_amount,
            "netWeight": total_net_weight,
            "currency": invoice["currency"],
        },
        "warnings": warnings,
        "sourceAnomalies": source_anomalies,
        "auditSamples": audit_samples,
        "generatedAt": now_stamp(),
    }


def ensure_cell(sheet_data, ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    row_num = int(match.group(2))
    col_name = match.group(1)
    row = sheet_data.find(f".//{{{NS_MAIN}}}row[@r='{row_num}']")
    if row is None:
        sheet_data_parent = sheet_data.find(f"{{{NS_MAIN}}}sheetData")
        row = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_num)})
        inserted = False
        for idx, child in enumerate(list(sheet_data_parent)):
            existing_row = int(child.attrib.get("r", "0") or "0")
            if existing_row > row_num:
                sheet_data_parent.insert(idx, row)
                inserted = True
                break
        if not inserted:
            sheet_data_parent.append(row)
    cell = row.find(f"{{{NS_MAIN}}}c[@r='{ref}']")
    if cell is None:
        cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": ref})
        inserted = False
        target_col = column_to_number(col_name)
        for idx, child in enumerate(list(row)):
            cref = child.attrib.get("r", "")
            cmatch = re.match(r"([A-Z]+)", cref)
            if cmatch and column_to_number(cmatch.group(1)) > target_col:
                row.insert(idx, cell)
                inserted = True
                break
        if not inserted:
            row.append(cell)
    return cell


def set_cell(root, ref, value):
    cell = ensure_cell(root, ref)
    for child in list(cell):
        cell.remove(child)
    if isinstance(value, (int, float)) and value != "":
        cell.attrib.pop("t", None)
        v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
        v.text = str(int(value)) if float(value).is_integer() else str(value)
    else:
        cell.attrib["t"] = "inlineStr"
        is_node = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
        t = ET.SubElement(is_node, f"{{{NS_MAIN}}}t")
        t.text = safe_text(value)


def copy_cell_style(root, source_ref, target_ref):
    source = root.find(f".//{{{NS_MAIN}}}c[@r='{source_ref}']")
    if source is None or "s" not in source.attrib:
        return
    ensure_cell(root, target_ref).attrib["s"] = source.attrib["s"]


def apply_cell_style(root, ref, style_id):
    if style_id is None:
        return
    cell = ensure_cell(root, ref)
    cell.attrib["s"] = str(style_id)


def output_text_is_abnormal(value):
    reason = text_value_state(value)
    if reason:
        return True
    text = safe_text(value)
    return text.upper() == "UNKNOWN" or text.startswith("未匹配")


def output_number_is_abnormal(value, require_positive=True):
    _, reason = numeric_value_state(value, require_positive=require_positive)
    return bool(reason)


def flag_text_cell(root, ref, value, style_id):
    if output_text_is_abnormal(value):
        apply_cell_style(root, ref, style_id)


def flag_number_cell(root, ref, value, style_id, unit_ref=None):
    if output_number_is_abnormal(value, require_positive=True):
        apply_cell_style(root, ref, style_id)
        if unit_ref:
            apply_cell_style(root, unit_ref, style_id)


def clear_cell(root, ref):
    set_cell(root, ref, "")


def worksheet_cell_text(cell, shared_strings):
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        inline = cell.find(f"{{{NS_MAIN}}}is")
        return "".join(t.text or "" for t in inline.findall(f".//{{{NS_MAIN}}}t")) if inline is not None else ""
    value = cell.find(f"{{{NS_MAIN}}}v")
    if value is None:
        return ""
    text = value.text or ""
    if cell_type == "s":
        idx = int(float(text)) if text else -1
        return shared_strings[idx] if 0 <= idx < len(shared_strings) else ""
    return text


def zip_shared_strings(zip_file):
    if "xl/sharedStrings.xml" not in zip_file.namelist():
        return []
    root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t")) for si in root.findall(f"{{{NS_MAIN}}}si")]


def find_row_containing_text(sheet_root, shared_strings, target_text):
    target = safe_text(target_text).lower()
    for cell in sheet_root.findall(f".//{{{NS_MAIN}}}c"):
        text = worksheet_cell_text(cell, shared_strings).lower()
        if target and target in text:
            match = re.match(r"[A-Z]+(\d+)", cell.attrib.get("r", ""))
            if match:
                return int(match.group(1))
    return None


def declaration_sheet_layout(sheet_root, shared_strings):
    header_row = find_row_containing_text(sheet_root, shared_strings, "项号") or 17
    subtotal_row = find_row_containing_text(sheet_root, shared_strings, "Sub Total") or 75
    line_start = header_row + 1
    if subtotal_row <= line_start + 1:
        subtotal_row = 75
    line_rows = []
    row_num = line_start
    while row_num + 1 < subtotal_row:
        line_rows.append(row_num)
        row_num += 3
    return line_start, subtotal_row, line_rows


def split_marker_cell_ref(ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    if not match:
        return "", 0
    return match.group(1), int(match.group(2))


def find_template_markers(sheet_root, shared_strings):
    markers = {}
    marker_pattern = re.compile(r"\{\{\s*([^{}]+?)\s*\}\}")
    for cell in sheet_root.findall(f".//{{{NS_MAIN}}}c"):
        ref = cell.attrib.get("r", "")
        text = worksheet_cell_text(cell, shared_strings)
        for match in marker_pattern.finditer(text):
            markers.setdefault(match.group(1).strip(), []).append(ref)
    return markers


def first_marker_ref(markers, marker_name):
    refs = markers.get(marker_name) or []
    return refs[0] if refs else None


def set_marker_value(sheet_root, markers, marker_name, value):
    ref = first_marker_ref(markers, marker_name)
    if ref:
        set_cell(sheet_root, ref, value)
    return ref


def commodity_marker_blocks(markers):
    item_refs = sorted(markers.get("商品项号", []), key=lambda ref: split_marker_cell_ref(ref)[1])
    if not item_refs:
        return []
    marker_names = {
        "商品项号", "商品编号", "商品名称", "数量", "数量单位", "金额", "币制",
        "原产国", "最终目的国", "境内货源地", "征免", "品牌", "净重", "净重单位",
    }
    item_rows = [split_marker_cell_ref(ref)[1] for ref in item_refs]
    result = []
    for idx, item_ref in enumerate(item_refs):
        start_row = item_rows[idx]
        end_row = item_rows[idx + 1] if idx + 1 < len(item_rows) else 10**9
        block = {"商品项号": item_ref}
        for marker_name in marker_names - {"商品项号"}:
            for ref in markers.get(marker_name, []):
                row_num = split_marker_cell_ref(ref)[1]
                if start_row <= row_num < end_row:
                    block[marker_name] = ref
                    break
        result.append(block)
    return result


def shift_cell_ref_row(ref, row_delta):
    match = re.fullmatch(r"(\$?[A-Z]+)(\$?)(\d+)", ref or "")
    if not match:
        return ref
    return f"{match.group(1)}{match.group(2)}{int(match.group(3)) + row_delta}"


def shift_range_rows(ref, threshold_row, row_delta):
    shifted = []
    for cell in (ref or "").split(":"):
        match = re.fullmatch(r"(\$?[A-Z]+)(\$?)(\d+)", cell)
        if not match:
            shifted.append(cell)
            continue
        row = int(match.group(3))
        if row >= threshold_row:
            row += row_delta
        shifted.append(f"{match.group(1)}{match.group(2)}{row}")
    return ":".join(shifted)


def offset_row_element(row_element, row_delta):
    row_element.attrib["r"] = str(
        int(row_element.attrib.get("r", "0")) + row_delta
    )
    for cell in row_element.findall(f"{{{NS_MAIN}}}c"):
        cell.attrib["r"] = shift_cell_ref_row(
            cell.attrib.get("r", ""),
            row_delta,
        )


def expand_marker_commodity_blocks(
    sheet_root,
    shared_strings,
    required_count,
):
    markers = find_template_markers(sheet_root, shared_strings)
    blocks = commodity_marker_blocks(markers)
    if required_count <= len(blocks) or not blocks:
        return None

    item_rows = [
        split_marker_cell_ref(block["商品项号"])[1]
        for block in blocks
    ]
    last_item_row = item_rows[-1]
    subtotal_ref = first_marker_ref(markers, "合计标签")
    fallback_stride = (
        item_rows[-1] - item_rows[-2]
        if len(item_rows) > 1
        else 3
    )
    subtotal_row = (
        split_marker_cell_ref(subtotal_ref)[1]
        if subtotal_ref
        else last_item_row + fallback_stride
    )
    stride = (
        item_rows[-1] - item_rows[-2]
        if len(item_rows) > 1
        else subtotal_row - last_item_row
    )
    if stride <= 0 or subtotal_row <= last_item_row:
        raise ValueError(
            "报关单模板商品明细标记行结构不正确，无法自动扩展"
        )

    additional_blocks = required_count - len(blocks)
    inserted_rows = additional_blocks * stride
    source_start = last_item_row
    source_end = subtotal_row - 1
    sheet_data = sheet_root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        raise ValueError("报关单模板缺少 sheetData，无法自动扩展")

    source_rows = [
        copy.deepcopy(row)
        for row in sheet_data.findall(f"{{{NS_MAIN}}}row")
        if source_start
        <= int(row.attrib.get("r", "0"))
        <= source_end
    ]
    if not source_rows:
        raise ValueError("报关单模板没有可复制的商品明细行")

    for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
        if int(row.attrib.get("r", "0")) >= subtotal_row:
            offset_row_element(row, inserted_rows)

    merge_cells = sheet_root.find(f"{{{NS_MAIN}}}mergeCells")
    source_merge_refs = []
    if merge_cells is not None:
        for merge_cell in merge_cells.findall(
            f"{{{NS_MAIN}}}mergeCell"
        ):
            ref = merge_cell.attrib.get("ref", "")
            row_numbers = []
            for cell in ref.split(":"):
                match = re.search(r"(\d+)$", cell)
                if match:
                    row_numbers.append(int(match.group(1)))
            if (
                row_numbers
                and min(row_numbers) >= source_start
                and max(row_numbers) <= source_end
            ):
                source_merge_refs.append(ref)
            merge_cell.attrib["ref"] = shift_range_rows(
                ref,
                subtotal_row,
                inserted_rows,
            )

    for block_index in range(1, additional_blocks + 1):
        row_delta = block_index * stride
        for source_row in source_rows:
            cloned_row = copy.deepcopy(source_row)
            offset_row_element(cloned_row, row_delta)
            sheet_data.append(cloned_row)
        if merge_cells is not None:
            for ref in source_merge_refs:
                cloned_merge = ET.SubElement(
                    merge_cells,
                    f"{{{NS_MAIN}}}mergeCell",
                )
                cloned_merge.attrib["ref"] = shift_range_rows(
                    ref,
                    source_start,
                    row_delta,
                )

    rows = list(sheet_data.findall(f"{{{NS_MAIN}}}row"))
    for row in rows:
        sheet_data.remove(row)
    for row in sorted(
        rows,
        key=lambda item: int(item.attrib.get("r", "0")),
    ):
        sheet_data.append(row)
    if merge_cells is not None:
        merge_cells.attrib["count"] = str(
            len(merge_cells.findall(f"{{{NS_MAIN}}}mergeCell"))
        )
    return {
        "insertRow": subtotal_row,
        "insertedRows": inserted_rows,
    }


def shift_drawing_rows(drawing_xml, insert_row, inserted_rows):
    root = ET.fromstring(drawing_xml)
    namespace = (
        "http://schemas.openxmlformats.org/drawingml/"
        "2006/spreadsheetDrawing"
    )
    threshold = insert_row - 1
    changed = False
    for row_node in root.findall(f".//{{{namespace}}}row"):
        try:
            value = int(row_node.text or "0")
        except ValueError:
            continue
        if value >= threshold:
            row_node.text = str(value + inserted_rows)
            changed = True
    return serialize_excel_xml(root) if changed else None


def fill_marker_template(sheet_root, markers, preview, red_style_id=None):
    declaration_date = preview.get("declarationDateSerial") or preview.get("exportDateSerial") or ""
    fixed_values = {
        "境内发货人代码": "91320594762449680U",
        "境内发货人": "凯斯库汽车部件（苏州）有限公司3205240783",
        "出口日期": "",
        "申报日期": declaration_date,
        "境外收货人": preview["consignee"],
        "运输方式": preview.get("transportMode") or "",
        "生产销售单位代码": "91320594762449680U",
        "生产销售单位": "凯斯库汽车部件（苏州）有限公司3205240783",
        "监管方式": "一般",
        "征免性质": "101",
        "合同协议号": preview["contractNo"],
        "贸易国": preview.get("tradeCountry") or "",
        "运抵国": preview.get("destinationCountry") or "",
        "离境口岸": "上海",
        "包装种类": preview["packageKind"],
        "件数": preview["packageCount"],
        "毛重": preview["grossWeight"],
        "净重": preview["netWeight"],
        "成交方式": preview["tradeTerm"],
        "运费币制": preview["totals"]["currency"],
    }
    for marker_name, value in fixed_values.items():
        set_marker_value(sheet_root, markers, marker_name, value)
    declaration_ref = first_marker_ref(markers, "申报日期") or "L4"
    if not markers.get("申报日期"):
        export_ref = first_marker_ref(markers, "出口日期") or "G4"
        copy_cell_style(sheet_root, export_ref, declaration_ref)
        set_cell(sheet_root, declaration_ref, declaration_date)

    blocks = commodity_marker_blocks(markers)
    if len(preview["commodityLines"]) > len(blocks):
        raise ValueError(f"报关单模板商品明细标记不足：模板可填写 {len(blocks)} 行，当前需要 {len(preview['commodityLines'])} 行")

    field_values = {
        "商品项号": lambda line: line["itemNo"],
        "商品编号": lambda line: line["hsCode"],
        "商品名称": lambda line: line["goodsName"],
        "数量": lambda line: line["quantity"],
        "数量单位": lambda line: "个",
        "金额": lambda line: line["amount"],
        "币制": lambda line: line["currency"],
        "原产国": lambda line: preview["originCountry"],
        "最终目的国": lambda line: preview["destinationCountry"],
        "境内货源地": lambda line: preview["domesticSource"],
        "征免": lambda line: "照章征税",
        "品牌": lambda line: line["brand"],
        "净重": lambda line: line["netWeight"],
        "净重单位": lambda line: "千克",
    }
    for idx, block in enumerate(blocks):
        line = preview["commodityLines"][idx] if idx < len(preview["commodityLines"]) else None
        for marker_name, ref in block.items():
            set_cell(sheet_root, ref, field_values[marker_name](line) if line else "")
        if line and red_style_id is not None:
            if "商品编号" in block:
                flag_text_cell(sheet_root, block["商品编号"], line["hsCode"], red_style_id)
            if "商品名称" in block:
                flag_text_cell(sheet_root, block["商品名称"], line["goodsName"], red_style_id)
            if "数量" in block:
                flag_number_cell(sheet_root, block["数量"], line["quantity"], red_style_id, block.get("数量单位"))
            if "金额" in block:
                flag_number_cell(sheet_root, block["金额"], line["amount"], red_style_id)
            if "币制" in block:
                flag_text_cell(sheet_root, block["币制"], line["currency"], red_style_id)
            if "净重" in block:
                flag_number_cell(sheet_root, block["净重"], line["netWeight"], red_style_id, block.get("净重单位"))

    subtotal_values = {
        "合计标签": "Sub Total",
        "合计数量": preview["totals"]["quantity"],
        "合计数量单位": "个",
        "合计金额": preview["totals"]["amount"],
        "合计币制": preview["totals"]["currency"],
        "合计净重": preview["netWeight"],
        "合计净重单位": "千克",
    }
    for marker_name, value in subtotal_values.items():
        ref = set_marker_value(sheet_root, markers, marker_name, value)
        if red_style_id is not None and ref:
            if marker_name in ("合计数量", "合计金额", "合计净重"):
                flag_number_cell(sheet_root, ref, value, red_style_id)
            elif marker_name == "合计币制":
                flag_text_cell(sheet_root, ref, value, red_style_id)


def ensure_red_bold_style(styles_xml):
    root = ET.fromstring(styles_xml)
    fonts = root.find(f"{{{NS_MAIN}}}fonts")
    fills = root.find(f"{{{NS_MAIN}}}fills")
    borders = root.find(f"{{{NS_MAIN}}}borders")
    cell_xfs = root.find(f"{{{NS_MAIN}}}cellXfs")
    if fonts is None or fills is None or borders is None or cell_xfs is None:
        return None, styles_xml

    font_id = len(list(fonts))
    font = ET.SubElement(fonts, f"{{{NS_MAIN}}}font")
    ET.SubElement(font, f"{{{NS_MAIN}}}b")
    ET.SubElement(font, f"{{{NS_MAIN}}}color", {"rgb": "FFFF0000"})
    ET.SubElement(font, f"{{{NS_MAIN}}}sz", {"val": "11"})
    ET.SubElement(font, f"{{{NS_MAIN}}}name", {"val": "Calibri"})
    fonts.attrib["count"] = str(font_id + 1)

    fill_id = 0
    border_id = 0
    style_id = len(list(cell_xfs))
    ET.SubElement(cell_xfs, f"{{{NS_MAIN}}}xf", {
        "numFmtId": "0",
        "fontId": str(font_id),
        "fillId": str(fill_id),
        "borderId": str(border_id),
        "xfId": "0",
        "applyFont": "1",
    })
    cell_xfs.attrib["count"] = str(style_id + 1)
    return style_id, serialize_excel_xml(root)


def simple_sheet_xml(rows):
    worksheet = ET.Element(f"{{{NS_MAIN}}}worksheet")
    max_cols = max((len(row) for row in rows), default=1)
    max_rows = max(len(rows), 1)
    ET.SubElement(worksheet, f"{{{NS_MAIN}}}dimension", {
        "ref": f"A1:{excel_col_name(max_cols - 1)}{max_rows}",
    })
    sheet_views = ET.SubElement(worksheet, f"{{{NS_MAIN}}}sheetViews")
    sheet_view = ET.SubElement(sheet_views, f"{{{NS_MAIN}}}sheetView", {"workbookViewId": "0"})
    ET.SubElement(sheet_view, f"{{{NS_MAIN}}}pane", {
        "ySplit": "5",
        "topLeftCell": "A6",
        "activePane": "bottomLeft",
        "state": "frozen",
    })
    ET.SubElement(sheet_view, f"{{{NS_MAIN}}}selection", {
        "pane": "bottomLeft",
        "activeCell": "A6",
        "sqref": "A6",
    })
    ET.SubElement(worksheet, f"{{{NS_MAIN}}}sheetFormatPr", {
        "defaultRowHeight": "15",
    })
    cols = ET.SubElement(worksheet, f"{{{NS_MAIN}}}cols")
    widths = [12, 16, 28, 18, 18, 14, 14, 14, 10, 15, 15, 16, 18]
    for idx, width in enumerate(widths, 1):
        ET.SubElement(cols, f"{{{NS_MAIN}}}col", {
            "min": str(idx),
            "max": str(idx),
            "width": str(width),
            "customWidth": "1",
        })
    sheet_data = ET.SubElement(worksheet, f"{{{NS_MAIN}}}sheetData")
    for row_idx, row_values in enumerate(rows, 1):
        row_el = ET.SubElement(sheet_data, f"{{{NS_MAIN}}}row", {"r": str(row_idx)})
        for col_idx, value in enumerate(row_values):
            ref = f"{excel_col_name(col_idx)}{row_idx}"
            cell = ET.SubElement(row_el, f"{{{NS_MAIN}}}c", {"r": ref})
            if isinstance(value, (int, float)) and value != "":
                v = ET.SubElement(cell, f"{{{NS_MAIN}}}v")
                v.text = str(int(value)) if float(value).is_integer() else str(value)
            else:
                cell.attrib["t"] = "inlineStr"
                is_node = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
                t = ET.SubElement(is_node, f"{{{NS_MAIN}}}t")
                t.text = safe_text(value)
    return serialize_excel_xml(worksheet)


def audit_rows(preview):
    samples = preview.get("auditSamples") or []
    if not samples:
        return [["随机抽检"], ["无可抽检数据"]]
    item_nos = sorted({sample.get("itemNo") for sample in samples if sample.get("itemNo")})
    selected_item_no = random.choice(item_nos) if item_nos else samples[0].get("itemNo", "")
    selected = [sample for sample in samples if sample.get("itemNo") == selected_item_no] or [random.choice(samples)]
    first = selected[0]
    rows = [
        ["随机抽检"],
        ["抽检方式", "随机抽取一个报关商品编号，并按 QAD PN 展示输入文件中的逐行数据，供客户回查 Invoice 和 Packing list"],
        ["报关项号", first.get("itemNo", ""), "商品编号", first.get("hsCode", ""), "商品名称", first.get("goodsName", "")],
        [],
        [
            "报关项号",
            "商品编号",
            "商品名称",
            "QAD PN",
            "Invoice 数量",
            "Invoice 单价",
            "Invoice 金额",
            "币制",
            "Packing NW(kg)",
            "Packing GW(kg)",
            "PO No.",
            "Invoice 来源行",
        ],
    ]
    total_quantity = 0.0
    total_amount = 0.0
    total_net_weight = 0.0
    total_gross_weight = 0.0
    for sample in selected:
        quantity = round2(to_number(sample.get("invoiceQuantity")))
        amount = round2(to_number(sample.get("invoiceAmount")))
        net_weight = round2(to_number(sample.get("packingNetWeight")))
        gross_weight = round2(to_number(sample.get("packingGrossWeight")))
        total_quantity += quantity
        total_amount += amount
        total_net_weight += net_weight
        total_gross_weight += gross_weight
        rows.append([
            sample.get("itemNo", ""),
            sample.get("hsCode", ""),
            sample.get("goodsName", ""),
            sample.get("qadPartNo", ""),
            quantity,
            sample.get("unitPrice", ""),
            amount,
            sample.get("currency", ""),
            net_weight,
            gross_weight,
            sample.get("poNo", ""),
            f"{sample.get('invoiceSourceSheet', '')} 第{sample.get('invoiceSourceRow', '')}行",
        ])
    rows.append([])
    rows.append([
        "本抽检组加总",
        "",
        "",
        "",
        "",
        round2(total_quantity),
        "",
        round2(total_amount),
        first.get("currency", ""),
        round2(total_net_weight),
        round2(total_gross_weight),
        "",
        "",
    ])
    return rows


def next_sheet_number(sheet_paths):
    numbers = []
    for path in sheet_paths:
        match = re.search(r"sheet(\d+)\.xml$", path)
        if match:
            numbers.append(int(match.group(1)))
    return max(numbers or [0]) + 1


def add_audit_sheet(workbook, rels, content_types_root, sheets, modified, preview):
    sheet_name = "随机抽检"
    if sheet_name in sheets:
        modified[sheets[sheet_name]] = simple_sheet_xml(audit_rows(preview))
        return

    sheet_no = next_sheet_number(sheets.values())
    sheet_path = f"xl/worksheets/sheet{sheet_no}.xml"
    rel_target = f"worksheets/sheet{sheet_no}.xml"

    existing_rids = []
    for rel in rels.findall(f"{{{NS_PACKAGE_REL}}}Relationship"):
        match = re.match(r"rId(\d+)$", rel.attrib.get("Id", ""))
        if match:
            existing_rids.append(int(match.group(1)))
    rid = f"rId{max(existing_rids or [0]) + 1}"
    ET.SubElement(rels, f"{{{NS_PACKAGE_REL}}}Relationship", {
        "Id": rid,
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
        "Target": rel_target,
    })

    sheets_node = workbook.find(f"{{{NS_MAIN}}}sheets")
    existing_sheet_ids = [int(sheet.attrib.get("sheetId", "0")) for sheet in sheets_node.findall(f"{{{NS_MAIN}}}sheet")]
    ET.SubElement(sheets_node, f"{{{NS_MAIN}}}sheet", {
        "name": sheet_name,
        "sheetId": str(max(existing_sheet_ids or [0]) + 1),
        f"{{{NS_REL}}}id": rid,
    })

    override_exists = any(
        item.attrib.get("PartName") == f"/{sheet_path}"
        for item in content_types_root.findall(f"{{{NS_CONTENT_TYPES}}}Override")
    )
    if not override_exists:
        ET.SubElement(content_types_root, f"{{{NS_CONTENT_TYPES}}}Override", {
            "PartName": f"/{sheet_path}",
            "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
        })

    modified[sheet_path] = simple_sheet_xml(audit_rows(preview))


def remove_calc_chain(rels, content_types_root):
    for rel in list(rels):
        if rel.attrib.get("Type", "").endswith("/calcChain") or rel.attrib.get("Target") == "calcChain.xml":
            rels.remove(rel)
    for item in list(content_types_root):
        if item.attrib.get("PartName") == "/xl/calcChain.xml":
            content_types_root.remove(item)


def normalize_workbook_open_state(workbook):
    book_views = workbook.find(f"{{{NS_MAIN}}}bookViews")
    if book_views is None:
        book_views = ET.Element(f"{{{NS_MAIN}}}bookViews")
        first_child = next(iter(list(workbook)), None)
        if first_child is None:
            workbook.append(book_views)
        else:
            workbook.insert(0, book_views)
    workbook_view = book_views.find(f"{{{NS_MAIN}}}workbookView")
    if workbook_view is None:
        workbook_view = ET.SubElement(book_views, f"{{{NS_MAIN}}}workbookView")
    workbook_view.attrib.update({
        "visibility": "visible",
        "activeTab": "0",
        "firstSheet": "0",
        "xWindow": "0",
        "yWindow": "0",
        "windowWidth": "24000",
        "windowHeight": "14000",
    })
    workbook_view.attrib.pop("minimized", None)

    sheets_node = workbook.find(f"{{{NS_MAIN}}}sheets")
    if sheets_node is not None:
        for idx, sheet in enumerate(sheets_node.findall(f"{{{NS_MAIN}}}sheet")):
            if idx == 0:
                sheet.attrib.pop("state", None)


def unhide_rows(sheet_root):
    changed = False
    sheet_format = sheet_root.find(f"{{{NS_MAIN}}}sheetFormatPr")
    if sheet_format is not None and sheet_format.attrib.pop("zeroHeight", None) is not None:
        changed = True
    for row in sheet_root.findall(f".//{{{NS_MAIN}}}row"):
        if row.attrib.pop("hidden", None) is not None:
            changed = True
        if row.attrib.pop("zeroHeight", None) is not None:
            changed = True
    return changed


def generate_workbook(preview):
    template = active_template_path()
    if not template.exists():
        raise ValueError("未找到报关单模板，请先在管理员页面上传模板")
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_name = f"报关单 {preview['contractNo']}.xlsx"
    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}-{output_name}"

    with zipfile.ZipFile(template, "r") as zin:
        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        content_types = ET.fromstring(zin.read("[Content_Types].xml"))
        shared_strings = zip_shared_strings(zin)
        red_style_id = None
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = {}
        for sheet in workbook.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = rel_map[rid].lstrip("/")
            sheets[sheet.attrib["name"]] = "xl/" + target if not target.startswith("xl/") else target

        main_path = sheets.get("Sheet1") or next(iter(sheets.values()))
        main_root = ET.fromstring(zin.read(main_path))
        unhide_rows(main_root)
        expansion = expand_marker_commodity_blocks(
            main_root,
            shared_strings,
            len(preview["commodityLines"]),
        )
        template_markers = find_template_markers(main_root, shared_strings)

        if template_markers.get("合同协议号") or template_markers.get("商品项号"):
            fill_marker_template(main_root, template_markers, preview, red_style_id)
        else:
            declaration_date = preview.get("declarationDateSerial") or preview.get("exportDateSerial") or ""
            copy_cell_style(main_root, "G4", "L4")
            values = {
                "C3": "91320594762449680U",
                "A4": "凯斯库汽车部件（苏州）有限公司3205240783",
                "G4": "",
                "L4": declaration_date,
                "A6": preview["consignee"],
                "E6": preview.get("transportMode") or "",
                "C7": "91320594762449680U",
                "A8": "凯斯库汽车部件（苏州）有限公司3205240783",
                "E8": "一般",
                "G8": "101",
                "A10": preview["contractNo"],
                "E10": preview.get("tradeCountry") or "",
                "G10": preview.get("destinationCountry") or "",
                "O10": "上海",
                "A12": preview["packageKind"],
                "E12": preview["packageCount"],
                "F12": preview["grossWeight"],
                "G12": preview["netWeight"],
                "J12": preview["tradeTerm"],
            }
            for ref, value in values.items():
                set_cell(main_root, ref, value)
            flag_number_cell(main_root, "E12", preview["packageCount"], red_style_id)
            flag_number_cell(main_root, "F12", preview["grossWeight"], red_style_id)
            flag_number_cell(main_root, "G12", preview["netWeight"], red_style_id)

            line_start, subtotal_row, line_rows = declaration_sheet_layout(main_root, shared_strings)
            if len(preview["commodityLines"]) > len(line_rows):
                raise ValueError(f"报关单模板商品明细行不足：模板可填写 {len(line_rows)} 行，当前需要 {len(preview['commodityLines'])} 行")

            for row in range(line_start, subtotal_row + 2):
                for col in ("A", "B", "D", "G", "H", "J", "K", "L", "M", "O", "R"):
                    clear_cell(main_root, f"{col}{row}")

            for line, row_num in zip(preview["commodityLines"], line_rows):
                set_cell(main_root, f"A{row_num}", line["itemNo"])
                set_cell(main_root, f"B{row_num}", line["hsCode"])
                set_cell(main_root, f"D{row_num}", line["goodsName"])
                set_cell(main_root, f"G{row_num}", line["quantity"])
                set_cell(main_root, f"H{row_num}", "个")
                set_cell(main_root, f"J{row_num}", line["amount"])
                set_cell(main_root, f"K{row_num}", line["currency"])
                set_cell(main_root, f"L{row_num}", preview["originCountry"])
                set_cell(main_root, f"M{row_num}", preview["destinationCountry"])
                set_cell(main_root, f"O{row_num}", preview["domesticSource"])
                set_cell(main_root, f"R{row_num}", "照章征税")
                set_cell(main_root, f"B{row_num + 1}", line["brand"])
                set_cell(main_root, f"G{row_num + 1}", line["netWeight"])
                set_cell(main_root, f"H{row_num + 1}", "千克")
                flag_text_cell(main_root, f"B{row_num}", line["hsCode"], red_style_id)
                flag_text_cell(main_root, f"D{row_num}", line["goodsName"], red_style_id)
                flag_number_cell(main_root, f"G{row_num}", line["quantity"], red_style_id, f"H{row_num}")
                flag_number_cell(main_root, f"J{row_num}", line["amount"], red_style_id)
                flag_text_cell(main_root, f"K{row_num}", line["currency"], red_style_id)
                flag_number_cell(main_root, f"G{row_num + 1}", line["netWeight"], red_style_id, f"H{row_num + 1}")

            set_cell(main_root, f"A{subtotal_row}", "Sub Total")
            set_cell(main_root, f"G{subtotal_row}", preview["totals"]["quantity"])
            set_cell(main_root, f"H{subtotal_row}", "个")
            set_cell(main_root, f"J{subtotal_row}", preview["totals"]["amount"])
            set_cell(main_root, f"K{subtotal_row}", preview["totals"]["currency"])
            set_cell(main_root, f"G{subtotal_row + 1}", preview["netWeight"])
            set_cell(main_root, f"H{subtotal_row + 1}", "千克")
            flag_number_cell(main_root, f"G{subtotal_row}", preview["totals"]["quantity"], red_style_id, f"H{subtotal_row}")
            flag_number_cell(main_root, f"J{subtotal_row}", preview["totals"]["amount"], red_style_id)
            flag_text_cell(main_root, f"K{subtotal_row}", preview["totals"]["currency"], red_style_id)
            flag_number_cell(main_root, f"G{subtotal_row + 1}", preview["netWeight"], red_style_id, f"H{subtotal_row + 1}")

        update_worksheet_dimension(main_root)
        modified = {main_path: serialize_excel_xml(main_root)}
        if expansion:
            for filename in zin.namelist():
                if not (
                    filename.startswith("xl/drawings/drawing")
                    and filename.endswith(".xml")
                ):
                    continue
                shifted_drawing = shift_drawing_rows(
                    zin.read(filename),
                    expansion["insertRow"],
                    expansion["insertedRows"],
                )
                if shifted_drawing is not None:
                    modified[filename] = shifted_drawing
        if "申报要素" in sheets:
            decl_path = sheets["申报要素"]
            decl_root = ET.fromstring(zin.read(decl_path))
            unhide_rows(decl_root)
            for row in range(1, 180):
                for col in ("A", "B", "C", "D"):
                    clear_cell(decl_root, f"{col}{row}")
            cursor = 1
            for line in preview["commodityLines"]:
                set_cell(decl_root, f"A{cursor}", f"HS CODE：{line['hsCode']}")
                set_cell(decl_root, f"A{cursor + 2}", "1、品名：")
                set_cell(decl_root, f"B{cursor + 2}", line["goodsName"])
                set_cell(decl_root, f"A{cursor + 3}", "2、品牌：")
                set_cell(decl_root, f"B{cursor + 3}", line["brand"])
                set_cell(decl_root, f"A{cursor + 4}", "3、型号：")
                set_cell(decl_root, f"B{cursor + 4}", "无")
                set_cell(decl_root, f"A{cursor + 5}", "4、用途：")
                set_cell(decl_root, f"B{cursor + 5}", "汽车零部件用")
                cursor += 8
            update_worksheet_dimension(decl_root)
            modified[decl_path] = serialize_excel_xml(decl_root)

        for sheet_path in sheets.values():
            if sheet_path in modified:
                continue
            sheet_root = ET.fromstring(zin.read(sheet_path))
            if unhide_rows(sheet_root):
                modified[sheet_path] = serialize_excel_xml(sheet_root)

        remove_calc_chain(rels, content_types)
        normalize_workbook_open_state(workbook)
        modified["xl/workbook.xml"] = serialize_excel_xml(workbook)
        modified["xl/_rels/workbook.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
        modified["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            written = set()
            for info in zin.infolist():
                if info.filename == "xl/calcChain.xml":
                    continue
                data = modified.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zout.writestr(info, data)
                written.add(info.filename)
            for filename, data in modified.items():
                if filename not in written:
                    zout.writestr(filename, data)

    return output_path, output_name


def parse_upload(handler):
    length = int(handler.headers.get("Content-Length", "0"))
    if length > MAX_UPLOAD_BYTES:
        raise ValueError("上传文件过大")
    env = {
        "REQUEST_METHOD": "POST",
        "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
        "CONTENT_LENGTH": str(length),
    }
    form = cgi.FieldStorage(fp=handler.rfile, headers=handler.headers, environ=env)
    return form


def save_field_file(field, directory, fallback):
    if field is None or not getattr(field, "filename", ""):
        return None
    name = Path(field.filename).name or fallback
    suffix = Path(name).suffix.lower()
    if suffix not in (".xls", ".xlsx"):
        raise ValueError(f"{name} 不是支持的 Excel 文件")
    path = directory / f"{uuid.uuid4().hex}-{name}"
    with path.open("wb") as out:
        shutil.copyfileobj(field.file, out)
    return path


def form_file_fields(form, names):
    fields = []
    for name in names:
        if name not in form:
            continue
        value = form[name]
        if isinstance(value, list):
            fields.extend(value)
        else:
            fields.append(value)
    return [field for field in fields if getattr(field, "filename", "")]


def classify_uploaded_excel(path):
    name = Path(path).name.lower()
    if "invoice" in name:
        return "invoice"
    if "packing" in name or "packinglist" in name or "packing-list" in name:
        return "packing"

    try:
        sheets = [matrix_from_sheet(s) for s in read_spreadsheet(path)]
    except Exception:
        return ""

    for sheet in sheets:
        sample = " ".join(safe_text(value).lower() for row in sheet[:35] for value in row[:12])
        if "commercial" in sample and "invoice" in sample:
            return "invoice"
        if "packing" in sample and "list" in sample:
            return "packing"
    return ""


def detect_invoice_and_packing(paths):
    invoice_file = None
    packing_file = None
    classifications = []
    for path in paths:
        kind = classify_uploaded_excel(path)
        classifications.append({"filename": Path(path).name, "kind": kind or "unknown"})
        if kind == "invoice" and invoice_file is None:
            invoice_file = path
        elif kind == "packing" and packing_file is None:
            packing_file = path

    if not invoice_file or not packing_file:
        known = "；".join(f"{item['filename']} -> {item['kind']}" for item in classifications)
        raise ValueError("未能自动识别 Invoice 和 Packing list，请确认两个文件名或表格标题包含 Invoice / Packing。识别结果：" + known)
    return invoice_file, packing_file, classifications


SESSIONS = {}


class AppHandler(BaseHTTPRequestHandler):
    server_version = "SuriWorkDeclaration/1.0"

    def log_message(self, fmt, *args):
        print(f"[{now_stamp()}] {self.address_string()} {fmt % args}")

    def do_GET(self):
        try:
            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path
            if path == "/":
                self.serve_static("index.html")
                return
            if path == f"/u/{PUBLIC_TOKEN}":
                self.serve_static("index.html")
                return
            if path == "/suri-admin":
                self.redirect("/")
                return
            if path == f"/admin/{ADMIN_TOKEN}":
                self.redirect("/suri-admin")
                return
            if path == "/favicon.ico":
                self.send_response(204)
                self.end_headers()
                return
            if path.startswith("/static/"):
                self.serve_static(path[len("/static/"):])
                return
            if path == "/api/declaration/config":
                json_response(self, 200, {"ok": True, "active": declaration_config_status()})
                return
            if path == "/api/merge/config":
                merge_service.json_response(
                    self,
                    200,
                    {"ok": True, "active": merge_service.config_status()},
                )
                return
            if path == "/api/merge/history":
                merge_service.AppHandler.handle_history_list(
                    self,
                    urllib.parse.parse_qs(parsed.query),
                )
                return
            if path.startswith("/api/merge/history/"):
                parts = path.strip("/").split("/")
                if len(parts) == 4:
                    row = merge_service._history_row(parts[3])
                    merge_service.json_response(
                        self,
                        200,
                        {
                            "ok": True,
                            "record": merge_service.row_to_history(row, include_preview=True),
                        },
                    )
                    return
                if len(parts) == 5 and parts[4] == "download":
                    merge_service.AppHandler.handle_history_download(
                        self,
                        parts[3],
                        urllib.parse.parse_qs(parsed.query),
                    )
                    return
            if path.startswith("/api/merge/download/"):
                merge_service.AppHandler.handle_session_download(
                    self,
                    path.rsplit("/", 1)[-1],
                )
                return
            if path.startswith("/download/"):
                self.serve_download(path.rsplit("/", 1)[-1])
                return
            if path == "/api/history":
                self.handle_history_list(urllib.parse.parse_qs(parsed.query))
                return
            if path.startswith("/api/history/"):
                parts = path.strip("/").split("/")
                if len(parts) == 3:
                    self.handle_history_detail(parts[2])
                    return
                if len(parts) == 4 and parts[3] == "download":
                    self.handle_history_download(parts[2], urllib.parse.parse_qs(parsed.query))
                    return
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            log_runtime_error("GET failed:\n" + traceback.format_exc())
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def do_POST(self):
        try:
            if self.path == "/api/parse":
                self.handle_parse()
            elif self.path == "/api/generate":
                self.handle_generate()
            elif self.path == "/api/admin/rules":
                self.handle_admin_rules()
            elif self.path == "/api/merge/config":
                merge_service.AppHandler.handle_config_upload(self)
            elif self.path == "/api/merge/preview":
                merge_service.AppHandler.handle_preview(self)
            elif self.path == "/api/merge/generate":
                merge_service.AppHandler.handle_generate(self)
            else:
                self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            if self.path.startswith("/api/merge/"):
                merge_service.log_runtime_error("POST failed:\n" + traceback.format_exc())
            log_runtime_error("POST failed:\n" + traceback.format_exc())
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def do_DELETE(self):
        try:
            parsed = urllib.parse.urlparse(self.path)
            parts = parsed.path.strip("/").split("/")
            if len(parts) == 4 and parts[:3] == ["api", "merge", "history"]:
                merge_service.AppHandler.handle_history_delete(self, parts[3])
                return
            if len(parts) == 3 and parts[:2] == ["api", "history"]:
                self.handle_history_delete(parts[2])
                return
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            log_runtime_error("DELETE failed:\n" + traceback.format_exc())
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def redirect(self, target):
        self.send_response(302)
        self.send_header("Location", target)
        self.end_headers()

    def serve_static(self, name):
        clean = posixpath.normpath(urllib.parse.unquote(name)).lstrip("/")
        path = STATIC_DIR / clean
        if not path.exists() or not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        mime = "text/html; charset=utf-8"
        if path.suffix == ".css":
            mime = "text/css; charset=utf-8"
        elif path.suffix == ".js":
            mime = "application/javascript; charset=utf-8"
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_path(self, path, filename, content_type="application/octet-stream"):
        path = Path(path)
        if not path.is_file():
            raise ValueError("下载文件不存在")
        encoded = urllib.parse.quote(filename)
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded}")
        self.send_header("Content-Length", str(path.stat().st_size))
        self.end_headers()
        with path.open("rb") as source:
            shutil.copyfileobj(source, self.wfile, length=1024 * 1024)

    def serve_download(self, token):
        item = SESSIONS.get(token)
        if not item:
            self.send_error(HTTPStatus.NOT_FOUND, "Download expired")
            return
        path = Path(item["path"])
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "File not found")
            return
        try:
            validate_xlsx_file(path, require_dimensions=True)
        except ValueError as exc:
            self.send_error(HTTPStatus.INTERNAL_SERVER_ERROR, str(exc))
            return
        data = path.read_bytes()
        filename = item["filename"]
        encoded = urllib.parse.quote(filename)
        self.send_response(200)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def handle_parse(self):
        record_id = uuid.uuid4().hex
        upload_dir = HISTORY_DIR / record_id / "uploads"
        upload_dir.mkdir(parents=True, exist_ok=True)
        try:
            form = parse_upload(self)
            invoice_file = save_field_file(form["invoice"] if "invoice" in form else None, upload_dir, "invoice.xls")
            packing_file = save_field_file(form["packing"] if "packing" in form else None, upload_dir, "packing.xls")
            if not invoice_file or not packing_file:
                files = []
                for field in form_file_fields(form, ("documents", "files")):
                    saved = save_field_file(field, upload_dir, field.filename)
                    if saved:
                        files.append(saved)
                if len(files) < 2:
                    raise ValueError("请一次选择 Invoice 和 Packing list 两个文件")
                invoice_file, packing_file, classifications = detect_invoice_and_packing(files)
            else:
                classifications = [
                    {"filename": Path(invoice_file).name, "kind": "invoice"},
                    {"filename": Path(packing_file).name, "kind": "packing"},
                ]
            invoice = parse_invoice(invoice_file)
            packing = parse_packing(packing_file, invoice.get("contractNo"))
            preview = merge_preview(invoice, packing, load_rules())
            history_id = create_history_record(preview, invoice_file, packing_file, classifications, record_id=record_id)
            session_id = uuid.uuid4().hex
            SESSIONS[session_id] = {"preview": preview, "historyId": history_id, "created": time.time()}
            json_response(self, 200, {
                "ok": True,
                "sessionId": session_id,
                "historyId": history_id,
                "preview": preview,
                "recognizedFiles": classifications,
            })
        finally:
            shutil.rmtree(upload_dir, ignore_errors=True)
            record_dir = HISTORY_DIR / record_id
            if not (record_dir / "preview.json").exists():
                shutil.rmtree(record_dir, ignore_errors=True)

    def handle_generate(self):
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        data = json.loads(body.decode("utf-8") or "{}")
        session_id = data.get("sessionId")
        preview = data.get("preview") or (SESSIONS.get(session_id) or {}).get("preview")
        if not preview:
            raise ValueError("预览结果已过期，请重新上传解析")
        path, filename = generate_workbook(preview)
        history_id = data.get("historyId") or (SESSIONS.get(session_id) or {}).get("historyId")
        if history_id:
            path = update_history_output(history_id, path, filename, preview)
        token = uuid.uuid4().hex
        SESSIONS[token] = {"path": str(path), "filename": filename, "created": time.time()}
        json_response(self, 200, {
            "ok": True,
            "downloadUrl": f"/download/{token}",
            "filename": filename,
            "historyId": history_id,
            "savedPath": str(path),
        })

    def handle_admin_rules(self):
        ensure_app_dirs()
        upload_dir = Path(tempfile.mkdtemp(prefix="declaration-config-", dir=DATA_DIR))
        try:
            form = parse_upload(self)
            candidates = {}
            for kind, fallback in (("template", "template.xlsx"), ("rules", "rules.xlsx")):
                field = form[kind] if kind in form else None
                if field is None or not getattr(field, "filename", ""):
                    continue
                original_name = safe_filename(field.filename, fallback)
                if Path(original_name).suffix.lower() != ".xlsx":
                    raise ValueError(f"{original_name} 必须是 .xlsx 文件")
                path = upload_dir / f"{kind}.xlsx"
                with path.open("wb") as out:
                    shutil.copyfileobj(field.file, out)
                validate_xlsx_file(path)
                if kind == "rules":
                    rules = load_rules(path)
                    if not any(not key.startswith("__") for key in rules):
                        raise ValueError("规则表中没有识别到有效的料号匹配记录")
                candidates[kind] = {
                    "path": path,
                    "filename": original_name,
                    "updatedAt": now_stamp(),
                }
            if not candidates:
                raise ValueError("请选择要上传的模板或规则表")

            updated = {}
            for kind, item in candidates.items():
                target_dir = TEMPLATE_DIR if kind == "template" else RULES_DIR
                target_name = "template.xlsx" if kind == "template" else "rules.xlsx"
                version_dir = target_dir / "versions"
                version_dir.mkdir(parents=True, exist_ok=True)
                version_name = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}-{item['filename']}"
                shutil.copy2(item["path"], version_dir / safe_filename(version_name, target_name))
                temporary_target = target_dir / f".{target_name}.{uuid.uuid4().hex}"
                shutil.copy2(item["path"], temporary_target)
                os.replace(temporary_target, target_dir / target_name)
                updated[kind] = {
                    "filename": item["filename"],
                    "updatedAt": item["updatedAt"],
                }

            meta_path = storage_meta_path()
            meta = json_loads(meta_path.read_text("utf-8"), {}) if meta_path.exists() else {}
            meta.update(updated)
            meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
            json_response(
                self,
                200,
                {
                    "ok": True,
                    "updated": updated,
                    "active": declaration_config_status(),
                },
            )
        finally:
            shutil.rmtree(upload_dir, ignore_errors=True)

    def handle_history_list(self, query=None):
        init_db()
        query = query or {}
        start_at = normalize_history_datetime((query.get("start") or [""])[0])
        end_at = normalize_history_datetime((query.get("end") or [""])[0], end_of_day=True)
        limit_raw = (query.get("limit") or ["5"])[0]
        page_raw = (query.get("page") or ["1"])[0]
        try:
            limit = max(1, min(int(limit_raw), 50))
        except ValueError:
            limit = 5
        try:
            page = max(1, int(page_raw))
        except ValueError:
            page = 1
        where = []
        params = []
        if start_at:
            where.append("created_at >= ?")
            params.append(start_at)
        if end_at:
            where.append("created_at <= ?")
            params.append(end_at)
        if start_at and end_at and start_at > end_at:
            raise ValueError("开始时间不能晚于结束时间")
        where_sql = f"WHERE {' AND '.join(where)}" if where else ""
        with db_connect() as conn:
            total = conn.execute(
                f"SELECT COUNT(*) FROM history_records {where_sql}",
                params,
            ).fetchone()[0]
            total_pages = max(1, math.ceil(total / limit))
            page = min(page, total_pages)
            offset = (page - 1) * limit
            rows = conn.execute(
                f"""
                SELECT * FROM history_records
                 {where_sql}
                 ORDER BY datetime(created_at) DESC, created_at DESC
                 LIMIT ? OFFSET ?
                """,
                (*params, limit, offset),
            ).fetchall()
        json_response(self, 200, {
            "ok": True,
            "history": [row_to_history(row) for row in rows],
            "filters": {"start": start_at, "end": end_at, "limit": limit, "page": page},
            "pagination": {
                "page": page,
                "limit": limit,
                "total": total,
                "totalPages": total_pages,
            },
        })

    def handle_history_detail(self, record_id):
        init_db()
        with db_connect() as conn:
            row = conn.execute("SELECT * FROM history_records WHERE id = ?", (record_id,)).fetchone()
        if not row:
            raise ValueError("未找到历史记录")
        json_response(self, 200, {"ok": True, "record": row_to_history(row, include_preview=True)})

    def handle_history_download(self, record_id, query):
        kind = (query.get("kind") or [""])[0]
        if kind not in {"invoice", "packing", "output", "preview"}:
            raise ValueError("下载类型不正确")
        with db_connect() as conn:
            row = conn.execute("SELECT * FROM history_records WHERE id = ?", (record_id,)).fetchone()
        if not row:
            raise ValueError("未找到历史记录")
        record = row_to_history(row)
        path = Path(record["filePaths"].get(kind) or "")
        if not path.exists():
            raise ValueError("历史文件不存在")
        if kind == "output":
            validate_xlsx_file(path, require_dimensions=True)
        names = {
            "invoice": record["invoiceName"] or "invoice.xls",
            "packing": record["packingName"] or "packing.xls",
            "output": record["outputName"] or "output.xlsx",
            "preview": f"{record['contractNo'] or record_id}-preview.json",
        }
        body = path.read_bytes()
        encoded = urllib.parse.quote(names[kind])
        content_type = "application/octet-stream"
        if kind == "preview":
            content_type = "application/json; charset=utf-8"
        elif kind == "output":
            content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{encoded}")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def handle_history_delete(self, record_id):
        init_db()
        with db_connect() as conn:
            row = conn.execute("SELECT * FROM history_records WHERE id = ?", (record_id,)).fetchone()
            if not row:
                raise ValueError("未找到历史记录")
            conn.execute("DELETE FROM history_records WHERE id = ?", (record_id,))
        shutil.rmtree(HISTORY_DIR / record_id, ignore_errors=True)
        json_response(self, 200, {"ok": True})


def run(host="127.0.0.1", port=None, open_browser=False):
    ensure_app_dirs()
    init_db()
    merge_service.ensure_app_dirs()
    merge_service.init_db()
    port = int(port if port is not None else os.environ.get("PORT", "8000"))
    url = f"http://{host}:{port}/"
    print(f"报关单生成页面: {url}")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    ThreadingHTTPServer((host, port), AppHandler).serve_forever()


def run_in_thread(host="127.0.0.1", port=None):
    ensure_app_dirs()
    init_db()
    merge_service.ensure_app_dirs()
    merge_service.init_db()
    port = int(port if port is not None else os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, f"http://{host}:{port}/"


if __name__ == "__main__":
    run(open_browser=os.environ.get("SURI_OPEN_BROWSER") == "1")

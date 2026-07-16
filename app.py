#!/usr/bin/env python3
import cgi
import html
import io
import json
import os
import posixpath
import re
import shutil
import struct
import tempfile
import time
import unicodedata
import urllib.parse
import uuid
import zipfile
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from xml.etree import ElementTree as ET


BASE_DIR = Path(__file__).resolve().parent
STATIC_DIR = BASE_DIR / "static"
STORAGE_DIR = BASE_DIR / "storage"
OUTPUT_DIR = BASE_DIR / "outputs"

DEFAULT_TEMPLATE = BASE_DIR / "报关单 IMOS 空白模板.xlsx"
DEFAULT_RULES = BASE_DIR / "最新逻辑_请优先看" / "报关单配置关系表.xlsx"
LEGACY_RULES = BASE_DIR / "2026+Daily+Export+List.xlsx"

PUBLIC_TOKEN = os.environ.get("PUBLIC_TOKEN", "imos-demo")
ADMIN_TOKEN = os.environ.get("ADMIN_TOKEN", "admin-demo")
MAX_UPLOAD_BYTES = int(os.environ.get("MAX_UPLOAD_BYTES", str(25 * 1024 * 1024)))

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS_MAIN)


def now_stamp():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def active_template_path():
    uploaded = STORAGE_DIR / "template.xlsx"
    return uploaded if uploaded.exists() else DEFAULT_TEMPLATE


def active_rules_path():
    uploaded = STORAGE_DIR / "rules.xlsx"
    if uploaded.exists():
        return uploaded
    if DEFAULT_RULES.exists():
        return DEFAULT_RULES
    return LEGACY_RULES


def storage_meta_path():
    return STORAGE_DIR / "meta.json"


def json_response(handler, status, payload):
    body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def validate_xlsx_file(path):
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
            ET.fromstring(workbook.read("xl/workbook.xml"))
            ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
            ET.fromstring(workbook.read("[Content_Types].xml"))
    except zipfile.BadZipFile as exc:
        raise ValueError("生成文件不是有效的 Excel 工作簿") from exc
    except ET.ParseError as exc:
        raise ValueError("生成文件内部结构异常") from exc


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
    match = re.search(r"(20\d{6})", text or "")
    if not match:
        return ""
    ymd = match.group(1)
    year, month, day = int(ymd[:4]), int(ymd[4:6]), int(ymd[6:8])
    import datetime as _dt
    base = _dt.date(1899, 12, 30)
    return (_dt.date(year, month, day) - base).days


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


def row_contains_any(row, *terms):
    text = " ".join(safe_text(v).lower() for v in row)
    return any(term.lower() in text for term in terms)


def is_enabled_value(value):
    text = safe_text(value).strip().lower()
    return text not in ("否", "no", "n", "false", "0", "禁用", "停用")


def normalized_header(row):
    return [safe_text(v).lower().replace(" ", "").replace("_", "") for v in row]


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


def parse_invoice(path):
    sheets = [matrix_from_sheet(s) for s in read_spreadsheet(path)]
    invoice_sheet = None
    header_index = None
    for sheet in sheets:
        for idx, row in enumerate(sheet):
            compact = re.sub(r"[\s_]+", "", " ".join(safe_text(v).lower() for v in row))
            if any(term in compact for term in ("qadpn", "partno")) and any(
                term in compact for term in ("qty", "quantity")
            ):
                invoice_sheet = sheet
                header_index = idx
                break
        if invoice_sheet:
            break
    if not invoice_sheet:
        raise ValueError("Invoice 中没有找到包含 QAD PN 和 Quantity/Qty 的明细表")

    header = [re.sub(r"[\s_]+", "", safe_text(v).lower()) for v in invoice_sheet[header_index]]

    def find_col(*names):
        for name in names:
            normalized_name = re.sub(r"[\s_]+", "", name.lower())
            for idx, value in enumerate(header):
                if normalized_name == value:
                    return idx
            for idx, value in enumerate(header):
                if normalized_name in value:
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
        "goods_name": find_col("description"),
        "hs_code": find_col("hs code"),
        "brand": find_col("品牌"),
        "model": find_col("车型"),
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
    for row in invoice_sheet[header_index + 1:]:
        part = safe_text(row[cols["part"]]) if cols["part"] is not None and cols["part"] < len(row) else ""
        if not part or not re.search(r"\d", part):
            continue
        quantity = to_number(row[cols["quantity"]]) if cols["quantity"] is not None and cols["quantity"] < len(row) else 0
        unit_price = to_number(row[cols["unit_price"]]) if cols["unit_price"] is not None and cols["unit_price"] < len(row) else 0
        amount = to_number(row[cols["amount"]]) if cols["amount"] is not None and cols["amount"] < len(row) else 0
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
    }


def parse_packing(path, expected_contract_no=""):
    entries = read_spreadsheet_entries(path)
    package_count = 0
    gross_weight = 0.0
    net_weight = 0.0
    part_weights = {}
    part_gross_weights = {}
    export_date = excel_serial_from_yyyymmdd(Path(path).name)
    parser_warnings = []
    expected_key = re.sub(r"\s+", "", safe_text(expected_contract_no)).upper()
    candidates = []

    for sheet_idx, entry in enumerate(entries):
        sheet = matrix_from_sheet(entry["cells"])
        header_idx = None
        for idx, row in enumerate(sheet):
            compact = re.sub(r"[\s._()]+", "", " ".join(safe_text(v).lower() for v in row))
            has_part = any(term in compact for term in ("qadpn", "partno", "fgpn"))
            has_qty = any(term in compact for term in ("qty", "quantity"))
            has_net = any(term in compact for term in ("nwkg", "nw"))
            has_gross = any(term in compact for term in ("gwkg", "gw"))
            if has_part and has_qty and has_net and has_gross:
                header_idx = idx
                break
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
        sheet = selected["sheet"]
        header_idx = selected["header"]
        header = [
            re.sub(r"[\s._()]+", "", safe_text(v).lower())
            for v in sheet[header_idx]
        ]

        def find_col(*terms):
            for term in terms:
                normalized_term = re.sub(r"[\s._()]+", "", term.lower())
                for col, value in enumerate(header):
                    if normalized_term in value:
                        return col
            return None

        part_col = find_col("qad pn", "part no", "fg pn", "fg_pn", "fgpn")
        qty_col = find_col("qty", "quantity")
        net_col = find_col("n.w", "n w", "nw")
        gross_col = find_col("g.w", "g w", "gw")
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
        for row in sheet[header_idx + 1:]:
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
        })
        group["quantity"] += to_number(item["quantity"])
        group["amount"] += to_number(item["amount"])
        group["netWeight"] += to_number(item.get("netWeight"))
        group["parts"].append(item["partNo"])

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
        "generatedAt": now_stamp(),
    }


def ensure_cell(sheet_data, ref):
    match = re.match(r"([A-Z]+)(\d+)", ref)
    row_num = int(match.group(2))
    col_name = match.group(1)
    row = sheet_data.find(f".//{{{NS_MAIN}}}row[@r='{row_num}']")
    if row is None:
        sheet_data_parent = sheet_data.find(f"{{{NS_MAIN}}}sheetData")
        row = ET.SubElement(sheet_data_parent, f"{{{NS_MAIN}}}row", {"r": str(row_num)})
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
        book_views = ET.SubElement(workbook, f"{{{NS_MAIN}}}bookViews")
    workbook_view = book_views.find(f"{{{NS_MAIN}}}workbookView")
    if workbook_view is None:
        workbook_view = ET.SubElement(book_views, f"{{{NS_MAIN}}}workbookView")
    workbook_view.attrib.update(
        {
            "visibility": "visible",
            "activeTab": "0",
            "firstSheet": "0",
            "xWindow": "0",
            "yWindow": "0",
            "windowWidth": "18000",
            "windowHeight": "12000",
        }
    )


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
    OUTPUT_DIR.mkdir(exist_ok=True)
    output_name = f"报关单 {preview['contractNo']}.xlsx"
    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}-{output_name}"

    with zipfile.ZipFile(template, "r") as zin:
        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        content_types = ET.fromstring(zin.read("[Content_Types].xml"))
        shared_strings = zip_shared_strings(zin)
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = {}
        for sheet in workbook.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = rel_map[rid].lstrip("/")
            sheets[sheet.attrib["name"]] = "xl/" + target if not target.startswith("xl/") else target

        main_path = sheets.get("Sheet1") or next(iter(sheets.values()))
        main_root = ET.fromstring(zin.read(main_path))
        unhide_rows(main_root)
        template_markers = find_template_markers(main_root, shared_strings)

        if template_markers.get("合同协议号") or template_markers.get("商品项号"):
            fill_marker_template(main_root, template_markers, preview)
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

            set_cell(main_root, f"A{subtotal_row}", "Sub Total")
            set_cell(main_root, f"G{subtotal_row}", preview["totals"]["quantity"])
            set_cell(main_root, f"H{subtotal_row}", "个")
            set_cell(main_root, f"J{subtotal_row}", preview["totals"]["amount"])
            set_cell(main_root, f"K{subtotal_row}", preview["totals"]["currency"])
            set_cell(main_root, f"G{subtotal_row + 1}", preview["netWeight"])
            set_cell(main_root, f"H{subtotal_row + 1}", "千克")

        modified = {main_path: ET.tostring(main_root, encoding="utf-8", xml_declaration=True)}

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
            modified[decl_path] = ET.tostring(decl_root, encoding="utf-8", xml_declaration=True)

        for sheet_path in sheets.values():
            if sheet_path in modified:
                continue
            sheet_root = ET.fromstring(zin.read(sheet_path))
            if unhide_rows(sheet_root):
                modified[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

        remove_calc_chain(rels, content_types)
        normalize_workbook_open_state(workbook)
        modified["xl/workbook.xml"] = ET.tostring(workbook, encoding="utf-8", xml_declaration=True)
        modified["xl/_rels/workbook.xml.rels"] = ET.tostring(rels, encoding="utf-8", xml_declaration=True)
        modified["[Content_Types].xml"] = ET.tostring(content_types, encoding="utf-8", xml_declaration=True)

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename == "xl/calcChain.xml":
                    continue
                data = modified.get(info.filename)
                if data is None:
                    data = zin.read(info.filename)
                zout.writestr(info, data)

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
    server_version = "IMOSDeclaration/1.0"

    def log_message(self, fmt, *args):
        print(f"[{now_stamp()}] {self.address_string()} {fmt % args}")

    def do_GET(self):
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
        if path.startswith("/download/"):
            self.serve_download(path.rsplit("/", 1)[-1])
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self):
        try:
            if self.path == "/api/parse":
                self.handle_parse()
            elif self.path == "/api/generate":
                self.handle_generate()
            elif self.path == "/api/admin/rules":
                self.handle_admin_rules()
            else:
                self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
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
            validate_xlsx_file(path)
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
        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            form = parse_upload(self)
            invoice_file = save_field_file(form["invoice"] if "invoice" in form else None, tmpdir, "invoice.xls")
            packing_file = save_field_file(form["packing"] if "packing" in form else None, tmpdir, "packing.xls")
            if not invoice_file or not packing_file:
                files = []
                for field in form_file_fields(form, ("documents", "files")):
                    saved = save_field_file(field, tmpdir, field.filename)
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
            session_id = uuid.uuid4().hex
            SESSIONS[session_id] = {"preview": preview, "created": time.time()}
            json_response(self, 200, {
                "ok": True,
                "sessionId": session_id,
                "preview": preview,
                "recognizedFiles": classifications,
            })

    def handle_generate(self):
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        data = json.loads(body.decode("utf-8") or "{}")
        session_id = data.get("sessionId")
        preview = data.get("preview") or (SESSIONS.get(session_id) or {}).get("preview")
        if not preview:
            raise ValueError("预览结果已过期，请重新上传解析")
        path, filename = generate_workbook(preview)
        token = uuid.uuid4().hex
        SESSIONS[token] = {"path": str(path), "filename": filename, "created": time.time()}
        json_response(self, 200, {"ok": True, "downloadUrl": f"/download/{token}", "filename": filename})

    def handle_admin_rules(self):
        STORAGE_DIR.mkdir(exist_ok=True)
        form = parse_upload(self)
        updated = {}
        if "template" in form and getattr(form["template"], "filename", ""):
            path = STORAGE_DIR / "template.xlsx"
            with path.open("wb") as out:
                shutil.copyfileobj(form["template"].file, out)
            updated["template"] = {"filename": form["template"].filename, "updatedAt": now_stamp()}
        if "rules" in form and getattr(form["rules"], "filename", ""):
            path = STORAGE_DIR / "rules.xlsx"
            with path.open("wb") as out:
                shutil.copyfileobj(form["rules"].file, out)
            updated["rules"] = {"filename": form["rules"].filename, "updatedAt": now_stamp()}
        if not updated:
            raise ValueError("请选择要上传的模板或规则表")
        meta_path = storage_meta_path()
        meta = {}
        if meta_path.exists():
            meta = json.loads(meta_path.read_text("utf-8"))
        meta.update(updated)
        meta_path.write_text(json.dumps(meta, ensure_ascii=False, indent=2), "utf-8")
        json_response(self, 200, {"ok": True, "updated": updated, "active": meta})


def run():
    port = int(os.environ.get("PORT", "8000"))
    STORAGE_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    print(f"报关单生成页面: http://localhost:{port}/")
    ThreadingHTTPServer(("0.0.0.0", port), AppHandler).serve_forever()


if __name__ == "__main__":
    run()

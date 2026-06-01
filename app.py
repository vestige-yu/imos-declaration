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
DEFAULT_RULES = BASE_DIR / "2026+Daily+Export+List.xlsx"

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
    return uploaded if uploaded.exists() else DEFAULT_RULES


def json_response(handler, status, payload):
    body = json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


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


def normalize_part(value):
    return safe_text(value).upper().replace(" ", "")


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


class XlsWorkbook:
    """Small BIFF8 reader for the old .xls files used by this workflow."""

    END_OF_CHAIN = 0xFFFFFFFE
    FREE_SECTOR = 0xFFFFFFFF

    def __init__(self, data):
        self.data = data
        self.workbook_stream = self._read_workbook_stream()
        self.shared_strings = self._read_sst()
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
            result[sheet.attrib["name"]] = "xl/" + target if not target.startswith("xl/") else target
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


def read_spreadsheet(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".xls":
        workbook = XlsWorkbook.load(path)
        return workbook.sheets
    if suffix == ".xlsx":
        book = XlsxBook(path)
        try:
            return [
                {
                    (r, c): value
                    for r, row in enumerate(book.rows(name))
                    for c, value in enumerate(row)
                    if safe_text(value)
                }
                for name in book.sheet_map
            ]
        finally:
            book.close()
    raise ValueError("仅支持 .xls 或 .xlsx 文件")


def matrix_from_sheet(sheet):
    if not sheet:
        return []
    max_row = max(row for row, _ in sheet)
    max_col = max(col for _, col in sheet)
    return [[sheet.get((r, c), "") for c in range(max_col + 1)] for r in range(max_row + 1)]


def row_contains(row, *terms):
    text = " ".join(safe_text(v).lower() for v in row)
    return all(term.lower() in text for term in terms)


def parse_invoice(path):
    sheets = [matrix_from_sheet(s) for s in read_spreadsheet(path)]
    invoice_sheet = None
    header_index = None
    for sheet in sheets:
        for idx, row in enumerate(sheet):
            if row_contains(row, "qad pn") and row_contains(row, "qty"):
                invoice_sheet = sheet
                header_index = idx
                break
        if invoice_sheet:
            break
    if not invoice_sheet:
        raise ValueError("Invoice 中没有找到包含 QAD PN 和 qty. 的明细表")

    header = [safe_text(v).lower() for v in invoice_sheet[header_index]]

    def find_col(*names):
        for name in names:
            for idx, value in enumerate(header):
                if name.lower() == value:
                    return idx
            for idx, value in enumerate(header):
                if name.lower() in value:
                    return idx
        return None

    cols = {
        "part": find_col("qad pn", "part no"),
        "imos": find_col("imos p/n"),
        "description_en": find_col("description"),
        "quantity": find_col("qty"),
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
    currency = "USD"
    export_date = excel_serial_from_yyyymmdd(Path(path).name)
    for row in invoice_sheet[:30]:
        for idx, value in enumerate(row):
            text = safe_text(value)
            if re.fullmatch(r"SP\d{8,}", text):
                contract = text
            if "CASCO Imos" in text:
                consignee = text.strip()
            if "Delivery Term:" in text:
                trade_term = text.split(":", 1)[-1].strip().split()[0]
            if text in ("USD", "EUR", "CNY"):
                currency = text

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
            "imosPartNo": safe_text(row[cols["imos"]]) if cols["imos"] is not None and cols["imos"] < len(row) else part,
            "descriptionEn": safe_text(row[cols["description_en"]]) if cols["description_en"] is not None and cols["description_en"] < len(row) else "",
            "quantity": quantity,
            "unitPrice": unit_price,
            "amount": amount,
            "currency": item_currency or currency,
            "poNo": safe_text(row[cols["po"]]) if cols["po"] is not None and cols["po"] < len(row) else "",
            "goodsName": safe_text(row[cols["goods_name"]]) if cols["goods_name"] is not None and cols["goods_name"] < len(row) else "",
            "hsCode": normalize_hs(row[cols["hs_code"]]) if cols["hs_code"] is not None and cols["hs_code"] < len(row) else "",
            "brand": normalize_brand(row[cols["brand"]]) if cols["brand"] is not None and cols["brand"] < len(row) else "",
            "model": safe_text(row[cols["model"]]) if cols["model"] is not None and cols["model"] < len(row) else "",
        })

    return {
        "contractNo": contract or Path(path).stem.split()[0],
        "consignee": consignee or "CASCO Imos Italia S.R.L.",
        "tradeTerm": trade_term or "CPT",
        "currency": currency or "USD",
        "exportDateSerial": export_date,
        "items": items,
    }


def parse_packing(path):
    sheets = [matrix_from_sheet(s) for s in read_spreadsheet(path)]
    package_count = 0
    gross_weight = 0.0
    net_weight = 0.0
    part_weights = {}
    export_date = excel_serial_from_yyyymmdd(Path(path).name)

    for sheet in sheets:
        header_idx = None
        for idx, row in enumerate(sheet):
            if row_contains(row, "qad pn") and row_contains(row, "n.w"):
                header_idx = idx
                break
        if header_idx is None:
            continue
        header = [safe_text(v).lower() for v in sheet[header_idx]]

        def find_col(*terms):
            for term in terms:
                for col, value in enumerate(header):
                    if term in value:
                        return col
            return None

        part_col = find_col("qad pn")
        qty_col = find_col("qty")
        net_col = find_col("n.w")
        gross_col = find_col("g.w")
        pallet_col = find_col("pallet")
        if part_col is None or qty_col is None or net_col is None or gross_col is None:
            continue

        pallets = set()
        for row in sheet[header_idx + 1:]:
            label = " ".join(safe_text(v).lower() for v in row)
            if "total" in label:
                package_count = int(to_number(row[pallet_col])) if pallet_col is not None and pallet_col < len(row) else package_count
                net_weight = round2(to_number(row[net_col])) if net_col < len(row) else net_weight
                gross_weight = round2(to_number(row[gross_col])) if gross_col < len(row) else gross_weight
                continue

            if max(part_col, qty_col, net_col) >= len(row):
                continue
            part = safe_text(row[part_col])
            quantity = to_number(row[qty_col])
            if quantity < 1 or not re.match(r"^\d{9,}-\w+", part):
                continue
            part_weights[normalize_part(part)] = round2(to_number(row[net_col]))
            if pallet_col is not None and pallet_col < len(row):
                for number in re.findall(r"\d+", safe_text(row[pallet_col])):
                    pallets.add(int(number))

        if not package_count and pallets:
            package_count = max(pallets)
        if part_weights:
            break

    if not part_weights:
        for sheet in sheets:
            for row in sheet:
                line = " ".join(safe_text(v) for v in row)
                if re.search(r"\btotal\b", line, re.I):
                    numbers = [to_number(v, None) for v in row]
                    numbers = [v for v in numbers if v is not None and v > 0]
                    if len(numbers) >= 2:
                        gross_weight = max(gross_weight, max(numbers))
                        positives = sorted(numbers)
                        net_weight = max(net_weight, positives[-2] if len(positives) > 1 else positives[-1])
                for value in row:
                    text = safe_text(value)
                    if re.search(r"\b\d+\s*(pallet|plt|托盘)", text, re.I):
                        package_count = max(package_count, int(re.search(r"\d+", text).group(0)))

    return {
        "packageCount": int(package_count) if package_count else "",
        "grossWeight": round2(gross_weight) if gross_weight else "",
        "netWeight": round2(net_weight) if net_weight else "",
        "partWeights": part_weights,
        "exportDateSerial": export_date,
    }


def load_rules():
    path = active_rules_path()
    if not path.exists():
        return {}
    rules = {}
    book = XlsxBook(path)
    try:
        for sheet_name in book.sheet_map:
            rows = book.rows(sheet_name)
            if not rows:
                continue
            header = [safe_text(v).lower() for v in rows[0]]
            part_cols = [i for i, v in enumerate(header) if "part" in v or "pn" in v]
            hs_col = next((i for i, v in enumerate(header) if "hs" in v and "code" in v), None)
            desc_col = next((i for i, v in enumerate(header) if "description" in v or "货物名称" in v), None)
            brand_col = next((i for i, v in enumerate(header) if "品牌" in v), None)
            if not part_cols or hs_col is None:
                continue
            for row in rows[1:]:
                for part_col in part_cols:
                    if part_col >= len(row):
                        continue
                    key = normalize_part(row[part_col])
                    if not key:
                        continue
                    rules.setdefault(key, {})
                    if hs_col < len(row) and normalize_hs(row[hs_col]):
                        rules[key]["hsCode"] = normalize_hs(row[hs_col])
                    if desc_col is not None and desc_col < len(row) and safe_text(row[desc_col]):
                        rules[key]["goodsName"] = safe_text(row[desc_col])
                    if brand_col is not None and brand_col < len(row) and safe_text(row[brand_col]):
                        rules[key]["brand"] = normalize_brand(row[brand_col])
    finally:
        book.close()
    return rules


def merge_preview(invoice, packing, rules):
    warnings = []
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
    for item in invoice["items"]:
        rule = rules.get(normalize_part(item["partNo"])) or rules.get(normalize_part(item["imosPartNo"])) or {}
        merged = dict(item)
        merged["hsCode"] = item.get("hsCode") or rule.get("hsCode", "")
        merged["goodsName"] = item.get("goodsName") or rule.get("goodsName", "")
        merged["brand"] = normalize_brand(item.get("brand") or rule.get("brand", ""))
        merged["netWeight"] = packing.get("partWeights", {}).get(normalize_part(item["partNo"]), "")
        if not merged["hsCode"] or not merged["goodsName"]:
            missing_rules.append(item["partNo"])
        enriched.append(merged)

    if missing_rules:
        warnings.append("以下 Part No. 未完整匹配到 HS/商品名称规则：" + "、".join(missing_rules[:20]))

    groups = {}
    for item in enriched:
        key = (
            item.get("hsCode") or "UNKNOWN",
            item.get("goodsName") or item.get("descriptionEn") or "未匹配商品名称",
            item.get("brand") or "无",
            item.get("currency") or invoice["currency"],
        )
        group = groups.setdefault(key, {
            "hsCode": key[0],
            "goodsName": key[1],
            "brand": key[2],
            "currency": key[3],
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
        "currency": invoice["currency"],
        "exportDateSerial": invoice.get("exportDateSerial") or packing.get("exportDateSerial") or "",
        "packageKind": "托盘",
        "packageCount": packing.get("packageCount") or "",
        "grossWeight": packing.get("grossWeight") or "",
        "netWeight": packing.get("netWeight") or total_net_weight,
        "originCountry": "中国",
        "destinationCountry": "意大利",
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


def clear_cell(root, ref):
    set_cell(root, ref, "")


def generate_workbook(preview):
    template = active_template_path()
    if not template.exists():
        raise ValueError("未找到报关单模板，请先在管理员页面上传模板")
    OUTPUT_DIR.mkdir(exist_ok=True)
    output_name = f"报关单 IMOS {preview['contractNo']}.xlsx"
    output_path = OUTPUT_DIR / f"{uuid.uuid4().hex}-{output_name}"

    with zipfile.ZipFile(template, "r") as zin:
        workbook = ET.fromstring(zin.read("xl/workbook.xml"))
        rels = ET.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
        rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        sheets = {}
        for sheet in workbook.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
            rid = sheet.attrib[f"{{{NS_REL}}}id"]
            target = rel_map[rid].lstrip("/")
            sheets[sheet.attrib["name"]] = "xl/" + target if not target.startswith("xl/") else target

        main_path = sheets.get("Sheet1") or next(iter(sheets.values()))
        main_root = ET.fromstring(zin.read(main_path))

        values = {
            "A4": "凯斯库汽车部件（苏州）有限公司",
            "G4": preview.get("exportDateSerial") or "",
            "A6": preview["consignee"],
            "E6": "水路运输",
            "A8": "凯斯库汽车部件（苏州）有限公司",
            "A10": preview["contractNo"],
            "E10": "意大利",
            "G10": "意大利",
            "A12": preview["packageKind"],
            "E12": preview["packageCount"],
            "F12": preview["grossWeight"],
            "G12": preview["netWeight"],
            "J12": preview["tradeTerm"],
        }
        for ref, value in values.items():
            set_cell(main_root, ref, value)

        for row in range(18, 75):
            for col in ("A", "B", "D", "G", "H", "J", "K", "L", "M", "O", "R"):
                clear_cell(main_root, f"{col}{row}")

        line_rows = [18 + i * 3 for i in range(19)]
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
            set_cell(main_root, f"B{row_num + 1}", line["brand"])
            set_cell(main_root, f"G{row_num + 1}", line["netWeight"])
            set_cell(main_root, f"H{row_num + 1}", "千克")

        set_cell(main_root, "A75", "Sub Total")
        set_cell(main_root, "G75", preview["totals"]["quantity"])
        set_cell(main_root, "H75", "个")
        set_cell(main_root, "J75", preview["totals"]["amount"])
        set_cell(main_root, "K75", preview["totals"]["currency"])
        set_cell(main_root, "G76", preview["netWeight"])
        set_cell(main_root, "H76", "千克")

        modified = {main_path: ET.tostring(main_root, encoding="utf-8", xml_declaration=True)}

        if "申报要素" in sheets:
            decl_path = sheets["申报要素"]
            decl_root = ET.fromstring(zin.read(decl_path))
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

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
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
            packing = parse_packing(packing_file)
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
        meta_path = STORAGE_DIR / "meta.json"
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
    print(f"外部页面: http://localhost:{port}/u/{html.escape(PUBLIC_TOKEN)}")
    print(f"管理员页面: http://localhost:{port}/admin/{html.escape(ADMIN_TOKEN)}")
    ThreadingHTTPServer(("0.0.0.0", port), AppHandler).serve_forever()


if __name__ == "__main__":
    run()

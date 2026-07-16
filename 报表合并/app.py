#!/usr/bin/env python3
from __future__ import annotations

import cgi
import io
import json
import math
import os
import posixpath
import re
import shutil
import sqlite3
import sys
import tempfile
import threading
import time
import traceback
import urllib.parse
import uuid
import webbrowser
import zipfile
from contextlib import contextmanager
from datetime import date, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Sequence

from openpyxl import load_workbook


APP_NAME = "SuriWorkReportMerge"
BASE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
CORE_DIR = RESOURCE_DIR / "src"
if str(CORE_DIR) not in sys.path:
    sys.path.insert(0, str(CORE_DIR))

from report_merge import MergeConfig, merge_reports, preview_reports  # noqa: E402
from report_merge.rules import load_rules  # noqa: E402
from report_merge.template import inspect_template  # noqa: E402


STATIC_DIR = RESOURCE_DIR / "static"
DEFAULT_TEMPLATE = RESOURCE_DIR / "2026 Daily Export List模板.xlsx"
DEFAULT_RULES = RESOURCE_DIR / "List模板.xlsx"
MAX_FILES = 30
MAX_UPLOAD_BYTES = int(
    os.environ.get("REPORT_MERGE_MAX_UPLOAD_BYTES", str(300 * 1024 * 1024))
)
MAX_PREVIEW_ROWS = int(os.environ.get("REPORT_MERGE_PREVIEW_ROWS", "1000"))
SESSION_TTL_SECONDS = 2 * 60 * 60
DOWNLOAD_PREFIX = "/download"


def app_data_dir() -> Path:
    override = os.environ.get("SURI_REPORT_MERGE_DATA_DIR")
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
TEMP_DIR = DATA_DIR / "temp"
DB_PATH = DATA_DIR / "app.db"
META_PATH = DATA_DIR / "config.json"


def now_stamp() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def ensure_app_dirs() -> None:
    for directory in (
        DATA_DIR,
        TEMPLATE_DIR,
        RULES_DIR,
        HISTORY_DIR,
        OUTPUT_DIR,
        TEMP_DIR,
    ):
        directory.mkdir(parents=True, exist_ok=True)


def active_template_path() -> Path:
    uploaded = TEMPLATE_DIR / "current.xlsx"
    return uploaded if uploaded.exists() else DEFAULT_TEMPLATE


def active_rules_path() -> Path:
    uploaded = RULES_DIR / "current.xlsx"
    return uploaded if uploaded.exists() else DEFAULT_RULES


def json_default(value: Any) -> str:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"{type(value).__name__} 不能转换为 JSON")


def json_dumps(value: Any) -> str:
    return json.dumps(
        value,
        ensure_ascii=False,
        indent=2,
        default=json_default,
    )


def json_loads(value: str | None, fallback: Any) -> Any:
    if not value:
        return fallback
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return fallback


def safe_filename(name: str | None, fallback: str) -> str:
    clean = Path(name or fallback).name
    clean = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", clean).strip(" .")
    return clean or fallback


def db_connect() -> sqlite3.Connection:
    ensure_app_dirs()
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


@contextmanager
def db_session():
    connection = db_connect()
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def init_db() -> None:
    with db_session() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS merge_history (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                status TEXT NOT NULL,
                source_count INTEGER NOT NULL,
                source_names_json TEXT NOT NULL,
                serials_json TEXT NOT NULL,
                row_count INTEGER NOT NULL,
                total_quantity REAL NOT NULL,
                total_amount REAL NOT NULL,
                warning_count INTEGER NOT NULL,
                naming_rule_count INTEGER NOT NULL,
                missing_count INTEGER NOT NULL,
                template_name TEXT NOT NULL,
                rules_name TEXT NOT NULL,
                output_name TEXT NOT NULL,
                file_paths_json TEXT NOT NULL,
                preview_json TEXT NOT NULL
            )
            """
        )
        connection.execute(
            """
            CREATE INDEX IF NOT EXISTS idx_merge_history_created_at
            ON merge_history(created_at DESC)
            """
        )
        columns = {
            row["name"]
            for row in connection.execute(
                "PRAGMA table_info(merge_history)"
            ).fetchall()
        }
        if "naming_rule_count" not in columns:
            connection.execute(
                """
                ALTER TABLE merge_history
                ADD COLUMN naming_rule_count INTEGER NOT NULL DEFAULT 0
                """
            )
            if "fallback_count" in columns:
                connection.execute(
                    """
                    UPDATE merge_history
                    SET naming_rule_count = fallback_count
                    """
                )


def load_config_meta() -> dict[str, Any]:
    if not META_PATH.exists():
        return {}
    return json_loads(META_PATH.read_text("utf-8"), {})


def _file_meta(
    kind: str,
    path: Path,
    default_path: Path,
    stored_meta: dict[str, Any],
) -> dict[str, Any]:
    entry = stored_meta.get(kind) or {}
    uploaded = path != default_path
    return {
        "filename": entry.get("filename") if uploaded else default_path.name,
        "updatedAt": entry.get("updatedAt") if uploaded else "内置默认",
        "source": "uploaded" if uploaded else "default",
    }


def config_status() -> dict[str, Any]:
    ensure_app_dirs()
    template_path = active_template_path()
    rules_path = active_rules_path()
    if not template_path.exists():
        raise ValueError("未找到默认输出模板")
    if not rules_path.exists():
        raise ValueError("未找到默认规则文件")
    template = inspect_template(template_path)
    rules = load_rules(rules_path)
    meta = load_config_meta()
    return {
        "template": {
            **_file_meta("template", template_path, DEFAULT_TEMPLATE, meta),
            "sheet": template.sheet_name,
            "headerRow": template.header_row,
            "supportedColumns": len(template.columns),
            "unknownHeaders": list(template.unknown_headers),
        },
        "rules": {
            **_file_meta("rules", rules_path, DEFAULT_RULES, meta),
            "sheet": rules.sheet_name,
            "recordCount": len(rules.records),
            "ambiguousGroupCount": len(rules.ambiguous_records),
            "warnings": [
                {
                    "code": warning.code,
                    "message": warning.message,
                }
                for warning in rules.warnings
            ],
        },
    }


def _warning_dict(warning) -> dict[str, Any]:
    severity = "warning"
    if warning.code in {"RULE_NOT_FOUND", "RULE_NAMING_AMBIGUOUS"}:
        severity = "error"
    return {
        "code": warning.code,
        "severity": severity,
        "message": warning.message,
        "sourceFile": warning.source_file,
        "qadPn": warning.qad_pn,
        "outputRow": warning.output_row,
    }


def _row_dict(row: dict[str, Any]) -> dict[str, Any]:
    result = dict(row)
    pickup_date = result.get("pickup_date")
    if isinstance(pickup_date, (date, datetime)):
        result["pickup_date"] = pickup_date.isoformat()
    return result


def preview_payload(
    preview,
    source_names: Sequence[str],
    row_limit: int | None = None,
) -> dict[str, Any]:
    rows = [_row_dict(row) for row in preview.rows]
    exact_count = sum(
        row.get("_rule_match_type") == "exact" for row in rows
    )
    naming_rule_count = sum(
        row.get("_rule_match_type") == "naming_rule" for row in rows
    )
    missing_count = sum(
        row.get("_rule_match_type") in {"ambiguous", "missing"}
        for row in rows
    )
    warnings = [_warning_dict(warning) for warning in preview.warnings]
    summary = preview.summary()
    summary.update(
        {
            "sourceCount": len(source_names),
            "sourceNames": list(source_names),
            "exactCount": exact_count,
            "namingRuleCount": naming_rule_count,
            "missingCount": missing_count,
            "requiresConfirmation": bool(warnings),
        }
    )
    visible_rows = rows if row_limit is None else rows[:row_limit]
    return {
        "summary": summary,
        "warnings": warnings,
        "rows": visible_rows,
        "rowDisplayLimit": row_limit,
        "rowsTruncated": len(visible_rows) < len(rows),
        "templateSheet": preview.template_sheet,
        "ruleSheet": preview.rule_sheet,
    }


def _copy_source_files(
    source_paths: Sequence[str | Path],
    source_names: Sequence[str] | None,
    record_dir: Path,
) -> tuple[list[Path], list[str]]:
    destination_root = record_dir / "sources"
    destination_root.mkdir(parents=True, exist_ok=True)
    copied_paths: list[Path] = []
    copied_names: list[str] = []
    names = list(source_names or [])
    for index, raw_path in enumerate(source_paths, 1):
        source = Path(raw_path)
        original_name = (
            names[index - 1]
            if index - 1 < len(names)
            else source.name
        )
        clean_name = safe_filename(
            original_name,
            f"source-{index}{source.suffix.lower()}",
        )
        target_dir = destination_root / f"{index:02d}"
        target_dir.mkdir(parents=True, exist_ok=True)
        destination = target_dir / clean_name
        if source.resolve() != destination.resolve():
            shutil.copy2(source, destination)
        copied_paths.append(destination)
        copied_names.append(clean_name)
    return copied_paths, copied_names


def create_preview_record(
    source_paths: Sequence[str | Path],
    source_names: Sequence[str] | None = None,
    record_id: str | None = None,
) -> tuple[str, dict[str, Any]]:
    ensure_app_dirs()
    init_db()
    record_id = record_id or uuid.uuid4().hex
    record_dir = HISTORY_DIR / record_id
    if record_dir.exists():
        raise ValueError("历史记录 ID 已存在")
    record_dir.mkdir(parents=True)
    try:
        copied_sources, copied_names = _copy_source_files(
            source_paths,
            source_names,
            record_dir,
        )
        template_path = record_dir / "template.xlsx"
        rules_path = record_dir / "rules.xlsx"
        shutil.copy2(active_template_path(), template_path)
        shutil.copy2(active_rules_path(), rules_path)
        active_config = config_status()

        preview = preview_reports(
            copied_sources,
            template_path,
            rules_path,
            MergeConfig(
                max_files=MAX_FILES,
                unmatched_policy="warn",
                allow_naming_rule_match=True,
            ),
        )
        full_payload = preview_payload(preview, copied_names, row_limit=None)
        response_payload = preview_payload(
            preview,
            copied_names,
            row_limit=MAX_PREVIEW_ROWS,
        )
        preview_path = record_dir / "preview.json"
        preview_path.write_text(json_dumps(full_payload), "utf-8")

        summary = full_payload["summary"]
        serials = list((summary.get("bySerial") or {}).keys())
        file_paths = {
            "sources": [str(path) for path in copied_sources],
            "template": str(template_path),
            "rules": str(rules_path),
            "preview": str(preview_path),
            "output": "",
        }
        with db_session() as connection:
            connection.execute(
                """
                INSERT INTO merge_history (
                    id, created_at, status, source_count, source_names_json,
                    serials_json, row_count, total_quantity, total_amount,
                    warning_count, naming_rule_count, missing_count,
                    template_name, rules_name, output_name,
                    file_paths_json, preview_json
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    record_id,
                    now_stamp(),
                    "previewed",
                    len(copied_sources),
                    json_dumps(copied_names),
                    json_dumps(serials),
                    int(summary["rowCount"]),
                    float(summary["totalQuantity"]),
                    float(summary["totalAmount"]),
                    len(full_payload["warnings"]),
                    int(summary["namingRuleCount"]),
                    int(summary["missingCount"]),
                    active_config["template"]["filename"],
                    active_config["rules"]["filename"],
                    "",
                    json_dumps(file_paths),
                    json_dumps(full_payload),
                ),
            )
        return record_id, response_payload
    except Exception:
        shutil.rmtree(record_dir, ignore_errors=True)
        raise


def _history_row(record_id: str) -> sqlite3.Row:
    init_db()
    with db_session() as connection:
        row = connection.execute(
            "SELECT * FROM merge_history WHERE id = ?",
            (record_id,),
        ).fetchone()
    if row is None:
        raise ValueError("未找到历史记录")
    return row


def _output_filename(row: sqlite3.Row) -> str:
    serials = json_loads(row["serials_json"], [])
    if len(serials) == 1:
        label = serials[0]
    elif serials:
        label = f"{serials[0]}等{row['source_count']}份"
    else:
        label = f"{row['source_count']}份报表"
    return safe_filename(f"报表合并-{label}.xlsx", "报表合并结果.xlsx")


def generate_history_record(record_id: str) -> tuple[Path, str, dict[str, Any]]:
    row = _history_row(record_id)
    paths = json_loads(row["file_paths_json"], {})
    source_paths = [Path(path) for path in paths.get("sources") or []]
    template_path = Path(paths.get("template") or "")
    rules_path = Path(paths.get("rules") or "")
    for path in [*source_paths, template_path, rules_path]:
        if not path.is_file():
            raise ValueError(f"历史记录所需文件不存在：{path.name}")

    record_dir = HISTORY_DIR / record_id
    output_path = record_dir / "output.xlsx"
    result = merge_reports(
        source_paths,
        template_path,
        rules_path,
        output_path,
        MergeConfig(
            max_files=MAX_FILES,
            unmatched_policy="warn",
            allow_naming_rule_match=True,
        ),
    )
    stored_preview = json_loads(row["preview_json"], {})
    stored_summary = stored_preview.get("summary") or {}
    current_summary = result.summary()
    if int(stored_summary.get("rowCount") or 0) != int(
        current_summary.get("rowCount") or 0
    ):
        raise ValueError("生成结果与已确认预览不一致，请重新上传并预览")
    for key in ("totalQuantity", "totalAmount"):
        if abs(
            float(stored_summary.get(key) or 0)
            - float(current_summary.get(key) or 0)
        ) > 0.000001:
            raise ValueError("生成结果与已确认预览不一致，请重新上传并预览")

    output_name = _output_filename(row)
    paths["output"] = str(output_path)
    with db_session() as connection:
        connection.execute(
            """
            UPDATE merge_history
               SET status = 'generated',
                   output_name = ?,
                   file_paths_json = ?
             WHERE id = ?
            """,
            (output_name, json_dumps(paths), record_id),
        )
    return output_path, output_name, current_summary


def row_to_history(
    row: sqlite3.Row,
    include_preview: bool = False,
) -> dict[str, Any]:
    result = {
        "id": row["id"],
        "createdAt": row["created_at"],
        "status": row["status"],
        "sourceCount": row["source_count"],
        "sourceNames": json_loads(row["source_names_json"], []),
        "serials": json_loads(row["serials_json"], []),
        "rowCount": row["row_count"],
        "totalQuantity": row["total_quantity"],
        "totalAmount": row["total_amount"],
        "warningCount": row["warning_count"],
        "namingRuleCount": row["naming_rule_count"],
        "missingCount": row["missing_count"],
        "templateName": row["template_name"],
        "rulesName": row["rules_name"],
        "outputName": row["output_name"],
    }
    if include_preview:
        result["preview"] = json_loads(row["preview_json"], {})
    return result


def normalize_history_datetime(value: str, end_of_day: bool = False) -> str:
    text = urllib.parse.unquote(value or "").replace("T", " ").strip()
    if not text:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return f"{text} {'23:59:59' if end_of_day else '00:00:00'}"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}", text):
        return f"{text}:{'59' if end_of_day else '00'}"
    if re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", text):
        return text
    raise ValueError("历史查询时间格式不正确")


def update_active_config(
    template_upload: tuple[Path, str] | None = None,
    rules_upload: tuple[Path, str] | None = None,
) -> dict[str, Any]:
    if template_upload is None and rules_upload is None:
        raise ValueError("请选择要上传的模板或规则文件")
    ensure_app_dirs()
    validated: dict[str, dict[str, Any]] = {}
    if template_upload is not None:
        template_path, original_name = template_upload
        spec = inspect_template(template_path)
        validated["template"] = {
            "path": template_path,
            "filename": safe_filename(original_name, "template.xlsx"),
            "sheet": spec.sheet_name,
            "headerRow": spec.header_row,
            "supportedColumns": len(spec.columns),
            "unknownHeaders": list(spec.unknown_headers),
        }
    if rules_upload is not None:
        rules_path, original_name = rules_upload
        rule_set = load_rules(rules_path)
        validated["rules"] = {
            "path": rules_path,
            "filename": safe_filename(original_name, "rules.xlsx"),
            "sheet": rule_set.sheet_name,
            "recordCount": len(rule_set.records),
            "ambiguousGroupCount": len(rule_set.ambiguous_records),
            "warnings": [
                warning.message for warning in rule_set.warnings
            ],
        }

    meta = load_config_meta()
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    for kind, item in validated.items():
        target_dir = TEMPLATE_DIR if kind == "template" else RULES_DIR
        version_dir = target_dir / "versions"
        version_dir.mkdir(parents=True, exist_ok=True)
        version_name = safe_filename(
            f"{timestamp}-{uuid.uuid4().hex[:8]}-{item['filename']}",
            f"{timestamp}-{kind}.xlsx",
        )
        shutil.copy2(item["path"], version_dir / version_name)
        temporary_target = target_dir / f".current-{uuid.uuid4().hex}.xlsx"
        shutil.copy2(item["path"], temporary_target)
        os.replace(temporary_target, target_dir / "current.xlsx")
        meta[kind] = {
            "filename": item["filename"],
            "updatedAt": now_stamp(),
            "versionFile": version_name,
        }
    META_PATH.write_text(json_dumps(meta), "utf-8")
    return {
        "updated": {
            kind: {
                key: value
                for key, value in item.items()
                if key != "path"
            }
            for kind, item in validated.items()
        },
        "active": config_status(),
    }


def validate_xlsx(path: Path) -> None:
    try:
        workbook = load_workbook(path, data_only=False, read_only=True)
        workbook.close()
    except Exception as exc:
        raise ValueError("文件不是有效的 .xlsx 工作簿") from exc


def parse_upload(handler: BaseHTTPRequestHandler):
    length = int(handler.headers.get("Content-Length", "0"))
    if length <= 0:
        raise ValueError("没有收到上传内容")
    if length > MAX_UPLOAD_BYTES:
        raise ValueError(
            f"本次上传超过 {MAX_UPLOAD_BYTES // 1024 // 1024} MB 限制"
        )
    environment = {
        "REQUEST_METHOD": "POST",
        "CONTENT_TYPE": handler.headers.get("Content-Type", ""),
        "CONTENT_LENGTH": str(length),
    }
    return cgi.FieldStorage(
        fp=handler.rfile,
        headers=handler.headers,
        environ=environment,
    )


def form_file_fields(form, names: Sequence[str]) -> list[Any]:
    result: list[Any] = []
    for name in names:
        if name not in form:
            continue
        value = form[name]
        result.extend(value if isinstance(value, list) else [value])
    return [
        field
        for field in result
        if getattr(field, "filename", "")
    ]


def save_upload_field(
    field,
    directory: Path,
    fallback: str,
    allowed_suffixes: set[str],
) -> tuple[Path, str]:
    original_name = safe_filename(
        getattr(field, "filename", ""),
        fallback,
    )
    suffix = Path(original_name).suffix.lower()
    if suffix not in allowed_suffixes:
        raise ValueError(
            f"{original_name} 格式不支持；允许："
            + "、".join(sorted(allowed_suffixes))
        )
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f"{uuid.uuid4().hex}{suffix}"
    with path.open("wb") as output:
        shutil.copyfileobj(field.file, output)
    if path.stat().st_size == 0:
        raise ValueError(f"{original_name} 是空文件")
    return path, original_name


def json_response(
    handler: BaseHTTPRequestHandler,
    status: int,
    payload: dict[str, Any],
) -> None:
    body = json_dumps(payload).encode("utf-8")
    handler.send_response(status)
    handler.send_header("Content-Type", "application/json; charset=utf-8")
    handler.send_header("Content-Length", str(len(body)))
    handler.end_headers()
    handler.wfile.write(body)


def log_runtime_error(message: str) -> None:
    try:
        ensure_app_dirs()
        with (DATA_DIR / "error.log").open("a", encoding="utf-8") as handle:
            handle.write(f"[{now_stamp()}]\n{message.rstrip()}\n\n")
    except Exception:
        pass


SESSIONS: dict[str, dict[str, Any]] = {}


def cleanup_sessions() -> None:
    cutoff = time.time() - SESSION_TTL_SECONDS
    stale = [
        token
        for token, item in SESSIONS.items()
        if float(item.get("created") or 0) < cutoff
    ]
    for token in stale:
        SESSIONS.pop(token, None)


class AppHandler(BaseHTTPRequestHandler):
    server_version = "SuriReportMerge/1.0"

    def log_message(self, fmt: str, *args) -> None:
        print(f"[{now_stamp()}] {self.address_string()} {fmt % args}")

    def do_GET(self) -> None:
        try:
            cleanup_sessions()
            parsed = urllib.parse.urlparse(self.path)
            path = parsed.path
            if path == "/":
                self.serve_static("index.html")
                return
            if path == "/favicon.ico":
                self.send_response(204)
                self.end_headers()
                return
            if path.startswith("/static/"):
                self.serve_static(path[len("/static/") :])
                return
            if path == "/api/config":
                json_response(
                    self,
                    200,
                    {"ok": True, "active": config_status()},
                )
                return
            if path == "/api/history":
                self.handle_history_list(
                    urllib.parse.parse_qs(parsed.query)
                )
                return
            if path.startswith("/api/history/"):
                parts = path.strip("/").split("/")
                if len(parts) == 3:
                    row = _history_row(parts[2])
                    json_response(
                        self,
                        200,
                        {
                            "ok": True,
                            "record": row_to_history(
                                row,
                                include_preview=True,
                            ),
                        },
                    )
                    return
                if len(parts) == 4 and parts[3] == "download":
                    self.handle_history_download(
                        parts[2],
                        urllib.parse.parse_qs(parsed.query),
                    )
                    return
            if path.startswith("/download/"):
                self.handle_session_download(path.rsplit("/", 1)[-1])
                return
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            log_runtime_error("GET failed:\n" + traceback.format_exc())
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def do_POST(self) -> None:
        try:
            cleanup_sessions()
            if self.path == "/api/config":
                self.handle_config_upload()
            elif self.path == "/api/preview":
                self.handle_preview()
            elif self.path == "/api/generate":
                self.handle_generate()
            else:
                self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            log_runtime_error("POST failed:\n" + traceback.format_exc())
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def do_DELETE(self) -> None:
        try:
            parts = urllib.parse.urlparse(self.path).path.strip("/").split("/")
            if len(parts) == 3 and parts[:2] == ["api", "history"]:
                self.handle_history_delete(parts[2])
                return
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
        except Exception as exc:
            log_runtime_error("DELETE failed:\n" + traceback.format_exc())
            json_response(self, 400, {"ok": False, "error": str(exc)})

    def serve_static(self, name: str) -> None:
        clean = posixpath.normpath(urllib.parse.unquote(name)).lstrip("/")
        path = (STATIC_DIR / clean).resolve()
        static_root = STATIC_DIR.resolve()
        if (
            os.path.commonpath([str(path), str(static_root)])
            != str(static_root)
            or not path.is_file()
        ):
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        mime = "application/octet-stream"
        if path.suffix == ".html":
            mime = "text/html; charset=utf-8"
        elif path.suffix == ".css":
            mime = "text/css; charset=utf-8"
        elif path.suffix == ".js":
            mime = "application/javascript; charset=utf-8"
        body = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_path(
        self,
        path: Path,
        filename: str,
        content_type: str = "application/octet-stream",
    ) -> None:
        if not path.is_file():
            raise ValueError("下载文件不存在")
        encoded = urllib.parse.quote(filename)
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header(
            "Content-Disposition",
            f"attachment; filename*=UTF-8''{encoded}",
        )
        self.send_header("Content-Length", str(path.stat().st_size))
        self.end_headers()
        with path.open("rb") as source:
            shutil.copyfileobj(source, self.wfile, length=1024 * 1024)

    def handle_session_download(self, token: str) -> None:
        item = SESSIONS.get(token)
        if not item:
            raise ValueError("下载链接已过期")
        path = Path(item["path"])
        validate_xlsx(path)
        self.send_path(
            path,
            item["filename"],
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    def handle_config_upload(self) -> None:
        upload_dir = Path(
            tempfile.mkdtemp(prefix="config-", dir=TEMP_DIR)
        )
        try:
            form = parse_upload(self)
            template_upload = None
            rules_upload = None
            if "template" in form and getattr(
                form["template"],
                "filename",
                "",
            ):
                template_upload = save_upload_field(
                    form["template"],
                    upload_dir,
                    "template.xlsx",
                    {".xlsx"},
                )
            if "rules" in form and getattr(
                form["rules"],
                "filename",
                "",
            ):
                rules_upload = save_upload_field(
                    form["rules"],
                    upload_dir,
                    "rules.xlsx",
                    {".xlsx"},
                )
            result = update_active_config(template_upload, rules_upload)
            json_response(self, 200, {"ok": True, **result})
        finally:
            shutil.rmtree(upload_dir, ignore_errors=True)

    def handle_preview(self) -> None:
        upload_dir = Path(
            tempfile.mkdtemp(prefix="sources-", dir=TEMP_DIR)
        )
        try:
            form = parse_upload(self)
            fields = form_file_fields(form, ("documents", "files"))
            if not 1 <= len(fields) <= MAX_FILES:
                raise ValueError(
                    f"请一次选择 1～{MAX_FILES} 份源报表；"
                    f"当前为 {len(fields)} 份"
                )
            paths: list[Path] = []
            names: list[str] = []
            for index, field in enumerate(fields, 1):
                path, name = save_upload_field(
                    field,
                    upload_dir,
                    f"source-{index}.xlsx",
                    {".xls", ".xlsx"},
                )
                paths.append(path)
                names.append(name)
            history_id, preview = create_preview_record(paths, names)
            session_id = uuid.uuid4().hex
            SESSIONS[session_id] = {
                "historyId": history_id,
                "created": time.time(),
            }
            json_response(
                self,
                200,
                {
                    "ok": True,
                    "sessionId": session_id,
                    "historyId": history_id,
                    "preview": preview,
                },
            )
        finally:
            shutil.rmtree(upload_dir, ignore_errors=True)

    def handle_generate(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length)
        data = json.loads(body.decode("utf-8") or "{}")
        if data.get("confirmed") is not True:
            raise ValueError("请先确认预览和黄色/红色异常")
        session_id = str(data.get("sessionId") or "")
        session = SESSIONS.get(session_id) or {}
        history_id = str(
            data.get("historyId")
            or session.get("historyId")
            or ""
        )
        if not history_id:
            raise ValueError("预览会话已过期，请重新上传并预览")
        output_path, output_name, summary = generate_history_record(history_id)
        token = uuid.uuid4().hex
        SESSIONS[token] = {
            "path": str(output_path),
            "filename": output_name,
            "created": time.time(),
        }
        json_response(
            self,
            200,
            {
                "ok": True,
                "downloadUrl": f"{DOWNLOAD_PREFIX}/{token}",
                "filename": output_name,
                "historyId": history_id,
                "savedPath": str(output_path),
                "summary": summary,
            },
        )

    def handle_history_list(self, query: dict[str, list[str]]) -> None:
        init_db()
        start_at = normalize_history_datetime(
            (query.get("start") or [""])[0]
        )
        end_at = normalize_history_datetime(
            (query.get("end") or [""])[0],
            end_of_day=True,
        )
        if start_at and end_at and start_at > end_at:
            raise ValueError("开始时间不能晚于结束时间")
        try:
            limit = max(
                1,
                min(int((query.get("limit") or ["5"])[0]), 50),
            )
        except ValueError:
            limit = 5
        try:
            page = max(1, int((query.get("page") or ["1"])[0]))
        except ValueError:
            page = 1

        conditions: list[str] = []
        params: list[Any] = []
        if start_at:
            conditions.append("created_at >= ?")
            params.append(start_at)
        if end_at:
            conditions.append("created_at <= ?")
            params.append(end_at)
        where_sql = (
            "WHERE " + " AND ".join(conditions)
            if conditions
            else ""
        )
        with db_session() as connection:
            total = connection.execute(
                f"SELECT COUNT(*) FROM merge_history {where_sql}",
                params,
            ).fetchone()[0]
            total_pages = max(1, math.ceil(total / limit))
            page = min(page, total_pages)
            rows = connection.execute(
                f"""
                SELECT * FROM merge_history
                {where_sql}
                ORDER BY datetime(created_at) DESC, created_at DESC
                LIMIT ? OFFSET ?
                """,
                (*params, limit, (page - 1) * limit),
            ).fetchall()
        json_response(
            self,
            200,
            {
                "ok": True,
                "history": [row_to_history(row) for row in rows],
                "filters": {
                    "start": start_at,
                    "end": end_at,
                    "limit": limit,
                    "page": page,
                },
                "pagination": {
                    "page": page,
                    "limit": limit,
                    "total": total,
                    "totalPages": total_pages,
                },
            },
        )

    def handle_history_download(
        self,
        record_id: str,
        query: dict[str, list[str]],
    ) -> None:
        kind = (query.get("kind") or [""])[0]
        if kind not in {"sources", "template", "rules", "preview", "output"}:
            raise ValueError("下载类型不正确")
        row = _history_row(record_id)
        paths = json_loads(row["file_paths_json"], {})
        if kind == "sources":
            zip_path = TEMP_DIR / f"{record_id}-{uuid.uuid4().hex}-源报表.zip"
            source_names = json_loads(row["source_names_json"], [])
            try:
                with zipfile.ZipFile(
                    zip_path,
                    "w",
                    zipfile.ZIP_DEFLATED,
                ) as archive:
                    for index, raw_path in enumerate(
                        paths.get("sources") or [],
                        1,
                    ):
                        source = Path(raw_path)
                        if source.is_file():
                            name = (
                                source_names[index - 1]
                                if index - 1 < len(source_names)
                                else source.name
                            )
                            archive.write(
                                source,
                                arcname=(
                                    f"{index:02d}-"
                                    f"{safe_filename(name, source.name)}"
                                ),
                            )
                self.send_path(
                    zip_path,
                    f"{record_id}-源报表.zip",
                    "application/zip",
                )
            finally:
                zip_path.unlink(missing_ok=True)
            return

        path = Path(paths.get(kind) or "")
        names = {
            "template": row["template_name"] or "template.xlsx",
            "rules": row["rules_name"] or "rules.xlsx",
            "preview": f"{record_id}-预览.json",
            "output": row["output_name"] or "报表合并结果.xlsx",
        }
        content_types = {
            "template": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "rules": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "preview": "application/json; charset=utf-8",
            "output": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }
        if kind in {"template", "rules", "output"}:
            validate_xlsx(path)
        self.send_path(path, names[kind], content_types[kind])

    def handle_history_delete(self, record_id: str) -> None:
        _history_row(record_id)
        with db_session() as connection:
            connection.execute(
                "DELETE FROM merge_history WHERE id = ?",
                (record_id,),
            )
        shutil.rmtree(HISTORY_DIR / record_id, ignore_errors=True)
        json_response(self, 200, {"ok": True})


def run(
    host: str = "127.0.0.1",
    port: int | None = None,
    open_browser: bool = False,
) -> None:
    ensure_app_dirs()
    init_db()
    port = int(port if port is not None else os.environ.get("PORT", "8000"))
    url = f"http://{host}:{port}/"
    print(f"报表合并页面: {url}")
    if open_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    ThreadingHTTPServer((host, port), AppHandler).serve_forever()


def run_in_thread(
    host: str = "127.0.0.1",
    port: int | None = None,
) -> tuple[ThreadingHTTPServer, str]:
    ensure_app_dirs()
    init_db()
    port = int(port if port is not None else os.environ.get("PORT", "8000"))
    server = ThreadingHTTPServer((host, port), AppHandler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server, f"http://{host}:{port}/"


if __name__ == "__main__":
    run(open_browser=os.environ.get("SURI_OPEN_BROWSER") == "1")

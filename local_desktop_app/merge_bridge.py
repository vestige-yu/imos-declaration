"""Load the report-merge application as the second module of the desktop app."""

import importlib.util
import sys
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", BASE_DIR))
SOURCE_APP_DIR = BASE_DIR.parent / "报表合并"

for core_dir in (
    RESOURCE_DIR / "report_merge_src",
    SOURCE_APP_DIR / "src",
):
    if core_dir.is_dir() and str(core_dir) not in sys.path:
        sys.path.insert(0, str(core_dir))

# These imports are intentionally explicit so PyInstaller includes dynamic dependencies.
import openpyxl  # noqa: F401,E402
import report_merge  # noqa: F401,E402


def _service_path():
    candidates = (
        RESOURCE_DIR / "merge_module" / "app.py",
        SOURCE_APP_DIR / "app.py",
    )
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise RuntimeError("未找到报表合并模块")


def _load_service():
    path = _service_path()
    spec = importlib.util.spec_from_file_location("suri_report_merge_app", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("报表合并模块无法加载")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.DOWNLOAD_PREFIX = "/api/merge/download"
    return module


service = _load_service()

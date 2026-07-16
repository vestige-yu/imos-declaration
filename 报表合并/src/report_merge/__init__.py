"""报表合并核心能力。"""

from .engine import MergeConfig, merge_reports, preview_reports
from .models import (
    MergePreview,
    MergeResult,
    MergeWarning,
    RuleRecord,
    SourceLine,
    SourceReport,
)

__all__ = [
    "MergeConfig",
    "MergePreview",
    "MergeResult",
    "MergeWarning",
    "RuleRecord",
    "SourceLine",
    "SourceReport",
    "merge_reports",
    "preview_reports",
]

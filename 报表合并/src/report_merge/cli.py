from __future__ import annotations

import argparse
import json
from pathlib import Path

from .engine import MergeConfig, merge_reports


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="把最多 30 份 Excel 源报表按客户模板和规则表合并为一份报表"
    )
    parser.add_argument(
        "sources",
        nargs="+",
        help="待合并的 .xls/.xlsx 源报表，可传 1 到 30 份",
    )
    parser.add_argument(
        "--template",
        required=True,
        help="客户上传的 .xlsx 输出模板",
    )
    parser.add_argument(
        "--rules",
        required=True,
        help="客户上传的 .xlsx 匹配规则",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="生成的 .xlsx 文件路径",
    )
    parser.add_argument(
        "--unmatched-policy",
        choices=("warn", "block"),
        default="warn",
        help="规则未匹配时继续并警告，或阻止生成",
    )
    parser.add_argument(
        "--strict-part-number",
        action="store_true",
        help="禁用客户确认的料号命名规则，只允许原文精确匹配",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    result = merge_reports(
        [Path(path) for path in args.sources],
        Path(args.template),
        Path(args.rules),
        Path(args.output),
        MergeConfig(
            unmatched_policy=args.unmatched_policy,
            allow_naming_rule_match=not args.strict_part_number,
        ),
    )
    print(json.dumps(result.summary(), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

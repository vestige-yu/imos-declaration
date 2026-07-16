from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Sequence

from .models import MergePreview, MergeResult, MergeWarning
from .rules import load_rules
from .source_parser import parse_source_report
from .template import inspect_template, write_output


@dataclass(frozen=True)
class MergeConfig:
    max_files: int = 30
    unmatched_policy: Literal["warn", "block"] = "warn"
    allow_naming_rule_match: bool = True


def _validate_common_inputs(
    source_files: Sequence[str | Path],
    template_path: str | Path,
    rules_path: str | Path,
    config: MergeConfig,
) -> tuple[list[Path], Path, Path]:
    if not source_files:
        raise ValueError("请至少上传 1 份源报表")
    if len(source_files) > config.max_files:
        raise ValueError(
            f"一次最多处理 {config.max_files} 份源报表；"
            f"当前为 {len(source_files)} 份"
        )
    sources = [Path(path) for path in source_files]
    for source in sources:
        if not source.is_file():
            raise ValueError(f"源报表不存在：{source}")
        if source.suffix.lower() not in {".xls", ".xlsx"}:
            raise ValueError(
                f"源报表 {source.name} 格式不支持；仅支持 .xls 和 .xlsx"
            )

    template = Path(template_path)
    rules = Path(rules_path)
    if not template.is_file() or template.suffix.lower() != ".xlsx":
        raise ValueError("输出模板必须是存在的 .xlsx 文件")
    if not rules.is_file() or rules.suffix.lower() != ".xlsx":
        raise ValueError("规则文件必须是存在的 .xlsx 文件")
    if config.unmatched_policy not in {"warn", "block"}:
        raise ValueError("unmatched_policy 仅支持 warn 或 block")
    return sources, template, rules


def _prepare_rows(
    sources: Sequence[Path],
    template: Path,
    rules_path: Path,
    active_config: MergeConfig,
) -> tuple[object, object, list[dict[str, object]], list[MergeWarning]]:
    template_spec = inspect_template(template)
    rule_set = load_rules(rules_path)
    warnings = list(rule_set.warnings)
    if template_spec.unknown_headers:
        warnings.append(
            MergeWarning(
                code="TEMPLATE_UNKNOWN_COLUMNS",
                message=(
                    "模板中存在尚未支持的列，输出将保留列名但不填值："
                    + "、".join(template_spec.unknown_headers)
                ),
            )
        )

    merged_rows: list[dict[str, object]] = []
    for source in sources:
        report = parse_source_report(source)
        for line in report.lines:
            match = rule_set.match(
                line.qad_pn,
                allow_naming_rule_match=active_config.allow_naming_rule_match,
            )
            rule = match.record
            match_type = match.match_type
            output_row = template_spec.header_row + len(merged_rows) + 1
            if match_type == "naming_rule" and rule is not None:
                applied = "、".join(match.applied_rules) or "客户确认规则"
                warnings.append(
                    MergeWarning(
                        code="RULE_NAMING_MATCH",
                        message=(
                            f"{line.qad_pn} 未原文精确匹配，已按"
                            f"“{applied}”使用 {rule.qad_pn} 的规则"
                        ),
                        source_file=line.source_file,
                        qad_pn=line.qad_pn,
                        output_row=output_row,
                    )
                )
            elif match_type == "ambiguous":
                candidate_text = "、".join(
                    record.qad_pn for record in match.candidates[:5]
                )
                warning = MergeWarning(
                    code="RULE_NAMING_AMBIGUOUS",
                    message=(
                        f"{line.qad_pn} 按料号命名规则对应多个相互冲突的"
                        f"规则料号（{candidate_text}），未自动选取"
                    ),
                    source_file=line.source_file,
                    qad_pn=line.qad_pn,
                    output_row=output_row,
                )
                if active_config.unmatched_policy == "block":
                    raise ValueError(
                        f"{line.source_file} 的 {line.qad_pn} 对应多个冲突规则，"
                        "已按 block 策略停止生成"
                    )
                warnings.append(warning)
            elif match_type == "missing":
                warning = MergeWarning(
                    code="RULE_NOT_FOUND",
                    message=f"{line.qad_pn} 在规则文件中没有匹配记录",
                    source_file=line.source_file,
                    qad_pn=line.qad_pn,
                    output_row=output_row,
                )
                if active_config.unmatched_policy == "block":
                    raise ValueError(
                        f"{line.source_file} 的 {line.qad_pn} 未匹配规则，"
                        "已按 block 策略停止生成"
                    )
                warnings.append(warning)

            merged_rows.append(
                {
                    "serial_no": report.serial_no,
                    "qad_pn": line.qad_pn,
                    "quantity": line.quantity,
                    "amount": line.amount,
                    "currency": line.currency or report.currency,
                    "pickup_date": report.pickup_date,
                    "ship_to": report.ship_to,
                    "delivery_way": report.delivery_way,
                    "item_no": line.item_no,
                    "goods_name": rule.goods_name if rule else "",
                    "hs_code": rule.hs_code if rule else "",
                    "brand": rule.brand if rule else "",
                    "description_en": line.description_en,
                    "unit_price": line.unit_price,
                    "po_no": line.po_no,
                    "source_file": line.source_file,
                    "_rule_match_type": match_type,
                    "_matched_qad_pn": rule.qad_pn if rule else "",
                    "_applied_part_rules": list(match.applied_rules),
                }
            )
    return template_spec, rule_set, merged_rows, warnings


def preview_reports(
    source_files: Sequence[str | Path],
    template_path: str | Path,
    rules_path: str | Path,
    config: MergeConfig | None = None,
) -> MergePreview:
    active_config = config or MergeConfig()
    sources, template, rules_path = _validate_common_inputs(
        source_files,
        template_path,
        rules_path,
        active_config,
    )
    template_spec, rule_set, merged_rows, warnings = _prepare_rows(
        sources,
        template,
        rules_path,
        active_config,
    )
    return MergePreview(
        template_sheet=template_spec.sheet_name,
        rule_sheet=rule_set.sheet_name,
        rows=merged_rows,
        warnings=warnings,
    )


def merge_reports(
    source_files: Sequence[str | Path],
    template_path: str | Path,
    rules_path: str | Path,
    output_path: str | Path,
    config: MergeConfig | None = None,
) -> MergeResult:
    active_config = config or MergeConfig()
    sources, template, rules_path = _validate_common_inputs(
        source_files,
        template_path,
        rules_path,
        active_config,
    )
    output = Path(output_path)
    if output.suffix.lower() != ".xlsx":
        raise ValueError("输出文件必须使用 .xlsx 扩展名")
    protected_inputs = {
        source.resolve() for source in sources
    } | {template.resolve(), rules_path.resolve()}
    if output.resolve() in protected_inputs:
        raise ValueError("输出路径不能覆盖任何源报表、模板或规则文件")
    template_spec, rule_set, merged_rows, warnings = _prepare_rows(
        sources,
        template,
        rules_path,
        active_config,
    )

    write_output(
        template,
        template_spec,
        merged_rows,
        output,
    )
    return MergeResult(
        output_path=output,
        template_sheet=template_spec.sheet_name,
        rule_sheet=rule_set.sheet_name,
        rows=merged_rows,
        warnings=warnings,
    )

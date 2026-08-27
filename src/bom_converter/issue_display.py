from __future__ import annotations

import json
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Mapping

from .models import ConversionResult, Issue


SEVERITY_PRESENTATION = {
    "info": ("处理提示", "info"),
    "warning": ("需要检查", "warning"),
    "error": ("不能正式使用", "error"),
}


ISSUE_GUIDANCE: dict[str, tuple[str, str, str]] = {
    "UNKNOWN_PROFILE": (
        "未匹配已知格式",
        "文件没有完整匹配当前内置的已知格式，需要先确认通用字段映射。",
        "返回检查字段，核对表头范围和通用映射后再转换。",
    ),
    "LOW_HEADER_CONFIDENCE": (
        "表头识别把握较低",
        "自动识别到的字段数量或表头特征不足，可能仍有列没有正确识别。",
        "检查表头范围、完整表头路径和未识别列，必要时使用手动配置格式。",
    ),
    "GENERATED_NAME": (
        "名称由组合规则生成",
        "原名称为空或当前规则要求组合名称，程序已按本机规则生成最终名称。",
        "检查名称生成规则窗口中的前五行预览，确认组合顺序和分隔符。",
    ),
    "LEVEL_MARKER_MULTIPLE": (
        "同一行存在多个层级标记",
        "一个零件行在层级列组中出现了多个有效标记，程序不能可靠判断唯一层级。",
        "打开来源工作表对应行，只保留一个正确层级标记后重新转换。",
    ),
    "LEVEL_MARKER_MISSING": (
        "没有检测到层级标记",
        "该零件行的层级列组中没有找到 1、Y、Yes、√、✓ 或 ● 等有效标记。",
        "检查来源行或层级列组设置，补充/确认正确层级后重新转换。",
    ),
    "UNIT_CHANGED": (
        "单位与已保存规则不同",
        "当前表头单位和历史规则记录的单位不同，直接复用旧规则可能造成数值含义变化。",
        "返回检查字段，确认当前来源单位；确认无误后再转换并按需更新规则。",
    ),
    "IMAGE_LIMIT": (
        "图片超过模板槽位",
        "来源行中的图片数量超过当前标准模板可写入的图片位置。",
        "检查图片槽位设置和标准模板容量，确认哪些图片需要保留后重新转换。",
    ),
    "WEIGHT_MISMATCH": (
        "重量关系不一致",
        "单件重量乘以数量与总重量不一致，程序保留了来源值，没有自行改写。",
        "核对来源行的单件重量、数量和总重量；若样本数值已脱敏，可记录后跳过算术核对。",
    ),
}


@dataclass(frozen=True)
class IssueDisplay:
    code: str
    severity: str
    severity_label: str
    title: str
    description: str
    source_sheet: str
    source_row: str
    field: str
    action: str
    tag: str

    def table_values(self) -> tuple[str, ...]:
        return (
            self.severity_label,
            self.title,
            self.description,
            self.source_sheet,
            self.source_row,
            self.field,
            self.action,
        )


@dataclass(frozen=True)
class ResultCounts:
    completely_successful_files: int
    files_with_warnings: int
    failed_files: int
    info_count: int
    warning_count: int
    error_count: int

    @property
    def file_count(self) -> int:
        return (
            self.completely_successful_files
            + self.files_with_warnings
            + self.failed_files
        )

    def summary_text(self) -> str:
        return (
            f"共 {self.file_count} 个文件：{self.completely_successful_files} 个完全成功文件，"
            f"{self.files_with_warnings} 个含警告文件，{self.failed_files} 个失败文件；"
            f"{self.warning_count} 条警告，{self.error_count} 条错误，"
            f"{self.info_count} 条处理提示。"
        )


def _text_or_dash(value: object) -> str:
    if value is None:
        return "—"
    text = str(value).strip()
    return text or "—"


def present_issue(
    issue: Issue,
    field_labels: Mapping[str, str] | None = None,
) -> IssueDisplay:
    severity = str(issue.severity or "warning").strip().lower()
    if severity not in SEVERITY_PRESENTATION:
        severity = "warning"
    severity_label, tag = SEVERITY_PRESENTATION[severity]
    default_title, default_description, action = ISSUE_GUIDANCE.get(
        issue.code,
        (
            "其他处理问题",
            "程序记录了一个尚未归入常见类型的问题。",
            "根据下方说明检查对应工作表、行和字段；不确定时不要把结果用于正式业务。",
        ),
    )
    title = f"{default_title}（{issue.code}）" if issue.code else default_title
    description = _text_or_dash(issue.message)
    if description == "—":
        description = default_description
    field = _text_or_dash(issue.field)
    if issue.field and field_labels and issue.field in field_labels:
        field = field_labels[issue.field]
    return IssueDisplay(
        code=issue.code,
        severity=severity,
        severity_label=severity_label,
        title=title,
        description=description,
        source_sheet=_text_or_dash(issue.source_sheet),
        source_row=_text_or_dash(issue.source_row),
        field=field,
        action=action,
        tag=tag,
    )


def count_results(results: Iterable[ConversionResult]) -> ResultCounts:
    successful = warning_files = failed = 0
    info_count = warning_count = error_count = 0
    for result in results:
        result_info = sum(str(issue.severity).lower() == "info" for issue in result.issues)
        result_warnings = sum(str(issue.severity).lower() == "warning" for issue in result.issues)
        result_errors = sum(str(issue.severity).lower() == "error" for issue in result.issues)
        info_count += result_info
        warning_count += result_warnings
        error_count += result_errors
        if result_errors:
            failed += 1
        elif result_warnings:
            warning_files += 1
        else:
            successful += 1
    return ResultCounts(successful, warning_files, failed, info_count, warning_count, error_count)


def issue_summary_text(result: ConversionResult, field_labels: Mapping[str, str] | None = None) -> str:
    displays = [present_issue(issue, field_labels) for issue in result.issues]
    warning_count = sum(item.severity == "warning" for item in displays)
    error_count = sum(item.severity == "error" for item in displays)
    info_count = sum(item.severity == "info" for item in displays)
    lines = [
        f"文件：{result.output_path.name}",
        f"处理提示 {info_count} 条；警告 {warning_count} 条；错误 {error_count} 条。",
    ]
    if not displays:
        lines.append("没有需要显示的问题。")
    for index, item in enumerate(displays, 1):
        lines.extend(
            (
                "",
                f"{index}. [{item.severity_label}] {item.title}",
                f"说明：{item.description}",
                f"位置：工作表 {item.source_sheet}，来源行 {item.source_row}，字段 {item.field}",
                f"建议：{item.action}",
            )
        )
    return "\n".join(lines)


def conversion_result_from_report(report_path: str | Path) -> ConversionResult:
    """Load an existing compatible report for local UI review without rewriting it."""

    path = Path(report_path).resolve()
    payload = json.loads(path.read_text(encoding="utf-8"))
    issues = [
        Issue(
            code=str(item.get("code") or ""),
            severity=str(item.get("severity") or "warning"),
            message=str(item.get("message") or ""),
            source_sheet=item.get("source_sheet"),
            source_row=item.get("source_row"),
            field=item.get("field"),
            source_value=item.get("source_value"),
            calculated_value=item.get("calculated_value"),
        )
        for item in payload.get("issues", [])
    ]
    source_name = str(payload.get("source_file") or "未知来源.xlsx")
    output_name = str(payload.get("output_file") or path.with_suffix(".xlsx").name)
    return ConversionResult(
        source_path=path.parent / source_name,
        output_path=path.parent / output_name,
        profile_id=str(payload.get("profile_id") or "unknown"),
        source_rows=int(payload.get("source_rows") or 0),
        output_rows=int(payload.get("output_rows") or 0),
        source_images=int(payload.get("source_images") or 0),
        output_images=int(payload.get("output_images") or 0),
        issues=issues,
        report_path=path,
    )

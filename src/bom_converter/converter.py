from __future__ import annotations

import json
from collections import Counter
from pathlib import Path

from .analyzer import analyze_workbook
from .confirmed_mapping import apply_confirmed_mappings, transform_confirmed_workbook
from .models import ColumnMappingConfig, ConversionResult, Issue, WorkbookAnalysis
from .template_writer import write_template_output
from .transform import transform_workbook
from .verifier import sha256_file, verify_output


def _write_report(output_path: Path, analysis: WorkbookAnalysis, issues: list[Issue], output_rows: int, output_images: int) -> Path:
    report_path = output_path.with_suffix(".report.json")
    severity = Counter(issue.severity for issue in issues)
    payload = {
        "offline_processing": True,
        "source_file": analysis.source_path.name,
        "output_file": output_path.name,
        "profile_id": analysis.profile_id,
        "profile_name": analysis.profile_name,
        "sheet": analysis.sheet_name,
        "header_rows": analysis.header_rows,
        "header_start_row": analysis.header_start_row,
        "header_end_row": analysis.header_end_row,
        "data_start_row": analysis.data_start_row,
        "header_confidence_reasons": analysis.header_region.confidence_reasons if analysis.header_region else [],
        "fingerprint": analysis.fingerprint,
        "source_rows": analysis.data_row_count,
        "output_rows": output_rows,
        "source_images": analysis.image_count,
        "output_images": output_images,
        "requires_review": analysis.requires_review,
        "ignored_nonempty_columns": analysis.ignored_nonempty_columns,
        "issue_summary": dict(severity),
        "issues": [issue.__dict__ for issue in issues],
        "mappings": [mapping.__dict__ for mapping in analysis.mappings],
        "privacy_note": "报告默认不保存完整零件名称、编号、供应商或图片内容。",
    }
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return report_path


def convert_file(source_path: str | Path, template_path: str | Path, output_dir: str | Path, mode: str = "quick") -> ConversionResult:
    source = Path(source_path).resolve()
    template = Path(template_path).resolve()
    if source == template:
        raise ValueError("来源文件不能与标准模板相同")
    source_hash = sha256_file(source)
    template_hash = sha256_file(template)
    analysis = analyze_workbook(source, mode=mode)
    if analysis.profile_id == "generic":
        raise ValueError("该文件未匹配公开演示格式；请在通用审核模式确认映射后再转换")
    return convert_analyzed_file(source, template, output_dir, analysis)


def convert_analyzed_file(
    source_path: str | Path,
    template_path: str | Path,
    output_dir: str | Path,
    analysis: WorkbookAnalysis,
    *,
    include_sheet_name: bool = True,
) -> ConversionResult:
    """Convert one already analyzed worksheet to its own standard workbook."""

    source = Path(source_path).resolve()
    template = Path(template_path).resolve()
    if source == template:
        raise ValueError("来源文件不能与标准模板相同")
    source_hash = sha256_file(source)
    template_hash = sha256_file(template)
    rows, issues = transform_workbook(source, analysis)
    output_stem = f"{source.stem}__{analysis.sheet_name}_" if include_sheet_name else source.stem
    output = write_template_output(template, source, output_dir, rows, analysis.profile_id, output_stem)
    output_rows, output_images, verification_issues = verify_output(output, template, len(rows), sum(len(row.images) for row in rows))
    issues.extend(verification_issues)
    if sha256_file(source) != source_hash:
        issues.append(Issue("SOURCE_MODIFIED", "error", "来源文件哈希发生变化"))
    if sha256_file(template) != template_hash:
        issues.append(Issue("TEMPLATE_MODIFIED", "error", "标准模板哈希发生变化"))
    report_path = _write_report(output, analysis, issues, output_rows, output_images)
    return ConversionResult(
        source, output, analysis.profile_id, len(rows), output_rows,
        sum(len(row.images) for row in rows), output_images, issues, report_path,
        analysis.sheet_name,
    )


def convert_many(sources: list[str | Path], template_path: str | Path, output_dir: str | Path, mode: str = "quick") -> list[ConversionResult]:
    return [convert_file(source, template_path, output_dir, mode=mode) for source in sources]


def convert_confirmed_file(
    source_path: str | Path,
    template_path: str | Path,
    output_dir: str | Path,
    analysis: WorkbookAnalysis,
    mappings: dict[int, str | None],
    units: dict[str, str] | None = None,
    column_configs: dict[int, ColumnMappingConfig] | None = None,
) -> ConversionResult:
    source = Path(source_path).resolve()
    template = Path(template_path).resolve()
    if source == template:
        raise ValueError("来源文件不能与标准模板相同")
    source_hash = sha256_file(source)
    template_hash = sha256_file(template)
    apply_confirmed_mappings(analysis, mappings)
    rows, issues = transform_confirmed_workbook(source, analysis, mappings, units, column_configs)
    output = write_template_output(
        template,
        source,
        output_dir,
        rows,
        "generic",
        f"{source.stem}__{analysis.sheet_name}_",
    )
    expected_images = sum(len(row.images) for row in rows)
    output_rows, output_images, verification_issues = verify_output(
        output, template, len(rows), expected_images
    )
    issues.extend(verification_issues)
    if sha256_file(source) != source_hash:
        issues.append(Issue("SOURCE_MODIFIED", "error", "来源文件哈希发生变化"))
    if sha256_file(template) != template_hash:
        issues.append(Issue("TEMPLATE_MODIFIED", "error", "标准模板哈希发生变化"))
    report_path = _write_report(output, analysis, issues, output_rows, output_images)
    return ConversionResult(
        source, output, "generic_confirmed", len(rows), output_rows,
        expected_images, output_images, issues, report_path, analysis.sheet_name,
    )

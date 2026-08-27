from __future__ import annotations

from collections import Counter
from pathlib import Path
import re

from .models import HeaderRegion, Issue, MappingDecision, SheetAnalysis, WorkbookAnalysis
from .header_region import build_header_region, detect_header_regions, extract_header_paths
from .mapping_memory import MappingMemory
from .name_rules import infer_name_rule_columns
from .profiles import GENERIC_SYNONYMS, PROFILES, PROFILE_BY_ID, SourceProfile, TARGET_HEADERS
from .text_utils import clean_text, column_letter, fingerprint, normalize_header
from .xlsx_reader import SheetData, XlsxReader


def extract_headers(sheet: SheetData, header_rows: tuple[int, ...], max_col: int | None = None) -> list[str]:
    paths = extract_header_paths(sheet, min(header_rows), max(header_rows), max_col)
    return ["/".join(path) if path else f"<未命名列 {column_letter(col)}>" for col, path in enumerate(paths, 1)]


def _profile_score(sheet: SheetData, profile: SourceProfile) -> float:
    headers = extract_headers(sheet, profile.header_rows, len(profile.columns))
    expected = [rule.source_header for rule in profile.columns]
    matches = sum(normalize_header(a) == normalize_header(b) for a, b in zip(headers, expected))
    width_penalty = min(abs(sheet.max_col - len(expected)) / max(len(expected), 1), 0.25)
    return max(matches / len(expected) - width_penalty, 0.0)


def detect_known_profile(reader: XlsxReader) -> tuple[SourceProfile | None, SheetData, float]:
    candidates: list[tuple[float, SourceProfile, SheetData]] = []
    for sheet_name in reader.sheet_names:
        sheet = reader.read_sheet(sheet_name)
        for profile in PROFILES:
            candidates.append((_profile_score(sheet, profile), profile, sheet))
    score, profile, sheet = max(candidates, key=lambda item: item[0])
    return (profile if score >= 0.82 else None), sheet, score


def detect_known_profile_for_sheet(sheet: SheetData) -> tuple[SourceProfile | None, float]:
    score, profile = max(((_profile_score(sheet, item), item) for item in PROFILES), key=lambda item: item[0])
    return (profile if score >= 0.82 else None), score


def _detect_generic_header(sheet: SheetData) -> int:
    scored: list[tuple[float, int]] = []
    for row in range(1, min(sheet.max_row, 30) + 1):
        values = [clean_text(sheet.get(row, col, merged=True)) for col in range(1, sheet.max_col + 1)]
        nonempty = [value for value in values if value]
        known = sum(normalize_header(value) in GENERIC_SYNONYMS for value in nonempty)
        status_hits = sum(value in {"已计算", "总行数", "在处理", "在处理行数", "Demo Line"} for value in nonempty)
        score = known * 3 + min(len(nonempty), 15) * 0.15 - status_hits * 5
        scored.append((score, row))
    return max(scored)[1]


def _generic_analysis(
    source: Path,
    reader: XlsxReader,
    sheet: SheetData,
    region: HeaderRegion,
    mapping_memory: MappingMemory | None,
    *,
    manually_selected: bool = False,
    candidate_sheet_count: int = 1,
) -> WorkbookAnalysis:
    images = reader.read_images(sheet.name)
    issues: list[Issue] = []
    headers = [column.logical_header for column in region.columns]
    workbook_fingerprint = fingerprint([sheet.name, *headers])
    saved_rule = mapping_memory.get_rule(workbook_fingerprint) if mapping_memory else None
    remembered = mapping_memory.recall(workbook_fingerprint) if mapping_memory else {}
    remembered_configs = mapping_memory.recall_configs(workbook_fingerprint) if mapping_memory else {}
    remembered_name_rule = mapping_memory.recall_name_rule(workbook_fingerprint) if mapping_memory and saved_rule else None
    ignored_columns = set(saved_rule.get("ignored_columns", [])) if saved_rule else set()
    mappings: list[MappingDecision] = []
    for column in region.columns:
        field = column.recommended_field
        process = column.recommended_process
        image_slot = None
        config = remembered_configs.get(column.source_col)
        if config:
            field = config.universal_field
            process = config.process_type
            image_slot = config.image_slot
            confidence = 1.0
            status = "mapped" if process != "ignore" else "ignored"
            note = "已应用本机保存的完整表头区域规则"
        elif column.source_col in remembered:
            field = remembered[column.source_col]
            confidence = 1.0
            status = "mapped"
            note = "已应用本机保存的字段映射；记忆文件不保存 BOM 行数据"
        else:
            confidence = column.confidence
            status = "review" if column.recommendation_conflict or not field else "mapped"
            if process == "ignore" and not field:
                status = "ignored"
            note = "；".join(column.confidence_reasons)
        mappings.append(MappingDecision(
            source_col=column.source_col,
            source_header=column.logical_header,
            sample_value=_sample_value(sheet, region.data_start_row, column.source_col),
            universal_field=field,
            target_header=TARGET_HEADERS.get(field or ""),
            confidence=confidence,
            status=status,
            note=note,
            column_letter=column.column_letter,
            header_path=list(column.header_path),
            parent_header=column.parent_header,
            child_header=column.child_header,
            source_unit=(config.unit if config else column.unit),
            process_type=process,
            image_slot=image_slot,
            level_group=(config.level_group if config else column.level_group),
            level_value=(config.level_value if config else column.level_value),
            default_value=(config.default_value if config else None),
            confidence_reasons=list(column.confidence_reasons),
        ))

    rows = _data_rows(sheet, region.data_start_row)
    unit_changed = False
    if saved_rule:
        if saved_rule.get("sheet_name") and saved_rule.get("sheet_name") != sheet.name:
            issues.append(Issue("SHEET_CHANGED", "warning", "已保存规则的工作表名称发生变化，需要检查", sheet.name))
        saved_start = saved_rule.get("header_start_row")
        saved_end = saved_rule.get("header_end_row")
        saved_data = saved_rule.get("data_start_row")
        region_changed = (
            (saved_start is not None and saved_start != region.header_start_row)
            or (saved_end is not None and saved_end != region.header_end_row)
            or (saved_data is not None and saved_data != region.data_start_row)
        )
        if region_changed:
            issues.append(Issue(
                "HEADER_REGION_CHANGED", "warning",
                f"已保存表头/数据区域为 {saved_start}-{saved_end}/{saved_data}，当前为 {region.header_start_row}-{region.header_end_row}/{region.data_start_row}",
                sheet.name,
            ))
        for mapping in mappings:
            saved_unit = saved_rule.get("units", {}).get(str(mapping.source_col)) or saved_rule.get("units", {}).get(mapping.universal_field or "")
            current_unit = mapping.source_unit or infer_sample_unit(mapping.universal_field, mapping.sample_value)
            if saved_unit and current_unit and saved_unit != current_unit:
                unit_changed = True
                issues.append(Issue(
                    "UNIT_CHANGED", "warning",
                    f"列 {mapping.column_letter} 的单位由已保存的 {saved_unit} 变为 {current_unit}，需要再次确认",
                    sheet.name, field=mapping.universal_field, source_value=current_unit,
                ))
        issues.append(Issue("SAVED_RULE_APPLIED", "info", "已应用本机表头区域和映射规则；规则库不保存 BOM 行数据", sheet.name))
    else:
        similar = mapping_memory.find_similar(headers) if mapping_memory else None
        if similar:
            issues.append(Issue(
                "SIMILAR_SAVED_FORMAT", "warning",
                f"与历史格式相似度 {similar.similarity:.0%}；新增列 {similar.added_headers or '无'}；缺少列 {similar.missing_headers or '无'}，请检查差异",
                sheet.name,
            ))
        else:
            issues.append(Issue("UNKNOWN_PROFILE", "warning", "未匹配公开演示格式，需要在通用审核模式确认表头区域和映射", sheet.name))
    if candidate_sheet_count > 1 and not manually_selected:
        issues.append(Issue(
            "MULTIPLE_HEADER_SHEETS", "warning",
            f"检测到 {candidate_sheet_count} 个可能包含 BOM 表头的工作表；已显示最高分候选，请人工选择工作表",
            sheet.name,
        ))
    if region.confidence < .72 and not saved_rule:
        issues.append(Issue(
            "LOW_HEADER_CONFIDENCE", "warning",
            "表头区域识别置信度较低：" + "；".join(region.confidence_reasons),
            sheet.name,
        ))
    requires_review = (
        not saved_rule
        or unit_changed
        or any(mapping.status == "review" for mapping in mappings)
        or any(issue.code in {"HEADER_REGION_CHANGED", "MULTIPLE_HEADER_SHEETS", "LOW_HEADER_CONFIDENCE"} for issue in issues)
    )
    name_rule = infer_name_rule_columns(mappings, remembered_configs, remembered_name_rule)
    return WorkbookAnalysis(
        source_path=source,
        profile_id="generic",
        profile_name="已记忆的通用格式" if saved_rule and not requires_review else "未确认的通用格式",
        sheet_name=sheet.name,
        header_rows=list(range(region.header_start_row, region.header_end_row + 1)),
        data_start_row=region.data_start_row,
        data_row_count=len(rows),
        image_count=len(images),
        fingerprint=workbook_fingerprint,
        profile_confidence=region.confidence,
        mappings=mappings,
        ignored_nonempty_columns=[m.source_header for m in mappings if m.status == "ignored" and m.sample_value is not None],
        requires_review=requires_review,
        issues=issues,
        header_region=region,
        name_rule=name_rule,
    )


def analyze_manual_workbook(
    path: str | Path,
    sheet_name: str,
    header_start_row: int,
    header_end_row: int,
    data_start_row: int,
    mapping_memory: MappingMemory | None = None,
) -> WorkbookAnalysis:
    source = Path(path).resolve()
    with XlsxReader(source) as reader:
        if sheet_name not in reader.sheet_names:
            raise ValueError(f"工作表不存在：{sheet_name}")
        sheet = reader.read_sheet(sheet_name)
        region = build_header_region(sheet, header_start_row, header_end_row, data_start_row)
        return _generic_analysis(source, reader, sheet, region, mapping_memory, manually_selected=True)


def _sample_value(sheet: SheetData, start_row: int, col: int) -> object:
    for row in range(start_row, min(sheet.max_row, start_row + 10) + 1):
        value = sheet.get(row, col)
        if clean_text(value):
            return value
    return None


def _data_rows(sheet: SheetData, start_row: int, mapped_cols: list[int] | None = None) -> list[int]:
    cols = mapped_cols or list(range(1, sheet.max_col + 1))
    rows: list[int] = []
    for row in range(start_row, sheet.max_row + 1):
        if any(clean_text(sheet.get(row, col)) for col in cols):
            rows.append(row)
    return rows


def infer_sample_unit(field: str | None, value: object) -> str | None:
    if field not in {"unit_weight", "total_weight", "dimensions", "thickness", "surface_area"}:
        return None
    text = clean_text(value)
    if not text:
        return None
    normalized = text.casefold().replace("²", "2").replace("^", "")
    candidates = {
        "unit_weight": ("mg", "kg", "g"),
        "total_weight": ("mg", "kg", "g"),
        "dimensions": ("mm", "cm", "m"),
        "thickness": ("mm", "cm", "m"),
        "surface_area": ("mm2", "cm2", "m2"),
    }[field]
    for unit in candidates:
        if re.search(rf"(?<![a-z]){re.escape(unit)}(?![a-z0-9])", normalized):
            return unit.replace("2", "²")
    return None


def analyze_workbook(
    path: str | Path,
    mode: str = "quick",
    profile_id: str | None = None,
    mapping_memory: MappingMemory | None = None,
) -> WorkbookAnalysis:
    source = Path(path).resolve()
    with XlsxReader(source) as reader:
        if profile_id:
            profile = PROFILE_BY_ID[profile_id]
            sheet = max((reader.read_sheet(name) for name in reader.sheet_names), key=lambda s: _profile_score(s, profile))
            score = _profile_score(sheet, profile)
        else:
            profile, sheet, score = detect_known_profile(reader)
        images = reader.read_images(sheet.name)
        issues: list[Issue] = []
        if profile is None:
            candidates = detect_header_regions(reader)
            if not candidates:
                raise ValueError("没有找到可供审核的表头区域")
            remembered_region = None
            if mapping_memory:
                for candidate in candidates:
                    candidate_headers = [column.logical_header for column in candidate.columns]
                    candidate_fingerprint = fingerprint([candidate.sheet_name, *candidate_headers])
                    if mapping_memory.get_rule(candidate_fingerprint):
                        remembered_region = candidate
                        break
            region = remembered_region or candidates[0]
            selected_sheet = reader.read_sheet(region.sheet_name)
            plausible_sheets = {item.sheet_name for item in candidates if item.confidence >= max(region.confidence - .12, .62)}
            return _generic_analysis(
                source, reader, selected_sheet, region, mapping_memory,
                candidate_sheet_count=1 if remembered_region else len(plausible_sheets),
            )
        selected_sheet_name = sheet.name
    return analyze_sheet(source, selected_sheet_name, mode=mode, mapping_memory=mapping_memory)


def analyze_sheet(
    path: str | Path,
    sheet_name: str,
    mode: str = "quick",
    mapping_memory: MappingMemory | None = None,
) -> WorkbookAnalysis:
    """Analyze one explicitly selected sheet without silently switching to another sheet."""

    source = Path(path).resolve()
    with XlsxReader(source) as reader:
        if sheet_name not in reader.sheet_names:
            raise ValueError(f"工作表不存在：{sheet_name}")
        sheet = reader.read_sheet(sheet_name)
        if sheet.max_row < 1 or sheet.max_col < 1:
            raise ValueError("工作表为空")
        profile, score = detect_known_profile_for_sheet(sheet)
        if profile is None:
            candidates = detect_header_regions(reader, sheet_names={sheet_name})
            if not candidates:
                raise ValueError("没有找到可供审核的表头区域")
            remembered = None
            if mapping_memory:
                for candidate in candidates:
                    headers = [column.logical_header for column in candidate.columns]
                    if mapping_memory.get_rule(fingerprint([sheet_name, *headers])):
                        remembered = candidate
                        break
            region = remembered or candidates[0]
            return _generic_analysis(
                source,
                reader,
                sheet,
                region,
                mapping_memory,
                manually_selected=True,
            )

        headers = extract_headers(sheet, profile.header_rows, len(profile.columns))
        mappings: list[MappingDecision] = []
        ignored_nonempty: list[str] = []
        mapped_data_cols: list[int] = []
        known_region = build_header_region(sheet, min(profile.header_rows), max(profile.header_rows), profile.data_start_row)
        region_by_col = {column.source_col: column for column in known_region.columns}
        for col, rule in enumerate(profile.columns, 1):
            sample = _sample_value(sheet, profile.data_start_row, col)
            target = TARGET_HEADERS.get(rule.field or "")
            if rule.field == "image":
                target = "Label Picture" if rule.image_slot == 7 else (f"Picture{rule.image_slot}" if rule.image_slot else None)
            region_column = region_by_col.get(col)
            mappings.append(MappingDecision(
                col,
                headers[col - 1],
                sample,
                rule.field,
                target,
                rule.confidence,
                rule.status,
                rule.note,
                column_letter(col),
                list(region_column.header_path) if region_column else [headers[col - 1]],
                source_unit=region_column.unit if region_column else None,
                process_type=("image" if rule.field == "image" else ("level_group" if rule.field == "level_marker" else "direct")),
                image_slot=rule.image_slot,
                level_group="bom_level" if rule.field == "level_marker" else None,
                level_value=(region_column.level_value if region_column else None),
            ))
            if rule.field and rule.field not in {"image", "level_marker"}:
                mapped_data_cols.append(col)
            if rule.status == "ignored" and sample is not None:
                ignored_nonempty.append(headers[col - 1])
        rows = _data_rows(sheet, profile.data_start_row, mapped_data_cols)
        requires_review = mode == "audit" or score < .95 or any(m.status == "review" or m.confidence < .8 for m in mappings)
        return WorkbookAnalysis(
            source,
            profile.profile_id,
            profile.name,
            sheet.name,
            list(profile.header_rows),
            profile.data_start_row,
            len(rows),
            len(reader.read_images(sheet.name)),
            fingerprint([sheet.name, *headers]),
            score,
            mappings,
            ignored_nonempty,
            requires_review,
            [],
            known_region,
        )


def analyze_workbook_sheets(
    path: str | Path,
    mode: str = "quick",
    mapping_memory: MappingMemory | None = None,
) -> list[SheetAnalysis]:
    """Enumerate and independently analyze every worksheet for UI selection."""

    source = Path(path).resolve()
    results: list[SheetAnalysis] = []
    with XlsxReader(source) as reader:
        sheet_metadata = [(name, reader.is_sheet_hidden(name)) for name in reader.sheet_names]
    for sheet_name, hidden in sheet_metadata:
        try:
            analysis = analyze_sheet(source, sheet_name, mode=mode, mapping_memory=mapping_memory)
            recognized = sum(bool(item.universal_field) for item in analysis.mappings)
            has_data = analysis.data_row_count > 0
            is_bom = has_data and recognized >= 2 and analysis.profile_confidence >= .58
            recommended = is_bom and not hidden and analysis.profile_confidence >= .65
            status = "隐藏工作表" if hidden else ("建议转换" if recommended else ("可能是 BOM，建议检查" if is_bom else "默认不选择"))
            results.append(SheetAnalysis(source, sheet_name, hidden, analysis, is_bom, recommended, status))
        except Exception as exc:
            results.append(SheetAnalysis(source, sheet_name, hidden, None, False, False, "无法识别", str(exc)))
    return results

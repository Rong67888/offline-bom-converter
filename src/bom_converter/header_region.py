from __future__ import annotations

from dataclasses import replace
import re
from typing import Iterable

from .models import HeaderColumn, HeaderRegion, Issue
from .text_utils import clean_text, column_letter, normalize_header, split_cell_reference
from .xlsx_reader import SheetData, XlsxReader


LEVEL_MARKERS = {"1", "y", "yes", "√", "✓", "●"}
TITLE_WORDS = {
    "项目编号", "projectno", "quotationbom", "designbom", "说明", "隐私说明",
    "候选表头行", "数据起始行", "预期处理", "simulationbom",
}


def _aliases() -> dict[str, tuple[str, ...]]:
    return {
        "sequence": ("序号", "item", "no"),
        "category": ("分类", "级别", "class"),
        "part_name": (
            "名称", "零件名称", "零件部件名称", "partname", "partcomponentname",
            "subcomponentpartname", "partnameandbreakdown", "总成零件名称",
            "子零件部件零件名称",
        ),
        "part_number": (
            "零件号", "零件分部件零件号", "partnumber", "partcomponentnumber",
            "subcomponentpartnumber", "partnumberlessfinish", "customerno",
            "总成零件号", "子零件部件零件号",
        ),
        "quantity": ("数量", "件数", "qty", "qtyveh", "quantity", "quantitypcs", "pcset", "件套", "单件数量"),
        "unit_weight": ("单件重量", "masscomp", "masskg", "netweight", "净重"),
        "total_weight": ("总重", "massasm", "grossweight", "毛重"),
        "dimensions": ("尺寸", "尺寸范围", "boundarydimension", "overalldimensions", "外形尺寸", "size"),
        "thickness": ("厚度", "料厚", "thickness"),
        "surface_area": ("表面积", "surfacearea"),
        "material_raw": ("材质", "材料", "材料名称", "materialname", "materialnameandgradecode"),
        "material_spec": ("材料标准", "materialspecification"),
        "manufacturer": ("供应商", "制造商", "manufacturer", "supplier"),
        "location": ("产地", "location", "source"),
        "production_process": ("工艺", "生产工艺", "manufacturingprocess", "requiredmanufacturingprocess"),
        "image": ("图片", "图", "image", "photo", "pic", "figure", "visiblepic", "partfigure", "partvisiblepic", "photoofcomponent"),
        "remark": ("备注", "补充说明", "remark", "remarks"),
        "code": ("编号", "编码", "code", "qad", "noqad", "oldsystemno"),
        "level": ("层级", "装配级别", "componentasmlevel", "level"),
    }


FIELD_ALIASES = _aliases()


def is_level_marker(value: object) -> bool:
    text = clean_text(value)
    return bool(text) and text.casefold() in LEVEL_MARKERS


def level_number_from_header(value: object) -> int | None:
    """Return the semantic level from L1..L99/Level 1 labels, not column position."""

    text = clean_text(value)
    if not text:
        return None
    normalized = normalize_header(text)
    match = re.fullmatch(r"(?:l|level|层级)?0*(\d{1,2})", normalized)
    if not match:
        return None
    level = int(match.group(1))
    return level if level > 0 else None


def resolve_level_markers(
    values: Iterable[tuple[int, object]],
    *,
    sheet_name: str,
    source_row: int,
) -> tuple[int | None, Issue | None]:
    marked = [level for level, value in values if is_level_marker(value)]
    if len(marked) == 1:
        return marked[0], None
    code = "LEVEL_MARKER_MULTIPLE" if marked else "LEVEL_MARKER_MISSING"
    message = (
        f"层级列组检测到多个有效标记 {marked}，未猜测 Level"
        if marked
        else "层级列组没有有效标记，未猜测 Level"
    )
    return None, Issue(code, "warning", message, sheet_name, source_row, "level")


def _split_alias_text(value: str) -> list[str]:
    return [part for part in re.split(r"[\r\n]+", value) if clean_text(part)]


def _unit_from_path(path: list[str]) -> str | None:
    text = " ".join(path).casefold().replace("²", "2").replace("㎏", "kg")
    for pattern, unit in (
        (r"(?<![a-z])mm2(?![a-z0-9])", "mm²"),
        (r"(?<![a-z])cm2(?![a-z0-9])", "cm²"),
        (r"(?<![a-z])m2(?![a-z0-9])", "m²"),
        (r"(?<![a-z])mg(?![a-z])", "mg"),
        (r"(?<![a-z])kg(?![a-z])", "kg"),
        (r"(?<![a-z])mm(?![a-z])", "mm"),
        (r"(?<![a-z])cm(?![a-z])", "cm"),
        (r"(?<![a-z])g(?![a-z])", "g"),
        (r"(?:^|[\s（(])m(?:$|[\s）)])", "m"),
    ):
        if re.search(pattern, text):
            return unit
    return None


def _field_for_text(text: str) -> str | None:
    normalized = normalize_header(text)
    if not normalized:
        return None
    matches: list[tuple[int, str]] = []
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            normalized_alias = normalize_header(alias)
            if normalized == normalized_alias or (len(normalized_alias) >= 4 and normalized_alias in normalized):
                matches.append((len(normalized_alias), field))
    return max(matches)[1] if matches else None


def _recommend(path: list[str]) -> tuple[str | None, str, float, list[str], bool, bool, str | None, int | None]:
    reasons: list[str] = []
    normalized_parts = [normalize_header(part) for value in path for part in _split_alias_text(value)]
    child = normalize_header(path[-1]) if path else ""
    parent = normalize_header(path[-2]) if len(path) > 1 else ""
    combined = "".join(normalized_parts)

    if child in {"x", "y", "z"} and any(token in combined for token in ("size", "dimension", "尺寸")):
        field = {"x": "length", "y": "width", "z": "height"}[child]
        reasons.append(f"父表头表示尺寸，子表头 {child.upper()} 映射为长/宽/高")
        return field, "dimension_axis", .98, reasons, False, False, None, None

    level_parent = any(token in combined for token in ("componentasmlevel", "装配级别"))
    if level_parent and child.isdigit() and 1 <= int(child) <= 13:
        level_value = int(child)
        reasons.append(f"识别为层级列组的第 {level_value} 级")
        return "level", "level_group", .99, reasons, False, False, "component_asm_level", level_value

    standalone_level = level_number_from_header(path[-1]) if path else None
    if standalone_level is not None and standalone_level <= 13:
        reasons.append(f"表头 L{standalone_level} 识别为层级列组第 {standalone_level} 级")
        return "level", "level_group", .98, reasons, False, False, "bom_level", standalone_level

    piece_fields: list[str] = []
    for value in path:
        aliases = _split_alias_text(value)
        for alias in aliases:
            field = _field_for_text(alias)
            if field:
                piece_fields.append(field)
    unique_fields = list(dict.fromkeys(piece_fields))
    bilingual = len(piece_fields) >= 2
    if len(unique_fields) > 1:
        reasons.append("同一列的多行/中英文表头推荐结果冲突，必须人工确认")
        return None, "direct", .25, reasons, bilingual, True, None, None
    if unique_fields:
        field = unique_fields[0]
        process = "dimension_string" if field == "dimensions" else ("image" if field == "image" else "direct")
        if bilingual:
            reasons.append("连续多行或换行中的别名推荐为同一通用字段")
            confidence = .96
        else:
            reasons.append("表头别名与通用字段词典匹配")
            confidence = .86
        return field, process, confidence, reasons, bilingual, False, None, None
    reasons.append("未找到可靠的通用字段别名，需要人工选择或忽略")
    return None, "ignore", 0.0, reasons, bilingual, False, None, None


def extract_header_paths(sheet: SheetData, start_row: int, end_row: int, max_col: int | None = None) -> list[list[str]]:
    result: list[list[str]] = []
    for col in range(1, (max_col or sheet.max_col) + 1):
        pieces: list[str] = []
        for row in range(start_row, end_row + 1):
            text = clean_text(sheet.get(row, col, merged=True))
            if text and (not pieces or normalize_header(text) != normalize_header(pieces[-1])):
                pieces.append(text)
        result.append(pieces)
    return result


def merged_structure_summary(sheet: SheetData, start_row: int, end_row: int) -> list[str]:
    result: list[str] = []
    for ref in sheet.merged_ranges:
        start, end = ref.split(":") if ":" in ref else (ref, ref)
        r1, _ = split_cell_reference(start)
        r2, _ = split_cell_reference(end)
        if r1 <= end_row and r2 >= start_row:
            result.append(ref)
    return sorted(result)


def build_header_region(
    sheet: SheetData,
    header_start_row: int,
    header_end_row: int,
    data_start_row: int | None = None,
) -> HeaderRegion:
    if not 1 <= header_start_row <= header_end_row <= sheet.max_row:
        raise ValueError("表头开始行和结束行必须位于工作表有效范围内")
    actual_data_start = data_start_row or header_end_row + 1
    if actual_data_start <= header_end_row:
        raise ValueError("数据开始行必须大于表头结束行")
    paths = extract_header_paths(sheet, header_start_row, header_end_row)
    columns: list[HeaderColumn] = []
    for col, path in enumerate(paths, 1):
        logical = " / ".join(path) if path else f"<未命名列 {column_letter(col)}>"
        field, process, confidence, reasons, bilingual, conflict, group, level_value = _recommend(path)
        columns.append(HeaderColumn(
            source_col=col,
            column_letter=column_letter(col),
            header_path=path,
            logical_header=logical,
            parent_header=path[0] if len(path) > 1 else None,
            child_header=path[-1] if len(path) > 1 else None,
            unit=_unit_from_path(path),
            recommended_field=field,
            recommended_process=process,
            confidence=confidence,
            confidence_reasons=reasons,
            bilingual_alias=bilingual,
            recommendation_conflict=conflict,
            level_group=group,
            level_value=level_value,
        ))

    nonempty = [column for column in columns if column.header_path]
    recognized = [column for column in nonempty if column.recommended_field]
    conflicts = [column for column in columns if column.recommendation_conflict]
    level_columns = [column for column in columns if column.recommended_process == "level_group"]
    dimension_axes = [column for column in columns if column.recommended_process == "dimension_axis"]
    bilingual_matches = [column for column in columns if column.bilingual_alias and column.recommended_field]
    density = len(nonempty) / max(len(columns), 1)
    recognized_ratio = len(recognized) / max(len(nonempty), 1)
    confidence = .08 + .42 * recognized_ratio + .12 * min(density, 1.0)
    reasons = [f"{len(recognized)}/{len(nonempty) or 0} 个非空列获得字段推荐"]
    if len(level_columns) >= 2:
        confidence += .18
        reasons.append(f"识别到 {len(level_columns)} 列连续层级组")
    if len(dimension_axes) == 3:
        confidence += .16
        reasons.append("识别到 Size/Dimension 下完整 X、Y、Z 子列")
    if len(bilingual_matches) >= 2:
        confidence += .12
        reasons.append("多列连续中英文别名结果一致")
    if conflicts:
        confidence -= .18
        reasons.append(f"{len(conflicts)} 列存在中英文推荐冲突")
    if any(normalize_header(value) in TITLE_WORDS for path in paths for value in path):
        confidence -= .25
        reasons.append("区域含项目标题/说明词，降低置信度")
    return HeaderRegion(
        sheet.name,
        header_start_row,
        header_end_row,
        actual_data_start,
        columns,
        max(0.0, min(confidence, 1.0)),
        reasons,
        merged_structure_summary(sheet, header_start_row, header_end_row),
    )


def _candidate_data_evidence(sheet: SheetData, region: HeaderRegion) -> tuple[float, list[str]]:
    rows = range(region.data_start_row, min(sheet.max_row, region.data_start_row + 4) + 1)
    populated_rows = 0
    for row in rows:
        populated = sum(bool(clean_text(sheet.get(row, col))) for col in range(1, sheet.max_col + 1))
        if populated >= 2:
            populated_rows += 1
    if populated_rows:
        return min(.04 * populated_rows, .12), [f"表头后检测到 {populated_rows} 行数据证据"]
    return -.12, ["表头后缺少连续数据证据"]


def _row_structure_adjustment(sheet: SheetData, row: int) -> tuple[float, str | None]:
    raw_values = [sheet.get(row, col) for col in range(1, sheet.max_col + 1) if clean_text(sheet.get(row, col))]
    covering_merges: list[tuple[int, int, int, int]] = []
    for ref in sheet.merged_ranges:
        start, end = ref.split(":") if ":" in ref else (ref, ref)
        r1, c1 = split_cell_reference(start)
        r2, c2 = split_cell_reference(end)
        if r1 <= row <= r2:
            covering_merges.append((r1, c1, r2, c2))
    if not raw_values:
        if covering_merges:
            return 0.0, None
        return -.16, f"第 {row} 行为空行"
    if len(raw_values) == 1 and any(c1 == 1 and c2 >= max(sheet.max_col - 1, 1) for _, c1, _, c2 in covering_merges):
        return -.22, f"第 {row} 行是跨整表标题/说明"
    header_hits = sum(bool(_field_for_text(clean_text(value))) for value in raw_values)
    child_tokens = all(normalize_header(value) in {"x", "y", "z"} or normalize_header(value).isdigit() for value in raw_values)
    if header_hits >= max(1, len(raw_values) // 3):
        return .03, None
    if child_tokens and covering_merges:
        return .03, None
    numeric_or_identifier = sum(
        isinstance(value, (int, float))
        or normalize_header(value).startswith("sim")
        or "虚构" in clean_text(value)
        for value in raw_values
    )
    if numeric_or_identifier >= max(1, len(raw_values) // 2):
        return -.20, f"第 {row} 行更像数据行"
    return -.03, None


def _merge_boundary_adjustment(sheet: SheetData, start_row: int, end_row: int) -> tuple[float, list[str]]:
    adjustment = 0.0
    reasons: list[str] = []
    for ref in sheet.merged_ranges:
        start, end = ref.split(":") if ":" in ref else (ref, ref)
        r1, _ = split_cell_reference(start)
        r2, _ = split_cell_reference(end)
        if r1 < start_row <= r2:
            adjustment -= .12
            reasons.append(f"候选开始行切入合并区域 {ref}")
            break
    for ref in sheet.merged_ranges:
        start, end = ref.split(":") if ":" in ref else (ref, ref)
        r1, _ = split_cell_reference(start)
        r2, _ = split_cell_reference(end)
        if r1 <= end_row < r2:
            adjustment -= .12
            reasons.append(f"候选结束行截断合并区域 {ref}")
            break
    return adjustment, reasons


def detect_header_regions(
    reader: XlsxReader,
    max_scan_row: int = 30,
    max_height: int = 5,
    sheet_names: set[str] | None = None,
) -> list[HeaderRegion]:
    candidates: list[HeaderRegion] = []
    for sheet_name in reader.sheet_names:
        if sheet_names is not None and sheet_name not in sheet_names:
            continue
        sheet = reader.read_sheet(sheet_name)
        scan_end = min(sheet.max_row, max_scan_row)
        for start in range(1, scan_end + 1):
            for end in range(start, min(scan_end, start + max_height - 1) + 1):
                if end >= sheet.max_row:
                    continue
                region = build_header_region(sheet, start, end, end + 1)
                evidence, reasons = _candidate_data_evidence(sheet, region)
                row_adjustments = [_row_structure_adjustment(sheet, row) for row in range(start, end + 1)]
                row_score = sum(item[0] for item in row_adjustments)
                row_reasons = [item[1] for item in row_adjustments if item[1]]
                boundary_score, boundary_reasons = _merge_boundary_adjustment(sheet, start, end)
                height_bonus = .02 * (end - start) if any(
                    column.recommended_process in {"level_group", "dimension_axis"} or column.bilingual_alias
                    for column in region.columns
                ) else 0.0
                candidates.append(replace(
                    region,
                    confidence=max(0.0, min(region.confidence + evidence + height_bonus + row_score + boundary_score, 1.0)),
                    confidence_reasons=region.confidence_reasons + reasons + row_reasons + boundary_reasons,
                ))
    return sorted(
        candidates,
        key=lambda item: (
            item.confidence,
            -(item.header_end_row - item.header_start_row),
            item.header_start_row,
        ),
        reverse=True,
    )

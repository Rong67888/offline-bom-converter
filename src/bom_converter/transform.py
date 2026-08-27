from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Any

from .header_region import level_number_from_header, resolve_level_markers
from .models import BomRow, Issue, WorkbookAnalysis
from .name_rules import generate_name
from .profiles import PROFILE_BY_ID, ColumnRule, SourceProfile
from .text_utils import as_number, clean_text, combine_unique, display_number
from .xlsx_reader import XlsxReader


def _unit_from_text(value: Any, units: tuple[str, ...]) -> str | None:
    if not isinstance(value, str):
        return None
    normalized = value.casefold().replace("²", "2").replace("^", "")
    for unit in sorted(units, key=len, reverse=True):
        if re.search(rf"(?<![a-z]){re.escape(unit.casefold().replace('²', '2'))}(?![a-z])", normalized):
            return unit
    return None


def convert_weight(value: Any, default_unit: str = "kg") -> float | None:
    number = as_number(value)
    if number is None:
        return None
    unit = _unit_from_text(value, ("mg", "kg", "g")) or default_unit
    factors = {"kg": 1.0, "g": 0.001, "mg": 0.000001}
    return number * factors[unit]


def convert_length(value: Any, default_unit: str = "mm") -> float | None:
    number = as_number(value)
    if number is None:
        return None
    unit = _unit_from_text(value, ("mm", "cm", "m")) or default_unit
    factors = {"mm": 1.0, "cm": 10.0, "m": 1000.0}
    return number * factors[unit]


def convert_area(value: Any, default_unit: str = "m²") -> float | None:
    number = as_number(value)
    if number is None:
        return None
    unit = _unit_from_text(value, ("mm²", "cm²", "m²")) or default_unit
    factors = {"m²": 1.0, "cm²": 0.0001, "mm²": 0.000001}
    return number * factors[unit]


def parse_dimensions(value: Any, default_unit: str = "mm") -> tuple[dict[str, float], str | None]:
    text = clean_text(value)
    if not text:
        return {}, None
    normalized = text.replace("×", "*").replace("x", "*").replace("X", "*")
    unit = _unit_from_text(normalized, ("mm", "cm", "m")) or default_unit
    labeled = {match.group(1).upper(): float(match.group(2)) for match in re.finditer(r"\b([LWH])\s*[:=]?\s*(\d+(?:\.\d+)?)", normalized, re.I)}
    if labeled:
        factors = {"mm": 1.0, "cm": 10.0, "m": 1000.0}
        mapped = {"L": "length", "W": "width", "H": "height"}
        return {mapped[key]: number * factors[unit] for key, number in labeled.items()}, None
    cleaned = re.sub(r"(?i)(mm|cm|m|毫米|厘米|米)", "", normalized)
    if not re.fullmatch(r"\s*\d+(?:\.\d+)?(?:\s*\*\s*\d+(?:\.\d+)?){0,}\s*", cleaned):
        return {}, "尺寸包含无法解释的说明文字"
    numbers = [float(item) for item in re.findall(r"\d+(?:\.\d+)?", cleaned)]
    factor = {"mm": 1.0, "cm": 10.0, "m": 1000.0}[unit]
    if len(numbers) == 3:
        return {key: number * factor for key, number in zip(("length", "width", "height"), numbers)}, None
    if len(numbers) == 1:
        return {"length": numbers[0] * factor}, None
    return {}, f"尺寸包含 {len(numbers)} 段数值，阶段 1 不自动猜测"


def classify_material(value: Any) -> tuple[str | None, str | None]:
    text = clean_text(value)
    if not text:
        return None, None
    key = text.casefold()
    if any(token in key for token in ("adc", "aluminum", "铝")):
        category = "Aluminum"
    elif any(token in key for token in ("steel", "钢", "dc52", "铁")):
        category = "Steel"
    elif any(token in key for token in ("glass", "玻璃")):
        category = "Glass"
    elif any(token in key for token in ("plastic", "塑料", "abs", "pbt", "pmma", "pp-", "pc+")):
        category = "Plastics"
    elif any(token in key for token in ("electronic", "电子")):
        category = "Other"
    else:
        category = None
    return category, text


def _mapped_data_columns(profile: SourceProfile) -> list[int]:
    return [index for index, rule in enumerate(profile.columns, 1) if rule.field and rule.field not in {"image", "level_marker"}]


def _row_has_data(sheet: object, row: int, profile: SourceProfile) -> bool:
    return any(clean_text(sheet.get(row, col)) for col in _mapped_data_columns(profile))


def _add_value(bucket: dict[str, Any], field: str, value: Any) -> None:
    if field in {"remark_piece", "torque"}:
        bucket.setdefault(field, []).append(value)
    elif field == "process_raw" and bucket.get(field):
        bucket[field] = combine_unique([bucket[field], value])
    elif field not in bucket or bucket[field] is None:
        bucket[field] = value


def transform_workbook(path: str | Path, analysis: WorkbookAnalysis) -> tuple[list[BomRow], list[Issue]]:
    if analysis.profile_id not in PROFILE_BY_ID:
        raise ValueError("未知格式只能分析，必须在通用审核模式确认映射后再转换")
    profile = PROFILE_BY_ID[analysis.profile_id]
    issues: list[Issue] = list(analysis.issues)
    result: list[BomRow] = []
    with XlsxReader(path) as reader:
        sheet = reader.read_sheet(analysis.sheet_name)
        images_by_row: dict[int, list] = defaultdict(list)
        for image in reader.read_images(analysis.sheet_name):
            images_by_row[image.source_row].append(image)
        for source_row in range(profile.data_start_row, sheet.max_row + 1):
            if not _row_has_data(sheet, source_row, profile):
                continue
            raw: dict[str, Any] = {}
            unused: dict[str, Any] = {}
            for col, rule in enumerate(profile.columns, 1):
                value = sheet.get(source_row, col)
                if rule.field in {None, "image", "level_marker"}:
                    if clean_text(value) and rule.field is None:
                        unused[rule.source_header] = value
                    continue
                if rule.status in {"internal", "ignored", "review"}:
                    if clean_text(value):
                        unused[rule.source_header] = value
                _add_value(raw, rule.field, value)
            values: dict[str, Any] = {}
            for field in ("category", "sequence", "part_number", "level", "electronics_spec", "electronics_silk", "electronics_package",
                          "pin_number", "pcb_side", "electronics_type", "surface_treatment", "manufacturer", "location", "code", "vpc"):
                if field in raw:
                    values[field] = clean_text(raw[field])
            if not values.get("part_number"):
                values["part_number"] = clean_text(raw.get("part_number_fallback"))
                if values["part_number"]:
                    issues.append(Issue("PART_NUMBER_FALLBACK", "warning", "正式零件号为空，使用已确认的备用零件号", sheet.name, source_row, "part_number", raw.get("part_number_fallback")))
            name_result = generate_name(
                None,
                original=raw.get("part_name"),
                standard=raw.get("standard_name"),
                gb=raw.get("standard_name_gb"),
                spec=raw.get("spec"),
            )
            values["part_name"] = name_result.final_name
            if name_result.final_name and name_result.final_name != name_result.original_name:
                issues.append(Issue("GENERATED_NAME", "info", "名称由标准件名称、GB 名称和规格组合生成", sheet.name, source_row, "part_name", calculated_value=name_result.final_name))
            if profile.level_columns:
                semantic_columns = []
                for col in profile.level_columns:
                    header = profile.columns[col - 1].source_header if 1 <= col <= len(profile.columns) else None
                    level_number = level_number_from_header(header)
                    if level_number is not None:
                        semantic_columns.append((level_number, sheet.get(source_row, col)))
                level, level_issue = resolve_level_markers(
                    semantic_columns,
                    sheet_name=sheet.name,
                    source_row=source_row,
                )
                values["level"] = level
                if level_issue:
                    issues.append(level_issue)
            else:
                level = as_number(raw.get("level"))
                values["level"] = display_number(level) if level is not None else None
            quantity = as_number(raw.get("quantity"))
            values["quantity"] = display_number(quantity)
            unit_weight = convert_weight(raw.get("unit_weight"), profile.default_units.get("unit_weight", "kg"))
            total_weight = convert_weight(raw.get("total_weight"), profile.default_units.get("total_weight", "kg"))
            values["unit_weight"] = display_number(unit_weight)
            values["total_weight"] = display_number(total_weight)
            if unit_weight is not None and quantity is not None and total_weight is not None:
                calculated = unit_weight * quantity
                tolerance = max(abs(total_weight) * .02, .001)
                if abs(calculated - total_weight) > tolerance:
                    issues.append(Issue("WEIGHT_MISMATCH", "warning", "单件重量×数量与总重不一致，已保留来源值", sheet.name, source_row, "total_weight", total_weight, calculated))
            dimensions, dim_error = parse_dimensions(raw.get("dimensions"), profile.default_units.get("dimensions", "mm"))
            values.update({key: display_number(number) for key, number in dimensions.items()})
            if dim_error:
                issues.append(Issue("DIMENSION_AMBIGUOUS", "warning", dim_error, sheet.name, source_row, "dimensions", raw.get("dimensions")))
            thickness = convert_length(raw.get("thickness"), profile.default_units.get("thickness", "mm"))
            values["thickness"] = display_number(thickness)
            area = convert_area(raw.get("surface_area"), profile.default_units.get("surface_area", "m²"))
            values["surface_area"] = display_number(area)
            material_type, material_spec = classify_material(raw.get("material_raw"))
            values["material_type"] = material_type
            values["material_spec"] = material_spec
            process = clean_text(raw.get("process_raw"))
            if process:
                if "装配" in process or "assembly" in process.casefold():
                    values["assembly_process"] = process
                else:
                    values["production_process"] = process
            if clean_text(raw.get("production_process")):
                values["production_process"] = clean_text(raw.get("production_process"))
            remarks = [clean_text(raw.get("remark"))]
            remarks.extend(clean_text(item) for item in raw.get("remark_piece", []))
            remarks = [item for item in remarks if item]
            values["remark"] = "\n".join(dict.fromkeys(remarks)) if remarks else None
            row_images = sorted(images_by_row.get(source_row, []), key=lambda item: (item.source_col, item.order))
            if len(row_images) > 7:
                issues.append(Issue("IMAGE_OVERFLOW", "error", f"该行有 {len(row_images)} 张图片，目标模板最多 7 张", sheet.name, source_row, "images"))
            result.append(BomRow(sheet.name, source_row, values, row_images, unused, remarks))
    return result, issues


def image_slot_for(profile: SourceProfile | None, source_col: int, used_slots: set[int]) -> int | None:
    if profile is not None and 1 <= source_col <= len(profile.columns):
        configured = profile.columns[source_col - 1].image_slot
        if configured and configured not in used_slots:
            return configured
    for slot in range(1, 8):
        if slot not in used_slots:
            return slot
    return None

from __future__ import annotations

from collections import defaultdict
from dataclasses import replace
from pathlib import Path
from typing import Any

from .header_region import resolve_level_markers
from .models import BomRow, ColumnMappingConfig, Issue, WorkbookAnalysis
from .name_rules import generate_name_from_columns
from .profiles import TARGET_HEADERS
from .text_utils import as_number, clean_text, display_number
from .transform import (
    classify_material,
    convert_area,
    convert_length,
    convert_weight,
    parse_dimensions,
)
from .xlsx_reader import XlsxReader


INTERMEDIATE_FIELDS = {
    "dimensions": "尺寸（自动拆分到长/宽/高）",
    "material_raw": "原始材料（自动分类并保留牌号）",
    "standard_name": "标准件名称（用于名称组合）",
    "standard_name_gb": "GB 名称（用于名称组合）",
    "spec": "规格（用于名称组合）",
    "english_name": "英文名称（当前仅保留审核信息）",
}

FIELD_LABELS = {
    **{field: header.replace("\n", " / ") for field, header in TARGET_HEADERS.items()},
    **INTERMEDIATE_FIELDS,
}

PROCESS_LABELS = {
    "direct": "直接复制",
    "ignore": "忽略",
    "name_component": "名称组合来源",
    "dimension_string": "尺寸字符串拆分",
    "dimension_axis": "长/宽/高",
    "level_group": "层级列组",
    "image": "图片",
    "fallback_part_number": "备用零件号",
    "fixed_default": "固定默认值",
}


def infer_confirmed_units(analysis: WorkbookAnalysis, mappings: dict[int, str | None]) -> dict[str, str]:
    from .analyzer import infer_sample_unit

    result: dict[str, str] = {}
    by_column = {mapping.source_col: mapping for mapping in analysis.mappings}
    for column, field in mappings.items():
        if not field:
            continue
        unit = infer_sample_unit(field, by_column[column].sample_value)
        if unit:
            result[field] = unit
    return result


def apply_confirmed_mappings(analysis: WorkbookAnalysis, mappings: dict[int, str | None]) -> None:
    for decision in analysis.mappings:
        field = mappings.get(decision.source_col)
        decision.universal_field = field
        decision.target_header = TARGET_HEADERS.get(field or "")
        decision.confidence = 1.0
        decision.status = "mapped" if field else "ignored"
        decision.note = "用户已在通用审核模式确认"
    analysis.requires_review = False


def configs_from_analysis(
    analysis: WorkbookAnalysis,
    mappings: dict[int, str | None] | None = None,
    units: dict[str, str] | None = None,
) -> dict[int, ColumnMappingConfig]:
    mappings = mappings or {decision.source_col: decision.universal_field for decision in analysis.mappings}
    units = units or {}
    result: dict[int, ColumnMappingConfig] = {}
    for decision in analysis.mappings:
        field = mappings.get(decision.source_col)
        process = decision.process_type
        if field is not None and process == "ignore":
            process = "direct"
        if field is None and process not in {"image", "level_group", "fixed_default"}:
            process = "ignore"
        result[decision.source_col] = ColumnMappingConfig(
            source_col=decision.source_col,
            process_type=process,
            universal_field=field,
            unit=decision.source_unit or units.get(field or ""),
            image_slot=decision.image_slot,
            level_group=decision.level_group,
            level_value=decision.level_value,
            default_value=decision.default_value,
        )
    return result


def _first(raw: dict[str, list[Any]], field: str) -> Any:
    for value in raw.get(field, []):
        if clean_text(value):
            return value
    return None


def transform_confirmed_workbook(
    path: str | Path,
    analysis: WorkbookAnalysis,
    mappings: dict[int, str | None],
    units: dict[str, str] | None = None,
    column_configs: dict[int, ColumnMappingConfig] | None = None,
) -> tuple[list[BomRow], list[Issue]]:
    units = units or {}
    column_configs = column_configs or configs_from_analysis(analysis, mappings, units)
    for config in column_configs.values():
        if config.universal_field and config.unit:
            units.setdefault(config.universal_field, config.unit)
    issues: list[Issue] = list(analysis.issues)
    result: list[BomRow] = []
    mapped_columns = [
        column for column, config in column_configs.items()
        if config.process_type not in {"ignore", "image", "fixed_default"}
    ]
    if not mapped_columns and not any(config.process_type == "fixed_default" for config in column_configs.values()):
        raise ValueError("至少需要确认一个输出字段，不能把所有列都设为忽略")

    with XlsxReader(path) as reader:
        sheet = reader.read_sheet(analysis.sheet_name)
        images_by_row: dict[int, list] = defaultdict(list)
        for image in reader.read_images(analysis.sheet_name):
            images_by_row[image.source_row].append(image)

        for source_row in range(analysis.data_start_row, sheet.max_row + 1):
            if not any(clean_text(sheet.get(source_row, column)) for column in mapped_columns):
                continue
            raw: dict[str, list[Any]] = defaultdict(list)
            unused: dict[str, Any] = {}
            for decision in analysis.mappings:
                value = sheet.get(source_row, decision.source_col)
                config = column_configs.get(decision.source_col, ColumnMappingConfig(decision.source_col, "ignore"))
                field = config.universal_field
                if config.process_type in {"direct", "name_component", "dimension_string", "dimension_axis"} and field:
                    raw[field].append(value)
                elif config.process_type == "fallback_part_number":
                    raw["part_number_fallback"].append(value)
                elif config.process_type == "fixed_default" and field:
                    raw[field].append(config.default_value)
                elif clean_text(value):
                    unused[decision.source_header] = value

            values: dict[str, Any] = {}
            direct_text_fields = (
                "category", "part_number", "level", "electronics_spec", "electronics_silk",
                "electronics_package", "pin_number", "pcb_side", "electronics_type",
                "production_process", "assembly_process", "surface_treatment", "manufacturer",
                "location", "remark", "code", "vpc",
            )
            for field in direct_text_fields:
                value = _first(raw, field)
                if value is not None:
                    values[field] = clean_text(value)

            if not values.get("part_number"):
                fallback = _first(raw, "part_number_fallback")
                if clean_text(fallback):
                    values["part_number"] = clean_text(fallback)
                    issues.append(Issue(
                        "PART_NUMBER_FALLBACK", "warning",
                        "正式零件号为空，使用已确认的备用零件号",
                        sheet.name, source_row, "part_number", fallback,
                    ))

            level_groups: dict[str, list[ColumnMappingConfig]] = defaultdict(list)
            for config in column_configs.values():
                if config.process_type == "level_group" and config.level_group and config.level_value:
                    level_groups[config.level_group].append(config)
            if level_groups:
                if len(level_groups) > 1:
                    issues.append(Issue(
                        "MULTIPLE_LEVEL_GROUPS", "warning",
                        "当前行存在多个层级列组，仅使用第一个已配置组",
                        sheet.name, source_row, "level",
                    ))
                group = next(iter(level_groups.values()))
                level, level_issue = resolve_level_markers(
                    ((config.level_value or 0, sheet.get(source_row, config.source_col)) for config in group),
                    sheet_name=sheet.name,
                    source_row=source_row,
                )
                values["level"] = level
                if level_issue:
                    issues.append(level_issue)

            sequence = as_number(_first(raw, "sequence"))
            quantity = as_number(_first(raw, "quantity"))
            values["sequence"] = display_number(sequence)
            values["quantity"] = display_number(quantity)

            name_result = generate_name_from_columns(
                analysis.name_rule,
                lambda column: sheet.get(source_row, column),
                fallback_original=_first(raw, "part_name"),
                fallback_standard=_first(raw, "standard_name"),
                fallback_gb=_first(raw, "standard_name_gb"),
                fallback_spec=_first(raw, "spec"),
            )
            values["part_name"] = name_result.final_name
            if name_result.final_name and name_result.final_name != name_result.original_name:
                issues.append(Issue(
                    "GENERATED_NAME", "info",
                    "名称已按本机确认的名称生成规则处理",
                    sheet.name, source_row, "part_name", calculated_value=name_result.final_name,
                ))

            unit_weight = convert_weight(_first(raw, "unit_weight"), units.get("unit_weight", "kg"))
            total_weight = convert_weight(_first(raw, "total_weight"), units.get("total_weight", "kg"))
            values["unit_weight"] = display_number(unit_weight)
            values["total_weight"] = display_number(total_weight)
            if unit_weight is not None and quantity is not None and total_weight is not None:
                calculated = unit_weight * quantity
                tolerance = max(abs(total_weight) * .02, .001)
                if abs(calculated - total_weight) > tolerance:
                    issues.append(Issue(
                        "WEIGHT_MISMATCH", "warning",
                        "单件重量×数量与总重不一致，已保留来源值",
                        sheet.name, source_row, "total_weight", total_weight, calculated,
                    ))

            dimensions, dimension_error = parse_dimensions(
                _first(raw, "dimensions"), units.get("dimensions", "mm")
            )
            values.update({field: display_number(number) for field, number in dimensions.items()})
            if dimension_error:
                issues.append(Issue(
                    "DIMENSION_AMBIGUOUS", "warning", dimension_error,
                    sheet.name, source_row, "dimensions", _first(raw, "dimensions"),
                ))
            for field in ("length", "width", "height", "unfold_length", "diameter"):
                direct = convert_length(_first(raw, field), units.get(field, "mm"))
                if direct is not None:
                    values[field] = display_number(direct)
            thickness = convert_length(_first(raw, "thickness"), units.get("thickness", "mm"))
            area = convert_area(_first(raw, "surface_area"), units.get("surface_area", "m²"))
            values["thickness"] = display_number(thickness)
            values["surface_area"] = display_number(area)

            material = _first(raw, "material_raw") or _first(raw, "material_spec")
            material_type, material_spec = classify_material(material)
            if _first(raw, "material_type") is not None:
                material_type = clean_text(_first(raw, "material_type"))
            values["material_type"] = material_type
            values["material_spec"] = material_spec

            image_configs = {
                column: config for column, config in column_configs.items()
                if config.process_type == "image"
            }
            source_images = sorted(images_by_row.get(source_row, []), key=lambda item: (item.source_col, item.order))
            if image_configs:
                row_images = [
                    replace(image, target_slot=image_configs[image.source_col].image_slot)
                    for image in source_images
                    if image.source_col in image_configs
                ]
            else:
                row_images = source_images
            if len(row_images) > 7:
                issues.append(Issue(
                    "IMAGE_OVERFLOW", "error",
                    f"该行有 {len(row_images)} 张图片，目标模板最多 7 张",
                    sheet.name, source_row, "images",
                ))
            result.append(BomRow(sheet.name, source_row, values, row_images, unused, []))
    return result, issues

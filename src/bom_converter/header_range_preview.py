from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

from .header_region import build_header_region, detect_header_regions
from .models import HeaderRegion
from .text_utils import clean_text, column_letter, normalize_header, split_cell_reference
from .xlsx_reader import SheetData, XlsxReader


@dataclass(frozen=True)
class RangeDiagnostic:
    severity: str
    code: str
    message: str
    row: int | None = None


@dataclass(frozen=True)
class SheetPreview:
    sheet_name: str
    column_letters: list[str]
    rows: list[tuple[int, list[Any]]]
    max_row: int
    max_col: int
    merged_ranges: list[str]


def load_sheet_preview(
    path: str | Path,
    sheet_name: str,
    *,
    max_rows: int = 50,
    max_cols: int = 24,
) -> SheetPreview:
    with XlsxReader(path) as reader:
        sheet = reader.read_sheet(sheet_name)
    column_count = min(max(sheet.max_col, 1), max_cols)
    rows = [
        (row, [sheet.get(row, column, merged=True) for column in range(1, column_count + 1)])
        for row in range(1, min(max(sheet.max_row, 1), max_rows) + 1)
    ]
    return SheetPreview(
        sheet.name,
        [column_letter(column) for column in range(1, column_count + 1)],
        rows,
        sheet.max_row,
        sheet.max_col,
        list(sheet.merged_ranges),
    )


def header_candidates(
    path: str | Path,
    sheet_name: str | None = None,
    *,
    limit: int = 12,
) -> list[HeaderRegion]:
    with XlsxReader(path) as reader:
        candidates = detect_header_regions(reader, max_scan_row=50)
    filtered = [item for item in candidates if not sheet_name or item.sheet_name == sheet_name]
    if not sheet_name:
        best_by_sheet: list[HeaderRegion] = []
        seen_sheets: set[str] = set()
        for candidate in filtered:
            if candidate.sheet_name not in seen_sheets:
                best_by_sheet.append(candidate)
                seen_sheets.add(candidate.sheet_name)
        filtered = [*best_by_sheet, *[item for item in filtered if item not in best_by_sheet]]
    result: list[HeaderRegion] = []
    seen: set[tuple[str, int, int, int]] = set()
    for candidate in filtered:
        key = (
            candidate.sheet_name,
            candidate.header_start_row,
            candidate.header_end_row,
            candidate.data_start_row,
        )
        if key in seen:
            continue
        seen.add(key)
        result.append(candidate)
        if len(result) >= limit:
            break
    return result


def _merge_bounds(reference: str) -> tuple[int, int, int, int]:
    start, end = reference.split(":") if ":" in reference else (reference, reference)
    row1, col1 = split_cell_reference(start)
    row2, col2 = split_cell_reference(end)
    return row1, col1, row2, col2


def _populated(sheet: SheetData, row: int) -> list[Any]:
    return [
        sheet.get(row, column)
        for column in range(1, sheet.max_col + 1)
        if clean_text(sheet.get(row, column))
    ]


def _looks_like_data(values: list[Any]) -> bool:
    if not values:
        return False
    numeric = sum(isinstance(value, (int, float)) for value in values)
    identifiers = sum(
        normalize_header(value).startswith(("sim", "id", "pn"))
        or any(char.isdigit() for char in str(value)) and " " not in str(value).strip()
        for value in values
    )
    return numeric + identifiers >= max(2, (len(values) + 1) // 2)


def validate_header_range(
    sheet: SheetData,
    header_start_row: int,
    header_end_row: int,
    data_start_row: int,
) -> list[RangeDiagnostic]:
    diagnostics: list[RangeDiagnostic] = []
    if min(header_start_row, header_end_row, data_start_row) < 1:
        return [RangeDiagnostic("error", "ROW_BELOW_ONE", "行号必须从 1 开始")]
    if header_start_row > header_end_row:
        diagnostics.append(RangeDiagnostic("error", "HEADER_ORDER", "表头开始行不能大于表头结束行"))
    if data_start_row <= header_end_row:
        diagnostics.append(RangeDiagnostic("error", "DATA_BEFORE_HEADER_END", "数据开始行必须大于表头结束行", data_start_row))
    if header_start_row > sheet.max_row:
        diagnostics.append(RangeDiagnostic("error", "HEADER_AFTER_SHEET", f"表头开始行超过工作表最后一行 {sheet.max_row}"))
    if any(item.severity == "error" and item.code in {"HEADER_ORDER", "ROW_BELOW_ONE"} for item in diagnostics):
        return diagnostics

    for reference in sheet.merged_ranges:
        row1, _, row2, _ = _merge_bounds(reference)
        if row1 < header_start_row <= row2:
            diagnostics.append(RangeDiagnostic(
                "error", "MERGE_CUT_AT_START", f"表头开始行截断合并单元格 {reference}", header_start_row,
            ))
        if row1 <= header_end_row < row2:
            diagnostics.append(RangeDiagnostic(
                "error", "MERGE_CUT_AT_END", f"表头结束行截断合并单元格 {reference}", header_end_row,
            ))
        if row1 < data_start_row <= row2:
            diagnostics.append(RangeDiagnostic(
                "error", "MERGE_CUT_AT_DATA", f"数据开始行落在合并单元格 {reference} 中间", data_start_row,
            ))

    if not diagnostics or not any(item.severity == "error" for item in diagnostics):
        region = build_header_region(sheet, header_start_row, header_end_row, data_start_row)
        recognized = sum(bool(column.recommended_field) for column in region.columns)
        if recognized < 2:
            diagnostics.append(RangeDiagnostic(
                "warning", "TOO_FEW_FIELDS", f"当前表头只识别出 {recognized} 个字段，建议检查范围或手动映射",
            ))

    for row in range(header_start_row, min(header_end_row, sheet.max_row) + 1):
        if _looks_like_data(_populated(sheet, row)):
            diagnostics.append(RangeDiagnostic(
                "warning", "DATA_IN_HEADER", f"第 {row} 行更像数据，可能不应包含在表头中", row,
            ))

    if data_start_row <= sheet.max_row and not _populated(sheet, data_start_row):
        next_row = next(
            (row for row in range(data_start_row + 1, sheet.max_row + 1) if _populated(sheet, row)),
            None,
        )
        message = f"数据开始行第 {data_start_row} 行为空"
        if next_row:
            message += f"；下一条可能数据在第 {next_row} 行"
        diagnostics.append(RangeDiagnostic("warning", "EMPTY_DATA_START", message, data_start_row))

    return diagnostics

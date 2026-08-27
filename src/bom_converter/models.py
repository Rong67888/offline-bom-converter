from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class ImageRef:
    source_path: Path
    archive_part: str
    source_sheet: str
    source_row: int
    source_col: int
    order: int
    extension: str
    byte_size: int
    width_px: int | None = None
    height_px: int | None = None
    target_slot: int | None = None


@dataclass
class HeaderColumn:
    source_col: int
    column_letter: str
    header_path: list[str]
    logical_header: str
    parent_header: str | None = None
    child_header: str | None = None
    unit: str | None = None
    recommended_field: str | None = None
    recommended_process: str = "ignore"
    confidence: float = 0.0
    confidence_reasons: list[str] = field(default_factory=list)
    bilingual_alias: bool = False
    recommendation_conflict: bool = False
    level_group: str | None = None
    level_value: int | None = None


@dataclass
class HeaderRegion:
    sheet_name: str
    header_start_row: int
    header_end_row: int
    data_start_row: int
    columns: list[HeaderColumn]
    confidence: float
    confidence_reasons: list[str] = field(default_factory=list)
    merged_structure_summary: list[str] = field(default_factory=list)


@dataclass
class ColumnMappingConfig:
    source_col: int
    process_type: str = "direct"
    universal_field: str | None = None
    unit: str | None = None
    image_slot: int | None = None
    level_group: str | None = None
    level_value: int | None = None
    default_value: Any = None


@dataclass
class NameGenerationRule:
    """Header-only name composition settings; never stores row values."""

    strategy: str = "fallback"
    original_name_col: int | None = None
    standard_name_col: int | None = None
    gb_name_col: int | None = None
    spec_col: int | None = None
    template: str = "{名称} {GB} {规格}"
    deduplicate: bool = True


@dataclass
class MappingDecision:
    source_col: int
    source_header: str
    sample_value: Any
    universal_field: str | None
    target_header: str | None
    confidence: float
    status: str = "mapped"
    note: str = ""
    column_letter: str = ""
    header_path: list[str] = field(default_factory=list)
    parent_header: str | None = None
    child_header: str | None = None
    source_unit: str | None = None
    process_type: str = "direct"
    image_slot: int | None = None
    level_group: str | None = None
    level_value: int | None = None
    default_value: Any = None
    confidence_reasons: list[str] = field(default_factory=list)


@dataclass
class Issue:
    code: str
    severity: str
    message: str
    source_sheet: str | None = None
    source_row: int | None = None
    field: str | None = None
    source_value: Any = None
    calculated_value: Any = None


@dataclass
class BomRow:
    source_sheet: str
    source_row: int
    values: dict[str, Any] = field(default_factory=dict)
    images: list[ImageRef] = field(default_factory=list)
    unused_fields: dict[str, Any] = field(default_factory=dict)
    remarks: list[str] = field(default_factory=list)


@dataclass
class WorkbookAnalysis:
    source_path: Path
    profile_id: str
    profile_name: str
    sheet_name: str
    header_rows: list[int]
    data_start_row: int
    data_row_count: int
    image_count: int
    fingerprint: str
    profile_confidence: float
    mappings: list[MappingDecision]
    ignored_nonempty_columns: list[str]
    requires_review: bool
    issues: list[Issue] = field(default_factory=list)
    header_region: HeaderRegion | None = None
    name_rule: NameGenerationRule | None = None

    @property
    def header_start_row(self) -> int:
        return self.header_region.header_start_row if self.header_region else min(self.header_rows)

    @property
    def header_end_row(self) -> int:
        return self.header_region.header_end_row if self.header_region else max(self.header_rows)

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["source_path"] = str(self.source_path)
        data["header_start_row"] = self.header_start_row
        data["header_end_row"] = self.header_end_row
        return data


@dataclass
class SheetAnalysis:
    """Per-sheet discovery result used by the UI; never stores row business values."""

    source_path: Path
    sheet_name: str
    hidden: bool
    analysis: WorkbookAnalysis | None
    is_bom: bool
    recommended_selected: bool
    status: str
    error: str | None = None


@dataclass
class ConversionResult:
    source_path: Path
    output_path: Path
    profile_id: str
    source_rows: int
    output_rows: int
    source_images: int
    output_images: int
    issues: list[Issue]
    report_path: Path | None = None
    source_sheet: str | None = None

    @property
    def success(self) -> bool:
        return self.source_rows == self.output_rows and self.source_images == self.output_images

    def to_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["source_path"] = str(self.source_path)
        data["output_path"] = str(self.output_path)
        data["report_path"] = str(self.report_path) if self.report_path else None
        data["success"] = self.success
        return data

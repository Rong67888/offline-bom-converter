from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ColumnRule:
    source_header: str
    field: str | None
    confidence: float = 1.0
    status: str = "mapped"
    note: str = ""
    image_slot: int | None = None


@dataclass(frozen=True)
class SourceProfile:
    profile_id: str
    name: str
    header_rows: tuple[int, ...]
    data_start_row: int
    columns: tuple[ColumnRule, ...]
    default_units: dict[str, str] = field(default_factory=dict)
    level_columns: tuple[int, ...] = ()


TARGET_HEADERS = {
    "category": "Category",
    "sequence": "Row No.",
    "part_name": "Part Name",
    "part_number": "Part Number",
    "level": "Assembly Level",
    "electronics_spec": "Component Specification",
    "electronics_silk": "Marking",
    "electronics_package": "Package",
    "pin_number": "Pin Count",
    "pcb_side": "Board Side",
    "electronics_type": "Component Type",
    "unit_weight": "Unit Weight (kg)",
    "quantity": "Quantity",
    "total_weight": "Total Weight (kg)",
    "material_type": "Material Category",
    "material_spec": "Material Specification",
    "length": "Length (mm)",
    "width": "Width (mm)",
    "height": "Height (mm)",
    "unfold_length": "Unfolded Length (mm)",
    "diameter": "Diameter (mm)",
    "thickness": "Thickness (mm)",
    "production_process": "Production Process",
    "assembly_process": "Assembly Process",
    "surface_treatment": "Surface Treatment",
    "surface_area": "Surface Area (m2)",
    "manufacturer": "Manufacturer",
    "location": "Origin",
    "remark": "Notes",
    "code": "Reference Code",
    "vpc": "Group Code",
}


def r(
    header: str,
    field: str | None,
    confidence: float = 1.0,
    *,
    status: str = "mapped",
    note: str = "",
    image_slot: int | None = None,
) -> ColumnRule:
    return ColumnRule(header, field, confidence, status, note, image_slot)


# Private-workbook profiles are deliberately absent. This small profile is
# entirely fictional and only demonstrates the bundled single-header SIM file.
PROFILES: tuple[SourceProfile, ...] = (
    SourceProfile(
        "sim_demo",
        "Fictional SIM single-header format",
        (4,),
        5,
        (
            r("Row", "sequence"),
            r("Part Name", "part_name"),
            r("Part Number", "part_number"),
            r("Assembly Level", "level"),
            r("Quantity", "quantity"),
            r("Unit Weight (g)", "unit_weight"),
            r("Size (cm)", "dimensions"),
            r("Material", "material_raw"),
            r("Process", "production_process"),
            r("Picture", "image", image_slot=1),
        ),
        {"unit_weight": "g", "dimensions": "cm"},
    ),
)


PROFILE_BY_ID = {profile.profile_id: profile for profile in PROFILES}


GENERIC_SYNONYMS = {
    "序号": "sequence", "行号": "sequence", "no": "sequence", "row": "sequence",
    "名称": "part_name", "零件名称": "part_name", "partname": "part_name", "name": "part_name",
    "零件号": "part_number", "partnumber": "part_number",
    "层级": "level", "装配层级": "level", "assemblylevel": "level", "level": "level",
    "数量": "quantity", "qty": "quantity", "quantity": "quantity",
    "单件重量": "unit_weight", "unitweight": "unit_weight",
    "总重": "total_weight", "totalweight": "total_weight",
    "尺寸": "dimensions", "外形尺寸": "dimensions", "size": "dimensions", "dimension": "dimensions",
    "material": "material_raw", "材质": "material_raw", "材料": "material_raw",
    "manufacturer": "manufacturer", "制造商": "manufacturer",
    "origin": "location", "产地": "location", "thickness": "thickness", "厚度": "thickness",
    "process": "process_raw", "工艺": "process_raw", "notes": "remark", "备注": "remark",
    "referencecode": "code", "编号": "code", "groupcode": "vpc", "分组号": "vpc",
}

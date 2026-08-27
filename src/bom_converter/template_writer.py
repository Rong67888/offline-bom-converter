from __future__ import annotations

import copy
import os
import posixpath
import re
import shutil
import tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZIP_STORED, ZipFile, ZipInfo

from .models import BomRow, ImageRef
from .profiles import PROFILE_BY_ID, TARGET_HEADERS
from .text_utils import column_letter
from .transform import image_slot_for
from .xlsx_reader import NS_DRAW, NS_DRAW_MAIN, NS_MAIN, NS_REL_DOC, NS_REL_PKG, XlsxReader, _resolve_part


NS_CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
NS_X14AC = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
NS_XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
NS_XR2 = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
NS_XR3 = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"
CALC_CHAIN_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
CALC_CHAIN_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"

ET.register_namespace("", NS_MAIN)
ET.register_namespace("r", NS_REL_DOC)
ET.register_namespace("xdr", NS_DRAW)
ET.register_namespace("a", NS_DRAW_MAIN)
ET.register_namespace("mc", NS_MC)
ET.register_namespace("x14ac", NS_X14AC)
ET.register_namespace("xr", NS_XR)
ET.register_namespace("xr2", NS_XR2)
ET.register_namespace("xr3", NS_XR3)


def _safe_filename_component(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip().rstrip(". ")
    return cleaned or "工作表"


def _unique_output_path(output_dir: Path, source_stem: str) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    base = output_dir / f"{source_stem}_标准格式.xlsx"
    if not base.exists():
        return base
    index = 2
    while True:
        candidate = output_dir / f"{source_stem}_标准格式_{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def _clear_cell(cell: ET.Element) -> None:
    for child in list(cell):
        cell.remove(child)
    cell.attrib.pop("t", None)


def _write_cell(cell: ET.Element, value: object) -> None:
    _clear_cell(cell)
    if value is None:
        return
    if isinstance(value, bool):
        cell.set("t", "b")
        ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = "1" if value else "0"
    elif isinstance(value, (int, float)):
        ET.SubElement(cell, f"{{{NS_MAIN}}}v").text = str(value)
    else:
        cell.set("t", "inlineStr")
        inline = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
        text = ET.SubElement(inline, f"{{{NS_MAIN}}}t")
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = str(value)


def _cell_map(row: ET.Element) -> dict[int, ET.Element]:
    result: dict[int, ET.Element] = {}
    for cell in row.findall(f"{{{NS_MAIN}}}c"):
        reference = cell.attrib.get("r", "")
        letters = "".join(char for char in reference if char.isalpha())
        col = 0
        for char in letters:
            col = col * 26 + ord(char.upper()) - 64
        if col:
            result[col] = cell
    return result


def _clone_data_row(prototype: ET.Element, row_number: int, max_col: int) -> tuple[ET.Element, dict[int, ET.Element]]:
    row = copy.deepcopy(prototype)
    row.set("r", str(row_number))
    cells = _cell_map(row)
    default_style = cells.get(4, next(iter(cells.values()), None))
    for col in range(1, max_col + 1):
        if col not in cells:
            cell = ET.Element(f"{{{NS_MAIN}}}c", {"r": f"{column_letter(col)}{row_number}"})
            if default_style is not None and "s" in default_style.attrib:
                cell.set("s", default_style.attrib["s"])
            row.append(cell)
            cells[col] = cell
        cell = cells[col]
        cell.set("r", f"{column_letter(col)}{row_number}")
        _clear_cell(cell)
    return row, cells


def _make_anchor(row_zero: int, col_zero: int, rel_id: str, picture_id: int, name: str, width_px: int | None, height_px: int | None) -> ET.Element:
    max_cx, max_cy = 1_305_000, 756_000
    if width_px and height_px:
        scale = min(max_cx / width_px, max_cy / height_px)
        cx, cy = int(width_px * scale), int(height_px * scale)
    else:
        cx, cy = max_cx, max_cy
    anchor = ET.Element(f"{{{NS_DRAW}}}twoCellAnchor")
    marker = ET.SubElement(anchor, f"{{{NS_DRAW}}}from")
    ET.SubElement(marker, f"{{{NS_DRAW}}}col").text = str(col_zero)
    ET.SubElement(marker, f"{{{NS_DRAW}}}colOff").text = "38100"
    ET.SubElement(marker, f"{{{NS_DRAW}}}row").text = str(row_zero)
    ET.SubElement(marker, f"{{{NS_DRAW}}}rowOff").text = "38100"
    marker_to = ET.SubElement(anchor, f"{{{NS_DRAW}}}to")
    ET.SubElement(marker_to, f"{{{NS_DRAW}}}col").text = str(col_zero)
    ET.SubElement(marker_to, f"{{{NS_DRAW}}}colOff").text = str(38100 + cx)
    ET.SubElement(marker_to, f"{{{NS_DRAW}}}row").text = str(row_zero)
    ET.SubElement(marker_to, f"{{{NS_DRAW}}}rowOff").text = str(38100 + cy)
    pic = ET.SubElement(anchor, f"{{{NS_DRAW}}}pic")
    nv = ET.SubElement(pic, f"{{{NS_DRAW}}}nvPicPr")
    ET.SubElement(nv, f"{{{NS_DRAW}}}cNvPr", {"id": str(picture_id), "name": name})
    nv_pic = ET.SubElement(nv, f"{{{NS_DRAW}}}cNvPicPr")
    ET.SubElement(nv_pic, f"{{{NS_DRAW_MAIN}}}picLocks", {"noChangeAspect": "1"})
    fill = ET.SubElement(pic, f"{{{NS_DRAW}}}blipFill")
    ET.SubElement(fill, f"{{{NS_DRAW_MAIN}}}blip", {f"{{{NS_REL_DOC}}}embed": rel_id})
    stretch = ET.SubElement(fill, f"{{{NS_DRAW_MAIN}}}stretch")
    ET.SubElement(stretch, f"{{{NS_DRAW_MAIN}}}fillRect")
    shape = ET.SubElement(pic, f"{{{NS_DRAW}}}spPr")
    transform = ET.SubElement(shape, f"{{{NS_DRAW_MAIN}}}xfrm")
    ET.SubElement(transform, f"{{{NS_DRAW_MAIN}}}off", {"x": "0", "y": "0"})
    ET.SubElement(transform, f"{{{NS_DRAW_MAIN}}}ext", {"cx": str(cx), "cy": str(cy)})
    geometry = ET.SubElement(shape, f"{{{NS_DRAW_MAIN}}}prstGeom", {"prst": "rect"})
    ET.SubElement(geometry, f"{{{NS_DRAW_MAIN}}}avLst")
    ET.SubElement(anchor, f"{{{NS_DRAW}}}clientData")
    return anchor


def _serialize(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _serialize_with_original_root(root: ET.Element, original: bytes, root_name: str) -> bytes:
    """Serialize edited XML while preserving the template's root namespace declarations.

    ElementTree only emits namespaces used by element/attribute QNames. Prefixes named
    solely inside mc:Ignorable are therefore otherwise dropped, although Microsoft Excel
    requires every one of those prefixes to remain declared in the current XML scope.
    """
    # Public/user-created templates may use either a default namespace
    # (``<worksheet>``) or an explicit prefix (``<x:worksheet>``). Register
    # the original declarations before serialization so child elements keep
    # the same valid prefix vocabulary as the preserved root element.
    opening_pattern = re.compile(
        rb"<(?:[A-Za-z_][A-Za-z0-9_.-]*:)?"
        + re.escape(root_name.encode("ascii"))
        + rb"\b[^>]*>"
    )
    original_match = opening_pattern.search(original)
    if original_match is None:
        raise ValueError(f"无法保留 {root_name} 根节点的命名空间声明")
    for namespace in re.finditer(
        rb"\bxmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?=([\"'])(.*?)\2",
        original_match.group(0),
    ):
        prefix = (namespace.group(1) or b"").decode("ascii")
        uri = namespace.group(3).decode("utf-8")
        if prefix not in {"xml", "xmlns"}:
            ET.register_namespace(prefix, uri)
    serialized = _serialize(root)
    pattern = opening_pattern
    serialized_match = pattern.search(serialized)
    if serialized_match is None:
        raise ValueError(f"无法保留 {root_name} 根节点的命名空间声明")
    preserved_root = original_match.group(0)
    declared_prefixes = {
        (item.group(1) or b"")
        for item in re.finditer(
            rb"\bxmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?=([\"'])(.*?)\2",
            preserved_root,
        )
    }
    # ElementTree may promote a namespace that was originally declared on a
    # child element (for example r:id in workbook.xml) to the serialized root.
    # Keep those promoted declarations as well, otherwise replacing the root
    # tag would leave an unbound prefix.
    for item in re.finditer(
        rb"\bxmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?=([\"'])(.*?)\2",
        serialized_match.group(0),
    ):
        prefix = item.group(1) or b""
        if prefix in declared_prefixes:
            continue
        preserved_root = preserved_root[:-1] + b" " + item.group(0) + b">"
        declared_prefixes.add(prefix)
    return (
        serialized[: serialized_match.start()]
        + preserved_root
        + serialized[serialized_match.end() :]
    )


def _serialize_default_namespace(root: ET.Element, namespace: str) -> bytes:
    ET.register_namespace("", namespace)
    try:
        return ET.tostring(root, encoding="utf-8", xml_declaration=True)
    finally:
        ET.register_namespace("", NS_MAIN)


def _prepare_drawing(template_zip: ZipFile, sheet_part: str, rows: list[BomRow], profile_id: str, target_columns: dict[str, int]) -> tuple[dict[str, bytes], list[tuple[str, ImageRef]]]:
    sheet_root = ET.fromstring(template_zip.read(sheet_part))
    drawing_ref = sheet_root.find(f"{{{NS_MAIN}}}drawing")
    if drawing_ref is None:
        raise ValueError("标准模板缺少 Drawing 关系，无法安全插入图片")
    sheet_rels_part = posixpath.join(posixpath.dirname(sheet_part), "_rels", posixpath.basename(sheet_part) + ".rels")
    sheet_rels_root = ET.fromstring(template_zip.read(sheet_rels_part))
    sheet_rel = next(rel for rel in sheet_rels_root if rel.attrib.get("Id") == drawing_ref.attrib[f"{{{NS_REL_DOC}}}id"])
    drawing_part = _resolve_part(sheet_part, sheet_rel.attrib["Target"])
    drawing_rels_part = posixpath.join(posixpath.dirname(drawing_part), "_rels", posixpath.basename(drawing_part) + ".rels")
    drawing_root = ET.fromstring(template_zip.read(drawing_part))
    drawing_rels_root = ET.fromstring(template_zip.read(drawing_rels_part))
    for anchor in list(drawing_root):
        marker = anchor.find(f"{{{NS_DRAW}}}from")
        if marker is not None and int(marker.findtext(f"{{{NS_DRAW}}}row", "0")) >= 6:
            drawing_root.remove(anchor)
    existing_ids = []
    for node in drawing_root.findall(f".//{{{NS_DRAW}}}cNvPr"):
        try:
            existing_ids.append(int(node.attrib.get("id", "0")))
        except ValueError:
            pass
    used_rel_ids = {rel.attrib.get("Id", "") for rel in drawing_rels_root}
    media_parts: list[tuple[str, ImageRef]] = []
    existing_parts = set(template_zip.namelist())
    profile = PROFILE_BY_ID.get(profile_id)
    picture_columns = {slot: target_columns["Label Picture" if slot == 7 else f"Picture{slot}"] for slot in range(1, 8)}
    picture_id = max(existing_ids, default=0) + 1
    media_index = 1
    for output_index, row in enumerate(rows, start=7):
        used_slots: set[int] = set()
        for image in row.images:
            slot = (
                image.target_slot
                if image.target_slot and image.target_slot not in used_slots
                else image_slot_for(profile, image.source_col, used_slots)
            )
            if slot is None:
                continue
            used_slots.add(slot)
            while True:
                rel_id = f"rIdBomConv{media_index}"
                media_part = f"xl/media/bomconv_{media_index}.{image.extension or 'jpeg'}"
                if rel_id not in used_rel_ids and media_part not in existing_parts:
                    break
                media_index += 1
            used_rel_ids.add(rel_id)
            existing_parts.add(media_part)
            rel = ET.SubElement(drawing_rels_root, f"{{{NS_REL_PKG}}}Relationship")
            rel.set("Id", rel_id)
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
            rel.set("Target", f"../media/{posixpath.basename(media_part)}")
            anchor = _make_anchor(output_index - 1, picture_columns[slot] - 1, rel_id, picture_id, f"BOM image {picture_id}", image.width_px, image.height_px)
            drawing_root.append(anchor)
            media_parts.append((media_part, image))
            picture_id += 1
            media_index += 1
    return {drawing_part: _serialize(drawing_root), drawing_rels_part: _serialize_default_namespace(drawing_rels_root, NS_REL_PKG)}, media_parts


def _ensure_content_types(
    data: bytes,
    media_parts: list[tuple[str, ImageRef]],
    removed_parts: set[str],
) -> bytes:
    root = ET.fromstring(data)
    for override in list(root.findall(f"{{{NS_CONTENT}}}Override")):
        part_name = override.attrib.get("PartName", "").lstrip("/")
        content_type = override.attrib.get("ContentType", "")
        if part_name in removed_parts or content_type == CALC_CHAIN_CONTENT_TYPE:
            root.remove(override)
    known = {item.attrib.get("Extension", "").casefold() for item in root.findall(f"{{{NS_CONTENT}}}Default")}
    mime = {"jpeg": "image/jpeg", "jpg": "image/jpeg", "png": "image/png", "gif": "image/gif", "bmp": "image/bmp"}
    for extension in sorted({image.extension.casefold() for _, image in media_parts}):
        if extension not in known:
            ET.SubElement(root, f"{{{NS_CONTENT}}}Default", {"Extension": extension, "ContentType": mime.get(extension, f"image/{extension}")})
    return _serialize_default_namespace(root, NS_CONTENT)


def _remove_calc_chain_relationships(data: bytes) -> tuple[bytes, set[str]]:
    root = ET.fromstring(data)
    removed_parts = {"xl/calcChain.xml"}
    for relationship in list(root.findall(f"{{{NS_REL_PKG}}}Relationship")):
        if relationship.attrib.get("Type") != CALC_CHAIN_REL_TYPE:
            continue
        target = relationship.attrib.get("Target", "")
        if target:
            removed_parts.add(_resolve_part("xl/workbook.xml", target))
        root.remove(relationship)
    return _serialize_default_namespace(root, NS_REL_PKG), removed_parts


def _set_full_recalculation(data: bytes) -> bytes:
    root = ET.fromstring(data)
    calc_pr = root.find(f"{{{NS_MAIN}}}calcPr")
    if calc_pr is None:
        calc_pr = ET.Element(f"{{{NS_MAIN}}}calcPr")
        ext_list = root.find(f"{{{NS_MAIN}}}extLst")
        if ext_list is None:
            root.append(calc_pr)
        else:
            root.insert(list(root).index(ext_list), calc_pr)
    calc_pr.attrib.clear()
    calc_pr.attrib.update({
        "calcId": "0",
        "calcMode": "auto",
        "calcCompleted": "0",
        "fullCalcOnLoad": "1",
        "forceFullCalc": "1",
    })
    return _serialize_with_original_root(root, data, "workbook")


def _build_sheet_xml(template_zip: ZipFile, sheet_part: str, rows: list[BomRow]) -> tuple[bytes, dict[str, int]]:
    original_sheet_xml = template_zip.read(sheet_part)
    root = ET.fromstring(original_sheet_xml)
    sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
    if sheet_data is None:
        raise ValueError("标准模板缺少 sheetData")
    existing_rows = {int(row.attrib["r"]): row for row in sheet_data.findall(f"{{{NS_MAIN}}}row")}
    if 5 not in existing_rows or 6 not in existing_rows:
        raise ValueError("标准模板必须保留第 5 行格式样板和第 6 行正式表头")
    # Read headers directly from template cells, including shared strings.
    shared: list[str] = []
    if "xl/sharedStrings.xml" in template_zip.namelist():
        strings_root = ET.fromstring(template_zip.read("xl/sharedStrings.xml"))
        shared = ["".join(t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t")) for si in strings_root]
    header_cells = _cell_map(existing_rows[6])
    target_columns: dict[str, int] = {}
    for col, cell in header_cells.items():
        if cell.attrib.get("t") == "s":
            node = cell.find(f"{{{NS_MAIN}}}v")
            value = shared[int(node.text)] if node is not None else ""
        elif cell.attrib.get("t") == "inlineStr":
            value = "".join(t.text or "" for t in cell.findall(f".//{{{NS_MAIN}}}t"))
        else:
            node = cell.find(f"{{{NS_MAIN}}}v")
            value = node.text if node is not None else ""
        if value:
            target_columns[value] = col
    expected = set(TARGET_HEADERS.values()) | {f"Picture{i}" for i in range(1, 7)} | {"Label Picture"}
    missing = sorted(expected - set(target_columns))
    if missing:
        raise ValueError(f"标准模板缺少目标列: {', '.join(missing)}")
    format_prototype = existing_rows[5]
    prototype_cells = _cell_map(format_prototype)
    max_col = max(target_columns.values())
    missing_format_columns = [column_letter(col) for col in range(1, max_col + 1) if col not in prototype_cells]
    if missing_format_columns:
        preview = "、".join(missing_format_columns[:12])
        suffix = "……" if len(missing_format_columns) > 12 else ""
        raise ValueError(f"标准模板第 5 行格式结构不完整，缺少单元格：{preview}{suffix}")
    for row_number, element in list(existing_rows.items()):
        if row_number >= 7:
            sheet_data.remove(element)
    header_by_field = TARGET_HEADERS
    for output_row, bom_row in enumerate(rows, start=7):
        row_element, cells = _clone_data_row(format_prototype, output_row, max_col)
        data = dict(bom_row.values)
        if data.get("sequence") is None:
            data["sequence"] = output_row - 6
        for field, header in header_by_field.items():
            _write_cell(cells[target_columns[header]], data.get(field))
        sheet_data.append(row_element)
    dimension = root.find(f"{{{NS_MAIN}}}dimension")
    if dimension is not None:
        dimension.set("ref", f"A1:{column_letter(max_col)}{max(6, len(rows) + 6)}")
    return _serialize_with_original_root(root, original_sheet_xml, "worksheet"), target_columns


def write_template_output(
    template_path: str | Path,
    source_path: str | Path,
    output_dir: str | Path,
    rows: list[BomRow],
    profile_id: str,
    output_stem: str | None = None,
) -> Path:
    template = Path(template_path).resolve()
    source = Path(source_path).resolve()
    output = _unique_output_path(Path(output_dir).resolve(), _safe_filename_component(output_stem or source.stem))
    if output == source or output == template:
        raise ValueError("输出路径不得覆盖来源文件或标准模板")
    temp_handle = tempfile.NamedTemporaryFile(prefix="bomconv_", suffix=".xlsx.tmp", dir=output.parent, delete=False)
    temp_path = Path(temp_handle.name)
    temp_handle.close()
    try:
        with ZipFile(template) as template_zip:
            with XlsxReader(template) as template_reader:
                sheet_part = template_reader.read_sheet("Temp").part
            sheet_xml, target_columns = _build_sheet_xml(template_zip, sheet_part, rows)
            drawing_replacements, media_parts = _prepare_drawing(template_zip, sheet_part, rows, profile_id, target_columns)
            replacements = {sheet_part: sheet_xml, **drawing_replacements}
            workbook_rels, removed_parts = _remove_calc_chain_relationships(
                template_zip.read("xl/_rels/workbook.xml.rels")
            )
            replacements["xl/_rels/workbook.xml.rels"] = workbook_rels
            replacements["xl/workbook.xml"] = _set_full_recalculation(
                template_zip.read("xl/workbook.xml")
            )
            replacements["[Content_Types].xml"] = _ensure_content_types(
                template_zip.read("[Content_Types].xml"),
                media_parts,
                removed_parts,
            )
            with ZipFile(temp_path, "w") as output_zip:
                for info in template_zip.infolist():
                    if info.filename in removed_parts:
                        continue
                    if info.filename in replacements:
                        output_zip.writestr(info, replacements[info.filename])
                        continue
                    with template_zip.open(info) as input_stream, output_zip.open(info, "w") as output_stream:
                        shutil.copyfileobj(input_stream, output_stream, 1024 * 1024)
                source_archives: dict[Path, ZipFile] = {}
                try:
                    for media_part, image in media_parts:
                        archive = source_archives.setdefault(image.source_path, ZipFile(image.source_path))
                        info = ZipInfo(media_part)
                        info.compress_type = ZIP_STORED
                        with archive.open(image.archive_part) as input_stream, output_zip.open(info, "w") as output_stream:
                            shutil.copyfileobj(input_stream, output_stream, 1024 * 1024)
                finally:
                    for archive in source_archives.values():
                        archive.close()
        os.replace(temp_path, output)
        return output
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise

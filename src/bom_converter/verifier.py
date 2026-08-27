from __future__ import annotations

import hashlib
import posixpath
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

from .models import Issue
from .xlsx_reader import NS_DRAW_MAIN, NS_MAIN, NS_REL_DOC, NS_REL_PKG, XlsxReader, _resolve_part

NS_CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
CALC_CHAIN_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain"
CALC_CHAIN_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"


def sha256_file(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def markup_compatibility_errors(data: bytes) -> list[str]:
    """Return mc:Ignorable prefixes that are not declared in their XML scope."""
    errors: list[str] = []
    pending_namespaces: list[tuple[str, str]] = []
    scope_stack: list[dict[str, str]] = [{}]
    try:
        for event, item in ET.iterparse(BytesIO(data), events=("start-ns", "start", "end")):
            if event == "start-ns":
                prefix, uri = item
                pending_namespaces.append((prefix or "", uri))
                continue
            if event == "start":
                scope = scope_stack[-1].copy()
                scope.update(pending_namespaces)
                pending_namespaces.clear()
                ignorable = item.attrib.get(f"{{{NS_MC}}}Ignorable", "")
                for prefix in ignorable.split():
                    if prefix not in scope:
                        errors.append(f"mc:Ignorable 引用了未声明前缀 {prefix}")
                scope_stack.append(scope)
                continue
            scope_stack.pop()
    except ET.ParseError as exc:
        errors.append(f"XML 解析失败: {exc}")
    return errors


def _workbook_structure_hash(data: bytes) -> str:
    root = ET.fromstring(data)
    calc_pr = root.find(f"{{{NS_MAIN}}}calcPr")
    if calc_pr is not None:
        root.remove(calc_pr)
    return hashlib.sha256(ET.tostring(root)).hexdigest()


def _top_signature(path: Path) -> tuple:
    with XlsxReader(path) as reader:
        sheet = reader.read_sheet("Temp")
        values = tuple(sorted((key, value) for key, value in sheet.values.items() if key[0] <= 6))
        formulas = tuple(sorted((key, value) for key, value in sheet.formulas.items() if key[0] <= 6))
        sheet_part = sheet.part
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read(sheet_part))
        styles = tuple(sorted(
            (cell.attrib.get("r"), cell.attrib.get("s"))
            for cell in root.findall(f".//{{{NS_MAIN}}}c")
            if int("".join(char for char in cell.attrib.get("r", "0") if char.isdigit()) or 0) <= 6
        ))
        cols = root.find(f"{{{NS_MAIN}}}cols")
        cols_xml = ET.tostring(cols) if cols is not None else b""
        # calcPr is deliberately changed after calcChain removal; all other workbook
        # structure must still match the template.
        workbook_hash = _workbook_structure_hash(archive.read("xl/workbook.xml"))
        styles_hash = hashlib.sha256(archive.read("xl/styles.xml")).hexdigest()
    return values, formulas, styles, cols_xml, workbook_hash, styles_hash


def _package_compatibility_issues(path: Path) -> list[Issue]:
    issues: list[Issue] = []
    with ZipFile(path) as archive:
        names = set(archive.namelist())
        for part in sorted(name for name in names if name.endswith((".xml", ".rels"))):
            for message in markup_compatibility_errors(archive.read(part)):
                issues.append(Issue("MC_NAMESPACE", "error", f"{part}: {message}"))

        if "xl/calcChain.xml" in names:
            issues.append(Issue("CALC_CHAIN_PART", "error", "输出仍包含失效的 xl/calcChain.xml"))

        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        calc_relationships = [
            rel for rel in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship")
            if rel.attrib.get("Type") == CALC_CHAIN_REL_TYPE
        ]
        if calc_relationships:
            issues.append(Issue("CALC_CHAIN_REL", "error", "workbook relationships 仍引用 calcChain"))

        content_root = ET.fromstring(archive.read("[Content_Types].xml"))
        calc_overrides = [
            item for item in content_root.findall(f"{{{NS_CONTENT}}}Override")
            if item.attrib.get("ContentType") == CALC_CHAIN_CONTENT_TYPE
            or item.attrib.get("PartName", "").casefold() == "/xl/calcchain.xml"
        ]
        if calc_overrides:
            issues.append(Issue("CALC_CHAIN_CONTENT_TYPE", "error", "[Content_Types].xml 仍声明 calcChain"))

        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        calc_pr = workbook_root.find(f"{{{NS_MAIN}}}calcPr")
        required = {
            "calcMode": "auto",
            "calcCompleted": "0",
            "fullCalcOnLoad": "1",
            "forceFullCalc": "1",
        }
        if calc_pr is None or any(calc_pr.attrib.get(key) != value for key, value in required.items()):
            issues.append(Issue("RECALCULATION", "error", "workbook 未设置为由 Excel 打开时完整重新计算"))
    return issues


def _drawing_media_integrity(path: Path, sheet_name: str = "Temp") -> tuple[int, list[str]]:
    missing: list[str] = []
    with XlsxReader(path) as reader:
        sheet_part = reader.read_sheet(sheet_name).part
    with ZipFile(path) as archive:
        sheet_root = ET.fromstring(archive.read(sheet_part))
        drawing = sheet_root.find(f"{{{NS_MAIN}}}drawing")
        if drawing is None:
            return 0, []
        sheet_rels_part = posixpath.join(posixpath.dirname(sheet_part), "_rels", posixpath.basename(sheet_part) + ".rels")
        sheet_rels_root = ET.fromstring(archive.read(sheet_rels_part))
        rel_id = drawing.attrib[f"{{{NS_REL_DOC}}}id"]
        relation = next(rel for rel in sheet_rels_root if rel.attrib["Id"] == rel_id)
        drawing_part = _resolve_part(sheet_part, relation.attrib["Target"])
        drawing_root = ET.fromstring(archive.read(drawing_part))
        drawing_rels_part = posixpath.join(posixpath.dirname(drawing_part), "_rels", posixpath.basename(drawing_part) + ".rels")
        drawing_rels_root = ET.fromstring(archive.read(drawing_rels_part))
        targets = {rel.attrib["Id"]: rel.attrib["Target"] for rel in drawing_rels_root}
        count = 0
        for blip in drawing_root.findall(f".//{{{NS_DRAW_MAIN}}}blip"):
            embed = blip.attrib.get(f"{{{NS_REL_DOC}}}embed")
            if not embed:
                continue
            count += 1
            target = targets.get(embed)
            if not target or _resolve_part(drawing_part, target) not in archive.namelist():
                missing.append(embed)
        return count, missing


def verify_output(output_path: str | Path, template_path: str | Path, expected_rows: int, expected_images: int) -> tuple[int, int, list[Issue]]:
    output = Path(output_path).resolve()
    template = Path(template_path).resolve()
    issues: list[Issue] = []
    try:
        with ZipFile(output) as archive:
            corrupt = archive.testzip()
            if corrupt:
                issues.append(Issue("ZIP_CORRUPT", "error", f"输出 XLSX 内部文件损坏: {corrupt}"))
    except BadZipFile:
        return 0, 0, [Issue("XLSX_INVALID", "error", "输出文件不是有效的 XLSX 压缩包")]
    issues.extend(_package_compatibility_issues(output))
    with XlsxReader(output) as reader:
        sheet = reader.read_sheet("Temp")
        output_rows = max(sheet.max_row - 6, 0)
    output_images, missing = _drawing_media_integrity(output)
    if output_rows != expected_rows:
        issues.append(Issue("ROW_COUNT", "error", f"输出行数 {output_rows} 与来源数据行数 {expected_rows} 不一致"))
    if output_images != expected_images:
        issues.append(Issue("IMAGE_COUNT", "error", f"输出图片锚点 {output_images} 与来源图片数 {expected_images} 不一致"))
    if missing:
        issues.append(Issue("IMAGE_RELATIONSHIP", "error", f"{len(missing)} 个图片关系缺少媒体文件"))
    if _top_signature(output) != _top_signature(template):
        issues.append(Issue("TEMPLATE_CHANGED", "error", "模板第 1-6 行、列宽、样式表或命名区域发生非预期变化"))
    return output_rows, output_images, issues

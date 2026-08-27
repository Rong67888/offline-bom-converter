from __future__ import annotations

import posixpath
import struct
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from .models import ImageRef
from .text_utils import split_cell_reference


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_DRAW_MAIN = "http://schemas.openxmlformats.org/drawingml/2006/main"


@dataclass
class SheetData:
    name: str
    part: str
    values: dict[tuple[int, int], object]
    formulas: dict[tuple[int, int], str]
    merged_ranges: list[str]
    max_row: int
    max_col: int
    hidden: bool = False

    def get(self, row: int, col: int, merged: bool = False) -> object:
        value = self.values.get((row, col))
        if value is not None or not merged:
            return value
        for ref in self.merged_ranges:
            start, end = ref.split(":") if ":" in ref else (ref, ref)
            r1, c1 = split_cell_reference(start)
            r2, c2 = split_cell_reference(end)
            if r1 <= row <= r2 and c1 <= col <= c2:
                return self.values.get((r1, c1))
        return None


def _resolve_part(base_part: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


class XlsxReader:
    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()
        if self.path.suffix.lower() != ".xlsx":
            raise ValueError("阶段 1 仅支持 .xlsx 文件")
        self._zip = ZipFile(self.path)
        self._shared_strings = self._read_shared_strings()
        self._sheets, self._sheet_states = self._read_sheet_parts()

    def close(self) -> None:
        self._zip.close()

    def __enter__(self) -> "XlsxReader":
        return self

    def __exit__(self, *_: object) -> None:
        self.close()

    @property
    def sheet_names(self) -> list[str]:
        return list(self._sheets)

    def is_sheet_hidden(self, name: str) -> bool:
        return self._sheet_states.get(name, "visible") != "visible"

    @property
    def sheet_states(self) -> dict[str, str]:
        return dict(self._sheet_states)

    def _read_shared_strings(self) -> list[str]:
        if "xl/sharedStrings.xml" not in self._zip.namelist():
            return []
        root = ET.fromstring(self._zip.read("xl/sharedStrings.xml"))
        return ["".join(t.text or "" for t in si.findall(f".//{{{NS_MAIN}}}t")) for si in root]

    def _relationships(self, rel_part: str) -> dict[str, str]:
        if rel_part not in self._zip.namelist():
            return {}
        root = ET.fromstring(self._zip.read(rel_part))
        return {rel.attrib["Id"]: rel.attrib["Target"] for rel in root.findall(f"{{{NS_REL_PKG}}}Relationship")}

    def _read_sheet_parts(self) -> tuple[dict[str, str], dict[str, str]]:
        workbook = ET.fromstring(self._zip.read("xl/workbook.xml"))
        rels = self._relationships("xl/_rels/workbook.xml.rels")
        result: dict[str, str] = {}
        states: dict[str, str] = {}
        for sheet in workbook.findall(f".//{{{NS_MAIN}}}sheet"):
            rel_id = sheet.attrib[f"{{{NS_REL_DOC}}}id"]
            name = sheet.attrib["name"]
            result[name] = _resolve_part("xl/workbook.xml", rels[rel_id])
            states[name] = sheet.attrib.get("state", "visible")
        return result, states

    def read_sheet(self, name: str) -> SheetData:
        part = self._sheets[name]
        root = ET.fromstring(self._zip.read(part))
        values: dict[tuple[int, int], object] = {}
        formulas: dict[tuple[int, int], str] = {}
        max_row = max_col = 0
        for cell in root.findall(f".//{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r")
            if not ref:
                continue
            row, col = split_cell_reference(ref)
            max_row, max_col = max(max_row, row), max(max_col, col)
            formula = cell.find(f"{{{NS_MAIN}}}f")
            if formula is not None:
                formulas[(row, col)] = formula.text or ""
            cell_type = cell.attrib.get("t")
            if cell_type == "inlineStr":
                value: object = "".join(t.text or "" for t in cell.findall(f".//{{{NS_MAIN}}}t"))
            else:
                node = cell.find(f"{{{NS_MAIN}}}v")
                raw = node.text if node is not None else None
                if raw is None:
                    value = None
                elif cell_type == "s":
                    value = self._shared_strings[int(raw)]
                elif cell_type in {"str", "e"}:
                    value = raw
                elif cell_type == "b":
                    value = raw == "1"
                else:
                    try:
                        number = float(raw)
                        value = int(number) if number.is_integer() else number
                    except ValueError:
                        value = raw
            values[(row, col)] = value
        merged = [item.attrib["ref"] for item in root.findall(f".//{{{NS_MAIN}}}mergeCell")]
        return SheetData(name, part, values, formulas, merged, max_row, max_col, self.is_sheet_hidden(name))

    def read_images(self, sheet_name: str) -> list[ImageRef]:
        sheet_part = self._sheets[sheet_name]
        sheet_root = ET.fromstring(self._zip.read(sheet_part))
        drawing = sheet_root.find(f"{{{NS_MAIN}}}drawing")
        if drawing is None:
            return []
        sheet_rels_part = posixpath.join(posixpath.dirname(sheet_part), "_rels", posixpath.basename(sheet_part) + ".rels")
        sheet_rels = self._relationships(sheet_rels_part)
        drawing_part = _resolve_part(sheet_part, sheet_rels[drawing.attrib[f"{{{NS_REL_DOC}}}id"]])
        drawing_root = ET.fromstring(self._zip.read(drawing_part))
        drawing_rels_part = posixpath.join(posixpath.dirname(drawing_part), "_rels", posixpath.basename(drawing_part) + ".rels")
        drawing_rels = self._relationships(drawing_rels_part)
        result: list[ImageRef] = []
        order = 0
        for anchor in list(drawing_root):
            marker = anchor.find(f"{{{NS_DRAW}}}from")
            blip = anchor.find(f".//{{{NS_DRAW_MAIN}}}blip")
            if marker is None or blip is None:
                continue
            rel_id = blip.attrib.get(f"{{{NS_REL_DOC}}}embed")
            if not rel_id or rel_id not in drawing_rels:
                continue
            image_part = _resolve_part(drawing_part, drawing_rels[rel_id])
            if image_part not in self._zip.namelist():
                continue
            row = int(marker.findtext(f"{{{NS_DRAW}}}row", "0")) + 1
            col = int(marker.findtext(f"{{{NS_DRAW}}}col", "0")) + 1
            info = self._zip.getinfo(image_part)
            width, height = self._image_dimensions(image_part)
            order += 1
            result.append(ImageRef(
                source_path=self.path,
                archive_part=image_part,
                source_sheet=sheet_name,
                source_row=row,
                source_col=col,
                order=order,
                extension=Path(image_part).suffix.lower().lstrip("."),
                byte_size=info.file_size,
                width_px=width,
                height_px=height,
            ))
        return result

    def _image_dimensions(self, part: str) -> tuple[int | None, int | None]:
        with self._zip.open(part) as stream:
            head = stream.read(32)
            if head.startswith(b"\x89PNG\r\n\x1a\n") and len(head) >= 24:
                return struct.unpack(">II", head[16:24])
            if head.startswith(b"GIF8") and len(head) >= 10:
                return struct.unpack("<HH", head[6:10])
            if head.startswith(b"\xff\xd8"):
                stream.seek(0)
                stream.read(2)
                while True:
                    marker_start = stream.read(1)
                    if not marker_start:
                        break
                    if marker_start != b"\xff":
                        continue
                    marker = stream.read(1)
                    while marker == b"\xff":
                        marker = stream.read(1)
                    if marker in {b"\xd8", b"\xd9"}:
                        continue
                    length_raw = stream.read(2)
                    if len(length_raw) != 2:
                        break
                    length = struct.unpack(">H", length_raw)[0]
                    if marker and marker[0] in range(0xC0, 0xC4):
                        data = stream.read(5)
                        if len(data) == 5:
                            height, width = struct.unpack(">HH", data[1:5])
                            return width, height
                        break
                    stream.seek(max(length - 2, 0), 1)
        return None, None

    def copy_part_to(self, archive_part: str, target_stream: object) -> None:
        with self._zip.open(archive_part) as source:
            while True:
                chunk = source.read(1024 * 1024)
                if not chunk:
                    break
                target_stream.write(chunk)

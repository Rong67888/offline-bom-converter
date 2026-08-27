from __future__ import annotations

import argparse
import hashlib
import json
import re
import struct
import subprocess
import sys
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


ROOT = Path(__file__).resolve().parents[1]
MAX_FILE_BYTES = 10 * 1024 * 1024
TEXT_SUFFIXES = {".py", ".md", ".txt", ".toml", ".json", ".svg", ".gitignore"}
ALLOWED_XLSX = {
    "assets/public_demo_template.xlsx",
    "samples/simulated/sim_single_header.xlsx",
    "samples/simulated/sim_complex_headers.xlsx",
    "samples/simulated/sim_multisheet_levels.xlsx",
}
ALLOWED_IMAGE_PREFIX = "docs/images/"
ALLOWED_IMAGES = {
    "docs/images/ui_quick_sim.png",
    "docs/images/ui_mapping_sim.svg",
    "docs/images/ui_results_sim.png",
}
FORBIDDEN_SUFFIXES = {".exe", ".dll", ".pdf", ".xls", ".xlsm", ".xlsb", ".doc", ".docx", ".ppt", ".pptx"}
FORBIDDEN_NAMES = {"mapping_rules.json", ".env"}
FORBIDDEN_DIRS = {
    "dist", "build", "outputs", ".tmp_artifact", "teacher_test_package",
    "version_archive", "previous_release", "__pycache__", ".pytest_cache",
}


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sensitive_terms() -> list[str]:
    terms = ["BOM" + str(index) for index in range(1, 6)]
    terms.append("".join(chr(value) for value in (80, 65, 84, 65, 67)))
    terms.append("".join(chr(value) for value in (115, 107, 121, 95, 114)))
    return terms


def scan_text(label: str, text: str, violations: list[str]) -> None:
    lowered = text.casefold()
    for term in sensitive_terms():
        if re.search(rf"(?<![a-z0-9]){re.escape(term.casefold())}(?![a-z0-9])", lowered):
            violations.append(f"{label}: contains excluded identifier {term!r}")
    absolute_patterns = (
        r"[A-Za-z]:[\\/](?:Users|college)[\\/]",
        r"(?:^|[\s\"'])/(?:home|Users)/[^\s\"']+",
        r"wxid_[A-Za-z0-9_]+",
    )
    for pattern in absolute_patterns:
        if re.search(pattern, text, re.IGNORECASE | re.MULTILINE):
            violations.append(f"{label}: contains a local path or attachment identifier")
    secret_patterns = (
        r"gh[pousr]_[A-Za-z0-9]{30,}",
        r"github_pat_[A-Za-z0-9_]{40,}",
        r"AKIA[0-9A-Z]{16}",
        r"(?i)(api[_-]?key|access[_-]?token|client[_-]?secret|password)\s*[:=]\s*[\"'][^\"']{8,}[\"']",
        r"-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----",
    )
    for pattern in secret_patterns:
        if re.search(pattern, text):
            violations.append(f"{label}: contains a credential-like value")


def xlsx_audit(path: Path, violations: list[str], inventory: dict[str, object]) -> None:
    label = relative(path)
    try:
        with ZipFile(path) as archive:
            names = archive.namelist()
            dangerous = [
                name for name in names
                if name.startswith(("xl/externalLinks/", "customXml/"))
                or "vbaProject" in name
                or "/comments" in name
                or "/threadedComments" in name
            ]
            if dangerous:
                violations.append(f"{label}: forbidden OOXML parts {dangerous}")
            hidden_sheets: list[str] = []
            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            for sheet in workbook_root.iter():
                if sheet.tag.endswith("}sheet") and sheet.attrib.get("state", "visible") != "visible":
                    hidden_sheets.append(sheet.attrib.get("name", "?"))
            if hidden_sheets:
                violations.append(f"{label}: hidden worksheets {hidden_sheets}")
            external_relationships: list[str] = []
            for name in names:
                if not name.endswith(".rels"):
                    continue
                root = ET.fromstring(archive.read(name))
                for rel in root:
                    if rel.attrib.get("TargetMode") == "External":
                        external_relationships.append(f"{name}:{rel.attrib.get('Target', '')}")
            if external_relationships:
                violations.append(f"{label}: external relationships {external_relationships}")
            media = {
                name: sha256(archive.read(name))
                for name in names if name.startswith("xl/media/") and not name.endswith("/")
            }
            xml_text = "\n".join(
                archive.read(name).decode("utf-8", errors="ignore")
                for name in names if name.endswith((".xml", ".rels"))
            )
            scan_text(label + " OOXML", xml_text, violations)
            inventory[label] = {
                "sheets": [sheet.attrib.get("name", "") for sheet in workbook_root.iter() if sheet.tag.endswith("}sheet")],
                "hidden_sheets": hidden_sheets,
                "comments": 0,
                "external_relationships": len(external_relationships),
                "custom_xml_parts": sum(name.startswith("customXml/") for name in names),
                "macro_parts": sum("vbaProject" in name for name in names),
                "media_count": len(media),
                "media_sha256": sorted(media.values()),
            }
    except (BadZipFile, KeyError, ET.ParseError) as exc:
        violations.append(f"{label}: invalid XLSX package: {exc}")


def image_audit(path: Path, violations: list[str], inventory: dict[str, object]) -> None:
    label = relative(path)
    data = path.read_bytes()
    record: dict[str, object] = {"bytes": len(data), "sha256": sha256(data)}
    if path.suffix.casefold() == ".svg":
        try:
            root = ET.fromstring(data)
        except ET.ParseError as exc:
            violations.append(f"{label}: invalid SVG: {exc}")
            return
        external: list[str] = []
        for element in root.iter():
            for name, value in element.attrib.items():
                if name.endswith("href") and value and not value.startswith("#"):
                    external.append(value)
        if external:
            violations.append(f"{label}: SVG contains external or embedded resource references")
        scan_text(label + " SVG", data.decode("utf-8"), violations)
        record.update({"format": "svg", "external_resources": len(external)})
    elif data.startswith(b"\x89PNG\r\n\x1a\n"):
        offset = 8
        chunks: list[str] = []
        text_chunks: list[str] = []
        width = height = None
        while offset + 12 <= len(data):
            length = struct.unpack(">I", data[offset:offset + 4])[0]
            kind = data[offset + 4:offset + 8]
            payload = data[offset + 8:offset + 8 + length]
            chunks.append(kind.decode("ascii", errors="replace"))
            if kind == b"IHDR" and len(payload) >= 8:
                width, height = struct.unpack(">II", payload[:8])
            if kind in {b"tEXt", b"iTXt", b"zTXt"}:
                text_chunks.append(payload.decode("utf-8", errors="ignore"))
            offset += 12 + length
            if kind == b"IEND":
                break
        metadata_text = "\n".join(text_chunks)
        scan_text(label + " PNG metadata", metadata_text, violations)
        record.update({
            "format": "png",
            "width": width,
            "height": height,
            "chunk_types": chunks,
            "text_metadata_chunks": len(text_chunks),
        })
    else:
        violations.append(f"{label}: unsupported or invalid sanitized image format")
    inventory[label] = record


def tracked_files() -> list[str]:
    completed = subprocess.run(
        ["git", "ls-files"], cwd=ROOT, check=True, capture_output=True, text=True, encoding="utf-8"
    )
    return [line.strip() for line in completed.stdout.splitlines() if line.strip()]


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit the privacy-safe public repository export.")
    parser.add_argument("--require-git-tracked", action="store_true")
    args = parser.parse_args()
    violations: list[str] = []
    files: list[Path] = []
    for path in ROOT.rglob("*"):
        if not path.is_file() or ".git" in path.parts:
            continue
        files.append(path)
        rel = relative(path)
        parts = set(path.relative_to(ROOT).parts)
        if parts & FORBIDDEN_DIRS:
            violations.append(f"{rel}: forbidden directory")
        if path.name in FORBIDDEN_NAMES or path.name.endswith(".report.json"):
            violations.append(f"{rel}: forbidden local-state/report file")
        if path.suffix.casefold() in FORBIDDEN_SUFFIXES:
            violations.append(f"{rel}: forbidden binary/document type")
        if path.suffix.casefold() == ".xlsx" and rel not in ALLOWED_XLSX:
            violations.append(f"{rel}: XLSX is not in the explicit public allowlist")
        if path.suffix.casefold() in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".svg"}:
            if not rel.startswith(ALLOWED_IMAGE_PREFIX) or rel not in ALLOWED_IMAGES:
                violations.append(f"{rel}: image is outside the explicit sanitized allowlist")
        if path.stat().st_size > MAX_FILE_BYTES:
            violations.append(f"{rel}: file exceeds {MAX_FILE_BYTES} bytes")
        if path.suffix.casefold() in TEXT_SUFFIXES or path.name == ".gitignore":
            scan_text(rel, path.read_text(encoding="utf-8", errors="strict"), violations)

    xlsx_inventory: dict[str, object] = {}
    for rel in sorted(ALLOWED_XLSX):
        path = ROOT / rel
        if not path.exists():
            violations.append(f"{rel}: required public fixture is missing")
        else:
            xlsx_audit(path, violations, xlsx_inventory)

    image_inventory: dict[str, object] = {}
    for rel in sorted(ALLOWED_IMAGES):
        path = ROOT / rel
        if not path.exists():
            violations.append(f"{rel}: required sanitized image is missing")
        else:
            image_audit(path, violations, image_inventory)

    tracked: list[str] | None = None
    if args.require_git_tracked:
        tracked = tracked_files()
        actual = sorted(relative(path) for path in files)
        if sorted(tracked) != actual:
            missing = sorted(set(actual) - set(tracked))
            extra = sorted(set(tracked) - set(actual))
            violations.append(f"git tracked list mismatch; untracked={missing}, missing={extra}")

    summary = {
        "ok": not violations,
        "file_count": len(files),
        "total_bytes": sum(path.stat().st_size for path in files),
        "xlsx_inventory": xlsx_inventory,
        "image_inventory": image_inventory,
        "git_tracked_count": None if tracked is None else len(tracked),
        "violations": violations,
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0 if not violations else 1


if __name__ == "__main__":
    sys.exit(main())

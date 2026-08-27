from __future__ import annotations

import hashlib
import inspect
import json
import re
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from bom_converter.analyzer import analyze_manual_workbook, analyze_sheet, analyze_workbook_sheets
from bom_converter.confirmed_mapping import configs_from_analysis, transform_confirmed_workbook
from bom_converter.converter import convert_confirmed_file, convert_file
from bom_converter.gui_v2 import BomConverterAppV2
from bom_converter.header_region import resolve_level_markers
from bom_converter.issue_display import present_issue
from bom_converter.mapping_memory import MappingRuleStore
from bom_converter.models import ColumnMappingConfig, Issue, NameGenerationRule
from bom_converter.name_rules import generate_name, render_name_template
from bom_converter.profiles import PROFILES
from bom_converter.text_utils import column_letter
from bom_converter.transform import convert_area, convert_length, convert_weight, parse_dimensions
from bom_converter.verifier import markup_compatibility_errors
from bom_converter.xlsx_reader import NS_MAIN, XlsxReader


ROOT = Path(__file__).resolve().parents[1]
SAMPLES = ROOT / "samples" / "simulated"
TEMPLATE = ROOT / "assets" / "public_demo_template.xlsx"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class PublicFeatureTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.temp = tempfile.TemporaryDirectory(prefix="offline_bom_public_tests_")
        cls.output_dir = Path(cls.temp.name)
        cls.single = SAMPLES / "sim_single_header.xlsx"
        cls.complex = SAMPLES / "sim_complex_headers.xlsx"
        cls.multisheet = SAMPLES / "sim_multisheet_levels.xlsx"
        cls.protected_hashes = {path: sha256(path) for path in (TEMPLATE, cls.single, cls.complex, cls.multisheet)}
        cls.quick_result = convert_file(cls.single, TEMPLATE, cls.output_dir, mode="quick")

    @classmethod
    def tearDownClass(cls) -> None:
        for path, before in cls.protected_hashes.items():
            if sha256(path) != before:
                raise AssertionError(f"Protected public fixture changed: {path.name}")
        cls.temp.cleanup()

    def test_only_fictional_known_profile_is_shipped(self) -> None:
        self.assertEqual([profile.profile_id for profile in PROFILES], ["sim_demo"])
        self.assertIn("Fictional", PROFILES[0].name)

    def test_excel_column_labels(self) -> None:
        self.assertEqual({1: "A", 26: "Z", 27: "AA", 28: "AB", 52: "AZ", 53: "BA"},
                         {index: column_letter(index) for index in (1, 26, 27, 28, 52, 53)})

    def test_single_header_profile_and_images(self) -> None:
        analysis = analyze_sheet(self.single, "SIM Simple", mode="quick")
        self.assertEqual((analysis.profile_id, analysis.header_start_row, analysis.header_end_row, analysis.data_start_row),
                         ("sim_demo", 4, 4, 5))
        self.assertEqual((analysis.data_row_count, analysis.image_count), (3, 1))

    def test_complex_header_range_and_xyz_paths(self) -> None:
        analysis = analyze_manual_workbook(self.complex, "SIM XYZ Bilingual", 4, 5, 6)
        self.assertEqual((analysis.header_start_row, analysis.header_end_row, analysis.data_start_row), (4, 5, 6))
        by_letter = {item.column_letter: item for item in analysis.mappings}
        self.assertEqual(by_letter["G"].header_path, ["Size (mm)", "X"])
        self.assertEqual(by_letter["H"].header_path, ["Size (mm)", "Y"])
        self.assertEqual(by_letter["I"].header_path, ["Size (mm)", "Z"])
        self.assertEqual((by_letter["G"].universal_field, by_letter["H"].universal_field, by_letter["I"].universal_field),
                         ("length", "width", "height"))
        self.assertEqual((by_letter["G"].source_unit, by_letter["H"].source_unit, by_letter["I"].source_unit),
                         ("mm", "mm", "mm"))

    def test_bilingual_aliases_share_one_source_column(self) -> None:
        analysis = analyze_manual_workbook(self.complex, "SIM XYZ Bilingual", 4, 5, 6)
        name_mapping = next(item for item in analysis.mappings if item.column_letter == "B")
        self.assertEqual(name_mapping.universal_field, "part_name")
        self.assertEqual(len(name_mapping.header_path), 2)
        self.assertEqual(name_mapping.header_path[-1], "Part Name")

    def test_multisheet_selection_excludes_non_bom_sheets(self) -> None:
        sheets = analyze_workbook_sheets(self.multisheet, mode="audit")
        selected = [item.sheet_name for item in sheets if item.recommended_selected]
        rejected = [item.sheet_name for item in sheets if not item.recommended_selected]
        self.assertEqual(selected, ["SIM Front Module", "SIM Rear Module"])
        self.assertIn("Overview", rejected)
        self.assertIn("Empty Notes", rejected)

    def test_level_group_l1_to_l13(self) -> None:
        analysis = analyze_sheet(self.multisheet, "SIM Front Module", mode="audit")
        mappings = {item.source_col: item.universal_field for item in analysis.mappings}
        configs = configs_from_analysis(analysis, mappings)
        configs[17] = ColumnMappingConfig(17, "image", image_slot=1)
        rows, issues = transform_confirmed_workbook(self.multisheet, analysis, mappings, column_configs=configs)
        self.assertEqual([row.values["level"] for row in rows], [5, 6, 13])
        self.assertEqual(sum(len(row.images) for row in rows), 1)
        self.assertFalse([item for item in issues if item.severity == "error"])

    def test_level_multiple_and_missing_markers_warn(self) -> None:
        level, issue = resolve_level_markers([(5, 1), (6, "Y")], sheet_name="SIM", source_row=8)
        self.assertIsNone(level)
        self.assertEqual(issue.code, "LEVEL_MARKER_MULTIPLE")
        level, issue = resolve_level_markers([(5, None), (6, "")], sheet_name="SIM", source_row=9)
        self.assertIsNone(level)
        self.assertEqual(issue.code, "LEVEL_MARKER_MISSING")

    def test_units_and_dimensions(self) -> None:
        self.assertAlmostEqual(convert_weight("125 g"), 0.125)
        self.assertAlmostEqual(convert_length("12 cm"), 120)
        self.assertAlmostEqual(convert_area("2500 cm2"), 0.25)
        dimensions, warning = parse_dimensions("12 × 8 × 3 cm")
        self.assertIsNone(warning)
        self.assertEqual(dimensions, {"length": 120.0, "width": 80.0, "height": 30.0})

    def test_name_rules_skip_blanks_deduplicate_and_keep_multiplication_sign(self) -> None:
        rendered = render_name_template("{名称}（{GB}）{规格}", original=None, standard="SIM Bolt", gb="", spec="M8×20")
        self.assertEqual(rendered, "SIM BoltM8×20")
        self.assertNotIn("（）", rendered)
        result = generate_name(NameGenerationRule(strategy="append", template="{标准件} {规格}", deduplicate=True),
                               original="SIM Clip", standard="SIM Clip", spec="M4×10")
        self.assertIn("×", result.final_name)
        self.assertEqual(result.final_name.count("SIM Clip"), 1)

    def test_mapping_memory_persists_rules_without_examples(self) -> None:
        with tempfile.TemporaryDirectory(prefix="public_mapping_rules_") as directory:
            path = Path(directory) / "mapping_rules.json"
            store = MappingRuleStore(path)
            store.save_rule(
                "sim-fingerprint",
                header_rows=[4, 5],
                headers=["Part Name", "Size (mm) / X"],
                mappings={1: "part_name", 2: "length"},
                units={"length": "mm"},
                sheet_name="SIM XYZ Bilingual",
                header_start_row=4,
                header_end_row=5,
                data_start_row=6,
                header_paths=[["零件名称", "Part Name"], ["Size (mm)", "X"]],
                name_rule=NameGenerationRule(strategy="fallback", template="{名称} {GB} {规格}"),
            )
            reopened = MappingRuleStore(path)
            self.assertEqual(reopened.recall("sim-fingerprint"), {1: "part_name", 2: "length"})
            payload = path.read_text(encoding="utf-8")
            self.assertNotIn("Aurora Bracket", payload)
            self.assertNotIn("SIM-AUR-001", payload)
            saved_rule = json.loads(payload)["rules"]["sim-fingerprint"]
            self.assertNotIn("sample_value", saved_rule)
            self.assertNotIn("sample_values", saved_rule)

    def test_quick_conversion_rows_images_and_no_errors(self) -> None:
        result = self.quick_result
        self.assertEqual((result.output_rows, result.output_images), (3, 1))
        self.assertFalse([item for item in result.issues if item.severity == "error"])

    def test_image_bytes_are_copied_without_reencoding(self) -> None:
        with ZipFile(self.single) as source_zip, ZipFile(self.quick_result.output_path) as output_zip:
            source_hashes = {
                hashlib.sha256(source_zip.read(name)).hexdigest()
                for name in source_zip.namelist() if name.startswith("xl/media/")
            }
            output_hashes = {
                hashlib.sha256(output_zip.read(name)).hexdigest()
                for name in output_zip.namelist() if name.startswith("xl/media/")
            }
        self.assertTrue(source_hashes)
        self.assertTrue(source_hashes.issubset(output_hashes))

    def test_data_rows_reuse_template_row_five_storage_style(self) -> None:
        with XlsxReader(TEMPLATE) as reader:
            template_part = reader.read_sheet("Temp").part
        with XlsxReader(self.quick_result.output_path) as reader:
            output_part = reader.read_sheet("Temp").part
        with ZipFile(TEMPLATE) as archive:
            template_root = ET.fromstring(archive.read(template_part))
        with ZipFile(self.quick_result.output_path) as archive:
            output_root = ET.fromstring(archive.read(output_part))
        template_row = template_root.find(f".//{{{NS_MAIN}}}row[@r='5']")
        self.assertIsNotNone(template_row)
        expected = {re.sub(r"\d+", "", cell.attrib["r"]): cell.attrib.get("s") for cell in template_row}
        for row_number in (7, 8, 9):
            row = output_root.find(f".//{{{NS_MAIN}}}row[@r='{row_number}']")
            actual = {re.sub(r"\d+", "", cell.attrib["r"]): cell.attrib.get("s") for cell in row}
            self.assertEqual(actual, expected)
            self.assertFalse(row.findall(f".//{{{NS_MAIN}}}f"))

    def test_template_has_twelve_generic_formulas_and_output_preserves_them(self) -> None:
        for path in (TEMPLATE, self.quick_result.output_path):
            with XlsxReader(path) as reader:
                formulas = reader.read_sheet("Temp").formulas
            self.assertEqual(len(formulas), 12)

    def test_markup_compatibility_prefixes_are_valid(self) -> None:
        for path in (TEMPLATE, self.quick_result.output_path):
            with ZipFile(path) as archive:
                errors = []
                for name in archive.namelist():
                    if name.endswith((".xml", ".rels")):
                        errors.extend(markup_compatibility_errors(archive.read(name)))
            self.assertEqual(errors, [])

    def test_two_selected_sheets_create_independent_outputs(self) -> None:
        outputs = []
        for sheet_name in ("SIM Front Module", "SIM Rear Module"):
            analysis = analyze_sheet(self.multisheet, sheet_name, mode="audit")
            mappings = {item.source_col: item.universal_field for item in analysis.mappings}
            configs = configs_from_analysis(analysis, mappings)
            configs[17] = ColumnMappingConfig(17, "image", image_slot=1)
            result = convert_confirmed_file(self.multisheet, TEMPLATE, self.output_dir, analysis, mappings, column_configs=configs)
            outputs.append(result)
        self.assertEqual([(item.source_sheet, item.output_rows, item.output_images) for item in outputs],
                         [("SIM Front Module", 3, 1), ("SIM Rear Module", 2, 1)])
        self.assertNotEqual(outputs[0].output_path, outputs[1].output_path)
        self.assertIn("SIM Front Module", outputs[0].output_path.name)
        self.assertIn("SIM Rear Module", outputs[1].output_path.name)

    def test_issue_details_have_plain_language(self) -> None:
        view = present_issue(Issue("LOW_HEADER_CONFIDENCE", "warning", "SIM", "SIM Sheet"))
        self.assertEqual(view.severity_label, "需要检查")
        self.assertTrue(view.title)
        self.assertTrue(view.action)

    def test_horizontal_matrix_and_new_task_contract_remain_in_ui(self) -> None:
        source = inspect.getsource(BomConverterAppV2)
        self.assertIn("MappingMatrix", source)
        self.assertIn("_start_new_task", source)
        method = inspect.getsource(BomConverterAppV2._start_new_task)
        self.assertIn("_clear_sources", method)
        self.assertIn('self.progress["value"] = 0', method)
        self.assertNotIn("self.template.set", method)
        self.assertNotIn("self.output_dir.set", method)

    def test_public_source_contains_no_private_profile_identifiers(self) -> None:
        source_text = "\n".join(path.read_text(encoding="utf-8") for path in (ROOT / "src").rglob("*.py"))
        for index in range(1, 6):
            self.assertIsNone(re.search(rf"\bbom{index}\b", source_text, re.IGNORECASE))

    def test_private_upload_helpers_are_guarded_and_non_destructive(self) -> None:
        ps1 = (ROOT / "上传到GitHub私有仓库.ps1").read_text(encoding="utf-8-sig")
        bat = (ROOT / "双击上传到GitHub私有仓库.bat").read_text(encoding="utf-8")
        self.assertIn("Rong67888", ps1)
        self.assertIn("offline-bom-converter", ps1)
        self.assertIn("--private", ps1)
        self.assertIn("auth status", ps1)
        self.assertIn("--require-git-tracked", ps1)
        self.assertIn("status', '--porcelain", ps1)
        self.assertIn("push --set-upstream origin main", ps1)
        self.assertIn("LocalCheckOnly", ps1)
        self.assertIn("github_upload_result.txt", ps1)
        self.assertIn("上传成功，仓库仍为Private", ps1)
        self.assertIn("existingDefaultBranch", ps1)
        forbidden = ("--force", "--public", "repo delete", "release create", "release upload")
        for fragment in forbidden:
            self.assertNotIn(fragment, ps1.casefold())
        self.assertIn("%~dp0", bat)
        self.assertIn("powershell.exe", bat.casefold())
        self.assertIn("pause", bat.casefold())


if __name__ == "__main__":
    unittest.main()

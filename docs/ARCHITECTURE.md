# Architecture / 架构说明

## 1. Design goals / 设计目标

The converter prioritizes explainability, offline operation, source-file protection, and format preservation. It does not treat a workbook as a flat CSV: merged headers, row numbers, units, drawings, formulas, styles, named ranges, and package relationships remain relevant.

转换器优先保证可解释、离线、来源文件保护和模板保真。工作簿不是简单 CSV；合并表头、实际行号、单位、Drawing 图片、公式、样式、命名区域和压缩包关系都需要被处理。

## 2. Modules / 模块职责

| Module | Responsibility |
| --- | --- |
| `xlsx_reader.py` | Read worksheet values, merges, formulas, visibility, and drawing anchors directly from the XLSX package. |
| `header_region.py` | Score 1–5 row header candidates, propagate merged cells logically, build full paths, infer parent/child/unit/level groups. |
| `analyzer.py` | Select known fictional demo or generic analysis, combine memory, confidence reasons, examples, issues, and per-sheet recommendations. |
| `mapping_matrix.py` | Present source columns horizontally with Excel A/AA labels and mapping properties vertically. |
| `manual_mapping_ui.py` / `gui_v2.py` | Provide file, sheet, range, mapping, name-rule, conversion, reset, and result workflows. |
| `mapping_memory.py` | Persist header fingerprints and rules locally with atomic replacement. |
| `name_rules.py` | Generate names with fallback/replace/append strategies, blank cleanup, and optional deduplication. |
| `confirmed_mapping.py` | Convert user-confirmed columns into neutral `BomRow` values. |
| `transform.py` | Unit, material, dimension, weight, and fictional demo-profile transforms. |
| `template_writer.py` | Clone the output template, reuse row-5 storage styles, write rows, copy image bytes, and repair calculation/namespace relationships. |
| `verifier.py` | Validate row/image counts, template structure, Markup Compatibility prefixes, calc-chain removal, and drawing relationships. |
| `issue_display.py` | Turn technical issue codes into user-facing severity, title, explanation, location, and action. |

## 3. HeaderRegion model / 表头区域模型

`HeaderRegion` records `header_start_row`, `header_end_row`, `data_start_row`, confidence, reasons, and one `HeaderColumn` per source column. Each column retains its Excel letter, numeric index, full header path, parent, child, unit, proposed field, process type, and level-group value.

For `Size (mm)` merged across three columns with child labels `X`, `Y`, `Z`, the logical paths become:

```text
G: Size (mm) / X -> length, unit mm
H: Size (mm) / Y -> width,  unit mm
I: Size (mm) / Z -> height, unit mm
```

The reader never writes propagated values back to the source workbook.

## 4. Neutral intermediate model / 通用中间层

`BomRow` separates conversion rules from workbook layout:

- `values`: normalized text/numeric fields;
- `images`: references to original ZIP media bytes plus source anchors and optional target slots;
- `source_sheet` and `source_row`: traceability;
- `unused_fields`: non-empty values excluded by the active mapping;
- `remarks`: explicit transformation notes.

Public target fields are generic and are defined in `profiles.py`. The public repository contains only one fully fictional quick-path profile. Unknown workbooks use confirmed generic mappings.

## 5. Multi-sheet workflow / 多 Sheet 流程

Every worksheet is analyzed independently. A worksheet is recommended only when it has plausible headers, mapped fields, and data rows. Overview and empty sheets remain visible for audit but are not selected by default. Each selected worksheet produces a separate output name and report.

## 6. Images and OpenXML / 图片与 OpenXML

Images are not decoded and re-saved. Their original media bytes are copied from the source ZIP package, then new DrawingML relationships and row/column anchors are created in the output package. This preserves byte hashes and prevents quality loss. Unsupported embedded objects are not treated as ordinary images.

The writer preserves or restores worksheet namespace declarations required by `mc:Ignorable`, removes stale `calcChain.xml` relationships, and requests a full Excel recalculation on open.

## 7. Template row-5 style rule / 模板第 5 行格式规则

The public template uses row 5 as a blank style prototype and row 6 as the target header. Generated data starts at row 7. Each destination cell receives the row-5 storage style for the same column, while row-5 text, values, formulas, and drawings are not copied. Rows 1–6, widths, style definitions, named ranges, and template formulas are verified after conversion.

## 8. Local memory and privacy / 本地记忆与隐私

The mapping database stores worksheet name, header range, normalized paths, merge summaries, source-column numbers, field/process/unit/image/level settings, ignored columns, and name-rule configuration. It does not store data rows, examples, part names, identifiers, or image bytes. Writes use a temporary file followed by atomic replacement.

## 9. Adding a public demo format / 新增公开演示格式

1. Create a fixed-seed fictional workbook under `samples/simulated/`.
2. Audit its cells, media, metadata, comments, hidden sheets, links, macros, and custom XML.
3. Prefer generic recognition. Add a `SourceProfile` only if a quick-path demonstration is necessary.
4. Use generic fictional headers and values; never adapt a non-public column order.
5. Add public tests for header range, mapping, transformation, images, privacy, and output integrity.

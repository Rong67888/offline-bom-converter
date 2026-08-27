from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
SRC = PROJECT_ROOT / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from bom_converter.gui import BomConverterApp as LegacyBomConverterApp  # noqa: E402
from bom_converter.gui import main as legacy_main  # noqa: E402
from bom_converter.gui_v2 import BomConverterAppV2, main  # noqa: E402


def infer_mode(executable: str) -> str:
    executable_name = Path(executable).stem
    # Unicode escapes avoid dependence on the active Windows console code page.
    audit_markers = ("\u901a\u7528", "\u5ba1\u6838")
    return "audit" if any(marker in executable_name for marker in audit_markers) else "quick"


def write_gui_smoke_result(path: Path, mode: str, *, legacy_ui: bool = False) -> None:
    from tkinter import Tk

    root = Tk()
    app_class = LegacyBomConverterApp if legacy_ui else BomConverterAppV2
    app = app_class(root, mode)
    root.update_idletasks()
    payload = {
        "window_title": root.title(),
        "mode": app.mode.get(),
        "ui_version": "原界面" if legacy_ui else "第二版正式界面",
        "template": app.template.get(),
        "output_dir": app.output_dir.get(),
        "python_executable": sys.executable,
        "python_version": sys.version.split()[0],
    }
    root.destroy()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_manual_ui_smoke_result(source: Path, path: Path) -> None:
    """Open the packaged V2 manual dialog and record UI-only facts.

    This is intentionally read-only: no mapping rule is saved and no conversion is run.
    """

    from tkinter import Tk
    from tempfile import TemporaryDirectory

    from bom_converter.gui_v2 import ManualMappingDialogV2
    from bom_converter.mapping_memory import MappingRuleStore

    root = Tk()
    root.withdraw()
    with TemporaryDirectory(prefix="bom_converter_ui_smoke_") as temp_dir:
        dialog = ManualMappingDialogV2(
            root,
            source,
            MappingRuleStore(Path(temp_dir) / "isolated_rules.json"),
            lambda *_args: None,
        )
        root.update_idletasks()
        payload = {
            "source": str(source.resolve()),
            "sheet": dialog.sheet_name.get(),
            "header_start_row": int(dialog.header_start.get()),
            "header_end_row": int(dialog.header_end.get()),
            "data_start_row": int(dialog.data_start.get()),
            "preview_rows": len(dialog.sheet_preview.get_children()),
            "preview_columns": list(dialog.sheet_preview["columns"]),
            "mapping_rows": len(dialog.table.get_children()),
            "has_name_rule_entry": hasattr(dialog, "_open_name_rule"),
            "range_is_legacy_default": (
                int(dialog.header_start.get()),
                int(dialog.header_end.get()),
                int(dialog.data_start.get()),
            ) == (1, 1, 2),
        }
        dialog.window.destroy()
    root.destroy()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_result_ui_smoke_result(report: Path, path: Path, mode: str) -> None:
    """Exercise the packaged result page against an existing local report."""

    import os
    from tkinter import Tk

    from bom_converter.gui_v2 import RESULT_COUNTER_TITLES

    root = Tk()
    app = BomConverterAppV2(root, mode)
    opened_targets: list[str] = []
    original_startfile = getattr(os, "startfile", None)
    try:
        os.startfile = lambda target: opened_targets.append(str(Path(target).resolve()))  # type: ignore[attr-defined]
        result = app.load_result_report(report)
        root.update_idletasks()
        result_item = app.result_table.selection()[0]
        issue_rows = [
            {
                "values": list(app.issue_table.item(item, "values")),
                "tags": list(app.issue_table.item(item, "tags")),
            }
            for item in app.issue_table.get_children()
        ]
        binding = app.result_table.bind("<Double-1>")
        detail = app._open_issue_details()
        root.update_idletasks()
        detail_title = detail.title() if detail else None
        detail_rows = 0
        if detail:
            detail_tables = [child for child in detail.winfo_children()[0].winfo_children() if child.winfo_class() == "TFrame"]
            for frame in detail_tables:
                for child in frame.winfo_children():
                    if child.winfo_class() == "Treeview":
                        detail_rows = len(child.get_children())
            detail.destroy()
        app.copy_issues_button.invoke()
        root.update_idletasks()
        clipboard = root.clipboard_get()
        app.open_report_button.invoke()
        app.open_result_dir_button.invoke()
        app.return_mapping_button.invoke()
        root.update_idletasks()
        payload = {
            "report": str(report.resolve()),
            "output_file": result.output_path.name,
            "counter_titles": list(RESULT_COUNTER_TITLES),
            "counter_values": [app.success_count.get(), app.warning_count.get(), app.failure_count.get()],
            "summary": app.result_summary_text.get(),
            "selected_summary": app.selected_result_text.get(),
            "result_row": list(app.result_table.item(result_item, "values")),
            "issue_rows": issue_rows,
            "issue_codes": [issue.code for issue in result.issues],
            "double_click_bound": bool(binding),
            "detail_title": detail_title,
            "detail_rows": detail_rows,
            "button_texts": [
                app.return_mapping_button.cget("text"),
                app.open_report_button.cget("text"),
                app.copy_issues_button.cget("text"),
                app.open_result_dir_button.cget("text"),
            ],
            "clipboard_has_both_codes": "UNKNOWN_PROFILE" in clipboard and "LOW_HEADER_CONFIDENCE" in clipboard,
            "opened_targets": opened_targets,
            "returned_to_mapping": app.result_page.winfo_manager() == "" and app.workspace.winfo_manager() == "grid",
            "offline": True,
        }
    finally:
        if original_startfile is not None:
            os.startfile = original_startfile  # type: ignore[attr-defined]
        elif hasattr(os, "startfile"):
            delattr(os, "startfile")
        root.destroy()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def run_packaged_conversion(
    sources: list[Path],
    template: Path,
    output_dir: Path,
    result_json: Path | None,
) -> int:
    from bom_converter.converter import convert_many

    results = convert_many(sources, template, output_dir, mode="quick")
    payload = [result.to_dict() for result in results]
    if result_json:
        result_json.parent.mkdir(parents=True, exist_ok=True)
        result_json.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    has_error = any(
        any(issue.severity == "error" for issue in result.issues)
        for result in results
    )
    return 2 if has_error else 0


def run_packaged_sheet_conversion(source: Path, sheet_names: list[str], template: Path, output_dir: Path, result_json: Path) -> int:
    from bom_converter.analyzer import analyze_sheet
    from bom_converter.confirmed_mapping import configs_from_analysis
    from bom_converter.converter import convert_analyzed_file, convert_confirmed_file

    results = []
    for sheet_name in sheet_names:
        analysis = analyze_sheet(source, sheet_name, mode="audit")
        if analysis.profile_id == "generic":
            mappings = {item.source_col: item.universal_field for item in analysis.mappings}
            configs = configs_from_analysis(analysis, mappings)
            result = convert_confirmed_file(source, template, output_dir, analysis, mappings, column_configs=configs)
        else:
            result = convert_analyzed_file(source, template, output_dir, analysis, include_sheet_name=True)
        results.append(result.to_dict())
    result_json.parent.mkdir(parents=True, exist_ok=True)
    result_json.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    return 0


def write_task_lifecycle_smoke_result(sources: list[Path], template: Path, output_dir: Path, result_json: Path) -> None:
    """Run three isolated conversion tasks in one real V2 controller instance."""

    from tkinter import Tk
    from bom_converter.analyzer import analyze_workbook_sheets
    from bom_converter.confirmed_mapping import configs_from_analysis
    from bom_converter.converter import convert_analyzed_file, convert_confirmed_file

    root = Tk()
    root.withdraw()
    app = BomConverterAppV2(root, "audit")
    app.template.set(str(template))
    app.output_dir.set(str(output_dir))
    tasks = []
    for source in sources:
        source = source.resolve()
        app.sources.append(source)
        app.file_table.insert("", "end", iid=str(source), values=(source.name, "等待识别", "—", "—", "—", "分析中"))
        sheets = analyze_workbook_sheets(source, mode="audit")
        app._handle_sheet_analyses(source, sheets)
        converted = []
        for key in sorted(app.selected_sheets, key=lambda item: item[1]):
            item = app.sheet_analyses[key]
            if not item.analysis:
                continue
            analysis = item.analysis
            if analysis.profile_id == "generic":
                mappings = {decision.source_col: decision.universal_field for decision in analysis.mappings}
                configs = configs_from_analysis(analysis, mappings)
                result = convert_confirmed_file(source, template, output_dir, analysis, mappings, column_configs=configs)
            else:
                result = convert_analyzed_file(source, template, output_dir, analysis, include_sheet_name=True)
            app._handle_conversion(result)
            converted.append({"sheet": analysis.sheet_name, "rows": result.output_rows, "images": result.output_images})
        tasks.append({
            "source": source.name,
            "selected_sheets": sorted(key[1] for key in app.selected_sheets),
            "converted": converted,
            "result_rows": len(app._conversion_results),
        })
        app._start_new_task()
        tasks[-1]["cleared_after_task"] = not app.sources and not app.sheet_analyses and not app._conversion_results
    payload = {
        "tasks": tasks,
        "same_window": True,
        "template_preserved": app.template.get() == str(template),
        "output_dir_preserved": app.output_dir.get() == str(output_dir),
        "final_state_empty": not app.sources and not app.sheet_analyses and not app._conversion_results,
    }
    root.destroy()
    result_json.parent.mkdir(parents=True, exist_ok=True)
    result_json.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    inferred_mode = infer_mode(sys.executable)
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--mode", choices=("quick", "audit"), default=inferred_mode)
    parser.add_argument("--smoke-test", type=Path)
    parser.add_argument("--manual-ui-smoke-source", type=Path)
    parser.add_argument("--manual-ui-smoke-json", type=Path)
    parser.add_argument("--result-ui-smoke-report", type=Path)
    parser.add_argument("--result-ui-smoke-json", type=Path)
    parser.add_argument("--legacy-ui", action="store_true")
    parser.add_argument("--convert-source", type=Path, action="append", default=[])
    parser.add_argument("--template", type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--result-json", type=Path)
    parser.add_argument("--sheet-source", type=Path)
    parser.add_argument("--convert-sheet", action="append", default=[])
    parser.add_argument("--sheet-result-json", type=Path)
    parser.add_argument("--task-source", type=Path, action="append", default=[])
    parser.add_argument("--task-lifecycle-smoke-json", type=Path)
    args, _ = parser.parse_known_args()
    if args.result_ui_smoke_report:
        if not args.result_ui_smoke_json:
            raise SystemExit("Result UI smoke test requires --result-ui-smoke-json.")
        write_result_ui_smoke_result(args.result_ui_smoke_report, args.result_ui_smoke_json, args.mode)
    elif args.manual_ui_smoke_source:
        if not args.manual_ui_smoke_json:
            raise SystemExit("Manual UI smoke test requires --manual-ui-smoke-json.")
        write_manual_ui_smoke_result(args.manual_ui_smoke_source, args.manual_ui_smoke_json)
    elif args.smoke_test:
        write_gui_smoke_result(args.smoke_test, args.mode, legacy_ui=args.legacy_ui)
    elif args.convert_source:
        if not args.template or not args.output_dir:
            raise SystemExit(
                "Packaged conversion test requires both --template and --output-dir."
            )
        raise SystemExit(
            run_packaged_conversion(
                args.convert_source,
                args.template,
                args.output_dir,
                args.result_json,
            )
        )
    elif args.sheet_source and args.convert_sheet:
        if not args.template or not args.output_dir or not args.sheet_result_json:
            raise SystemExit("Sheet conversion test requires --template, --output-dir and --sheet-result-json.")
        raise SystemExit(run_packaged_sheet_conversion(
            args.sheet_source,
            args.convert_sheet,
            args.template,
            args.output_dir,
            args.sheet_result_json,
        ))
    elif args.task_source:
        if not args.template or not args.output_dir or not args.task_lifecycle_smoke_json:
            raise SystemExit("Task lifecycle smoke requires --template, --output-dir and --task-lifecycle-smoke-json.")
        write_task_lifecycle_smoke_result(args.task_source, args.template, args.output_dir, args.task_lifecycle_smoke_json)
    else:
        (legacy_main if args.legacy_ui else main)(args.mode)

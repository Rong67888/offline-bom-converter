from __future__ import annotations

import argparse
import json
import os
import queue
import threading
from pathlib import Path
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    BooleanVar,
    StringVar,
    Text,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
)
from tkinter import ttk

from .analyzer import analyze_manual_workbook, analyze_sheet, analyze_workbook, analyze_workbook_sheets
from .confirmed_mapping import (
    FIELD_LABELS,
    PROCESS_LABELS,
    configs_from_analysis,
    infer_confirmed_units,
    transform_confirmed_workbook,
)
from .converter import convert_analyzed_file, convert_confirmed_file
from .gui import BomConverterApp, _install_root, _resource_root
from .header_range_preview import header_candidates, load_sheet_preview, validate_header_range
from .issue_display import (
    conversion_result_from_report,
    count_results,
    issue_summary_text,
    present_issue,
)
from .manual_mapping_ui import ManualMappingDialog
from .mapping_matrix import MatrixColumn, MappingMatrix
from .models import ColumnMappingConfig, ConversionResult, Issue, NameGenerationRule, SheetAnalysis, WorkbookAnalysis
from .name_rule_ui import NameRuleDialog
from .name_rules import infer_name_rule_columns
from .profiles import TARGET_HEADERS
from .xlsx_reader import XlsxReader


COLORS = {
    "navy": "#18324A",
    "primary": "#2F6F9F",
    "primary_dark": "#24597E",
    "canvas": "#EEF3F7",
    "surface": "#FFFFFF",
    "line": "#D7E1E8",
    "text": "#1F3446",
    "muted": "#607487",
    "soft_blue": "#E8F1F8",
    "success": "#1F7A4D",
    "success_bg": "#E8F6EE",
    "warning": "#946200",
    "warning_bg": "#FFF4D6",
    "danger": "#B42318",
    "danger_bg": "#FDECEA",
}


STATUS_TEXT = {
    "mapped": "已匹配",
    "ignored": "已忽略",
    "review": "需要确认",
}

RESULT_COUNTER_TITLES = ("完全成功文件", "含警告文件", "失败文件")


V2_FIELD_LABELS = {
    "category": "分类",
    "sequence": "序号",
    "part_name": "零件名称",
    "part_number": "零件号",
    "level": "层级",
    "electronics_spec": "元器件型号规格",
    "electronics_silk": "元器件丝印",
    "electronics_package": "元器件封装",
    "pin_number": "元器件引脚数",
    "pcb_side": "PCB 正反面",
    "electronics_type": "元器件类型",
    "unit_weight": "单件重量",
    "quantity": "数量",
    "total_weight": "总重量",
    "material_type": "材料种类",
    "material_spec": "材料牌号 / 规格",
    "length": "长度",
    "width": "宽度",
    "height": "高度",
    "unfold_length": "展开长度",
    "diameter": "直径",
    "thickness": "厚度",
    "production_process": "生产工艺",
    "assembly_process": "装配工艺",
    "surface_treatment": "表面处理",
    "surface_area": "表面处理面积",
    "manufacturer": "供应商 / 制造商",
    "location": "产地",
    "remark": "备注",
    "code": "代码",
    "vpc": "整车分解代码",
    "dimensions": "尺寸字符串",
    "material_raw": "原始材料",
    "standard_name": "标准件名称",
    "standard_name_gb": "GB 名称",
    "spec": "规格",
    "english_name": "英文名称",
}


def confidence_text(value: float, *, confirmed: bool = False) -> str:
    if confirmed:
        return "已人工确认"
    if value >= 0.85:
        return "高"
    if value >= 0.65:
        return "中"
    return "低"


def recommendation_source(analysis: WorkbookAnalysis) -> str:
    if analysis.profile_id != "generic":
        return "已知格式"
    if any(issue.code == "SAVED_RULE_APPLIED" for issue in analysis.issues):
        return "历史规则"
    return "自动推荐"


def header_region_text(analysis: WorkbookAnalysis) -> str:
    if analysis.header_start_row == analysis.header_end_row:
        header = f"第 {analysis.header_start_row} 行"
    else:
        header = f"第 {analysis.header_start_row}–{analysis.header_end_row} 行"
    return f"{header}；数据第 {analysis.data_start_row} 行起"


def _build_unique_field_options() -> tuple[list[str], dict[str, str | None], dict[str, str]]:
    options = ["忽略"]
    display_to_field: dict[str, str | None] = {"忽略": None}
    field_to_display: dict[str, str] = {}
    used: dict[str, int] = {}
    for field in sorted(FIELD_LABELS):
        label = V2_FIELD_LABELS.get(field, FIELD_LABELS.get(field, field))
        used[label] = used.get(label, 0) + 1
        display = label if used[label] == 1 else f"{label}（选项 {used[label]}）"
        options.append(display)
        display_to_field[display] = field
        field_to_display[field] = display
    return options, display_to_field, field_to_display


FIELD_OPTIONS_V2, DISPLAY_TO_FIELD_V2, FIELD_TO_DISPLAY_V2 = _build_unique_field_options()
PROCESS_OPTIONS_V2 = list(PROCESS_LABELS.values())
DISPLAY_TO_PROCESS_V2 = {label: key for key, label in PROCESS_LABELS.items()}
PROCESS_TO_DISPLAY_V2 = dict(PROCESS_LABELS)
UNIT_OPTIONS = ["", "kg", "g", "mg", "mm", "cm", "m", "mm²", "cm²", "m²"]
SLOT_OPTIONS = [""] + [str(value) for value in range(1, 8)]


class ManualMappingDialogV2(ManualMappingDialog):
    """第二版手动配置窗口；业务读写仍调用现有 ManualMappingDialog 方法。"""

    def __init__(self, *args: object, **kwargs: object) -> None:
        parent = args[0] if args else kwargs.get("parent")
        self.internal_detail = StringVar(master=parent, value="内部字段：尚未选择列")
        self.range_hint = StringVar(master=parent, value="正在读取自动建议和工作表预览……")
        self._range_diagnostics = []
        self._candidate_regions = []
        super().__init__(*args, **kwargs)
        self.process.set(PROCESS_TO_DISPLAY_V2.get("ignore", "忽略"))
        self.field.set("忽略")
        self.window.title(f"手动配置格式 · UI 第二版 — {self.source.name}")
        self.window.geometry("1280x700")
        self.window.minsize(1020, 600)
        self.window.transient(self.window.master)

    def _build(self, sheet_names: list[str]) -> None:
        self.window.configure(background=COLORS["canvas"])
        outer = ttk.Frame(self.window, style="App.TFrame", padding=12)
        outer.pack(fill=BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        header = ttk.Frame(outer, style="Header.TFrame", padding=(14, 7))
        header.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="手动配置格式", style="HeaderTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            header,
            text="仅保存表头结构和处理规则，不保存示例值、零件数据或图片",
            style="HeaderHint.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        region = ttk.LabelFrame(outer, text="1  选择工作表和数据区域", style="Card.TLabelframe", padding=10)
        region.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        region.columnconfigure(8, weight=1)
        ttk.Label(region, text="工作表").grid(row=0, column=0, sticky="w")
        sheet_box = ttk.Combobox(region, textvariable=self.sheet_name, values=sheet_names, state="readonly", width=24)
        sheet_box.grid(row=0, column=1, padx=(5, 12))
        sheet_box.bind("<<ComboboxSelected>>", self._sheet_changed)
        for index, (label, variable) in enumerate(
            (("表头开始行", self.header_start), ("表头结束行", self.header_end), ("数据开始行", self.data_start))
        ):
            column = 2 + index * 2
            ttk.Label(region, text=label).grid(row=0, column=column, sticky="e", padx=(6, 3))
            entry = ttk.Entry(region, textvariable=variable, width=7)
            entry.grid(row=0, column=column + 1, sticky="w")
            entry.bind("<Return>", lambda _event: self._analyze())
            entry.bind("<FocusOut>", lambda _event: self._analyze())
        ttk.Button(region, text="读取并预览", command=self._analyze, style="Primary.TButton").grid(row=0, column=8, sticky="e", padx=(12, 0))
        range_actions = ttk.Frame(region)
        range_actions.grid(row=1, column=0, columnspan=9, sticky="ew", pady=(7, 0))
        ttk.Button(range_actions, text="接受自动建议", command=self._accept_automatic_suggestion, style="Primary.TButton").pack(side=LEFT)
        ttk.Button(range_actions, text="查看其他候选", command=self._show_candidates).pack(side=LEFT, padx=5)
        ttk.Button(range_actions, text="恢复自动范围", command=self._restore_automatic_range).pack(side=LEFT)
        ttk.Label(range_actions, textvariable=self.range_hint, style="Muted.TLabel").pack(side=RIGHT)
        ttk.Label(region, textvariable=self.reason, style="Muted.TLabel", wraplength=1180).grid(row=2, column=0, columnspan=9, sticky="w", pady=(6, 0))

        body = ttk.Frame(outer, style="App.TFrame")
        body.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self.manual_notebook = ttk.Notebook(body)
        self.manual_notebook.grid(row=0, column=0, sticky="nsew")
        range_tab = ttk.Frame(self.manual_notebook, padding=8)
        mapping_tab = ttk.Frame(self.manual_notebook, padding=8)
        self.manual_notebook.add(range_tab, text="范围预览（前 50 行）")
        self.manual_notebook.add(mapping_tab, text="字段与名称设置")

        range_tab.columnconfigure(0, weight=1)
        range_tab.rowconfigure(0, weight=1)
        preview_card = ttk.LabelFrame(range_tab, text="Excel 行号与列字母只读预览", padding=7)
        preview_card.grid(row=0, column=0, sticky="nsew")
        preview_card.columnconfigure(0, weight=1)
        preview_card.rowconfigure(0, weight=1)
        self.sheet_preview = ttk.Treeview(preview_card, columns=(), show="headings", height=10, selectmode="browse")
        preview_y = ttk.Scrollbar(preview_card, orient="vertical", command=self.sheet_preview.yview)
        preview_x = ttk.Scrollbar(preview_card, orient="horizontal", command=self.sheet_preview.xview)
        self.sheet_preview.configure(yscrollcommand=preview_y.set, xscrollcommand=preview_x.set)
        self.sheet_preview.grid(row=0, column=0, sticky="nsew")
        preview_y.grid(row=0, column=1, sticky="ns")
        preview_x.grid(row=1, column=0, sticky="ew")
        self.sheet_preview.tag_configure("title", background="#EEF3F7", foreground=COLORS["muted"])
        self.sheet_preview.tag_configure("header", background=COLORS["soft_blue"], foreground=COLORS["navy"])
        self.sheet_preview.tag_configure("data", background=COLORS["success_bg"], foreground=COLORS["success"])
        self.sheet_preview.tag_configure("problem", background=COLORS["danger_bg"], foreground=COLORS["danger"])
        self.sheet_preview.tag_configure("blank", background="#F8FAFC", foreground="#94A3B8")
        row_actions = ttk.Frame(range_tab)
        row_actions.grid(row=1, column=0, sticky="ew", pady=(7, 0))
        ttk.Label(row_actions, text="先在表格中点一行，再设置：").pack(side=LEFT)
        ttk.Button(row_actions, text="设为表头开始行", command=lambda: self._set_selected_preview_row("header_start")).pack(side=LEFT, padx=4)
        ttk.Button(row_actions, text="设为表头结束行", command=lambda: self._set_selected_preview_row("header_end")).pack(side=LEFT, padx=4)
        ttk.Button(row_actions, text="设为数据开始行", command=lambda: self._set_selected_preview_row("data_start")).pack(side=LEFT, padx=4)
        ttk.Label(
            row_actions,
            text="蓝色=表头；绿色=数据开始；红色=范围问题；灰色=标题或空白",
            style="Muted.TLabel",
        ).pack(side=RIGHT)

        mapping_tab.columnconfigure(0, weight=1)
        mapping_tab.rowconfigure(0, weight=1)
        table_card = ttk.LabelFrame(mapping_tab, text="2  检查完整表头并逐列设置", style="Card.TLabelframe", padding=8)
        table_card.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        table_card.columnconfigure(0, weight=1)
        table_card.rowconfigure(0, weight=1)
        self.table = MappingMatrix(
            table_card,
            on_edit=self._open_manual_field_editor,
            on_select=self._load_selected,
            height=340,
        )
        self.table.grid(row=0, column=0, sticky="nsew")

        editor = ttk.LabelFrame(mapping_tab, text="3  编辑所选列", style="Card.TLabelframe", padding=10)
        editor.grid(row=0, column=1, sticky="nsew")
        editor.columnconfigure(1, weight=1)
        controls = (
            ("处理方式", self.process, PROCESS_OPTIONS_V2),
            ("识别为", self.field, FIELD_OPTIONS_V2),
            ("来源单位", self.unit, UNIT_OPTIONS),
            ("图片槽位", self.image_slot, SLOT_OPTIONS),
        )
        for row, (label, variable, values) in enumerate(controls):
            ttk.Label(editor, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Combobox(editor, textvariable=variable, values=values, state="readonly", width=24).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Label(editor, text="层级组").grid(row=4, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.level_group).grid(row=4, column=1, sticky="ew", pady=4)
        ttk.Label(editor, text="层级数字").grid(row=5, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.level_value).grid(row=5, column=1, sticky="ew", pady=4)
        ttk.Label(editor, text="固定默认值").grid(row=6, column=0, sticky="w", pady=4)
        ttk.Entry(editor, textvariable=self.default_value).grid(row=6, column=1, sticky="ew", pady=4)
        ttk.Label(editor, textvariable=self.internal_detail, style="Muted.TLabel", wraplength=220).grid(row=7, column=0, columnspan=2, sticky="w", pady=(8, 5))
        ttk.Button(editor, text="应用到所选列", command=self._apply_selected, style="Secondary.TButton").grid(row=8, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(editor, text="名称生成规则……", command=self._open_name_rule, style="Primary.TButton").grid(row=9, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        actions = ttk.Frame(outer, style="Card.TFrame", padding=(10, 8))
        actions.grid(row=3, column=0, sticky="ew")
        ttk.Button(actions, text="预览前五行", command=self._preview_rows, style="Secondary.TButton").pack(side=LEFT)
        ttk.Button(actions, text="仅保存规则", command=self._save_only, style="Secondary.TButton").pack(side=LEFT, padx=6)
        ttk.Button(actions, text="转换并保存规则", command=lambda: self._convert(True), style="Primary.TButton").pack(side=RIGHT)
        ttk.Button(actions, text="转换一次", command=lambda: self._convert(False), style="Secondary.TButton").pack(side=RIGHT, padx=6)

        self._candidate_regions = header_candidates(self.source)
        self._refresh_sheet_preview()

    def _current_range_diagnostics(self):
        try:
            start, end, data = self._parse_rows()
            with XlsxReader(self.source) as reader:
                sheet = reader.read_sheet(self.sheet_name.get())
            return validate_header_range(sheet, start, end, data)
        except Exception as exc:
            from .header_range_preview import RangeDiagnostic

            return [RangeDiagnostic("error", "INVALID_RANGE", str(exc))]

    def _refresh_sheet_preview(self) -> None:
        if not hasattr(self, "sheet_preview"):
            return
        try:
            preview = load_sheet_preview(self.source, self.sheet_name.get(), max_rows=50, max_cols=24)
            start, end, data = self._parse_rows()
        except Exception as exc:
            self.range_hint.set(f"无法读取预览：{exc}")
            return
        diagnostics = self._current_range_diagnostics()
        self._range_diagnostics = diagnostics
        problem_rows = {item.row for item in diagnostics if item.row}
        columns = ("row", *preview.column_letters)
        self.sheet_preview.configure(columns=columns)
        self.sheet_preview.heading("row", text="行号")
        self.sheet_preview.column("row", width=54, minwidth=54, stretch=False, anchor="center")
        for letter in preview.column_letters:
            self.sheet_preview.heading(letter, text=letter)
            self.sheet_preview.column(letter, width=115, minwidth=70, stretch=False)
        for item in self.sheet_preview.get_children():
            self.sheet_preview.delete(item)
        for row_number, values in preview.rows:
            rendered = []
            for value in values:
                text = "" if value is None else str(value).replace("\r", " ").replace("\n", " / ")
                rendered.append(text if len(text) <= 60 else text[:57] + "...")
            if row_number in problem_rows:
                tag = "problem"
            elif row_number == data:
                tag = "data"
            elif start <= row_number <= end:
                tag = "header"
            elif row_number < start:
                tag = "title"
            elif not any(rendered):
                tag = "blank"
            else:
                tag = ""
            self.sheet_preview.insert("", END, iid=str(row_number), values=(row_number, *rendered), tags=(tag,) if tag else ())
        errors = sum(item.severity == "error" for item in diagnostics)
        warnings = sum(item.severity == "warning" for item in diagnostics)
        self.range_hint.set(
            f"显示前 {min(preview.max_row, 50)} 行、{len(preview.column_letters)} 列；范围问题 {errors} 个，提醒 {warnings} 个"
        )

    def _set_selected_preview_row(self, target: str) -> None:
        selected = self.sheet_preview.selection()
        if not selected:
            messagebox.showwarning("未选择行", "请先在预览表格中选择一行。", parent=self.window)
            return
        row = selected[0]
        variable = {
            "header_start": self.header_start,
            "header_end": self.header_end,
            "data_start": self.data_start,
        }[target]
        variable.set(row)
        self._analyze()

    def _sheet_changed(self, _event: object = None) -> None:
        candidates = [item for item in self._candidate_regions if item.sheet_name == self.sheet_name.get()]
        if candidates:
            best = candidates[0]
            self.header_start.set(str(best.header_start_row))
            self.header_end.set(str(best.header_end_row))
            self.data_start.set(str(best.data_start_row))
        self._analyze()

    def _restore_automatic_range(self) -> None:
        sheet, start, end, data = self.automatic_range
        self.sheet_name.set(sheet)
        self.header_start.set(str(start))
        self.header_end.set(str(end))
        self.data_start.set(str(data))
        self._refresh_sheet_preview()
        self.reason.set("已恢复程序最初的自动范围；点击“接受自动建议”可重新识别字段。")

    def _accept_automatic_suggestion(self) -> None:
        self._restore_automatic_range()
        self._analyze()
        if self.analysis:
            self.manual_notebook.select(1)

    def _show_candidates(self) -> None:
        popup = Toplevel(self.window)
        popup.title("其他表头候选")
        popup.geometry("900x480")
        popup.transient(self.window)
        outer = ttk.Frame(popup, padding=10)
        outer.pack(fill=BOTH, expand=True)
        ttk.Label(outer, text="候选按识别把握排序。选择一项后仍应检查蓝色表头区域。", style="Muted.TLabel").pack(fill="x", pady=(0, 6))
        columns = ("sheet", "header", "data", "confidence", "reason")
        table = ttk.Treeview(outer, columns=columns, show="headings", selectmode="browse")
        for name, title, width in (
            ("sheet", "工作表", 180),
            ("header", "表头范围", 110),
            ("data", "数据开始", 90),
            ("confidence", "识别把握", 85),
            ("reason", "主要理由", 400),
        ):
            table.heading(name, text=title)
            table.column(name, width=width, stretch=name == "reason")
        table.pack(fill=BOTH, expand=True)
        candidates = self._candidate_regions[:20]
        for index, candidate in enumerate(candidates):
            table.insert(
                "",
                END,
                iid=str(index),
                values=(
                    candidate.sheet_name,
                    f"{candidate.header_start_row}–{candidate.header_end_row}",
                    candidate.data_start_row,
                    confidence_text(candidate.confidence),
                    "；".join(candidate.confidence_reasons[:3]),
                ),
            )

        def use_selected() -> None:
            selected = table.selection()
            if not selected:
                messagebox.showwarning("未选择候选", "请先选择一项。", parent=popup)
                return
            candidate = candidates[int(selected[0])]
            self.sheet_name.set(candidate.sheet_name)
            self.header_start.set(str(candidate.header_start_row))
            self.header_end.set(str(candidate.header_end_row))
            self.data_start.set(str(candidate.data_start_row))
            popup.destroy()
            self._analyze()

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(7, 0))
        ttk.Button(actions, text="关闭", command=popup.destroy).pack(side=RIGHT)
        ttk.Button(actions, text="使用所选候选", command=use_selected, style="Primary.TButton").pack(side=RIGHT, padx=6)
        if candidates:
            table.selection_set("0")

    def _analyze(self) -> None:
        diagnostics = self._current_range_diagnostics()
        errors = [item for item in diagnostics if item.severity == "error"]
        if errors:
            self.analysis = None
            self.configs = {}
            self._refresh_table()
            self.reason.set("范围不能使用：" + "；".join(item.message for item in errors))
            self._refresh_sheet_preview()
            return
        super()._analyze()
        if self.analysis:
            self.name_rule = infer_name_rule_columns(
                self.analysis.mappings,
                self.configs,
                self.analysis.name_rule or self.name_rule,
            )
            self.analysis.name_rule = self.name_rule
            reasons = "；".join(self.analysis.header_region.confidence_reasons)
            warnings = [item.message for item in diagnostics if item.severity == "warning"]
            suffix = ("；范围提醒：" + "；".join(warnings)) if warnings else ""
            self.reason.set(
                f"表头第 {self.analysis.header_start_row}–{self.analysis.header_end_row} 行，"
                f"数据从第 {self.analysis.data_start_row} 行开始；识别把握：{confidence_text(self.analysis.profile_confidence)}。{reasons}{suffix}"
            )
        self._refresh_sheet_preview()

    def _open_name_rule(self) -> None:
        if not self.analysis:
            messagebox.showwarning("尚未识别表头", "请先选择有效范围并重新识别。", parent=self.window)
            return
        self.name_rule_dialog = NameRuleDialog(
            self.window,
            self.source,
            self.analysis,
            self.configs,
            self.name_rule,
            self._apply_name_rule,
        )

    def _apply_name_rule(self, rule: NameGenerationRule) -> None:
        self.name_rule = rule
        if self.analysis:
            self.analysis.name_rule = rule
        self.reason.set("名称生成规则已应用到当前格式；转换并保存规则后，下次导入会自动复用。")

    def _refresh_table(self) -> None:
        if not self.analysis:
            self.table.set_columns([])
            return
        columns: list[MatrixColumn] = []
        for decision in self.analysis.mappings:
            config = self.configs[decision.source_col]
            status = "已忽略" if config.process_type == "ignore" else ("需要确认" if decision.status == "review" else "已匹配")
            sample = "—" if decision.sample_value is None else str(decision.sample_value)
            columns.append(MatrixColumn(str(decision.source_col), f"{decision.column_letter}列", {
                "path": decision.source_header,
                "sample": sample[:80],
                "field": FIELD_TO_DISPLAY_V2.get(config.universal_field or "", "忽略"),
                "target": TARGET_HEADERS.get(config.universal_field or "", "—").replace("\n", " / "),
                "confidence": confidence_text(decision.confidence),
                "status": status,
                "recommendation": "自动推荐" if decision.confidence < 1 else "历史规则",
            }))
        self.table.set_columns(columns)

    def _open_manual_field_editor(self, key: str) -> None:
        config = self.configs[int(key)]
        current = FIELD_TO_DISPLAY_V2.get(config.universal_field or "", "忽略")
        self.table.open_field_editor(key, FIELD_OPTIONS_V2, current, lambda display: self._commit_manual_field(key, display))

    def _commit_manual_field(self, key: str, display: str) -> None:
        column = int(key)
        field = DISPLAY_TO_FIELD_V2.get(display)
        config = self.configs[column]
        config.universal_field = field
        config.process_type = "ignore" if field is None else ("direct" if config.process_type == "ignore" else config.process_type)
        self._refresh_table()
        self.table.selection_set(key)
        self._load_selected()

    def _load_selected(self, _event: object = None) -> None:
        selected = self.table.selection()
        if not selected:
            return
        config = self.configs[int(selected[0])]
        self.process.set(PROCESS_TO_DISPLAY_V2.get(config.process_type, "忽略"))
        self.field.set(FIELD_TO_DISPLAY_V2.get(config.universal_field or "", "忽略"))
        self.unit.set(config.unit or "")
        self.image_slot.set(str(config.image_slot or ""))
        self.level_group.set(config.level_group or "component_asm_level")
        self.level_value.set(str(config.level_value or ""))
        self.default_value.set("" if config.default_value is None else str(config.default_value))
        self.internal_detail.set(f"内部字段：{config.universal_field or '—'}\n仅用于维护定位，不写入主表格。")

    def _apply_selected(self) -> None:
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("未选择列", "请先在左侧表格中选择一列。", parent=self.window)
            return
        column = int(selected[0])
        process = DISPLAY_TO_PROCESS_V2.get(self.process.get(), "ignore")
        field = DISPLAY_TO_FIELD_V2.get(self.field.get())
        try:
            slot = int(self.image_slot.get()) if self.image_slot.get() else None
            level_value = int(self.level_value.get()) if self.level_value.get() else None
        except ValueError:
            messagebox.showerror("设置错误", "图片槽位和层级数字必须是整数。", parent=self.window)
            return
        if process == "ignore":
            field = None
        elif process == "level_group":
            field = "level"
        elif process == "image":
            field = "image"
        self.configs[column] = ColumnMappingConfig(
            source_col=column,
            process_type=process,
            universal_field=field,
            unit=self.unit.get() or None,
            image_slot=slot,
            level_group=self.level_group.get().strip() or None,
            level_value=level_value,
            default_value=self.default_value.get() if process == "fixed_default" else None,
        )
        self._refresh_table()
        self.table.selection_set(str(column))
        self._load_selected()
        if self.analysis:
            self.name_rule = infer_name_rule_columns(self.analysis.mappings, self.configs, self.name_rule)
            self.analysis.name_rule = self.name_rule


class BomConverterAppV2(BomConverterApp):
    """独立 UI 第二版。继承现有控制器，只替换界面编排和用户文字。"""

    def __init__(self, root: Tk, default_mode: str = "audit", preview_state: str | None = None) -> None:
        self.only_review = BooleanVar(master=root, value=False)
        self.search_text = StringVar(master=root, value="")
        self.detail_text = StringVar(master=root, value="选择一行可查看内部字段和推荐说明。")
        self.mode_caption = StringVar(master=root, value="")
        self.settings_summary_text = StringVar(master=root, value="输出位置已就绪")
        self.log_visible = False
        self._mapping_rows: dict[str, tuple[str, ...]] = {}
        self._mapping_internal: dict[str, str] = {}
        self._manual_confirmed: dict[Path, set[int]] = {}
        self.sheet_analyses: dict[tuple[Path, str], SheetAnalysis] = {}
        self.selected_sheets: set[tuple[Path, str]] = set()
        self.sheet_mapping_choices: dict[tuple[Path, str], dict[int, str | None]] = {}
        self.sheet_manual_configs: dict[tuple[Path, str], dict[int, ColumnMappingConfig]] = {}
        self.sheet_confirmed: dict[tuple[Path, str], set[int]] = {}
        self.current_sheet_key: tuple[Path, str] | None = None
        self._conversion_results: list[ConversionResult] = []
        self._result_by_item: dict[str, ConversionResult] = {}
        self.result_summary_text = StringVar(master=root, value="尚无转换结果。")
        self.selected_result_text = StringVar(master=root, value="选择一个输出文件，即可在这里直接查看完整问题列表。")
        self.last_issue_detail_window: Toplevel | None = None
        self.preview_dialog: ManualMappingDialogV2 | None = None
        super().__init__(root, default_mode)
        self.root.title("BOM 格式转换工具（离线版）")
        self.root.geometry("1280x720+20+20")
        self.root.minsize(1040, 640)
        self.mode.trace_add("write", lambda *_args: self._on_mode_change())
        self.search_text.trace_add("write", lambda *_args: self._refresh_mapping_view())
        self.template.trace_add("write", lambda *_args: self._update_settings_summary())
        self.output_dir.trace_add("write", lambda *_args: self._update_settings_summary())
        self._update_settings_summary()
        self._on_mode_change()
        if preview_state:
            self.load_preview_state(preview_state)

    def _build_style(self) -> None:
        self.root.configure(background=COLORS["canvas"])
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        font = ("Microsoft YaHei UI", 9)
        style.configure(".", font=font)
        style.configure("App.TFrame", background=COLORS["canvas"])
        style.configure("Card.TFrame", background=COLORS["surface"])
        style.configure("Header.TFrame", background=COLORS["navy"])
        style.configure("TLabel", background=COLORS["canvas"], foreground=COLORS["text"])
        style.configure("Card.TLabel", background=COLORS["surface"], foreground=COLORS["text"])
        style.configure("Muted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"])
        style.configure("HeaderTitle.TLabel", background=COLORS["navy"], foreground="#FFFFFF", font=("Microsoft YaHei UI", 15, "bold"))
        style.configure("HeaderHint.TLabel", background=COLORS["navy"], foreground="#D7E5F0", font=("Microsoft YaHei UI", 9))
        style.configure("Offline.TLabel", background="#274A66", foreground="#EAF4FB", padding=(10, 5))
        style.configure("Section.TLabel", background=COLORS["surface"], foreground=COLORS["navy"], font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Primary.TButton", background=COLORS["primary"], foreground="#FFFFFF", padding=(14, 7), borderwidth=0, font=("Microsoft YaHei UI", 9, "bold"))
        style.map("Primary.TButton", background=[("active", COLORS["primary_dark"]), ("disabled", "#9CB1C1")])
        style.configure("Secondary.TButton", background="#F5F8FA", foreground=COLORS["text"], padding=(12, 7), borderwidth=1)
        style.map("Secondary.TButton", background=[("active", COLORS["soft_blue"])])
        style.configure("Danger.TButton", background=COLORS["danger_bg"], foreground=COLORS["danger"], padding=(12, 7))
        style.configure("TEntry", padding=6, fieldbackground="#FFFFFF")
        style.configure("TCombobox", padding=5, fieldbackground="#FFFFFF")
        style.configure("Card.TLabelframe", background=COLORS["surface"], bordercolor=COLORS["line"], relief="solid")
        style.configure("Card.TLabelframe.Label", background=COLORS["surface"], foreground=COLORS["navy"], font=("Microsoft YaHei UI", 10, "bold"))
        style.configure("Treeview", rowheight=29, background="#FFFFFF", fieldbackground="#FFFFFF", foreground=COLORS["text"], bordercolor=COLORS["line"], borderwidth=1)
        style.configure("Treeview.Heading", background="#DCE7EF", foreground=COLORS["navy"], font=("Microsoft YaHei UI", 9, "bold"), relief="flat")
        style.map("Treeview", background=[("selected", "#CFE3F2")], foreground=[("selected", COLORS["navy"])])
        style.configure("Step.TLabel", background="#DCE7EF", foreground=COLORS["muted"], padding=(8, 5), font=("Microsoft YaHei UI", 8, "bold"))
        style.configure("StepActive.TLabel", background=COLORS["primary"], foreground="#FFFFFF", padding=(8, 5), font=("Microsoft YaHei UI", 8, "bold"))
        for kind, foreground, background in (
            ("Success", COLORS["success"], COLORS["success_bg"]),
            ("Warning", COLORS["warning"], COLORS["warning_bg"]),
            ("Danger", COLORS["danger"], COLORS["danger_bg"]),
        ):
            style.configure(f"Counter{kind}.TFrame", background=background)
            style.configure(f"Counter{kind}Value.TLabel", background=background, foreground=foreground, font=("Microsoft YaHei UI", 22, "bold"))
            style.configure(f"Counter{kind}Text.TLabel", background=background, foreground=foreground, font=("Microsoft YaHei UI", 9, "bold"))

    def _build(self) -> None:
        outer = ttk.Frame(self.root, style="App.TFrame", padding=12)
        outer.pack(fill=BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(2, weight=1)

        header = ttk.Frame(outer, style="Header.TFrame", padding=(16, 8))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)
        ttk.Label(header, text="BOM 格式转换工具", style="HeaderTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="添加文件 → 检查识别结果 → 必要时确认字段 → 选择位置 → 转换 → 查看结果", style="HeaderHint.TLabel").grid(row=1, column=0, sticky="w", pady=(3, 0))
        ttk.Label(header, text="所有文件仅在本机离线处理", style="Offline.TLabel").grid(row=0, column=1, rowspan=2, sticky="e", padx=(16, 0))

        nav = ttk.Frame(outer, style="Card.TFrame", padding=(10, 7))
        nav.grid(row=1, column=0, sticky="ew", pady=(8, 8))
        ttk.Label(nav, text="工作模式", style="Card.TLabel").pack(side=LEFT, padx=(0, 8))
        ttk.Radiobutton(nav, text="快速模式", variable=self.mode, value="quick").pack(side=LEFT, padx=4)
        ttk.Radiobutton(nav, text="通用审核模式", variable=self.mode, value="audit").pack(side=LEFT, padx=4)
        ttk.Label(nav, textvariable=self.mode_caption, style="Muted.TLabel").pack(side=LEFT, padx=12)
        steps = ttk.Frame(nav, style="Card.TFrame")
        steps.pack(side=RIGHT)
        self.step_labels: list[ttk.Label] = []
        for index, text in enumerate(("1 添加", "2 检查", "3 确认", "4 位置", "5 转换", "6 结果"), 1):
            label = ttk.Label(steps, text=text, style="StepActive.TLabel" if index == 1 else "Step.TLabel")
            label.pack(side=LEFT, padx=2)
            self.step_labels.append(label)

        self.content = ttk.Frame(outer, style="App.TFrame")
        self.content.grid(row=2, column=0, sticky="nsew")
        self.content.columnconfigure(0, weight=1)
        self.content.rowconfigure(0, weight=1)

        self.workspace = ttk.Frame(self.content, style="App.TFrame")
        self.workspace.grid(row=0, column=0, sticky="nsew")
        self.workspace.columnconfigure(0, weight=1)
        self.workspace.rowconfigure(1, weight=1)

        files = ttk.LabelFrame(self.workspace, text="1  添加 BOM 文件并检查识别结果", style="Card.TLabelframe", padding=8)
        files.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        files.columnconfigure(0, weight=1)
        toolbar = ttk.Frame(files, style="Card.TFrame")
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Label(toolbar, text="文件列表", style="Section.TLabel").pack(side=LEFT)
        ttk.Button(toolbar, text="添加文件", command=self._choose_sources, style="Primary.TButton").pack(side=RIGHT)
        ttk.Button(toolbar, text="重新检查", command=self._reanalyze_selected, style="Secondary.TButton").pack(side=RIGHT, padx=6)
        ttk.Button(toolbar, text="清空", command=self._clear_sources, style="Secondary.TButton").pack(side=RIGHT)
        file_columns = ("file", "profile", "header", "rows", "images", "status")
        self.file_table = ttk.Treeview(files, columns=file_columns, show="headings", height=3, selectmode="browse")
        for column, title, width in (
            ("file", "文件名", 300),
            ("profile", "识别格式", 210),
            ("header", "表头区域", 190),
            ("rows", "数据行", 72),
            ("images", "图片数", 72),
            ("status", "状态", 210),
        ):
            self.file_table.heading(column, text=title)
            self.file_table.column(column, width=width, minwidth=65, anchor="center" if column in {"rows", "images"} else "w")
        fy = ttk.Scrollbar(files, orient="vertical", command=self.file_table.yview)
        fx = ttk.Scrollbar(files, orient="horizontal", command=self.file_table.xview)
        self.file_table.configure(yscrollcommand=fy.set, xscrollcommand=fx.set)
        self.file_table.grid(row=1, column=0, sticky="ew")
        fy.grid(row=1, column=1, sticky="ns")
        fx.grid(row=2, column=0, sticky="ew")
        self.file_table.bind("<<TreeviewSelect>>", self._on_file_selected)
        for tag, foreground, background in (
            ("ready", COLORS["success"], COLORS["success_bg"]),
            ("review", COLORS["warning"], COLORS["warning_bg"]),
            ("error", COLORS["danger"], COLORS["danger_bg"]),
            ("pending", COLORS["muted"], "#F3F6F8"),
        ):
            self.file_table.tag_configure(tag, foreground=foreground, background=background)

        ttk.Label(files, text="工作表（双击可勾选/取消；每个选中工作表会生成独立文件）", style="Muted.TLabel").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(6, 3)
        )
        sheet_columns = ("selected", "sheet", "hidden", "header", "data", "rows", "images", "confidence", "bom", "status")
        self.sheet_table = ttk.Treeview(files, columns=sheet_columns, show="headings", height=4, selectmode="browse")
        for column, title, width in (
            ("selected", "选择", 54), ("sheet", "工作表", 170), ("hidden", "可见性", 88),
            ("header", "建议表头", 100), ("data", "数据开始", 78), ("rows", "数据行", 65),
            ("images", "图片", 55), ("confidence", "识别把握", 78), ("bom", "像BOM", 65), ("status", "状态", 170),
        ):
            self.sheet_table.heading(column, text=title)
            self.sheet_table.column(column, width=width, minwidth=50, stretch=column in {"sheet", "status"}, anchor="center" if column not in {"sheet", "status"} else "w")
        sy = ttk.Scrollbar(files, orient="vertical", command=self.sheet_table.yview)
        sx = ttk.Scrollbar(files, orient="horizontal", command=self.sheet_table.xview)
        self.sheet_table.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.sheet_table.grid(row=4, column=0, sticky="ew")
        sy.grid(row=4, column=1, sticky="ns")
        sx.grid(row=5, column=0, sticky="ew")
        self.sheet_table.bind("<<TreeviewSelect>>", self._on_sheet_selected)
        self.sheet_table.bind("<Double-1>", self._toggle_selected_sheet)

        self.mapping_card = ttk.LabelFrame(self.workspace, text="2  检查并确认字段", style="Card.TLabelframe", padding=8)
        self.mapping_card.grid(row=1, column=0, sticky="nsew", pady=(0, 8))
        self.mapping_card.columnconfigure(0, weight=1)
        self.mapping_card.rowconfigure(1, weight=1)
        filterbar = ttk.Frame(self.mapping_card, style="Card.TFrame")
        filterbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
        ttk.Checkbutton(filterbar, text="只显示需要确认", variable=self.only_review, command=self._refresh_mapping_view).pack(side=LEFT)
        ttk.Label(filterbar, text="搜索来源列或目标字段", style="Card.TLabel").pack(side=LEFT, padx=(14, 4))
        ttk.Entry(filterbar, textvariable=self.search_text, width=28).pack(side=LEFT)
        ttk.Button(filterbar, text="手动配置格式", command=self._open_manual_mapping, style="Secondary.TButton").pack(side=RIGHT)
        ttk.Button(filterbar, text="恢复自动推荐", command=self._restore_recommendations, style="Secondary.TButton").pack(side=RIGHT, padx=6)
        ttk.Button(filterbar, text="批量忽略选中列", command=self._batch_ignore, style="Danger.TButton").pack(side=RIGHT)
        self.mapping_table = MappingMatrix(
            self.mapping_card,
            on_edit=self._open_mapping_editor,
            on_select=self._show_mapping_detail,
            height=320,
        )
        self.mapping_table.grid(row=1, column=0, columnspan=2, sticky="nsew")
        detail = ttk.Frame(self.mapping_card, style="Card.TFrame")
        detail.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Label(detail, textvariable=self.detail_text, style="Muted.TLabel").pack(side=LEFT)
        ttk.Button(detail, text="确认并转换", command=lambda: self._confirm_and_convert(False), style="Primary.TButton").pack(side=RIGHT)
        ttk.Button(detail, text="确认、转换并记住", command=lambda: self._confirm_and_convert(True), style="Secondary.TButton").pack(side=RIGHT, padx=6)
        ttk.Button(detail, text="仅保存规则", command=self._save_rule_only, style="Secondary.TButton").pack(side=RIGHT)

        self.settings_card = ttk.LabelFrame(self.workspace, text="3  选择模板和输出目录", style="Card.TLabelframe", padding=8)
        self.settings_card.grid(row=2, column=0, sticky="ew")
        self.settings_card.columnconfigure(0, weight=1)
        self.settings_body = ttk.Frame(self.settings_card, style="Card.TFrame")
        self.settings_body.grid(row=0, column=0, sticky="ew")
        self.settings_body.columnconfigure(1, weight=1)
        ttk.Label(self.settings_body, text="标准模板", style="Card.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Entry(self.settings_body, textvariable=self.template).grid(row=0, column=1, sticky="ew", padx=7)
        ttk.Button(self.settings_body, text="选择模板", command=self._choose_template, style="Secondary.TButton").grid(row=0, column=2)
        ttk.Label(self.settings_body, text="输出目录", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(self.settings_body, textvariable=self.output_dir).grid(row=1, column=1, sticky="ew", padx=7, pady=(6, 0))
        ttk.Button(self.settings_body, text="选择目录", command=self._choose_output, style="Secondary.TButton").grid(row=1, column=2, pady=(6, 0))
        actionbar = ttk.Frame(self.settings_body, style="Card.TFrame")
        actionbar.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.convert_button = ttk.Button(actionbar, text="开始转换", command=self._start_conversion, style="Primary.TButton")
        self.convert_button.pack(side=LEFT)
        ttk.Button(actionbar, text="打开输出目录", command=self._open_output_dir, style="Secondary.TButton").pack(side=LEFT, padx=6)
        self.log_toggle = ttk.Button(actionbar, text="详细日志", command=self._toggle_log, style="Secondary.TButton")
        self.log_toggle.pack(side=LEFT)
        self.progress = ttk.Progressbar(actionbar, mode="determinate", maximum=100)
        self.progress.pack(side=LEFT, fill="x", expand=True, padx=10)
        ttk.Label(actionbar, textvariable=self.status_text, style="Muted.TLabel").pack(side=RIGHT)
        self.collapse_settings_button = ttk.Button(
            actionbar,
            text="收起位置设置",
            command=self._collapse_settings,
            style="Secondary.TButton",
        )

        self.settings_summary = ttk.Frame(self.settings_card, style="Card.TFrame")
        ttk.Label(self.settings_summary, textvariable=self.settings_summary_text, style="Card.TLabel").pack(side=LEFT)
        ttk.Button(self.settings_summary, text="修改位置", command=self._expand_settings, style="Secondary.TButton").pack(side=RIGHT)
        ttk.Button(self.settings_summary, text="开始转换", command=self._start_conversion, style="Primary.TButton").pack(side=RIGHT, padx=6)

        self.result_page = ttk.Frame(self.content, style="App.TFrame")
        self.result_page.columnconfigure(0, weight=1)
        self.result_page.rowconfigure(2, weight=1)
        self.result_page.rowconfigure(3, weight=2)
        ttk.Label(self.result_page, text="转换结果", style="Section.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 8))
        counters = ttk.Frame(self.result_page, style="App.TFrame")
        counters.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        counters.columnconfigure((0, 1, 2), weight=1)
        self.success_count = StringVar(value="0")
        self.warning_count = StringVar(value="0")
        self.failure_count = StringVar(value="0")
        for column, kind, title, variable in (
            (0, "Success", RESULT_COUNTER_TITLES[0], self.success_count),
            (1, "Warning", RESULT_COUNTER_TITLES[1], self.warning_count),
            (2, "Danger", RESULT_COUNTER_TITLES[2], self.failure_count),
        ):
            card = ttk.Frame(counters, style=f"Counter{kind}.TFrame", padding=14)
            card.grid(row=0, column=column, sticky="ew", padx=(0 if column == 0 else 5, 0 if column == 2 else 5))
            ttk.Label(card, textvariable=variable, style=f"Counter{kind}Value.TLabel").pack()
            ttk.Label(card, text=title, style=f"Counter{kind}Text.TLabel").pack()
        results_card = ttk.LabelFrame(self.result_page, text="每个文件的输出情况（警告数和错误数均为条数）", style="Card.TLabelframe", padding=8)
        results_card.grid(row=2, column=0, sticky="nsew", pady=(0, 8))
        results_card.columnconfigure(0, weight=1)
        results_card.rowconfigure(1, weight=1)
        ttk.Label(results_card, textvariable=self.result_summary_text, style="Muted.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
        )
        result_columns = ("output", "rows", "images", "warnings", "errors", "status", "source", "sheet")
        self.result_table = ttk.Treeview(results_card, columns=result_columns, show="headings", height=5)
        for column, title, width in (
            ("output", "输出文件", 320),
            ("rows", "输出数据行", 105),
            ("images", "输出图片数", 105),
            ("warnings", "警告数", 85),
            ("errors", "错误数", 85),
            ("status", "结果", 120),
            ("source", "来源文件", 170),
            ("sheet", "来源工作表", 130),
        ):
            self.result_table.heading(column, text=title)
            self.result_table.column(column, width=width, anchor="center" if column not in {"source", "sheet", "output"} else "w")
        ry = ttk.Scrollbar(results_card, orient="vertical", command=self.result_table.yview)
        self.result_table.configure(yscrollcommand=ry.set)
        self.result_table.grid(row=1, column=0, sticky="nsew")
        ry.grid(row=1, column=1, sticky="ns")
        for tag, foreground, background in (
            ("ready", COLORS["success"], COLORS["success_bg"]),
            ("review", COLORS["warning"], COLORS["warning_bg"]),
            ("error", COLORS["danger"], COLORS["danger_bg"]),
        ):
            self.result_table.tag_configure(tag, foreground=foreground, background=background)
        self.result_table.bind("<<TreeviewSelect>>", self._show_selected_issues)
        self.result_table.bind("<Double-1>", self._open_issue_details)

        issues_card = ttk.LabelFrame(
            self.result_page,
            text="所选文件的完整问题列表",
            style="Card.TLabelframe",
            padding=8,
        )
        issues_card.grid(row=3, column=0, sticky="nsew", pady=(0, 8))
        issues_card.columnconfigure(0, weight=1)
        issues_card.rowconfigure(1, weight=1)
        ttk.Label(issues_card, textvariable=self.selected_result_text, style="Muted.TLabel").grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 5)
        )
        self.issue_table = self._create_issue_table(issues_card, height=6)
        self.issue_table.grid(row=1, column=0, sticky="nsew")
        issue_y = ttk.Scrollbar(issues_card, orient="vertical", command=self.issue_table.yview)
        issue_x = ttk.Scrollbar(issues_card, orient="horizontal", command=self.issue_table.xview)
        self.issue_table.configure(yscrollcommand=issue_y.set, xscrollcommand=issue_x.set)
        issue_y.grid(row=1, column=1, sticky="ns")
        issue_x.grid(row=2, column=0, sticky="ew")

        result_actions = ttk.Frame(self.result_page, style="Card.TFrame", padding=10)
        result_actions.grid(row=4, column=0, sticky="ew")
        self.return_mapping_button = ttk.Button(
            result_actions,
            text="返回检查字段",
            command=self._return_to_mapping,
            style="Secondary.TButton",
        )
        self.return_mapping_button.pack(side=LEFT)
        self.continue_files_button = ttk.Button(
            result_actions,
            text="继续添加文件",
            command=self._continue_adding_files,
            style="Secondary.TButton",
        )
        self.continue_files_button.pack(side=LEFT, padx=6)
        self.new_task_button = ttk.Button(
            result_actions,
            text="开始新任务",
            command=self._start_new_task,
            style="Secondary.TButton",
        )
        self.new_task_button.pack(side=LEFT)
        self.open_report_button = ttk.Button(
            result_actions,
            text="打开同名 .report.json",
            command=self._open_selected_report,
            style="Secondary.TButton",
            state="disabled",
        )
        self.open_report_button.pack(side=LEFT, padx=6)
        self.copy_issues_button = ttk.Button(
            result_actions,
            text="复制问题摘要",
            command=self._copy_selected_issue_summary,
            style="Secondary.TButton",
            state="disabled",
        )
        self.copy_issues_button.pack(side=LEFT)
        self.result_log_toggle = ttk.Button(result_actions, text="详细日志", command=self._toggle_log, style="Secondary.TButton")
        self.result_log_toggle.pack(side=LEFT, padx=6)
        self.open_result_dir_button = ttk.Button(
            result_actions,
            text="打开输出目录",
            command=self._open_selected_output_dir,
            style="Primary.TButton",
        )
        self.open_result_dir_button.pack(side=RIGHT)

        self.log_frame = ttk.Frame(outer, style="Card.TFrame", padding=7)
        self.log_frame.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        self.log = Text(self.log_frame, wrap="word", state="disabled", height=5, background="#F8FAFC", foreground=COLORS["text"], relief="flat")
        self.log.pack(fill="x")
        self.log_frame.grid_remove()

    def _on_mode_change(self) -> None:
        if self.mode.get() == "quick":
            self.mode_caption.set("适合已知格式；只有异常时才需要审核字段")
            self.mapping_card.grid_remove()
            self.workspace.rowconfigure(1, weight=0)
            self.file_table.configure(height=3)
            self.settings_summary.grid_remove()
            self.settings_body.grid(row=0, column=0, sticky="ew")
            self.collapse_settings_button.pack_forget()
        else:
            self.mode_caption.set("适合陌生或变化格式；可检查和保存字段规则")
            self.mapping_card.grid()
            self.workspace.rowconfigure(1, weight=1)
            self.file_table.configure(height=1)
            self._collapse_settings()

    def _update_settings_summary(self) -> None:
        template_name = Path(self.template.get()).name or "未选择模板"
        output_name = Path(self.output_dir.get()).name or self.output_dir.get() or "未选择目录"
        self.settings_summary_text.set(f"模板：{template_name}    输出目录：{output_name}")

    def _expand_settings(self) -> None:
        self.settings_summary.grid_remove()
        self.settings_body.grid(row=0, column=0, sticky="ew")
        if self.mode.get() == "audit" and not self.collapse_settings_button.winfo_ismapped():
            self.collapse_settings_button.pack(side=RIGHT, padx=(6, 0))

    def _collapse_settings(self) -> None:
        if not hasattr(self, "settings_summary"):
            return
        self.settings_body.grid_remove()
        self.collapse_settings_button.pack_forget()
        self._update_settings_summary()
        self.settings_summary.grid(row=0, column=0, sticky="ew")

    def _choose_sources(self) -> None:
        files = filedialog.askopenfilenames(filetypes=[("Excel 工作簿", "*.xlsx")])
        paths: list[Path] = []
        for name in files:
            path = Path(name).resolve()
            if path.name.startswith("~$") or path in self.sources:
                continue
            self.sources.append(path)
            paths.append(path)
            self.file_table.insert(
                "",
                END,
                iid=str(path),
                values=(path.name, "等待识别", "—", "—", "—", "分析中"),
                tags=("pending",),
            )
        if paths:
            self.status_text.set(f"正在分析 {len(paths)} 个文件……")
            threading.Thread(target=self._worker_analyze, args=(paths, self.mode.get()), daemon=True).start()

    @staticmethod
    def _sheet_item_id(path: Path, sheet_name: str) -> str:
        return f"sheet::{len(str(path))}::{path}::{sheet_name}"

    def _worker_analyze(self, paths: list[Path], mode: str) -> None:
        total = len(paths)
        for index, path in enumerate(paths, 1):
            try:
                sheets = analyze_workbook_sheets(path, mode=mode, mapping_memory=self.mapping_store)
                self.events.put(("sheet_analyses", (path, sheets)))
            except Exception as exc:
                self.events.put(("file_error", (path, str(exc))))
            self.events.put(("progress", index * 100 / max(total, 1)))
        self.events.put(("done", "文件和工作表分析完成"))

    def _handle_sheet_analyses(self, path: Path, sheets: list[SheetAnalysis]) -> None:
        for key in [item for item in self.sheet_analyses if item[0] == path]:
            self.sheet_analyses.pop(key, None)
            self.selected_sheets.discard(key)
        usable = [item for item in sheets if item.analysis]
        for item in sheets:
            key = (path, item.sheet_name)
            self.sheet_analyses[key] = item
            if item.recommended_selected:
                self.selected_sheets.add(key)
        primary = next((item.analysis for item in sheets if item.recommended_selected and item.analysis), None)
        primary = primary or next((item.analysis for item in sheets if item.analysis), None)
        if primary:
            self.analyses[path] = primary
        selected_count = sum((path, item.sheet_name) in self.selected_sheets for item in sheets)
        status = f"已分析 {len(sheets)} 个工作表；默认选择 {selected_count} 个"
        tag = "ready" if selected_count else "review"
        self.file_table.item(
            str(path),
            values=(path.name, f"{len(sheets)} 个工作表", "逐表分析", sum(item.analysis.data_row_count for item in usable), sum(item.analysis.image_count for item in usable), status),
            tags=(tag,),
        )
        self._show_sheets_for(path)
        if primary and not self.current_sheet_key:
            key = (path, primary.sheet_name)
            self._activate_sheet(key)

    def _show_sheets_for(self, path: Path) -> None:
        for item in self.sheet_table.get_children():
            self.sheet_table.delete(item)
        for key, item in self.sheet_analyses.items():
            if key[0] != path:
                continue
            analysis = item.analysis
            values = (
                "☑" if key in self.selected_sheets else "☐",
                item.sheet_name,
                "隐藏工作表" if item.hidden else "可见",
                header_region_text(analysis) if analysis else "—",
                analysis.data_start_row if analysis else "—",
                analysis.data_row_count if analysis else "—",
                analysis.image_count if analysis else "—",
                confidence_text(analysis.profile_confidence) if analysis else "—",
                "是" if item.is_bom else "否",
                item.status if not item.error else f"{item.status}：{item.error}",
            )
            iid = self._sheet_item_id(*key)
            self.sheet_table.insert("", END, iid=iid, values=values, tags=("ready" if key in self.selected_sheets else "pending",))
        for tag, foreground, background in (
            ("ready", COLORS["success"], COLORS["success_bg"]),
            ("pending", COLORS["muted"], "#F3F6F8"),
        ):
            self.sheet_table.tag_configure(tag, foreground=foreground, background=background)

    def _key_from_sheet_item(self, item_id: str) -> tuple[Path, str] | None:
        for key in self.sheet_analyses:
            if self._sheet_item_id(*key) == item_id:
                return key
        return None

    def _on_file_selected(self, _event: object = None) -> None:
        source = self._selected_source()
        if source:
            self._show_sheets_for(source)

    def _on_sheet_selected(self, _event: object = None) -> None:
        selected = self.sheet_table.selection()
        key = self._key_from_sheet_item(selected[0]) if selected else None
        if key:
            self._activate_sheet(key)

    def _activate_sheet(self, key: tuple[Path, str]) -> None:
        item = self.sheet_analyses.get(key)
        if not item or not item.analysis:
            return
        source, _sheet = key
        self.current_sheet_key = key
        self.current_source = source
        self.analyses[source] = item.analysis
        if key not in self.sheet_mapping_choices:
            self.sheet_mapping_choices[key] = {decision.source_col: decision.universal_field for decision in item.analysis.mappings}
        self.mapping_choices[source] = self.sheet_mapping_choices[key]
        if key in self.sheet_manual_configs:
            self.manual_configs[source] = self.sheet_manual_configs[key]
        else:
            self.manual_configs.pop(source, None)
        self._manual_confirmed[source] = self.sheet_confirmed.setdefault(key, set())
        self._load_mapping_table(source, restore=False)
        self.detail_text.set(f"当前工作表：{item.sheet_name}；该工作表的映射和名称规则独立保存。")

    def _toggle_selected_sheet(self, _event: object = None) -> None:
        selected = self.sheet_table.selection()
        key = self._key_from_sheet_item(selected[0]) if selected else None
        if not key:
            return
        if key in self.selected_sheets:
            self.selected_sheets.remove(key)
        else:
            self.selected_sheets.add(key)
        self._show_sheets_for(key[0])
        iid = self._sheet_item_id(*key)
        if self.sheet_table.exists(iid):
            self.sheet_table.selection_set(iid)
        self._activate_sheet(key)

    def _clear_sources(self) -> None:
        super()._clear_sources()
        self.sheet_analyses.clear()
        self.selected_sheets.clear()
        self.sheet_mapping_choices.clear()
        self.sheet_manual_configs.clear()
        self.sheet_confirmed.clear()
        self.current_sheet_key = None
        for item in self.sheet_table.get_children():
            self.sheet_table.delete(item)
        self._mapping_rows.clear()
        self._mapping_internal.clear()
        self._manual_confirmed.clear()
        self._conversion_results.clear()
        self._result_by_item.clear()
        for item in self.result_table.get_children():
            self.result_table.delete(item)
        for item in self.issue_table.get_children():
            self.issue_table.delete(item)
        self.selected_result_text.set("选择一个输出文件，即可在这里直接查看完整问题列表。")
        self.open_report_button.configure(state="disabled")
        self.copy_issues_button.configure(state="disabled")
        self._update_result_counts()

    def _poll(self) -> None:
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "analysis":
                    path, analysis = value
                    self._handle_analysis(path, analysis)
                elif kind == "sheet_analyses":
                    path, sheets = value
                    self._handle_sheet_analyses(path, sheets)
                elif kind == "conversion":
                    self._handle_conversion(value)
                elif kind == "file_error":
                    path, error = value
                    if str(path) in self.file_table.get_children():
                        old = self.file_table.item(str(path), "values")
                        self.file_table.item(
                            str(path),
                            values=(*old[:5], f"错误：{error}"),
                            tags=("error",),
                        )
                    self._append(f"错误：{Path(path).name}: {error}")
                elif kind == "sheet_error":
                    path, sheet, error = value
                    self._append(f"错误：{Path(path).name} / 工作表 {sheet or '—'}：{error}")
                elif kind == "progress":
                    self.progress["value"] = float(value)
                elif kind == "done":
                    self.convert_button.configure(state="normal")
                    self.status_text.set(str(value))
        except queue.Empty:
            pass
        self.root.after(120, self._poll)

    def _analysis_status(self, analysis: WorkbookAnalysis) -> tuple[str, str]:
        issue_codes = {issue.code for issue in analysis.issues}
        if any(code in issue_codes for code in {"UNIT_CHANGED", "HEADER_REGION_CHANGED", "SIMILAR_SAVED_FORMAT"}):
            return "格式有变化，需要确认", "review"
        if any(code in issue_codes for code in {"LEVEL_MARKER_MULTIPLE", "LEVEL_MARKER_MISSING", "IMAGE_SLOT_OVERFLOW"}):
            return "内容存在异常，需要确认", "review"
        if analysis.profile_id == "generic" and analysis.requires_review:
            return "陌生格式，需要确认", "review"
        if analysis.profile_id == "generic":
            return "历史规则完整匹配", "ready"
        if analysis.requires_review:
            return "识别把握较低，需要确认", "review"
        return "可以转换", "ready"

    def _handle_analysis(self, path: Path, analysis: WorkbookAnalysis) -> None:
        self.analyses[path] = analysis
        self.mapping_choices.pop(path, None)
        self.manual_configs.pop(path, None)
        status, tag = self._analysis_status(analysis)
        self.file_table.item(
            str(path),
            values=(path.name, analysis.profile_name, header_region_text(analysis), analysis.data_row_count, analysis.image_count, status),
            tags=(tag,),
        )
        issue = next((item for item in analysis.issues if item.severity in {"warning", "error"}), None)
        if issue:
            self._append(f"{path.name}：{issue.message}")
        if not self.file_table.selection():
            self.file_table.selection_set(str(path))
            self.current_source = path
            self._load_mapping_table(path, restore=False)

    def _reanalyze_selected(self) -> None:
        source = self._selected_source()
        if not source:
            messagebox.showwarning("未选择文件", "请先在文件列表中选择一个文件。")
            return
        self.file_table.item(str(source), values=(source.name, "等待识别", "—", "—", "—", "重新检查中"), tags=("pending",))
        threading.Thread(target=self._worker_analyze, args=([source], self.mode.get()), daemon=True).start()

    def _open_manual_mapping(self) -> None:
        source = self._selected_source()
        if not source:
            messagebox.showwarning("未选择文件", "请先选择一份来源工作簿。")
            return
        self.preview_dialog = ManualMappingDialogV2(
            self.root,
            source,
            self.mapping_store,
            lambda analysis, mappings, units, configs: self._accept_manual_conversion(source, analysis, mappings, units, configs),
            initial_sheet=self.current_sheet_key[1] if self.current_sheet_key and self.current_sheet_key[0] == source else None,
        )

    def _accept_manual_conversion(
        self,
        source: Path,
        analysis: WorkbookAnalysis,
        mappings: dict[int, str | None],
        units: dict[str, str],
        configs: dict[int, ColumnMappingConfig],
    ) -> None:
        self.analyses[source] = analysis
        self.mapping_choices[source] = mappings
        self.manual_configs[source] = configs
        self.current_source = source
        self._manual_confirmed[source] = set(configs)
        key = (source, analysis.sheet_name)
        self.current_sheet_key = key
        self.sheet_mapping_choices[key] = mappings
        self.sheet_manual_configs[key] = configs
        self.sheet_confirmed[key] = set(configs)
        self.selected_sheets.add(key)
        if key in self.sheet_analyses:
            self.sheet_analyses[key].analysis = analysis
        self.file_table.item(
            str(source),
            values=(source.name, f"手动配置：{analysis.sheet_name}", header_region_text(analysis), analysis.data_row_count, analysis.image_count, "已人工确认"),
            tags=("ready",),
        )
        self._load_mapping_table(source, restore=False)
        self._run_jobs([("custom", source, analysis, mappings, units, configs)])

    def _load_mapping_table(self, source: Path, restore: bool) -> None:
        analysis = self.analyses[source]
        if restore or source not in self.mapping_choices:
            self.mapping_choices[source] = {item.source_col: item.universal_field for item in analysis.mappings}
        choices = self.mapping_choices[source]
        confirmed = self._manual_confirmed.get(source, set())
        source_label = recommendation_source(analysis)
        self._mapping_rows.clear()
        self._mapping_internal.clear()
        for decision in analysis.mappings:
            field = choices.get(decision.source_col)
            target = TARGET_HEADERS.get(field or "", "—").replace("\n", " / ")
            sample = "" if decision.sample_value is None else str(decision.sample_value)
            if len(sample) > 60:
                sample = sample[:57] + "..."
            status = "已忽略" if not field else ("需要确认" if decision.status == "review" else "已匹配")
            recommendation = "人工确认" if decision.source_col in confirmed else source_label
            values = (
                f"{decision.column_letter or decision.source_col} 列",
                decision.source_header,
                sample,
                V2_FIELD_LABELS.get(field or "", "忽略"),
                target,
                confidence_text(decision.confidence, confirmed=decision.source_col in confirmed),
                status,
                recommendation,
            )
            key = str(decision.source_col)
            self._mapping_rows[key] = values
            self._mapping_internal[key] = field or "—"
        self._refresh_mapping_view()

    def _refresh_mapping_view(self) -> None:
        if not hasattr(self, "mapping_table"):
            return
        selected = set(self.mapping_table.selection())
        query = self.search_text.get().strip().casefold()
        visible: list[str] = []
        for key, values in self._mapping_rows.items():
            if self.only_review.get() and values[6] != "需要确认":
                continue
            searchable = " ".join((*values, self._mapping_internal.get(key, ""))).casefold()
            if query and query not in searchable:
                continue
            visible.append(key)
        columns = [
            MatrixColumn(key, self._mapping_rows[key][0], {
                "path": self._mapping_rows[key][1],
                "sample": self._mapping_rows[key][2],
                "field": self._mapping_rows[key][3],
                "target": self._mapping_rows[key][4],
                "confidence": self._mapping_rows[key][5],
                "status": self._mapping_rows[key][6],
                "recommendation": self._mapping_rows[key][7],
            })
            for key in self._mapping_rows
        ]
        self.mapping_table.set_columns(columns)
        self.mapping_table.set_visible(visible)
        for key in selected:
            if self.mapping_table.exists(key):
                self.mapping_table.selection_add(key)

    def _open_mapping_editor(self, item: str) -> None:
        current = self._mapping_rows.get(item, ("", "", "", "忽略"))[3]
        self.mapping_table.open_field_editor(
            item,
            FIELD_OPTIONS_V2,
            current,
            lambda display: self._commit_mapping_value(item, display),
        )

    def _edit_mapping(self, event: object) -> None:
        # Backward-compatible event entry; MappingMatrix normally invokes _open_mapping_editor.
        selected = self.mapping_table.selection()
        if selected:
            self._open_mapping_editor(selected[0])

    def _commit_mapping_value(self, item: str, display: str) -> None:
        class _Value:
            def get(self) -> str:
                return display
            def destroy(self) -> None:
                return None
        self._commit_mapping_editor_v2(item, _Value())

    def _commit_mapping_editor_v2(self, item: str, editor: ttk.Combobox) -> None:
        if not self.current_source:
            editor.destroy()
            return
        field = DISPLAY_TO_FIELD_V2.get(editor.get())
        column = int(item)
        self.mapping_choices[self.current_source][column] = field
        self._manual_confirmed.setdefault(self.current_source, set()).add(column)
        if self.current_source in self.manual_configs:
            config = self.manual_configs[self.current_source][column]
            config.universal_field = field
            config.process_type = "ignore" if field is None else ("direct" if config.process_type == "ignore" else config.process_type)
        row = list(self._mapping_rows[item])
        row[3] = V2_FIELD_LABELS.get(field or "", "忽略")
        row[4] = TARGET_HEADERS.get(field or "", "—").replace("\n", " / ")
        row[5] = "已人工确认"
        row[6] = "已忽略" if field is None else "已匹配"
        row[7] = "人工确认"
        self._mapping_rows[item] = tuple(row)
        self._mapping_internal[item] = field or "—"
        editor.destroy()
        self._refresh_mapping_view()
        if self.mapping_table.exists(item):
            self.mapping_table.selection_set(item)
        self._show_mapping_detail()

    def _batch_ignore(self) -> None:
        if not self.current_source:
            return
        selected = self.mapping_table.selection()
        if not selected:
            messagebox.showwarning("未选择列", "请先选择一列或多列。")
            return
        for item in selected:
            column = int(item)
            self.mapping_choices[self.current_source][column] = None
            self._manual_confirmed.setdefault(self.current_source, set()).add(column)
            row = list(self._mapping_rows[item])
            row[3], row[4], row[5], row[6], row[7] = "忽略", "—", "已人工确认", "已忽略", "人工确认"
            self._mapping_rows[item] = tuple(row)
            self._mapping_internal[item] = "—"
        self._refresh_mapping_view()

    def _restore_recommendations(self) -> None:
        if not self.current_source:
            return
        analysis = (
            analyze_sheet(self.current_source, self.current_sheet_key[1], mode="audit")
            if self.current_sheet_key and self.current_sheet_key[0] == self.current_source
            else analyze_workbook(self.current_source, mode="audit")
        )
        self.analyses[self.current_source] = analysis
        self.mapping_choices.pop(self.current_source, None)
        self.manual_configs.pop(self.current_source, None)
        self._manual_confirmed.pop(self.current_source, None)
        self._load_mapping_table(self.current_source, restore=True)
        self._append("已恢复自动推荐；尚未写入规则库。")

    def _start_conversion(self) -> None:
        if not self.sources:
            messagebox.showwarning("缺少文件", "请先添加至少一份来源 BOM。")
            return
        selected = [key for key in self.selected_sheets if key in self.sheet_analyses]
        if not selected:
            messagebox.showwarning("未选择工作表", "请至少勾选一个要转换的工作表。")
            return
        jobs: list[tuple] = []
        for key in sorted(selected, key=lambda item: (str(item[0]).casefold(), item[1].casefold())):
            sheet_item = self.sheet_analyses[key]
            analysis = sheet_item.analysis
            if analysis is None:
                continue
            source = key[0]
            if analysis.profile_id != "generic":
                jobs.append(("known_sheet", source, analysis))
                continue
            mappings = self.sheet_mapping_choices.get(key) or {
                decision.source_col: decision.universal_field for decision in analysis.mappings
            }
            if analysis.requires_review and key not in self.sheet_confirmed:
                self._activate_sheet(key)
                messagebox.showwarning(
                    "需要确认映射",
                    f"{source.name} 的工作表“{analysis.sheet_name}”需要先确认字段，再开始转换。",
                )
                return
            units = infer_confirmed_units(analysis, mappings)
            configs = self.sheet_manual_configs.get(key) or configs_from_analysis(analysis, mappings, units)
            jobs.append(("custom_sheet", source, analysis, mappings, units, configs))
        if not jobs:
            messagebox.showwarning("没有可转换工作表", "所选工作表没有可用分析结果。")
            return
        self._run_jobs(jobs)

    def _confirm_and_convert(self, remember: bool) -> None:
        if self.current_sheet_key and self.current_source:
            self.sheet_confirmed[self.current_sheet_key] = set(self.mapping_choices.get(self.current_source, {}))
            self.sheet_mapping_choices[self.current_sheet_key] = self.mapping_choices.get(self.current_source, {})
            if self.current_source in self.manual_configs:
                self.sheet_manual_configs[self.current_sheet_key] = self.manual_configs[self.current_source]
        super()._confirm_and_convert(remember)

    def _worker_convert(self, jobs: list[tuple], template_path: str, output_dir: str) -> None:
        total = len(jobs)
        for index, job in enumerate(jobs, 1):
            try:
                if job[0] == "known_sheet":
                    _, source, analysis = job
                    result = convert_analyzed_file(source, template_path, output_dir, analysis, include_sheet_name=True)
                else:
                    _, source, analysis, mappings, units, configs = job
                    result = convert_confirmed_file(source, template_path, output_dir, analysis, mappings, units, configs)
                self.events.put(("conversion", result))
            except Exception as exc:
                source = job[1]
                sheet = job[2].sheet_name if len(job) > 2 and isinstance(job[2], WorkbookAnalysis) else None
                self.events.put(("sheet_error", (source, sheet, str(exc))))
            self.events.put(("progress", index * 100 / max(total, 1)))
        self.events.put(("done", "转换任务结束"))

    def _show_mapping_detail(self, _event: object = None) -> None:
        selected = self.mapping_table.selection()
        if not selected:
            self.detail_text.set("选择一行可查看内部字段和推荐说明。")
            return
        key = selected[0]
        values = self._mapping_rows.get(key)
        if values:
            self.detail_text.set(f"内部字段：{self._mapping_internal.get(key, '—')}    推荐来源：{values[7]}    示例值只显示在当前窗口，不保存。")

    def _create_issue_table(self, parent: object, height: int = 8) -> ttk.Treeview:
        columns = ("severity", "title", "description", "sheet", "row", "field", "action")
        table = ttk.Treeview(parent, columns=columns, show="headings", height=height)
        for column, title, width, anchor in (
            ("severity", "严重程度", 115, "center"),
            ("title", "问题标题", 280, "w"),
            ("description", "说明", 440, "w"),
            ("sheet", "来源工作表", 140, "w"),
            ("row", "来源行", 80, "center"),
            ("field", "涉及字段", 150, "w"),
            ("action", "建议操作", 470, "w"),
        ):
            table.heading(column, text=title)
            table.column(column, width=width, minwidth=60, anchor=anchor, stretch=False)
        for tag, foreground, background in (
            ("info", COLORS["primary"], COLORS["soft_blue"]),
            ("warning", COLORS["warning"], COLORS["warning_bg"]),
            ("error", COLORS["danger"], COLORS["danger_bg"]),
        ):
            table.tag_configure(tag, foreground=foreground, background=background)
        return table

    def _selected_result(self) -> ConversionResult | None:
        selected = self.result_table.selection()
        return self._result_by_item.get(selected[0]) if selected else None

    def _show_selected_issues(self, _event: object = None) -> None:
        for item in self.issue_table.get_children():
            self.issue_table.delete(item)
        result = self._selected_result()
        if result is None:
            self.selected_result_text.set("选择一个输出文件，即可在这里直接查看完整问题列表。")
            self.open_report_button.configure(state="disabled")
            self.copy_issues_button.configure(state="disabled")
            return
        displays = [present_issue(issue, V2_FIELD_LABELS) for issue in result.issues]
        for index, display in enumerate(displays):
            self.issue_table.insert(
                "",
                END,
                iid=f"issue-{index}",
                values=display.table_values(),
                tags=(display.tag,),
            )
        warnings = sum(item.severity == "warning" for item in displays)
        errors = sum(item.severity == "error" for item in displays)
        infos = sum(item.severity == "info" for item in displays)
        if displays:
            self.selected_result_text.set(
                f"{result.output_path.name}：{infos} 条处理提示，{warnings} 条警告，{errors} 条错误。"
            )
        else:
            self.selected_result_text.set(f"{result.output_path.name}：没有业务警告或错误。")
        report = self._selected_report_path(result)
        self.open_report_button.configure(state="normal" if report and report.exists() else "disabled")
        self.copy_issues_button.configure(state="normal")

    @staticmethod
    def _selected_report_path(result: ConversionResult) -> Path | None:
        if result.report_path:
            return Path(result.report_path)
        candidate = result.output_path.with_suffix(".report.json")
        return candidate if candidate.exists() else None

    def _open_selected_report(self) -> None:
        result = self._selected_result()
        report = self._selected_report_path(result) if result else None
        if not report or not report.exists():
            messagebox.showinfo("没有报告", "所选文件没有可打开的同名 .report.json。")
            return
        try:
            os.startfile(report)  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("无法打开报告", str(exc))

    def _copy_selected_issue_summary(self) -> None:
        result = self._selected_result()
        if result is None:
            messagebox.showinfo("尚未选择", "请先在结果表格中选择一个输出文件。")
            return
        summary = issue_summary_text(result, V2_FIELD_LABELS)
        self.root.clipboard_clear()
        self.root.clipboard_append(summary)
        self.root.update_idletasks()
        self.status_text.set("问题摘要已复制，可粘贴到邮件或验收记录中。")

    def _open_selected_output_dir(self) -> None:
        result = self._selected_result()
        target = result.output_path.parent if result else Path(self.output_dir.get()).resolve()
        try:
            target.mkdir(parents=True, exist_ok=True)
            os.startfile(target)  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("无法打开目录", str(exc))

    def _return_to_mapping(self) -> None:
        if self.mode.get() != "audit":
            self.mode.set("audit")
        self._show_workspace()
        for index, label in enumerate(self.step_labels, 1):
            label.configure(style="StepActive.TLabel" if index == 3 else "Step.TLabel")
        self.status_text.set("已返回字段检查；修改并确认后可以重新转换。")

    def _open_issue_details(self, _event: object = None) -> Toplevel | None:
        result = self._selected_result()
        if result is None:
            return None
        window = Toplevel(self.root)
        self.last_issue_detail_window = window
        window.title("警告与错误详情")
        window.geometry("1220x560+45+45")
        window.minsize(900, 420)
        window.configure(background=COLORS["canvas"])
        body = ttk.Frame(window, style="App.TFrame", padding=12)
        body.pack(fill=BOTH, expand=True)
        ttk.Label(body, text="警告与错误详情", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            body,
            text=f"输出文件：{result.output_path.name}。示例值和业务数据不会写入新的文件。",
            style="TLabel",
        ).pack(anchor="w", pady=(2, 8))
        table_frame = ttk.Frame(body, style="Card.TFrame")
        table_frame.pack(fill=BOTH, expand=True)
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)
        table = self._create_issue_table(table_frame, height=12)
        table.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(table_frame, orient="vertical", command=table.yview)
        scroll_x = ttk.Scrollbar(table_frame, orient="horizontal", command=table.xview)
        table.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        for index, issue in enumerate(result.issues):
            display = present_issue(issue, V2_FIELD_LABELS)
            table.insert("", END, iid=f"detail-{index}", values=display.table_values(), tags=(display.tag,))
        ttk.Button(body, text="关闭", command=window.destroy, style="Primary.TButton").pack(anchor="e", pady=(8, 0))
        return window

    def load_result_report(self, report_path: str | Path) -> ConversionResult:
        """Locally display an existing compatible report; the report is never rewritten."""

        result = conversion_result_from_report(report_path)
        self._handle_conversion(result)
        return result

    def _handle_conversion(self, result: ConversionResult) -> None:
        warnings = sum(issue.severity == "warning" for issue in result.issues)
        errors = sum(issue.severity == "error" for issue in result.issues)
        self._conversion_results.append(result)
        item_id = f"result-{len(self._conversion_results) - 1}"
        self._result_by_item[item_id] = result
        self.result_table.insert(
            "",
            END,
            iid=item_id,
            values=(
                result.output_path.name,
                result.output_rows,
                result.output_images,
                warnings,
                errors,
                "失败" if errors else ("完成，有警告" if warnings else "成功"),
                result.source_path.name,
                result.source_sheet or "—",
            ),
            tags=("error" if errors else ("review" if warnings else "ready"),),
        )
        if str(result.source_path) in self.file_table.get_children():
            old = self.file_table.item(str(result.source_path), "values")
            self.file_table.item(str(result.source_path), values=(*old[:5], "失败" if errors else "转换完成"), tags=("error" if errors else "ready",))
        self._append(f"完成：{result.output_path}")
        self._update_result_counts()
        self._show_results()
        self.result_table.selection_set(item_id)
        self.result_table.focus(item_id)
        self._show_selected_issues()

    def _update_result_counts(self) -> None:
        counts = count_results(self._conversion_results)
        self.success_count.set(str(counts.completely_successful_files))
        self.warning_count.set(str(counts.files_with_warnings))
        self.failure_count.set(str(counts.failed_files))
        self.result_summary_text.set(counts.summary_text() if counts.file_count else "尚无转换结果。")

    def _show_results(self) -> None:
        self.workspace.grid_remove()
        self.result_page.grid(row=0, column=0, sticky="nsew")
        for index, label in enumerate(self.step_labels, 1):
            label.configure(style="StepActive.TLabel" if index == 6 else "Step.TLabel")

    def _show_workspace(self) -> None:
        self.result_page.grid_remove()
        self.workspace.grid(row=0, column=0, sticky="nsew")
        for index, label in enumerate(self.step_labels, 1):
            label.configure(style="StepActive.TLabel" if index == 1 else "Step.TLabel")

    def _continue_adding_files(self) -> None:
        self._show_workspace()
        self.status_text.set("可继续添加文件；当前任务中的文件和结果已保留。")

    def _start_new_task(self) -> None:
        if self.preview_dialog and self.preview_dialog.window.winfo_exists():
            self.preview_dialog.window.destroy()
        self.preview_dialog = None
        self._clear_sources()
        self.progress["value"] = 0
        self.status_text.set("新任务已就绪：模板、输出目录和本机历史规则已保留。")
        self.log.configure(state="normal")
        self.log.delete("1.0", END)
        self.log.configure(state="disabled")
        self._show_workspace()

    def _toggle_log(self) -> None:
        self.log_visible = not self.log_visible
        if self.log_visible:
            self.log_frame.grid()
            self.log_toggle.configure(text="收起日志")
            self.result_log_toggle.configure(text="收起日志")
        else:
            self.log_frame.grid_remove()
            self.log_toggle.configure(text="详细日志")
            self.result_log_toggle.configure(text="详细日志")

    def load_preview_state(self, state: str) -> None:
        """用完全虚构的界面数据展示四种状态，不写入规则或 Excel。"""
        for item in self.file_table.get_children():
            self.file_table.delete(item)
        self._mapping_rows.clear()
        self._conversion_results.clear()
        self._result_by_item.clear()
        for item in self.result_table.get_children():
            self.result_table.delete(item)
        for item in self.issue_table.get_children():
            self.issue_table.delete(item)
        self._update_result_counts()
        if state == "quick":
            self.mode.set("quick")
            demos = (
                ("sim_single_header.xlsx", "公开演示格式", "第 4 行；数据第 5 行起", 3, 1, "可以转换", "ready"),
                ("仿真_历史规则_B.xlsx", "已记忆的通用格式", "第 4–5 行；数据第 6 行起", 26, 2, "历史规则完整匹配", "ready"),
                ("仿真_表头变化_C.xlsx", "陌生格式", "第 7–8 行；数据第 9 行起", 12, 1, "需要进入通用审核", "review"),
            )
            for index, values in enumerate(demos):
                self.file_table.insert("", END, iid=f"demo-{index}", values=values[:-1], tags=(values[-1],))
            self.status_text.set("2 个文件可以转换；1 个文件需要审核")
        elif state in {"audit", "manual"}:
            self.mode.set("audit")
            self.file_table.insert("", END, iid="demo-audit", values=("仿真_多行表头.xlsx", "陌生格式", "第 4–5 行；数据第 6 行起", 12, 2, "需要确认"), tags=("review",))
            rows = (
                ("11", ("K 列", "Size (mm) / X", "1250", "长度", "Length(mm)", "高", "已匹配", "自动推荐"), "length"),
                ("12", ("L 列", "Size (mm) / Y", "480", "宽度", "Width(mm)", "高", "已匹配", "自动推荐"), "width"),
                ("13", ("M 列", "Size (mm) / Z", "65", "高度", "Height(mm)", "高", "已匹配", "自动推荐"), "height"),
                ("14", ("N 列", "Material / Spec", "SIM-MAT", "材料规格", "Material Spec", "中", "需要确认", "自动推荐"), "material_spec"),
                ("15", ("O 列", "Extra Note", "仅作仿真", "忽略", "—", "低", "需要确认", "自动推荐"), "—"),
            )
            for key, values, internal in rows:
                self._mapping_rows[key] = values
                self._mapping_internal[key] = internal
            self._refresh_mapping_view()
            self.status_text.set("请确认黄色项目后再转换")
            if state == "manual":
                self.root.after(180, self._open_preview_manual_dialog)
        elif state == "results":
            self.mode.set("quick")
            demo_results = (
                ConversionResult(Path("仿真_A.xlsx"), Path("仿真_A_标准格式.xlsx"), "demo", 18, 18, 4, 4, []),
                ConversionResult(
                    Path("仿真_B.xlsx"),
                    Path("仿真_B_标准格式.xlsx"),
                    "generic",
                    26,
                    26,
                    2,
                    2,
                    [
                        Issue("UNKNOWN_PROFILE", "warning", "未匹配已知格式，需要检查通用映射", "仿真表", None, None),
                        Issue("LOW_HEADER_CONFIDENCE", "warning", "表头识别把握较低，需要检查未识别列", "仿真表", None, None),
                    ],
                ),
                ConversionResult(Path("仿真_C.xlsx"), Path("仿真_C_标准格式.xlsx"), "demo", 12, 12, 1, 1, []),
            )
            for result in demo_results:
                self._handle_conversion(result)
            self.result_table.selection_set("result-1")
            self.result_table.focus("result-1")
            self._show_selected_issues()
        else:
            raise ValueError(f"未知预览状态：{state}")

    def _open_preview_manual_dialog(self) -> None:
        source = _resource_root() / "samples" / "simulated" / "sim_complex_headers.xlsx"
        self.preview_dialog = ManualMappingDialogV2(self.root, source, self.mapping_store, lambda *_args: None)
        self.preview_dialog.sheet_name.set("SIM XYZ Bilingual")
        self.preview_dialog.header_start.set("4")
        self.preview_dialog.header_end.set("5")
        self.preview_dialog.data_start.set("6")
        self.preview_dialog._analyze()
        if self.preview_dialog.table.exists("11"):
            self.preview_dialog.table.selection_set("11")
            self.preview_dialog._load_selected()


def core_interface_probe() -> dict[str, object]:
    """只调用现有分析和中间转换接口，不写模板、不保存规则。"""
    source = _resource_root() / "samples" / "simulated" / "sim_complex_headers.xlsx"
    analysis = analyze_manual_workbook(source, "SIM XYZ Bilingual", 4, 5, 6)
    mappings = {item.source_col: item.universal_field for item in analysis.mappings}
    units = infer_confirmed_units(analysis, mappings)
    configs = configs_from_analysis(analysis, mappings, units)
    next_slot = 1
    for config in configs.values():
        if config.process_type == "image" and not config.image_slot:
            config.image_slot = next_slot
            next_slot += 1
    rows, issues = transform_confirmed_workbook(source, analysis, mappings, units, configs)
    return {
        "sheet": analysis.sheet_name,
        "header_start_row": analysis.header_start_row,
        "header_end_row": analysis.header_end_row,
        "data_start_row": analysis.data_start_row,
        "data_rows": len(rows),
        "images": sum(len(row.images) for row in rows),
        "errors": sum(issue.severity == "error" for issue in issues),
    }


def write_smoke_result(path: Path, app: BomConverterAppV2) -> None:
    app.root.update_idletasks()
    source = _resource_root() / "samples" / "simulated" / "sim_complex_headers.xlsx"
    payload = {
        "window_title": app.root.title(),
        "mode": app.mode.get(),
        "window_size": [app.root.winfo_width(), app.root.winfo_height()],
        "theme_backend": "原生 tkinter.ttk / clam（ttkbootstrap 因当前 Tcl 缺少 msgcat 未启用）",
        "offline": True,
        "ui_version": "第二版正式界面",
        "formal_entry_ui_v2": True,
        "core_interface_probe": core_interface_probe() if source.exists() else None,
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main(default_mode: str = "audit", preview_state: str | None = None) -> None:
    root = Tk()
    BomConverterAppV2(root, default_mode, preview_state)
    root.mainloop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="BOM 转换器 UI 第二版独立预览")
    parser.add_argument("--mode", choices=("quick", "audit"), default="audit")
    parser.add_argument("--preview-state", choices=("quick", "audit", "manual", "results"))
    parser.add_argument("--smoke-test", type=Path)
    args = parser.parse_args()
    root = Tk()
    app = BomConverterAppV2(root, args.mode, args.preview_state)
    if args.smoke_test:
        write_smoke_result(args.smoke_test, app)
        if app.preview_dialog:
            app.preview_dialog.window.destroy()
        root.destroy()
    else:
        root.mainloop()

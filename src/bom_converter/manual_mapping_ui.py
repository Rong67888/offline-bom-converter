from __future__ import annotations

from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Toplevel, messagebox
from tkinter import ttk
from typing import Callable

from .analyzer import analyze_manual_workbook, analyze_workbook
from .confirmed_mapping import FIELD_LABELS, PROCESS_LABELS, configs_from_analysis, transform_confirmed_workbook
from .mapping_memory import MappingRuleStore
from .models import ColumnMappingConfig, NameGenerationRule, WorkbookAnalysis
from .name_rules import infer_name_rule_columns
from .profiles import TARGET_HEADERS
from .xlsx_reader import XlsxReader


PROCESS_DISPLAY = {label: key for key, label in PROCESS_LABELS.items()}
PROCESS_OPTIONS = [f"{label}（{key}）" for key, label in PROCESS_LABELS.items()]
DISPLAY_TO_PROCESS = {display: display.rsplit("（", 1)[1][:-1] for display in PROCESS_OPTIONS}
FIELD_OPTIONS = ["不指定"] + [f"{FIELD_LABELS.get(field, field)}（{field}）" for field in sorted(FIELD_LABELS)]
DISPLAY_TO_FIELD = {display: (None if display == "不指定" else display.rsplit("（", 1)[1][:-1]) for display in FIELD_OPTIONS}
UNIT_OPTIONS = ["", "kg", "g", "mg", "mm", "cm", "m", "mm²", "cm²", "m²"]
SLOT_OPTIONS = [""] + [str(value) for value in range(1, 8)]


def _process_display(process: str) -> str:
    return f"{PROCESS_LABELS.get(process, process)}（{process}）"


def _field_display(field: str | None) -> str:
    return "不指定" if not field else f"{FIELD_LABELS.get(field, field)}（{field}）"


class ManualMappingDialog:
    """Small Tk fallback editor for unfamiliar workbook structures."""

    def __init__(
        self,
        parent: object,
        source: Path,
        store: MappingRuleStore,
        on_convert: Callable[[WorkbookAnalysis, dict[int, str | None], dict[str, str], dict[int, ColumnMappingConfig]], None],
        initial_sheet: str | None = None,
    ) -> None:
        self.source = source
        self.store = store
        self.on_convert = on_convert
        self.analysis: WorkbookAnalysis | None = None
        self.configs: dict[int, ColumnMappingConfig] = {}
        self.automatic_analysis: WorkbookAnalysis | None = None
        try:
            self.automatic_analysis = analyze_workbook(source, mode="audit", mapping_memory=store)
        except Exception:
            self.automatic_analysis = None
        self.name_rule = (
            self.automatic_analysis.name_rule
            if self.automatic_analysis and self.automatic_analysis.name_rule
            else NameGenerationRule()
        )
        self.window = Toplevel(parent)
        self.window.title(f"手动配置格式 — {source.name}")
        self.window.geometry("1380x820")
        self.window.minsize(1120, 700)

        with XlsxReader(source) as reader:
            sheet_names = reader.sheet_names
        suggested_sheet = initial_sheet if initial_sheet in sheet_names else (self.automatic_analysis.sheet_name if self.automatic_analysis else sheet_names[0])
        suggested_start = self.automatic_analysis.header_start_row if self.automatic_analysis else 1
        suggested_end = self.automatic_analysis.header_end_row if self.automatic_analysis else 1
        suggested_data = self.automatic_analysis.data_start_row if self.automatic_analysis else 2
        self.automatic_range = (suggested_sheet, suggested_start, suggested_end, suggested_data)
        self.sheet_name = StringVar(value=suggested_sheet)
        self.header_start = StringVar(value=str(suggested_start))
        self.header_end = StringVar(value=str(suggested_end))
        self.data_start = StringVar(value=str(suggested_data))
        self.process = StringVar(value=_process_display("ignore"))
        self.field = StringVar(value="不指定")
        self.unit = StringVar(value="")
        self.image_slot = StringVar(value="")
        self.level_group = StringVar(value="component_asm_level")
        self.level_value = StringVar(value="")
        self.default_value = StringVar(value="")
        self.reason = StringVar(value="请先选择工作表和表头范围，然后点击“读取并预览组合表头”。")
        self._build(sheet_names)
        self._analyze()

    def _build(self, sheet_names: list[str]) -> None:
        outer = ttk.Frame(self.window, padding=12)
        outer.pack(fill=BOTH, expand=True)
        ttk.Label(
            outer,
            text="手动配置格式：只保存表头结构和规则，不保存示例值、零件名称、编号或图片",
            style="Header.TLabel",
        ).pack(fill="x", pady=(0, 8))

        region = ttk.LabelFrame(outer, text="1. 选择工作表和表头区域", padding=8)
        region.pack(fill="x", pady=(0, 8))
        ttk.Label(region, text="工作表").grid(row=0, column=0, sticky="w")
        ttk.Combobox(region, textvariable=self.sheet_name, values=sheet_names, state="readonly", width=28).grid(row=0, column=1, padx=5)
        for index, (label, variable) in enumerate((("表头开始行", self.header_start), ("表头结束行", self.header_end), ("数据开始行", self.data_start)), 2):
            ttk.Label(region, text=label).grid(row=0, column=index * 2 - 2, sticky="e", padx=(10, 2))
            ttk.Entry(region, textvariable=variable, width=7).grid(row=0, column=index * 2 - 1)
        ttk.Button(region, text="读取并预览组合表头", command=self._analyze, style="Primary.TButton").grid(row=0, column=8, padx=(12, 0))
        ttk.Label(region, textvariable=self.reason, wraplength=1240).grid(row=1, column=0, columnspan=9, sticky="w", pady=(7, 0))

        table_frame = ttk.LabelFrame(outer, text="2. 预览完整表头并逐列设置处理方式", padding=8)
        table_frame.pack(fill=BOTH, expand=True, pady=(0, 8))
        columns = ("column", "path", "process", "field", "unit", "slot", "group", "confidence")
        self.table = ttk.Treeview(table_frame, columns=columns, show="headings", height=14, selectmode="browse")
        for name, title, width in (
            ("column", "来源列", 70), ("path", "完整表头路径", 390), ("process", "处理方式", 160),
            ("field", "通用字段", 220), ("unit", "来源单位", 80), ("slot", "图片槽", 65),
            ("group", "层级组/级数", 160), ("confidence", "置信度", 75),
        ):
            self.table.heading(name, text=title)
            self.table.column(name, width=width, anchor="w" if name not in {"column", "unit", "slot", "confidence"} else "center")
        self.table.pack(fill=BOTH, expand=True)
        self.table.bind("<<TreeviewSelect>>", self._load_selected)

        editor = ttk.Frame(table_frame)
        editor.pack(fill="x", pady=(8, 0))
        for col, (label, variable, values, width) in enumerate((
            ("处理方式", self.process, PROCESS_OPTIONS, 23),
            ("通用字段", self.field, FIELD_OPTIONS, 31),
            ("来源单位", self.unit, UNIT_OPTIONS, 9),
            ("图片槽位", self.image_slot, SLOT_OPTIONS, 8),
        )):
            ttk.Label(editor, text=label).grid(row=0, column=col * 2, sticky="e", padx=(4, 2))
            ttk.Combobox(editor, textvariable=variable, values=values, state="readonly", width=width).grid(row=0, column=col * 2 + 1)
        ttk.Label(editor, text="层级组").grid(row=1, column=0, sticky="e", pady=(5, 0))
        ttk.Entry(editor, textvariable=self.level_group, width=22).grid(row=1, column=1, pady=(5, 0))
        ttk.Label(editor, text="层级数字").grid(row=1, column=2, sticky="e", pady=(5, 0))
        ttk.Entry(editor, textvariable=self.level_value, width=10).grid(row=1, column=3, pady=(5, 0))
        ttk.Label(editor, text="固定默认值").grid(row=1, column=4, sticky="e", pady=(5, 0))
        ttk.Entry(editor, textvariable=self.default_value, width=24).grid(row=1, column=5, pady=(5, 0))
        ttk.Button(editor, text="应用到所选列", command=self._apply_selected).grid(row=1, column=7, padx=(8, 0), pady=(5, 0))

        actions = ttk.Frame(outer)
        actions.pack(fill="x")
        ttk.Button(actions, text="预览前几行转换结果", command=self._preview_rows).pack(side=LEFT)
        ttk.Button(actions, text="仅保存规则", command=self._save_only).pack(side=LEFT, padx=5)
        ttk.Button(actions, text="转换一次", command=lambda: self._convert(False), style="Primary.TButton").pack(side=RIGHT)
        ttk.Button(actions, text="转换并保存规则", command=lambda: self._convert(True)).pack(side=RIGHT, padx=5)

    def _parse_rows(self) -> tuple[int, int, int]:
        try:
            values = (int(self.header_start.get()), int(self.header_end.get()), int(self.data_start.get()))
        except ValueError as exc:
            raise ValueError("表头和数据行必须输入正整数") from exc
        if min(values) < 1:
            raise ValueError("表头和数据行必须从 1 开始")
        return values

    def _analyze(self) -> None:
        try:
            start, end, data = self._parse_rows()
            analysis = analyze_manual_workbook(self.source, self.sheet_name.get(), start, end, data, self.store)
        except Exception as exc:
            messagebox.showerror("无法读取表头区域", str(exc), parent=self.window)
            return
        self.analysis = analysis
        mappings = {item.source_col: item.universal_field for item in analysis.mappings}
        self.configs = configs_from_analysis(analysis, mappings)
        next_slot = 1
        for config in self.configs.values():
            if config.process_type == "image" and not config.image_slot:
                config.image_slot = next_slot
                next_slot = min(next_slot + 1, 7)
        self.name_rule = infer_name_rule_columns(analysis.mappings, self.configs, analysis.name_rule or self.name_rule)
        self.analysis.name_rule = self.name_rule
        self._refresh_table()
        self.reason.set(
            f"已组合 {analysis.header_start_row}–{analysis.header_end_row} 行，数据从第 {analysis.data_start_row} 行开始；"
            f"表头置信度 {analysis.profile_confidence:.0%}。" + "；".join(analysis.header_region.confidence_reasons)
        )

    def _refresh_table(self) -> None:
        for item in self.table.get_children():
            self.table.delete(item)
        if not self.analysis:
            return
        for decision in self.analysis.mappings:
            config = self.configs[decision.source_col]
            group = config.level_group or ""
            if config.level_value:
                group = f"{group} / {config.level_value}"
            self.table.insert("", END, iid=str(decision.source_col), values=(
                f"{decision.column_letter} / {decision.source_col}",
                decision.source_header,
                _process_display(config.process_type),
                _field_display(config.universal_field),
                config.unit or "",
                config.image_slot or "",
                group,
                f"{decision.confidence:.0%}",
            ))

    def _load_selected(self, _event: object = None) -> None:
        selected = self.table.selection()
        if not selected:
            return
        config = self.configs[int(selected[0])]
        self.process.set(_process_display(config.process_type))
        self.field.set(_field_display(config.universal_field))
        self.unit.set(config.unit or "")
        self.image_slot.set(str(config.image_slot or ""))
        self.level_group.set(config.level_group or "component_asm_level")
        self.level_value.set(str(config.level_value or ""))
        self.default_value.set("" if config.default_value is None else str(config.default_value))

    def _apply_selected(self) -> None:
        selected = self.table.selection()
        if not selected:
            messagebox.showwarning("未选择列", "请先在表格中选择一列。", parent=self.window)
            return
        column = int(selected[0])
        process = DISPLAY_TO_PROCESS.get(self.process.get(), "ignore")
        field = DISPLAY_TO_FIELD.get(self.field.get())
        try:
            slot = int(self.image_slot.get()) if self.image_slot.get() else None
            level_value = int(self.level_value.get()) if self.level_value.get() else None
        except ValueError:
            messagebox.showerror("设置错误", "图片槽位和层级数字必须是整数。", parent=self.window)
            return
        if process == "ignore":
            field = None
        if process == "level_group":
            field = "level"
        if process == "image":
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

    def _mapping_payload(self) -> tuple[dict[int, str | None], dict[str, str]]:
        if not self.analysis:
            raise ValueError("请先读取并预览表头区域")
        self.analysis.name_rule = self.name_rule
        mappings = {
            column: (None if config.process_type == "ignore" else config.universal_field)
            for column, config in self.configs.items()
        }
        if not any(config.process_type != "ignore" for config in self.configs.values()):
            raise ValueError("不能忽略所有列")
        units = {
            config.universal_field: config.unit
            for config in self.configs.values()
            if config.universal_field and config.unit
        }
        return mappings, units

    def _preview_rows(self) -> None:
        try:
            mappings, units = self._mapping_payload()
            rows, issues = transform_confirmed_workbook(self.source, self.analysis, mappings, units, self.configs)
        except Exception as exc:
            messagebox.showerror("无法预览", str(exc), parent=self.window)
            return
        preview = Toplevel(self.window)
        preview.title("前五行转换预览")
        preview.geometry("980x520")
        text = __import__("tkinter").Text(preview, wrap="word")
        text.pack(fill=BOTH, expand=True, padx=8, pady=8)
        for row in rows[:5]:
            text.insert(END, f"来源第 {row.source_row} 行\n")
            text.insert(END, "；".join(f"{TARGET_HEADERS.get(key, key)}={value}" for key, value in row.values.items() if value not in {None, ""}) + "\n")
            text.insert(END, f"图片 {len(row.images)} 张\n\n")
        warnings = [issue.message for issue in issues if issue.severity in {"warning", "error"}]
        if warnings:
            text.insert(END, "需要注意：\n" + "\n".join(f"- {item}" for item in warnings[:20]))
        text.configure(state="disabled")

    def _save_rule(self) -> None:
        mappings, units = self._mapping_payload()
        self.store.save_rule(
            self.analysis.fingerprint,
            sheet_name=self.analysis.sheet_name,
            header_rows=self.analysis.header_rows,
            header_start_row=self.analysis.header_start_row,
            header_end_row=self.analysis.header_end_row,
            data_start_row=self.analysis.data_start_row,
            headers=[item.source_header for item in self.analysis.mappings],
            header_paths=[item.header_path for item in self.analysis.mappings],
            merged_structure_summary=self.analysis.header_region.merged_structure_summary,
            mappings=mappings,
            column_configs=self.configs,
            units={str(column): config.unit for column, config in self.configs.items() if config.unit},
            ignored_columns=[column for column, config in self.configs.items() if config.process_type == "ignore"],
            name_rule=self.name_rule,
        )

    def _save_only(self) -> None:
        try:
            self._save_rule()
        except Exception as exc:
            messagebox.showerror("无法保存规则", str(exc), parent=self.window)
            return
        messagebox.showinfo("规则已保存", "已保存工作表、表头范围、完整表头路径和处理规则；没有保存 BOM 数据行或图片。", parent=self.window)

    def _convert(self, remember: bool) -> None:
        try:
            mappings, units = self._mapping_payload()
            if remember:
                self._save_rule()
        except Exception as exc:
            messagebox.showerror("配置不完整", str(exc), parent=self.window)
            return
        self.on_convert(self.analysis, mappings, units, dict(self.configs))
        self.window.destroy()

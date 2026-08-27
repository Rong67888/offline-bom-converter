from __future__ import annotations

from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, BooleanVar, StringVar, Toplevel, messagebox
from tkinter import ttk
from typing import Callable

from .models import ColumnMappingConfig, NameGenerationRule, WorkbookAnalysis
from .name_rules import (
    STRATEGY_LABELS,
    TEMPLATE_PRESETS,
    generate_name_from_columns,
    infer_name_rule_columns,
)
from .xlsx_reader import XlsxReader


class NameRuleDialog:
    """Plain-language editor. Preview values stay in memory and are never saved."""

    def __init__(
        self,
        parent: object,
        source: Path,
        analysis: WorkbookAnalysis,
        configs: dict[int, ColumnMappingConfig],
        initial_rule: NameGenerationRule | None,
        on_apply: Callable[[NameGenerationRule], None],
    ) -> None:
        self.source = source
        self.analysis = analysis
        self.configs = configs
        self.on_apply = on_apply
        self.rule = infer_name_rule_columns(analysis.mappings, configs, initial_rule)
        self.window = Toplevel(parent)
        self.window.title("名称生成规则")
        self.window.geometry("1080x650")
        self.window.minsize(900, 560)
        self.window.transient(parent)

        self.strategy = StringVar(value=STRATEGY_LABELS.get(self.rule.strategy, STRATEGY_LABELS["fallback"]))
        self.preset = StringVar(value=self._preset_for_template(self.rule.template))
        self.template = StringVar(value=self.rule.template)
        self.deduplicate = BooleanVar(value=self.rule.deduplicate)
        self.column_options, self.display_to_column, self.column_to_display = self._column_options()
        self.original_col = StringVar(value=self.column_to_display.get(self.rule.original_name_col, "不使用"))
        self.standard_col = StringVar(value=self.column_to_display.get(self.rule.standard_name_col, "不使用"))
        self.gb_col = StringVar(value=self.column_to_display.get(self.rule.gb_name_col, "不使用"))
        self.spec_col = StringVar(value=self.column_to_display.get(self.rule.spec_col, "不使用"))
        self.note = StringVar(value="预览值只在当前窗口显示，不保存到本地规则。")
        self._build()
        self._refresh_preview()

    def _column_options(self) -> tuple[list[str], dict[str, int | None], dict[int | None, str]]:
        options = ["不使用"]
        display_to_column: dict[str, int | None] = {"不使用": None}
        column_to_display: dict[int | None, str] = {None: "不使用"}
        for decision in self.analysis.mappings:
            display = f"{decision.column_letter} 列：{decision.source_header}"
            options.append(display)
            display_to_column[display] = decision.source_col
            column_to_display[decision.source_col] = display
        return options, display_to_column, column_to_display

    @staticmethod
    def _preset_for_template(template: str) -> str:
        for name, value in TEMPLATE_PRESETS.items():
            if name != "自定义" and value == template:
                return name
        return "自定义"

    def _build(self) -> None:
        outer = ttk.Frame(self.window, padding=12)
        outer.pack(fill=BOTH, expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        ttk.Label(outer, text="名称生成规则", style="Header.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text="选择用于名称的来源列和组合方式；保存时只记录列号、策略和模板，不保存下面显示的具体名称。",
        ).grid(row=1, column=0, sticky="w", pady=(2, 8))

        settings = ttk.LabelFrame(outer, text="1  选择来源列和策略", padding=10)
        settings.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        for column in range(4):
            settings.columnconfigure(column * 2 + 1, weight=1)
        controls = (
            ("原名称列", self.original_col),
            ("标准件名称列", self.standard_col),
            ("GB 名称列", self.gb_col),
            ("规格列", self.spec_col),
        )
        for index, (label, variable) in enumerate(controls):
            row, pair = divmod(index, 2)
            column = pair * 4
            ttk.Label(settings, text=label).grid(row=row, column=column, sticky="e", padx=(0, 4), pady=4)
            box = ttk.Combobox(settings, textvariable=variable, values=self.column_options, state="readonly", width=34)
            box.grid(row=row, column=column + 1, sticky="ew", padx=(0, 14), pady=4)
            box.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())

        ttk.Label(settings, text="使用策略").grid(row=2, column=0, sticky="e", padx=(0, 4), pady=4)
        strategy_box = ttk.Combobox(
            settings,
            textvariable=self.strategy,
            values=list(STRATEGY_LABELS.values()),
            state="readonly",
            width=24,
        )
        strategy_box.grid(row=2, column=1, sticky="w", pady=4)
        strategy_box.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())
        ttk.Checkbutton(
            settings,
            text="删除重复字段，并在预览中说明",
            variable=self.deduplicate,
            command=self._refresh_preview,
        ).grid(row=2, column=4, columnspan=2, sticky="w", pady=4)

        ttk.Label(settings, text="组合模板").grid(row=3, column=0, sticky="e", padx=(0, 4), pady=4)
        preset_box = ttk.Combobox(
            settings,
            textvariable=self.preset,
            values=list(TEMPLATE_PRESETS),
            state="readonly",
            width=16,
        )
        preset_box.grid(row=3, column=1, sticky="w", pady=4)
        preset_box.bind("<<ComboboxSelected>>", self._preset_changed)
        ttk.Entry(settings, textvariable=self.template).grid(row=3, column=4, columnspan=3, sticky="ew", pady=4)
        ttk.Button(settings, text="刷新名称预览", command=self._refresh_preview).grid(row=3, column=7, padx=(8, 0))
        ttk.Label(
            settings,
            text="可用占位符：{原名称}、{名称}或{标准件名称}、{GB}、{规格}。示例：{名称}（{GB}）{规格}",
        ).grid(row=4, column=0, columnspan=8, sticky="w", pady=(4, 0))

        preview_card = ttk.LabelFrame(outer, text="2  来源字段—最终名称预览（至少五行，只显示不保存）", padding=8)
        preview_card.grid(row=3, column=0, sticky="nsew", pady=(0, 8))
        preview_card.columnconfigure(0, weight=1)
        preview_card.rowconfigure(0, weight=1)
        columns = ("row", "original", "standard", "gb", "spec", "final", "removed")
        self.preview = ttk.Treeview(preview_card, columns=columns, show="headings", height=7)
        for name, title, width in (
            ("row", "来源行", 65),
            ("original", "原名称", 130),
            ("standard", "标准件名称", 130),
            ("gb", "GB 名称", 120),
            ("spec", "规格", 110),
            ("final", "最终名称", 230),
            ("removed", "去重说明", 240),
        ):
            self.preview.heading(name, text=title)
            self.preview.column(name, width=width, minwidth=60, stretch=name in {"final", "removed"})
        ybar = ttk.Scrollbar(preview_card, orient="vertical", command=self.preview.yview)
        xbar = ttk.Scrollbar(preview_card, orient="horizontal", command=self.preview.xview)
        self.preview.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        self.preview.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")

        actions = ttk.Frame(outer)
        actions.grid(row=4, column=0, sticky="ew")
        ttk.Label(actions, textvariable=self.note).pack(side=LEFT)
        ttk.Button(actions, text="取消", command=self.window.destroy).pack(side=RIGHT)
        ttk.Button(actions, text="应用名称规则", command=self._apply, style="Primary.TButton").pack(side=RIGHT, padx=6)

    def _preset_changed(self, _event: object = None) -> None:
        selected = self.preset.get()
        if selected != "自定义":
            self.template.set(TEMPLATE_PRESETS[selected])
        self._refresh_preview()

    def current_rule(self) -> NameGenerationRule:
        strategy_by_label = {label: key for key, label in STRATEGY_LABELS.items()}
        template = self.template.get().strip()
        if not template:
            raise ValueError("组合模板不能为空")
        return NameGenerationRule(
            strategy=strategy_by_label.get(self.strategy.get(), "fallback"),
            original_name_col=self.display_to_column.get(self.original_col.get()),
            standard_name_col=self.display_to_column.get(self.standard_col.get()),
            gb_name_col=self.display_to_column.get(self.gb_col.get()),
            spec_col=self.display_to_column.get(self.spec_col.get()),
            template=template,
            deduplicate=bool(self.deduplicate.get()),
        )

    def _refresh_preview(self) -> None:
        try:
            rule = self.current_rule()
        except ValueError as exc:
            self.note.set(str(exc))
            return
        for item in self.preview.get_children():
            self.preview.delete(item)
        with XlsxReader(self.source) as reader:
            sheet = reader.read_sheet(self.analysis.sheet_name)
        for source_row in range(self.analysis.data_start_row, self.analysis.data_start_row + 5):
            value = lambda column: sheet.get(source_row, column) if column else None
            result = generate_name_from_columns(rule, lambda column: sheet.get(source_row, column))
            self.preview.insert(
                "",
                END,
                values=(
                    source_row,
                    value(rule.original_name_col) or "",
                    value(rule.standard_name_col) or "",
                    value(rule.gb_name_col) or "",
                    value(rule.spec_col) or "",
                    result.final_name or "",
                    "；".join(result.removed_duplicates) or "未删除重复字段",
                ),
            )
        self.note.set("预览已刷新；空字段会跳过，名称中的 × 和中文括号会保留。")

    def _apply(self) -> None:
        try:
            rule = self.current_rule()
        except ValueError as exc:
            messagebox.showerror("名称规则不完整", str(exc), parent=self.window)
            return
        self.rule = rule
        self.on_apply(rule)
        self.window.destroy()

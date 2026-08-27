from __future__ import annotations

import os
import queue
import sys
import threading
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, StringVar, Tk, filedialog, messagebox
from tkinter import ttk

from .analyzer import analyze_workbook
from .confirmed_mapping import FIELD_LABELS, configs_from_analysis, infer_confirmed_units
from .converter import convert_confirmed_file, convert_file
from .manual_mapping_ui import ManualMappingDialog
from .mapping_memory import MappingRuleStore
from .models import ColumnMappingConfig, ConversionResult, WorkbookAnalysis
from .profiles import TARGET_HEADERS


def _resource_root() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parents[2]


def _install_root() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[2]


def _mapping_store_path() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(os.environ.get("LOCALAPPDATA", _install_root())) / "BOMConverter"
    else:
        base = _install_root() / "config"
    return base / "mapping_rules.json"


def _field_display(field: str | None) -> str:
    if not field:
        return "忽略"
    return f"{field} — {FIELD_LABELS.get(field, field)}"


FIELD_OPTIONS = ["忽略"] + [_field_display(field) for field in sorted(FIELD_LABELS)]
DISPLAY_TO_FIELD = {display: (None if display == "忽略" else display.split(" — ", 1)[0]) for display in FIELD_OPTIONS}


class BomConverterApp:
    def __init__(self, root: Tk, default_mode: str = "quick"):
        self.root = root
        self.root.title("BOM 格式转换工具（离线版）")
        self.root.geometry("1240x860")
        self.root.minsize(1080, 720)
        self.sources: list[Path] = []
        self.analyses: dict[Path, WorkbookAnalysis] = {}
        self.mapping_choices: dict[Path, dict[int, str | None]] = {}
        self.manual_configs: dict[Path, dict[int, ColumnMappingConfig]] = {}
        self.current_source: Path | None = None
        self.mode = StringVar(value=default_mode)
        self.template = StringVar(value=str(_resource_root() / "assets" / "public_demo_template.xlsx"))
        self.output_dir = StringVar(value=str(_install_root() / "outputs"))
        self.status_text = StringVar(value="就绪：请先添加来源 BOM")
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.mapping_store = MappingRuleStore(_mapping_store_path())
        self._mapping_editor: ttk.Combobox | None = None
        self._build_style()
        self._build()
        self.root.after(120, self._poll)

    def _build_style(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background="#F3F6F9")
        style.configure("TLabel", background="#F3F6F9", foreground="#243447")
        style.configure("Header.TLabel", font=("Microsoft YaHei UI", 16, "bold"), foreground="#17365D")
        style.configure("Step.TLabel", font=("Microsoft YaHei UI", 10, "bold"), foreground="#FFFFFF", background="#426B8E", padding=(10, 5))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 9, "bold"), foreground="#FFFFFF", background="#2F6B9A")
        style.map("Primary.TButton", background=[("active", "#25577D")])
        style.configure("Treeview", rowheight=27, background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#243447")
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 9, "bold"), background="#D8E4EE", foreground="#17365D")
        style.map("Treeview", background=[("selected", "#BDD7EE")], foreground=[("selected", "#17365D")])
        style.configure("TLabelframe", background="#F3F6F9")
        style.configure("TLabelframe.Label", background="#F3F6F9", foreground="#17365D", font=("Microsoft YaHei UI", 9, "bold"))

    def _build(self) -> None:
        outer = ttk.Frame(self.root, padding=14)
        outer.pack(fill=BOTH, expand=True)

        title_row = ttk.Frame(outer)
        title_row.pack(fill="x", pady=(0, 8))
        ttk.Label(title_row, text="BOM 格式转换工具", style="Header.TLabel").pack(side=LEFT)
        ttk.Label(title_row, text="所有 Excel 仅在本机离线处理").pack(side=RIGHT)

        mode_row = ttk.Frame(outer)
        mode_row.pack(fill="x", pady=(0, 8))
        ttk.Label(mode_row, text="工作模式：").pack(side=LEFT)
        ttk.Radiobutton(mode_row, text="快速模式", variable=self.mode, value="quick").pack(side=LEFT, padx=5)
        ttk.Radiobutton(mode_row, text="通用审核模式", variable=self.mode, value="audit").pack(side=LEFT, padx=5)
        ttk.Label(mode_row, text="添加文件  →  检查映射  →  开始转换  →  查看结果", style="Step.TLabel").pack(side=RIGHT)

        files = ttk.LabelFrame(outer, text="步骤 1：添加文件并识别格式", padding=8)
        files.pack(fill="x", pady=(0, 8))
        file_columns = ("file", "profile", "rows", "images", "status")
        self.file_table = ttk.Treeview(files, columns=file_columns, show="headings", height=5, selectmode="browse")
        for column, text, width in (
            ("file", "文件名", 310), ("profile", "识别格式", 260), ("rows", "数据行", 80),
            ("images", "图片", 70), ("status", "状态", 220),
        ):
            self.file_table.heading(column, text=text)
            self.file_table.column(column, width=width, anchor="w" if column in {"file", "profile", "status"} else "center")
        self.file_table.tag_configure("ready", foreground="#1B5E20")
        self.file_table.tag_configure("review", foreground="#9A6700")
        self.file_table.tag_configure("error", foreground="#B42318")
        self.file_table.tag_configure("pending", foreground="#52677A")
        self.file_table.pack(side=LEFT, fill="x", expand=True)
        self.file_table.bind("<<TreeviewSelect>>", self._on_file_selected)
        file_buttons = ttk.Frame(files)
        file_buttons.pack(side=RIGHT, fill="y", padx=(8, 0))
        ttk.Button(file_buttons, text="添加文件...", command=self._choose_sources, style="Primary.TButton").pack(fill="x", pady=2)
        ttk.Button(file_buttons, text="手动配置格式...", command=self._open_manual_mapping).pack(fill="x", pady=2)
        ttk.Button(file_buttons, text="重新检查", command=self._reanalyze_selected).pack(fill="x", pady=2)
        ttk.Button(file_buttons, text="清空列表", command=self._clear_sources).pack(fill="x", pady=2)

        mapping_frame = ttk.LabelFrame(outer, text="步骤 2：检查和确认字段映射（双击“通用字段”可下拉修改）", padding=8)
        mapping_frame.pack(fill=BOTH, expand=True, pady=(0, 8))
        mapping_columns = ("source", "sample", "field", "target", "confidence", "status")
        self.mapping_table = ttk.Treeview(mapping_frame, columns=mapping_columns, show="headings", height=10, selectmode="browse")
        for column, text, width in (
            ("source", "来源列", 190), ("sample", "示例值（只显示，不保存）", 260),
            ("field", "通用字段", 265), ("target", "目标模板列", 230),
            ("confidence", "置信度", 75), ("status", "处理状态", 85),
        ):
            self.mapping_table.heading(column, text=text)
            self.mapping_table.column(column, width=width, anchor="w" if column not in {"confidence", "status"} else "center")
        self.mapping_table.pack(fill=BOTH, expand=True)
        self.mapping_table.bind("<Double-1>", self._edit_mapping)

        mapping_actions = ttk.Frame(mapping_frame)
        mapping_actions.pack(fill="x", pady=(7, 0))
        ttk.Button(mapping_actions, text="确认并转换", command=lambda: self._confirm_and_convert(False), style="Primary.TButton").pack(side=LEFT)
        ttk.Button(mapping_actions, text="确认、转换并记住", command=lambda: self._confirm_and_convert(True)).pack(side=LEFT, padx=5)
        ttk.Button(mapping_actions, text="仅保存规则", command=self._save_rule_only).pack(side=LEFT, padx=5)
        ttk.Button(mapping_actions, text="恢复自动推荐", command=self._restore_recommendations).pack(side=LEFT, padx=5)
        ttk.Label(mapping_actions, text=f"规则库：{self.mapping_store.path}").pack(side=RIGHT)

        settings = ttk.LabelFrame(outer, text="步骤 3：选择模板、输出目录并开始转换", padding=8)
        settings.pack(fill="x", pady=(0, 8))
        ttk.Label(settings, text="标准模板").grid(row=0, column=0, sticky="w")
        ttk.Entry(settings, textvariable=self.template).grid(row=0, column=1, sticky="ew", padx=7)
        ttk.Button(settings, text="浏览...", command=self._choose_template).grid(row=0, column=2)
        ttk.Label(settings, text="输出目录").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(settings, textvariable=self.output_dir).grid(row=1, column=1, sticky="ew", padx=7, pady=(6, 0))
        ttk.Button(settings, text="浏览...", command=self._choose_output).grid(row=1, column=2, pady=(6, 0))
        settings.columnconfigure(1, weight=1)
        actions = ttk.Frame(settings)
        actions.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8, 0))
        self.convert_button = ttk.Button(actions, text="开始转换", command=self._start_conversion, style="Primary.TButton")
        self.convert_button.pack(side=LEFT)
        ttk.Button(actions, text="打开输出目录", command=self._open_output_dir).pack(side=LEFT, padx=6)
        self.progress = ttk.Progressbar(actions, mode="determinate", maximum=100)
        self.progress.pack(side=LEFT, fill="x", expand=True, padx=10)
        ttk.Label(actions, textvariable=self.status_text).pack(side=RIGHT)

        log_frame = ttk.LabelFrame(outer, text="步骤 4：详细日志（排错用）", padding=6)
        log_frame.pack(fill="x")
        self.log = __import__("tkinter").Text(log_frame, wrap="word", state="disabled", height=5, background="#F8FAFC", foreground="#334155")
        self.log.pack(fill="x")

    def _choose_sources(self) -> None:
        files = filedialog.askopenfilenames(filetypes=[("Excel 工作簿", "*.xlsx")])
        paths: list[Path] = []
        for name in files:
            path = Path(name).resolve()
            if path.name.startswith("~$") or path in self.sources:
                continue
            self.sources.append(path)
            paths.append(path)
            self.file_table.insert("", END, iid=str(path), values=(path.name, "等待识别", "-", "-", "分析中"), tags=("pending",))
        if paths:
            self.status_text.set(f"正在分析 {len(paths)} 个文件...")
            mode = self.mode.get()
            threading.Thread(target=self._worker_analyze, args=(paths, mode), daemon=True).start()

    def _clear_sources(self) -> None:
        self.sources.clear()
        self.analyses.clear()
        self.mapping_choices.clear()
        self.manual_configs.clear()
        self.current_source = None
        for item in self.file_table.get_children():
            self.file_table.delete(item)
        for item in self.mapping_table.get_children():
            self.mapping_table.delete(item)
        self.status_text.set("已清空；请添加来源 BOM")

    def _reanalyze_selected(self) -> None:
        source = self._selected_source()
        if not source:
            messagebox.showwarning("未选择文件", "请先在文件表格中选择一个文件。")
            return
        self.file_table.item(str(source), values=(source.name, "等待识别", "-", "-", "重新分析中"), tags=("pending",))
        threading.Thread(target=self._worker_analyze, args=([source], self.mode.get()), daemon=True).start()

    def _open_manual_mapping(self) -> None:
        source = self._selected_source()
        if not source:
            messagebox.showwarning("未选择文件", "请先在文件表格中选择一份来源工作簿。")
            return
        ManualMappingDialog(self.root, source, self.mapping_store, lambda analysis, mappings, units, configs: self._accept_manual_conversion(source, analysis, mappings, units, configs))

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
        status, tag = self._analysis_status(analysis)
        self.file_table.item(
            str(source),
            values=(source.name, f"手动配置：{analysis.sheet_name} {analysis.header_start_row}-{analysis.header_end_row} 行", analysis.data_row_count, analysis.image_count, status),
            tags=(tag,),
        )
        self._load_mapping_table(source, restore=False)
        self._run_jobs([("custom", source, analysis, mappings, units, configs)])

    def _worker_analyze(self, paths: list[Path], mode: str) -> None:
        total = len(paths)
        for index, path in enumerate(paths, 1):
            try:
                analysis = analyze_workbook(path, mode=mode, mapping_memory=self.mapping_store)
                self.events.put(("analysis", (path, analysis)))
            except Exception as exc:
                self.events.put(("file_error", (path, str(exc))))
            self.events.put(("progress", index * 100 / max(total, 1)))
        self.events.put(("done", "文件分析完成"))

    def _analysis_status(self, analysis: WorkbookAnalysis) -> tuple[str, str]:
        issue_codes = {issue.code for issue in analysis.issues}
        if "SIMILAR_SAVED_FORMAT" in issue_codes:
            return "相似历史格式，需检查差异", "review"
        if analysis.profile_id == "generic" and analysis.requires_review:
            return "未知/变化格式，需确认映射", "review"
        if analysis.profile_id == "generic":
            return "已应用本地记忆，可转换", "ready"
        if analysis.requires_review and self.mode.get() == "audit":
            return "已知格式，等待审核确认", "review"
        if analysis.requires_review:
            return "低置信度，需检查", "review"
        return "完整匹配，可直接转换", "ready"

    def _on_file_selected(self, _event: object = None) -> None:
        source = self._selected_source()
        if source and source in self.analyses:
            self.current_source = source
            self._load_mapping_table(source, restore=False)

    def _selected_source(self) -> Path | None:
        selected = self.file_table.selection()
        return Path(selected[0]) if selected else None

    def _load_mapping_table(self, source: Path, restore: bool) -> None:
        analysis = self.analyses[source]
        if restore or source not in self.mapping_choices:
            self.mapping_choices[source] = {
                mapping.source_col: mapping.universal_field
                for mapping in analysis.mappings
                if mapping.status != "ignored" or mapping.universal_field
            }
            for mapping in analysis.mappings:
                self.mapping_choices[source].setdefault(mapping.source_col, None)
        for item in self.mapping_table.get_children():
            self.mapping_table.delete(item)
        choices = self.mapping_choices[source]
        for mapping in analysis.mappings:
            field = choices.get(mapping.source_col)
            target = TARGET_HEADERS.get(field or "", "—")
            sample = "" if mapping.sample_value is None else str(mapping.sample_value)
            if len(sample) > 80:
                sample = sample[:77] + "..."
            self.mapping_table.insert(
                "", END, iid=str(mapping.source_col),
                values=(
                    f"{mapping.source_col}: {mapping.source_header}", sample,
                    _field_display(field), target.replace("\n", " / "),
                    f"{mapping.confidence:.0%}", "使用" if field else "忽略",
                ),
            )

    def _edit_mapping(self, event: object) -> None:
        if not self.current_source:
            return
        region = self.mapping_table.identify("region", event.x, event.y)
        column = self.mapping_table.identify_column(event.x)
        item = self.mapping_table.identify_row(event.y)
        if region != "cell" or column != "#3" or not item:
            return
        bbox = self.mapping_table.bbox(item, column)
        if not bbox:
            return
        x, y, width, height = bbox
        current = self.mapping_table.set(item, "field")
        editor = ttk.Combobox(self.mapping_table, values=FIELD_OPTIONS, state="readonly")
        editor.set(current if current in FIELD_OPTIONS else "忽略")
        editor.place(x=x, y=y, width=max(width, 280), height=height)
        editor.focus_set()
        editor.bind("<<ComboboxSelected>>", lambda _e: self._commit_mapping_editor(item, editor))
        editor.bind("<FocusOut>", lambda _e: editor.destroy())
        self._mapping_editor = editor

    def _commit_mapping_editor(self, item: str, editor: ttk.Combobox) -> None:
        if not self.current_source:
            editor.destroy()
            return
        display = editor.get()
        field = DISPLAY_TO_FIELD.get(display)
        column = int(item)
        self.mapping_choices[self.current_source][column] = field
        if self.current_source in self.manual_configs:
            config = self.manual_configs[self.current_source][column]
            config.universal_field = field
            if field is None:
                config.process_type = "ignore"
            elif config.process_type == "ignore":
                config.process_type = "direct"
        self.mapping_table.set(item, "field", _field_display(field))
        self.mapping_table.set(item, "target", TARGET_HEADERS.get(field or "", "—").replace("\n", " / "))
        self.mapping_table.set(item, "status", "使用" if field else "忽略")
        editor.destroy()

    def _restore_recommendations(self) -> None:
        if not self.current_source:
            return
        analysis = analyze_workbook(self.current_source, mode="audit")
        self.analyses[self.current_source] = analysis
        self.mapping_choices.pop(self.current_source, None)
        self.manual_configs.pop(self.current_source, None)
        self._load_mapping_table(self.current_source, restore=True)
        self._append("已恢复程序自动推荐；尚未写入规则库。")

    def _save_current_rule(self) -> bool:
        source = self.current_source
        if not source:
            messagebox.showwarning("未选择文件", "请先选择要保存规则的文件。")
            return False
        analysis = self.analyses[source]
        mappings = self.mapping_choices[source]
        if not any(mappings.values()):
            messagebox.showwarning("映射为空", "至少需要保留一个通用字段。")
            return False
        units = infer_confirmed_units(analysis, mappings)
        self.mapping_store.save_rule(
            analysis.fingerprint,
            sheet_name=analysis.sheet_name,
            header_rows=analysis.header_rows,
            header_start_row=analysis.header_start_row,
            header_end_row=analysis.header_end_row,
            data_start_row=analysis.data_start_row,
            headers=[mapping.source_header for mapping in analysis.mappings],
            header_paths=[mapping.header_path for mapping in analysis.mappings],
            merged_structure_summary=analysis.header_region.merged_structure_summary if analysis.header_region else [],
            mappings=mappings,
            units=units,
            ignored_columns=[column for column, field in mappings.items() if not field],
            column_configs=self.manual_configs.get(source) or configs_from_analysis(analysis, mappings, units),
            name_rule=analysis.name_rule,
        )
        self._append(f"已保存本地规则：{self.mapping_store.path}")
        self._append("规则只包含规范化表头、列号、字段、单位和忽略列；不包含示例值、零件数据或图片。")
        return True

    def _save_rule_only(self) -> None:
        if self._save_current_rule():
            messagebox.showinfo("规则已保存", "映射规则已保存在本机；没有保存 BOM 行数据或图片。")

    def _confirm_and_convert(self, remember: bool) -> None:
        source = self.current_source or self._selected_source()
        if not source:
            messagebox.showwarning("未选择文件", "请先选择一个文件并检查映射。")
            return
        if remember and not self._save_current_rule():
            return
        analysis = self.analyses[source]
        mappings = self.mapping_choices[source]
        units = infer_confirmed_units(analysis, mappings)
        configs = self.manual_configs.get(source) or configs_from_analysis(analysis, mappings, units)
        self._run_jobs([("custom", source, analysis, mappings, units, configs)])

    def _start_conversion(self) -> None:
        if not self.sources:
            messagebox.showwarning("缺少文件", "请先添加至少一份来源 BOM。")
            return
        jobs: list[tuple] = []
        for source in self.sources:
            analysis = self.analyses.get(source)
            if analysis is None:
                messagebox.showwarning("尚未分析", f"{source.name} 尚未完成格式分析。")
                return
            if analysis.profile_id != "generic":
                jobs.append(("known", source))
                continue
            if analysis.requires_review:
                self.file_table.selection_set(str(source))
                self.file_table.focus(str(source))
                self.current_source = source
                self._load_mapping_table(source, restore=False)
                messagebox.showwarning(
                    "需要确认映射",
                    f"{source.name} 是未知、变化或低置信度格式。请检查映射后使用“确认并转换”或“确认、转换并记住”。",
                )
                return
            mappings = {decision.source_col: decision.universal_field for decision in analysis.mappings}
            units = infer_confirmed_units(analysis, mappings)
            configs = self.manual_configs.get(source) or configs_from_analysis(analysis, mappings, units)
            jobs.append(("custom", source, analysis, mappings, units, configs))
        self._run_jobs(jobs)

    def _run_jobs(self, jobs: list[tuple]) -> None:
        self.convert_button.configure(state="disabled")
        self.progress["value"] = 0
        self.status_text.set(f"正在转换 {len(jobs)} 个文件...")
        # Tk 的变量只能在界面线程读取。先复制成普通字符串，再交给后台线程，
        # 避免某些 Windows/Tk 版本在批量转换时偶发卡死或报线程错误。
        template_path = self.template.get()
        output_dir = self.output_dir.get()
        threading.Thread(
            target=self._worker_convert,
            args=(jobs, template_path, output_dir),
            daemon=True,
        ).start()

    def _worker_convert(self, jobs: list[tuple], template_path: str, output_dir: str) -> None:
        total = len(jobs)
        for index, job in enumerate(jobs, 1):
            try:
                if job[0] == "known":
                    result = convert_file(job[1], template_path, output_dir, mode="quick")
                else:
                    _, source, analysis, mappings, units, configs = job
                    result = convert_confirmed_file(
                        source,
                        template_path,
                        output_dir,
                        analysis,
                        mappings,
                        units,
                        configs,
                    )
                self.events.put(("conversion", result))
            except Exception as exc:
                self.events.put(("file_error", (job[1], str(exc))))
            self.events.put(("progress", index * 100 / max(total, 1)))
        self.events.put(("done", "转换任务结束"))

    def _choose_template(self) -> None:
        name = filedialog.askopenfilename(filetypes=[("Excel 工作簿", "*.xlsx")])
        if name:
            self.template.set(name)

    def _choose_output(self) -> None:
        name = filedialog.askdirectory()
        if name:
            self.output_dir.set(name)

    def _open_output_dir(self) -> None:
        target = Path(self.output_dir.get()).resolve()
        target.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(target)  # type: ignore[attr-defined]
        except Exception as exc:
            messagebox.showerror("无法打开目录", str(exc))

    def _append(self, text: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(END, text + "\n")
        self.log.see(END)
        self.log.configure(state="disabled")

    def _handle_analysis(self, path: Path, analysis: WorkbookAnalysis) -> None:
        self.analyses[path] = analysis
        self.mapping_choices.pop(path, None)
        self.manual_configs.pop(path, None)
        status, tag = self._analysis_status(analysis)
        self.file_table.item(
            str(path),
            values=(path.name, analysis.profile_name, analysis.data_row_count, analysis.image_count, status),
            tags=(tag,),
        )
        issue = next((item for item in analysis.issues if item.code in {"SIMILAR_SAVED_FORMAT", "UNIT_CHANGED", "HEADER_ROW_CHANGED", "HEADER_REGION_CHANGED", "MULTIPLE_HEADER_SHEETS", "LOW_HEADER_CONFIDENCE"}), None)
        if issue:
            self._append(f"{path.name}: {issue.message}")
        if not self.file_table.selection():
            self.file_table.selection_set(str(path))
            self.current_source = path
            self._load_mapping_table(path, restore=False)

    def _handle_conversion(self, result: ConversionResult) -> None:
        warnings = sum(issue.severity == "warning" for issue in result.issues)
        errors = sum(issue.severity == "error" for issue in result.issues)
        tag = "error" if errors else "ready"
        status = f"已完成；warning {warnings}，error {errors}"
        if str(result.source_path) in self.file_table.get_children():
            old = self.file_table.item(str(result.source_path), "values")
            self.file_table.item(str(result.source_path), values=(*old[:4], status), tags=(tag,))
        self._append(f"完成：{result.output_path}")
        self._append(f"行 {result.output_rows}/{result.source_rows}；图片 {result.output_images}/{result.source_images}；报告 {result.report_path}")

    def _poll(self) -> None:
        try:
            while True:
                kind, value = self.events.get_nowait()
                if kind == "analysis":
                    path, analysis = value
                    self._handle_analysis(path, analysis)
                elif kind == "conversion":
                    self._handle_conversion(value)
                elif kind == "file_error":
                    path, error = value
                    if str(path) in self.file_table.get_children():
                        old = self.file_table.item(str(path), "values")
                        self.file_table.item(str(path), values=(*old[:4], f"错误：{error}"), tags=("error",))
                    self._append(f"错误：{Path(path).name}: {error}")
                elif kind == "progress":
                    self.progress["value"] = float(value)
                elif kind == "done":
                    self.convert_button.configure(state="normal")
                    self.status_text.set(str(value))
        except queue.Empty:
            pass
        self.root.after(120, self._poll)


def main(default_mode: str = "quick") -> None:
    root = Tk()
    BomConverterApp(root, default_mode)
    root.mainloop()

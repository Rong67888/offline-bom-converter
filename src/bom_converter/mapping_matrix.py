from __future__ import annotations

from dataclasses import dataclass
from tkinter import Canvas, END, Toplevel
from tkinter import ttk
from typing import Callable, Iterable


PROPERTY_ROWS = (
    ("path", "完整表头"),
    ("sample", "示例值\n（只显示，不保存）"),
    ("field", "识别为"),
    ("target", "最终写入"),
    ("confidence", "识别把握"),
    ("status", "处理状态"),
    ("recommendation", "推荐来源"),
)


@dataclass
class MatrixColumn:
    key: str
    excel_label: str
    values: dict[str, str]


class MappingMatrix(ttk.Frame):
    """Excel-like horizontal field mapping matrix with a frozen property column."""

    column_width = 170
    property_width = 150
    row_height = 48
    header_height = 38

    def __init__(
        self,
        parent: object,
        *,
        on_edit: Callable[[str], None] | None = None,
        on_select: Callable[[], None] | None = None,
        height: int = 330,
    ) -> None:
        super().__init__(parent)
        self.on_edit = on_edit
        self.on_select = on_select
        self._columns: dict[str, MatrixColumn] = {}
        self._visible: list[str] = []
        self._selected: list[str] = []
        self._focus: str | None = None
        self._row_keys = [item[0] for item in PROPERTY_ROWS]

        self.columnconfigure(1, weight=1)
        self.rowconfigure(0, weight=1)
        self.property_canvas = Canvas(self, width=self.property_width, height=height, highlightthickness=0, background="#F5F8FA")
        self.data_canvas = Canvas(self, height=height, highlightthickness=0, background="#FFFFFF")
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self._yview)
        self.hbar = ttk.Scrollbar(self, orient="horizontal", command=self.data_canvas.xview)
        self.property_canvas.grid(row=0, column=0, sticky="ns")
        self.data_canvas.grid(row=0, column=1, sticky="nsew")
        self.vbar.grid(row=0, column=2, sticky="ns")
        self.hbar.grid(row=1, column=1, sticky="ew")
        self.data_canvas.configure(xscrollcommand=self.hbar.set, yscrollcommand=self._yscroll_set)
        self.property_canvas.configure(yscrollcommand=self._yscroll_set)
        self.data_canvas.bind("<Button-1>", self._click)
        self.data_canvas.bind("<Double-1>", self._double_click)
        self.data_canvas.bind("<MouseWheel>", self._mousewheel)
        self.property_canvas.bind("<MouseWheel>", self._mousewheel)
        self._draw()

    @property
    def horizontal_scrollbar(self) -> ttk.Scrollbar:
        return self.hbar

    @property
    def vertical_scrollbar(self) -> ttk.Scrollbar:
        return self.vbar

    def set_columns(self, columns: Iterable[MatrixColumn]) -> None:
        selected = set(self._selected)
        self._columns = {item.key: item for item in columns}
        self._visible = list(self._columns)
        self._selected = [key for key in self._visible if key in selected]
        self._draw()

    def set_visible(self, keys: Iterable[str]) -> None:
        self._visible = [str(key) for key in keys if str(key) in self._columns]
        self._selected = [key for key in self._selected if key in self._visible]
        self._draw()

    def values_for(self, key: str) -> dict[str, str]:
        return dict(self._columns[str(key)].values)

    def update_column(self, key: str, **values: str) -> None:
        self._columns[str(key)].values.update({name: str(value) for name, value in values.items()})
        self._draw()

    # Small Treeview-compatible selection surface used by the existing controller/tests.
    def get_children(self, _item: str = "") -> tuple[str, ...]:
        return tuple(self._visible)

    def exists(self, key: str) -> bool:
        return str(key) in self._visible

    def selection(self) -> tuple[str, ...]:
        return tuple(self._selected)

    def selection_set(self, keys: str | Iterable[str]) -> None:
        requested = [keys] if isinstance(keys, str) else list(keys)
        self._selected = [str(key) for key in requested if str(key) in self._visible]
        self._focus = self._selected[-1] if self._selected else None
        self._draw()

    def selection_add(self, key: str) -> None:
        key = str(key)
        if key in self._visible and key not in self._selected:
            self._selected.append(key)
            self._focus = key
            self._draw()

    def focus(self, key: str | None = None) -> str:
        if key is not None and str(key) in self._visible:
            self._focus = str(key)
        return self._focus or ""

    def delete(self, key: str) -> None:
        key = str(key)
        self._columns.pop(key, None)
        self._visible = [item for item in self._visible if item != key]
        self._selected = [item for item in self._selected if item != key]
        self._draw()

    def insert(self, _parent: str, _index: str, *, iid: str, values: tuple[str, ...], tags: tuple[str, ...] = ()) -> None:
        # Compatibility order: source/path/sample/field/target/confidence/status/recommendation.
        padded = list(values) + [""] * max(0, 8 - len(values))
        column = MatrixColumn(str(iid), padded[0], {
            "path": padded[1], "sample": padded[2], "field": padded[3], "target": padded[4],
            "confidence": padded[5], "status": padded[6], "recommendation": padded[7],
        })
        self._columns[column.key] = column
        if column.key not in self._visible:
            self._visible.append(column.key)
        self._draw()

    def bind(self, sequence: str | None = None, func: object = None, add: object = None) -> str:
        # Preserve ordinary widget bindings; virtual Treeview selection is supplied by callbacks.
        return super().bind(sequence, func, add)

    def open_field_editor(self, key: str, options: list[str], current: str, commit: Callable[[str], None]) -> Toplevel:
        popup = Toplevel(self)
        popup.title(f"修改 {self._columns[key].excel_label} 的识别字段")
        popup.transient(self.winfo_toplevel())
        popup.resizable(False, False)
        box = ttk.Combobox(popup, values=options, state="readonly", width=34)
        box.set(current if current in options else options[0])
        box.pack(padx=10, pady=10)
        ttk.Button(popup, text="确认", command=lambda: (commit(box.get()), popup.destroy())).pack(pady=(0, 10))
        box.focus_set()
        return popup

    def _status_colors(self, status: str) -> tuple[str, str]:
        if "错误" in status or "不能" in status:
            return "#B42318", "#FDECEA"
        if "确认" in status:
            return "#946200", "#FFF4D6"
        if "忽略" in status:
            return "#607487", "#EEF1F3"
        return "#1F7A4D", "#E8F6EE"

    def _draw(self) -> None:
        self.property_canvas.delete("all")
        self.data_canvas.delete("all")
        total_height = self.header_height + len(PROPERTY_ROWS) * self.row_height
        total_width = max(1, len(self._visible)) * self.column_width
        self.property_canvas.configure(scrollregion=(0, 0, self.property_width, total_height))
        self.data_canvas.configure(scrollregion=(0, 0, total_width, total_height))
        self.property_canvas.create_rectangle(0, 0, self.property_width, self.header_height, fill="#DCE7EF", outline="#B8CAD6")
        self.property_canvas.create_text(12, self.header_height / 2, text="属性", anchor="w", font=("Microsoft YaHei UI", 9, "bold"), fill="#18324A")
        for row_index, (_key, title) in enumerate(PROPERTY_ROWS):
            y1 = self.header_height + row_index * self.row_height
            self.property_canvas.create_rectangle(0, y1, self.property_width, y1 + self.row_height, fill="#F5F8FA", outline="#D7E1E8")
            self.property_canvas.create_text(10, y1 + self.row_height / 2, text=title, anchor="w", width=self.property_width - 18, font=("Microsoft YaHei UI", 9, "bold"), fill="#1F3446")
        for col_index, key in enumerate(self._visible):
            column = self._columns[key]
            x1 = col_index * self.column_width
            selected = key in self._selected
            head_fill = "#CFE3F2" if selected else "#DCE7EF"
            self.data_canvas.create_rectangle(x1, 0, x1 + self.column_width, self.header_height, fill=head_fill, outline="#8BAFC6", width=2 if selected else 1)
            self.data_canvas.create_text(x1 + self.column_width / 2, self.header_height / 2, text=column.excel_label, font=("Microsoft YaHei UI", 9, "bold"), fill="#18324A")
            status_fg, status_bg = self._status_colors(column.values.get("status", ""))
            for row_index, (row_key, _title) in enumerate(PROPERTY_ROWS):
                y1 = self.header_height + row_index * self.row_height
                fill = status_bg if row_key == "status" else ("#F8FBFD" if selected else "#FFFFFF")
                self.data_canvas.create_rectangle(x1, y1, x1 + self.column_width, y1 + self.row_height, fill=fill, outline="#D7E1E8", width=2 if selected else 1)
                self.data_canvas.create_text(
                    x1 + 7,
                    y1 + self.row_height / 2,
                    text=column.values.get(row_key, "") or "—",
                    anchor="w",
                    width=self.column_width - 14,
                    font=("Microsoft YaHei UI", 9, "bold" if row_key in {"field", "status"} else "normal"),
                    fill=status_fg if row_key == "status" else "#1F3446",
                )

    def _locate(self, event: object) -> tuple[str | None, str | None]:
        x = self.data_canvas.canvasx(event.x)
        y = self.data_canvas.canvasy(event.y)
        col = int(x // self.column_width)
        if col < 0 or col >= len(self._visible):
            return None, None
        if y < self.header_height:
            return self._visible[col], "header"
        row = int((y - self.header_height) // self.row_height)
        if row < 0 or row >= len(self._row_keys):
            return self._visible[col], None
        return self._visible[col], self._row_keys[row]

    def _click(self, event: object) -> None:
        key, _row = self._locate(event)
        if not key:
            return
        if getattr(event, "state", 0) & 0x0004:
            if key in self._selected:
                self._selected.remove(key)
            else:
                self._selected.append(key)
        else:
            self._selected = [key]
        self._focus = key
        self._draw()
        if self.on_select:
            self.on_select()

    def _double_click(self, event: object) -> None:
        key, row = self._locate(event)
        if key and row == "field" and self.on_edit:
            self.on_edit(key)

    def _mousewheel(self, event: object) -> None:
        self._yview("scroll", int(-getattr(event, "delta", 0) / 120), "units")

    def _yview(self, *args: object) -> None:
        self.property_canvas.yview(*args)
        self.data_canvas.yview(*args)

    def _yscroll_set(self, first: str, last: str) -> None:
        self.vbar.set(first, last)
        if self.property_canvas.yview() != self.data_canvas.yview():
            self.property_canvas.yview_moveto(first)

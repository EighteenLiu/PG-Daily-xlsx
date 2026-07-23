from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from services.excel_processor import ExcelProcessor, LedgerData, LedgerProcessingError
from services.accuracy_service import AccuracyService
from services.notice_service import NoticeService
from services.resident_service import ResidentService
from utils.logger import setup_logging
from utils.paths import app_root, ensure_app_dirs


class LedgerApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ensure_app_dirs()
        self.logger = setup_logging()
        self.processor = ExcelProcessor()
        self.resident_service = ResidentService()
        self.accuracy_service = AccuracyService()
        self.notice_service = NoticeService()
        self.original: LedgerData | None = None
        self.base: LedgerData | None = None
        self.problem: LedgerData | None = None
        self.resident: LedgerData | None = None
        self.split_ledgers: dict[str, LedgerData] = {}
        self.selected_split_name: str | None = None
        self.resident_generated = False
        self.accuracy_updated = False
        self.current_stage = "原始台账"
        self.buttons: dict[str, ctk.CTkButton] = {}

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.title("台账处理与拆分系统")
        self.geometry("1260x780")
        self.minsize(1060, 660)
        self._build_ui()

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        top = ctk.CTkFrame(self, corner_radius=0)
        top.grid(row=0, column=0, sticky="ew")
        top.grid_columnconfigure(1, weight=1)

        title = ctk.CTkLabel(top, text="台账处理与拆分系统", font=ctk.CTkFont(size=22, weight="bold"))
        title.grid(row=0, column=0, padx=18, pady=(14, 4), sticky="w")
        self.status_label = ctk.CTkLabel(top, text="请选择原始台账", anchor="e")
        self.status_label.grid(row=0, column=1, padx=18, pady=(14, 4), sticky="e")

        self.step_label = ctk.CTkLabel(
            top,
            text="流程：上传原始台账 -> 生成基础台账 -> 生成问题台账 -> 生成专项表 / 更新统计表 / 拆分导出",
        )
        self.step_label.grid(row=1, column=0, columnspan=2, padx=18, pady=(0, 14), sticky="w")

        body = ctk.CTkFrame(self, corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew")
        body.grid_columnconfigure(0, weight=0)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkScrollableFrame(body, width=300, corner_radius=8)
        sidebar.grid(row=0, column=0, sticky="nsw", padx=(16, 8), pady=16)
        sidebar.grid_columnconfigure(0, weight=1)

        row = 0
        row = self._add_section(
            sidebar,
            row,
            "文件与预览",
            [
                ("上传原始台账", self.upload_original),
                ("预览当前", self.preview_current),
                ("重新上传当前", self.reupload_current),
            ],
        )
        row = self._add_section(
            sidebar,
            row,
            "基础与问题台账",
            [
                ("生成基础台账", self.generate_base),
                ("导入基础台账", self.upload_base),
                ("生成问题台账", self.generate_problem),
                ("导入问题台账", self.upload_problem),
            ],
        )
        row = self._add_section(
            sidebar,
            row,
            "专项表与统计",
            [
                ("生成居民自主投放表", self.generate_resident),
                ("预览居民自主投放表", self.preview_resident),
                ("更新准确率统计表", self.update_accuracy),
            ],
        )
        row = self._add_section(
            sidebar,
            row,
            "台账拆分",
            [
                ("生成拆分预览", self.generate_split),
                ("导出拆分Excel", self.export_all),
            ],
        )

        self.info_box = ctk.CTkTextbox(sidebar, height=120, wrap="word")
        self.info_box.grid(row=row, column=0, sticky="ew", padx=12, pady=(12, 16))
        self.info_box.insert("1.0", "右侧导出当前台账会导出当前预览内容\n导出拆分Excel会批量导出全部拆分表")
        self.info_box.configure(state="disabled")

        table_frame = ctk.CTkFrame(body)
        table_frame.grid(row=0, column=1, sticky="nsew", padx=(8, 16), pady=16)
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(1, weight=1)

        self.table_title = ctk.CTkLabel(table_frame, text="预览", font=ctk.CTkFont(size=16, weight="bold"))
        self.table_title.grid(row=0, column=0, sticky="w", padx=12, pady=12)

        self.export_current_button = ctk.CTkButton(
            table_frame,
            text="导出当前台账",
            command=self.export_current,
            width=120,
            height=32,
        )
        self.export_current_button.grid(row=0, column=0, sticky="e", padx=(12, 410), pady=12)
        self.buttons["导出当前台账"] = self.export_current_button

        self.notice_button = ctk.CTkButton(
            table_frame,
            text="生成通告",
            command=self.generate_notice,
            width=110,
            height=32,
        )
        self.notice_button.grid(row=0, column=0, sticky="e", padx=(12, 292), pady=12)
        self.buttons["生成通告"] = self.notice_button

        self.split_selector = ctk.CTkOptionMenu(
            table_frame,
            values=["暂无拆分表"],
            command=self._select_split_preview,
            state="disabled",
            width=260,
        )
        self.split_selector.grid(row=0, column=0, sticky="e", padx=12, pady=12)

        self.tree = ttk.Treeview(table_frame, show="headings")
        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(12, 0), pady=(0, 12))
        y_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 12))
        x_scroll.grid(row=2, column=0, sticky="ew", padx=(12, 0), pady=(0, 12))
        self._update_button_states()

    def _add_section(
        self,
        parent: ctk.CTkBaseClass,
        row: int,
        title: str,
        actions: list[tuple[str, object]],
    ) -> int:
        section = ctk.CTkFrame(parent, corner_radius=8)
        section.grid(row=row, column=0, sticky="ew", padx=8, pady=(8 if row == 0 else 10, 0))
        section.grid_columnconfigure(0, weight=1)
        label = ctk.CTkLabel(section, text=title, font=ctk.CTkFont(size=14, weight="bold"), anchor="w")
        label.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 4))
        for idx, (text, command) in enumerate(actions, start=1):
            button = ctk.CTkButton(section, text=text, command=command, height=34)
            button.grid(row=idx, column=0, sticky="ew", padx=12, pady=(6, 10 if idx == len(actions) else 0))
            self.buttons[text] = button
        return row + 1

    def upload_original(self) -> None:
        path = self._ask_excel_file()
        if not path:
            return
        self._run("上传原始台账", lambda: self._set_original(self.processor.load_ledger(path)))

    def preview_current(self) -> None:
        data = self._current_data()
        if not data:
            self._warn("当前没有可预览的台账")
            return
        self._display_ledger(data, self.current_stage)

    def reupload_current(self) -> None:
        if self.current_stage == "基础台账":
            self.upload_base()
        elif self.current_stage == "问题台账":
            self.upload_problem()
        else:
            self.upload_original()

    def next_step(self) -> None:
        if self.current_stage == "原始台账":
            self.generate_base()
        elif self.current_stage == "基础台账":
            self.generate_problem()
        elif self.current_stage == "问题台账":
            self.generate_split()
        elif self.current_stage == "拆分预览":
            self.export_all()
        else:
            self._warn("请先上传原始台账")

    def generate_base(self) -> None:
        if not self.original:
            self._warn("请先上传原始台账")
            return
        self._run("生成基础台账", lambda: self._set_base(self.processor.make_base_ledger(self.original)))

    def upload_base(self) -> None:
        path = self._ask_excel_file()
        if not path:
            return
        self._run("导入基础台账", lambda: self._set_base(self.processor.load_ledger(path)))

    def generate_problem(self) -> None:
        if not self.base:
            self._warn("请先生成或导入基础台账")
            return
        self._run("生成问题台账", lambda: self._set_problem(self.processor.make_problem_ledger(self.base)))

    def upload_problem(self) -> None:
        path = self._ask_excel_file()
        if not path:
            return
        self._run("导入问题台账", lambda: self._set_problem(self.processor.load_ledger(path)))

    def generate_resident(self) -> None:
        if not self.problem:
            self._warn("请先生成或导入问题台账")
            return

        def action() -> Path:
            self.resident = self.resident_service.make_resident_ledger(self.problem)
            out_dir = self._date_output_dir(self.problem)
            out_path = out_dir / f"{self.problem.date_label or '未识别日期'}居民自主投放.xlsx"
            saved = self.processor.save_ledger(self.resident, out_path)
            self.resident_generated = True
            self.current_stage = "居民自主投放表"
            self._display_ledger(self.resident, "居民自主投放表")
            self._set_status(f"居民自主投放表：{len(self.resident.rows)} 条记录，已生成 {saved}")
            self._update_button_states()
            return saved

        self._run("生成居民自主投放表", action)

    def preview_resident(self) -> None:
        if not self.resident:
            self._warn("请先生成居民自主投放表")
            return
        self.current_stage = "居民自主投放表"
        self._display_ledger(self.resident, "居民自主投放表")
        self._update_button_states()

    def update_accuracy(self) -> None:
        if not self.problem:
            self._warn("请先生成或导入问题台账")
            return
        path = filedialog.askopenfilename(
            title="选择小区村值守率及投放准确率统计表",
            initialdir=str(app_root() / "output"),
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return

        def action() -> Path:
            saved = self.accuracy_service.update_statistics(self.problem, path)
            self.accuracy_updated = True
            self._set_status(f"准确率统计表已更新：{saved}")
            self._update_button_states()
            return saved

        self._run("更新准确率统计表", action)

    def generate_notice(self) -> None:
        if not self.problem:
            self._warn("请先生成或导入问题台账")
            return
        path = filedialog.asksaveasfilename(
            title="生成通告",
            defaultextension=".txt",
            initialdir=str(app_root() / "output"),
            initialfile="通告.txt",
            filetypes=[("文本文件", "*.txt")],
        )
        if path:
            self._run("生成通告", lambda: self.notice_service.generate_notice(self.problem, path))

    def generate_split(self) -> None:
        if not self.problem:
            self._warn("请先生成或导入问题台账")
            return

        def action() -> None:
            self.split_ledgers = self.processor.split_by_location(self.problem)
            self.selected_split_name = next(iter(self.split_ledgers), None)
            self._refresh_split_selector()
            if self.selected_split_name:
                self._display_split_ledger(self.selected_split_name)
            else:
                self._show_table(["导出文件名", "记录数"], [], "拆分预览")
            self.current_stage = "拆分预览"
            self._set_status(f"已生成 {len(self.split_ledgers)} 个拆分文件预览")
            self._update_button_states()

        self._run("拆分预览", action)

    def export_current(self) -> None:
        data = self._current_data()
        if not data:
            self._warn("当前没有可导出的台账")
            return
        default = self._default_export_name(data, self.current_stage)
        path = filedialog.asksaveasfilename(
            title="导出当前台账",
            defaultextension=".xlsx",
            initialdir=str(app_root() / "output"),
            initialfile=default,
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if path:
            self._run("导出当前台账", lambda: self.processor.save_ledger(data, path))

    def export_all(self) -> None:
        if not self.split_ledgers:
            self.generate_split()
            if not self.split_ledgers:
                return
        directory = filedialog.askdirectory(
            title="选择拆分文件导出目录",
            initialdir=str(app_root() / "output"),
        )
        if directory:
            self._run(
                "导出拆分Excel",
                lambda: self._export_all_to_dir(directory),
            )

    def _export_all_to_dir(self, directory: str) -> None:
        saved = self.processor.save_split_ledgers(self.split_ledgers, directory)
        self._set_status(f"已导出 {len(saved)} 个文件：{directory}")
        messagebox.showinfo("导出完成", f"已导出 {len(saved)} 个文件")

    def _set_original(self, data: LedgerData) -> None:
        self.original = data
        self.base = None
        self.problem = None
        self.resident = None
        self.split_ledgers = {}
        self.selected_split_name = None
        self.resident_generated = False
        self.accuracy_updated = False
        self.current_stage = "原始台账"
        self._display_ledger(data, "原始台账")
        self._refresh_split_selector()
        self._update_button_states()

    def _set_base(self, data: LedgerData) -> None:
        self.base = data
        self.problem = None
        self.resident = None
        self.split_ledgers = {}
        self.selected_split_name = None
        self.resident_generated = False
        self.accuracy_updated = False
        self.current_stage = "基础台账"
        self._display_ledger(data, "基础台账")
        self._refresh_split_selector()
        self._update_button_states()

    def _set_problem(self, data: LedgerData) -> None:
        self.problem = data
        self.resident = None
        self.split_ledgers = {}
        self.selected_split_name = None
        self.resident_generated = False
        self.accuracy_updated = False
        self.current_stage = "问题台账"
        self._display_ledger(data, "问题台账")
        self._refresh_split_selector()
        self._update_button_states()

    def _display_ledger(self, data: LedgerData, title: str) -> None:
        self._show_table(data.headers, data.preview_rows(), title)
        self._set_status(f"{title}：{len(data.rows)} 条记录，{len(data.headers)} 列")

    def _show_table(self, headers: list[str], rows: list[list[object]], title: str) -> None:
        self.table_title.configure(text=title)
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = headers
        for header in headers:
            self.tree.heading(header, text=header)
            self.tree.column(header, width=max(110, min(len(header) * 18 + 40, 240)), stretch=True)
        for row in rows[:300]:
            self.tree.insert("", "end", values=[("" if value is None else value) for value in row])

    def _current_data(self) -> LedgerData | None:
        if self.current_stage == "居民自主投放表":
            return self.resident
        if self.current_stage == "拆分预览" and self.selected_split_name:
            return self.split_ledgers.get(self.selected_split_name)
        if self.current_stage == "问题台账":
            return self.problem
        if self.current_stage == "基础台账":
            return self.base
        if self.current_stage == "原始台账":
            return self.original
        return None

    def _ask_excel_file(self) -> str:
        return filedialog.askopenfilename(
            title="选择Excel台账",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
        )

    def _run(self, label: str, func) -> None:
        try:
            result = func()
            self.logger.info("%s完成", label)
            if isinstance(result, Path):
                self._set_status(f"{label}完成：{result}")
                messagebox.showinfo("完成", f"{label}完成")
        except LedgerProcessingError as exc:
            self.logger.exception("%s失败", label)
            messagebox.showerror("处理失败", str(exc))
        except Exception as exc:
            self.logger.exception("%s失败", label)
            messagebox.showerror("处理失败", f"{label}失败：{exc}")

    def _set_status(self, text: str) -> None:
        self.status_label.configure(text=text)

    def _warn(self, text: str) -> None:
        messagebox.showwarning("提示", text)

    def _date_output_dir(self, data: LedgerData) -> Path:
        name = data.date_label or "未识别日期"
        directory = app_root() / "output" / name
        directory.mkdir(parents=True, exist_ok=True)
        return directory

    def _update_button_states(self) -> None:
        rules = {
            "预览当前": self._current_data() is not None,
            "重新上传当前": True,
            "下一步": self._next_step_enabled(),
            "生成基础台账": self.original is not None,
            "导入基础台账": True,
            "导出当前台账": self._current_data() is not None,
            "生成问题台账": self.base is not None,
            "导入问题台账": True,
            "生成居民自主投放表": self.problem is not None,
            "预览居民自主投放表": self.resident is not None,
            "更新准确率统计表": self.problem is not None,
            "生成通告": self.problem is not None,
            "生成拆分预览": self.problem is not None,
            "导出拆分Excel": bool(self.split_ledgers),
        }
        for text, enabled in rules.items():
            button = self.buttons.get(text)
            if button:
                button.configure(state=("normal" if enabled else "disabled"))

    def _next_step_enabled(self) -> bool:
        if self.current_stage == "原始台账":
            return self.original is not None
        if self.current_stage == "基础台账":
            return self.base is not None
        if self.current_stage == "问题台账":
            return self.problem is not None
        if self.current_stage == "居民自主投放表":
            return self.resident is not None
        if self.current_stage == "拆分预览":
            return bool(self.split_ledgers)
        return False

    def _default_export_name(self, data: LedgerData, stage: str) -> str:
        date_label = data.date_label or "未识别日期"
        if stage == "拆分预览" and self.selected_split_name:
            return self.selected_split_name
        if stage == "居民自主投放表":
            return f"{date_label}居民自主投放.xlsx"
        return f"{date_label}{stage}.xlsx"

    def _refresh_split_selector(self) -> None:
        names = list(self.split_ledgers)
        if names:
            self.split_selector.configure(values=names, state="normal")
            selected = self.selected_split_name if self.selected_split_name in self.split_ledgers else names[0]
            self.selected_split_name = selected
            self.split_selector.set(selected)
        else:
            self.split_selector.configure(values=["暂无拆分表"], state="disabled")
            self.split_selector.set("暂无拆分表")

    def _select_split_preview(self, name: str) -> None:
        if name not in self.split_ledgers:
            return
        self.selected_split_name = name
        self.current_stage = "拆分预览"
        self._display_split_ledger(name)
        self._update_button_states()

    def _display_split_ledger(self, name: str) -> None:
        ledger = self.split_ledgers.get(name)
        if not ledger:
            return
        self._show_table(ledger.headers, ledger.preview_rows(), f"拆分预览：{name}")
        self._set_status(f"{name}：{len(ledger.rows)} 条记录，{len(ledger.headers)} 列")

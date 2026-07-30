from __future__ import annotations

import argparse
import json
import os
import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import invoice_workflow as workflow
from invoice_workflow import AdjustmentRow, OcrConfirmRow, WorkflowState


APP_TITLE = "大統進貨助手"
SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}


class InvoiceOcrApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.state = WorkflowState()
        self.image_path = tk.StringVar(value="")
        self.selected_image_paths: list[Path] = []
        self.selected_image_display = ""
        self.output_dir = tk.StringVar(value=str(workflow.PROJECT_ROOT / "建檔進貨用" / "進貨圖片轉試算表"))
        self.product_csv_path = tk.StringVar(value=str(workflow.product_csv() or ""))
        self.vendor_name = tk.StringVar(value="尚未辨識")
        self.status = tk.StringVar(value="請選擇一張或多張同廠商進貨圖片。")
        self.worker: threading.Thread | None = None
        self.started_at = 0.0
        self.current_output: Path | None = None
        self.edit_entry: ttk.Entry | None = None
        self.table_mode = "ocr"
        self.workflow_stage = "initial"
        self.product_match_ready = False
        self.ocr_candidate_values: dict[str, str] = {}
        self.adjustment_source_rows: dict[str, int] = {}
        self.adjustment_draft_rows: list[AdjustmentRow] | None = None
        self.match_changed_since_back = False
        self.active_cell_item = ""
        self.active_cell_column = ""
        self.new_row_counter = 0

        self.title(APP_TITLE)
        self.geometry("1280x800")
        self.minsize(1080, 700)
        self.configure(padx=14, pady=12)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.create_widgets()
        self.check_runtime_on_start()

    def create_widgets(self) -> None:
        menu_bar = tk.Menu(self)
        category_menu = tk.Menu(menu_bar, tearoff=False)
        category_menu.add_command(label="管理大類…", command=self.open_category_manager)
        menu_bar.add_cascade(label="大類", menu=category_menu)
        self.configure(menu=menu_bar)

        header = ttk.Label(self, text=APP_TITLE, font=("Microsoft JhengHei UI", 18, "bold"))
        header.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        ttk.Label(self, text="進貨圖片").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.image_path).grid(row=1, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
        ttk.Button(self, text="選擇圖片（可多選）", command=self.pick_image).grid(row=1, column=5, sticky="ew", pady=6)

        ttk.Label(self, text="輸出資料夾").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.output_dir).grid(row=2, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
        ttk.Button(self, text="選擇資料夾", command=self.pick_output_dir).grid(row=2, column=5, sticky="ew", pady=6)

        ttk.Label(self, text="產品資料輸出").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.product_csv_path).grid(
            row=3, column=1, columnspan=4, sticky="ew", padx=8, pady=6
        )
        ttk.Button(self, text="選擇產品 CSV", command=self.pick_product_csv).grid(
            row=3, column=5, sticky="ew", pady=6
        )

        ttk.Label(self, text="偵測廠商").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.vendor_name, state="readonly").grid(
            row=4, column=1, columnspan=5, sticky="ew", padx=8, pady=6
        )

        self.step_labels: dict[str, ttk.Label] = {}
        steps = [
            ("ocr", "1 OCR 原文確認"),
            ("match", "2 已建檔勾選"),
            ("review", "3 名稱調整"),
            ("build", "4 正式輸出"),
            ("cleanup", "5 清理中間檔"),
        ]
        step_frame = ttk.Frame(self)
        step_frame.grid(row=5, column=0, columnspan=6, sticky="ew", pady=(8, 6))
        for idx, (key, text) in enumerate(steps):
            label = ttk.Label(step_frame, text=f"□ {text}", font=("Microsoft JhengHei UI", 10, "bold"))
            label.grid(row=0, column=idx, sticky="w", padx=(0, 18))
            self.step_labels[key] = label

        button_frame = ttk.Frame(self)
        button_frame.grid(row=6, column=0, columnspan=6, sticky="ew", pady=(4, 8))
        self.start_button = ttk.Button(button_frame, text="開始 OCR", command=self.start_ocr)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.match_button = ttk.Button(
            button_frame,
            text="開始產品重複比對",
            command=self.start_product_match,
            state="disabled",
        )
        self.match_button.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.review_button = ttk.Button(button_frame, text="確認文字並開始調整", command=self.prepare_review, state="disabled")
        self.review_button.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.candidate_button = ttk.Button(
            button_frame,
            text="選擇候選",
            command=self.open_candidate_picker,
            state="disabled",
        )
        self.candidate_button.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        self.add_item_button = ttk.Button(button_frame, text="拆分/新增品項", command=self.split_selected_item, state="disabled")
        self.add_item_button.grid(row=0, column=4, sticky="ew", padx=(0, 8))
        self.adjust_button = ttk.Button(button_frame, text="確認資料", command=self.generate_adjusted, state="disabled")
        self.adjust_button.grid(row=0, column=5, sticky="ew", padx=(0, 8))
        self.build_button = ttk.Button(button_frame, text="產生建檔/採購檔", command=self.build_outputs, state="disabled")
        self.build_button.grid(row=0, column=6, sticky="ew", padx=(0, 8))
        self.cleanup_button = ttk.Button(button_frame, text="清理中間檔", command=self.cleanup_files, state="disabled")
        self.cleanup_button.grid(row=0, column=7, sticky="ew", padx=(0, 8))
        self.open_button = ttk.Button(button_frame, text="開啟目前檔案", command=self.open_current_output, state="disabled")
        self.open_button.grid(row=0, column=8, sticky="ew")
        self.back_button = ttk.Button(
            button_frame,
            text="回到上一步",
            command=self.go_back,
            state="disabled",
        )
        self.back_button.grid(row=1, column=0, sticky="ew", padx=(0, 8), pady=(6, 0))
        self.copy_column_button = ttk.Button(
            button_frame,
            text="複製整欄",
            command=self.copy_selected_column,
            state="disabled",
        )
        self.copy_column_button.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(6, 0))
        self.delete_item_button = ttk.Button(
            button_frame,
            text="刪除選取項目",
            command=self.delete_selected_items,
            state="disabled",
        )
        self.delete_item_button.grid(row=1, column=2, sticky="ew", padx=(0, 8), pady=(6, 0))
        self.deep_candidate_button = ttk.Button(
            button_frame,
            text="更低候選單筆檢查",
            command=self.start_lower_candidate_search,
            state="disabled",
        )
        self.deep_candidate_button.grid(row=1, column=3, sticky="ew", padx=(0, 8), pady=(6, 0))
        ttk.Label(
            button_frame,
            text="先點選任一格或欄位標題，再按「複製整欄」（快捷鍵 Ctrl+Shift+C）",
        ).grid(row=1, column=4, columnspan=5, sticky="w", pady=(6, 0))
        for col in range(9):
            button_frame.columnconfigure(col, weight=1)

        table_frame = ttk.LabelFrame(self, text="資料確認")
        table_frame.grid(row=7, column=0, columnspan=6, sticky="nsew", pady=(4, 8))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)
        columns = ("已建檔", "OCR品名", "數量", "進價", "金額", "已建檔代號", "已建檔品名", "比對狀態", "相似候選")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=13)
        self.configure_tree_columns(columns, [70, 300, 70, 80, 80, 100, 260, 95, 430], {"OCR品名", "已建檔品名", "相似候選"})
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<ButtonRelease-1>", self.select_adjustment_cell)
        self.tree.bind("<Double-1>", self.begin_edit_cell)
        self.tree.bind("<Button-3>", self.show_column_copy_menu)
        self.tree.bind("<Control-Shift-C>", self.copy_selected_column)
        self.tree.bind("<Control-Shift-c>", self.copy_selected_column)
        self.tree.bind("<Delete>", self.delete_selected_items)
        self.column_copy_menu = tk.Menu(self, tearoff=False)
        self.column_copy_menu.add_command(label="複製此整欄", command=self.copy_selected_column)

        status_frame = ttk.LabelFrame(self, text="狀態")
        status_frame.grid(row=8, column=0, columnspan=6, sticky="nsew")
        status_frame.rowconfigure(1, weight=1)
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status, font=("Microsoft JhengHei UI", 11, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )
        self.detail_text = tk.Text(status_frame, height=8, wrap="word", state="disabled")
        self.detail_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.columnconfigure(1, weight=1)
        self.rowconfigure(7, weight=2)
        self.rowconfigure(8, weight=1)

    def configure_tree_columns(self, columns: tuple[str, ...], widths: list[int], stretch_columns: set[str]) -> None:
        self.tree.configure(columns=columns)
        for column, width in zip(columns, widths):
            self.tree.heading(
                column,
                text=column,
                command=lambda selected_column=column: self.select_column_for_copy(selected_column),
            )
            self.tree.column(column, width=width, stretch=(column in stretch_columns))

    def open_category_manager(self) -> None:
        dialog = tk.Toplevel(self)
        dialog.title("大類管理")
        dialog.geometry("620x540")
        dialog.minsize(520, 420)
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(
            dialog,
            text="APP 內建大類｜可新增或刪除；修改後會保存到 APP 設定。",
            font=("Microsoft JhengHei UI", 11, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 8))

        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=12)
        category_tree = ttk.Treeview(
            list_frame,
            columns=("大類代號", "大類名稱"),
            show="headings",
            selectmode="extended",
        )
        category_tree.heading("大類代號", text="大類代號")
        category_tree.heading("大類名稱", text="大類名稱")
        category_tree.column("大類代號", width=100, stretch=False, anchor="center")
        category_tree.column("大類名稱", width=380, stretch=True)
        category_scroll = ttk.Scrollbar(
            list_frame,
            orient="vertical",
            command=category_tree.yview,
        )
        category_tree.configure(yscrollcommand=category_scroll.set)
        category_tree.pack(side="left", fill="both", expand=True)
        category_scroll.pack(side="right", fill="y")

        form_frame = ttk.LabelFrame(dialog, text="新增大類")
        form_frame.pack(fill="x", padx=12, pady=(10, 8))
        code_value = tk.StringVar()
        name_value = tk.StringVar()
        ttk.Label(form_frame, text="代號").grid(row=0, column=0, padx=(8, 4), pady=8)
        code_entry = ttk.Entry(form_frame, textvariable=code_value, width=12)
        code_entry.grid(row=0, column=1, padx=(0, 12), pady=8)
        ttk.Label(form_frame, text="名稱").grid(row=0, column=2, padx=(0, 4), pady=8)
        name_entry = ttk.Entry(form_frame, textvariable=name_value)
        name_entry.grid(row=0, column=3, sticky="ew", padx=(0, 8), pady=8)
        form_frame.columnconfigure(3, weight=1)

        def reload_categories() -> None:
            for item in category_tree.get_children():
                category_tree.delete(item)
            for row in workflow.category_rule_rows():
                category_tree.insert(
                    "",
                    "end",
                    iid=f"category-{row['code']}",
                    values=(row["code"], row["name"]),
                )

        def refresh_current_adjustment_names() -> None:
            if self.table_mode != "adjust":
                return
            names = workflow.category_name_map()
            for item in self.tree.get_children():
                values = list(self.tree.item(item, "values"))
                code = workflow.category_code(values[2])
                if code in names:
                    values[3] = names[code]
                    self.tree.item(item, values=values)

        def add_category() -> None:
            code = code_value.get().strip()
            name = name_value.get().strip()
            if not re.fullmatch(r"\d{1,3}", code):
                messagebox.showwarning(
                    APP_TITLE,
                    "大類代號請輸入 1 至 3 位數字。",
                    parent=dialog,
                )
                return
            if not name:
                messagebox.showwarning(APP_TITLE, "請輸入大類名稱。", parent=dialog)
                return
            rows = workflow.category_rule_rows()
            if any(row["code"] == code for row in rows):
                messagebox.showwarning(
                    APP_TITLE,
                    f"大類代號 {code} 已存在，請先刪除舊項目或使用其他代號。",
                    parent=dialog,
                )
                return
            rows.append(
                {
                    "code": code,
                    "name": name,
                    "label": name,
                    "keywords": name,
                    "note": "使用者從 APP 新增",
                }
            )
            workflow.save_category_rule_rows(rows)
            code_value.set("")
            name_value.set("")
            reload_categories()
            refresh_current_adjustment_names()
            self.status.set(f"已新增大類：{code} {name}")
            code_entry.focus_set()

        def delete_categories() -> None:
            selected = list(category_tree.selection())
            if not selected:
                messagebox.showwarning(APP_TITLE, "請先選取要刪除的大類。", parent=dialog)
                return
            selected_codes = {
                str(category_tree.item(item, "values")[0]) for item in selected
            }
            selected_text = "、".join(
                f"{category_tree.item(item, 'values')[0]} {category_tree.item(item, 'values')[1]}"
                for item in selected[:8]
            )
            if not messagebox.askyesno(
                APP_TITLE,
                f"確定刪除 {len(selected)} 個大類？\n\n{selected_text}\n\n"
                "已使用被刪除大類的商品，之後必須重新選擇有效大類。",
                parent=dialog,
            ):
                return
            remaining = [
                row
                for row in workflow.category_rule_rows()
                if row["code"] not in selected_codes
            ]
            workflow.save_category_rule_rows(remaining)
            reload_categories()
            refresh_current_adjustment_names()
            self.status.set(f"已刪除 {len(selected_codes)} 個大類。")

        def reset_categories() -> None:
            if not messagebox.askyesno(
                APP_TITLE,
                "確定恢復 APP 內建預設大類？\n\n自行新增或刪除的變更會被重設。",
                parent=dialog,
            ):
                return
            workflow.reset_category_rule_rows()
            reload_categories()
            refresh_current_adjustment_names()
            self.status.set("已恢復 APP 內建預設大類。")

        ttk.Button(form_frame, text="新增", command=add_category).grid(
            row=0,
            column=4,
            padx=(0, 8),
            pady=8,
        )
        code_entry.bind("<Return>", lambda _event: name_entry.focus_set())
        name_entry.bind("<Return>", lambda _event: add_category())

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(button_frame, text="刪除選取", command=delete_categories).pack(
            side="left"
        )
        ttk.Button(button_frame, text="恢復內建預設", command=reset_categories).pack(
            side="left",
            padx=(8, 0),
        )
        ttk.Button(button_frame, text="關閉", command=dialog.destroy).pack(side="right")

        reload_categories()

    def set_ocr_table(self) -> None:
        self.table_mode = "ocr"
        columns = ("已建檔", "OCR品名", "數量", "進價", "金額", "已建檔代號", "已建檔品名", "比對狀態", "相似候選")
        self.configure_tree_columns(columns, [70, 300, 70, 80, 80, 100, 260, 95, 430], {"OCR品名", "已建檔品名", "相似候選"})
        self.add_item_button.configure(state="disabled")

    def set_adjustment_table(self) -> None:
        self.table_mode = "adjust"
        columns = ("產品代號", "產品名稱", "大類代號", "大類名稱", "數量", "成本", "金額", "狀態")
        self.configure_tree_columns(
            columns,
            [95, 360, 85, 180, 70, 80, 80, 110],
            {"產品名稱", "大類名稱"},
        )
        self.candidate_button.configure(state="disabled")
        self.deep_candidate_button.configure(state="disabled")

    def check_runtime_on_start(self) -> None:
        issues = workflow.validate_runtime()
        if issues:
            self.status.set("環境需要檢查。")
            self.set_detail("\n".join(issues))
            return
        csv_path = workflow.product_csv()
        self.set_detail(f"產品資料：{csv_path}\n請確認每天使用前已選擇今天匯出的產品資料輸出 CSV。")

    def pick_image(self) -> None:
        selected = filedialog.askopenfilenames(
            title="選擇一張或多張同廠商進貨圖片",
            filetypes=[("圖片檔", "*.jpg *.jpeg *.png *.webp"), ("所有檔案", "*.*")],
        )
        if selected:
            self.selected_image_paths = [Path(path) for path in selected]
            if len(self.selected_image_paths) == 1:
                display = str(self.selected_image_paths[0])
            else:
                display = (
                    f"已選擇 {len(self.selected_image_paths)} 張｜"
                    + "｜".join(path.name for path in self.selected_image_paths)
                )
            self.selected_image_display = display
            self.image_path.set(display)
            self.status.set(f"已選擇 {len(self.selected_image_paths)} 張圖片，可以開始 OCR。")

    def pick_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="選擇輸出資料夾", initialdir=self.output_dir.get())
        if selected:
            self.output_dir.set(selected)

    def pick_product_csv(self) -> None:
        initial_path = Path(self.product_csv_path.get().strip('" ') or workflow.REFERENCE_DIR)
        initial_dir = initial_path.parent if initial_path.suffix else initial_path
        selected = filedialog.askopenfilename(
            title="選擇產品資料輸出 CSV",
            initialdir=str(initial_dir if initial_dir.exists() else workflow.REFERENCE_DIR),
            filetypes=[("CSV 檔案", "*.csv *.CSV"), ("所有檔案", "*.*")],
        )
        if not selected:
            return
        try:
            csv_path = workflow.prepare_product_csv_for_use(
                Path(selected),
                persist=True,
            )
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.product_csv_path.set(str(csv_path))
        self.status.set("產品資料輸出路徑已更新。")
        self.set_detail(f"目前產品資料：{csv_path}")

    def validate_start_inputs(self) -> tuple[list[Path], Path] | None:
        image_text = self.image_path.get().strip('" ')
        output_text = self.output_dir.get().strip('" ')
        if not image_text:
            messagebox.showwarning(APP_TITLE, "請先選擇進貨單圖片。")
            return None
        images = (
            self.selected_image_paths
            if self.selected_image_paths and image_text == self.selected_image_display
            else [Path(image_text)]
        )
        for image in images:
            if not image.exists():
                messagebox.showwarning(APP_TITLE, f"找不到圖片：{image}")
                return None
            if image.suffix.lower() not in SUPPORTED_EXTENSIONS:
                messagebox.showwarning(APP_TITLE, "目前只支援 jpg、jpeg、png、webp 圖片。")
                return None
        if not output_text:
            messagebox.showwarning(APP_TITLE, "請指定輸出資料夾。")
            return None
        issues = workflow.validate_runtime()
        if issues:
            messagebox.showerror(APP_TITLE, "\n".join(issues))
            return None
        return images, Path(output_text)

    def start_ocr(self) -> None:
        selected = self.validate_start_inputs()
        if selected is None:
            return
        images, output_dir = selected
        output_dir.mkdir(parents=True, exist_ok=True)
        self.state = WorkflowState(
            image_path=images[0],
            image_paths=images,
            output_dir=output_dir,
        )
        self.workflow_stage = "initial"
        self.adjustment_draft_rows = None
        self.match_changed_since_back = False
        self.active_cell_item = ""
        self.active_cell_column = ""
        self.vendor_name.set("辨識中…")
        self.product_match_ready = False
        self.clear_tree()
        self.set_ocr_table()
        self.reset_steps()
        self.disable_buttons()
        self.start_button.configure(state="disabled")
        self.started_at = time.monotonic()
        self.status.set(f"OCR 辨識中，共 {len(images)} 張圖片。")
        self.set_detail("目前只進行 OCR；辨識完成並確認產品資料最新後，才會開始重複產品比對。")
        self.run_in_thread(self.worker_ocr)

    def update_batch_progress(self, current: int, total: int, image: Path) -> None:
        self.status.set(f"OCR 辨識中：第 {current}/{total} 張｜{image.name}")

    def worker_ocr(self) -> None:
        workflow.run_ocr_batch(
            self.state,
            progress_callback=lambda current, total, image: self.after(
                0,
                self.update_batch_progress,
                current,
                total,
                image,
            ),
        )
        self.after(0, self.vendor_name.set, self.state.vendor or "未辨識廠商")
        self.after(0, self.ocr_finished)

    def ocr_finished(self) -> None:
        self.worker = None
        self.workflow_stage = "ocr"
        self.product_match_ready = False
        self.vendor_name.set(self.state.vendor or "未辨識廠商")
        self.mark_step("ocr", True)
        self.mark_step("match", False)
        self.populate_ocr_tree(workflow.load_ocr_confirm_rows(self.state))
        self.current_output = self.state.raw_xlsx
        self.status.set("OCR 已完成，請先確認原始文字與產品資料版本。")
        lines = [
            "可雙擊表格修改 OCR 品名、數量、進價、金額。",
            "確認 OCR 內容後，請先確認「產品資料輸出」是今天最新版本。",
            "接著按「開始產品重複比對」，才會產生已建檔與相似候選。",
            "",
            f"廠商：{self.state.vendor}",
            f"圖片張數：{len(workflow.state_image_paths(self.state))}",
            f"商品筆數：{self.state.row_count}",
            f"合併總額：{self.state.invoice_total}",
        ]
        if self.state.ocr_issues:
            lines.append("")
            lines.append("OCR 疑點：")
            lines.extend(f"- {issue}" for issue in self.state.ocr_issues[:12])
        self.set_detail("\n".join(lines))
        self.start_button.configure(state="normal")
        self.match_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.refresh_navigation_buttons()
        issue_text = (
            f"\n另外偵測到 {len(self.state.ocr_issues)} 項 OCR 疑點，請特別檢查。"
            if self.state.ocr_issues
            else ""
        )
        messagebox.showinfo(
            "OCR 完成",
            f"OCR 已完成。\n\n"
            f"偵測廠商：{self.state.vendor or '未辨識廠商'}\n"
            f"圖片張數：{len(workflow.state_image_paths(self.state))}\n"
            f"商品筆數：{self.state.row_count}"
            f"{issue_text}\n\n"
            "請確認產品輸出資料為最新後再繼續。\n"
            "確認後請按「開始產品重複比對」。",
            parent=self,
        )

    def start_product_match(self) -> None:
        if self.state.raw_xlsx is None or not self.state.raw_xlsx.exists():
            messagebox.showwarning(APP_TITLE, "請先完成 OCR。")
            return
        product_csv_text = self.product_csv_path.get().strip('" ')
        if not product_csv_text:
            messagebox.showerror(APP_TITLE, "請選擇最新的產品資料輸出 CSV。")
            return
        try:
            csv_path = workflow.prepare_product_csv_for_use(
                Path(product_csv_text),
                persist=True,
            )
            workflow.save_raw_ocr_rows(self.state, self.collect_ocr_rows())
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.disable_buttons(keep_start=True)
        self.started_at = time.monotonic()
        self.status.set("正在進行產品重複比對。")
        self.set_detail(f"使用產品資料：{csv_path}\n正在比對已建檔商品與相似候選。")
        self.run_in_thread(self.worker_product_match)

    def worker_product_match(self) -> None:
        workflow.run_match(self.state)
        self.after(0, self.product_match_finished)

    def product_match_finished(self) -> None:
        self.worker = None
        self.workflow_stage = "match"
        self.adjustment_draft_rows = None
        self.match_changed_since_back = False
        self.product_match_ready = True
        self.mark_step("match", True)
        self.populate_ocr_tree(workflow.load_ocr_confirm_rows(self.state))
        self.current_output = self.state.match_xlsx
        self.status.set("產品重複比對完成，請確認已建檔與相似候選。")
        self.set_detail(
            "一般候選為相似度 60% 以上；低相似候選為 55%～59.99%，只供人工查看。\n"
            "沒有適合候選時，可選取單筆後按「更低候選單筆檢查」，查看約 35%～54.99% 的結果。\n"
            "相似候選不會自動勾選；只有精確命中或有效產品代號才會自動勾選。\n"
            "候選超過一筆時，選取商品後按「選擇候選」，或直接雙擊「相似候選」欄。\n"
            "候選清單可用滑鼠滾輪瀏覽，雙擊候選即可帶入代號與正式品名。\n"
            "若候選沒有正確商品，可雙擊第一筆「已建檔代號」後貼上整欄代號。\n"
            "確認完成後按「確認文字並開始調整」。"
        )
        self.start_button.configure(state="normal")
        self.match_button.configure(state="normal")
        self.review_button.configure(state="normal")
        self.candidate_button.configure(state="normal")
        self.deep_candidate_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.refresh_navigation_buttons()

    def prepare_review(self) -> None:
        if not self.state.match_xlsx:
            messagebox.showwarning(APP_TITLE, "請先完成 OCR 與產品比對。")
            return
        if self.adjustment_draft_rows is not None and not self.match_changed_since_back:
            draft_rows = self.adjustment_draft_rows
            self.adjustment_draft_rows = None
            self.review_prepared(draft_rows)
            self.status.set("已回到先前保留的名稱調整內容。")
            return
        if self.match_changed_since_back:
            self.adjustment_draft_rows = None
        try:
            workflow.save_ocr_confirm_rows(self.state, self.collect_ocr_rows())
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.disable_buttons(keep_start=True)
        self.status.set("正在套用命名規則。")
        self.set_detail("已確認 OCR 文字，正在產生需要人工調整的新品/類似品清單。")
        self.run_in_thread(self.worker_prepare_review)

    def worker_prepare_review(self) -> None:
        rows = workflow.prepare_review_table(self.state)
        self.after(0, self.review_prepared, rows)

    def review_prepared(self, rows: list[AdjustmentRow]) -> None:
        self.worker = None
        self.workflow_stage = "adjust"
        self.match_changed_since_back = False
        self.set_adjustment_table()
        self.populate_adjustment_tree(rows)
        self.mark_step("review", False)
        self.current_output = self.state.match_xlsx
        if not rows:
            self.status.set("全部商品已確認為已建檔，正在準備正式輸出。")
            self.set_detail("沒有需要名稱調整的新品或類似品，正在背景建立正式輸出所需的內部交接資料。")
            self.run_in_thread(self.worker_generate_adjusted)
            return
        self.status.set("請確認需要建檔或調整名稱的商品。")
        lines = [
            "單擊任一格即可選取內容，使用 Ctrl+C 複製、Ctrl+V 貼上；切換格子時會保存修改。",
            "產品代號、產品名稱、大類代號與大類名稱都支援由目前列往下整欄貼上。",
            "大類已拆成「大類代號」與「大類名稱」，例如：38｜公仔/吊卡/PVC。",
            "需要拆成多個品項時，選取來源列後按「拆分/新增品項」，再調整拆分後數量與金額。",
            "確認後按「確認資料」。",
        ]
        if self.state.excluded_items:
            lines.append("")
            lines.append("已排除項目：")
            lines.extend(f"- {name}" for name in self.state.excluded_items)
        self.set_detail("\n".join(lines))
        self.start_button.configure(state="normal")
        self.review_button.configure(state="normal")
        self.adjust_button.configure(state="normal")
        self.add_item_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.refresh_navigation_buttons()

    def generate_adjusted(self) -> None:
        if self.table_mode != "adjust":
            messagebox.showwarning(APP_TITLE, "請先確認 OCR 文字並進入名稱調整。")
            return
        rows = self.collect_adjustment_rows()
        try:
            workflow.save_adjustment_rows(self.state, rows)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.disable_buttons(keep_start=True)
        self.status.set("正在確認資料並準備正式輸出。")
        self.set_detail("正在檢查產品代號、正式名稱與大類。")
        self.run_in_thread(self.worker_generate_adjusted)

    def worker_generate_adjusted(self) -> None:
        result = workflow.generate_adjusted_xlsx(self.state)
        self.after(0, self.adjusted_finished, result)

    def adjusted_finished(self, result: dict[str, object]) -> None:
        self.worker = None
        self.workflow_stage = "ready"
        self.mark_step("review", True)
        self.current_output = self.state.match_xlsx
        self.status.set("資料確認完成，可以產生正式檔。")
        self.set_detail("產品代號、正式名稱與大類檢查通過。內部交接資料已在背景準備完成。")
        self.start_button.configure(state="normal")
        self.review_button.configure(state="normal")
        self.adjust_button.configure(state="normal")
        self.add_item_button.configure(state="disabled")
        self.build_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.refresh_navigation_buttons()

    def build_outputs(self) -> None:
        if not self.state.adjusted_xlsx:
            messagebox.showwarning(APP_TITLE, "請先完成資料確認。")
            return
        confirmed = messagebox.askyesno(
            APP_TITLE,
            "請確認：進貨單資料已檢查並調整完成，可以進行建檔與採購匯入檔產生？",
        )
        if not confirmed:
            return
        self.disable_buttons(keep_start=True)
        self.status.set("正在產生建檔用與採購單匯入檔。")
        self.set_detail("正在套用範本與廠商代號。")
        self.run_in_thread(self.worker_build_outputs)

    def worker_build_outputs(self) -> None:
        result = workflow.build_import_files(self.state)
        self.after(0, self.build_finished, result)

    def build_finished(self, result: dict[str, object]) -> None:
        self.worker = None
        self.workflow_stage = "built"
        self.mark_step("build", True)
        self.current_output = self.state.purchase_file or self.state.new_product_file
        self.status.set("正式檔已產生。")
        tax_codes = result.get("taxAdjustedProductCodes") or []
        lines = [
            f"建檔用檔案：{self.state.new_product_file if self.state.new_product_file else '無（全部已建檔）'}",
            f"採購單用檔案：{self.state.purchase_file}",
            f"商品筆數：{result.get('rowCount')}",
            f"新品筆數：{result.get('newProductRowCount')}",
            f"已建檔筆數：{result.get('existingProductRowCount')}",
            f"採購筆數：{result.get('rowCount')}",
            f"廠商代號：{result.get('vendorCode')}",
            f"含稅調整：{', '.join(tax_codes) if tax_codes else '無'}",
            f"排除項目：{', '.join(self.state.excluded_items) if self.state.excluded_items else '無'}",
            "特殊處理：依現有建檔/採購規則輸出",
            "來源圖片：",
        ]
        source_paths = (
            self.state.processed_image_paths
            or self.state.image_paths
            or ([self.state.processed_image_path or self.state.image_path] if (self.state.processed_image_path or self.state.image_path) else [])
        )
        lines.extend(f"- {path}" for path in source_paths)
        if result.get("internalAdjustedFileRemoved"):
            lines.append("內部交接資料：正式檔驗證成功後已自動清除")
        elif result.get("internalAdjustedCleanupWarning"):
            lines.append(f"內部交接資料：保留（自動清除失敗：{result.get('internalAdjustedCleanupWarning')}）")
        self.set_detail("\n".join(lines))
        self.start_button.configure(state="normal")
        self.cleanup_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.refresh_navigation_buttons()
        if messagebox.askyesno(APP_TITLE, "正式成品已產生。要現在清理本次中間檔嗎？"):
            self.cleanup_files()

    def cleanup_files(self) -> None:
        try:
            deleted = workflow.cleanup_intermediate_files(self.state)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.mark_step("cleanup", True)
        self.workflow_stage = "cleaned"
        self.status.set("中間檔清理完成。")
        self.set_detail("已清理：\n" + "\n".join(str(path) for path in deleted) if deleted else "沒有可清理的中間檔。")
        self.cleanup_button.configure(state="disabled")
        self.refresh_navigation_buttons()

    def run_in_thread(self, target) -> None:
        def runner() -> None:
            try:
                target()
            except Exception as exc:
                self.after(0, self.worker_failed, str(exc))

        self.worker = threading.Thread(target=runner, daemon=True)
        self.worker.start()
        self.after(500, self.update_elapsed_status)

    def worker_failed(self, message: str) -> None:
        self.worker = None
        self.status.set("流程停止。")
        self.set_detail(message[-5000:])
        self.start_button.configure(state="normal")
        if self.state.raw_xlsx and self.state.raw_xlsx.exists() and not self.state.match_xlsx:
            self.match_button.configure(state="normal")
        if self.state.match_xlsx and self.state.match_xlsx.exists():
            self.review_button.configure(state="normal")
            if self.table_mode == "ocr":
                self.candidate_button.configure(state="normal")
                self.deep_candidate_button.configure(
                    state="normal" if self.product_match_ready else "disabled"
                )
        if self.state.adjusted_xlsx and self.state.adjusted_xlsx.exists():
            self.build_button.configure(state="normal")
        self.open_button.configure(state="normal" if self.current_output and self.current_output.exists() else "disabled")
        self.refresh_navigation_buttons()

    def update_elapsed_status(self) -> None:
        if not self.worker or not self.worker.is_alive():
            return
        elapsed = int(time.monotonic() - self.started_at)
        self.status.set(f"處理中，已執行 {elapsed} 秒。")
        self.after(1000, self.update_elapsed_status)

    def populate_ocr_tree(self, rows: list[OcrConfirmRow]) -> None:
        self.set_ocr_table()
        self.clear_tree()
        self.ocr_candidate_values = {}
        for row in rows:
            item_id = str(row.excel_row)
            self.ocr_candidate_values[item_id] = row.candidates
            candidate_display = self.candidate_display_text(row.candidates)
            self.tree.insert(
                "",
                "end",
                iid=item_id,
                values=(
                    "☑" if row.is_existing else "☐",
                    row.raw_name,
                    row.quantity,
                    row.unit_cost,
                    row.amount,
                    row.matched_code,
                    row.matched_name,
                    row.status,
                    candidate_display,
                ),
            )
        self.refresh_navigation_buttons()

    def candidate_display_text(self, candidates_text: str) -> str:
        options = workflow.parse_candidates(candidates_text)
        if not options:
            return ""
        tier_names = {
            "normal": "一般候選",
            "low": "低相似候選",
            "deep": "更低候選",
        }
        count_names = {
            "normal": "一般",
            "low": "低相似",
            "deep": "更低",
        }
        count_parts = [
            f"{count_names[tier]} {sum(1 for option in options if option[3] == tier)}"
            for tier in ("normal", "low", "deep")
            if any(option[3] == tier for option in options)
        ]
        first = options[0]
        summary = f"[{tier_names[first[3]]}] {first[0]} {first[1]}"
        if len(options) > 1:
            summary += f"（{'｜'.join(count_parts)}，雙擊選擇）"
        return summary

    def populate_adjustment_tree(self, rows: list[AdjustmentRow]) -> None:
        self.clear_tree()
        self.adjustment_source_rows = {}
        category_names = workflow.category_name_map()
        for row in rows:
            self.adjustment_source_rows[row.row_id] = row.source_row
            category_code = workflow.category_code(row.category)
            category_name = category_names.get(category_code, "")
            if not category_name:
                display_parts = str(row.category_display or "").split(maxsplit=1)
                if len(display_parts) == 2 and display_parts[0] == category_code:
                    category_name = display_parts[1]
            self.tree.insert(
                "",
                "end",
                iid=row.row_id,
                values=(
                    row.product_code,
                    row.name,
                    category_code,
                    category_name,
                    row.quantity,
                    row.unit_cost,
                    row.amount,
                    row.status,
                ),
            )
        self.refresh_navigation_buttons()

    def clear_tree(self) -> None:
        self.cancel_cell_edit()
        self.active_cell_item = ""
        self.active_cell_column = ""
        for item in self.tree.get_children():
            self.tree.delete(item)

    def select_column_for_copy(self, selected_column: str) -> None:
        columns = list(self.tree["columns"])
        if selected_column not in columns:
            return
        self.active_cell_column = f"#{columns.index(selected_column) + 1}"
        self.status.set(f"已選擇「{selected_column}」欄，按「複製整欄」即可複製全部商品列。")
        self.refresh_navigation_buttons()

    def show_column_copy_menu(self, event) -> str:
        region = self.tree.identify("region", event.x, event.y)
        if region not in {"cell", "heading"}:
            return "break"
        column_id = self.tree.identify_column(event.x)
        if not column_id:
            return "break"
        self.active_cell_column = column_id
        if region == "cell":
            item = self.tree.identify_row(event.y)
            if item:
                self.active_cell_item = item
                self.tree.selection_set(item)
                self.tree.focus(item)
        try:
            self.column_copy_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.column_copy_menu.grab_release()
        return "break"

    def copy_selected_column(self, _event=None) -> str:
        children = list(self.tree.get_children())
        if not children:
            messagebox.showwarning(APP_TITLE, "目前表格沒有可複製的商品資料。")
            return "break"
        columns = list(self.tree["columns"])
        column_id = self.active_cell_column or ("#2" if len(columns) >= 2 else "#1")
        try:
            column_index = int(column_id.removeprefix("#")) - 1
        except ValueError:
            column_index = 1 if len(columns) >= 2 else 0
        if column_index < 0 or column_index >= len(columns):
            column_index = 1 if len(columns) >= 2 else 0
        column_id = f"#{column_index + 1}"
        self.active_cell_column = column_id
        column_name = columns[column_index]
        values = [
            str(list(self.tree.item(item, "values"))[column_index])
            for item in children
        ]
        self.clipboard_clear()
        # Tk will expose LF correctly to Excel on Windows.  Supplying CRLF here
        # can be normalized a second time by the clipboard bridge and produce
        # an empty spreadsheet row between product names.
        self.clipboard_append("\n".join(values))
        self.update_idletasks()
        self.status.set(f"已複製「{column_name}」整欄，共 {len(values)} 筆，可直接貼到 Excel。")
        return "break"

    def delete_selected_items(self, _event=None) -> str:
        if self.workflow_stage not in {"ocr", "match", "adjust"}:
            messagebox.showwarning(APP_TITLE, "目前階段不能刪除商品項目。")
            return "break"
        selected = list(self.tree.selection())
        if not selected and self.active_cell_item:
            selected = [self.active_cell_item]
        if not selected:
            messagebox.showwarning(APP_TITLE, "請先選取要刪除的商品列。")
            return "break"
        name_index = 1
        selected_names = [
            str(list(self.tree.item(item, "values"))[name_index])
            for item in selected
            if self.tree.exists(item)
        ]
        preview = "\n".join(f"- {name}" for name in selected_names[:8])
        if len(selected_names) > 8:
            preview += f"\n- 另有 {len(selected_names) - 8} 筆"
        if not messagebox.askyesno(
            APP_TITLE,
            f"確定刪除選取的 {len(selected_names)} 筆商品？\n\n{preview}\n\n"
            "刪除後會同步更新目前的中間 Excel 與總額。",
        ):
            return "break"

        for item in selected:
            self.ocr_candidate_values.pop(str(item), None)
            self.adjustment_source_rows.pop(str(item), None)
            if self.tree.exists(item):
                self.tree.delete(item)
        self.active_cell_item = ""

        try:
            if self.table_mode == "ocr":
                rows = self.collect_ocr_rows()
                if self.product_match_ready and self.state.match_xlsx:
                    workflow.save_ocr_confirm_rows(self.state, rows)
                    self.populate_ocr_tree(workflow.load_ocr_confirm_rows(self.state))
                else:
                    workflow.save_raw_ocr_rows(self.state, rows)
                    raw_state = WorkflowState(raw_xlsx=self.state.raw_xlsx)
                    self.populate_ocr_tree(workflow.load_ocr_confirm_rows(raw_state))
            else:
                workflow.save_adjustment_rows(
                    self.state,
                    self.collect_adjustment_rows(),
                )
                self.populate_adjustment_tree(workflow.load_adjustment_rows(self.state))
                self.state.row_count = len(workflow.load_ocr_confirm_rows(self.state))
        except Exception as exc:
            if self.table_mode == "ocr":
                self.populate_ocr_tree(workflow.load_ocr_confirm_rows(self.state))
            else:
                self.populate_adjustment_tree(workflow.load_adjustment_rows(self.state))
            messagebox.showerror(APP_TITLE, f"刪除失敗，已重新載入原資料：\n{exc}")
            return "break"

        self.status.set(
            f"已刪除 {len(selected_names)} 筆商品；目前共 {self.state.row_count} 筆，"
            f"總額 {self.state.invoice_total}。"
        )
        self.refresh_navigation_buttons()
        return "break"

    def refresh_navigation_buttons(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            self.back_button.configure(state="disabled")
            self.copy_column_button.configure(state="disabled")
            self.delete_item_button.configure(state="disabled")
            return
        can_go_back = self.workflow_stage in {"match", "adjust", "ready"}
        self.back_button.configure(state="normal" if can_go_back else "disabled")
        self.copy_column_button.configure(
            state="normal" if self.tree.get_children() else "disabled"
        )
        can_delete = self.workflow_stage in {"ocr", "match", "adjust"}
        self.delete_item_button.configure(
            state="normal" if can_delete and self.tree.get_children() else "disabled"
        )

    def mark_match_table_changed(self) -> None:
        if self.workflow_stage == "match" and self.adjustment_draft_rows is not None:
            self.match_changed_since_back = True

    def go_back(self) -> None:
        if self.worker is not None and self.worker.is_alive():
            messagebox.showwarning(APP_TITLE, "目前仍在處理中，完成後才能回到上一步。")
            return
        self.cancel_cell_edit()
        if self.workflow_stage == "ready":
            self.workflow_stage = "adjust"
            self.mark_step("review", False)
            self.disable_buttons(keep_start=True)
            self.start_button.configure(state="normal")
            self.review_button.configure(state="normal")
            self.adjust_button.configure(state="normal")
            self.add_item_button.configure(state="normal")
            self.open_button.configure(state="normal")
            self.current_output = self.state.match_xlsx
            self.status.set("已回到名稱調整，原本輸入內容仍保留。")
            self.set_detail(
                "可繼續修改產品代號、產品名稱、大類代號、大類名稱、數量、成本與金額。\n"
                "完成後再按「確認資料」，重新建立正式輸出所需資料。"
            )
        elif self.workflow_stage == "adjust":
            self.adjustment_draft_rows = self.collect_adjustment_rows()
            self.match_changed_since_back = False
            self.product_match_ready = True
            self.populate_ocr_tree(workflow.load_ocr_confirm_rows(self.state))
            self.workflow_stage = "match"
            self.disable_buttons(keep_start=True)
            self.start_button.configure(state="normal")
            self.match_button.configure(state="normal")
            self.review_button.configure(state="normal")
            self.candidate_button.configure(state="normal")
            self.deep_candidate_button.configure(state="normal")
            self.open_button.configure(state="normal")
            self.current_output = self.state.match_xlsx
            self.status.set("已回到已建檔勾選；名稱調整內容已暫時保留。")
            self.set_detail(
                "若不修改已建檔勾選，按「確認文字並開始調整」會恢復剛才的名稱調整內容。\n"
                "若修改比對結果，程式會依新的勾選重新產生名稱調整資料。"
            )
        elif self.workflow_stage == "match":
            if self.state.raw_xlsx is None or not self.state.raw_xlsx.exists():
                messagebox.showerror(APP_TITLE, "找不到 OCR 原始檔，無法回到上一步。")
                return
            self.adjustment_draft_rows = None
            self.match_changed_since_back = False
            self.product_match_ready = False
            raw_state = WorkflowState(raw_xlsx=self.state.raw_xlsx)
            self.populate_ocr_tree(workflow.load_ocr_confirm_rows(raw_state))
            self.workflow_stage = "ocr"
            self.mark_step("match", False)
            self.disable_buttons(keep_start=True)
            self.start_button.configure(state="normal")
            self.match_button.configure(state="normal")
            self.open_button.configure(state="normal")
            self.current_output = self.state.raw_xlsx
            self.status.set("已回到 OCR 原文確認。")
            self.set_detail(
                "可重新修改 OCR 品名、數量、進價與金額。\n"
                "完成後按「開始產品重複比對」，程式會依修正後內容重新比對。"
            )
        else:
            messagebox.showinfo(APP_TITLE, "目前已是第一個可操作步驟。")
        self.refresh_navigation_buttons()

    def collect_ocr_rows(self) -> list[OcrConfirmRow]:
        rows: list[OcrConfirmRow] = []
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            rows.append(
                OcrConfirmRow(
                    excel_row=int(item),
                    is_existing=str(values[0]).strip() == "☑",
                    raw_name=str(values[1]).strip(),
                    quantity=str(values[2]).strip(),
                    unit_cost=str(values[3]).strip(),
                    amount=str(values[4]).strip(),
                    matched_code=str(values[5]).strip(),
                    matched_name=str(values[6]).strip(),
                    status=str(values[7]).strip(),
                    candidates=self.ocr_candidate_values.get(str(item), str(values[8]).strip()),
                )
            )
        return rows

    def collect_adjustment_rows(self) -> list[AdjustmentRow]:
        rows: list[AdjustmentRow] = []
        for item in self.tree.get_children():
            values = list(self.tree.item(item, "values"))
            excel_row = int(item) if str(item).isdigit() else 0
            rows.append(
                AdjustmentRow(
                    row_id=str(item),
                    excel_row=excel_row,
                    source_row=self.adjustment_source_rows.get(str(item), excel_row),
                    product_code=str(values[0]).strip(),
                    name=str(values[1]).strip(),
                    category=workflow.category_code(str(values[2]).strip()),
                    category_display=" ".join(
                        part
                        for part in (
                            workflow.category_code(str(values[2]).strip()),
                            str(values[3]).strip(),
                        )
                        if part
                    ),
                    quantity=str(values[4]).strip(),
                    unit_cost=str(values[5]).strip(),
                    amount=str(values[6]).strip(),
                    status=str(values[7]).strip(),
                )
            )
        return rows

    def open_candidate_picker(self) -> None:
        if self.table_mode != "ocr" or not self.product_match_ready:
            messagebox.showwarning(APP_TITLE, "候選選擇只在 OCR／已建檔確認階段使用。")
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(APP_TITLE, "請先選取要查看候選的商品。")
            return
        self.open_candidate_picker_for_item(selected[0])

    def start_lower_candidate_search(self) -> None:
        if self.table_mode != "ocr" or not self.product_match_ready:
            messagebox.showwarning(APP_TITLE, "更低候選檢查只在已完成產品比對後使用。")
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(APP_TITLE, "請先選取要單獨檢查的商品。")
            return
        item = selected[0]
        values = list(self.tree.item(item, "values"))
        source_name = str(values[1]).strip()
        if not source_name:
            messagebox.showwarning(APP_TITLE, "這筆商品沒有 OCR 品名，請先填入名稱。")
            return
        self.disable_buttons(keep_start=True)
        self.started_at = time.monotonic()
        self.status.set("正在執行更低候選單筆檢查。")
        self.set_detail(
            f"檢查品名：{source_name}\n"
            "只搜尋約 35%～54.99% 的更低候選；結果只供人工查看，不會自動勾選。"
        )
        self.run_in_thread(lambda: self.worker_lower_candidate_search(item, source_name))

    def worker_lower_candidate_search(self, item: str, source_name: str) -> None:
        options = workflow.find_lower_similarity_candidates(source_name)
        self.after(0, self.lower_candidate_search_finished, item, options)

    def lower_candidate_search_finished(
        self,
        item: str,
        options: list[tuple[str, str, str, str]],
    ) -> None:
        self.worker = None
        if not self.tree.exists(item):
            self.worker_failed("原商品列已不存在，無法顯示更低候選。")
            return
        existing_lines = [
            line
            for line in self.ocr_candidate_values.get(str(item), "").splitlines()
            if line.strip() and not line.strip().startswith("[更低候選]")
        ]
        deep_lines = [
            f"[更低候選] {code} {name} ({score})"
            for code, name, score, _tier in options
        ]
        combined = "\n".join(existing_lines + deep_lines)
        self.ocr_candidate_values[str(item)] = combined
        values = list(self.tree.item(item, "values"))
        values[8] = self.candidate_display_text(combined)
        self.tree.item(item, values=values)
        self.mark_match_table_changed()
        self.start_button.configure(state="normal")
        self.match_button.configure(state="normal")
        self.review_button.configure(state="normal")
        self.candidate_button.configure(state="normal")
        self.deep_candidate_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.refresh_navigation_buttons()
        if not options:
            self.status.set("更低候選單筆檢查完成，仍找不到 35% 以上的候選。")
            messagebox.showinfo(
                APP_TITLE,
                "這筆商品在更低門檻下仍沒有候選。\n\n"
                "建議確認 OCR 品名是否完整，或將它保留為新品。",
            )
            return
        self.status.set(f"更低候選單筆檢查完成，找到 {len(options)} 筆，只供人工選擇。")
        self.tree.selection_set(item)
        self.tree.focus(item)
        self.open_candidate_picker_for_item(item)

    def open_candidate_picker_for_item(self, item: str) -> None:
        candidates_text = self.ocr_candidate_values.get(str(item), "")
        options = workflow.parse_candidates(candidates_text)
        if not options:
            messagebox.showinfo(
                APP_TITLE,
                "這筆商品沒有相似候選。\n\n"
                "若你確定已建檔，可勾選「已建檔」並在「已建檔代號」輸入六位代號；"
                "程式會從今天的產品資料自動帶入正式品名。",
            )
            return

        dialog = tk.Toplevel(self)
        dialog.title("選擇已建檔候選")
        dialog.geometry("760x380")
        dialog.minsize(620, 300)
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(
            dialog,
            text=(
                "可用滑鼠滾輪瀏覽；雙擊候選或按「套用選取」。\n"
                "低相似與更低候選只供人工查看，不會自動勾選為已建檔。"
            ),
            font=("Microsoft JhengHei UI", 10, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 6))

        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        candidate_list = tk.Listbox(
            list_frame,
            activestyle="dotbox",
            exportselection=False,
            font=("Microsoft JhengHei UI", 10),
        )
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=candidate_list.yview)
        candidate_list.configure(yscrollcommand=scrollbar.set)
        candidate_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for code, name, score, tier in options:
            tier_text = {
                "normal": "一般候選",
                "low": "低相似候選",
                "deep": "更低候選",
            }[tier]
            score_text = f"｜相似度 {score}" if score else ""
            candidate_list.insert("end", f"[{tier_text}] {code}｜{name}{score_text}")
            if tier in {"low", "deep"}:
                candidate_list.itemconfig(candidate_list.size() - 1, foreground="#b05a00")
        candidate_list.selection_set(0)
        candidate_list.activate(0)
        candidate_list.focus_set()

        def scroll_candidates(event) -> str:
            step = -1 if event.delta > 0 else 1
            candidate_list.yview_scroll(step, "units")
            return "break"

        def apply_selected(_event=None) -> None:
            selection = candidate_list.curselection()
            if not selection:
                messagebox.showwarning(APP_TITLE, "請先選取一筆候選。", parent=dialog)
                return
            code, name, _score, _tier = options[int(selection[0])]
            values = list(self.tree.item(item, "values"))
            values[0] = "☑"
            values[5] = code
            values[6] = name
            values[7] = "已建檔"
            self.tree.item(item, values=values)
            self.mark_match_table_changed()
            self.tree.selection_set(item)
            self.tree.focus(item)
            self.status.set(f"已套用候選：{code} {name}")
            dialog.destroy()

        candidate_list.bind("<MouseWheel>", scroll_candidates)
        candidate_list.bind("<Double-1>", apply_selected)

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(button_frame, text="套用選取", command=apply_selected).pack(side="right")
        ttk.Button(button_frame, text="取消", command=dialog.destroy).pack(side="right", padx=(0, 8))

    def split_selected_item(self) -> None:
        if self.table_mode != "adjust":
            messagebox.showwarning(APP_TITLE, "請先進入名稱調整頁。")
            return
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning(APP_TITLE, "請先選取要拆分的來源品項。")
            return
        source_item = selected[0]
        values = list(self.tree.item(source_item, "values"))
        self.new_row_counter += 1
        new_id = f"new-{self.new_row_counter}"
        source_row = self.adjustment_source_rows.get(source_item, int(source_item) if source_item.isdigit() else 0)
        self.adjustment_source_rows[new_id] = source_row
        new_values = values[:]
        new_values[0] = ""
        self.tree.insert("", "end", iid=new_id, values=new_values)
        self.status.set("已新增拆分品項，請調整原列與新列的數量、成本與金額。")

    def select_adjustment_cell(self, event) -> None:
        if self.tree.identify("region", event.x, event.y) == "cell":
            item = self.tree.identify_row(event.y)
            column_id = self.tree.identify_column(event.x)
            if item and column_id:
                self.active_cell_item = item
                self.active_cell_column = column_id
        if self.table_mode == "adjust":
            self.begin_edit_cell(event, select_all=True)

    def begin_edit_cell(self, event, select_all: bool = False) -> None:
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item:
            return
        self.active_cell_item = item
        self.active_cell_column = column_id
        column_index = int(column_id.replace("#", "")) - 1
        read_only = False
        if self.table_mode == "ocr":
            if self.product_match_ready:
                if column_id == "#9":
                    self.open_candidate_picker_for_item(item)
                    return
                if column_id == "#1":
                    values = list(self.tree.item(item, "values"))
                    if str(values[0]).strip() == "☑":
                        values[0] = "☐"
                    else:
                        current_code = workflow.normalize_product_code(values[5])
                        candidates = workflow.parse_candidates(self.ocr_candidate_values.get(str(item), ""))
                        if current_code:
                            values[5] = current_code
                            values[6] = workflow.catalog_product_name(current_code) or str(values[6]).strip()
                        elif len(candidates) > 1:
                            self.open_candidate_picker_for_item(item)
                            return
                        elif candidates:
                            code, name, _score, _tier = candidates[0]
                            values[5] = code
                            values[6] = name
                        values[0] = "☑"
                        values[7] = "已建檔"
                    self.tree.item(item, values=values)
                    self.mark_match_table_changed()
                    return
                editable_columns = {"#2", "#3", "#4", "#5", "#6", "#7"}
            else:
                editable_columns = {"#2", "#3", "#4", "#5"}
        else:
            editable_columns = {"#1", "#2", "#3", "#4", "#5", "#6", "#7"}
            read_only = column_id == "#8"
        if column_id not in editable_columns and not read_only:
            return
        bbox = self.tree.bbox(item, column_id)
        if not bbox:
            return
        values = list(self.tree.item(item, "values"))
        x, y, width, height = bbox
        if self.edit_entry is not None:
            self.edit_entry.destroy()
        self.edit_entry = ttk.Entry(self.tree)
        self.edit_entry.insert(0, str(values[column_index]))
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.focus_set()
        if select_all:
            self.edit_entry.selection_range(0, tk.END)
        if read_only:
            self.edit_entry.configure(state="readonly")

        def paste_product_codes(_event=None) -> str:
            if self.edit_entry is None:
                return "break"
            try:
                clipboard_text = self.clipboard_get()
            except tk.TclError:
                return "break"
            result = self.apply_pasted_product_codes(item, clipboard_text)
            entry = self.edit_entry
            self.edit_entry = None
            entry.destroy()
            messages = [f"已從目前列往下填入 {result['applied']} 筆產品代號。"]
            if result["overflow"]:
                messages.append(f"另有 {result['overflow']} 筆超出目前表格列數，未填入。")
            if result["invalid"]:
                messages.append(f"無法辨識的內容：{', '.join(result['invalid'])}")
            if result["unknown"]:
                messages.append(f"產品資料找不到的代號：{', '.join(result['unknown'])}")
            self.status.set(" ".join(messages))
            if result["overflow"] or result["invalid"] or result["unknown"]:
                messagebox.showwarning(APP_TITLE, "\n".join(messages))
            return "break"

        def paste_column_values(_event=None) -> str:
            if self.edit_entry is None:
                return "break"
            try:
                clipboard_text = self.clipboard_get()
            except tk.TclError:
                return "break"
            result = self.apply_pasted_column_values(
                item,
                column_index,
                clipboard_text,
            )
            entry = self.edit_entry
            self.edit_entry = None
            entry.destroy()
            column_name = list(self.tree["columns"])[column_index]
            messages = [f"已從目前列往下填入「{column_name}」{result['applied']} 筆。"]
            if result["overflow"]:
                messages.append(f"另有 {result['overflow']} 筆超出目前表格列數，未填入。")
            self.status.set(" ".join(messages))
            if result["overflow"]:
                messagebox.showwarning(APP_TITLE, "\n".join(messages))
            return "break"

        def commit(_event=None) -> None:
            if self.edit_entry is None:
                return
            if not read_only:
                new_value = self.edit_entry.get()
                if self.table_mode == "adjust" and column_id == "#3":
                    category_code = workflow.category_code(new_value)
                    values[2] = category_code
                    values[3] = workflow.category_name_map().get(
                        category_code,
                        str(values[3]).strip(),
                    )
                elif self.table_mode == "adjust" and column_id == "#4":
                    category_name = str(new_value).strip()
                    values[3] = category_name
                    reverse_names = {
                        name: code for code, name in workflow.category_name_map().items()
                    }
                    if category_name in reverse_names:
                        values[2] = reverse_names[category_name]
                else:
                    values[column_index] = new_value
                if self.table_mode == "ocr" and column_id == "#6":
                    normalized_code = workflow.normalize_product_code(new_value)
                    if normalized_code:
                        values[5] = normalized_code
                        official_name = workflow.catalog_product_name(normalized_code)
                        if official_name:
                            values[0] = "☑"
                            values[6] = official_name
                            values[7] = "已建檔"
                self.tree.item(item, values=values)
                if self.table_mode == "ocr":
                    self.mark_match_table_changed()
            self.edit_entry.destroy()
            self.edit_entry = None

        is_product_code_column = (
            self.table_mode == "ocr" and column_id == "#6"
        ) or (
            self.table_mode == "adjust" and column_id == "#1"
        )
        if is_product_code_column:
            self.edit_entry.bind("<<Paste>>", paste_product_codes)
        is_multi_value_column = column_id in editable_columns and not read_only
        if is_product_code_column:
            is_multi_value_column = False
        if is_multi_value_column:
            self.edit_entry.bind("<<Paste>>", paste_column_values)
        self.edit_entry.bind("<Return>", commit)
        self.edit_entry.bind("<FocusOut>", commit)
        self.edit_entry.bind("<Escape>", lambda _event: self.cancel_cell_edit())

    def cancel_cell_edit(self) -> None:
        if self.edit_entry is not None:
            self.edit_entry.destroy()
            self.edit_entry = None

    def apply_pasted_product_codes(self, start_item: str, clipboard_text: str) -> dict[str, object]:
        codes, invalid = workflow.parse_pasted_product_codes(clipboard_text)
        children = list(self.tree.get_children())
        if start_item not in children:
            return {"applied": 0, "overflow": len(codes), "invalid": invalid, "unknown": []}
        start_index = children.index(start_item)
        target_items = children[start_index : start_index + len(codes)]
        product_names = workflow.catalog_product_names(codes) if self.table_mode == "ocr" else {}
        unknown: list[str] = []

        for target_item, code in zip(target_items, codes):
            row_values = list(self.tree.item(target_item, "values"))
            if self.table_mode == "ocr":
                row_values[5] = code
                official_name = product_names.get(code, "")
                if official_name:
                    row_values[0] = "☑"
                    row_values[6] = official_name
                    row_values[7] = "已建檔"
                else:
                    row_values[0] = "☐"
                    unknown.append(code)
            else:
                row_values[0] = code
            self.tree.item(target_item, values=row_values)

        if target_items:
            if self.table_mode == "ocr":
                self.mark_match_table_changed()
            self.tree.selection_set(target_items[-1])
            self.tree.focus(target_items[-1])
            self.tree.see(target_items[-1])
        return {
            "applied": len(target_items),
            "overflow": max(0, len(codes) - len(target_items)),
            "invalid": invalid,
            "unknown": unknown,
        }

    def apply_pasted_column_values(
        self,
        start_item: str,
        column_index: int,
        clipboard_text: str,
    ) -> dict[str, int]:
        pasted_values = workflow.parse_pasted_column_values(clipboard_text)
        children = list(self.tree.get_children())
        if start_item not in children:
            return {"applied": 0, "overflow": len(pasted_values)}
        start_index = children.index(start_item)
        target_items = children[start_index : start_index + len(pasted_values)]
        category_names = workflow.category_name_map()
        reverse_category_names = {name: code for code, name in category_names.items()}

        for target_item, pasted_value in zip(target_items, pasted_values):
            row_values = list(self.tree.item(target_item, "values"))
            if self.table_mode == "adjust" and column_index == 2:
                category_code = workflow.category_code(pasted_value)
                row_values[2] = category_code
                if category_code in category_names:
                    row_values[3] = category_names[category_code]
            elif self.table_mode == "adjust" and column_index == 3:
                row_values[3] = pasted_value
                if pasted_value in reverse_category_names:
                    row_values[2] = reverse_category_names[pasted_value]
            else:
                row_values[column_index] = pasted_value
            self.tree.item(target_item, values=row_values)

        if target_items:
            if self.table_mode == "ocr":
                self.mark_match_table_changed()
            self.tree.selection_set(target_items[-1])
            self.tree.focus(target_items[-1])
            self.tree.see(target_items[-1])
        return {
            "applied": len(target_items),
            "overflow": max(0, len(pasted_values) - len(target_items)),
        }

    def disable_buttons(self, keep_start: bool = False) -> None:
        for button in [
            self.match_button,
            self.review_button,
            self.candidate_button,
            self.deep_candidate_button,
            self.add_item_button,
            self.adjust_button,
            self.build_button,
            self.cleanup_button,
            self.open_button,
            self.back_button,
            self.copy_column_button,
            self.delete_item_button,
        ]:
            button.configure(state="disabled")
        if not keep_start:
            self.start_button.configure(state="disabled")

    def reset_steps(self) -> None:
        labels = {
            "ocr": "1 OCR 原文確認",
            "match": "2 已建檔勾選",
            "review": "3 名稱調整",
            "build": "4 正式輸出",
            "cleanup": "5 清理中間檔",
        }
        for key, text in labels.items():
            self.step_labels[key].configure(text=f"□ {text}")

    def mark_step(self, key: str, done: bool) -> None:
        text = self.step_labels[key].cget("text")[2:]
        self.step_labels[key].configure(text=("■ " if done else "▣ ") + text)

    def set_detail(self, text: str) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", text)
        self.detail_text.configure(state="disabled")

    def open_current_output(self) -> None:
        if self.current_output and self.current_output.exists():
            os.startfile(self.current_output)

    def on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_TITLE, "流程還在執行中，確定要關閉嗎？"):
                return
        self.destroy()


def run_self_test() -> int:
    synced = workflow.sync_reference_data()
    issues = workflow.validate_runtime()
    csv_path = workflow.product_csv()
    current_csv = False
    csv_error = ""
    if csv_path:
        try:
            workflow.assert_product_csv_current(csv_path)
            current_csv = True
        except RuntimeError as exc:
            csv_error = str(exc)
    result = {
        "ok": not issues and csv_path is not None,
        "project_root": str(workflow.PROJECT_ROOT),
        "python": str(workflow.python_exe()),
        "ocr_script": str(workflow.OCR_SCRIPT),
        "reference_data": str(workflow.REFERENCE_DIR),
        "synced_reference_files": synced,
        "product_csv": str(csv_path or ""),
        "product_csv_current": current_csv,
        "product_csv_error": csv_error,
        "product_match_ready": current_csv,
        "issues": issues,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--self-test", action="store_true", help="檢查本機 OCR、參考資料與產品資料狀態，不開啟視窗。")
    args = parser.parse_args()
    if args.self_test:
        return run_self_test()
    app = InvoiceOcrApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

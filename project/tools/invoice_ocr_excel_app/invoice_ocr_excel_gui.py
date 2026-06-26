from __future__ import annotations

import argparse
import json
import os
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import invoice_workflow as workflow
from invoice_workflow import AdjustmentRow, OcrConfirmRow, WorkflowState


APP_TITLE = "進貨單圖片轉 Excel"
SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}


class InvoiceOcrApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.state = WorkflowState()
        self.image_path = tk.StringVar(value="")
        self.output_dir = tk.StringVar(value=str(workflow.PROJECT_ROOT / "建檔進貨用" / "進貨圖片轉試算表"))
        self.status = tk.StringVar(value="請選擇一張進貨單圖片。")
        self.worker: threading.Thread | None = None
        self.started_at = 0.0
        self.current_output: Path | None = None
        self.edit_entry: ttk.Entry | None = None
        self.table_mode = "ocr"
        self.adjustment_source_rows: dict[str, int] = {}
        self.new_row_counter = 0

        self.title(APP_TITLE)
        self.geometry("1120x720")
        self.minsize(980, 620)
        self.configure(padx=14, pady=12)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.create_widgets()
        self.check_runtime_on_start()

    def create_widgets(self) -> None:
        header = ttk.Label(self, text=APP_TITLE, font=("Microsoft JhengHei UI", 18, "bold"))
        header.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        ttk.Label(self, text="圖片路徑").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.image_path).grid(row=1, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
        ttk.Button(self, text="選擇圖片", command=self.pick_image).grid(row=1, column=5, sticky="ew", pady=6)

        ttk.Label(self, text="輸出資料夾").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.output_dir).grid(row=2, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
        ttk.Button(self, text="選擇資料夾", command=self.pick_output_dir).grid(row=2, column=5, sticky="ew", pady=6)

        self.step_labels: dict[str, ttk.Label] = {}
        steps = [
            ("ocr", "1 OCR 原文確認"),
            ("match", "2 已建檔勾選"),
            ("review", "3 名稱調整"),
            ("build", "4 正式輸出"),
            ("cleanup", "5 清理中間檔"),
        ]
        step_frame = ttk.Frame(self)
        step_frame.grid(row=3, column=0, columnspan=6, sticky="ew", pady=(8, 6))
        for idx, (key, text) in enumerate(steps):
            label = ttk.Label(step_frame, text=f"□ {text}", font=("Microsoft JhengHei UI", 10, "bold"))
            label.grid(row=0, column=idx, sticky="w", padx=(0, 18))
            self.step_labels[key] = label

        button_frame = ttk.Frame(self)
        button_frame.grid(row=4, column=0, columnspan=6, sticky="ew", pady=(4, 8))
        self.start_button = ttk.Button(button_frame, text="開始 OCR + 產品比對", command=self.start_ocr_match)
        self.start_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.review_button = ttk.Button(button_frame, text="確認文字並開始調整", command=self.prepare_review, state="disabled")
        self.review_button.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.add_item_button = ttk.Button(button_frame, text="拆分/新增品項", command=self.split_selected_item, state="disabled")
        self.add_item_button.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.adjust_button = ttk.Button(button_frame, text="確認名稱並產生調整檔", command=self.generate_adjusted, state="disabled")
        self.adjust_button.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        self.build_button = ttk.Button(button_frame, text="產生建檔/採購檔", command=self.build_outputs, state="disabled")
        self.build_button.grid(row=0, column=4, sticky="ew", padx=(0, 8))
        self.cleanup_button = ttk.Button(button_frame, text="清理中間檔", command=self.cleanup_files, state="disabled")
        self.cleanup_button.grid(row=0, column=5, sticky="ew", padx=(0, 8))
        self.open_button = ttk.Button(button_frame, text="開啟目前檔案", command=self.open_current_output, state="disabled")
        self.open_button.grid(row=0, column=6, sticky="ew")
        for col in range(7):
            button_frame.columnconfigure(col, weight=1)

        table_frame = ttk.LabelFrame(self, text="資料確認")
        table_frame.grid(row=5, column=0, columnspan=6, sticky="nsew", pady=(4, 8))
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
        self.tree.bind("<Double-1>", self.begin_edit_cell)

        status_frame = ttk.LabelFrame(self, text="狀態")
        status_frame.grid(row=6, column=0, columnspan=6, sticky="nsew")
        status_frame.rowconfigure(1, weight=1)
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status, font=("Microsoft JhengHei UI", 11, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )
        self.detail_text = tk.Text(status_frame, height=8, wrap="word", state="disabled")
        self.detail_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=2)
        self.rowconfigure(6, weight=1)

    def configure_tree_columns(self, columns: tuple[str, ...], widths: list[int], stretch_columns: set[str]) -> None:
        self.tree.configure(columns=columns)
        for column, width in zip(columns, widths):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, stretch=(column in stretch_columns))

    def set_ocr_table(self) -> None:
        self.table_mode = "ocr"
        columns = ("已建檔", "OCR品名", "數量", "進價", "金額", "已建檔代號", "已建檔品名", "比對狀態", "相似候選")
        self.configure_tree_columns(columns, [70, 300, 70, 80, 80, 100, 260, 95, 430], {"OCR品名", "已建檔品名", "相似候選"})
        self.add_item_button.configure(state="disabled")

    def set_adjustment_table(self) -> None:
        self.table_mode = "adjust"
        columns = ("產品代號", "產品名稱", "大類", "數量", "成本", "金額", "狀態")
        self.configure_tree_columns(columns, [95, 380, 180, 70, 80, 80, 110], {"產品名稱", "大類"})

    def check_runtime_on_start(self) -> None:
        issues = workflow.validate_runtime()
        if issues:
            self.status.set("環境需要檢查。")
            self.set_detail("\n".join(issues))
            return
        csv_path = workflow.product_csv()
        self.set_detail(f"產品資料：{csv_path}\n請確認每天使用前已更新 reference_data 內的產品資料輸出.CSV。")

    def pick_image(self) -> None:
        selected = filedialog.askopenfilename(
            title="選擇進貨單圖片",
            filetypes=[("圖片檔", "*.jpg *.jpeg *.png *.webp"), ("所有檔案", "*.*")],
        )
        if selected:
            self.image_path.set(selected)
            self.status.set("已選擇圖片，可以開始。")

    def pick_output_dir(self) -> None:
        selected = filedialog.askdirectory(title="選擇輸出資料夾", initialdir=self.output_dir.get())
        if selected:
            self.output_dir.set(selected)

    def validate_start_inputs(self) -> tuple[Path, Path] | None:
        image_text = self.image_path.get().strip('" ')
        output_text = self.output_dir.get().strip('" ')
        if not image_text:
            messagebox.showwarning(APP_TITLE, "請先選擇進貨單圖片。")
            return None
        image = Path(image_text)
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
        csv_path = workflow.product_csv()
        if csv_path is None:
            messagebox.showerror(APP_TITLE, "找不到產品資料輸出.CSV。")
            return None
        try:
            workflow.assert_product_csv_current(csv_path)
        except RuntimeError as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return None
        return image, Path(output_text)

    def start_ocr_match(self) -> None:
        selected = self.validate_start_inputs()
        if selected is None:
            return
        image, output_dir = selected
        output_dir.mkdir(parents=True, exist_ok=True)
        self.state = WorkflowState(image_path=image, output_dir=output_dir)
        self.clear_tree()
        self.set_ocr_table()
        self.reset_steps()
        self.disable_buttons()
        self.start_button.configure(state="disabled")
        self.started_at = time.monotonic()
        self.status.set("OCR 與產品比對中。")
        self.set_detail("正在辨識圖片並比對既有產品，完成後會先讓你確認原始文字。")
        self.run_in_thread(self.worker_ocr_match)

    def worker_ocr_match(self) -> None:
        workflow.run_ocr(self.state)
        workflow.run_match(self.state)
        self.after(0, self.ocr_match_finished)

    def ocr_match_finished(self) -> None:
        self.mark_step("ocr", True)
        self.mark_step("match", True)
        self.populate_ocr_tree(workflow.load_ocr_confirm_rows(self.state))
        self.current_output = self.state.match_xlsx
        self.status.set("請確認 OCR 原始文字、成本與已建檔勾選。")
        lines = [
            "可雙擊表格修改 OCR 品名、數量、進價、金額。",
            "若確認是已建檔商品，勾選「已建檔」，並確認已建檔代號與品名正確。",
            "勾選已建檔且有六位代號的商品會跳過名稱調整，直接進採購輸出。",
            "",
            f"廠商：{self.state.vendor}",
            f"商品筆數：{self.state.row_count}",
        ]
        if self.state.ocr_issues:
            lines.append("")
            lines.append("OCR 疑點：")
            lines.extend(f"- {issue}" for issue in self.state.ocr_issues[:12])
        self.set_detail("\n".join(lines))
        self.start_button.configure(state="normal")
        self.review_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def prepare_review(self) -> None:
        if not self.state.match_xlsx:
            messagebox.showwarning(APP_TITLE, "請先完成 OCR 與產品比對。")
            return
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
        self.set_adjustment_table()
        self.populate_adjustment_tree(rows)
        self.mark_step("review", False)
        self.current_output = self.state.match_xlsx
        if not rows:
            self.status.set("全部商品已確認為已建檔，正在建立正式輸出前的交接檔。")
            self.set_detail("沒有需要名稱調整的新品或類似品，會直接產生 [調整].xlsx 供正式輸出使用。")
            self.run_in_thread(self.worker_generate_adjusted)
            return
        self.status.set("請確認需要建檔或調整名稱的商品。")
        lines = [
            "可雙擊表格的「產品代號 / 產品名稱 / 大類 / 數量 / 成本 / 金額」直接修改。",
            "大類請使用「代號 中文名稱」，例如：38 公仔/吊卡/PVC。",
            "需要拆成多個品項時，選取來源列後按「拆分/新增品項」，再調整拆分後數量與金額。",
            "確認後按「確認名稱並產生調整檔」。",
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
        self.status.set("正在產生 [調整].xlsx。")
        self.set_detail("正在檢查產品代號、正式名稱與大類。")
        self.run_in_thread(self.worker_generate_adjusted)

    def worker_generate_adjusted(self) -> None:
        result = workflow.generate_adjusted_xlsx(self.state)
        self.after(0, self.adjusted_finished, result)

    def adjusted_finished(self, result: dict[str, object]) -> None:
        self.mark_step("review", True)
        self.current_output = self.state.adjusted_xlsx
        self.status.set("調整檔已產生，可以正式輸出。")
        self.set_detail(f"{result.get('message', '')}\n\n調整檔：{self.state.adjusted_xlsx}")
        self.start_button.configure(state="normal")
        self.review_button.configure(state="normal")
        self.adjust_button.configure(state="normal")
        self.add_item_button.configure(state="disabled")
        self.build_button.configure(state="normal")
        self.open_button.configure(state="normal")

    def build_outputs(self) -> None:
        if not self.state.adjusted_xlsx:
            messagebox.showwarning(APP_TITLE, "請先產生 [調整].xlsx。")
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
        ]
        self.set_detail("\n".join(lines))
        self.start_button.configure(state="normal")
        self.cleanup_button.configure(state="normal")
        self.open_button.configure(state="normal")
        if messagebox.askyesno(APP_TITLE, "正式成品已產生。要現在清理本次中間檔嗎？"):
            self.cleanup_files()

    def cleanup_files(self) -> None:
        try:
            deleted = workflow.cleanup_intermediate_files(self.state)
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))
            return
        self.mark_step("cleanup", True)
        self.status.set("中間檔清理完成。")
        self.set_detail("已清理：\n" + "\n".join(str(path) for path in deleted) if deleted else "沒有可清理的中間檔。")
        self.cleanup_button.configure(state="disabled")

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
        self.status.set("流程停止。")
        self.set_detail(message[-5000:])
        self.start_button.configure(state="normal")
        if self.state.match_xlsx and self.state.match_xlsx.exists():
            self.review_button.configure(state="normal")
        if self.state.adjusted_xlsx and self.state.adjusted_xlsx.exists():
            self.build_button.configure(state="normal")
        self.open_button.configure(state="normal" if self.current_output and self.current_output.exists() else "disabled")

    def update_elapsed_status(self) -> None:
        if not self.worker or not self.worker.is_alive():
            return
        elapsed = int(time.monotonic() - self.started_at)
        self.status.set(f"處理中，已執行 {elapsed} 秒。")
        self.after(1000, self.update_elapsed_status)

    def populate_ocr_tree(self, rows: list[OcrConfirmRow]) -> None:
        self.set_ocr_table()
        self.clear_tree()
        for row in rows:
            self.tree.insert(
                "",
                "end",
                iid=str(row.excel_row),
                values=(
                    "☑" if row.is_existing else "☐",
                    row.raw_name,
                    row.quantity,
                    row.unit_cost,
                    row.amount,
                    row.matched_code,
                    row.matched_name,
                    row.status,
                    row.candidates,
                ),
            )

    def populate_adjustment_tree(self, rows: list[AdjustmentRow]) -> None:
        self.clear_tree()
        self.adjustment_source_rows = {}
        for row in rows:
            self.adjustment_source_rows[row.row_id] = row.source_row
            self.tree.insert(
                "",
                "end",
                iid=row.row_id,
                values=(
                    row.product_code,
                    row.name,
                    row.category_display,
                    row.quantity,
                    row.unit_cost,
                    row.amount,
                    row.status,
                ),
            )

    def clear_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)

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
                    candidates=str(values[8]).strip(),
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
                    category_display=str(values[2]).strip(),
                    quantity=str(values[3]).strip(),
                    unit_cost=str(values[4]).strip(),
                    amount=str(values[5]).strip(),
                    status=str(values[6]).strip(),
                )
            )
        return rows

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

    def begin_edit_cell(self, event) -> None:
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item = self.tree.identify_row(event.y)
        column_id = self.tree.identify_column(event.x)
        if not item:
            return
        column_index = int(column_id.replace("#", "")) - 1
        if self.table_mode == "ocr":
            if column_id == "#1":
                values = list(self.tree.item(item, "values"))
                if str(values[0]).strip() == "☑":
                    values[0] = "☐"
                else:
                    values[0] = "☑"
                    if not str(values[5]).strip():
                        code, name = workflow.first_candidate(str(values[8]))
                        values[5] = code
                        values[6] = name
                self.tree.item(item, values=values)
                return
            editable_columns = {"#2", "#3", "#4", "#5", "#6", "#7"}
        else:
            editable_columns = {"#1", "#2", "#3", "#4", "#5", "#6"}
        if column_id not in editable_columns:
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

        def commit(_event=None) -> None:
            if self.edit_entry is None:
                return
            new_value = self.edit_entry.get()
            if self.table_mode == "adjust" and column_id == "#3":
                new_value = workflow.category_display(new_value)
            values[column_index] = new_value
            self.tree.item(item, values=values)
            self.edit_entry.destroy()
            self.edit_entry = None

        self.edit_entry.bind("<Return>", commit)
        self.edit_entry.bind("<FocusOut>", commit)

    def disable_buttons(self, keep_start: bool = False) -> None:
        for button in [self.review_button, self.add_item_button, self.adjust_button, self.build_button, self.cleanup_button, self.open_button]:
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
        "ok": not issues and current_csv,
        "project_root": str(workflow.PROJECT_ROOT),
        "reference_data": str(workflow.REFERENCE_DIR),
        "product_csv": str(csv_path or ""),
        "product_csv_current": current_csv,
        "product_csv_error": csv_error,
        "issues": issues,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--self-test", action="store_true", help="檢查本機 OCR、參考資料與產品資料日期，不開啟視窗。")
    args = parser.parse_args()
    if args.self_test:
        return run_self_test()
    app = InvoiceOcrApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

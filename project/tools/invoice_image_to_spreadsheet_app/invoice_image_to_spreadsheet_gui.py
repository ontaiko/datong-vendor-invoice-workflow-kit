from __future__ import annotations

import argparse
import os
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import ocr_runner
from ocr_runner import OcrRunResult


APP_TITLE = "進貨單 OCR 批次轉試算表"


class BatchInvoiceOcrApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.items: dict[str, dict[str, object]] = {}
        self.worker: threading.Thread | None = None
        self.started_at = 0.0
        self.status = tk.StringVar(value="請選擇一張或多張進貨單圖片。")

        self.title(APP_TITLE)
        self.geometry("1120x680")
        self.minsize(980, 560)
        self.configure(padx=14, pady=12)
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.create_widgets()
        self.check_runtime_on_start()

    def create_widgets(self) -> None:
        header = ttk.Label(self, text=APP_TITLE, font=("Microsoft JhengHei UI", 18, "bold"))
        header.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        button_frame = ttk.Frame(self)
        button_frame.grid(row=1, column=0, columnspan=6, sticky="ew", pady=(0, 8))
        self.pick_button = ttk.Button(button_frame, text="選擇圖片", command=self.pick_images)
        self.pick_button.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.start_button = ttk.Button(button_frame, text="開始批次 OCR", command=self.start_batch, state="disabled")
        self.start_button.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.open_excel_button = ttk.Button(button_frame, text="開啟選取 Excel", command=self.open_selected_excel, state="disabled")
        self.open_excel_button.grid(row=0, column=2, sticky="ew", padx=(0, 8))
        self.open_folder_button = ttk.Button(button_frame, text="開啟圖片資料夾", command=self.open_selected_folder, state="disabled")
        self.open_folder_button.grid(row=0, column=3, sticky="ew", padx=(0, 8))
        self.clear_button = ttk.Button(button_frame, text="清除清單", command=self.clear_items, state="disabled")
        self.clear_button.grid(row=0, column=4, sticky="ew")
        for col in range(5):
            button_frame.columnconfigure(col, weight=1)

        table_frame = ttk.LabelFrame(self, text="批次清單")
        table_frame.grid(row=2, column=0, columnspan=6, sticky="nsew", pady=(2, 8))
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        columns = ("狀態", "圖片檔名", "輸出 Excel", "商品筆數", "廠商", "疑點或錯誤")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=16)
        widths = [90, 260, 300, 80, 130, 430]
        stretch_columns = {"圖片檔名", "輸出 Excel", "疑點或錯誤"}
        for column, width in zip(columns, widths):
            self.tree.heading(column, text=column)
            self.tree.column(column, width=width, stretch=(column in stretch_columns))

        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        self.tree.bind("<<TreeviewSelect>>", lambda _event: self.update_selection_buttons())

        status_frame = ttk.LabelFrame(self, text="狀態")
        status_frame.grid(row=3, column=0, columnspan=6, sticky="nsew")
        status_frame.rowconfigure(1, weight=1)
        status_frame.columnconfigure(0, weight=1)
        ttk.Label(status_frame, textvariable=self.status, font=("Microsoft JhengHei UI", 11, "bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )
        self.detail_text = tk.Text(status_frame, height=8, wrap="word", state="disabled")
        self.detail_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=3)
        self.rowconfigure(3, weight=1)

    def check_runtime_on_start(self) -> None:
        issues = ocr_runner.validate_runtime()
        if issues:
            self.status.set("OCR 環境需要檢查。")
            self.set_detail("\n\n".join(issues))
            self.start_button.configure(state="disabled")
            return
        self.set_detail(
            "已確認 OCR 環境可用。\n"
            "請選擇一張或多張圖片，輸出的 Excel 會放在各自圖片所在資料夾。"
        )

    def pick_images(self) -> None:
        selected = filedialog.askopenfilenames(
            title="選擇進貨單圖片",
            filetypes=[("圖片檔", "*.jpg *.jpeg *.png *.webp"), ("所有檔案", "*.*")],
        )
        if not selected:
            return
        added = 0
        for file_name in selected:
            path = Path(file_name).resolve()
            key = str(path)
            if key in self.items:
                continue
            if path.suffix.lower() not in ocr_runner.SUPPORTED_EXTENSIONS:
                continue
            item_id = key
            self.items[item_id] = {"image": path, "output": None, "status": "等待中"}
            self.tree.insert(
                "",
                "end",
                iid=item_id,
                values=("等待中", path.name, "", "", "", ""),
            )
            added += 1
        self.status.set(f"已加入 {added} 張圖片，清單共 {len(self.items)} 張。")
        self.set_detail("按「開始批次 OCR」後會依序處理，單張失敗會繼續下一張。")
        self.update_main_buttons()

    def start_batch(self) -> None:
        pending = [item_id for item_id, data in self.items.items() if data.get("status") in {"等待中", "失敗"}]
        if not pending:
            messagebox.showinfo(APP_TITLE, "沒有等待處理的圖片。")
            return
        self.disable_during_work()
        self.started_at = time.monotonic()
        self.status.set("批次 OCR 處理中。")
        self.set_detail("正在逐張辨識圖片，請先不要關閉視窗。")
        self.worker = threading.Thread(target=self.worker_batch, args=(pending,), daemon=True)
        self.worker.start()
        self.after(1000, self.update_elapsed_status)

    def worker_batch(self, item_ids: list[str]) -> None:
        total = len(item_ids)
        for index, item_id in enumerate(item_ids, start=1):
            image = self.items[item_id]["image"]
            if not isinstance(image, Path):
                continue
            self.after(0, self.mark_processing, item_id, index, total)
            result = ocr_runner.run_ocr(image)
            self.after(0, self.apply_result, item_id, result, index, total)
        self.after(0, self.batch_finished)

    def mark_processing(self, item_id: str, index: int, total: int) -> None:
        image = self.items[item_id]["image"]
        if not isinstance(image, Path):
            return
        self.items[item_id]["status"] = "處理中"
        self.tree.item(item_id, values=("處理中", image.name, "", "", "", ""))
        self.tree.selection_set(item_id)
        self.tree.see(item_id)
        self.status.set(f"處理中：{index} / {total} - {image.name}")

    def apply_result(self, item_id: str, result: OcrRunResult, index: int, total: int) -> None:
        image = result.image
        output_text = str(result.output) if result.output else ""
        issue_text = result.error or "；".join(result.issues[:8])
        self.items[item_id].update(
            {
                "status": result.status,
                "output": result.output,
                "vendor": result.vendor,
                "row_count": result.row_count,
                "issues": result.issues,
                "error": result.error,
            }
        )
        self.tree.item(
            item_id,
            values=(result.status, image.name, output_text, result.row_count or "", result.vendor, issue_text),
        )
        self.status.set(f"已完成：{index} / {total} - {image.name}")

    def batch_finished(self) -> None:
        counts = self.count_statuses()
        self.status.set(
            f"批次完成：完成 {counts['完成']} 張，需覆核 {counts['需覆核']} 張，失敗 {counts['失敗']} 張。"
        )
        lines = [
            "批次處理完成。",
            f"完成：{counts['完成']} 張",
            f"需覆核：{counts['需覆核']} 張",
            f"失敗：{counts['失敗']} 張",
        ]
        if counts["需覆核"]:
            lines.append("")
            lines.append("需覆核的 Excel 已產生，請打開檔案內的 OCR測試紀錄工作表查看疑點。")
        if counts["失敗"]:
            lines.append("")
            lines.append("失敗圖片已保留在清單中，可修正後再按開始批次 OCR 重跑失敗項目。")
        self.set_detail("\n".join(lines))
        self.worker = None
        self.update_main_buttons()
        self.update_selection_buttons()

    def count_statuses(self) -> dict[str, int]:
        result = {"完成": 0, "需覆核": 0, "失敗": 0, "等待中": 0, "處理中": 0}
        for data in self.items.values():
            status = str(data.get("status", ""))
            if status in result:
                result[status] += 1
        return result

    def update_elapsed_status(self) -> None:
        if not self.worker or not self.worker.is_alive():
            return
        elapsed = int(time.monotonic() - self.started_at)
        self.status.set(f"批次 OCR 處理中，已執行 {elapsed} 秒。")
        self.after(1000, self.update_elapsed_status)

    def clear_items(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        self.items.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.status.set("清單已清除。")
        self.set_detail("請選擇一張或多張進貨單圖片。")
        self.update_main_buttons()
        self.update_selection_buttons()

    def selected_item(self) -> str | None:
        selected = self.tree.selection()
        return selected[0] if selected else None

    def open_selected_excel(self) -> None:
        item_id = self.selected_item()
        if not item_id:
            return
        output = self.items[item_id].get("output")
        if isinstance(output, Path) and output.exists():
            os.startfile(output)

    def open_selected_folder(self) -> None:
        item_id = self.selected_item()
        if not item_id:
            return
        image = self.items[item_id].get("image")
        if isinstance(image, Path) and image.parent.exists():
            os.startfile(image.parent)

    def update_main_buttons(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        has_items = bool(self.items)
        has_pending = any(data.get("status") in {"等待中", "失敗"} for data in self.items.values())
        self.pick_button.configure(state="normal")
        self.start_button.configure(state="normal" if has_pending else "disabled")
        self.clear_button.configure(state="normal" if has_items else "disabled")

    def update_selection_buttons(self) -> None:
        if self.worker and self.worker.is_alive():
            self.open_excel_button.configure(state="disabled")
            self.open_folder_button.configure(state="disabled")
            return
        item_id = self.selected_item()
        if not item_id:
            self.open_excel_button.configure(state="disabled")
            self.open_folder_button.configure(state="disabled")
            return
        output = self.items[item_id].get("output")
        image = self.items[item_id].get("image")
        self.open_excel_button.configure(state="normal" if isinstance(output, Path) and output.exists() else "disabled")
        self.open_folder_button.configure(state="normal" if isinstance(image, Path) and image.parent.exists() else "disabled")

    def disable_during_work(self) -> None:
        for button in [self.pick_button, self.start_button, self.open_excel_button, self.open_folder_button, self.clear_button]:
            button.configure(state="disabled")

    def set_detail(self, text: str) -> None:
        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("1.0", text)
        self.detail_text.configure(state="disabled")

    def on_close(self) -> None:
        if self.worker and self.worker.is_alive():
            if not messagebox.askyesno(APP_TITLE, "批次 OCR 還在執行中，確定要關閉嗎？"):
                return
        self.destroy()


def run_self_test() -> int:
    return ocr_runner.run_self_test()


def main() -> int:
    parser = argparse.ArgumentParser(description=APP_TITLE)
    parser.add_argument("--self-test", action="store_true", help="檢查 OCR 環境，不開啟視窗。")
    args = parser.parse_args()
    if args.self_test:
        return run_self_test()
    app = BatchInvoiceOcrApp()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

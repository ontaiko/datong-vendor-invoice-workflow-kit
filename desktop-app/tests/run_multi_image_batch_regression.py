from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import invoice_workflow as workflow


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--first-raw", type=Path, required=True)
    parser.add_argument("--second-raw", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()

    first_raw = args.first_raw.resolve()
    second_raw = args.second_raw.resolve()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    if not first_raw.exists() or not second_raw.exists():
        raise RuntimeError("找不到批次合併回歸測試來源。")

    source_rows = {
        "first.jpg": workflow.load_ocr_confirm_rows(
            workflow.WorkflowState(raw_xlsx=first_raw)
        ),
        "second.jpg": workflow.load_ocr_confirm_rows(
            workflow.WorkflowState(raw_xlsx=second_raw)
        ),
    }
    source_files = {
        "first.jpg": first_raw,
        "second.jpg": second_raw,
    }
    expected_rows = sum(len(rows) for rows in source_rows.values())
    expected_total = sum(
        workflow.decimal_from_value(workflow.read_invoice_total(path))
        for path in source_files.values()
    )

    original_run_ocr = workflow.run_ocr
    vendor_by_image = {
        "first.jpg": "麗嬰國際股份有限公司",
        "second.jpg": "麗嬰國際股份有限公司",
    }

    def fake_run_ocr(
        state: workflow.WorkflowState,
        _project_root: Path = workflow.PROJECT_ROOT,
    ) -> dict[str, object]:
        if state.image_path is None:
            raise RuntimeError("測試圖片遺失。")
        key = state.image_path.name
        raw_path = source_files[key]
        state.raw_xlsx = raw_path
        state.raw_xlsx_files = [raw_path]
        state.vendor = vendor_by_image[key]
        state.row_count = len(source_rows[key])
        state.invoice_total = workflow.read_invoice_total(raw_path)
        state.needs_ocr_review = False
        state.ocr_issues = []
        return {
            "output": str(raw_path),
            "vendor": state.vendor,
            "row_count": state.row_count,
        }

    workflow.run_ocr = fake_run_ocr
    try:
        state = workflow.WorkflowState(
            image_path=Path("first.jpg"),
            image_paths=[Path("first.jpg"), Path("second.jpg")],
            output_dir=output_dir,
        )
        progress: list[tuple[int, int, str]] = []
        summary = workflow.run_ocr_batch(
            state,
            progress_callback=lambda current, total, image: progress.append(
                (current, total, image.name)
            ),
        )
        merged_rows = workflow.load_ocr_confirm_rows(state)
        if len(merged_rows) != expected_rows:
            raise RuntimeError(
                f"同廠商合併筆數錯誤：{len(merged_rows)}，預期 {expected_rows}"
            )
        if workflow.decimal_from_value(state.invoice_total) != expected_total:
            raise RuntimeError(
                f"同廠商合併總額錯誤：{state.invoice_total}，預期 {expected_total}"
            )
        if progress != [(1, 2, "first.jpg"), (2, 2, "second.jpg")]:
            raise RuntimeError(f"多圖 OCR 進度錯誤：{progress}")
        if int(summary.get("image_count", 0)) != 2:
            raise RuntimeError(f"多圖 OCR 摘要錯誤：{summary}")

        vendor_by_image["second.jpg"] = "不同廠商"
        mismatch_rejected = False
        try:
            workflow.run_ocr_batch(
                workflow.WorkflowState(
                    image_path=Path("first.jpg"),
                    image_paths=[Path("first.jpg"), Path("second.jpg")],
                    output_dir=output_dir,
                )
            )
        except RuntimeError as exc:
            mismatch_rejected = "不同廠商" in str(exc) and "不可合併" in str(exc)
        if not mismatch_rejected:
            raise RuntimeError("不同廠商圖片沒有被拒絕合併。")
    finally:
        workflow.run_ocr = original_run_ocr

    print(
        json.dumps(
            {
                "ok": True,
                "image_count": 2,
                "merged_rows": expected_rows,
                "merged_total": workflow.decimal_text(expected_total),
                "same_vendor_merged": True,
                "different_vendor_rejected": True,
                "ocr_progress_callback": True,
                "output": str(state.raw_xlsx),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

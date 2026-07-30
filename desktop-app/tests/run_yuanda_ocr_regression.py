from __future__ import annotations

import argparse
import json
import sys
from decimal import Decimal
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import invoice_workflow as workflow


EXPECTED_DOCUMENTS = [
    [
        ("(再販26)寶可夢 Pokepeace角色造型吊飾", "40", "54", "2160"),
        ("西村優志角色大頭小物包鉤環開心熊貓", "9", "260", "2340"),
        ("西村優志角色大頭小物包鉤環愛心兔", "6", "260", "1560"),
        ("西村優志角色大頭小物包鉤環壞壞貓", "3", "260", "780"),
        ("(再販26)夢奇奇造型公仔", "36", "26", "936"),
        ("葬送的芙莉蓮JIGSAW拼圖ATB-82", "3", "580", "1740"),
    ],
    [
        ("(再販26)蠟筆小新大滿足公仔-零食派對篇", "12", "170", "2040"),
        ("(再販25)星之卡比的噗噗噗咖啡時光", "24", "144.5", "3468"),
        ("PalVerse 庫洛魔法使透明牌篇盲盒", "12", "200", "2400"),
    ],
    [
        ("(再販26)寶可夢 Little Night收藏P2", "12", "212.5", "2550"),
        ("寶可夢鑽石星塵", "12", "231", "2772"),
        ("(再販26)寶可夢華麗裝飾框收藏P2", "12", "214.5", "2574"),
        ("PalVerse名偵探柯南 vol.2 盲盒(附特典)", "12", "200", "2400"),
        ("PalVerse Hololive vol.1盲盒", "12", "200", "2400"),
    ],
]
EXPECTED_TOTALS = [Decimal("9516"), Decimal("7908"), Decimal("12696")]


def row_signature(row: workflow.OcrConfirmRow) -> tuple[str, str, str, str]:
    return (
        row.raw_name,
        workflow.decimal_text(workflow.decimal_from_value(row.quantity)),
        workflow.decimal_text(workflow.decimal_from_value(row.unit_cost)),
        workflow.decimal_text(workflow.decimal_from_value(row.amount)),
    )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--raw", type=Path, action="append", required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()

    raw_files = [path.resolve() for path in args.raw]
    if len(raw_files) != 3:
        raise RuntimeError("元大 OCR 回歸測試必須提供三個 --raw 檔案。")
    if any(not path.exists() for path in raw_files):
        raise RuntimeError("找不到元大 OCR 回歸測試來源。")

    states: list[workflow.WorkflowState] = []
    document_results: list[dict[str, object]] = []
    for index, (raw_file, expected_rows, expected_total) in enumerate(
        zip(raw_files, EXPECTED_DOCUMENTS, EXPECTED_TOTALS),
        start=1,
    ):
        state = workflow.WorkflowState(
            image_path=Path(f"S__6397514{index + 3}.jpg"),
            raw_xlsx=raw_file,
            raw_xlsx_files=[raw_file],
            vendor="元大玩具股份有限公司",
            invoice_total=workflow.read_invoice_total(raw_file),
        )
        rows = workflow.load_ocr_confirm_rows(state)
        actual_rows = [row_signature(row) for row in rows]
        if actual_rows != expected_rows:
            raise RuntimeError(
                f"第 {index} 張元大 OCR 列內容錯誤：\n"
                + json.dumps(actual_rows, ensure_ascii=False, indent=2)
            )
        actual_total = workflow.decimal_from_value(state.invoice_total)
        if actual_total != expected_total:
            raise RuntimeError(
                f"第 {index} 張元大 OCR 總額 {actual_total}，預期 {expected_total}"
            )
        state.row_count = len(rows)
        states.append(state)
        document_results.append(
            {
                "rows": len(rows),
                "quantity": sum(
                    workflow.decimal_from_value(row.quantity) for row in rows
                ),
                "amount": actual_total,
            }
        )

    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    merged_path = workflow.merge_ocr_workbooks(states, output_dir)
    merged_state = workflow.WorkflowState(raw_xlsx=merged_path)
    merged_rows = workflow.load_ocr_confirm_rows(merged_state)
    merged_total = workflow.decimal_from_value(
        workflow.read_invoice_total(merged_path)
    )
    if len(merged_rows) != 14 or merged_total != Decimal("30120"):
        raise RuntimeError(
            f"元大合併結果錯誤：{len(merged_rows)} 列、總額 {merged_total}"
        )

    print(
        json.dumps(
            {
                "ok": True,
                "documents": [
                    {
                        "rows": result["rows"],
                        "quantity": workflow.decimal_text(result["quantity"]),
                        "amount": workflow.decimal_text(result["amount"]),
                    }
                    for result in document_results
                ],
                "merged_rows": len(merged_rows),
                "merged_quantity": workflow.decimal_text(
                    sum(
                        workflow.decimal_from_value(row.quantity)
                        for row in merged_rows
                    )
                ),
                "merged_total": workflow.decimal_text(merged_total),
                "output": str(merged_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

from __future__ import annotations

import csv
import json
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import invoice_workflow as workflow


def main() -> int:
    output_dir = (
        workflow.PROJECT_ROOT
        / ".codex-tmp"
        / "invoice-app-tests"
        / "lower-candidate-search"
    )
    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / "產品資料輸出.CSV"
    with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["1.產品代號", "2.產品名稱"],
        )
        writer.writeheader()
        writer.writerows(
            [
                {"1.產品代號": "000101", "2.產品名稱": "寶可夢護手霜"},
                {"1.產品代號": "000102", "2.產品名稱": "寶可夢吊飾"},
                {"1.產品代號": "000103", "2.產品名稱": "完全無關測試商品"},
            ]
        )
    identity_path = output_dir / "產品比對身份關鍵詞.csv"
    with identity_path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["類型", "關鍵詞", "啟用", "備註"],
        )
        writer.writeheader()

    original_override = workflow._PRODUCT_CSV_OVERRIDE
    workflow._PRODUCT_CSV_OVERRIDE = csv_path
    try:
        candidates = workflow.find_lower_similarity_candidates(
            "寶可夢吊飾",
            minimum_score=0.35,
            maximum_score=0.55,
            max_candidates=5,
        )
    finally:
        workflow._PRODUCT_CSV_OVERRIDE = original_override

    if not candidates:
        raise RuntimeError("更低候選搜尋沒有回傳測試候選。")
    if candidates[0][0] != "000101" or candidates[0][3] != "deep":
        raise RuntimeError(f"更低候選搜尋結果錯誤：{candidates}")
    if any(code == "000102" for code, _name, _score, _tier in candidates):
        raise RuntimeError("一般／精確候選不應混入更低候選區間。")

    pasted = workflow.parse_pasted_column_values("第一列\r\n\r\n第三列\r\n")
    if pasted != ["第一列", "", "第三列"]:
        raise RuntimeError(f"整欄貼上沒有保留中間空白列：{pasted}")

    print(
        json.dumps(
            {
                "ok": True,
                "candidate_count": len(candidates),
                "first_candidate": candidates[0],
                "manual_only_tier": "deep",
                "blank_row_preserved": True,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

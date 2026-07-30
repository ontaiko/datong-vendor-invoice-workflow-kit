from __future__ import annotations

import importlib.util
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw


APP_DIR = Path(__file__).resolve().parents[1]
SCRIPT_PATH = APP_DIR / "scripts" / "local_paddleocr_invoice_to_xlsx.py"
SPEC = importlib.util.spec_from_file_location("invoice_ocr_engine", SCRIPT_PATH)
if SPEC is None or SPEC.loader is None:
    raise RuntimeError(f"無法載入 OCR 模組：{SCRIPT_PATH}")
ocr = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = ocr
SPEC.loader.exec_module(ocr)

WORKFLOW_SPEC = importlib.util.spec_from_file_location("invoice_workflow", APP_DIR / "invoice_workflow.py")
if WORKFLOW_SPEC is None or WORKFLOW_SPEC.loader is None:
    raise RuntimeError("無法載入 invoice_workflow")
workflow = importlib.util.module_from_spec(WORKFLOW_SPEC)
sys.modules[WORKFLOW_SPEC.name] = workflow
WORKFLOW_SPEC.loader.exec_module(workflow)


def entry(text: str, x: float, y: float, score: float = 0.99, width: int = 100, height: int = 20):
    return ocr.OcrEntry(
        text=text,
        score=score,
        box=[int(x - width / 2), int(y - height / 2), int(x + width / 2), int(y + height / 2)],
        x=x,
        y=y,
    )


def assert_page_split() -> None:
    def page_image(page_count: int) -> Image.Image:
        image = Image.new("RGB", (900, 1800), "white")
        draw = ImageDraw.Draw(image)
        boundaries = [int(1800 * index / page_count) for index in range(page_count + 1)]
        for index in range(page_count):
            top = boundaries[index] + 35
            bottom = boundaries[index + 1] - 35
            for y in range(top, bottom, 28):
                draw.line((80, y, 820, y), fill="black", width=3)
            draw.rectangle((70, top, 830, bottom), outline="black", width=3)
        return image

    if len(ocr.split_stacked_pages(page_image(2))) != 2:
        raise RuntimeError("雙頁分割回歸失敗")
    if len(ocr.split_stacked_pages(page_image(3))) != 3:
        raise RuntimeError("三頁分割回歸失敗")


def assert_wanrong_boundaries() -> None:
    entries = [
        entry("貨號", 150, 100),
        entry("品名規格", 420, 100),
        entry("數量", 650, 100),
        entry("單位", 720, 100),
        entry("售價", 800, 100),
        entry("折數", 880, 100),
        entry("進價", 960, 100),
        entry("金額", 1100, 100),
        entry("BAP-111111 第一商品主名稱", 300, 200, width=300),
        entry("第一商品續行一般版", 420, 230),
        entry("1", 650, 202),
        entry("PCS", 720, 202),
        entry("350.000", 960, 202),
        entry("350", 1100, 202),
        entry("BAP-222222", 150, 260),
        entry("第二商品主名稱", 420, 260),
        entry("第二商品續行豪華版", 420, 290),
        entry("2", 650, 262),
        entry("PCS", 720, 262),
        entry("400.000", 960, 262),
        entry("800", 1100, 262),
        entry("總計", 960, 360),
        entry("1150", 1100, 360),
        entry("萬榮國際企業股份有限公司", 850, 40),
    ]
    rows, issues = ocr.parse_wanrong_rows(entries)
    if issues or len(rows) != 2:
        raise RuntimeError(f"萬榮列界回歸失敗：{issues} / {rows}")
    if rows[0].name != "第一商品主名稱第一商品續行一般版":
        raise RuntimeError(f"第一筆續行跨列：{rows[0].name}")
    if rows[1].name != "第二商品主名稱第二商品續行豪華版":
        raise RuntimeError(f"第二筆續行跨列：{rows[1].name}")
    if (rows[0].quantity, rows[0].unit_cost, rows[0].amount) != (1, 350, 350):
        raise RuntimeError(f"第一筆數字欄錯配：{rows[0]}")
    if (rows[1].quantity, rows[1].unit_cost, rows[1].amount) != (2, 400, 800):
        raise RuntimeError(f"第二筆數字欄錯配：{rows[1]}")

    mismatched = list(entries)
    amount_index = next(
        index for index, item in enumerate(mismatched) if item.text == "800" and item.y == 262
    )
    mismatched[amount_index] = entry("900", 1100, 262)
    bad_rows, _ = ocr.parse_wanrong_rows(mismatched)
    if bad_rows[1].unit_cost != 400 or bad_rows[1].check != "不符":
        raise RuntimeError("萬榮不符列不應用金額÷數量覆寫進價")


def assert_total_and_page_audits() -> None:
    first_page = [
        entry("合計", 900, 900),
        entry("4098", 1100, 900),
        entry("頁次：1/3", 1050, 80),
    ]
    last_page = [
        entry("折數", 880, 300),
        entry("40", 900, 300),
        entry("單列金額", 960, 500),
        entry("4098", 1100, 500),
        entry("總計", 960, 900),
        entry("46700", 1100, 900),
        entry("頁次：3/3", 1050, 80),
    ]
    result = {
        "entries": ocr.merge_page_entries([first_page, last_page]),
        "page_entries": [first_page, last_page],
    }
    if ocr.infer_result_total(result, []) != 46700:
        raise RuntimeError("總額未限定最末頁總計右側")
    page_issues = ocr.page_number_audit_issues(result)
    if not any("缺少頁碼 [2]" in issue for issue in page_issues):
        raise RuntimeError(f"缺頁稽核未觸發：{page_issues}")


def assert_name_safety() -> None:
    for header in ("品名規格", "品名规格", "商品名稱", "商品名称"):
        cleaned = ocr.normalize_ocr_name(
            f"{header} *吉伊卡哇 櫻花吊偶2(全4種) ********以下空白"
        )
        if cleaned != "吉伊卡哇櫻花吊偶2(全4種)":
            raise RuntimeError(f"分隔線／表頭清理失敗（{header}）：{cleaned}")
    issues = ocr.suspicious_name_issues("BNFIGURBQ PAJAMARS明日方舟VOL.1盲盒")
    if not any("黏字" in issue for issue in issues):
        raise RuntimeError("英文黏字未標記人工確認")
    row = ocr.ProductRow("BAP-1", "商品", None, 400, None, "需確認", "數量未穩定辨識")
    combined = ocr.collect_review_issues(["整體疑點"], [row])
    if combined != ["整體疑點", "BAP-1: 數量未穩定辨識"]:
        raise RuntimeError(f"列級 issues 未同步：{combined}")


def assert_yuanda_safe_amount_fill() -> None:
    rows = [
        ocr.ProductRow(str(index), f"商品{index}", 10, 235, None, "需確認", "金額未穩定辨識")
        for index in range(1, 4)
    ]
    entries = [entry("原幣合計：7,050", 800, 900)]
    corrected = ocr.fill_yuanda_missing_amounts_when_total_matches(rows, entries)
    if [row.amount for row in corrected] != [2350, 2350, 2350]:
        raise RuntimeError(f"元大安全補值失敗：{corrected}")
    if any(row.check != "通過" or not row.audit_note for row in corrected):
        raise RuntimeError("元大安全補值缺少通過狀態或稽核紀錄")

    gift_rows = [
        ocr.ProductRow("1", "贈品", 1, 100, None, "需確認", "金額未穩定辨識")
    ]
    if ocr.fill_yuanda_missing_amounts_when_total_matches(gift_rows, [entry("原幣合計：100", 800, 900)])[0].amount:
        raise RuntimeError("贈品列不得自動補值")


def assert_tax_summary_and_workflow_quality() -> None:
    entries = [
        entry("合計金额", 600, 340),
        entry("20,170", 720, 340),
        entry("税额", 600, 370),
        entry("1,009", 720, 370),
        entry("總金额", 600, 400),
        entry("21,179", 720, 400),
    ]
    summary = ocr.infer_tax_summary(entries)
    if summary != {"未稅合計": 20170, "稅額": 1009, "含稅總額": 21179}:
        raise RuntimeError(f"南波稅額摘要錯誤：{summary}")
    split_summary = {
        "row_count": 13,
        "estimated_row_count": 13,
        "complete_row_count": 11,
        "valid_row_count": 11,
        "total_matches": False,
        "page_error_count": 1,
        "page_issue_count": 1,
        "page_count": 2,
    }
    single_summary = {
        "row_count": 3,
        "estimated_row_count": 9,
        "complete_row_count": 1,
        "valid_row_count": 1,
        "total_matches": False,
        "page_error_count": 0,
        "page_issue_count": 1,
        "page_count": 1,
    }
    if workflow.ocr_summary_quality(split_summary) <= workflow.ocr_summary_quality(single_summary):
        raise RuntimeError("真正多頁結果不得因單頁失敗而無條件回退")


def main() -> int:
    assert_page_split()
    assert_wanrong_boundaries()
    assert_total_and_page_audits()
    assert_name_safety()
    assert_yuanda_safe_amount_fill()
    assert_tax_summary_and_workflow_quality()
    print("OCR safety regression: PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

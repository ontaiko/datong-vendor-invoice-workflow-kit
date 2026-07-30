from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
SCRIPT_PATH = APP_DIR / "scripts" / "local_paddleocr_invoice_to_xlsx.py"


def load_ocr_module():
    spec = importlib.util.spec_from_file_location("invoice_ocr_liying_regression", SCRIPT_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"無法載入 OCR 腳本：{SCRIPT_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def make_entry(module, text: str, x1: int, y1: int, x2: int, y2: int):
    return module.OcrEntry(
        text=text,
        score=0.99,
        box=[x1, y1, x2, y2],
        x=(x1 + x2) / 2,
        y=(y1 + y2) / 2,
    )


def main() -> int:
    module = load_ocr_module()
    entries = [
        make_entry(module, "產品名稱", 343, 418, 414, 440),
        make_entry(module, "零售價", 477, 419, 532, 442),
        make_entry(module, "數量", 551, 420, 591, 443),
        make_entry(module, "單價", 623, 421, 660, 442),
        make_entry(module, "金額", 702, 421, 742, 445),
    ]
    source_rows = [
        ("CU45031", (87, 450, 147, 468), ("線條小狗絨毛吊飾-慶", 292, 445, 448, 465), ("典小白", 292, 458, 344, 479), 350, 5, 245, 1225),
        ("CU45032", (84, 479, 146, 496), ("線條小狗絨毛吊飾-慶", 291, 473, 448, 496), ("典小金毛", 290, 488, 360, 508), 350, 5, 245, 1225),
        ("CU45041", (80, 504, 140, 522), ("線條小狗絨毛吊飾-晚", 290, 501, 447, 525), ("安星星小金毛", 289, 516, 392, 538), 280, 5, 196, 980),
        ("CU45042", (75, 532, 138, 550), ("線條小狗絨毛吊飾-晚", 288, 530, 447, 554), ("安月亮小白", 287, 545, 373, 567), 280, 5, 196, 980),
        ("CU45043", (69, 557, 134, 581), ("線條小狗絨毛吊飾-星", 285, 559, 447, 585), ("願小白", 285, 574, 339, 595), 280, 10, 196, 1960),
        ("CU45044", (66, 586, 131, 608), ("線條小狗絨毛吊飾-星", 283, 587, 447, 614), ("願小金毛", 283, 602, 355, 624), 280, 10, 196, 1960),
        ("CU45045", (62, 618, 128, 636), ("線條小狗絨毛吊飾-雨", 281, 616, 447, 643), ("夜奇遇", 282, 630, 337, 652), 450, 5, 315, 1575),
        ("CU73950", (57, 644, 124, 665), ("線條小狗絨毛吊飾-蜜", 279, 644, 446, 672), ("蜂小白", 278, 658, 334, 681), 280, 5, 196, 980),
        ("CU73951", (53, 675, 118, 693), ("線條小狗絨毛吊飾-蜜", 277, 674, 446, 702), ("蜂小金毛", 274, 687, 352, 714), 280, 5, 196, 980),
    ]
    expected_names = [
        "線條小狗絨毛吊飾-慶典小白",
        "線條小狗絨毛吊飾-慶典小金毛",
        "線條小狗絨毛吊飾-晚安星星小金毛",
        "線條小狗絨毛吊飾-晚安月亮小白",
        "線條小狗絨毛吊飾-星願小白",
        "線條小狗絨毛吊飾-星願小金毛",
        "線條小狗絨毛吊飾-雨夜奇遇",
        "線條小狗絨毛吊飾-蜜蜂小白",
        "線條小狗絨毛吊飾-蜜蜂小金毛",
    ]

    for code, anchor_box, first_name, second_name, retail, quantity, cost, amount in source_rows:
        entries.append(make_entry(module, code, *anchor_box))
        entries.append(make_entry(module, *first_name))
        entries.append(make_entry(module, *second_name))
        anchor_y = (anchor_box[1] + anchor_box[3]) / 2
        entries.extend(
            [
                make_entry(module, str(retail), 500, int(anchor_y - 9), 532, int(anchor_y + 9)),
                make_entry(module, str(quantity), 580, int(anchor_y - 9), 600, int(anchor_y + 9)),
                make_entry(module, f"{cost}.00", 620, int(anchor_y - 9), 676, int(anchor_y + 9)),
                make_entry(module, f"{amount:,}", 716, int(anchor_y - 9), 770, int(anchor_y + 9)),
            ]
        )

    rows, issues = module.parse_liying_rows(entries)
    actual_names = [row.name for row in rows]
    if issues:
        raise RuntimeError(f"麗嬰解析產生非預期整體疑點：{issues}")
    if actual_names != expected_names:
        raise RuntimeError(
            "麗嬰跨列品名分組錯誤：\n"
            + "\n".join(
                f"{index + 1}. {actual!r} != {expected!r}"
                for index, (actual, expected) in enumerate(zip(actual_names, expected_names))
                if actual != expected
            )
        )
    if sum(int(row.amount or 0) for row in rows) != 11865:
        raise RuntimeError("麗嬰測試列金額合計不是 11865。")
    if any(row.check != "通過" for row in rows):
        raise RuntimeError("麗嬰測試列金額核對未全部通過。")
    print("LIYING_NAME_BOUNDARY_REGRESSION_OK rows=9 total=11865")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

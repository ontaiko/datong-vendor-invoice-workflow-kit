#!/usr/bin/env python3
import argparse
import csv
import json
import re
import unicodedata
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from zoneinfo import ZoneInfo


NON_IDENTITY_PARENS = re.compile(r"\((?:全?\d+種|再販|單)\)", re.IGNORECASE)
LEADING_CATEGORY = re.compile(r"^\([^)]*\)")
IGNORED_CHARS = re.compile(r"[\s\-_/\\・:：,，.。+＋]+")
REFERENCE_DATE = re.compile(r"(\d{7})$")


def normalize_name(value):
    text = unicodedata.normalize("NFKC", str(value or "")).strip()
    text = LEADING_CATEGORY.sub("", text)
    text = NON_IDENTITY_PARENS.sub("", text)
    return IGNORED_CHARS.sub("", text).casefold()


def bigrams(text):
    if len(text) < 2:
        return {text} if text else set()
    return {text[index : index + 2] for index in range(len(text) - 1)}


def similarity(left, right):
    if not left or not right:
        return 0.0
    left_bigrams = bigrams(left)
    right_bigrams = bigrams(right)
    dice = (
        (2 * len(left_bigrams & right_bigrams)) / (len(left_bigrams) + len(right_bigrams))
        if left_bigrams and right_bigrams
        else 0.0
    )
    sequence = SequenceMatcher(None, left, right).ratio()
    containment = min(len(left), len(right)) / max(len(left), len(right)) if left in right or right in left else 0.0
    return max(dice, sequence, containment)


def read_csv_rows(path):
    last_error = None
    for encoding in ("utf-8-sig", "cp950", "utf-8"):
        try:
            with path.open("r", encoding=encoding, newline="") as handle:
                return list(csv.DictReader(handle))
        except UnicodeDecodeError as error:
            last_error = error
    raise last_error


def roc_date_label(value):
    return f"{value.year - 1911:03d}{value.month:02d}{value.day:02d}"


def assert_current_reference_csv(path):
    match = REFERENCE_DATE.search(path.stem)
    if not match:
        raise SystemExit(
            f"產品資料參考檔缺少日期標籤：{path.name}。"
            "請更新為今天的產品資料輸出 CSV，檔名末尾需包含 7 位數民國年月日。"
        )

    actual = match.group(1)
    expected = roc_date_label(datetime.now(ZoneInfo("Asia/Taipei")))
    if actual != expected:
        raise SystemExit(
            f"產品資料參考檔日期不是今天：{path.name}，目前標籤為 {actual}，"
            f"台北時區今天應為 {expected}。請更新產品資料輸出 CSV 後再繼續。"
        )


def display_name(item):
    return item.get("suggestedName") or item.get("name") or item.get("rawName") or ""


def compare_item(item, products, threshold, max_candidates):
    source_name = display_name(item)
    normalized = normalize_name(source_name)
    exact_matches = [product for product in products if normalized and product["normalizedName"] == normalized]

    result = dict(item)
    result.update(
        {
            "matchStatus": "new",
            "matchedProductCode": "",
            "matchedProductName": "",
            "similarCandidates": [],
            "matchNote": "產品資料 CSV 未找到合理候選。",
            "existingProduct": False,
        }
    )

    if len(exact_matches) == 1:
        product = exact_matches[0]
        result.update(
            {
                "matchStatus": "exact",
                "matchedProductCode": product["productCode"],
                "matchedProductName": product["productName"],
                "matchNote": "正規化名稱完全一致，自動沿用既有產品代號。",
                "existingProduct": True,
            }
        )
        if not result.get("productCode"):
            result["productCode"] = product["productCode"]
        if not result.get("name"):
            result["name"] = product["productName"]
        return result

    scored = []
    for product in products:
        score = similarity(normalized, product["normalizedName"])
        if score >= threshold:
            scored.append(
                {
                    "productCode": product["productCode"],
                    "productName": product["productName"],
                    "score": round(score, 4),
                }
            )

    scored.sort(key=lambda candidate: (-candidate["score"], candidate["productCode"]))
    result["similarCandidates"] = scored[:max_candidates]
    if scored:
        result["matchStatus"] = "similar"
        if len(exact_matches) > 1:
            result["matchNote"] = "有多筆正規化名稱一致的既有商品，需人工確認產品代號。"
        else:
            result["matchNote"] = "找到相似候選，需人工確認是否為同一商品。"
    return result


def main():
    parser = argparse.ArgumentParser(description="Match invoice products against an exported product CSV.")
    parser.add_argument("--products-json", required=True, type=Path)
    parser.add_argument("--csv", required=True, type=Path)
    parser.add_argument("--output-json", required=True, type=Path)
    parser.add_argument("--similar-threshold", type=float, default=0.45)
    parser.add_argument("--max-candidates", type=int, default=5)
    args = parser.parse_args()

    assert_current_reference_csv(args.csv)
    items = json.loads(args.products_json.read_text(encoding="utf-8-sig"))
    rows = read_csv_rows(args.csv)
    products = []
    for row in rows:
        code = str(row.get("1.產品代號", "")).strip()
        name = str(row.get("2.產品名稱", "")).strip()
        normalized = normalize_name(name)
        if code and name and normalized:
            products.append({"productCode": code, "productName": name, "normalizedName": normalized})

    matched = [compare_item(item, products, args.similar_threshold, args.max_candidates) for item in items]
    args.output_json.parent.mkdir(parents=True, exist_ok=True)
    args.output_json.write_text(json.dumps(matched, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()

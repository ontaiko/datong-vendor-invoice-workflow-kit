#!/usr/bin/env python3
from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path


def find_product_csv(workspace_root: Path) -> Path | None:
    reference_dir = workspace_root / "參考資料"
    preferred = reference_dir / "產品資料輸出.CSV"
    if preferred.exists():
        return preferred
    candidates = [
        path
        for pattern in ("產品資料輸出*.CSV", "產品資料輸出*.csv")
        for path in reference_dir.glob(pattern)
        if path.is_file()
    ]
    return max(candidates, key=lambda path: path.stat().st_mtime) if candidates else None


def main() -> int:
    parser = argparse.ArgumentParser(description="Check that the product export CSV was created or modified today.")
    parser.add_argument("--workspace-root", required=True, type=Path)
    args = parser.parse_args()

    csv_path = find_product_csv(args.workspace_root.expanduser().resolve())
    if csv_path is None:
        print("找不到參考資料/產品資料輸出*.CSV。")
        return 2

    stat = csv_path.stat()
    today = datetime.now().date()
    created = datetime.fromtimestamp(stat.st_ctime)
    modified = datetime.fromtimestamp(stat.st_mtime)
    if created.date() != today and modified.date() != today:
        print(
            f"產品資料檔不是今天版本：{csv_path}；"
            f"建立 {created:%Y-%m-%d %H:%M:%S}；修改 {modified:%Y-%m-%d %H:%M:%S}。"
        )
        return 2

    print(f"產品資料檔檢查通過：{csv_path}；修改 {modified:%Y-%m-%d %H:%M:%S}。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

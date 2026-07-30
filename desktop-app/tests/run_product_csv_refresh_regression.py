from __future__ import annotations

import os
import sys
from datetime import datetime as RealDateTime
from datetime import timedelta
from pathlib import Path


APP_DIR = Path(__file__).resolve().parents[1]
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import invoice_workflow as workflow


class TomorrowDateTime(RealDateTime):
    @classmethod
    def now(cls, tz=None):
        current = RealDateTime.now(tz)
        future = current + timedelta(days=1)
        return cls.fromtimestamp(future.timestamp(), tz)


def write_catalog(path: Path, product_name: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        f"1.產品代號,2.產品名稱\n000001,{product_name}\n",
        encoding="utf-8-sig",
    )


def main() -> int:
    test_root = (
        workflow.PROJECT_ROOT
        / ".codex-tmp"
        / "invoice-app-tests"
        / "product-csv-refresh"
    )
    workspace_root = test_root / "workspace"
    app_reference = test_root / "app-reference"
    source = workspace_root / "參考資料" / "產品資料輸出.CSV"
    target = app_reference / "產品資料輸出.CSV"

    write_catalog(source, "過期工作區資料")
    write_catalog(target, "過期APP副本")

    original_reference_dir = workflow.REFERENCE_DIR
    original_override = workflow._PRODUCT_CSV_OVERRIDE
    original_datetime = workflow.datetime
    workflow.REFERENCE_DIR = app_reference
    workflow._PRODUCT_CSV_OVERRIDE = original_override
    workflow.datetime = TomorrowDateTime
    try:
        first_failed = False
        try:
            workflow.prepare_product_csv_for_use(
                target,
                project_root=workspace_root,
                persist=False,
            )
        except RuntimeError as exc:
            first_failed = "不是今天建立或修改" in str(exc)
        if not first_failed:
            raise RuntimeError("第一次過期產品資料沒有被阻擋。")
        if workflow._PRODUCT_CSV_OVERRIDE != original_override:
            raise RuntimeError("驗證失敗時不應將過期路徑寫入目前產品資料狀態。")

        write_catalog(source, "今天更新的工作區資料")
        future_timestamp = TomorrowDateTime.now().timestamp()
        os.utime(source, (future_timestamp, future_timestamp))
        selected = workflow.prepare_product_csv_for_use(
            target,
            project_root=workspace_root,
            persist=False,
        )
        if selected != target.resolve():
            raise RuntimeError(f"重新載入後使用了錯誤路徑：{selected}")
        if "今天更新的工作區資料" not in target.read_text(encoding="utf-8-sig"):
            raise RuntimeError("APP 沒有在第二次檢查時重新載入工作區最新產品資料。")

        write_catalog(source, "過期工作區資料不可覆蓋")
        write_catalog(target, "直接更新APP副本")
        os.utime(target, (future_timestamp, future_timestamp))
        synced = workflow.sync_reference_data(workspace_root)
        if "產品資料輸出.CSV" in synced:
            raise RuntimeError("同步流程不應使用過期工作區資料覆蓋今天的 APP 副本。")
        if "直接更新APP副本" not in target.read_text(encoding="utf-8-sig"):
            raise RuntimeError("啟動時同步流程覆蓋了直接更新的 APP 副本。")
        selected = workflow.prepare_product_csv_for_use(
            target,
            project_root=workspace_root,
            persist=False,
        )
        if "直接更新APP副本" not in selected.read_text(encoding="utf-8-sig"):
            raise RuntimeError("已直接更新且通過驗證的 APP 副本被舊工作區檔案覆蓋。")
    finally:
        workflow.REFERENCE_DIR = original_reference_dir
        workflow._PRODUCT_CSV_OVERRIDE = original_override
        workflow.datetime = original_datetime

    print("PRODUCT_CSV_REFRESH_REGRESSION_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

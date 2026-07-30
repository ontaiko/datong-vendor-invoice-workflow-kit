from __future__ import annotations

import hashlib
import json
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo


APP_DIR = Path(__file__).resolve().parents[1]
OUTPUT = APP_DIR / "package-manifest.json"

REQUIRED_FILES = [
    "invoice_ocr_excel_gui.py",
    "invoice_workflow.py",
    "build_exe.ps1",
    "app_settings.example.json",
    "installer/build-installer.ps1",
    "installer/run-installer-smoke-test.ps1",
    "installer/setup-runtime.ps1",
    "installer/verify-installation.ps1",
    "installer/大統進貨助手.iss",
    "scripts/local_paddleocr_invoice_to_xlsx.py",
    "scripts/match-existing-products.py",
    "scripts/review-invoice-product-check.py",
    "scripts/fill-import-templates.ps1",
    "reference_data/OCR設定.json",
    "reference_data/產品資料輸出.CSV",
    "reference_data/產品比對身份關鍵詞.csv",
    "reference_data/品牌括號命名規則.csv",
    "reference_data/廠商代號.xlsx",
    "reference_data/建檔用.xls",
    "reference_data/採購單匯入範例.xls",
]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def main() -> int:
    missing = [relative for relative in REQUIRED_FILES if not (APP_DIR / relative).exists()]
    if missing:
        raise SystemExit("Missing package files:\n" + "\n".join(missing))

    files = []
    for relative in REQUIRED_FILES:
        path = APP_DIR / relative
        files.append(
            {
                "path": relative.replace("\\", "/"),
                "size": path.stat().st_size,
                "sha256": sha256(path),
            }
        )

    manifest = {
        "schema_version": 1,
        "app_name": "大統進貨助手",
        "generated_at": datetime.now(ZoneInfo("Asia/Taipei")).isoformat(timespec="seconds"),
        "file_count": len(files),
        "files": files,
    }
    OUTPUT.write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"ok": True, "output": str(OUTPUT), "file_count": len(files)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

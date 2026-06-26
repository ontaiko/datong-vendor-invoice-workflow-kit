from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import hashlib
from dataclasses import dataclass
from pathlib import Path
from typing import Any


PROJECT_ROOT = Path(os.environ.get("DATONG_WORKSPACE", Path.home() / "Documents" / "大統工作助手"))
CODEX_HOME = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))
OCR_SCRIPT = CODEX_HOME / "skills" / "extract-vendor-invoice-image" / "scripts" / "local_paddleocr_invoice_to_xlsx.py"
SETTINGS_PATH = PROJECT_ROOT / "參考資料" / "OCR設定.json"
TMP_ROOT = PROJECT_ROOT / ".codex-tmp" / "invoice-image-to-spreadsheet"
SYSTEM_PYTHON = Path.home() / "AppData" / "Local" / "Programs" / "Python" / "Python312" / "python.exe"
VENV_PYTHON = PROJECT_ROOT / ".venv-paddleocr" / "Scripts" / "python.exe"
SUPPORTED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
_SELECTED_PYTHON: Path | None = None


@dataclass
class OcrRunResult:
    image: Path
    status: str
    output: Path | None
    vendor: str
    row_count: int
    needs_review: bool
    issues: list[str]
    error: str = ""


def compact_error(text: str, limit: int = 3000) -> str:
    cleaned = "\n".join(line.rstrip() for line in str(text or "").splitlines() if line.strip())
    return cleaned[-limit:] if len(cleaned) > limit else cleaned


def python_candidates() -> list[Path]:
    return [path for path in (SYSTEM_PYTHON, VENV_PYTHON) if path.exists()]


def import_check(python: Path) -> tuple[bool, str]:
    completed = subprocess.run(
        [str(python), "-X", "utf8", "-c", "import paddleocr, paddle"],
        cwd=str(PROJECT_ROOT),
        env=runtime_env(),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        check=False,
    )
    output = "\n".join(part for part in [completed.stdout, completed.stderr] if part)
    return completed.returncode == 0, compact_error(output)


def select_python() -> Path:
    global _SELECTED_PYTHON
    if _SELECTED_PYTHON and _SELECTED_PYTHON.exists():
        return _SELECTED_PYTHON
    errors: list[str] = []
    for python in python_candidates():
        ok, detail = import_check(python)
        if ok:
            _SELECTED_PYTHON = python
            return python
        errors.append(f"{python}\n{detail}")
    if not errors:
        raise RuntimeError(
            "找不到可執行的 OCR Python。\n"
            f"已檢查：\n- {SYSTEM_PYTHON}\n- {VENV_PYTHON}"
        )
    raise RuntimeError("找不到可用的 PaddleOCR 環境。\n\n" + "\n\n".join(errors))


def validate_runtime() -> list[str]:
    issues: list[str] = []
    if not OCR_SCRIPT.exists():
        issues.append(f"找不到 OCR 腳本：{OCR_SCRIPT}")
    if not SETTINGS_PATH.exists():
        issues.append(f"找不到 OCR 設定檔：{SETTINGS_PATH}")
    try:
        select_python()
    except RuntimeError as exc:
        issues.append(str(exc))
    return issues


def runtime_env() -> dict[str, str]:
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PADDLE_PDX_ENABLE_MKLDNN_BYDEFAULT"] = "0"
    return env


def parse_json_summary(output: str) -> dict[str, Any]:
    matches = list(re.finditer(r"\{[\s\S]*?\}", output))
    for match in reversed(matches):
        try:
            data = json.loads(match.group(0))
        except json.JSONDecodeError:
            continue
        if isinstance(data, dict) and data.get("ok") is True:
            return data
    raise RuntimeError("OCR 已結束，但無法讀取結果摘要。")


def safe_tmp_dir(image: Path) -> Path:
    token = re.sub(r"[^A-Za-z0-9_.-]+", "_", image.stem).strip("._") or "invoice"
    suffix = hashlib.sha1(str(image.resolve()).encode("utf-8")).hexdigest()[:8]
    return TMP_ROOT / f"{token}-{suffix}"


def assert_supported_image(image: Path) -> None:
    if not image.exists():
        raise RuntimeError(f"找不到圖片：{image}")
    if image.suffix.lower() not in SUPPORTED_EXTENSIONS:
        supported = "、".join(sorted(SUPPORTED_EXTENSIONS))
        raise RuntimeError(f"不支援的圖片格式：{image.suffix}。支援格式：{supported}")


def run_ocr(image: Path) -> OcrRunResult:
    image = image.expanduser().resolve()
    try:
        assert_supported_image(image)
        python = select_python()
        output_dir = image.parent
        tmp_dir = safe_tmp_dir(image)
        command = [
            str(python),
            "-X",
            "utf8",
            str(OCR_SCRIPT),
            str(image),
            "--output-dir",
            str(output_dir),
            "--tmp-dir",
            str(tmp_dir),
            "--settings",
            str(SETTINGS_PATH),
        ]
        completed = subprocess.run(
            command,
            cwd=str(PROJECT_ROOT),
            env=runtime_env(),
            text=True,
            encoding="utf-8",
            errors="replace",
            capture_output=True,
            check=False,
        )
        output = "\n".join(part for part in [completed.stdout, completed.stderr] if part)
        if completed.returncode != 0:
            raise RuntimeError(output or f"OCR 程式結束代碼：{completed.returncode}")
        summary = parse_json_summary(output)
        output_path = Path(str(summary.get("output", ""))).resolve()
        needs_review = bool(summary.get("needs_review"))
        issues = [str(issue) for issue in (summary.get("issues") or []) if str(issue).strip()]
        return OcrRunResult(
            image=image,
            status="需覆核" if needs_review else "完成",
            output=output_path,
            vendor=str(summary.get("vendor", "")),
            row_count=int(summary.get("row_count", 0) or 0),
            needs_review=needs_review,
            issues=issues,
        )
    except Exception as exc:
        return OcrRunResult(
            image=image,
            status="失敗",
            output=None,
            vendor="",
            row_count=0,
            needs_review=False,
            issues=[],
            error=compact_error(str(exc)),
        )


def run_self_test() -> int:
    issues = validate_runtime()
    result = {
        "ok": not issues,
        "project_root": str(PROJECT_ROOT),
        "ocr_script": str(OCR_SCRIPT),
        "settings": str(SETTINGS_PATH),
        "python": str(select_python()) if not issues else "",
        "issues": issues,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(run_self_test())

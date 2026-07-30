from __future__ import annotations

import os
import tempfile
import time
from pathlib import Path

import invoice_workflow as workflow


def write_with_mtime(path: Path, text: str, mtime: float) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")
    os.utime(path, (mtime, mtime))


def main() -> int:
    original_reference_dir = workflow.REFERENCE_DIR
    try:
        with tempfile.TemporaryDirectory(prefix="datong-reference-sync-") as temp_dir:
            root = Path(temp_dir)
            project_root = root / "workspace"
            source_dir = project_root / "參考資料"
            target_dir = root / "installed" / "reference_data"
            workflow.REFERENCE_DIR = target_dir

            now = time.time()
            name = "OCR設定.json"
            source = source_dir / name
            target = target_dir / name

            write_with_mtime(source, '{"source":"older"}', now - 60)
            write_with_mtime(target, '{"target":"newer-user-edit"}', now)
            synced = workflow.sync_reference_data(project_root)
            assert name not in synced
            assert "newer-user-edit" in target.read_text(encoding="utf-8")

            write_with_mtime(source, '{"source":"newer-workspace"}', now + 60)
            synced = workflow.sync_reference_data(project_root)
            assert name in synced
            assert "newer-workspace" in target.read_text(encoding="utf-8")
    finally:
        workflow.REFERENCE_DIR = original_reference_dir

    print("REFERENCE_DATA_SYNC_REGRESSION_OK")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

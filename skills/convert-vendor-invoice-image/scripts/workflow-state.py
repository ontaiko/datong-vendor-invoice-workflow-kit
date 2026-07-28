#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


VALID_STAGES = {
    "reference-precheck",
    "extract",
    "match",
    "review",
    "build",
    "cleanup",
    "complete",
}
VALID_STATUSES = {"active", "waiting", "completed", "failed"}

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8")


def now_iso() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")


def resolved(path: str | Path) -> Path:
    return Path(path).expanduser().resolve()


def unique(values: list[str]) -> list[str]:
    return list(dict.fromkeys(values))


def load_state(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise SystemExit(f"State file not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def save_state(path: Path, state: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    state["updated_at"] = now_iso()
    path.write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def parse_pairs(values: list[str], value_type: type = str) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for value in values:
        if "=" not in value:
            raise SystemExit(f"Expected KEY=VALUE, got: {value}")
        key, raw = value.split("=", 1)
        key = key.strip()
        if not key:
            raise SystemExit(f"Empty key in: {value}")
        result[key] = value_type(raw.strip())
    return result


def command_init(args: argparse.Namespace) -> int:
    state_path = resolved(args.state)
    if state_path.exists() and not args.force:
        raise SystemExit(f"State file already exists: {state_path}")
    state = {
        "version": 1,
        "workspace_root": str(resolved(args.workspace_root)),
        "source_files": unique([str(resolved(p)) for p in args.source]),
        "stage": "reference-precheck",
        "status": "active",
        "current_skill": "",
        "artifacts": {},
        "intermediate_files": [],
        "final_files": [],
        "counts": {},
        "pending": [],
        "excluded": [],
        "notes": [],
        "created_at": now_iso(),
    }
    save_state(state_path, state)
    print(state_path)
    return 0


def command_update(args: argparse.Namespace) -> int:
    state_path = resolved(args.state)
    state = load_state(state_path)
    if args.stage:
        state["stage"] = args.stage
    if args.status:
        state["status"] = args.status
    if args.current_skill is not None:
        state["current_skill"] = args.current_skill
    state["artifacts"].update(parse_pairs(args.artifact))
    state["counts"].update(parse_pairs(args.count, int))
    for key, path in list(state["artifacts"].items()):
        state["artifacts"][key] = str(resolved(path))
    if args.clear_pending:
        state["pending"] = []
    if args.clear_excluded:
        state["excluded"] = []
    if args.clear_notes:
        state["notes"] = []
    state["pending"] = unique(state["pending"] + args.pending)
    state["excluded"] = unique(state["excluded"] + args.exclude)
    state["notes"] = unique(state["notes"] + args.note)
    state["intermediate_files"] = unique(
        state["intermediate_files"]
        + [str(resolved(path)) for path in args.intermediate]
    )
    state["final_files"] = unique(
        state["final_files"] + [str(resolved(path)) for path in args.final]
    )
    save_state(state_path, state)
    print(state_path)
    return 0


def command_show(args: argparse.Namespace) -> int:
    state = load_state(resolved(args.state))
    if not args.compact:
        print(json.dumps(state, ensure_ascii=False, indent=2))
        return 0
    compact = {
        "stage": state.get("stage"),
        "status": state.get("status"),
        "current_skill": state.get("current_skill"),
        "artifacts": state.get("artifacts", {}),
        "counts": state.get("counts", {}),
        "pending_count": len(state.get("pending", [])),
        "excluded_count": len(state.get("excluded", [])),
    }
    print(json.dumps(compact, ensure_ascii=False, separators=(",", ":")))
    return 0


def is_within(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def command_cleanup(args: argparse.Namespace) -> int:
    state_path = resolved(args.state)
    state = load_state(state_path)
    workspace = resolved(state["workspace_root"])
    sources = {resolved(path) for path in state.get("source_files", [])}
    finals = {resolved(path) for path in state.get("final_files", [])}
    missing_finals = [str(path) for path in finals if not path.is_file()]
    if not finals or missing_finals:
        print(
            json.dumps(
                {"deleted": [], "error": "final validation failed", "missing": missing_finals},
                ensure_ascii=False,
            )
        )
        return 2

    reference_dir = workspace / "參考資料"
    deleted: list[str] = []
    skipped: list[str] = []
    for raw in state.get("intermediate_files", []):
        path = resolved(raw)
        protected = (
            path in sources
            or path in finals
            or not is_within(path, workspace)
            or is_within(path, reference_dir)
        )
        if protected or not path.is_file():
            skipped.append(str(path))
            continue
        path.unlink()
        deleted.append(str(path))

    state["stage"] = "complete"
    state["status"] = "completed"
    state["cleanup"] = {"deleted": deleted, "skipped": skipped}
    save_state(state_path, state)
    if args.remove_state:
        state_path.unlink()
        try:
            state_path.parent.rmdir()
        except OSError:
            pass
    print(
        json.dumps(
            {"deleted_count": len(deleted), "skipped_count": len(skipped)},
            ensure_ascii=False,
        )
    )
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Maintain compact vendor invoice workflow state.")
    subparsers = parser.add_subparsers(dest="command", required=True)

    init = subparsers.add_parser("init")
    init.add_argument("--state", required=True)
    init.add_argument("--workspace-root", required=True)
    init.add_argument("--source", action="append", required=True)
    init.add_argument("--force", action="store_true")
    init.set_defaults(func=command_init)

    update = subparsers.add_parser("update")
    update.add_argument("--state", required=True)
    update.add_argument("--stage", choices=sorted(VALID_STAGES))
    update.add_argument("--status", choices=sorted(VALID_STATUSES))
    update.add_argument("--current-skill")
    update.add_argument("--artifact", action="append", default=[])
    update.add_argument("--intermediate", action="append", default=[])
    update.add_argument("--final", action="append", default=[])
    update.add_argument("--count", action="append", default=[])
    update.add_argument("--pending", action="append", default=[])
    update.add_argument("--exclude", action="append", default=[])
    update.add_argument("--note", action="append", default=[])
    update.add_argument("--clear-pending", action="store_true")
    update.add_argument("--clear-excluded", action="store_true")
    update.add_argument("--clear-notes", action="store_true")
    update.set_defaults(func=command_update)

    show = subparsers.add_parser("show")
    show.add_argument("--state", required=True)
    show.add_argument("--compact", action="store_true")
    show.set_defaults(func=command_show)

    cleanup = subparsers.add_parser("cleanup")
    cleanup.add_argument("--state", required=True)
    cleanup.add_argument("--remove-state", action="store_true")
    cleanup.set_defaults(func=command_cleanup)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())

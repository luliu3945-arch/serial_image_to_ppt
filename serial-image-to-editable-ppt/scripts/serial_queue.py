from __future__ import annotations

import argparse
import json
import re
from datetime import datetime, timezone
from pathlib import Path


STATE_NAME = ".serial_image_to_ppt_state.json"


def now() -> str:
    return datetime.now(timezone.utc).isoformat()


def state_path(output_dir: Path) -> Path:
    return output_dir / STATE_NAME


def load(output_dir: Path) -> dict:
    path = state_path(output_dir)
    if not path.exists():
        raise SystemExit(f"Queue state not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def save(output_dir: Path, state: dict) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    state["updated_at"] = now()
    state_path(output_dir).write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def page_range(start: int, end: int) -> list[int]:
    if start > end:
        raise SystemExit("Serial queue requires ascending order: --start must be <= --end")
    return list(range(start, end + 1))


def discover_pages(input_dir: Path, pattern: str) -> list[int]:
    if "{page}" not in pattern:
        raise SystemExit("Pattern must contain {page} when --end is omitted")
    escaped = re.escape(pattern).replace(re.escape("{page}"), r"(\d+)")
    matcher = re.compile(f"^{escaped}$")
    found = []
    for item in input_dir.iterdir():
        match = matcher.match(item.name)
        if match:
            found.append(int(match.group(1)))
    return sorted(set(found))


def has_complete_evidence(output_dir: Path, page: int) -> bool:
    bundle = output_dir / f"page_{page}_refined_editable_output"
    required = [
        output_dir / f"page_{page}_refined_editable.pptx",
        bundle / "visual_elements_manifest.json",
        bundle / "review" / f"page_{page}_refined_elements_overlay.png",
        bundle / "quality_preview" / f"page_{page}_refined_preview.png",
        bundle / "quality_report" / "quality_report.json",
        bundle / "quality_report" / f"page_{page}_diff.png",
    ]
    return all(item.exists() and item.stat().st_size > 0 for item in required)


def current(state: dict) -> int | None:
    completed = set(state["completed"])
    return next((page for page in state["pages"] if page not in completed), None)


def main() -> None:
    parser = argparse.ArgumentParser(description="Manage a strict serial page rebuild queue.")
    sub = parser.add_subparsers(dest="command", required=True)
    init = sub.add_parser("init")
    init.add_argument("--input-dir", type=Path, required=True)
    init.add_argument("--output-dir", type=Path, required=True)
    init.add_argument("--start", type=int, default=1)
    init.add_argument("--end", type=int)
    init.add_argument("--pattern", default="page_{page}.png")
    init.add_argument("--adopt-existing", action="store_true")
    for name in ("status", "current", "clean"):
        cmd = sub.add_parser(name)
        cmd.add_argument("--output-dir", type=Path, required=True)
    passed = sub.add_parser("pass")
    passed.add_argument("--output-dir", type=Path, required=True)
    passed.add_argument("--page", type=int, required=True)
    failed = sub.add_parser("fail")
    failed.add_argument("--output-dir", type=Path, required=True)
    failed.add_argument("--page", type=int, required=True)
    failed.add_argument("--reason", required=True)
    args = parser.parse_args()

    if args.command == "init":
        if not args.input_dir.is_absolute():
            raise SystemExit("A user-provided absolute --input-dir path is required before hook initialization")
        if not args.input_dir.exists():
            raise SystemExit(f"Input directory not found: {args.input_dir}")
        end = args.end
        if end is None:
            discovered = discover_pages(args.input_dir, args.pattern)
            if not discovered:
                raise SystemExit(f"No files matching {args.pattern} in {args.input_dir}")
            end = max(discovered)
        pages = page_range(args.start, end)
        missing = [page for page in pages if not (args.input_dir / args.pattern.format(page=page)).exists()]
        if missing:
            raise SystemExit(f"Missing source pages: {missing}")
        completed = [page for page in pages if args.adopt_existing and has_complete_evidence(args.output_dir, page)]
        state = {
            "version": 1,
            "input_dir": str(args.input_dir.resolve()),
            "output_dir": str(args.output_dir.resolve()),
            "pattern": args.pattern,
            "start": args.start,
            "end": end,
            "pages": pages,
            "completed": completed,
            "failures": [],
            "created_at": now(),
        }
        save(args.output_dir, state)
        print(json.dumps({"current": current(state), "completed": completed, "remaining": len(pages) - len(completed)}))
        return

    state = load(args.output_dir)
    active = current(state)
    if args.command == "status":
        print(json.dumps({"current": active, "completed": state["completed"], "remaining": len(state["pages"]) - len(state["completed"]), "state": str(state_path(args.output_dir))}, ensure_ascii=False))
    elif args.command == "current":
        source = None if active is None else str(Path(state["input_dir"]) / state["pattern"].format(page=active))
        print(json.dumps({"current": active, "source": source, "done": active is None}, ensure_ascii=False))
    elif args.command == "pass":
        if active != args.page:
            raise SystemExit(f"Serial gate violation: current page is {active}, not {args.page}")
        if not has_complete_evidence(args.output_dir, args.page):
            raise SystemExit(f"QA evidence incomplete for page {args.page}")
        state["completed"].append(args.page)
        save(args.output_dir, state)
        print(json.dumps({"passed": args.page, "next": current(state), "remaining": len(state["pages"]) - len(state["completed"])}))
    elif args.command == "fail":
        if active != args.page:
            raise SystemExit(f"Serial gate violation: current page is {active}, not {args.page}")
        state["failures"].append({"page": args.page, "reason": args.reason, "time": now()})
        save(args.output_dir, state)
        print(json.dumps({"failed": args.page, "current": active, "reason": args.reason}, ensure_ascii=False))
    elif args.command == "clean":
        if active is not None:
            raise SystemExit(f"Cannot clean: page {active} and later pages are still pending")
        path = state_path(args.output_dir)
        path.unlink()
        print(json.dumps({"cleaned": str(path), "completed": len(state["completed"])}))


if __name__ == "__main__":
    main()

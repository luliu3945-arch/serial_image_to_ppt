from __future__ import annotations

import argparse
import csv
import json
import shutil
from pathlib import Path


def find_one(root: Path, name: str) -> Path:
    matches = sorted(path for path in root.rglob(name) if path.is_file() and path.stat().st_size > 0)
    if len(matches) != 1:
        raise SystemExit(f"Expected exactly one {name} under {root}; found {len(matches)}")
    return matches[0]


def main() -> None:
    parser = argparse.ArgumentParser(description="Stage baseline decomposition manifests into one refined page bundle.")
    parser.add_argument("--baseline-dir", required=True, type=Path)
    parser.add_argument("--bundle", required=True, type=Path)
    args = parser.parse_args()

    if not args.baseline_dir.is_dir():
        raise SystemExit(f"Baseline directory not found: {args.baseline_dir}")

    json_source = find_one(args.baseline_dir, "visual_elements_manifest.json")
    csv_source = find_one(args.baseline_dir, "visual_elements_manifest.csv")

    json_rows = json.loads(json_source.read_text(encoding="utf-8"))
    if not isinstance(json_rows, list) or not json_rows:
        raise SystemExit(f"Baseline JSON manifest is empty or invalid: {json_source}")

    with csv_source.open(newline="", encoding="utf-8-sig") as handle:
        csv_rows = list(csv.reader(handle))
    if len(csv_rows) < 2:
        raise SystemExit(f"Baseline CSV manifest is empty or invalid: {csv_source}")

    args.bundle.mkdir(parents=True, exist_ok=True)
    json_target = args.bundle / "baseline_visual_elements_manifest.json"
    csv_target = args.bundle / "baseline_visual_elements_manifest.csv"
    shutil.copy2(json_source, json_target)
    shutil.copy2(csv_source, csv_target)

    result = {
        "baseline_dir": str(args.baseline_dir.resolve()),
        "bundle": str(args.bundle.resolve()),
        "json_rows": len(json_rows),
        "csv_rows": len(csv_rows) - 1,
        "json": str(json_target.resolve()),
        "csv": str(csv_target.resolve()),
    }
    print(json.dumps(result, ensure_ascii=False))


if __name__ == "__main__":
    main()

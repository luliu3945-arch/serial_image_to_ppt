from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


def companion_root() -> Path:
    codex_home = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))
    root = codex_home / "skills" / "codeximage-to-editable-ppt-v1"
    if not root.exists():
        raise SystemExit(f"Missing companion skill: {root}")
    return root


def main() -> None:
    parser = argparse.ArgumentParser(description="Run one-page baseline decomposition with strict single-worker settings.")
    parser.add_argument("--input", required=True, type=Path)
    parser.add_argument("--outdir", required=True, type=Path)
    parser.add_argument("--python", default=sys.executable)
    parser.add_argument("--dpi", type=int, default=300)
    args = parser.parse_args()
    if not args.input.exists():
        raise SystemExit(f"Input not found: {args.input}")
    runner = companion_root() / "scripts" / "run_batches.py"
    command = [
        args.python,
        str(runner),
        str(args.input),
        "--outdir",
        str(args.outdir),
        "--batch-size",
        "1",
        "--batch-workers",
        "1",
        "--workers",
        "1",
        "--dpi",
        str(args.dpi),
        "--granularity",
        "fine",
        "--review",
        "--clean",
        "--clean-batch-root",
        "--no-merge",
    ]
    result = subprocess.run(command, check=False)
    raise SystemExit(result.returncode)


if __name__ == "__main__":
    main()


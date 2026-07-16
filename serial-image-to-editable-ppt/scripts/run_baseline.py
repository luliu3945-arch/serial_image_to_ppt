from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path


def companion_root() -> Path:
    override = os.environ.get("SERIAL_IMAGE_TO_PPT_COMPANION")
    candidates = []
    if override:
        candidates.append(Path(override).expanduser())
    # The normal installation keeps both skills beside each other.
    candidates.append(Path(__file__).resolve().parents[2] / "codeximage-to-editable-ppt-v1")
    # Current Codex user-skill location, followed by legacy compatibility paths.
    candidates.append(Path.home() / ".agents" / "skills" / "codeximage-to-editable-ppt-v1")
    codex_home = Path(os.environ.get("CODEX_HOME", Path.home() / ".codex"))
    candidates.append(codex_home / "skills" / "codeximage-to-editable-ppt-v1")
    candidates.append(Path.home() / ".codex" / "skills" / "codeximage-to-editable-ppt-v1")
    for root in candidates:
        if (root / "SKILL.md").is_file():
            return root
    searched = "\n  - ".join(str(path) for path in candidates)
    raise SystemExit(f"Missing companion skill. Searched:\n  - {searched}")


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

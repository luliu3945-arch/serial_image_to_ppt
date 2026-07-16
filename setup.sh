#!/usr/bin/env bash
set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SKILL_ROOT="${SKILL_ROOT:-$HOME/.agents/skills}"
RUNTIME_ROOT="${RUNTIME_ROOT:-$HOME/.serial-image-to-ppt}"
PYTHON_COMMAND="${PYTHON_COMMAND:-python3}"
VENV_DIR="$RUNTIME_ROOT/venv"
VENV_PYTHON="$VENV_DIR/bin/python"
SKILLS=(serial-image-to-editable-ppt codeximage-to-editable-ppt-v1)

echo "1/5 Checking repository files..."
for name in "${SKILLS[@]}"; do
  test -f "$REPO_ROOT/$name/SKILL.md" || { echo "Missing $name/SKILL.md" >&2; exit 1; }
done

echo "2/5 Checking Python 3.10+..."
"$PYTHON_COMMAND" -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'

echo "3/5 Creating/updating isolated Python runtime..."
mkdir -p "$RUNTIME_ROOT"
if [[ ! -x "$VENV_PYTHON" ]]; then
  "$PYTHON_COMMAND" -m venv "$VENV_DIR"
fi
"$VENV_PYTHON" -m pip install --upgrade pip
"$VENV_PYTHON" -m pip install -r "$REPO_ROOT/requirements.txt"

echo "4/5 Installing both skills to $SKILL_ROOT..."
mkdir -p "$SKILL_ROOT"
for name in "${SKILLS[@]}"; do
  mkdir -p "$SKILL_ROOT/$name"
  cp -R "$REPO_ROOT/$name/." "$SKILL_ROOT/$name/"
done

"$VENV_PYTHON" -c 'import json, pathlib, sys, datetime
target = pathlib.Path(sys.argv[1])
payload = {"python": str(pathlib.Path(sys.executable).resolve()), "repository": str(pathlib.Path(sys.argv[2]).resolve()), "installed_at": datetime.datetime.now(datetime.timezone.utc).isoformat()}
target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")' \
  "$SKILL_ROOT/serial-image-to-editable-ppt/runtime.json" "$REPO_ROOT"

echo "5/5 Running environment check..."
"$VENV_PYTHON" "$REPO_ROOT/scripts/doctor.py" --strict --skill-root "$SKILL_ROOT"

echo
echo "Installation complete. Restart Codex if the skills do not appear immediately."
echo "Skills: $SKILL_ROOT"
echo "Python runtime: $VENV_PYTHON"

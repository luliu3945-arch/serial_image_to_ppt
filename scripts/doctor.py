from __future__ import annotations

import argparse
import importlib
import json
import os
import platform
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PACKAGE_IMPORTS = {
    "opencv-python": "cv2",
    "numpy": "numpy",
    "pillow": "PIL",
    "python-pptx": "pptx",
    "pytesseract": "pytesseract",
    "PyYAML": "yaml",
}


def command_version(command: str, args: list[str]) -> dict[str, object]:
    executable = shutil.which(command)
    if not executable:
        return {"available": False, "path": None, "version": None}
    try:
        result = subprocess.run(
            [executable, *args],
            capture_output=True,
            text=True,
            errors="replace",
            timeout=15,
            check=False,
        )
        output = (result.stdout or result.stderr).strip().splitlines()
        version = output[0] if output else f"exit {result.returncode}"
        available = result.returncode == 0
    except Exception as exc:  # pragma: no cover - diagnostic fallback
        version = f"check failed: {exc}"
        available = False
    return {"available": available, "path": executable, "version": version}


def package_status() -> dict[str, dict[str, object]]:
    result = {}
    for package, module_name in PACKAGE_IMPORTS.items():
        try:
            module = importlib.import_module(module_name)
            version = getattr(module, "__version__", "installed")
            result[package] = {"available": True, "version": str(version)}
        except Exception as exc:
            result[package] = {"available": False, "version": None, "error": str(exc)}
    return result


def powerpoint_status() -> dict[str, object]:
    if os.name != "nt":
        return {"available": False, "reason": "PowerPoint COM validation is Windows-only"}
    powershell = shutil.which("powershell") or shutil.which("pwsh")
    if not powershell:
        return {"available": False, "reason": "PowerShell not found"}
    script = (
        "$ErrorActionPreference='Stop';"
        "$p=New-Object -ComObject PowerPoint.Application;"
        "$v=$p.Version;$p.Quit();Write-Output $v"
    )
    try:
        result = subprocess.run(
            [powershell, "-NoProfile", "-Command", script],
            capture_output=True,
            text=True,
            errors="replace",
            timeout=30,
            check=False,
        )
        if result.returncode == 0 and result.stdout.strip():
            return {"available": True, "version": result.stdout.strip()}
        return {"available": False, "reason": (result.stderr or result.stdout).strip()}
    except Exception as exc:
        return {"available": False, "reason": str(exc)}


def tesseract_languages() -> list[str]:
    executable = shutil.which("tesseract")
    if not executable:
        return []
    try:
        result = subprocess.run(
            [executable, "--list-langs"],
            capture_output=True,
            text=True,
            errors="replace",
            timeout=15,
            check=False,
        )
        lines = [line.strip() for line in result.stdout.splitlines() if line.strip()]
        return [line for line in lines if not line.lower().startswith("list of available")]
    except Exception:
        return []


def skill_status(extra_root: Path | None = None) -> dict[str, object]:
    names = ["serial-image-to-editable-ppt", "codeximage-to-editable-ppt-v1"]
    roots = [
        Path.home() / ".agents" / "skills",
        Path(os.environ.get("CODEX_HOME", Path.home() / ".codex")) / "skills",
        Path.home() / ".codex" / "skills",
    ]
    if extra_root is not None:
        roots.insert(0, extra_root.expanduser().resolve())
    roots = list(dict.fromkeys(root.resolve() for root in roots))
    found = {}
    for name in names:
        matches = [str(root / name) for root in roots if (root / name / "SKILL.md").is_file()]
        found[name] = matches
    source_complete = all((ROOT / name / "SKILL.md").is_file() for name in names)
    return {"search_roots": [str(root) for root in roots], "found": found, "source_complete": source_complete}


def main() -> None:
    parser = argparse.ArgumentParser(description="Check the serial-image-to-PPT runtime and optional production tools.")
    parser.add_argument("--json", action="store_true", help="Print machine-readable JSON only")
    parser.add_argument("--strict", action="store_true", help="Fail when core Python or skill requirements are missing")
    parser.add_argument("--skill-root", type=Path, help="Additional installed skill root to validate")
    parser.add_argument("--require-powerpoint", action="store_true", help="Treat Microsoft PowerPoint as required")
    parser.add_argument("--require-ocr", action="store_true", help="Treat Tesseract with eng and chi_sim as required")
    parser.add_argument("--require-document-input", action="store_true", help="Treat LibreOffice and pdftoppm as required")
    args = parser.parse_args()

    packages = package_status()
    skills = skill_status(args.skill_root)
    commands = {
        "git": command_version("git", ["--version"]),
        "node": command_version("node", ["--version"]),
        "tesseract": command_version("tesseract", ["--version"]),
        "libreoffice": command_version("libreoffice", ["--version"]),
        "soffice": command_version("soffice", ["--version"]),
        "pdftoppm": command_version("pdftoppm", ["-v"]),
        "codex": command_version("codex", ["--version"]),
    }
    powerpoint = powerpoint_status()
    languages = tesseract_languages()
    python_ok = sys.version_info >= (3, 10)
    packages_ok = all(item["available"] for item in packages.values())
    skills_ok = all(skills["found"][name] for name in skills["found"])
    failures = []
    if not python_ok:
        failures.append("Python 3.10 or newer is required")
    if not packages_ok:
        failures.append("One or more Python packages are missing")
    if not skills_ok:
        failures.append("Both skill folders must be available")
    if args.require_powerpoint and not powerpoint.get("available"):
        failures.append("Microsoft PowerPoint desktop/COM is required for this check")
    if args.require_ocr and not {"eng", "chi_sim"}.issubset(set(languages)):
        failures.append("Tesseract OCR languages eng and chi_sim are required for this check")
    if args.require_document_input:
        office_ok = commands["libreoffice"]["available"] or commands["soffice"]["available"]
        if not office_ok or not commands["pdftoppm"]["available"]:
            failures.append("LibreOffice and pdftoppm are required for PPT/PDF input")

    report = {
        "ok": not failures,
        "failures": failures,
        "platform": platform.platform(),
        "python": {"version": platform.python_version(), "executable": sys.executable, "ok": python_ok},
        "python_packages": packages,
        "skills": skills,
        "commands": commands,
        "powerpoint": powerpoint,
        "tesseract_languages": languages,
        "notes": {
            "powerpoint": "Required for true PowerPoint acceptance previews and merge_delivery.ps1 on Windows.",
            "ocr": "Optional for image-only manual reconstruction; required for automated OCR text extraction.",
            "document_input": "LibreOffice and pdftoppm are only needed when the source is PPT/PPTX/PDF instead of PNG/JPG.",
            "artifact_tool": "Provided by the Codex presentations runtime; do not install an unrelated npm package with the same name.",
        },
    }
    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"Environment: {'PASS' if report['ok'] else 'NEEDS ATTENTION'}")
        print(f"Python: {report['python']['version']} ({report['python']['executable']})")
        for name, item in packages.items():
            print(f"  {'OK' if item['available'] else 'MISSING':7} Python package {name} {item.get('version') or ''}")
        for name, matches in skills["found"].items():
            print(f"  {'OK' if matches else 'MISSING':7} Skill {name}: {matches[0] if matches else 'not found'}")
        print(f"  {'OK' if powerpoint.get('available') else 'OPTIONAL/MISSING':16} Microsoft PowerPoint {powerpoint.get('version', '')}")
        print(f"  {'OK' if languages else 'OPTIONAL/MISSING':16} Tesseract languages: {', '.join(languages) or 'none'}")
        office = commands["libreoffice"]["available"] or commands["soffice"]["available"]
        print(f"  {'OK' if office else 'OPTIONAL/MISSING':16} LibreOffice")
        print(f"  {'OK' if commands['pdftoppm']['available'] else 'OPTIONAL/MISSING':16} pdftoppm")
        if failures:
            print("Failures:")
            for failure in failures:
                print(f"  - {failure}")
    if args.strict and failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()

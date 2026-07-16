---
name: serial-image-to-editable-ppt
description: Strictly serial orchestration for rebuilding numbered slide screenshots into refined editable PPTX files. Use when Codex must require a user-provided source image address before creating any hook, process a page range one image at a time, take the next numbered image only after the current slide passes artifact and true-PowerPoint visual review plus structural, manifest, and bounds QA, resume safely after interruptions, preserve baseline evidence, audit and package all decks, optionally merge the approved page decks in order, and remove the temporary queue state.
---

# Serial Image to Editable PPT

Rebuild a numbered image sequence into separate editable PPTX files with a hard one-page-at-a-time gate. Reuse `codeximage-to-editable-ppt-v1` for slide reconstruction; this skill supplies queue control, naming, QA, resumption, audit, packaging, ordered merge, and cleanup.

## Required companion skill

Locate and read `codeximage-to-editable-ppt-v1/SKILL.md` before rebuilding the first page. Follow its refined editable rebuild rules: preserve complex visual regions as tightly cropped independent PNGs, recreate text and simple geometry as native PowerPoint objects, export a full-slide preview, and iterate visually.

If the companion skill is unavailable, stop and report the missing dependency. Do not silently replace the refined rebuild with a full-slide background image.

## Environment preflight

Before initializing the first queue on a machine, locate `runtime.json` beside this `SKILL.md` when it exists. Use its absolute `python` value for every script command instead of an arbitrary `python` on `PATH`. The repository installer writes this file after creating the tested virtual environment.

If `runtime.json` is absent, require Python 3.10+ and verify that `opencv-python`, `numpy`, `pillow`, `python-pptx`, `pytesseract`, and `PyYAML` import successfully. When the repository checkout is available, run `scripts/doctor.py --strict` from the repository root. Stop before queue creation if core Python packages or the companion skill are missing.

Treat environment features by input and acceptance mode:

- image-only baseline work does not require LibreOffice or Poppler;
- OCR requires the external Tesseract executable and suitable languages such as `eng` and `chi_sim`, not only the `pytesseract` Python package;
- PPT/PPTX/PDF source conversion requires LibreOffice plus `pdftoppm`;
- true PowerPoint previews and `merge_delivery.ps1` require Windows with Microsoft PowerPoint desktop/COM;
- artifact-tool comes from the Codex presentations runtime and must not be replaced with an unrelated npm package.

If an optional capability is unavailable, state the reduced verification level. Do not claim true-PowerPoint acceptance or COM merge without PowerPoint.

## Pre-hook source-address gate

Do not create, initialize, restore, or infer a hook/queue until the user explicitly provides one of these:

- an absolute path to the directory containing the numbered source images; or
- an absolute path to one numbered source image supplied in the current request.

An attachment with an exposed absolute local path counts as user-provided. A path remembered from another task, inferred from the current working directory, found by broad filesystem search, or copied from an unrelated earlier workflow does not count.

If the current request lacks a source address, stop before any queue mutation and ask only for the image directory or image path. Use this prompt:

```text
请先提供待处理图片所在目录，或任意一张编号图片的绝对路径；收到并验证路径后我再创建串行 hook。
```

Validate that the supplied path exists and is readable. If it is a file, derive the source directory from its parent and infer the filename pattern only when the filename contains an unambiguous page number. If the path or numbering is ambiguous, ask for correction. Show the discovered `1 → N` range before initializing the hook. Never create a placeholder hook while waiting for the path.

## Inputs and naming

After the source-address gate passes, collect or infer:

- source directory;
- inclusive start and end page numbers; default to page `1` through the highest discovered page;
- filename pattern, default `page_{page}.png`;
- output directory;
- processing order: always ascending `1 → N`, or ascending within an explicitly requested subset.

Use these output names:

- final deck: `page_<N>_refined_editable.pptx`;
- evidence bundle: `page_<N>_refined_editable_output/`;
- temporary baseline: `serial_hook/page_<N>_baseline/`;
- queue state: `<output-dir>/.serial_image_to_ppt_state.json`.

## Initialize or resume the queue

Initialize once:

```powershell
python scripts/serial_queue.py init `
  --input-dir "<source-dir>" --output-dir "<output-dir>" `
  --pattern "page_{page}.png"
```

The command automatically discovers the highest page and creates an ascending `1 → N` queue. Pass `--start` and `--end` only for a user-requested subset or explicit alternative order.

On every continuation, call `status` first and trust the state file plus filesystem rather than conversation memory:

```powershell
python scripts/serial_queue.py status --output-dir "<output-dir>"
```

If the state file is missing but final files already exist, reinitialize with `--adopt-existing`; it marks a page complete only when its deck and required QA evidence pass inspection.

## Hard serial gate

Process exactly the page returned by `current`:

```powershell
python scripts/serial_queue.py current --output-dir "<output-dir>"
```

Never build, crop, preview, or QA the next page while the current page is unresolved. Parallel work and subagents are forbidden unless the user explicitly overrides serial execution.

For each page, perform the following sequence.

### 1. Inspect the source

View the source at original detail. Identify:

- editable titles, labels, descriptions, tables, lines, borders, arrows, and color blocks;
- complex photos, software screenshots, icons, barcodes, diagrams, and device renders to crop;
- slide dimensions and repeated footer/header patterns.

### 2. Run baseline decomposition

Use the wrapper to invoke the companion skill's batch runner with one input and one worker:

```powershell
python scripts/run_baseline.py `
  --input "<source-file>" `
  --outdir "<output-dir>/serial_hook/page_<N>_baseline"
```

Do not enable LibreOffice-dependent quality checks when LibreOffice is unavailable.

Stage the baseline manifests into the refined output bundle:

```powershell
python scripts/stage_baseline_evidence.py `
  --baseline-dir "<output-dir>/serial_hook/page_<N>_baseline" `
  --bundle "<output-dir>/page_<N>_refined_editable_output"
```

The command must create:

- `baseline_visual_elements_manifest.json`;
- `baseline_visual_elements_manifest.csv`.

### 3. Build the refined editable slide

Use artifact-tool through the presentations runtime. Build a single-slide deck at the source aspect ratio.

- Crop complex visuals tightly and store them under `split_png_elements/image01/`.
- Recreate all practical text as editable text boxes.
- Recreate simple lines, arrows, cards, tables, dividers, status marks, and blocks as native shapes.
- Add meaningful alt text to every inserted image.
- Record every final element in `visual_elements_manifest.json` and CSV.
- Save the final deck both at the output root and inside its evidence bundle.

Read [references/rebuild-contract.md](references/rebuild-contract.md) before writing a new per-page builder or when a build/export issue occurs.

### 4. Preview and iterate

Export an artifact-tool full-slide PNG at source dimensions and inspect it at original detail. Fix visible defects such as:

- wrong wrapping or auto-fit;
- clipped text;
- stretched crops;
- incorrect symbols or emoji substitutions;
- uneven card spacing;
- missing connectors;
- poor alignment against the source.

Do not accept a page merely because the export command produced a file.

When Microsoft PowerPoint is available, open the rebuilt PPTX through PowerPoint and export another PNG at the source dimensions. Inspect this true PowerPoint preview at original detail. Use it, not the artifact-only preview, as the acceptance preview and as the `qa_page.py --preview` input.

If PowerPoint is unavailable, use the best available real presentation renderer and disclose the limitation. Never skip the full-size visual review.

### 5. Run the page gate

Generate the QA evidence and inspect the PPTX:

```powershell
python scripts/qa_page.py `
  --page <N> --source "<source-file>" --preview "<preview.png>" `
  --pptx "<output-dir>/page_<N>_refined_editable.pptx" `
  --bundle "<output-dir>/page_<N>_refined_editable_output"
```

The page passes only when all are true:

- the PPTX opens and contains exactly one slide;
- it contains editable text and independent pictures;
- manifest count equals actual shape count;
- no manifest element is outside the source canvas;
- both baseline manifests exist in the refined bundle;
- overlay, preview, original, recomposed, diff, and QA JSON exist;
- artifact-tool and true-renderer full-size inspections are acceptable.

Mark the page complete only after the gate passes:

```powershell
python scripts/serial_queue.py pass --output-dir "<output-dir>" --page <N>
```

If it fails, keep the same current page, record the reason, repair, and rerun:

```powershell
python scripts/serial_queue.py fail --output-dir "<output-dir>" --page <N> --reason "<concise reason>"
```

## Final audit and package

After `current` reports no remaining page, audit the entire inclusive range:

```powershell
python scripts/audit_delivery.py `
  --output-dir "<output-dir>" --start 1 --end <N> --create-zip
```

The audit must pass every page. It writes JSON and CSV audit manifests and a ZIP containing every individual PPTX plus both audit files. If it fails, read `missing_artifacts` in the audit JSON, restore the missing evidence, and rerun the complete audit before cleanup.

If the user wants one ordered PPTX after the individual audit passes, merge the page decks with PowerPoint's native slide insertion:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/merge_delivery.ps1 `
  -OutputDir "<output-dir>" -Start 1 -End <N> `
  -OutputFile "<output-dir>/pages_1_to_<N>_merged_editable.pptx" `
  -PreviewDir "<output-dir>/merged_preview"
```

Verify the merged slide count, ascending page order, exported previews, and editability. Do not concatenate raw artifact-tool protos from independently generated page decks because their internal slide and media IDs may collide.

Then clean the temporary hook/state:

```powershell
python scripts/serial_queue.py clean --output-dir "<output-dir>"
```

`clean` must refuse to run while any page is pending. Keep baseline and QA evidence unless the user explicitly asks to delete it; only remove the queue state hook.

If a heartbeat automation was created for the run, delete or disable it only after audit, optional merge validation, and queue cleanup all succeed.

For multi-dozen-page runs, audit recovery, true-renderer details, and observed failure modes, read [references/production-runbook.md](references/production-runbook.md).

## User updates

At the start of each page, state the page being processed. After it passes, report object, text, picture, and bounds counts, then name the next page. On completion, report the audited page count, totals, ZIP location, audit locations, and that the hook was removed.

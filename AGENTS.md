# Agent operating guide

This repository contains two cooperating Codex skills:

- `serial-image-to-editable-ppt`: queue control, strict serial gating, QA, audit, packaging, merge, and cleanup.
- `codeximage-to-editable-ppt-v1`: one-slide editable reconstruction.

Read both `SKILL.md` files before processing the first source image. For a production run, also read:

- `serial-image-to-editable-ppt/references/rebuild-contract.md`
- `serial-image-to-editable-ppt/references/production-runbook.md`

## Non-negotiable rules

1. Require a user-provided absolute source image path before creating or restoring a queue.
2. Run `status` and then `current` at the start of every continuation.
3. Process only the page returned by `current`.
4. Do not parallelize pages, delegate pages, pre-crop a later page, or skip a failed page unless the user explicitly overrides the serial contract.
5. Do not use a full-slide screenshot as the final editable slide.
6. Do not call `pass` until the page has baseline evidence, an editable PPTX, a true PowerPoint preview when PowerPoint is available, complete QA evidence, matching manifest/shape counts, and zero bounds issues.
7. Do not clean the queue or delete a heartbeat automation until the full delivery audit passes.

## Required per-page sequence

Use the repository scripts from `serial-image-to-editable-ppt/scripts/`.

```powershell
python serial_queue.py status --output-dir "<output-dir>"
python serial_queue.py current --output-dir "<output-dir>"
```

For the one returned page:

1. Inspect the source image at original resolution.
2. Run one-worker baseline decomposition.
3. Stage the baseline manifests into the refined bundle.
4. Rebuild the slide with artifact-tool and the companion skill.
5. Inspect the artifact-tool preview at original resolution and iterate.
6. Export a preview with real Microsoft PowerPoint on Windows when available.
7. Run `qa_page.py` using the real PowerPoint preview.
8. Inspect the QA result and only then call `serial_queue.py pass`.
9. Immediately run `status` and `current` again; the returned page is the only next page.

Baseline commands:

```powershell
python run_baseline.py `
  --input "<source-file>" `
  --outdir "<output-dir>/serial_hook/page_<N>_baseline"

python stage_baseline_evidence.py `
  --baseline-dir "<output-dir>/serial_hook/page_<N>_baseline" `
  --bundle "<output-dir>/page_<N>_refined_editable_output"
```

The staging command is mandatory. A real 48-page production run initially failed final audit on pages 42–48 because these two files were absent:

- `baseline_visual_elements_manifest.json`
- `baseline_visual_elements_manifest.csv`

## Refined rebuild policy

- Crop photos, screenshots, device renders, intricate icons, and complex diagrams as independent PNGs.
- Rebuild titles, body text, labels, table text, page numbers, cards, arrows, dividers, borders, and simple diagrams as native PowerPoint objects.
- Add alt text to every picture.
- Keep a per-page builder source so every correction is reproducible.
- Rebuild after every correction; do not patch only the preview.
- Treat artifact-tool exit code 1 as non-success unless the PPTX and preview both exist and pass inspection.
- Avoid negative-width or negative-height line bounds. Use explicit elbow segments for rising connectors if the renderer clips a negative-height line.

## True PowerPoint preview

On Windows with PowerPoint installed, open the one-page deck through COM and export at the source dimensions. Use that exported PNG as the `--preview` input to `qa_page.py`.

Do not rely only on artifact-tool rendering. The true PowerPoint pass catches font substitution, unexpected wrapping, connector rendering, and image clipping.

## Completion sequence

After `current` reports `done: true`:

1. Run `audit_delivery.py --create-zip`.
2. If audit fails, read the generated JSON and repair every missing artifact. Do not clean the queue.
3. Rerun the audit until `all_passed: true`.
4. If the user wants one deck, run `merge_delivery.ps1` only after the individual audit passes.
5. Verify the merged slide count, page order, and exported previews.
6. Run `serial_queue.py clean`.
7. Delete or disable the task heartbeat automation.

Example merge:

```powershell
powershell -ExecutionPolicy Bypass -File merge_delivery.ps1 `
  -OutputDir "<output-dir>" -Start 1 -End 48 `
  -OutputFile "<output-dir>/pages_1_to_48_merged_editable.pptx" `
  -PreviewDir "<output-dir>/merged_preview"
```

## Recovery rules

- Trust the state file and filesystem, not conversation memory.
- If the state file exists, resume with `status/current`.
- If it is missing, reinitialize with `--adopt-existing`; never infer completion from PPTX presence alone.
- A failed audit is not a completed delivery.
- A page stays current after `fail`; repair and rerun its gate.

## Repository validation

Before committing changes:

```powershell
python -m compileall serial-image-to-editable-ppt/scripts
python <skill-creator>/scripts/quick_validate.py serial-image-to-editable-ppt
python <skill-creator>/scripts/quick_validate.py codeximage-to-editable-ppt-v1
```

Test new deterministic scripts on a small disposable page range. Keep generated PPTX files, queue state, previews, and audit outputs outside the repository.

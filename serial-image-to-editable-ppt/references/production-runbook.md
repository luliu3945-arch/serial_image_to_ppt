# Production runbook

Use this reference for long-running, heartbeat-driven, or multi-dozen-page deliveries.

## Production loop

For every wake-up or continuation:

```text
status -> current -> inspect source -> baseline -> stage baseline evidence
-> refined editable build -> artifact preview -> true PowerPoint preview
-> qa_page -> pass -> status -> current
```

Only the page returned by `current` may be touched. A wake-up is one opportunity to advance one page, not permission to batch or parallelize.

## Evidence staging

`run_baseline.py` writes the baseline manifests inside a nested batch item directory. Copying them by hand is easy to miss. Always run:

```powershell
python scripts/stage_baseline_evidence.py `
  --baseline-dir "<output-dir>/serial_hook/page_<N>_baseline" `
  --bundle "<output-dir>/page_<N>_refined_editable_output"
```

Run it before the page gate. The final audit requires both staged files.

## Two-preview rule

Use two renderers when Microsoft PowerPoint is available:

1. Artifact-tool preview for fast iteration and layout debugging.
2. PowerPoint COM export for the acceptance preview.

Inspect both at full size. Pass the PowerPoint export to `qa_page.py`. Common defects visible only after the true renderer include:

- title or banner wrapping;
- cropped labels inside a retained diagram;
- missing or broken connectors;
- font fallback and symbol substitution;
- text shrinkage that is technically in-bounds but visually unacceptable.

## Builder practices that survived a 48-page run

- Keep complex diagrams and photos as tightly cropped independent images.
- Keep surrounding labels, explanations, borders, and arrows editable.
- Reuse stable header/footer construction, but create a separate builder for each page.
- Use explicit elbow segments instead of negative-height lines when a rising connector is clipped.
- After every visual correction, rebuild the PPTX, re-export, inspect, and rerun QA.
- Record final object counts in user updates so regressions are visible.

## Audit recovery

The audit JSON is the source of truth. If `all_passed` is false:

1. Read `missing_artifacts` for every failed page.
2. Restore only the missing evidence.
3. Rerun the complete audit range.
4. Create the ZIP only after every page passes.

Observed production failure: a 48-page run first audited as `41/48` because pages 42–48 lacked staged baseline manifests. After staging those files, the same run audited `48/48` and produced the delivery ZIP.

## Ordered merge

Merge only after the individual audit passes. Use PowerPoint's native `InsertFromFile` through `merge_delivery.ps1`; do not concatenate raw artifact-tool protos because independently generated decks can reuse internal slide and media IDs.

After merging:

- assert the merged slide count equals the requested range;
- export every merged slide with PowerPoint;
- verify slide `N` matches page `N`;
- inspect the merged deck with artifact-tool or PowerPoint to confirm objects remain editable;
- preserve the individual PPTX files and audit evidence.

In the 48-page production run, the merged file contained 48 slides and all 48 exported slides matched their corresponding verified page; the largest mean RGB difference was `6.892`, consistent with PowerPoint re-rendering rather than a layout change.

## Hook and automation cleanup

Cleanup is the final action, not part of audit recovery:

```powershell
python scripts/serial_queue.py clean --output-dir "<output-dir>"
```

Then remove the heartbeat automation that was keeping the task alive. Keep baseline folders and QA evidence unless the user explicitly requests deletion.

# Quality review workflow

1. Open `review/slideXX_detected_elements_overlay.png`.
2. Check whether text, icons, figures, charts, and decorations are split at the desired granularity.
3. Open `quality_report/slideXX_diff.png`.
4. Check `visual_elements_manifest.csv` for simple elements where `transparent_background` is `no`, especially `icon`, `logo`, `symbol`, `button`, `badge`, `arrow`, `line`, `divider`, `shape`, `simple_shape`, and text-only elements. Review `alpha_note` for the failure reason.
5. Large diff regions usually mean one of the following:
   - missing background decoration;
   - element alpha mask too aggressive;
   - simple icon/logo/line/shape background not removed cleanly;
   - text converted with wrong font size or line breaks;
   - OCR recognized text incorrectly;
   - two visual objects were merged or one object was split too much.
6. Rerun with a different `--granularity` setting or disable editable text for problematic pages.

Suggested commands:

```bash
# Less fragmented
python scripts/decompose_visual_elements.py input.pptx --outdir output_normal --granularity normal --ocr --editable-text --review --quality-check

# More detailed
python scripts/decompose_visual_elements.py input.pptx --outdir output_fine --granularity fine --ocr --editable-text --review --quality-check
```

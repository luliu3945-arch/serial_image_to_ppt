# Refined rebuild contract

## Output bundle contract

Each page bundle must contain:

```text
page_<N>_refined_editable_output/
├── page_<N>_refined_editable.pptx
├── visual_elements_manifest.json
├── visual_elements_manifest.csv
├── baseline_visual_elements_manifest.json
├── baseline_visual_elements_manifest.csv
├── split_png_elements.zip
├── split_png_elements/image01/*.png
├── review/page_<N>_refined_elements_overlay.png
├── quality_preview/page_<N>_refined_preview.png
└── quality_report/
    ├── quality_report.json
    ├── page_<N>_original.png
    ├── page_<N>_recomposed.png
    └── page_<N>_diff.png
```

## Crop policy

- Crop photos, software screens, barcodes, device renders, intricate icons, and diagrams.
- Keep each logically independent visual as a separate PNG when practical.
- Do not include editable labels or surrounding card borders inside a crop unless they are inseparable from a complex screenshot.
- Preserve aspect ratio with `fit: contain`; never stretch.

## Editable object policy

- Rebuild titles, labels, prose, scores, table text, footer text, and page numbers as text boxes.
- Rebuild cards, lines, arrows, checkmarks, dividers, table grids, and color blocks as native shapes.
- Use source-matched fonts, sizes, colors, alignment, and spacing.
- Use `autoFit: shrinkText` as a guard, then inspect for unwanted shrinkage or wrapping.

## Manifest contract

Record one row per final PowerPoint object. Required fields include:

- `image_id`, `element_id`, `element_type`;
- `x`, `y`, `width`, `height`, `z_order`;
- `is_text_only`, `recognized_text`;
- `file_name`, `relative_path` for cropped images;
- source canvas width and height;
- conversion note and review flag.

Manifest count must equal actual shape count in the final PPTX.

## Artifact-tool behavior

- Set `HOME` before configuring the artifact-tool workspace on Windows.
- A command may return a timeout or exit code 1 after writing inspect sidecars. Treat it as a successful build only when the final PPTX and preview exist and pass validation.
- Rebuild from the per-page source script after every layout correction.
- Do not use a full-slide background image as the final deliverable.

## Visual review checklist

- Compare the entire slide, not isolated crops.
- Check title and footer alignment first.
- Check all text for clipping, odd line breaks, and tiny fallback sizing.
- Check repeated card widths, gaps, and borders.
- Check arrows and connector continuity.
- Check that symbols did not render as unrelated glyphs.
- Check image aspect ratios and whitespace around crops.
- Re-export after every correction.


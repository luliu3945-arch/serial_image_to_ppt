# Refined Editable Rebuild Workflow

Use this workflow as the default V1 final-output path, regardless of language. This includes English, Chinese, mixed-language raster slides, dense diagrams, academic figures, and slide screenshots with many labels. Automated decomposition is still useful, but it is a reference step rather than the final deliverable.

## 1. Always run refined rebuild unless explicitly skipped

Run the normal decomposition first unless the user explicitly asks only for planning or raw extraction. Then produce the refined rebuild as the primary final output.

Recommended baseline command:

```bash
python scripts/decompose_visual_elements.py input.png \
  --outdir output \
  --dpi 300 \
  --granularity fine \
  --ocr \
  --ocr-lang chi_sim+eng \
  --ocr-confidence-threshold 75 \
  --editable-text \
  --review
```

If the local OCR runtime does not have the needed language package, use the available language for baseline visual discovery, but do not trust unsupported-language OCR text.

In V1, refined rebuild is the default. Treat these conditions as extra reasons to distrust direct OCR recomposition:

- Any automatic output is visually or semantically poor, regardless of language.
- English OCR is available but the recognized text is wrong, merged, clipped, badly wrapped, or assigned to the wrong region.
- Chinese text exists and the OCR language list does not include `chi_sim`, `chi_tra`, or another suitable CJK package.
- OCR text is mostly gibberish or unrelated fragments.
- One OCR text element contains many unrelated slide regions.
- Important labels are merged with figures, icons, chart bars, or lines.
- The recomposed PPTX preview shows severe wrapping, overlap, or clipping.
- The user needs an actually editable deck, not just a rough visual decomposition.

## 2. Output contract

For a refined rebuild, produce a predictable artifact bundle:

```text
<name>_refined_editable.pptx
<name>_refined_editable_output/
  source_pages/
  split_png_elements/
  split_png_elements.zip
  recomposed_from_elements.pptx
  visual_elements_manifest.csv
  visual_elements_manifest.json
  image_source_report.csv
  image_source_report.json
  review/
  quality_preview/
  quality_report/
```

The root PPTX and `recomposed_from_elements.pptx` may be the same file copied to two locations for convenience.

## 3. Region map before editing

Before writing the rebuild script, identify the slide regions:

- page background and decorative header/footer;
- title text and section headers;
- main panels, tables, and chart frames;
- chart axes, bars, ticks, and numeric labels;
- legends and captions;
- scientific figures, diagrams, logos, and icons;
- warning/callout boxes and footer captions.

Use the original image pixel coordinates as the source of truth. Keep the coordinate system consistent throughout the script.

## 4. What to rebuild as native PowerPoint objects

Prefer native editable objects for:

- text boxes;
- straight lines and dashed dividers;
- rectangles, rounded rectangles, pills, badges, and color blocks;
- arrows and simple symbols;
- chart axes, ticks, bars, range bands, and labels;
- table/grid borders when simple enough.

Prefer independent PNG crops for:

- photos;
- logos;
- dense scientific illustrations;
- material microstructure diagrams;
- hand-drawn or shaded icons;
- complex plots where redrawing would risk scientific errors;
- decorative typography that is not worth reconstructing.

Do not use one full-slide screenshot as the final reconstruction. Cropped PNGs must be independent movable elements.

## 5. Text reconstruction rules

Use reliable OCR text only as an aid. The source image and user-provided references are the ground truth.

For all V1 refined rebuilds:

- do not hard-convert unreliable OCR output into the final editable deck;
- manually or programmatically reconstruct readable titles, labels, descriptions, captions, coordinate/axis labels, legends, and table text as editable text boxes;
- crop complex graphics, icons, diagrams, figures, photos, and expensive-to-redraw visual regions as independent PNG assets;
- rebuild linework, color blocks, arrows, table borders, chart axes, ticks, bars, and simple shapes as native PowerPoint objects;
- export a PowerPoint preview and adjust font size, position, width, and line breaks until major overlap and clipping are resolved.

For Chinese slides without Chinese OCR support:

- manually transcribe readable text from the source image;
- use `Microsoft YaHei`, `SimHei`, or a deck-consistent CJK font;
- use conservative font sizes because PowerPoint often renders CJK larger than screenshot pixels suggest;
- prefer wider text boxes and explicit line breaks for dense labels;
- use vertical line-by-line text for narrow rotated Chinese axis labels if PowerPoint rotation causes overlap;
- record `manual_text_reconstruction` in `conversion_note`;
- if text is unreadable, keep it as PNG and mark `manual_review_required`.

For formulas, chemical notation, subscripts, or special symbols:

- use editable text if it is simple and legible;
- use Unicode subscripts where practical, such as `CO₂`, `Na₂SO₄`, and `Ca(NO₃)₂`;
- keep as PNG if symbol fidelity is uncertain.

## 6. Rebuild script pattern

Use a deterministic script with these parts:

1. constants for source path, output paths, source width/height, and slide size;
2. `px`, `py`, `pw`, `ph` helpers to convert source pixels to PowerPoint EMUs;
3. helpers for adding text, shapes, lines, arrows, pictures, and manifest rows;
4. asset cropping from the source image into `split_png_elements/<page_id>/`;
5. slide construction using native objects plus independent picture crops;
6. manifest CSV/JSON writing;
7. review overlay generation;
8. PowerPoint-rendered preview export when available;
9. diff-based quality report.

Keep the script scoped to the slide or batch being rebuilt. It is acceptable for the refined script to be task-specific because the goal is a usable editable deck, not a generic perfect vector recovery engine.

## 7. Quality loop

After generating the PPTX:

1. Open it structurally with `python-pptx` and count slides, text boxes, and picture assets.
2. Export a PNG preview from PowerPoint if available.
3. Compare the preview with the source visually.
4. Fix obvious issues:
   - text too large or wrapped badly;
   - text clipping;
   - panel border radius too large;
   - arrows or dashed lines misaligned;
   - important image crops too large, too small, or clipped.
5. Repeat until the output is usable.

Quality is judged by editability and visual fidelity together. A slightly approximate but editable text layout is better than an OCR-generated deck with wrong text.

## 8. Manifest notes for refined rebuilds

Use standard manifest fields, and add clear notes:

- `converted_to_editable_text=yes` for reconstructed editable text boxes;
- `conversion_note=manual_text_reconstruction` for manually transcribed text;
- `conversion_note=ocr_language_missing_chi_sim` when Chinese OCR was unavailable;
- `conversion_note=rectangular_crop_fallback` for complex visual PNG crops;
- `needs_manual_review=yes` only when the element is uncertain or unreadable.

For PNG crops, keep `file_name`, `relative_path`, pixel bbox, normalized bbox, and PPTX EMU coordinates.

## 9. Recommended final response

Tell the user:

- where the refined PPTX is;
- where the output bundle is;
- how many editable text boxes and picture assets were produced;
- whether OCR was limited and how text was handled;
- whether a PowerPoint preview and quality report were generated.

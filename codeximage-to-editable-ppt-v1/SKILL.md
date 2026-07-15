---
name: codeximage-to-editable-ppt-v1
description: >-
  Direct refined editable rebuild for image-based PPT/PPTX files and slide screenshots. Use when converting raster slides into usable editable PPTX: run baseline decomposition as reference, then make the primary final deck by cropping complex graphics/icons/diagrams as independent PNGs, recreating titles/labels/descriptions/axis and table text as editable text boxes, rebuilding lines/color blocks/arrows/table borders/chart primitives with native PowerPoint shapes, exporting a preview, and iterating layout. Use scripts/run_batches.py for repeated or batch work.
---

# V1 Direct Refined Image-PPT-to-Editable-PPT Skill

## Scope and positioning

This skill is designed **specifically for image-based PPT/PPTX files and slide screenshots**. In this context, an image-based PPT means each slide is mainly or entirely composed of one full-slide raster image, a screenshot, or multiple raster layers. The goal is **not** to recover original PowerPoint vector objects. The goal is to:

1. extract or render each raster slide page at high quality;
2. visually decompose the raster page into independent PNG elements;
3. identify text-only PNG elements;
4. export all PNG elements into page-level folders;
5. generate a detailed manifest;
6. rebuild a PPTX from independent PNG elements plus editable text boxes and native PowerPoint shapes.

Do **not** optimize this skill for native-object PPT recovery. This skill assumes the visible slide is a raster image that must be decomposed visually.

V1 is more opinionated than the baseline optimized skill: the direct OCR-based `recomposed_from_elements.pptx` is not the primary deliverable. It is a reference and backup artifact. The primary deliverable is a refined editable rebuild.

## Core workflow

Use this staged workflow:

### V1 direct refined rebuild contract

For every input slide, produce a refined editable rebuild as the primary result whenever practical. Do this regardless of whether the slide is English, Chinese, mixed-language, or mostly visual.

The default V1 behavior is:

1. Run baseline decomposition first to create source pages, element crops, review overlays, manifests, and a rough recomposed PPTX.
2. Inspect the baseline manifest, review overlay, and rough recomposed preview if available.
3. Do not hard-convert unreliable OCR text into the final deck. Use the source image as the visual ground truth.
4. Crop complex graphics, icons, diagrams, scientific illustrations, photos, logos, and expensive-to-redraw chart artwork as independent PNG elements.
5. Recreate titles, labels, descriptions, captions, legends, coordinate labels, axis labels, table text, and other readable text as editable PowerPoint text boxes.
6. Rebuild simple structure with native PowerPoint objects: page backgrounds, panels, rounded rectangles, straight lines, dashed dividers, arrows, color blocks, chart axes, ticks, bars, badges, table/grid borders, and simple symbols.
7. Export a PowerPoint-rendered preview of the refined PPTX.
8. Adjust font size, text-box width, position, line breaks, and object geometry until severe overlap, clipping, and ugly wrapping are resolved.
9. Write the full artifact bundle: refined PPTX, source page, split PNG elements, ZIP, manifest CSV/JSON, review overlay, quality preview, diff image, and quality report.

Only skip the refined rebuild when the user explicitly asks for automatic decomposition only, dry-run planning, or raw element extraction without PPTX polishing.

### Language and OCR capability gate

Before trusting OCR output for editable text conversion, explicitly check whether the local OCR runtime can recognize the slide language.

1. Inspect the input visually or from filenames/context for the dominant text language.
2. Check available OCR languages, for example with `pytesseract.get_languages(config="")` when using Tesseract.
3. If the slide contains Chinese text but `chi_sim` or another suitable Chinese OCR package is unavailable, mark OCR as insufficient for Chinese text.
4. In that case, still run baseline decomposition for visual element discovery, source-page export, element crops, review overlays, and manifests, but do not rely on OCR-recognized Chinese text for the final editable PPTX.
5. Prefer a refined editable rebuild: use the raster slide as the visual reference, crop complex artwork as independent PNG elements, and manually or programmatically recreate the Chinese text as editable PowerPoint text boxes.
6. Record the reason in the manifest and quality notes, for example `ocr_language_missing_chi_sim`, `manual_text_reconstruction`, or `ocr_gibberish_replaced`.

Failure triggers that require refined rebuild instead of direct OCR recomposition:

- recognized text is mostly gibberish, romanized fragments, punctuation, or unrelated English;
- multiple unrelated text zones are merged into one large OCR element;
- text is merged with icons, charts, tables, or decorative backgrounds;
- Chinese text is present but the OCR runtime only exposes `eng`/`osd`;
- PowerPoint preview shows severe wrapping, overlap, or clipped editable text;
- chart labels, axis labels, captions, or legends are too small or dense for reliable OCR.

### Refined editable rebuild workflow as the default final output

For any language, including English and Chinese, do not treat the direct OCR-based recomposed PPTX as the final result by default. Produce two outputs whenever practical:

1. the baseline automated output from `decompose_visual_elements.py`;
2. a curated `*_refined_editable.pptx` and matching `*_refined_editable_output/` folder.

The refined rebuild must follow this pattern:

1. Run baseline decomposition first with review outputs enabled.
2. Inspect `visual_elements_manifest.csv`, the detected-element overlay, and, when possible, a PowerPoint-exported PNG preview.
3. Define the slide's visual regions: title/header, panels, tables, charts, legends, icons, figures, captions, footers, and decorative lines.
4. Do not hard-convert unreliable OCR text into editable text. Use the source image as ground truth, and use OCR only as an aid when it is visibly correct.
5. Crop complex graphics, icons, diagrams, photos, logos, scientific illustrations, and expensive-to-redraw chart artwork as independent PNG elements.
6. Rebuild titles, labels, captions, descriptions, coordinate/axis labels, legends, table text, and other readable text as editable PowerPoint text boxes.
7. Rebuild simple design structure as editable PowerPoint objects: backgrounds, panels, rounded rectangles, lines, dashed dividers, arrows, chart axes, bars, color blocks, table/grid borders, badges, labels, and simple symbols.
8. Preserve wording, line breaks, approximate alignment, color, and hierarchy.
9. For Chinese text, default to a common CJK-capable font such as `Microsoft YaHei`, `SimHei`, or the closest font already used by the deck. Use conservative font sizes because PowerPoint export often renders CJK text larger than screenshot pixels imply.
10. For English text, still inspect OCR output and preview rendering. If OCR text is wrong, merged, clipped, or badly wrapped, manually or programmatically reconstruct it as editable text from the source image.
11. Export a rendered preview of the rebuilt PPTX, inspect it visually, then adjust font scale, text-box width, line breaks, and object positions until there is no severe overlap or clipping.
12. Write the same artifact family as the baseline workflow: source page, split PNG elements, ZIP, CSV/JSON manifest, review overlay, quality preview, diff image, and quality report.
13. Name the refined deck and folder predictably, for example `slide_08_refined_editable.pptx` and `slide_08_refined_editable_output/`.

The refined rebuild is the primary V1 deliverable. Do not present a direct OCR-based PPTX as the primary result unless the user explicitly asks for automatic output only.

### Recommended implementation pattern for refined rebuilds

When using Python and `python-pptx`, use a deterministic per-slide rebuild script so that edits can be repeated and improved.

The script should:

- define the source image size and target PowerPoint slide size;
- provide pixel-to-EMU conversion helpers;
- crop complex visual assets into `split_png_elements/<page_id>/`;
- insert those crops as independent pictures, not as one full-slide screenshot;
- create editable text boxes for reliable or manually reconstructed text;
- create editable PowerPoint shapes for panels, dividers, arrows, chart bars, badges, and simple icons when feasible;
- write `visual_elements_manifest.csv` and `.json` with the same minimum fields as the normal workflow;
- create a review overlay with element bounding boxes;
- export or accept a PowerPoint-rendered preview and write a lightweight diff-based quality report;
- keep the script beside the task output only when it helps reproducibility.

Use manual text reconstruction only for text that can be read from the source image or from a user-provided reference. If text is unreadable, keep it as PNG and mark `manual_review_required`.

For the detailed checklist and implementation pattern, read `references/refined_rebuild_workflow.md` before performing a refined rebuild.

### Batch processing contract

When the user asks to process multiple files, a folder, repeated batches, or any task that may require more than one decomposition run, do not rely on conversation memory to remember batch parameters or manually loop over commands.

Use `scripts/run_batches.py` as the batch entrypoint. If the requested batch behavior is not supported yet, modify `scripts/run_batches.py` first, then run it. Keep batch behavior in the script rather than in ad hoc chat instructions.

The batch runner must:

1. collect input files from explicit file paths or input directories;
2. group inputs into batches of two files by default (`--batch-size 2`);
3. create one stable batch folder per batch group, such as `batch_001_2_items/`;
4. automatically write exactly one fixed-template `batch_instruction.md` inside each batch folder before execution;
5. write the batch instruction with the batch rules, the input list, one shared parameter block, one command per input item, expected outputs, and status fields;
6. execute item commands through the global worker pool rather than sequentially;
7. write `batch_run.log` for each batch;
8. write `batch_summary.csv` and `batch_summary.json` at the batch root.
9. after all item runs succeed, merge every item `recomposed_from_elements.pptx` into `merged_recomposed_from_elements.pptx` at the batch root, strictly following the original input order, not parallel completion order.

Default parallelism:

- `--batch-size 2`: put two input files in each batch.
- `--batch-workers 2`: keep at most two batches active at the same time.
- `--workers`: max concurrent item processes. If omitted, compute it as `batch-size * batch-workers`, so the default is 4 concurrent item processes.
- Preserve batch-level organization even when item execution is parallel: each batch still has one `batch_instruction.md` and one `batch_run.log`.
- Final merge is enabled by default. Use `--no-merge` only when the user explicitly does not want a combined deck.

Default command pattern:

```bash
python scripts/run_batches.py input_folder_or_files \
  --outdir batch_output \
  --batch-size 2 \
  --batch-workers 2 \
  --dpi 300 \
  --granularity fine \
  --ocr \
  --ocr-lang chi_sim+eng \
  --ocr-confidence-threshold 75 \
  --editable-text \
  --review \
  --quality-check
```

For planning or QA without running decomposition, use `--dry-run`. This still generates every `batch_instruction.md` and the batch summary.

Final merged output:

- `merged_recomposed_from_elements.pptx`: combined deck at the batch root.
- `merge_report.csv` and `merge_report.json`: ordered merge audit showing each input item, item PPTX path, manifest path, inclusion status, and slide count.

### Stage 1 - Input inspection and page source extraction

- Accept `.ppt`, `.pptx`, `.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`, `.tif`, `.tiff`.
- If input is PPT/PPTX, detect whether each slide contains a full-slide embedded raster image.
- If a full-slide embedded raster image is available, extract the original embedded image directly from the PPTX package whenever possible. This avoids blur caused by screenshotting or rendering.
- If the original embedded image cannot be extracted, or the slide has multiple raster layers, render the slide at high DPI, normally 240-300 DPI.
- Produce a `source_pages/` folder and an `image_source_report.csv` describing each page source.

Recommended source type values:

- `full_slide_image_extracted`
- `multi_image_slide_rendered`
- `rendered_slide_image`
- `input_image`
- `unknown_raster_layout`

### Stage 2 - Text-first detection

Perform text detection before general visual segmentation so text is not incorrectly merged with nearby icons, lines, or backgrounds.

Detect and export text-only PNG elements for:

- title text
- body text
- headers and footers
- captions
- legend text
- axis labels
- table text
- callout text
- notes
- formula labels
- formula numbers

Every text-only element must still be exported as a PNG backup, even if it is later converted into an editable text box in the recomposed PPTX.

### Stage 3 - Visual element segmentation

Detect and export independent non-text visual objects, including:

- backgrounds and background decorations
- logos
- icons
- photos and figures
- illustrations
- lines and dividers
- geometric shapes
- color blocks
- shadows
- formulas and special symbols
- tables
- charts
- legends and chart areas
- mixed image-text regions that cannot be reliably separated

A full-page background may be exported only when it is a pure background layer or solid background representation. Do not satisfy the request by exporting only whole-page screenshots.

### Stage 4 - Manifest generation

Generate both CSV and JSON manifest files whenever possible. The manifest must be usable as the source of truth for PPTX recomposition.

Minimum fields:

- `slide_id` or `image_id`
- `element_id`
- `file_name`
- `relative_path`
- `element_type`
- `is_text_only`
- `recognized_text`
- `OCR_confidence`
- `x`
- `y`
- `width`
- `height`
- `z_order`
- `opacity`
- `rotation`
- `crop_mask_type`
- `transparent_background`
- `alpha_note`
- `shadow_effect`
- `converted_to_editable_text`
- `conversion_note`

Recommended additional fields:

- `page_source_type`
- `source_page_path`
- `source_width`
- `source_height`
- `bbox_pixels`
- `bbox_normalized`
- `pptx_left_emu`
- `pptx_top_emu`
- `pptx_width_emu`
- `pptx_height_emu`
- `dominant_color`
- `font_family_guess`
- `font_size_guess`
- `font_color_guess`
- `text_alignment_guess`
- `language_guess`
- `split_confidence`
- `parent_group_id`
- `needs_manual_review`

### Stage 5 - Baseline PPTX recomposition

After PNG elements and the manifest are generated, rebuild a baseline PPTX. In V1, this baseline PPTX is a reference/backup artifact, not the primary final output.

- Insert non-text PNG elements using recorded position, size, z-order, opacity, rotation, and crop information.
- For elements marked as text-only, use `recognized_text`, original coordinates, estimated font size, font color, alignment, and layer order to insert an editable PowerPoint text box instead of the PNG image when OCR confidence is sufficient.
- Keep original text PNG files in the ZIP as visual backups.
- If a text element cannot be reliably recognized or visually reconstructed, keep it as PNG and record a reason.

After this baseline PPTX is created, proceed to the V1 refined rebuild contract above unless the user explicitly requested automatic output only.

Common reasons:

- `low_ocr_confidence`
- `mixed_text_and_graphics`
- `decorative_typography`
- `complex_formula`
- `symbol_recognition_uncertain`
- `text_too_small_or_blurred`
- `overlapping_elements`
- `manual_review_required`

### Stage 6 - Review and quality report

When possible, generate:

- review overlay images with element boxes and element IDs;
- a recomposed PPTX render;
- diff images comparing the source page and recomposed page;
- a quality report CSV.

Suggested outputs:

```text
output/
|-- source_pages/
|-- split_png_elements/
|-- split_png_elements.zip
|-- recomposed_from_elements.pptx
|-- visual_elements_manifest.csv
|-- visual_elements_manifest.json
|-- image_source_report.csv
|-- review/
|   |-- slide01_detected_elements_overlay.png
|   `-- ...
`-- quality_report/
    |-- slide01_original.png
    |-- slide01_recomposed.png
    |-- slide01_diff.png
    `-- quality_report.csv
```

## File and folder naming rules

Create a root folder such as `split_png_elements/`. Inside it, create one subfolder per page/image:

```text
split_png_elements/
|-- slide01/
|   |-- slide01_background_01.png
|   |-- slide01_decoration_01.png
|   |-- slide01_logo_01.png
|   |-- slide01_title_txt_01.png
|   |-- slide01_author_txt_01.png
|   |-- slide01_journal_txt_01.png
|   |-- slide01_doi_txt_01.png
|   |-- slide01_keyword_txt_01.png
|   |-- slide01_icon_01.png
|   |-- slide01_icon_02.png
|   |-- slide01_line_01.png
|   `-- slide01_figure_01.png
`-- slide02/
    |-- slide02_background_01.png
    |-- slide02_title_txt_01.png
    |-- slide02_chart_01.png
    |-- slide02_axis_label_txt_01.png
    |-- slide02_legend_txt_01.png
    |-- slide02_table_01.png
    `-- slide02_icon_01.png
```

For independent image inputs, use `image01/`, `image02/`, etc.

### Naming convention

Use:

```text
<page_id>_<element_type>_<sequence>.png
```

If an element is pure text, the filename must include `txt`, e.g.:

- `slide01_title_txt_01.png`
- `slide01_body_txt_02.png`
- `slide01_caption_txt_01.png`
- `slide01_axis_label_txt_01.png`
- `slide01_formula_label_txt_01.png`

Avoid naming a file simply `txt.png`; use meaningful text categories.

Recommended `element_type` values:

- `background`
- `decoration`
- `logo`
- `icon`
- `figure`
- `photo`
- `illustration`
- `line`
- `divider`
- `arrow`
- `shape`
- `simple_shape`
- `block`
- `button`
- `badge`
- `shadow`
- `title_txt`
- `body_txt`
- `caption_txt`
- `label_txt`
- `axis_label_txt`
- `legend_txt`
- `table`
- `chart`
- `chart_plot_area`
- `formula`
- `formula_label_txt`
- `symbol`
- `mixed`
- `unknown`

## Cropping and transparency rules

- PNG elements should use transparent backgrounds whenever possible.
- Crop each element to its minimum visible bounds. Cropping is not limited to rectangles.
- For circular, elliptical, curved, or irregular elements, use the PNG alpha channel as a mask following the true visible contour so that the outside area is transparent.
- Avoid preserving page background, white fill, rectangular backing, or extra whitespace outside the real visible element contour.
- Preserve smooth anti-aliased edges and avoid white halos, black halos, jagged edges, or background remnants.
- Record the crop shape or mask type as one of:
  - `rectangle`
  - `circle`
  - `ellipse`
  - `irregular_mask`
  - `full_background`
- Record `transparent_background` as `yes` or `no`.
- Record `alpha_note` for how the alpha channel was produced or why transparency was not reliable.

PNG canvases are rectangular, but alpha transparency should preserve non-rectangular shapes.

### Alpha transparency handling

Because image-based PPT slides are flattened raster images, original transparency information is usually unavailable. The workflow must infer transparency from pixels.

For text, icons, logos, simple shapes, lines, and dividers, exported PNG elements should use an alpha mask whenever possible, so that the non-element area becomes transparent.

For elements recognized as `icon`, `logo`, `symbol`, `button`, `badge`, `arrow`, `line`, `divider`, `shape`, or `simple_shape`, transparent background is mandatory unless a fill color, circular outline, button plate, badge backing, or similar base is part of the element's own visible design. Remove the external slide/page background and keep only the intended visible element plus its legitimate internal backing.

For photos, complex illustrations, charts, tables, screenshots, and background decorations, the workflow may preserve the original rectangular crop if alpha extraction would damage visual fidelity.

If a mandatory-transparent simple element cannot be reliably transparentized because of complex background, low-confidence mask, blended edge colors, or shadow dependence on the background, keep the safer crop, mark `transparent_background` as `no`, set `needs_manual_review` to `yes`, and explain the reason in `alpha_note`.

If any element is exported with a non-transparent background, the manifest must record the reason in `alpha_note`. Use one or more of these values when applicable:

- `low_mask_confidence`
- `complex_background`
- `edge_color_blending`
- `anti_aliased_text_edges`
- `shadow_requires_background`
- `photo_or_chart_region`
- `rectangular_crop_fallback`
- `transparent_required_but_unreliable`

Scripts implementing this skill should support:

```text
--alpha-mode conservative|auto|aggressive
--force-transparent-for text,icon,logo,symbol,button,badge,arrow,line,divider,shape,simple_shape
```

`--alpha-mode conservative` should avoid uncertain masking and preserve rectangular crops more often. `--alpha-mode auto` should infer transparency for likely foreground-only elements while preserving complex regions. `--alpha-mode aggressive` may use stronger background removal for text, icons, lines, logos, and simple shapes, but must mark uncertain results for review.

`--force-transparent-for` should accept a comma-separated list of element classes and attempt alpha-mask extraction for those classes even when the global alpha mode is conservative. When forced masking is attempted but not reliable, keep the safer rectangular crop and write the failure reason to `alpha_note`.

## Granularity modes

Use a configurable granularity mode:

- `coarse`: title/body/figure/table/chart regions, minimal fragmentation.
- `normal`: common visual objects: text blocks, icons, lines, images, tables, charts.
- `fine`: extra labels, legends, chart annotations, small icons, thin dividers.
- `ultra`: smallest practical visible components; may require manual review.

## Quality expectations and limitations

- This skill cannot perfectly reconstruct original vector objects from a raster image.
- It should maximize visual fidelity, PNG element separation, and editable text conversion.
- Use review overlays and diff reports to identify pages needing manual refinement.
- For complex formulas, decorative typography, low-resolution text, or tightly overlapping objects, keep a PNG fallback and record the reason.

# Label Plate Composer

Browser tool that builds plate PDFs from:

1. **Outline label PDFs** (one per style)
2. **Ratio master Excel** (size / EAN — or UPC / Colour — and A–D plate counts; columns found by header)
3. **Layout master PDF** (slots auto-detected from ~30×74 or 30×85 mm cut boxes, square or rounded)

## How it works

- **Plate A/B/C/D = pages 1–4**
- Excel cell values are how many of that size go on each plate
- Labels are placed in production order: Excel row order, repeated by count
- **Layout slots are detected dynamically** from the uploaded master (portrait + landscape cut rectangles), sorted top→bottom, left→right
- Excel totals per plate must equal the detected ups count
- Landscape slots receive artwork rotated 90° clockwise
- Label artwork is placed as vectors from the outline PDFs (not flattened to images)
- Each size is cropped to its own cut box before stamping, so Illustrator does not keep sibling labels behind the art

## Run locally

PDF.js needs HTTP (not `file://`):

```bash
cd label-composer
python -m http.server 8765
```

Open `http://localhost:8765`

## Sample files in this folder

- `outlines/` — style outline PDFs
- `ratio-master.xlsx`
- `layout-master.pdf`
- `sample-labels.pdf` — expected final result reference
- `Test -4 Urgent outline font/` — M&S 30×85 mm set (outlined fonts, rounded cut paths, UPC/Colour Excel) with its expected `Final …pdf`

## Outlined-font outlines (M&S)

When an outline PDF has no per-label clip groups (text converted to outlines), labels are split by geometry:
each stroked label-sized path is a cut box, and every drawn object goes to the cut box that contains it.
Sheet headers outside the boxes are dropped. Labels keep the outline's drawing order, which follows the
Excel (these sheets are laid out column by column). Rounded master cut paths are redrawn as-is, and a
"…Plate-A- 45 sheet _ Date : …, Done by- …" master line is refilled in place for each plate.

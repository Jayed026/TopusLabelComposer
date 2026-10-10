# Label Plate Composer

Browser tool that builds plate PDFs from:

1. **Outline label PDFs** (one per style)
2. **Ratio master Excel** (size / EAN / A–D plate counts)
3. **Layout master PDF** (slots auto-detected from ~30×74 mm cut boxes)

## How it works

- **Plate A/B/C/D = pages 1–4**
- Excel cell values are how many of that size go on each plate
- Labels are placed in production order: Excel row order, repeated by count
- **Layout slots are detected dynamically** from the uploaded master (portrait + landscape cut rectangles), sorted top→bottom, left→right
- Excel totals per plate must equal the detected ups count
- Landscape slots receive artwork rotated 90° clockwise

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

"""Validate plate composition against sample-labels.pdf"""
import io
from pathlib import Path

import openpyxl
import pymupdf
from PIL import Image
from pyzbar.pyzbar import decode

BASE = Path(r"c:\Bs23\Project\Learning\Topu\Test\label-composer")
OUTLINES = BASE / "outlines"
LAYOUT = BASE / "layout-master.pdf"
SAMPLE = BASE / "sample-labels.pdf"
EXCEL = BASE / "ratio-master.xlsx"
OUT = BASE / "test-composed.pdf"

PAGE_W, PAGE_H = 1029.6, 558.0
LABEL_W, LABEL_H = 85.039, 209.762
PLATES = ["A", "B", "C", "D"]


def parse_excel():
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    ws = wb.active
    rows = []
    styles = []
    style = None
    impressions = {"A": 0, "B": 0, "C": 0, "D": 0}
    data = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(data):
        if i == 0:
            continue
        if row[1] is None and any(isinstance(v, str) and "Imp_" in v for v in row[3:7] if v):
            for idx, plate in enumerate(PLATES):
                m = str(row[3 + idx] or "")
                impressions[plate] = int("".join(ch for ch in m if ch.isdigit()) or 0)
            continue
        if row[0]:
            style = row[0]
            styles.append(style)
        if not style or not row[1]:
            continue
        rows.append(
            {
                "style": style,
                "size": str(row[1]),
                "ean": str(int(row[2])) if row[2] else "",
                "A": int(row[3] or 0),
                "B": int(row[4] or 0),
                "C": int(row[5] or 0),
                "D": int(row[6] or 0),
            }
        )
    return rows, styles, impressions


def frames_for(doc_page):
    frames = []
    for d in doc_page.get_drawings():
        r = d["rect"]
        w, h = abs(r.width), abs(r.height)
        if 80 <= w <= 90 and 200 <= h <= 220:
            frames.append(pymupdf.Rect(r))
    uniq = []
    for r in sorted(frames, key=lambda r: (-round(r.y0, 0), round(r.x0, 0))):
        if not any(abs(r.x0 - u.x0) < 1 and abs(r.y0 - u.y0) < 1 for u in uniq):
            uniq.append(r)
    # sort top-to-bottom, left-to-right in visual space
    uniq.sort(key=lambda r: (round(r.y0 / 5), round(r.x0 / 5)))
    return uniq


def extract_labels(rows):
    by_style = {}
    for r in rows:
        by_style.setdefault(r["style"], []).append(r)
    label_map = {}
    for pdf in sorted(OUTLINES.glob("*.pdf")):
        style = None
        for s in by_style:
            if s in pdf.name.upper():
                style = s
                break
        assert style, pdf.name
        doc = pymupdf.open(pdf)
        page = doc[0]
        fr = frames_for(page)
        expected = by_style[style]
        assert len(fr) == len(expected), (style, len(fr), len(expected))
        for rect, row in zip(fr, expected):
            pix = page.get_pixmap(matrix=pymupdf.Matrix(2.5, 2.5), clip=rect, alpha=False)
            label_map[(row["style"], row["ean"])] = pix
    return label_map


def slots():
    vert_xs = [127.0, 212.0, 297.1, 382.1, 467.2, 552.2, 648.6, 733.6, 818.7, 903.7, 988.7]
    horiz_xs = [251.7, 461.5, 779.0, 988.7]
    out = []
    for x in horiz_xs:
        x0 = x - LABEL_H
        y0 = 534.1 - LABEL_W
        out.append(("H", pymupdf.Rect(x0, PAGE_H - 534.1, x, PAGE_H - y0)))
    for y_top in [443.4, 233.6]:
        for x in vert_xs:
            x0 = x - LABEL_W
            y0 = y_top - LABEL_H
            out.append(("V", pymupdf.Rect(x0, PAGE_H - y_top, x, PAGE_H - y0)))
    return out


def compose():
    rows, styles, impressions = parse_excel()
    label_map = extract_labels(rows)
    queues = {p: [] for p in PLATES}
    for r in rows:
        art = label_map[(r["style"], r["ean"])]
        for p in PLATES:
            queues[p].extend([art] * r[p])

    layout = pymupdf.open(LAYOUT)
    out = pymupdf.open()
    slot_list = slots()
    wo = styles[0] + "+" + "+".join(s.replace("BD0148", "") for s in styles[1:])

    for plate in PLATES:
        out.insert_pdf(layout, from_page=0, to_page=0)
        page = out[-1]
        # whiteout right text
        page.draw_rect(pymupdf.Rect(PAGE_W - 42, 20, PAGE_W - 12, PAGE_H - 20), color=(1, 1, 1), fill=(1, 1, 1))
        text = f"Plate {plate} - {impressions[plate]} _ W/O _{wo}_ Date : 22.04.2026, Done by- Test"
        page.insert_text((PAGE_W - 28, 40), text, fontsize=8, rotate=90)

        for i, (orient, rect) in enumerate(slot_list):
            pix = queues[plate][i]
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            if orient == "H":
                # rotate 90 CW so punch hole goes to the right
                img = img.rotate(-90, expand=True)
            bio = io.BytesIO()
            img.save(bio, format="PNG")
            page.insert_image(rect, stream=bio.getvalue(), keep_proportion=False)

    out.save(OUT)
    print("wrote", OUT)

    # compare barcodes plate A
    sample = pymupdf.open(SAMPLE)
    composed = pymupdf.open(OUT)
    mism = 0
    for plate_i, plate in enumerate(PLATES):
        for i, (orient, rect) in enumerate(slot_list):
            for doc, tag in [(sample, "sample"), (composed, "ours")]:
                page = doc[plate_i]
                clip = pymupdf.Rect(rect.x0 - 1, rect.y0 - 1, rect.x1 + 1, rect.y1 + 1)
                pix = page.get_pixmap(matrix=pymupdf.Matrix(2.5, 2.5), clip=clip)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                codes = decode(img) or decode(img.rotate(90, expand=True)) or decode(img.rotate(-90, expand=True))
                val = codes[0].data.decode() if codes else None
                if tag == "sample":
                    s_val = val
                else:
                    if val != s_val:
                        mism += 1
                        if mism <= 12:
                            print(f"mismatch {plate} slot {i+1}: sample={s_val} ours={val}")
    print("mismatches", mism)


if __name__ == "__main__":
    compose()

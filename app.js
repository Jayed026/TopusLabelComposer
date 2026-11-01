/* global pdfjsLib, PDFLib, XLSX */

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const PLATES = ["A", "B", "C", "D"];
/** 30×74 mm / 30×85 mm label short/long sides in PDF points (±tolerance). */
const LABEL_SHORT_MIN = 78;
const LABEL_SHORT_MAX = 95;
const LABEL_LONG_MIN = 195;
const LABEL_LONG_MAX = 260;

const els = {
  form: document.getElementById("compose-form"),
  labelFiles: document.getElementById("label-files"),
  excelFile: document.getElementById("excel-file"),
  layoutFile: document.getElementById("layout-file"),
  dateText: document.getElementById("date-text"),
  doneBy: document.getElementById("done-by"),
  generateBtn: document.getElementById("generate-btn"),
  resetBtn: document.getElementById("reset-btn"),
  progressBar: document.getElementById("progress-bar"),
  statusText: document.getElementById("status-text"),
  log: document.getElementById("log"),
  previewSection: document.getElementById("preview-section"),
  plateSummary: document.getElementById("plate-summary"),
  downloadLink: document.getElementById("download-link"),
};

els.dateText.value = formatDate(new Date());

els.form.addEventListener("submit", async (event) => {
  event.preventDefault();
  await runCompose();
});

els.resetBtn.addEventListener("click", () => {
  els.form.reset();
  els.dateText.value = formatDate(new Date());
  els.doneBy.value = "Composer";
  els.log.innerHTML = "";
  els.plateSummary.innerHTML = "";
  els.previewSection.classList.add("hidden");
  els.downloadLink.classList.add("hidden");
  setProgress(0, "Waiting for files…");
});

async function runCompose() {
  const labelFiles = [...els.labelFiles.files];
  const excelFile = els.excelFile.files[0];
  const layoutFile = els.layoutFile.files[0];

  if (!labelFiles.length || !excelFile || !layoutFile) {
    setStatus("Please provide label PDFs, Excel, and layout PDF.");
    return;
  }

  els.generateBtn.disabled = true;
  els.log.innerHTML = "";
  els.previewSection.classList.add("hidden");
  els.downloadLink.classList.add("hidden");

  try {
    setProgress(5, "Reading Excel ratio master…");
    const ratio = await parseRatioExcel(excelFile);
    log(`Styles in Excel: ${ratio.styles.join(", ")}`);
    log(
      `Impressions — A:${ratio.impressions.A} B:${ratio.impressions.B} C:${ratio.impressions.C} D:${ratio.impressions.D}`
    );

    setProgress(12, "Detecting slots from layout master…");
    const layout = await detectLayoutFromMaster(layoutFile);
    log(
      `Layout ${layout.pageW.toFixed(1)}×${layout.pageH.toFixed(1)} pts · ${layout.slots.length} slots ` +
        `(${layout.slots.filter((s) => s.orient === "H").length} H / ${layout.slots.filter((s) => s.orient === "V").length} V)`
    );

    setProgress(20, "Reading vector labels from outline PDFs…");
    const labelMap = await extractLabelsFromOutlines(labelFiles, ratio.rows);
    log(`Matched label artworks: ${Object.keys(labelMap).length}`);

    setProgress(45, "Building plate queues from Excel…");
    const plateQueues = buildPlateQueues(ratio.rows, labelMap);
    const slotCount = layout.slots.length;
    renderPlateSummary(plateQueues, ratio.impressions, slotCount);

    for (const plate of PLATES) {
      if (plateQueues[plate].length !== slotCount) {
        throw new Error(
          `Plate ${plate} has ${plateQueues[plate].length} labels; layout has ${slotCount} slots. ` +
            `Excel totals must match the master ups count.`
        );
      }
    }

    setProgress(55, "Composing final PDF pages…");
    const pdfBytes = await composeFinalPdf({
      layoutFile,
      layout,
      plateQueues,
      impressions: ratio.impressions,
      styles: ratio.styles,
      dateText: els.dateText.value.trim() || formatDate(new Date()),
      doneBy: els.doneBy.value.trim() || "Composer",
    });

    setProgress(95, "Preparing download…");
    const blob = new Blob([pdfBytes], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    els.downloadLink.href = url;
    els.downloadLink.download = buildOutputName(ratio.styles);
    els.downloadLink.classList.remove("hidden");
    els.previewSection.classList.remove("hidden");

    setProgress(100, "Done — download your composed PDF.");
  } catch (err) {
    console.error(err);
    setProgress(0, `Error: ${err.message}`);
    log(String(err.stack || err.message));
  } finally {
    els.generateBtn.disabled = false;
  }
}

/**
 * Detect label slots from the uploaded master layout PDF.
 * Finds ~30×74 mm cut boxes (portrait or landscape) and sorts top→bottom, left→right.
 */
async function detectLayoutFromMaster(file) {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
  const page = await pdf.getPage(1);
  const viewport = page.getViewport({ scale: 1 });
  const pageW = viewport.width;
  const pageH = viewport.height;

  let raw = await extractLabelSizedRectsFromPdf(buffer, pageH, { bothOrients: true });
  if (!raw.length) {
    raw = await extractLabelSizedRectsViaOps(buffer, pageH, { bothOrients: true });
  }
  // Rounded / path-drawn cut boxes (no `re`): scan paths with their transforms.
  if (!raw.length) {
    const doc = await PDFLib.PDFDocument.load(buffer.slice(0));
    raw = extractLabelSizedPathsFromContent(getPageContentText(doc.getPages()[0]));
  }
  // Last resort: scan full file bytes (uncompressed content streams)
  if (!raw.length) {
    const text = new TextDecoder("latin1").decode(new Uint8Array(buffer));
    const frames = [];
    collectLabelSizedRects(text, pageH, frames, { bothOrients: true });
    raw = dedupeFrames(frames);
  }

  if (!raw.length) {
    throw new Error(
      "No label slots found in layout master. The master must contain ~30×74 mm cut rectangles."
    );
  }

  // PDF bottom-left y descending ≈ top of page first, then left→right.
  raw.sort((a, b) => {
    const rowDiff = Math.round(b.y / 5) - Math.round(a.y / 5);
    if (rowDiff !== 0) return rowDiff;
    return a.x - b.x;
  });

  const slots = raw.map((f) => {
    const orient = f.w > f.h ? "H" : "V";
    return {
      orient,
      rect: {
        pdfX: f.x,
        pdfY: f.y,
        w: f.w,
        h: f.h,
      },
      // Original cut path (rounded corners etc.) when detected from paths.
      cut: f.cut || null,
    };
  });

  return { pageW, pageH, slots };
}

async function parseRatioExcel(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });

  if (!matrix.length) throw new Error("Excel sheet is empty.");

  const rows = [];
  const styles = [];
  let currentStyle = null;
  let impressions = { A: 0, B: 0, C: 0, D: 0 };
  const cols = detectRatioColumns(matrix[0] || []);
  const cell = (row, col) => (col >= 0 && row[col] != null ? String(row[col]).trim() : "");

  for (let i = 1; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const col0 = cell(row, 0);
    const size = cell(row, cols.size);
    const ean = cell(row, cols.ean).replace(/\D/g, "");

    // Footer impression row: Imp_142 sheet, Imp_ 20 sheet, etc.
    if (!size && !ean && PLATES.some((p) => /Imp_/i.test(cell(row, cols[p])))) {
      impressions = {
        A: parseImp(row[cols.A]),
        B: parseImp(row[cols.B]),
        C: parseImp(row[cols.C]),
        D: parseImp(row[cols.D]),
      };
      continue;
    }

    if (col0) {
      currentStyle = col0;
      if (!styles.includes(currentStyle)) styles.push(currentStyle);
    }
    if (!currentStyle || (!size && !ean)) continue;

    rows.push({
      style: currentStyle,
      size,
      ean,
      colour: cell(row, cols.colour),
      A: toCount(row[cols.A]),
      B: toCount(row[cols.B]),
      C: toCount(row[cols.C]),
      D: toCount(row[cols.D]),
    });
  }

  if (!rows.length) throw new Error("No size rows found in Excel.");
  return { rows, styles, impressions };
}

/**
 * Map header names to columns. Supports "Size / EAN Code" (Dunnes) and
 * "UPC / Colour" (M&S) sheets; falls back to the legacy fixed positions.
 */
function detectRatioColumns(header) {
  const names = header.map((v) => String(v ?? "").trim().toLowerCase());
  const find = (re) => names.findIndex((v) => re.test(v));
  const cols = {
    size: find(/size/),
    ean: find(/ean|upc|barcode/),
    colour: find(/colou?r/),
  };
  for (const p of PLATES) {
    cols[p] = find(new RegExp(`^(${p}\\s*plate|plate\\s*${p})$`, "i"));
  }
  if (PLATES.some((p) => cols[p] < 0)) PLATES.forEach((p, i) => (cols[p] = 3 + i));
  if (cols.size < 0 && cols.ean < 0) {
    cols.size = 1;
    cols.ean = 2;
  }
  return cols;
}

/** Human label for an Excel row: size if present, else UPC/EAN. */
function rowName(row) {
  return row.size || row.ean;
}

function parseImp(value) {
  const m = String(value || "").match(/Imp[_\s-]*(\d+)/i);
  return m ? Number(m[1]) : 0;
}

function toCount(value) {
  const n = Number(value);
  return Number.isFinite(n) && n > 0 ? n : 0;
}

async function extractLabelsFromOutlines(files, excelRows) {
  const byStyle = groupRowsByStyle(excelRows);
  const labelMap = {};

  let index = 0;
  for (const file of files) {
    index += 1;
    const style = detectStyleFromName(file.name);
    if (!style) {
      throw new Error(`Cannot detect style ID in filename: ${file.name}`);
    }
    const expected = byStyle[style];
    if (!expected || !expected.length) {
      throw new Error(`Style ${style} from ${file.name} is not in the Excel file.`);
    }

    setProgress(
      15 + Math.round((index / files.length) * 25),
      `Reading outline ${index}/${files.length}: ${style}`
    );

    const buffer = await file.arrayBuffer();
    const blocks = await extractCroppedLabelBlocks(buffer);
    if (blocks.length !== expected.length) {
      throw new Error(
        `${style}: found ${blocks.length} label frames but Excel has ${expected.length} sizes.`
      );
    }

    blocks.forEach((block, i) => {
      const row = expected[i];
      const key = labelKey(row.style, row.ean || row.size);
      labelMap[key] = {
        style: row.style,
        size: row.size,
        ean: row.ean,
        fileName: file.name,
        pdfBytes: buffer,
        clip: { x: block.x, y: block.y, w: block.w, h: block.h },
        content: block.content,
      };
    });

    log(`${style}: cropped ${blocks.length} vector sizes from ${file.name}`);
  }

  // Ensure every excel row with any plate qty has artwork
  for (const row of excelRows) {
    if (row.A + row.B + row.C + row.D === 0) continue;
    const key = labelKey(row.style, row.ean || row.size);
    if (!labelMap[key]) {
      throw new Error(`Missing artwork for ${row.style} / ${rowName(row)} (EAN ${row.ean}).`);
    }
  }

  return labelMap;
}

function groupRowsByStyle(rows) {
  const map = {};
  for (const row of rows) {
    if (!map[row.style]) map[row.style] = [];
    map[row.style].push(row);
  }
  return map;
}

function detectStyleFromName(name) {
  const decoded = decodeURIComponent(String(name));
  const m = decoded.match(/(BD\d{8,}W)/i);
  return m ? m[1].toUpperCase() : null;
}

function labelKey(style, eanOrSize) {
  return `${style}::${eanOrSize}`;
}

/**
 * Pull each label's own content stream from an outline PDF (not the whole sheet).
 * Rects are PDF user space (origin at the bottom-left), top of page first.
 */
async function extractCroppedLabelBlocks(buffer) {
  const { PDFDocument } = PDFLib;
  const src = await PDFDocument.load(buffer.slice(0));
  const srcPage = src.getPages()[0];
  const contentText = getPageContentText(srcPage);
  let blocks = extractLabelContentBlocks(contentText);
  // Outlined-font artwork (no per-label clip groups): split by geometry instead.
  // Keep drawing order there — those sheets may be laid out column-major
  // (top→bottom, then next column) and their drawing order follows the Excel.
  if (!blocks.length) {
    blocks = extractLabelBlocksBySpatialScan(contentText, makeXObjectBoxResolver(srcPage));
    if (!blocks.length) throw new Error("Could not detect label frames in outline PDF.");
    return blocks;
  }
  // sortX/sortY are the original page positions (clip is local 0..w x 0..h).
  blocks.sort((a, b) => {
    const rowDiff = Math.round(b.sortY / 8) - Math.round(a.sortY / 8);
    if (rowDiff !== 0) return rowDiff;
    return a.sortX - b.sortX;
  });
  return blocks;
}

function bytesToLatin1(bytes) {
  let s = "";
  for (let i = 0; i < bytes.length; i++) s += String.fromCharCode(bytes[i]);
  return s;
}

function latin1ToBytes(text) {
  const out = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) out[i] = text.charCodeAt(i) & 0xff;
  return out;
}

function rectsIntersect(a, b, pad = 2) {
  return !(
    a.x + a.w + pad < b.x ||
    b.x + b.w + pad < a.x ||
    a.y + a.h + pad < b.y ||
    b.y + b.h + pad < a.y
  );
}

function getPageContentText(page) {
  const { PDFRawStream, PDFContentStream, decodePDFRawStream } = PDFLib;
  const { Contents } = page.node.normalizedEntries();
  if (!Contents) throw new Error("Missing page contents");
  const parts = [];
  for (let i = 0; i < Contents.size(); i++) {
    const stream = Contents.lookup(i);
    let bytes;
    if (stream instanceof PDFRawStream) {
      bytes = decodePDFRawStream(stream).decode();
    } else if (stream instanceof PDFContentStream) {
      bytes = stream.getUnencodedContents();
    } else {
      throw new Error("Unrecognized content stream type in outline PDF.");
    }
    parts.push(bytesToLatin1(bytes));
  }
  return parts.join("\n");
}

/**
 * Outline pages store each size as a cut-box stroke plus a clipped artwork
 * block. Keep only the artwork (drop die-stroke / nested clips), trim stray
 * transforms, then rewrite into local 0..w x 0..h form space so Illustrator
 * does not keep ghost frames behind the art.
 */
function extractLabelContentBlocks(contentText) {
  const clips = findClipGroups(contentText);
  const labels = clips.filter((c) => isLabelSized(c.w, c.h, false));

  return labels.map((label, i) => {
    const nextLabel = labels[i + 1];
    let end = nextLabel ? nextLabel.index : contentText.length;
    for (const c of clips) {
      if (c.index <= label.index) continue;
      if (c.index >= end) break;
      if (c.w > 500 && c.h > 500) continue;
      if (!rectsIntersect(label, c)) {
        end = c.index;
        break;
      }
    }

    const slice = contentText.slice(label.index, end);
    let content = extractArtworkOnly(slice, label);
    content = trimContentToLabelBounds(content, label);
    content = stripNestedClipRects(content);
    content = wrapContentInLocalFormSpace(content, label);

    return {
      x: 0,
      y: 0,
      w: label.w,
      h: label.h,
      sortX: label.x,
      sortY: label.y,
      content: content.replace(/\s+$/, "\n"),
    };
  });
}

function findClipGroups(contentText) {
  const clipRe = /(?:^|\n)(q\n)([-\d.]+) ([-\d.]+) ([-\d.]+) ([-\d.]+) re\nW n\n/g;
  const clips = [];
  let m;
  while ((m = clipRe.exec(contentText)) !== null) {
    let x = Number(m[2]);
    let y = Number(m[3]);
    let w = Number(m[4]);
    let h = Number(m[5]);
    if (w < 0) {
      x += w;
      w = -w;
    }
    if (h < 0) {
      y += h;
      h = -h;
    }
    const qIndex = m[0].startsWith("\n") ? m.index + 1 : m.index;
    clips.push({ index: qIndex, x, y, w, h });
  }
  return clips;
}

/** Skip the short cut-box stroke; keep fills + main artwork for this label. */
function extractArtworkOnly(slice, label) {
  if (!slice.startsWith("q")) return slice;
  const cutEnd = findBalancedQ(slice, 0);
  if (cutEnd < 0) return slice;
  let body = slice.slice(cutEnd);
  body = body.replace(
    /q\n0 0 [-\d.]+ [-\d.]+ re\nW n\n/g,
    `q\n${label.x} ${label.y} ${label.w} ${label.h} re\nW n\n`
  );
  return body;
}

function findBalancedQ(text, qIndex) {
  let depth = 1;
  let i = qIndex + 1;
  const n = text.length;
  while (i < n && depth > 0) {
    const ch = text[i];
    if (ch === "q" || ch === "Q") {
      const prev = i > 0 ? text[i - 1] : "\n";
      const next = i + 1 < n ? text[i + 1] : "\n";
      if (/\s/.test(prev) && /\s/.test(next)) {
        if (ch === "q") depth += 1;
        else {
          depth -= 1;
          if (depth === 0) return i + 1;
        }
      }
    }
    i += 1;
  }
  return -1;
}

/** Drop trailing artwork whose placement transform sits outside the cut box. */
function trimContentToLabelBounds(content, label) {
  const re = /q 1 0 0 1 ([-\d.]+) ([-\d.]+) cm/g;
  let m;
  let cutAt = null;
  while ((m = re.exec(content)) !== null) {
    const tx = Number(m[1]);
    const ty = Number(m[2]);
    const inside =
      tx >= label.x - 5 &&
      tx <= label.x + label.w + 5 &&
      ty >= label.y - 5 &&
      ty <= label.y + label.h + 5;
    if (!inside) {
      cutAt = m.index;
      break;
    }
  }
  if (cutAt == null) return content;

  let trimmed = content.slice(0, cutAt).replace(/\s+$/, "\n");
  let depth = 0;
  for (const line of trimmed.split("\n")) {
    const s = line.trim();
    if (s === "q" || s.startsWith("q ")) depth += 1;
    else if (s === "Q") depth -= 1;
  }
  while (depth > 0) {
    trimmed += "Q\n";
    depth -= 1;
  }
  return trimmed;
}

/** Remove inner clip rects; outer local wrap already clips to the cut box. */
function stripNestedClipRects(content) {
  return content.replace(/q\n[-\d.]+ [-\d.]+ [-\d.]+ [-\d.]+ re\nW n\n/g, "q\n");
}

/** Move page-space label content into a 0..w x 0..h form with a tight clip. */
function wrapContentInLocalFormSpace(content, label) {
  return (
    `q\n` +
    `0 0 ${label.w} ${label.h} re\n` +
    `W n\n` +
    `1 0 0 1 ${-label.x} ${-label.y} cm\n` +
    content.replace(/\s+$/, "\n") +
    `Q\n`
  );
}

/** Remove the layout master's placeholder Plate A/B/C/D margin text (avoids duplicates). */
function stripLayoutPlateText(contentText) {
  return contentText.replace(/BT\n[\s\S]*?\nET\n/g, (block) =>
    /\(Plate\s+[A-D]/.test(block) ? "" : block
  );
}

/* ------------------------------------------------------------------------- *
 * Generic content-stream scanner                                            *
 * Used when outlines / masters have no `re`-based frames or clip groups,    *
 * e.g. M&S files with outlined fonts and rounded cut paths.                 *
 * ------------------------------------------------------------------------- */

const PATH_BUILD_OPS = new Set(["m", "l", "c", "v", "y", "h", "re"]);
const PATH_PAINT_OPS = new Set(["S", "s", "f", "F", "f*", "B", "B*", "b", "b*", "n"]);
const STROKE_PAINT_OPS = new Set(["S", "s", "B", "B*", "b", "b*"]);
const TEXT_SHOW_OPS = new Set(["Tj", "TJ", "'", '"']);
const TEXT_POSITION_OPS = new Set(["Tm", "Td", "TD", "T*"]);
const IDENTITY = [1, 0, 0, 1, 0, 0];
/** Slack (pt) around a label's cut box when deciding which label owns artwork. */
const LABEL_OWNER_PAD = 4;

function isPdfWhite(c) {
  return c === 32 || c === 10 || c === 13 || c === 9 || c === 12 || c === 0;
}

function isPdfDelim(c) {
  return (
    c === 40 || c === 41 || c === 60 || c === 62 || c === 91 ||
    c === 93 || c === 123 || c === 125 || c === 47 || c === 37
  );
}

function skipPdfWhite(text, i) {
  const n = text.length;
  while (i < n) {
    const c = text.charCodeAt(i);
    if (isPdfWhite(c)) i++;
    else if (c === 37) {
      while (i < n && text.charCodeAt(i) !== 10 && text.charCodeAt(i) !== 13) i++;
    } else break;
  }
  return i;
}

function readPdfRegular(text, i) {
  const n = text.length;
  while (i < n) {
    const c = text.charCodeAt(i);
    if (isPdfWhite(c) || isPdfDelim(c)) break;
    i++;
  }
  return i;
}

function skipPdfString(text, i) {
  let depth = 0;
  for (const n = text.length; i < n; i++) {
    const c = text.charCodeAt(i);
    if (c === 92) i++;
    else if (c === 40) depth++;
    else if (c === 41 && --depth === 0) return i + 1;
  }
  return text.length;
}

function skipPdfDict(text, i) {
  let depth = 0;
  const n = text.length;
  while (i < n) {
    const c = text.charCodeAt(i);
    if (c === 40) {
      i = skipPdfString(text, i);
    } else if (c === 60 && text.charCodeAt(i + 1) === 60) {
      depth++;
      i += 2;
    } else if (c === 62 && text.charCodeAt(i + 1) === 62) {
      i += 2;
      if (--depth === 0) return i;
    } else i++;
  }
  return n;
}

/** Read one operand at `i`: [token, next] — or null when `i` starts an operator. */
function readPdfOperand(text, i) {
  const c = text.charCodeAt(i);
  if (c === 40) {
    const end = skipPdfString(text, i);
    return [{ t: "str", raw: text.slice(i, end) }, end];
  }
  if (c === 60) {
    if (text.charCodeAt(i + 1) === 60) return [{ t: "dict" }, skipPdfDict(text, i)];
    const close = text.indexOf(">", i);
    const end = close < 0 ? text.length : close + 1;
    return [{ t: "hex", raw: text.slice(i, end) }, end];
  }
  if (c === 91) {
    const items = [];
    let j = i + 1;
    for (;;) {
      j = skipPdfWhite(text, j);
      if (j >= text.length) break;
      if (text.charCodeAt(j) === 93) {
        j++;
        break;
      }
      const r = readPdfOperand(text, j);
      if (r) {
        items.push(r[0]);
        j = r[1];
      } else {
        j = Math.max(j + 1, readPdfRegular(text, j));
      }
    }
    return [{ t: "arr", items }, j];
  }
  if (c === 47) {
    const end = readPdfRegular(text, i + 1);
    return [{ t: "name", v: text.slice(i, end) }, end];
  }
  const end = readPdfRegular(text, i);
  if ((c >= 48 && c <= 57) || c === 43 || c === 45 || c === 46) {
    return [{ t: "num", v: Number(text.slice(i, end)) || 0 }, end];
  }
  const word = text.slice(i, end);
  if (word === "true" || word === "false" || word === "null") return [{ t: "kw", v: word }, end];
  return null;
}

/** Tokenize a content stream into `{ op, args, start, end }` (offsets span operands + operator). */
function parseContentOps(text) {
  const ops = [];
  const n = text.length;
  let args = [];
  let argStart = -1;
  let i = 0;
  for (;;) {
    i = skipPdfWhite(text, i);
    if (i >= n) break;
    const operand = readPdfOperand(text, i);
    if (operand) {
      if (argStart < 0) argStart = i;
      args.push(operand[0]);
      i = operand[1];
      continue;
    }
    let end = readPdfRegular(text, i);
    if (end === i) {
      i++;
      continue;
    }
    const op = text.slice(i, end);
    if (op === "BI") {
      // Inline image: skip binary data up to the EI keyword.
      const id = text.indexOf("ID", end);
      const eiRe = /\sEI(?=\s|$)/g;
      eiRe.lastIndex = id < 0 ? end : id + 2;
      const m = eiRe.exec(text);
      end = m ? m.index + 3 : n;
    }
    ops.push({ op, args, start: argStart < 0 ? i : argStart, end });
    args = [];
    argStart = -1;
    i = end;
  }
  return ops;
}

function opNums(o, count) {
  const out = [];
  for (let k = 0; k < count; k++) {
    const a = o.args[k];
    out.push(a && a.t === "num" ? a.v : 0);
  }
  return out;
}

/** m × n for PDF row-vector matrices [a b c d e f]. */
function mulMatrix(m, n) {
  return [
    m[0] * n[0] + m[1] * n[2],
    m[0] * n[1] + m[1] * n[3],
    m[2] * n[0] + m[3] * n[2],
    m[2] * n[1] + m[3] * n[3],
    m[4] * n[0] + m[5] * n[2] + n[4],
    m[4] * n[1] + m[5] * n[3] + n[5],
  ];
}

function bboxAddPoint(b, m, x, y) {
  const px = m[0] * x + m[2] * y + m[4];
  const py = m[1] * x + m[3] * y + m[5];
  if (!b) return [px, py, px, py];
  if (px < b[0]) b[0] = px;
  if (py < b[1]) b[1] = py;
  if (px > b[2]) b[2] = px;
  if (py > b[3]) b[3] = py;
  return b;
}

function bboxOfBox(m, [x0, y0, x1, y1]) {
  let b = bboxAddPoint(null, m, x0, y0);
  b = bboxAddPoint(b, m, x1, y0);
  b = bboxAddPoint(b, m, x0, y1);
  return bboxAddPoint(b, m, x1, y1);
}

function bboxUnion(a, b) {
  if (!a) return b ? b.slice() : null;
  if (!b) return a;
  return [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[2], b[2]), Math.max(a[3], b[3])];
}

function bboxInside(b, zone) {
  return b[0] >= zone[0] && b[1] >= zone[1] && b[2] <= zone[2] && b[3] <= zone[3];
}

function bboxIntersects(b, zone) {
  return !(b[2] < zone[0] || b[0] > zone[2] || b[3] < zone[1] || b[1] > zone[3]);
}

function decodePdfString(tok) {
  if (tok.t === "hex") {
    const hex = tok.raw.slice(1, -1).replace(/\s+/g, "");
    let s = "";
    for (let i = 0; i < hex.length; i += 2) {
      s += String.fromCharCode(parseInt((hex.slice(i, i + 2) + "0").slice(0, 2), 16));
    }
    return s;
  }
  const raw = tok.raw.slice(1, -1);
  let s = "";
  for (let i = 0; i < raw.length; i++) {
    const ch = raw[i];
    if (ch !== "\\") {
      s += ch;
      continue;
    }
    const nx = raw[++i];
    if (nx === undefined) break;
    if (nx === "\r") {
      if (raw[i + 1] === "\n") i++;
      continue;
    }
    if (nx === "\n") continue;
    const esc = { n: "\n", r: "\r", t: "\t", b: "\b", f: "\f" }[nx];
    if (esc) {
      s += esc;
    } else if (/[0-7]/.test(nx)) {
      let oct = nx;
      while (oct.length < 3 && /[0-7]/.test(raw[i + 1] || "")) oct += raw[++i];
      s += String.fromCharCode(parseInt(oct, 8) & 0xff);
    } else {
      s += nx;
    }
  }
  return s;
}

/** Decoded text of a Tj / TJ / ' / " operator (big TJ gaps become spaces). */
function shownText(o) {
  const last = o.args[o.args.length - 1];
  if (!last) return "";
  if (last.t === "arr") {
    return last.items
      .map((it) => (it.t === "num" ? (it.v <= -250 ? " " : "") : decodePdfString(it)))
      .join("");
  }
  return last.t === "str" || last.t === "hex" ? decodePdfString(last) : "";
}

function strokeStateText(s) {
  return [s.gs, s.cs, s.color, s.width, s.dash].filter(Boolean).map((x) => `${x}\n`).join("");
}

/**
 * Parse a content stream into a tree of q/Q groups and atomic paint nodes,
 * each with a page-space bbox (CTM applied). Node types:
 * - group: q…Q, bbox = union of painted children
 * - path:  construction ops + paint op (clip-only paths become `state`)
 * - text:  BT…ET, bbox approximated from glyph origins
 * - paint: Do / inline image
 * - state: anything else (colour, cm, gs, clip, marked content…)
 */
function buildContentTree(text, resolveXObject) {
  const root = { type: "group", children: [], bbox: null, parent: null };
  let group = root;
  let ctm = IDENTITY;
  let stroke = { gs: "", cs: "", color: "", width: "", dash: "" };
  const saved = [];
  let path = null;
  let textNode = null;
  let tm = IDENTITY;
  let tlm = IDENTITY;
  let fontSize = 1;
  let leading = 0;

  const add = (node) => {
    node.parent = group;
    group.children.push(node);
  };
  const closeGroup = (end, unclosed) => {
    const g = group;
    g.end = end;
    g.unclosed = unclosed;
    for (const c of g.children) if (c.type !== "state") g.bbox = bboxUnion(g.bbox, c.bbox);
    group = g.parent;
    group.children.push(g);
  };
  const moveLine = (tx, ty) => {
    tlm = mulMatrix([1, 0, 0, 1, tx, ty], tlm);
    tm = tlm;
  };

  for (const o of parseContentOps(text)) {
    const op = o.op;

    if (textNode) {
      if (op === "ET") {
        textNode.end = o.end;
        add(textNode);
        textNode = null;
        continue;
      }
      if (op === "Tm") {
        tm = tlm = opNums(o, 6);
      } else if (op === "Td" || op === "TD") {
        const [tx, ty] = opNums(o, 2);
        if (op === "TD") leading = -ty;
        moveLine(tx, ty);
      } else if (op === "T*") {
        moveLine(0, -leading);
      } else if (op === "TL") {
        leading = opNums(o, 1)[0];
      } else if (op === "Tf") {
        fontSize = opNums(o, 2)[1];
      }
      if (TEXT_SHOW_OPS.has(op)) {
        if (op === "'" || op === '"') moveLine(0, -leading);
        const m = mulMatrix(tm, ctm);
        const chars = shownText(o).length;
        textNode.bbox = bboxAddPoint(textNode.bbox, m, 0, -0.25 * fontSize);
        textNode.bbox = bboxAddPoint(textNode.bbox, m, 0.55 * fontSize * chars, 0.8 * fontSize);
      } else if (!TEXT_POSITION_OPS.has(op)) {
        textNode.stateOps.push(text.slice(o.start, o.end));
      }
      continue;
    }

    if (PATH_BUILD_OPS.has(op)) {
      if (!path) path = { start: o.start, bbox: null, clipAt: -1 };
      if (op === "re") {
        const [x, y, w, h] = opNums(o, 4);
        path.bbox = bboxUnion(path.bbox, bboxOfBox(ctm, [x, y, x + w, y + h]));
      } else {
        const a = opNums(o, o.args.length);
        for (let k = 0; k + 1 < a.length; k += 2) path.bbox = bboxAddPoint(path.bbox, ctm, a[k], a[k + 1]);
      }
      continue;
    }
    if (op === "W" || op === "W*") {
      if (path && path.clipAt < 0) path.clipAt = o.start;
      continue;
    }
    if (PATH_PAINT_OPS.has(op)) {
      if (!path || path.clipAt >= 0 || op === "n") {
        const clipBox = path && path.clipAt >= 0 ? path.bbox : null;
        add({ type: "state", start: path ? path.start : o.start, end: o.end, clipBox });
      } else {
        add({
          type: "path",
          start: path.start,
          end: o.end,
          bbox: path.bbox,
          stroke: STROKE_PAINT_OPS.has(op),
          ctm,
          pathText: text.slice(path.start, o.start).trim(),
          strokeState: stroke,
        });
      }
      path = null;
      continue;
    }

    switch (op) {
      case "BT":
        textNode = { type: "text", start: o.start, bbox: null, stateOps: [] };
        tm = tlm = IDENTITY;
        continue;
      case "q":
        saved.push({ ctm, stroke });
        group = { type: "group", children: [], bbox: null, start: o.start, parent: group };
        continue;
      case "Q": {
        const s = saved.pop();
        if (s) ({ ctm, stroke } = s);
        if (group !== root) closeGroup(o.end, false);
        else add({ type: "state", start: o.start, end: o.end });
        continue;
      }
      case "cm":
        ctm = mulMatrix(opNums(o, 6), ctm);
        break;
      case "Do": {
        const name = o.args[0] && o.args[0].v;
        const form = resolveXObject ? resolveXObject(name) : null;
        const m = form ? mulMatrix(form.matrix, ctm) : ctm;
        add({ type: "paint", start: o.start, end: o.end, bbox: bboxOfBox(m, form ? form.bbox : [0, 0, 1, 1]) });
        continue;
      }
      case "BI":
        add({ type: "paint", start: o.start, end: o.end, bbox: bboxOfBox(ctm, [0, 0, 1, 1]) });
        continue;
      case "CS":
        stroke = { ...stroke, cs: text.slice(o.start, o.end), color: "" };
        break;
      case "SC":
      case "SCN":
        stroke = { ...stroke, color: text.slice(o.start, o.end) };
        break;
      case "RG":
      case "K":
      case "G":
        stroke = { ...stroke, cs: "", color: text.slice(o.start, o.end) };
        break;
      case "w":
        stroke = { ...stroke, width: text.slice(o.start, o.end) };
        break;
      case "gs":
        stroke = { ...stroke, gs: text.slice(o.start, o.end) };
        break;
      case "d":
        stroke = { ...stroke, dash: text.slice(o.start, o.end) };
        break;
    }
    add({ type: "state", start: o.start, end: o.end });
  }

  while (group !== root) closeGroup(text.length, true);
  return root;
}

function walkContentTree(node, fn) {
  for (const child of node.children) {
    fn(child);
    if (child.type === "group") walkContentTree(child, fn);
  }
}

/** Resolve a page XObject name to its form BBox/Matrix (images → null = unit square). */
function makeXObjectBoxResolver(page) {
  const { PDFName, PDFDict, PDFArray, PDFStream, PDFNumber } = PDFLib;
  const resources = page.node.Resources();
  const xobjects = resources && resources.lookupMaybe(PDFName.of("XObject"), PDFDict);
  if (!xobjects) return null;
  const nums = (arr) =>
    arr &&
    arr.asArray().map((v) => {
      const n = page.doc.context.lookup(v);
      return n instanceof PDFNumber ? n.asNumber() : 0;
    });
  return (name) => {
    const stream = name ? xobjects.lookup(PDFName.of(name.slice(1))) : null;
    if (!(stream instanceof PDFStream)) return null;
    if (stream.dict.get(PDFName.of("Subtype")) !== PDFName.of("Form")) return null;
    const b = nums(stream.dict.lookupMaybe(PDFName.of("BBox"), PDFArray));
    if (!b || b.length < 4) return null;
    return {
      bbox: [Math.min(b[0], b[2]), Math.min(b[1], b[3]), Math.max(b[0], b[2]), Math.max(b[1], b[3])],
      matrix: nums(stream.dict.lookupMaybe(PDFName.of("Matrix"), PDFArray)) || IDENTITY,
    };
  };
}

/**
 * Split an outline sheet into per-label content by geometry: each label's cut
 * box is a stroked label-sized path; every painted object is assigned to the
 * cut box that contains it (objects outside all boxes — sheet headers etc. —
 * are dropped, and the label's own die stroke is removed).
 */
function extractLabelBlocksBySpatialScan(contentText, resolveXObject) {
  const root = buildContentTree(contentText, resolveXObject);

  const labels = [];
  walkContentTree(root, (node) => {
    if (node.type !== "path" || !node.stroke || !node.bbox) return;
    const [x0, y0, x1, y1] = node.bbox;
    if (!isLabelSized(x1 - x0, y1 - y0, false)) return;
    node.isCut = true;
    for (let p = node.parent; p; p = p.parent) p.hasCut = true;
    if (!labels.some((l) => Math.abs(l.x - x0) < 1.5 && Math.abs(l.y - y0) < 1.5)) {
      labels.push({ x: x0, y: y0, w: x1 - x0, h: y1 - y0 });
    }
  });
  if (!labels.length) return [];

  const zones = labels.map((l) => [
    l.x - LABEL_OWNER_PAD,
    l.y - LABEL_OWNER_PAD,
    l.x + l.w + LABEL_OWNER_PAD,
    l.y + l.h + LABEL_OWNER_PAD,
  ]);
  walkContentTree(root, (node) => {
    if (node.type === "group" || node.type === "state") return;
    node.owner = -1;
    if (!node.bbox) return;
    node.owner = zones.findIndex((z) => bboxInside(node.bbox, z));
    if (node.owner < 0) {
      const cx = (node.bbox[0] + node.bbox[2]) / 2;
      const cy = (node.bbox[1] + node.bbox[3]) / 2;
      node.owner = zones.findIndex((z) => cx >= z[0] && cx <= z[2] && cy >= z[1] && cy <= z[3]);
    }
  });

  const slice = (node) => contentText.slice(node.start, node.end) + (node.unclosed ? "\nQ" : "");
  function emit(node, li, out) {
    const zone = zones[li];
    for (const child of node.children) {
      if (child.type === "state") {
        // A clip covering the whole label (e.g. the sheet's page clip) is redundant
        // with the form's own clip, and Illustrator shows it as a stray frame.
        if (child.clipBox && bboxInside(zone, child.clipBox)) continue;
        out.push(slice(child));
      } else if (child.type === "group") {
        if (!child.bbox) out.push(slice(child));
        else if (!bboxIntersects(child.bbox, zone)) continue;
        else if (!child.hasCut && bboxInside(child.bbox, zone)) out.push(slice(child));
        else {
          out.push("q");
          emit(child, li, out);
          out.push("Q");
        }
      } else if (child.isCut) {
        continue;
      } else if (child.owner === li) {
        out.push(slice(child));
      } else if (child.type === "text" && child.stateOps.length) {
        // Keep persistent text state (font, colour) set inside a dropped text block.
        out.push(`BT\n${child.stateOps.join("\n")}\nET`);
      }
    }
  }

  return labels.map((label, li) => {
    const out = [];
    emit(root, li, out);
    return {
      x: 0,
      y: 0,
      w: label.w,
      h: label.h,
      sortX: label.x,
      sortY: label.y,
      content: wrapContentInLocalFormSpace(`${out.join("\n")}\n`, label),
    };
  });
}

/** Label-sized painted paths (any orientation, any transform) from a master page. */
function extractLabelSizedPathsFromContent(contentText) {
  const frames = [];
  walkContentTree(buildContentTree(contentText, null), (node) => {
    if (node.type !== "path" || !node.bbox) return;
    const [x0, y0, x1, y1] = node.bbox;
    if (!isLabelSized(x1 - x0, y1 - y0, true)) return;
    frames.push({
      x: x0,
      y: y0,
      w: x1 - x0,
      h: y1 - y0,
      cut: node.stroke
        ? { ctm: node.ctm, path: node.pathText, state: strokeStateText(node.strokeState) }
        : null,
    });
  });
  // Prefer the stroked cut path over a same-sized fill when deduping.
  frames.sort((a, b) => Number(Boolean(b.cut)) - Number(Boolean(a.cut)));
  return dedupeFrames(frames);
}

/**
 * Remove the "…Plate-A- 45 sheet _ Date : …, Done by- …" line from a master
 * text block that also holds other lines. Returns the cleaned content and the
 * line's text + placement so it can be redrawn per plate, or null.
 */
function stripPlateTextLine(contentText) {
  const ops = parseContentOps(contentText);
  const lines = [];
  let line = null;
  let inText = false;
  let tm = IDENTITY;
  let tlm = IDENTITY;
  let fontSize = 1;
  let leading = 0;
  const newLine = () => {
    line = { shows: [], text: "", matrix: null, fontSize };
    lines.push(line);
  };
  const moveLine = (tx, ty) => {
    tlm = mulMatrix([1, 0, 0, 1, tx, ty], tlm);
    tm = tlm;
  };

  for (const o of ops) {
    if (o.op === "BT") {
      inText = true;
      tm = tlm = IDENTITY;
      newLine();
      continue;
    }
    if (!inText) continue;
    switch (o.op) {
      case "ET":
        inText = false;
        break;
      case "Tf":
        fontSize = opNums(o, 2)[1];
        break;
      case "TL":
        leading = opNums(o, 1)[0];
        break;
      case "Tm":
        tm = tlm = opNums(o, 6);
        newLine();
        break;
      case "Td":
      case "TD": {
        const [tx, ty] = opNums(o, 2);
        if (o.op === "TD") leading = -ty;
        moveLine(tx, ty);
        if (ty !== 0) newLine();
        break;
      }
      case "T*":
        moveLine(0, -leading);
        newLine();
        break;
      case "'":
      case '"':
        moveLine(0, -leading);
        newLine();
      // falls through
      case "Tj":
      case "TJ":
        if (!line.matrix) {
          line.matrix = tm;
          line.fontSize = fontSize;
        }
        line.shows.push(o);
        line.text += shownText(o);
        break;
    }
  }

  const target = lines.find((l) => /Plate[\s\-_]*[A-D]\b/i.test(l.text));
  if (!target) return null;
  let content = contentText;
  for (const o of [...target.shows].sort((a, b) => b.start - a.start)) {
    // ' and " also advance the line — keep that so later lines stay put.
    const keep = o.op === "'" || o.op === '"' ? "T*" : "";
    content = content.slice(0, o.start) + keep + content.slice(o.end);
  }
  const m = target.matrix;
  return {
    content,
    line: {
      text: target.text.replace(/\s+/g, " ").trim(),
      x: m[4],
      y: m[5],
      rotate: (Math.atan2(m[1], m[0]) * 180) / Math.PI,
      size: target.fontSize * Math.hypot(m[2], m[3]),
    },
  };
}

/** Fill a master's plate line template with this plate's values. */
function fillPlateLineTemplate(text, { plate, impression, woCompact, dateText, doneBy }) {
  return text
    .replace(/(Plate[\s\-_]*)[A-D]([\s\-_]*)\d*/i, (_, pre, mid) => `${pre}${plate}${mid}${impression}`)
    .replace(/BD\d{8,}W(?:\+\d+W)*/i, () => woCompact)
    .replace(/(Date\s*:\s*)[^,_]*/i, (_, pre) => `${pre}${dateText}`)
    .replace(/(Done\s*by\s*-?\s*).*$/i, (_, pre) => `${pre}${doneBy}`);
}

function fmtNum(n) {
  return String(Number(n.toFixed(4)));
}

async function getFirstPageHeight(buffer) {
  const pdf = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
  const page = await pdf.getPage(1);
  return page.getViewport({ scale: 1 }).height;
}

/**
 * Returns unique label-sized rects in PDF bottom-left coords: {x,y,w,h,pageHeight}.
 */
async function extractLabelSizedRectsFromPdf(buffer, pageHeight, options) {
  const bytes = new Uint8Array(buffer);
  const frames = [];
  const marker = asciiBytes("/Length ");

  for (let i = 0; i < bytes.length - 20; i++) {
    if (!matchAt(bytes, i, marker)) continue;
    let p = i + marker.length;
    let lenStr = "";
    while (p < bytes.length && bytes[p] >= 48 && bytes[p] <= 57) {
      lenStr += String.fromCharCode(bytes[p]);
      p++;
    }
    if (!lenStr) continue;
    const slice = bytes.subarray(p, Math.min(bytes.length, p + 40));
    const head = new TextDecoder("latin1").decode(slice);
    const sm = head.match(/^>>\s*stream\r?\n/);
    if (!sm) continue;

    const length = Number(lenStr);
    const dataStart = p + sm[0].length;
    const chunk = bytes.subarray(dataStart, dataStart + length);
    let text = "";
    try {
      const infl = await inflateZlib(chunk);
      text = new TextDecoder("latin1").decode(infl);
    } catch {
      // Uncompressed stream
      text = new TextDecoder("latin1").decode(chunk);
    }
    collectLabelSizedRects(text, pageHeight, frames, options);
  }

  return dedupeFrames(frames);
}

function asciiBytes(str) {
  const out = new Uint8Array(str.length);
  for (let i = 0; i < str.length; i++) out[i] = str.charCodeAt(i);
  return out;
}

function matchAt(bytes, index, needle) {
  if (index + needle.length > bytes.length) return false;
  for (let i = 0; i < needle.length; i++) {
    if (bytes[index + i] !== needle[i]) return false;
  }
  return true;
}

function isLabelSized(w, h, bothOrients) {
  const short = Math.min(w, h);
  const long = Math.max(w, h);
  if (short < LABEL_SHORT_MIN || short > LABEL_SHORT_MAX) return false;
  if (long < LABEL_LONG_MIN || long > LABEL_LONG_MAX) return false;
  if (!bothOrients && !(w < h)) return false; // outline: portrait only
  return true;
}

function collectLabelSizedRects(text, pageHeight, frames, options = {}) {
  const bothOrients = Boolean(options.bothOrients);
  const re = /([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+([-\d.]+)\s+re/g;
  let rm;
  while ((rm = re.exec(text)) !== null) {
    let x = Number(rm[1]);
    let y = Number(rm[2]);
    let w = Number(rm[3]);
    let h = Number(rm[4]);
    if (w < 0) {
      x += w;
      w = -w;
    }
    if (h < 0) {
      y += h;
      h = -h;
    }
    if (!isLabelSized(w, h, bothOrients)) continue;
    // Skip near-zero junk
    if (x < 1 || y < 1) continue;
    frames.push({ x, y, w, h, pageHeight });
  }
}

async function extractLabelSizedRectsViaOps(buffer, pageHeight, options) {
  const pdf = await pdfjsLib.getDocument({ data: buffer.slice(0) }).promise;
  const page = await pdf.getPage(1);
  const ops = await page.getOperatorList();
  const frames = [];
  for (let i = 0; i < ops.fnArray.length; i++) {
    if (ops.fnArray[i] !== pdfjsLib.OPS.rectangle) continue;
    let [x, y, w, h] = ops.argsArray[i];
    if (w < 0) {
      x += w;
      w = -w;
    }
    if (h < 0) {
      y += h;
      h = -h;
    }
    if (!isLabelSized(w, h, Boolean(options.bothOrients))) continue;
    if (x < 1 || y < 1) continue;
    frames.push({ x, y, w, h, pageHeight });
  }
  return dedupeFrames(frames);
}

async function inflateZlib(chunk) {
  try {
    return await inflateWith(chunk, "deflate");
  } catch {
    return await inflateWith(chunk, "deflate-raw");
  }
}

async function inflateWith(chunk, format) {
  const ds = new DecompressionStream(format);
  const stream = new Blob([chunk]).stream().pipeThrough(ds);
  const ab = await new Response(stream).arrayBuffer();
  return new Uint8Array(ab);
}

function dedupeFrames(frames) {
  const out = [];
  for (const f of frames) {
    if (!out.some((u) => Math.abs(u.x - f.x) < 1.5 && Math.abs(u.y - f.y) < 1.5)) {
      out.push(f);
    }
  }
  return out;
}

function buildPlateQueues(rows, labelMap) {
  const queues = { A: [], B: [], C: [], D: [] };
  for (const row of rows) {
    const key = labelKey(row.style, row.ean || row.size);
    const art = labelMap[key];
    for (const plate of PLATES) {
      const count = row[plate];
      for (let i = 0; i < count; i++) {
        queues[plate].push(art);
      }
    }
  }
  return queues;
}

function renderPlateSummary(queues, impressions, slotCount) {
  els.plateSummary.innerHTML = PLATES.map((plate) => {
    const items = queues[plate]
      .map((l) => `${l.style}/${rowName(l)}`)
      .join(", ");
    const ok = queues[plate].length === slotCount;
    return `<article class="plate-card">
      <strong>Plate ${plate} — ${queues[plate].length}/${slotCount} labels · Imp ${impressions[plate] || 0}${ok ? "" : " ⚠ mismatch"}</strong>
      <code>${escapeHtml(items)}</code>
    </article>`;
  }).join("");
  els.previewSection.classList.remove("hidden");
}

async function composeFinalPdf({
  layoutFile,
  layout,
  plateQueues,
  impressions,
  styles,
  dateText,
  doneBy,
}) {
  const layoutBytes = await layoutFile.arrayBuffer();
  const { PDFDocument, PDFName, rgb, degrees } = PDFLib;
  const layoutPdf = await PDFDocument.load(layoutBytes);
  // Delete placeholder "Plate A - …" text from the master so we don't stack a
  // second copy (whiteout hides it on screen, but Illustrator still sees both).
  // `plateLine` is set when the placeholder shares a text block with other
  // lines (e.g. "…_Plate-A- 45 sheet _ Date : …"); we redraw it in place.
  let plateLine = null;
  {
    const layoutPage = layoutPdf.getPages()[0];
    const original = getPageContentText(layoutPage);
    let cleaned = stripLayoutPlateText(original);
    if (cleaned === original) {
      const stripped = stripPlateTextLine(original);
      if (stripped) {
        cleaned = stripped.content;
        plateLine = stripped.line;
      }
    }
    layoutPage.node.set(
      PDFName.of("Contents"),
      layoutPdf.context.register(layoutPdf.context.flateStream(latin1ToBytes(cleaned)))
    );
  }
  const outPdf = await PDFDocument.create();
  const woCompact = compactStyleList(styles);
  const pageW = layout.pageW;
  const pageH = layout.pageH;
  const slots = layout.slots;
  const formByLabel = new Map();
  const resourcesBySource = new Map();
  const marginFont = plateLine
    ? await outPdf.embedFont(PDFLib.StandardFonts.HelveticaOblique)
    : null;

  /**
   * White underlay + cut stroke that follow the master's own cut path, as raw
   * page operators (no form XObject, so Illustrator gets no extra clip frame).
   * The stroke state refers to the master page's resources (e.g. /CS0, /GS1).
   */
  function cutOpsForSlot(slotIndex) {
    const cut = slots[slotIndex].cut;
    if (!cut) return null;
    const placed = `${cut.ctm.map(fmtNum).join(" ")} cm\n${cut.path}\n`;
    return {
      under: PDFLib.PDFOperator.of(`q\n1 g\n${placed}f\nQ`),
      over: PDFLib.PDFOperator.of(`q\n${cut.state}${placed}S\nQ`),
    };
  }

  async function resourcesForOutline(pdfBytes) {
    const cached = resourcesBySource.get(pdfBytes);
    if (cached) return cached;

    // Import fonts/XObjects once via a tiny donor form (not the full sheet art).
    const src = await PDFDocument.load(pdfBytes.slice(0));
    const srcPage = src.getPages()[0];
    srcPage.node.set(
      PDFName.of("Contents"),
      src.context.register(src.context.flateStream(latin1ToBytes("\n")))
    );
    const donor = await outPdf.embedPage(srcPage, {
      left: 0,
      bottom: 0,
      right: 1,
      top: 1,
    });
    await donor.embed();
    const resources = outPdf.context.lookup(donor.ref).dict.get(PDFName.of("Resources"));
    resourcesBySource.set(pdfBytes, resources);
    return resources;
  }

  async function formForLabel(label) {
    const key = labelKey(label.style, label.ean || label.size);
    const cached = formByLabel.get(key);
    if (cached) return cached;

    const resources = await resourcesForOutline(label.pdfBytes);
    const { w, h } = label.clip;
    // Content is already rewritten into local 0..w x 0..h space. Keep an
    // identity Matrix so Illustrator does not invent ghost frames from a
    // page-space BBox/Matrix pair.
    const ref = outPdf.context.register(
      outPdf.context.flateStream(latin1ToBytes(label.content), {
        Type: "XObject",
        Subtype: "Form",
        FormType: 1,
        BBox: [0, 0, w, h],
        Matrix: [1, 0, 0, 1, 0, 0],
        Resources: resources,
      })
    );
    const form = { ref, width: w, height: h };
    formByLabel.set(key, form);
    return form;
  }

  for (let p = 0; p < PLATES.length; p++) {
    setProgress(55 + p * 10, `Stamping Plate ${PLATES[p]}…`);
    const plate = PLATES[p];
    const [page] = await outPdf.copyPages(layoutPdf, [0]);
    outPdf.addPage(page);

    if (plateLine) {
      const margin = fillPlateLineTemplate(plateLine.text, {
        plate,
        impression: impressions[plate] || 0,
        woCompact,
        dateText,
        doneBy,
      });
      page.drawText(margin, {
        x: plateLine.x,
        y: plateLine.y,
        size: plateLine.size,
        rotate: degrees(plateLine.rotate),
        font: marginFont,
        color: rgb(0, 0, 0),
      });
    } else {
      const margin =
        `Plate ${plate} - ${impressions[plate] || 0} _ W/O _${woCompact}_ Date : ${dateText}, Done by- ${doneBy}`;
      page.drawText(margin, {
        x: pageW - 28,
        y: 40,
        size: 8,
        rotate: degrees(90),
        color: rgb(0, 0, 0),
        maxWidth: Math.max(40, pageH - 60),
      });
    }

    const queue = plateQueues[plate];
    for (let i = 0; i < slots.length; i++) {
      const label = queue[i];
      const form = await formForLabel(label);
      drawCroppedLabel(page, form, slots[i], cutOpsForSlot(i));
    }
  }

  return outPdf.save();
}

/** Layout-master die/cut mark color (cyan). */
const CUT_MARK_RGB = [0, 0.681, 0.938];
const CUT_MARK_WIDTH = 1;

/**
 * Stamp one cropped label form into a slot.
 * Landscape slots rotate the portrait art 90° clockwise (punch hole to the right).
 * Re-draws the cyan cut rectangle afterward — the white underlay covers the
 * layout master’s marks, and outline artwork no longer carries its own die stroke.
 * With `cutOps` (path-detected masters) the underlay and cut line follow the
 * master's original cut path, e.g. rounded corners, in its own cut colour.
 */
function drawCroppedLabel(page, form, slot, cutOps) {
  const {
    rgb,
    degrees,
    pushGraphicsState,
    popGraphicsState,
    translate,
    scale,
    rotateRadians,
    drawObject,
    toRadians,
  } = PDFLib;
  const { pdfX, pdfY, w: slotW, h: slotH } = slot.rect;

  if (cutOps) {
    page.pushOperators(cutOps.under);
  } else {
    page.drawRectangle({
      x: pdfX,
      y: pdfY,
      width: slotW,
      height: slotH,
      color: rgb(1, 1, 1),
    });
  }

  const key = page.node.newXObject("Label", form.ref);
  if (slot.orient === "H") {
    page.pushOperators(
      pushGraphicsState(),
      translate(pdfX, pdfY + slotH),
      rotateRadians(toRadians(degrees(-90))),
      scale(slotH / form.width, slotW / form.height),
      drawObject(key),
      popGraphicsState()
    );
  } else {
    page.pushOperators(
      pushGraphicsState(),
      translate(pdfX, pdfY),
      scale(slotW / form.width, slotH / form.height),
      drawObject(key),
      popGraphicsState()
    );
  }

  if (cutOps) {
    page.pushOperators(cutOps.over);
    return;
  }
  page.drawRectangle({
    x: pdfX,
    y: pdfY,
    width: slotW,
    height: slotH,
    borderColor: rgb(...CUT_MARK_RGB),
    borderWidth: CUT_MARK_WIDTH,
  });
}

function compactStyleList(styles) {
  if (!styles.length) return "XXXXXX";
  const first = styles[0];
  // Match production naming: BD01485474W+6551W+6575W+... / BD01493206W+3207W+...
  const rest = styles.slice(1).map((s) => s.replace(/^BD\d{4}/i, ""));
  return [first, ...rest].join("+");
}

function buildOutputName(styles) {
  return `${compactStyleList(styles)}_composed_plates.pdf`;
}

function formatDate(d) {
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function setProgress(pct, message) {
  els.progressBar.style.width = `${Math.max(0, Math.min(100, pct))}%`;
  if (message) setStatus(message);
}

function setStatus(message) {
  els.statusText.textContent = message;
}

function log(message) {
  const li = document.createElement("li");
  li.textContent = message;
  els.log.appendChild(li);
  els.log.scrollTop = els.log.scrollHeight;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

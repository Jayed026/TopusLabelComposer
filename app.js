/* global pdfjsLib, PDFLib, XLSX */

pdfjsLib.GlobalWorkerOptions.workerSrc =
  "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";

const PLATES = ["A", "B", "C", "D"];
const RENDER_SCALE = 2.5;
/** 30×74 mm label short/long sides in PDF points (±tolerance). */
const LABEL_SHORT_MIN = 78;
const LABEL_SHORT_MAX = 95;
const LABEL_LONG_MIN = 195;
const LABEL_LONG_MAX = 225;

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

    setProgress(20, "Extracting labels from outline PDFs…");
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

  for (let i = 1; i < matrix.length; i++) {
    const row = matrix[i] || [];
    const col0 = row[0] != null ? String(row[0]).trim() : "";
    const size = row[1] != null ? String(row[1]).trim() : "";
    const eanRaw = row[2];

    // Footer impression row: Imp_142 sheet, etc.
    if (!size && [row[3], row[4], row[5], row[6]].some((v) => /Imp_/i.test(String(v || "")))) {
      impressions = {
        A: parseImp(row[3]),
        B: parseImp(row[4]),
        C: parseImp(row[5]),
        D: parseImp(row[6]),
      };
      continue;
    }

    if (col0) {
      currentStyle = col0;
      if (!styles.includes(currentStyle)) styles.push(currentStyle);
    }
    if (!currentStyle || !size) continue;

    const ean = String(eanRaw ?? "").replace(/\D/g, "");
    rows.push({
      style: currentStyle,
      size,
      ean,
      A: toCount(row[3]),
      B: toCount(row[4]),
      C: toCount(row[5]),
      D: toCount(row[6]),
    });
  }

  if (!rows.length) throw new Error("No size rows found in Excel.");
  return { rows, styles, impressions };
}

function parseImp(value) {
  const m = String(value || "").match(/Imp[_\s-]?(\d+)/i);
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

    const frames = await findLabelFrames(file);
    if (frames.length !== expected.length) {
      throw new Error(
        `${style}: found ${frames.length} label frames but Excel has ${expected.length} sizes.`
      );
    }

    const pageImage = await renderPdfPageToCanvas(file, 0, RENDER_SCALE);
    frames.forEach((frame, i) => {
      const row = expected[i];
      const key = labelKey(row.style, row.ean || row.size);
      const crop = cropCanvas(pageImage.canvas, frame, pageImage.scale, pageImage.pageHeight);
      labelMap[key] = {
        style: row.style,
        size: row.size,
        ean: row.ean,
        canvas: crop,
        fileName: file.name,
      };
    });

    log(`${style}: extracted ${frames.length} sizes from ${file.name}`);
  }

  // Ensure every excel row with any plate qty has artwork
  for (const row of excelRows) {
    if (row.A + row.B + row.C + row.D === 0) continue;
    const key = labelKey(row.style, row.ean || row.size);
    if (!labelMap[key]) {
      throw new Error(`Missing artwork for ${row.style} / ${row.size} (EAN ${row.ean}).`);
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
 * Find 30×74mm portrait label frames in outline PDFs (inflated content streams).
 */
async function findLabelFrames(file) {
  const buffer = await file.arrayBuffer();
  const pageHeight = await getFirstPageHeight(buffer);
  let frames = await extractLabelSizedRectsFromPdf(buffer, pageHeight, { bothOrients: false });
  // convert PDF bottom-left → visual top-left for canvas crop
  frames = frames
    .sort((a, b) => {
      const rowDiff = Math.round(b.y / 8) - Math.round(a.y / 8);
      if (rowDiff !== 0) return rowDiff;
      return a.x - b.x;
    })
    .map((f) => ({
      x: f.x,
      y: pageHeight - f.y - f.h,
      w: f.w,
      h: f.h,
    }));

  if (!frames.length) {
    const viaOps = await extractLabelSizedRectsViaOps(buffer, pageHeight, { bothOrients: false });
    frames = viaOps
      .sort((a, b) => {
        const rowDiff = Math.round(b.y / 8) - Math.round(a.y / 8);
        if (rowDiff !== 0) return rowDiff;
        return a.x - b.x;
      })
      .map((f) => ({
        x: f.x,
        y: pageHeight - f.y - f.h,
        w: f.w,
        h: f.h,
      }));
  }
  if (!frames.length) {
    throw new Error("Could not detect label frames in outline PDF.");
  }
  return frames;
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

async function renderPdfPageToCanvas(file, pageIndex, scale) {
  const buffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
  const page = await pdf.getPage(pageIndex + 1);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.ceil(viewport.width);
  canvas.height = Math.ceil(viewport.height);
  const ctx = canvas.getContext("2d", { alpha: false });
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);
  await page.render({ canvasContext: ctx, viewport }).promise;
  return { canvas, scale, pageHeight: viewport.height / scale };
}

function cropCanvas(sourceCanvas, frame, scale) {
  const sx = Math.max(0, Math.floor(frame.x * scale));
  const sy = Math.max(0, Math.floor(frame.y * scale));
  const sw = Math.min(sourceCanvas.width - sx, Math.ceil(frame.w * scale));
  const sh = Math.min(sourceCanvas.height - sy, Math.ceil(frame.h * scale));

  const out = document.createElement("canvas");
  out.width = sw;
  out.height = sh;
  const ctx = out.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, sw, sh);
  ctx.drawImage(sourceCanvas, sx, sy, sw, sh, 0, 0, sw, sh);
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
      .map((l) => `${l.style}/${l.size}`)
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
  const { PDFDocument, rgb, degrees } = PDFLib;
  const layoutPdf = await PDFDocument.load(layoutBytes);
  const outPdf = await PDFDocument.create();
  const woCompact = compactStyleList(styles);
  const pageW = layout.pageW;
  const pageH = layout.pageH;
  const slots = layout.slots;

  for (let p = 0; p < PLATES.length; p++) {
    setProgress(55 + p * 10, `Stamping Plate ${PLATES[p]}…`);
    const plate = PLATES[p];
    const [page] = await outPdf.copyPages(layoutPdf, [0]);
    outPdf.addPage(page);

    // Cover placeholder right-margin text, then write plate-specific text.
    page.drawRectangle({
      x: pageW - 42,
      y: 20,
      width: 30,
      height: Math.max(40, pageH - 40),
      color: rgb(1, 1, 1),
    });

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

    const queue = plateQueues[plate];
    for (let i = 0; i < slots.length; i++) {
      const slot = slots[i];
      const label = queue[i];
      const png = await canvasToPngBytes(label.canvas, slot.orient);
      const image = await outPdf.embedPng(png);
      page.drawImage(image, {
        x: slot.rect.pdfX,
        y: slot.rect.pdfY,
        width: slot.rect.w,
        height: slot.rect.h,
      });
    }
  }

  return outPdf.save();
}

function compactStyleList(styles) {
  if (!styles.length) return "XXXXXX";
  const first = styles[0];
  // Match production naming: BD01485474W+6551W+6575W+...
  const rest = styles.slice(1).map((s) => s.replace(/^BD0148/i, ""));
  return [first, ...rest].join("+");
}

async function canvasToPngBytes(canvas, orient) {
  const source = orient === "H" ? rotateCanvasCW(canvas) : canvas;
  const blob = await new Promise((resolve) => source.toBlob(resolve, "image/png"));
  if (!blob) throw new Error("Failed to encode label image.");
  return new Uint8Array(await blob.arrayBuffer());
}

/** Rotate portrait label 90° clockwise so punch hole sits on the right in H slots. */
function rotateCanvasCW(canvas) {
  const out = document.createElement("canvas");
  out.width = canvas.height;
  out.height = canvas.width;
  const ctx = out.getContext("2d");
  ctx.translate(out.width, 0);
  ctx.rotate(Math.PI / 2);
  ctx.drawImage(canvas, 0, 0);
  return out;
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

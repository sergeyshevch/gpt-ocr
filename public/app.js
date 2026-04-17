const STORAGE_KEY = "openai_api_key";
const STORAGE_BATCH = "gpt_ocr_batch_size";
const STORAGE_DETAIL = "gpt_ocr_detail_low";
const STORAGE_FORMAT = "gpt_ocr_output_format";
const STORAGE_CONCURRENCY = "gpt_ocr_concurrency";
const STORAGE_LANG = "gpt_ocr_expected_lang";

const MAX_FILES = 150;
const MODEL = "gpt-4o-mini";
const API_URL = "https://api.openai.com/v1/chat/completions";

const SYSTEM_MARKDOWN_SINGLE = `You are an OCR assistant for presentation slides. Recognize and transcribe ALL visible text from the image exactly as written—do not edit, rephrase, correct spelling or grammar, reorder, summarize, or alter the content in any way.

Language: keep the exact same language(s) and script as on the slide (including mixed languages). Never translate—do not rewrite into English or any other language.

Preserve structure and reading order:
- Use line breaks to match vertical layout where it helps readability.
- Use Markdown headings (#, ##) only when the slide clearly uses title vs body hierarchy.
- Use bullet or numbered lists when the slide uses lists.
- Represent tables as GitHub-flavored Markdown tables when columns/rows are clear; otherwise align columns with spaces/monospace blocks.
- Do not invent content; if text is unreadable, write [illegible].

Do not add a slide number, "Slide N", or similar label—output only the slide content.

Output only the transcribed content for this single image—no preamble or explanation.`;

const SYSTEM_MARKDOWN_BATCH = `You are an OCR assistant for presentation slides. You will receive several images in order; each image is one slide. Recognize and transcribe ALL visible text exactly as written—do not edit, rephrase, correct spelling or grammar, reorder, summarize, or alter the content in any way.

For EVERY image, output exactly one block in this exact format (use the slide numbers given in the user message):

--- Slide N ---
<Markdown transcription for that slide only>

Rules:
- N must match the slide index you were told for that image position (do not skip or merge slides).
- Language: keep the exact same language(s) and script as on each slide; never translate.
- Preserve structure: headings as Markdown #/##, lists, GFM tables where clear, line breaks for layout.
- No text before the first --- Slide --- line and no summary after the last block.
- Inside each block (after the delimiter line), transcribe only the slide—do not repeat slide numbers or "Slide N" as a heading inside the body.
- Do not invent or edit content; use [illegible] for unreadable text.`;

const SYSTEM_HTML_SINGLE = `You are an OCR assistant for presentation slides. Recognize and transcribe ALL visible text from the image exactly as written—do not edit, rephrase, correct spelling or grammar, reorder, summarize, or alter the content in any way.

Language: keep the exact same language(s) and script as on the slide (including mixed languages). Never translate—do not rewrite into English or any other language.

Output a single HTML fragment only (no <!DOCTYPE>, no <html>, <head>, or <body> wrapper).
Use semantic tags: <h1>–<h3> for titles, <p> for paragraphs, <br> only when line breaks are meaningful, <ul>/<ol>/<li> for lists.
For tables use <table border="1"> with <thead>/<tbody>, <tr>, <th>, <td> when the layout is clearly tabular.
Preserve reading order. Do not invent or edit content; use <span class="illegible">[illegible]</span> for unreadable text.

Do not add a slide number, "Slide N", or similar label in the HTML—only the slide content.

Output only the HTML for this one slide—no preamble or markdown.`;

const SYSTEM_HTML_BATCH = `You are an OCR assistant for presentation slides. You will receive several images in order; each image is one slide. Recognize and transcribe ALL visible text exactly as written—do not edit, rephrase, correct spelling or grammar, reorder, summarize, or alter the content in any way.

For EVERY image, output exactly one block in this exact format (slide numbers are given in the user message):

--- Slide N ---
<HTML fragment for that slide only — same rules as single-slide HTML: no document wrapper, semantic tags, tables with border="1" where appropriate>

Rules:
- N must match the slide index for that image position.
- Language: keep the exact same language(s) and script as on each slide; never translate.
- No text before the first --- Slide --- line. No markdown—only HTML inside each block.
- No summary after the last block.
- Inside each block's HTML, do not add headings or text that only label the slide index—only the slide's real content.
- Do not invent or edit content; use <span class="illegible">[illegible]</span> for unreadable text.`;

const $ = (id) => document.getElementById(id);

const apiKeyInput = $("api-key");
const filesInput = $("files");
const fileHint = $("file-hint");
const form = $("ocr-form");
const submitBtn = $("submit-btn");
const copyBtn = $("copy-btn");
const copyWordBtn = $("copy-word-btn");
const previewDocsBtn = $("preview-docs-btn");
const batchSizeSelect = $("batch-size");
const concurrencySelect = $("concurrency");
const detailLowInput = $("detail-low");
const expectedLangInput = $("expected-lang");
const outputFormatSelect = $("output-format");
const progressBar = $("progress-bar");
const progressLabel = $("progress-label");
const statusEl = $("status");
const output = $("output");
const bgWarning = $("bg-warning");

let isProcessing = false;
let wakeLock = null;

async function requestWakeLock() {
  if (!("wakeLock" in navigator)) return;
  try {
    wakeLock = await navigator.wakeLock.request("screen");
    wakeLock.addEventListener("release", () => { wakeLock = null; });
  } catch { /* user denied or not supported */ }
}

function releaseWakeLock() {
  if (wakeLock) {
    wakeLock.release().catch(() => {});
    wakeLock = null;
  }
}

document.addEventListener("visibilitychange", () => {
  if (!isProcessing) {
    bgWarning.hidden = true;
    return;
  }
  if (document.visibilityState === "hidden") {
    bgWarning.hidden = false;
  } else {
    bgWarning.hidden = true;
    requestWakeLock();
  }
});

function loadKey() {
  const stored = localStorage.getItem(STORAGE_KEY);
  if (stored) apiKeyInput.value = stored;
}

function loadOptions() {
  const b = localStorage.getItem(STORAGE_BATCH);
  if (b && ["1", "2", "3", "5"].includes(b)) batchSizeSelect.value = b;
  const c = localStorage.getItem(STORAGE_CONCURRENCY);
  if (c && ["1", "2", "3", "5"].includes(c)) concurrencySelect.value = c;
  const d = localStorage.getItem(STORAGE_DETAIL);
  if (d === "1" || d === "true") detailLowInput.checked = true;
  const lang = localStorage.getItem(STORAGE_LANG);
  if (lang) expectedLangInput.value = lang;
  const f = localStorage.getItem(STORAGE_FORMAT);
  if (f === "html" || f === "markdown") outputFormatSelect.value = f;
}

function saveKey() {
  const v = apiKeyInput.value.trim();
  if (v) localStorage.setItem(STORAGE_KEY, v);
}

function saveOptions() {
  localStorage.setItem(STORAGE_BATCH, batchSizeSelect.value);
  localStorage.setItem(STORAGE_CONCURRENCY, concurrencySelect.value);
  localStorage.setItem(STORAGE_DETAIL, detailLowInput.checked ? "1" : "0");
  localStorage.setItem(STORAGE_LANG, expectedLangInput.value.trim());
  localStorage.setItem(STORAGE_FORMAT, outputFormatSelect.value);
}

function setFileHint(count) {
  if (count === 0) {
    fileHint.textContent = "";
    return;
  }
  if (count > MAX_FILES) {
    fileHint.textContent = `Selected ${count} files. Only the first ${MAX_FILES} will be used.`;
    fileHint.classList.add("error");
  } else {
    fileHint.textContent = `${count} file(s) selected.`;
    fileHint.classList.remove("error");
  }
}

filesInput.addEventListener("change", () => {
  setFileHint(filesInput.files.length);
});

apiKeyInput.addEventListener("blur", saveKey);
batchSizeSelect.addEventListener("change", saveOptions);
concurrencySelect.addEventListener("change", saveOptions);
detailLowInput.addEventListener("change", saveOptions);
expectedLangInput.addEventListener("change", saveOptions);
outputFormatSelect.addEventListener("change", saveOptions);

const HEIC_TYPES = ["image/heic", "image/heif"];

function isHeic(file) {
  if (HEIC_TYPES.includes(file.type)) return true;
  const name = file.name.toLowerCase();
  return name.endsWith(".heic") || name.endsWith(".heif");
}

let _libheif = null;

async function getLibheif() {
  if (_libheif) return _libheif;
  const factory = (await import("https://cdn.jsdelivr.net/npm/libheif-js@1.19.8/libheif-wasm/libheif-bundle.mjs")).default;
  _libheif = factory();
  return _libheif;
}

async function convertHeicToJpeg(file) {
  const lib = await getLibheif();

  const buffer = await file.arrayBuffer();
  const decoder = new lib.HeifDecoder();
  const images = decoder.decode(new Uint8Array(buffer));
  if (!images || images.length === 0) {
    throw new Error(`Could not decode HEIC file: ${file.name}`);
  }

  const image = images[0];
  const width = image.get_width();
  const height = image.get_height();

  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  const imageData = ctx.createImageData(width, height);

  await new Promise((resolve, reject) => {
    image.display(imageData, (displayData) => {
      if (!displayData) return reject(new Error(`HEIF decode error: ${file.name}`));
      resolve();
    });
  });

  ctx.putImageData(imageData, 0, 0);

  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => (blob ? resolve(blob) : reject(new Error("Canvas toBlob failed"))),
      "image/jpeg",
      0.92
    );
  });
}

async function convertFiles(files, onProgress) {
  const result = [];
  for (let i = 0; i < files.length; i++) {
    result.push(isHeic(files[i]) ? await convertHeicToJpeg(files[i]) : files[i]);
    if (onProgress) onProgress(i + 1, files.length);
  }
  return result;
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(file);
  });
}

function getFilesList() {
  const list = Array.from(filesInput.files);
  list.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: "base" }));
  return list.slice(0, MAX_FILES);
}

function setProgress(done, total) {
  const pct = total === 0 ? 0 : Math.round((done / total) * 100);
  progressBar.value = pct;
  progressBar.max = 100;
  progressLabel.textContent = `${done} / ${total}`;
}

function setBusy(busy) {
  submitBtn.disabled = busy;
  filesInput.disabled = busy;
  apiKeyInput.disabled = busy;
  batchSizeSelect.disabled = busy;
  concurrencySelect.disabled = busy;
  detailLowInput.disabled = busy;
  expectedLangInput.disabled = busy;
  outputFormatSelect.disabled = busy;
}

/** Enable copy/preview only when the output area has text (avoids stuck disabled state). */
function updateOutputActions() {
  const has = output.value.trim().length > 0;
  copyBtn.disabled = !has;
  copyWordBtn.disabled = !has;
  if (previewDocsBtn) {
    previewDocsBtn.disabled = !has;
    previewDocsBtn.title = has
      ? "Opens one new tab with formatted content (no nested frames)—copy, then paste into Google Docs."
      : "Run recognition first; preview uses the text in the output box below.";
  }
}

function getImageDetail() {
  return detailLowInput.checked ? "low" : "high";
}

function getExpectedLang() {
  return expectedLangInput.value.trim();
}

function buildSystemPrompt(base) {
  const lang = getExpectedLang();
  if (!lang) return base;
  return (
    base +
    `\n\nIMPORTANT — expected language: the slides are in ${lang}. Output ALL text in ${lang} exactly as written. Do not transliterate, translate, or substitute words into any other language. Even if some words look similar to another language, keep them in ${lang} as they appear on the slide.`
  );
}

function getOutputFormat() {
  return outputFormatSelect.value === "html" ? "html" : "markdown";
}

function getBatchSize() {
  const n = parseInt(batchSizeSelect.value, 10);
  return Number.isFinite(n) && n >= 1 ? Math.min(5, n) : 1;
}

function getConcurrency() {
  const n = parseInt(concurrencySelect.value, 10);
  return Number.isFinite(n) && n >= 1 ? Math.min(5, n) : 1;
}

/** Split model output on --- Slide N --- markers; returns { slide, body } in order. */
function splitSlideResponse(text) {
  const trimmed = text.trim();
  const re = /^--- Slide (\d+) ---\s*$/gm;
  const markers = [];
  let m;
  while ((m = re.exec(trimmed)) !== null) {
    markers.push({
      n: Number(m[1], 10),
      headerEnd: m.index + m[0].length,
      headerStart: m.index,
    });
  }
  if (markers.length === 0) {
    return [];
  }
  const blocks = [];
  for (let i = 0; i < markers.length; i++) {
    const bodyStart = markers[i].headerEnd;
    const bodyEnd =
      i + 1 < markers.length ? markers[i + 1].headerStart : trimmed.length;
    blocks.push({
      slide: markers[i].n,
      body: trimmed.slice(bodyStart, bodyEnd).trim(),
    });
  }
  return blocks;
}

function normalizeBatchBlocks(raw, batchStart, count) {
  const blocks = splitSlideResponse(raw);
  if (blocks.length === 0) {
    throw new Error(
      'Could not find "--- Slide N ---" sections in the model reply. Try “Images per request” = 1.'
    );
  }
  if (blocks.length < count) {
    throw new Error(
      `Expected ${count} slide section(s), found ${blocks.length}. Try a smaller batch size.`
    );
  }
  const ordered = blocks.slice(0, count);
  return ordered.map((b, i) => ({
    slide: batchStart + i,
    body: b.body,
  }));
}

function buildUserContentBatch(start, end, dataUrls, format) {
  const n = dataUrls.length;
  const lang = getExpectedLang();
  const langNote = lang
    ? ` The expected language is ${lang}—output only in ${lang}, do not translate or mix languages.`
    : "";
  const lines = [
    `There are ${n} images in order.`,
    `They are slides ${start} through ${end} (inclusive).`,
    `Transcribe each slide in the same language(s) and script as printed on the slide; do not translate.${langNote}`,
    `For each image, output one block starting with exactly "--- Slide K ---" where K is that slide's number (${start}…${end}).`,
    format === "html"
      ? "Inside each block, output only an HTML fragment (no wrapper document)."
      : "Inside each block, output only Markdown for that slide.",
  ];
  const content = [{ type: "text", text: lines.join("\n") }];
  const detail = getImageDetail();
  for (const url of dataUrls) {
    content.push({
      type: "image_url",
      image_url: { url, detail },
    });
  }
  return content;
}

const ocrWorker = new Worker("ocr-worker.js");
let workerMsgId = 0;
const workerCallbacks = new Map();

ocrWorker.addEventListener("message", (e) => {
  const { id, result, error } = e.data;
  const cb = workerCallbacks.get(id);
  if (!cb) return;
  workerCallbacks.delete(id);
  if (error) cb.reject(new Error(error));
  else cb.resolve(result);
});

function callOpenAI(apiKey, systemPrompt, userContent) {
  return new Promise((resolve, reject) => {
    const id = ++workerMsgId;
    workerCallbacks.set(id, { resolve, reject });
    ocrWorker.postMessage({
      id,
      apiUrl: API_URL,
      apiKey,
      model: MODEL,
      systemPrompt,
      userContent,
    });
  });
}

function buildRichDocInnerHtml(combinedText, format) {
  const escaped = combinedText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
  if (format === "html") {
    const blocks = splitSlideResponse(combinedText);
    if (blocks.length === 0) {
      return combinedText;
    }
    return blocks
      .map((b) => `<section class="slide">${b.body}</section>`)
      .join('<hr style="margin:1.25em 0;border:none;border-top:1px solid #ccc"/>');
  }
  return `<pre style="white-space:pre-wrap;font-family:Calibri,Segoe UI,sans-serif;font-size:11pt;margin:0">${escaped}</pre>`;
}

function buildWordHtml(combinedText, format) {
  const inner = buildRichDocInnerHtml(combinedText, format);
  return `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>${inner}</body></html>`;
}

/** Strip script tags from model HTML before embedding in a preview document. */
function stripScriptsFromHtml(html) {
  return html
    .replace(/<script\b[\s\S]*?<\/script>/gi, "")
    .replace(/<\/script/gi, "<\\/script");
}

/**
 * One self-contained HTML document (single blob URL). Nested blob + iframe is blocked on
 * GitHub Pages / some browsers (“Not allowed to load local resource: blob:…”).
 */
function buildPreviewStandaloneDocument(combinedText, format) {
  const rawInner = buildRichDocInnerHtml(combinedText, format);
  const inner = stripScriptsFromHtml(rawInner);
  const tip =
    format !== "html"
      ? `<p class="banner-tip">Tip: use <strong>HTML</strong> output for headings, lists, and tables in Docs.</p>`
      : "";
  const css = `
    body { margin: 0; background: #fff; color: #202124; font-family: system-ui, -apple-system, Segoe UI, sans-serif; font-size: 14px; }
    .banner { background: #e8f0fe; border-bottom: 1px solid #b8c9ea; padding: 12px 16px; line-height: 1.55; }
    .banner-tip { margin: 8px 0 0; opacity: 0.95; font-size: 13px; }
    kbd { font-family: inherit; border: 1px solid #888; border-radius: 4px; padding: 1px 6px; background: #fff; }
    main#gptocr-main {
      padding: 1rem 1.25rem 2rem;
      font-family: Arial, Helvetica, sans-serif;
      line-height: 1.45;
      color: #111;
    }
    main#gptocr-main pre { font-family: Consolas, "Courier New", monospace; font-size: 13px; }
    main#gptocr-main table { border-collapse: collapse; }
    main#gptocr-main td, main#gptocr-main th { border: 1px solid #333; padding: 4px 8px; vertical-align: top; }
    main#gptocr-main ul, main#gptocr-main ol { margin: 0.5em 0; padding-left: 1.5em; }
    main#gptocr-main .slide h2 { font-size: 15pt; }
  `;
  return `<!DOCTYPE html><html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OCR preview — Google Docs</title>
<style>${css}</style>
</head><body>
  <div class="banner">
    <strong>Google Docs:</strong> slide content below is selected automatically when this tab opens.
    Press <kbd>Ctrl</kbd>+<kbd>C</kbd> / <kbd>⌘</kbd>+<kbd>C</kbd>, switch to your Doc, and paste.
    If nothing is selected, click inside the slide area, then <kbd>Ctrl</kbd>+<kbd>A</kbd> / <kbd>⌘</kbd>+<kbd>A</kbd>, then copy.
    ${tip}
  </div>
  <main id="gptocr-main">${inner}</main>
  <script>
(function () {
  var el = document.getElementById("gptocr-main");
  if (!el || !window.getSelection) return;
  try {
    var sel = window.getSelection();
    var range = document.createRange();
    range.selectNodeContents(el);
    sel.removeAllRanges();
    sel.addRange(range);
  } catch (e) {}
})();
  <\/script>
</body></html>`;
}

/**
 * Google Docs often ignores text/html written via the Clipboard API. Copying from a normal
 * rendered page produces clipboard data Docs accepts; we open that page in a new tab.
 */
function openPreviewForGoogleDocs() {
  const plain = output.value;
  if (!plain.trim()) return;

  const fmt = output.dataset.outputFormat || getOutputFormat();
  const html = buildPreviewStandaloneDocument(plain, fmt);
  const url = URL.createObjectURL(
    new Blob([html], { type: "text/html;charset=utf-8" })
  );
  const w = window.open(url, "_blank", "noopener,noreferrer");
  if (!w) {
    URL.revokeObjectURL(url);
    statusEl.classList.add("error");
    statusEl.textContent =
      "Pop-up blocked — allow pop-ups for this site, then try “Preview for Google Docs” again.";
    return;
  }
  statusEl.classList.remove("error");
  statusEl.textContent =
    "Preview opened — copy the selected slide area, then paste into Google Docs.";
}

async function copyRichToClipboard(plain, htmlDocument) {
  const htmlBlob = new Blob([htmlDocument], { type: "text/html" });
  const plainBlob = new Blob([plain], { type: "text/plain" });
  await navigator.clipboard.write([
    new ClipboardItem({
      "text/html": htmlBlob,
      "text/plain": plainBlob,
    }),
  ]);
}

const MAX_RETRIES = 3;

/** Rebuild output textarea from the slots array (only non-null entries, in order). */
function renderSlots(slots, format) {
  output.value = slots.filter((s) => s !== null).join("\n\n");
  output.dataset.outputFormat = format;
  updateOutputActions();
}

/** Fill slots for a single batch; throws on failure. */
async function processBatch(batchFiles, fileIndex, slots, apiKey, format) {
  const dataUrls = await Promise.all(
    batchFiles.map((f) => readFileAsDataUrl(f))
  );

  const start = fileIndex + 1;
  const end = fileIndex + batchFiles.length;

  let raw;
  if (batchFiles.length === 1) {
    const system = buildSystemPrompt(
      format === "html" ? SYSTEM_HTML_SINGLE : SYSTEM_MARKDOWN_SINGLE
    );
    const detail = getImageDetail();
    const lang = getExpectedLang();
    const langNote = lang
      ? ` The expected language is ${lang}—output only in ${lang}, do not translate or mix languages.`
      : "";
    const userContent = [
      {
        type: "text",
        text: `Transcribe this slide in the same language as on the slide; do not translate.${langNote}`,
      },
      {
        type: "image_url",
        image_url: { url: dataUrls[0], detail },
      },
    ];
    raw = await callOpenAI(apiKey, system, userContent);
    slots[fileIndex] = raw;
  } else {
    const system = buildSystemPrompt(
      format === "html" ? SYSTEM_HTML_BATCH : SYSTEM_MARKDOWN_BATCH
    );
    const userContent = buildUserContentBatch(start, end, dataUrls, format);
    raw = await callOpenAI(apiKey, system, userContent);
    const normalized = normalizeBatchBlocks(raw, start, batchFiles.length);
    for (let j = 0; j < normalized.length; j++) {
      slots[fileIndex + j] = normalized[j].body;
    }
  }
}

/**
 * Run a list of async tasks with a concurrency cap.
 * Each task is { run: async () => void, fileCount: number }.
 * onTaskDone is called after each task finishes (success or fail).
 * Returns an array of { task, error } for every failed task.
 */
async function runWithConcurrency(tasks, concurrency, onTaskDone) {
  const failed = [];
  let nextIdx = 0;

  async function worker() {
    while (nextIdx < tasks.length) {
      const idx = nextIdx++;
      const task = tasks[idx];
      try {
        await task.run();
      } catch (err) {
        failed.push({ task, error: err });
      }
      onTaskDone(task);
    }
  }

  const workers = Array.from({ length: Math.min(concurrency, tasks.length) }, () => worker());
  await Promise.all(workers);
  return failed;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  saveKey();
  saveOptions();

  const apiKey = apiKeyInput.value.trim();
  if (!apiKey) {
    statusEl.textContent = "Enter your API key.";
    statusEl.classList.add("error");
    return;
  }

  const files = getFilesList();
  if (files.length === 0) {
    statusEl.textContent = "Select at least one image.";
    statusEl.classList.add("error");
    return;
  }

  const format = getOutputFormat();
  const batchSize = getBatchSize();
  const concurrency = getConcurrency();

  statusEl.classList.remove("error");
  statusEl.textContent = "Starting…";
  output.value = "";
  delete output.dataset.outputFormat;
  updateOutputActions();
  setProgress(0, files.length);
  setBusy(true);
  isProcessing = true;
  bgWarning.hidden = true;
  await requestWakeLock();

  const hasHeic = files.some(isHeic);
  let readyFiles = files;
  if (hasHeic) {
    statusEl.textContent = `Converting HEIC images to JPEG…`;
    try {
      readyFiles = await convertFiles(files, (done, total) => {
        setProgress(done, total);
        statusEl.textContent = `Converting HEIC images… ${done}/${total}`;
      });
    } catch (err) {
      statusEl.classList.add("error");
      statusEl.textContent = `HEIC conversion failed: ${err.message}`;
      setBusy(false);
      isProcessing = false;
      releaseWakeLock();
      return;
    }
    setProgress(0, readyFiles.length);
  }

  const slots = new Array(readyFiles.length).fill(null);
  let processed = 0;

  const tasks = [];
  for (let i = 0; i < readyFiles.length; i += batchSize) {
    const batchFiles = readyFiles.slice(i, i + batchSize);
    const fileIndex = i;
    tasks.push({
      fileIndex,
      batchFiles,
      fileCount: batchFiles.length,
      run: () => processBatch(batchFiles, fileIndex, slots, apiKey, format),
    });
  }

  try {
    statusEl.textContent = `Processing ${readyFiles.length} slides (${concurrency} in parallel)…`;

    let failed = await runWithConcurrency(tasks, concurrency, (task) => {
      processed += task.fileCount;
      setProgress(processed, readyFiles.length);
      renderSlots(slots, format);
    });

    for (let attempt = 1; attempt <= MAX_RETRIES && failed.length > 0; attempt++) {
      const isRateLimit = failed.some((f) =>
        /rate limit|429|tokens? per min|TPM|RPM/i.test(f.error?.message || "")
      );
      const delaySec = isRateLimit ? attempt * 5 : attempt * 10;
      const retryConcurrency = isRateLimit ? 1 : Math.max(1, Math.floor(concurrency / 2));

      statusEl.classList.remove("error");
      statusEl.textContent = `Retry ${attempt}/${MAX_RETRIES}: ${failed.length} batch(es), waiting ${delaySec}s…`;
      await new Promise((r) => setTimeout(r, delaySec * 1000));

      statusEl.textContent = `Retry ${attempt}/${MAX_RETRIES}: ${failed.length} batch(es)…`;
      const retryTasks = failed.map((f) => ({
        ...f.task,
        run: () => processBatch(f.task.batchFiles, f.task.fileIndex, slots, apiKey, format),
      }));

      failed = await runWithConcurrency(retryTasks, retryConcurrency, () => {
        renderSlots(slots, format);
      });
    }

    if (failed.length > 0) {
      const ranges = failed.map((f) => {
        const s = f.task.fileIndex + 1;
        const e = f.task.fileIndex + f.task.batchFiles.length;
        return s === e ? `${s}` : `${s}–${e}`;
      });
      statusEl.classList.add("error");
      statusEl.textContent = `Done with errors. Failed slides: ${ranges.join(", ")}. ${failed[0].error?.message || ""}`;
    } else {
      statusEl.classList.remove("error");
      statusEl.textContent = "Done.";
    }
  } catch (err) {
    statusEl.classList.add("error");
    statusEl.textContent = err instanceof Error ? err.message : String(err);
  } finally {
    renderSlots(slots, format);
    setBusy(false);
    isProcessing = false;
    bgWarning.hidden = true;
    releaseWakeLock();
  }
});

copyBtn.addEventListener("click", async () => {
  try {
    await navigator.clipboard.writeText(output.value);
    statusEl.classList.remove("error");
    statusEl.textContent = "Copied as plain text.";
  } catch {
    statusEl.classList.add("error");
    statusEl.textContent = "Could not copy (browser blocked clipboard).";
  }
});

copyWordBtn.addEventListener("click", async () => {
  try {
    const plain = output.value;
    const fmt = output.dataset.outputFormat || getOutputFormat();
    const html = buildWordHtml(plain, fmt);
    await copyRichToClipboard(plain, html);
    statusEl.classList.remove("error");
    statusEl.textContent =
      fmt === "html"
        ? "Copied with HTML (works well in Word). For Google Docs use “Preview for Google Docs”."
        : "Copied; use HTML output + “Preview for Google Docs” for rich paste into Docs.";
  } catch {
    statusEl.classList.add("error");
    statusEl.textContent = "Could not copy (browser blocked clipboard).";
  }
});

if (previewDocsBtn) {
  previewDocsBtn.addEventListener("click", () => {
    openPreviewForGoogleDocs();
  });
}

loadKey();
loadOptions();
setFileHint(0);
updateOutputActions();

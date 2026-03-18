const UPLOAD_ENDPOINT = "/api/users/upload/xlsx/to-json/events";

const form = document.getElementById("uploadForm");
const fileInput = document.getElementById("fileInput");
const fileHint = document.getElementById("fileHint");
const modeSelect = document.getElementById("modeSelect");
const sheetNameField = document.getElementById("sheetNameField");
const sheetNameInput = document.getElementById("sheetNameInput");
const sheetIndexField = document.getElementById("sheetIndexField");
const sheetIndexInput = document.getElementById("sheetIndexInput");
const modeHelp = document.getElementById("modeHelp");
const submitBtn = document.getElementById("submitBtn");
const clearBtn = document.getElementById("clearBtn");
const statusEl = document.getElementById("status");
const responseChip = document.getElementById("responseChip");
const summaryEl = document.getElementById("summary");
const insightsEl = document.getElementById("insights");
const sheetPreviewEl = document.getElementById("sheetPreview");
const placeholderEl = document.getElementById("placeholder");
const rawJsonPanelEl = document.getElementById("rawJsonPanel");
const jsonOutputEl = document.getElementById("jsonOutput");
const toggleRawBtn = document.getElementById("toggleRawBtn");
const copyJsonBtn = document.getElementById("copyJsonBtn");
let activeStreamState = null;
let lastStreamState = null;
let renderScheduled = false;
let statusScheduled = false;
const MAX_RAW_TEXT_BYTES = 250_000;
const PREVIEW_RENDER_DELAY_MS = 300;
const STATUS_RENDER_DELAY_MS = 250;

const modeDescriptions = {
  first: "Returns the first sheet only as a single object with rows and recordCount.",
  all: "Returns every sheet inside a sheets array, plus sheetCount and totalRecordCount.",
  name: "Looks up one sheet by name and returns it inside the sheets array wrapper.",
  index: "Looks up one sheet by zero-based index and returns it inside the sheets array wrapper."
};

updateModeFields();
setStatus("Choose an .xlsx file to preview the JSON conversion.", "info");

modeSelect.addEventListener("change", updateModeFields);
fileInput.addEventListener("change", updateFileHint);
form.addEventListener("submit", handleSubmit);
clearBtn.addEventListener("click", clearResponse);
toggleRawBtn.addEventListener("click", toggleRawStreaming);
copyJsonBtn.addEventListener("click", copyJsonToClipboard);

function updateModeFields() {
  const mode = modeSelect.value;
  const needsSheetName = mode === "name";
  const needsSheetIndex = mode === "index";

  sheetNameField.hidden = !needsSheetName;
  sheetIndexField.hidden = !needsSheetIndex;

  sheetNameInput.required = needsSheetName;
  sheetIndexInput.required = needsSheetIndex;

  if (!needsSheetName) {
    sheetNameInput.value = "";
  }
  if (!needsSheetIndex) {
    sheetIndexInput.value = "";
  }

  modeHelp.textContent = modeDescriptions[mode] || "";
}

function updateFileHint() {
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    fileHint.innerHTML = "The backend currently supports only <code>.xlsx</code> uploads.";
    return;
  }

  fileHint.innerHTML =
    "Selected file: <code>" +
    escapeHtml(file.name) +
    "</code> (" +
    formatBytes(file.size) +
    ")";
}

async function handleSubmit(event) {
  event.preventDefault();

  const validationError = validateForm();
  if (validationError) {
    setStatus(validationError, "error");
    return;
  }

  const file = fileInput.files[0];
  const mode = modeSelect.value;
  const query = new URLSearchParams({ mode });

  if (mode === "name") {
    query.set("sheetName", sheetNameInput.value.trim());
  } else if (mode === "index") {
    query.set("sheetIndex", sheetIndexInput.value.trim());
  }

  const formData = new FormData();
  formData.append("file", file);

  setBusy(true);
  prepareStreamingView();
  setStatus("Uploading workbook and starting streamed event rendering...", "info");
  responseChip.textContent = "Connecting...";

  try {
    const response = await fetch(UPLOAD_ENDPOINT + "?" + query.toString(), {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      const rawText = await readPlainTextResponse(response);
      throw new Error(extractErrorMessage(rawText, response.status, response.statusText));
    }

    const streamState = createStreamState(file, mode);
    activeStreamState = streamState;
    lastStreamState = streamState;
    updateRawPanelState(streamState);
    renderRawText(streamState);
    await consumeNdjsonResponse(response, streamState);
    if (!streamState.rawStreamingEnabled) {
      streamState.rawText = streamState.fullRawText.slice(-MAX_RAW_TEXT_BYTES);
      streamState.rawTextTruncated = streamState.fullRawText.length > MAX_RAW_TEXT_BYTES;
      renderRawText(streamState);
    }
    renderStreamState(streamState);
    responseChip.textContent = streamState.completed ? "Stream complete" : "Stream finished";
    setStatus("Upload complete. Streamed sheet preview rendered below.", "success");
  } catch (error) {
    clearResponseViewOnly();
    responseChip.textContent = "Request failed";
    setStatus(error.message || "Something went wrong while uploading the file.", "error");
  } finally {
    activeStreamState = null;
    setBusy(false);
  }
}

function validateForm() {
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    return "Choose an Excel file before uploading.";
  }

  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    return "Only .xlsx files are supported by the current backend endpoint.";
  }

  if (modeSelect.value === "name" && !sheetNameInput.value.trim()) {
    return "Enter a sheet name for mode=name.";
  }

  if (modeSelect.value === "index") {
    const value = sheetIndexInput.value.trim();
    if (value === "") {
      return "Enter a sheet index for mode=index.";
    }
    if (!/^\d+$/.test(value)) {
      return "Sheet index must be a non-negative integer.";
    }
  }

  return "";
}

function renderStreamState(streamState) {
  const stats = buildSummaryFromState(streamState);

  responseChip.textContent = stats.chipLabel;
  placeholderEl.hidden = true;
  rawJsonPanelEl.hidden = false;

  summaryEl.hidden = false;
  summaryEl.innerHTML = stats.items
    .map(
      item =>
        '<article class="meta-card"><span class="meta-label">' +
        escapeHtml(item.label) +
        '</span><div class="meta-value">' +
        escapeHtml(item.value) +
        "</div></article>"
    )
    .join("");

  renderInsights(streamState.sheets);
  renderSheetPreview(streamState.sheets);
}

function buildSummaryFromState(streamState) {
  const sheetNames = streamState.sheets.map(sheet => sheet.sheetName || "(blank name)").join(", ");

  return {
    chipLabel: streamState.sheets.length ? "Live sheet event stream" : "Waiting for sheet data",
    items: [
      { label: "File", value: streamState.fileName },
      { label: "Mode", value: streamState.mode },
      { label: "Sheets Matched", value: stringifyNumber(resolveSheetCount(streamState)) },
      { label: "Total Rows", value: stringifyNumber(resolveTotalRecordCount(streamState)) },
      { label: "Sheet Names", value: sheetNames || "None" },
      { label: "Payload Size", value: formatBytes(streamState.fullRawText.length) }
    ]
  };
}

function clearResponse() {
  form.reset();
  updateModeFields();
  updateFileHint();
  clearResponseViewOnly();
  responseChip.textContent = "No response yet";
  setStatus("Choose an .xlsx file to preview the JSON conversion.", "info");
}

function clearResponseViewOnly() {
  summaryEl.hidden = true;
  summaryEl.innerHTML = "";
  insightsEl.hidden = true;
  insightsEl.innerHTML = "";
  sheetPreviewEl.hidden = true;
  sheetPreviewEl.innerHTML = "";
  placeholderEl.hidden = false;
  rawJsonPanelEl.hidden = true;
  jsonOutputEl.textContent = "";
  renderScheduled = false;
  statusScheduled = false;
  lastStreamState = null;
  updateRawPanelState(null);
}

function prepareStreamingView() {
  clearResponseViewOnly();
  placeholderEl.hidden = true;
  rawJsonPanelEl.hidden = false;
  updateRawPanelState(null);
}

function setBusy(isBusy) {
  submitBtn.disabled = isBusy;
  clearBtn.disabled = isBusy;
  fileInput.disabled = isBusy;
  modeSelect.disabled = isBusy;
  sheetNameInput.disabled = isBusy;
  sheetIndexInput.disabled = isBusy;
  submitBtn.textContent = isBusy ? "Uploading..." : "Upload and show JSON";
}

function setStatus(message, tone) {
  statusEl.textContent = message || "";
  statusEl.className = "status" + (tone ? " " + tone : "");
}

function extractErrorMessage(rawText, status, statusText) {
  if (!rawText) {
    return "Request failed with " + status + " " + statusText + ".";
  }

  try {
    const parsed = JSON.parse(rawText);
    if (parsed.message) {
      return parsed.message;
    }
    if (parsed.error) {
      return parsed.error;
    }
  } catch (error) {
    // Keep the plain text fallback below.
  }

  return rawText.length > 300 ? rawText.slice(0, 300) + "..." : rawText;
}

function stringifyNumber(value) {
  if (typeof value === "number") {
    return value.toLocaleString();
  }
  if (value === null || value === undefined || value === "") {
    return "-";
  }
  return String(value);
}

function formatBytes(bytes) {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return "-";
  }
  if (bytes < 1024) {
    return bytes + " B";
  }
  if (bytes < 1024 * 1024) {
    return (bytes / 1024).toFixed(1) + " KB";
  }
  return (bytes / (1024 * 1024)).toFixed(2) + " MB";
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

async function readPlainTextResponse(response) {
  if (!response.body || typeof response.body.getReader !== "function") {
    const fallbackText = await response.text();
    jsonOutputEl.textContent = fallbackText;
    return fallbackText;
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let rawText = "";

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        rawText += decoder.decode();
        break;
      }
      rawText += decoder.decode(value, { stream: true });
    }
  } finally {
    reader.releaseLock();
  }

  jsonOutputEl.textContent = rawText;
  return rawText;
}

async function consumeNdjsonResponse(response, streamState) {
  if (!response.body || typeof response.body.getReader !== "function") {
    throw new Error("This browser does not support streamed response reading.");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  responseChip.textContent = "Streaming live...";

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) {
        const finalText = decoder.decode();
        streamState.rawText += finalText;
        buffer += finalText;
        break;
      }

      const chunkText = decoder.decode(value, { stream: true });
      appendRawText(streamState, chunkText);
      streamState.chunkCount += 1;
      buffer += chunkText;
      processNdjsonBuffer(streamState, buffer);
      buffer = streamState.pendingBuffer;
      scheduleStatusRender(streamState);
      scheduleStreamRender(streamState);
    }
  } finally {
    reader.releaseLock();
  }

  processNdjsonBuffer(streamState, buffer + "\n");
  streamState.pendingBuffer = "";
  streamState.completed = true;
  updateStreamingPreview(streamState);
}

function processNdjsonBuffer(streamState, source) {
  const lines = source.split(/\r?\n/);
  streamState.pendingBuffer = lines.pop() || "";

  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed) {
      continue;
    }
    processNdjsonEvent(streamState, trimmed);
  }
}

function processNdjsonEvent(streamState, line) {
  let event;
  try {
    event = JSON.parse(line);
  } catch (error) {
    return;
  }

  switch (event.type) {
    case "streamStart":
      streamState.streamType = event.type;
      break;
    case "sheetStart":
      ensureSheet(streamState, event.sheetIndex, event.sheetName);
      break;
    case "headers":
      ensureSheet(streamState, event.sheetIndex, event.sheetName).headers = Array.isArray(event.headers)
        ? event.headers
        : [];
      break;
    case "row": {
      const sheet = ensureSheet(streamState, event.sheetIndex, event.sheetName);
      if (sheet.rows.length < 5 && event.data && typeof event.data === "object") {
        sheet.rows.push(event.data);
      }
      sheet.recordCount += 1;
      break;
    }
    case "sheetEnd": {
      const sheet = ensureSheet(streamState, event.sheetIndex, event.sheetName);
      sheet.recordCount = typeof event.recordCount === "number" ? event.recordCount : sheet.recordCount;
      break;
    }
    case "complete":
      streamState.completed = true;
      streamState.completeSheetCount = event.sheetCount;
      streamState.completeTotalRecordCount = event.totalRecordCount;
      break;
    default:
      break;
  }
}

function updateStreamingPreview(streamState) {
  renderRawText(streamState);
  responseChip.textContent = "Streaming live...";
  setStatus(
    "Receiving streamed events... " +
      stringifyNumber(streamState.chunkCount) +
      " chunks, " +
      formatBytes(streamState.fullRawText.length) +
      " loaded, " +
      stringifyNumber(streamState.sheets.length) +
      " sheets discovered." +
      (streamState.rawStreamingEnabled ? "" : " Live raw view is paused for smoother performance."),
    "info"
  );
}

function renderInsights(sheets) {
  const totalSheets = sheets.length;
  const totalRows = sheets.reduce((sum, sheet) => sum + (Number(sheet.recordCount) || 0), 0);
  const widestSheet = sheets.reduce((max, sheet) => Math.max(max, getFieldNames(sheet).length), 0);
  const previewRows = sheets.reduce((sum, sheet) => sum + Math.min(sheet.rows.length, 3), 0);

  const cards = [
    {
      title: "Sheets",
      value: stringifyNumber(totalSheets),
      note: totalSheets ? "Workbook sections rendered below." : "No matching sheet was returned."
    },
    {
      title: "Rows",
      value: stringifyNumber(totalRows),
      note: "Count is based on the backend JSON response."
    },
    {
      title: "Fields",
      value: stringifyNumber(widestSheet),
      note: "Maximum columns detected in the previewed sheets."
    },
    {
      title: "Preview Rows",
      value: stringifyNumber(previewRows),
      note: "A small sample is shown so the page stays readable."
    }
  ];

  insightsEl.hidden = false;
  insightsEl.innerHTML = cards
    .map(
      card =>
        '<article class="insight-card"><p class="insight-title">' +
        escapeHtml(card.title) +
        '</p><p class="insight-value">' +
        escapeHtml(card.value) +
        '</p><p class="insight-note">' +
        escapeHtml(card.note) +
        "</p></article>"
    )
    .join("");
}

function renderSheetPreview(sheets) {
  sheetPreviewEl.hidden = false;

  if (!sheets.length) {
    sheetPreviewEl.innerHTML =
      '<article class="placeholder">No sheet data matched the selected mode, so there is nothing to preview.</article>';
    return;
  }

  sheetPreviewEl.innerHTML = sheets
    .map((sheet, index) => renderSheetCard(sheet, index === 0))
    .join("");
}

function renderSheetCard(sheet, isOpenByDefault) {
  const fieldNames = getFieldNames(sheet);
  const previewRows = sheet.rows.slice(0, 5);
  const fieldsMarkup = fieldNames.length
    ? fieldNames
        .slice(0, 12)
        .map(field => '<span class="field-pill">' + escapeHtml(field) + "</span>")
        .join("")
    : '<span class="field-pill">No fields found</span>';

  const moreFieldsMarkup =
    fieldNames.length > 12
      ? '<span class="field-pill">+' + escapeHtml(fieldNames.length - 12) + " more</span>"
      : "";

  const tableMarkup = previewRows.length
    ? renderPreviewTable(fieldNames, previewRows)
    : '<div class="placeholder">This sheet has no data rows in the JSON response.</div>';

  return (
    '<details class="sheet-card"' +
    (isOpenByDefault ? " open" : "") +
    '><summary><div><h3 class="sheet-title">' +
    escapeHtml(sheet.sheetName || "Unnamed sheet") +
    '</h3><p class="sheet-meta">Sheet index: ' +
    escapeHtml(stringifyNumber(sheet.sheetIndex)) +
    " | Rows: " +
    escapeHtml(stringifyNumber(sheet.recordCount)) +
    '</p></div><div class="sheet-badge">' +
    escapeHtml(fieldNames.length) +
    ' fields</div></summary><div class="sheet-body"><div class="field-pill-row">' +
    fieldsMarkup +
    moreFieldsMarkup +
    "</div>" +
    tableMarkup +
    "</div></details>"
  );
}

function renderPreviewTable(fieldNames, rows) {
  const columns = fieldNames.slice(0, 8);

  return (
    '<div class="table-shell"><table class="preview-table"><thead><tr>' +
    columns.map(column => "<th>" + escapeHtml(column) + "</th>").join("") +
    '</tr></thead><tbody>' +
    rows
      .map(
        row =>
          "<tr>" +
          columns
            .map(column => {
              const value = row && Object.prototype.hasOwnProperty.call(row, column) ? row[column] : "";
              return renderCell(value);
            })
            .join("") +
          "</tr>"
      )
      .join("") +
    "</tbody></table></div>"
  );
}

function renderCell(value) {
  if (value === null || value === undefined || value === "") {
    return '<td class="cell-empty">empty</td>';
  }
  return "<td>" + escapeHtml(value) + "</td>";
}

function getFieldNames(sheet) {
  if (sheet && Array.isArray(sheet.headers) && sheet.headers.length) {
    return sheet.headers;
  }
  if (!sheet || !Array.isArray(sheet.rows) || !sheet.rows.length) {
    return [];
  }
  return Object.keys(sheet.rows[0]);
}

function createStreamState(file, mode) {
  return {
    fileName: file.name,
    mode,
    rawText: "",
    fullRawText: "",
    chunkCount: 0,
    sheets: [],
    sheetMap: new Map(),
    pendingBuffer: "",
    completed: false,
    completeSheetCount: null,
    completeTotalRecordCount: null,
    streamType: "",
    rawStreamingEnabled: false,
    rawTextTruncated: false
  };
}

function ensureSheet(streamState, sheetIndex, sheetName) {
  const key = String(sheetIndex);
  if (streamState.sheetMap.has(key)) {
    return streamState.sheetMap.get(key);
  }

  const sheet = {
    sheetIndex,
    sheetName: sheetName || "",
    headers: [],
    rows: [],
    recordCount: 0
  };
  streamState.sheetMap.set(key, sheet);
  streamState.sheets.push(sheet);
  streamState.sheets.sort((left, right) => {
    const leftIndex = typeof left.sheetIndex === "number" ? left.sheetIndex : Number.MAX_SAFE_INTEGER;
    const rightIndex = typeof right.sheetIndex === "number" ? right.sheetIndex : Number.MAX_SAFE_INTEGER;
    return leftIndex - rightIndex;
  });
  return sheet;
}

function scheduleStreamRender(streamState) {
  if (renderScheduled) {
    return;
  }

  renderScheduled = true;
  window.setTimeout(() => {
    renderScheduled = false;
    if (activeStreamState === streamState) {
      renderStreamState(streamState);
    }
  }, PREVIEW_RENDER_DELAY_MS);
}

function scheduleStatusRender(streamState) {
  if (statusScheduled) {
    return;
  }

  statusScheduled = true;
  window.setTimeout(() => {
    statusScheduled = false;
    if (activeStreamState === streamState) {
      updateStreamingPreview(streamState);
    }
  }, STATUS_RENDER_DELAY_MS);
}

function resolveSheetCount(streamState) {
  if (typeof streamState.completeSheetCount === "number") {
    return streamState.completeSheetCount;
  }
  return streamState.sheets.length;
}

function resolveTotalRecordCount(streamState) {
  if (typeof streamState.completeTotalRecordCount === "number") {
    return streamState.completeTotalRecordCount;
  }
  return streamState.sheets.reduce((sum, sheet) => sum + (Number(sheet.recordCount) || 0), 0);
}

async function copyJsonToClipboard() {
  const streamState = activeStreamState || lastStreamState;
  if (!streamState || !streamState.fullRawText) {
    return;
  }

  const originalLabel = copyJsonBtn.textContent;

  try {
    await navigator.clipboard.writeText(streamState.fullRawText);
    copyJsonBtn.textContent = "Copied";
    setTimeout(() => {
      copyJsonBtn.textContent = originalLabel;
    }, 1400);
  } catch (error) {
    setStatus("Could not copy the JSON automatically.", "error");
  }
}

function appendRawText(streamState, chunkText) {
  streamState.fullRawText += chunkText;

  if (streamState.rawStreamingEnabled) {
    streamState.rawText += chunkText;
    if (streamState.rawText.length > MAX_RAW_TEXT_BYTES) {
      streamState.rawText = streamState.rawText.slice(streamState.rawText.length - MAX_RAW_TEXT_BYTES);
      streamState.rawTextTruncated = true;
    }
  }
}

function renderRawText(streamState) {
  if (!streamState) {
    jsonOutputEl.textContent =
      "Live raw event rendering is paused for performance.\n\n" +
      "Structured preview is still updating while the stream is being read.\n" +
      "Use 'Enable Live Raw' only when you need to inspect the event lines.";
    return;
  }

  if (!streamState.rawStreamingEnabled && !streamState.completed) {
    jsonOutputEl.textContent =
      "Live raw event rendering is paused for performance.\n\n" +
      "Structured preview is still updating while the stream is being read.\n" +
      "Use 'Enable Live Raw' only when you need to inspect the event lines.";
    return;
  }

  jsonOutputEl.textContent = streamState.rawTextTruncated
    ? "(Showing only the latest raw stream window)\n\n" + streamState.rawText
    : streamState.rawText;
  jsonOutputEl.scrollTop = jsonOutputEl.scrollHeight;
}

function toggleRawStreaming() {
  const streamState = activeStreamState || lastStreamState;
  if (!streamState) {
    return;
  }

  streamState.rawStreamingEnabled = !streamState.rawStreamingEnabled;

  if (streamState.rawStreamingEnabled) {
    streamState.rawText = streamState.fullRawText.slice(-MAX_RAW_TEXT_BYTES);
    streamState.rawTextTruncated = streamState.fullRawText.length > MAX_RAW_TEXT_BYTES;
  } else {
    streamState.rawText = streamState.completed
      ? streamState.fullRawText.slice(-MAX_RAW_TEXT_BYTES)
      : "";
  }

  updateRawPanelState(streamState);
  renderRawText(streamState);
}

function updateRawPanelState(streamState) {
  const enabled = !!(streamState && streamState.rawStreamingEnabled);
  const completed = !!(streamState && streamState.completed);
  toggleRawBtn.textContent = enabled
    ? "Pause Live Raw"
    : completed
      ? "Show Stored Raw"
      : "Enable Live Raw";
  toggleRawBtn.classList.toggle("active", enabled);
}

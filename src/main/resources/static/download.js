// ── Endpoints ────────────────────────────────────────────────────────────────
const ENDPOINTS = {
  // CSV — auto headers
  csvStream        : '/api/users/download/csv',                         // stream — zero heap
  csvBytes         : '/api/users/download/csv/bytes',                   // byte[] — Content-Length
  csvBase64        : '/api/users/download/csv/base64',                  // base64 — JSON body
  csvFile          : '/api/users/download/csv/file',                    // server disk write
  // CSV — custom headers
  csvCustom        : '/api/users/download/csv/custom-headers',          // custom — stream
  csvCustomBytes   : '/api/users/download/csv/custom-headers/bytes',    // custom — bytes
  csvCustomBase64  : '/api/users/download/csv/custom-headers/base64',   // custom — base64
  csvCustomFile    : '/api/users/download/csv/custom-headers/file',     // custom — server file
  // Excel — auto headers
  excel            : '/api/users/download/xlsx',                        // Excel stream
  excelBytes       : '/api/users/download/xlsx/bytes',                  // Excel bytes
  excelBase64      : '/api/users/download/xlsx/base64',                 // Excel base64
  excelFile        : '/api/users/download/xlsx/file',                   // Excel server file
  // Excel — custom headers
  excelCustom      : '/api/users/download/xlsx/custom-headers',         // Excel custom — stream
  excelCustomBytes : '/api/users/download/xlsx/custom-headers/bytes',   // Excel custom — bytes
  excelCustomBase64: '/api/users/download/xlsx/custom-headers/base64',  // Excel custom — base64
  excelCustomFile  : '/api/users/download/xlsx/custom-headers/file',    // Excel custom — file
};

// ── UI helpers ────────────────────────────────────────────────────────────────
function uiBusy(btnId, busy) {
  const btn = document.getElementById(btnId);
  if (btn) btn.disabled = busy;
  const p = document.getElementById('progress');
  if (busy) {
    p.style.display = 'block';
    p.removeAttribute('value'); // indeterminate spinner
  } else {
    setTimeout(() => { p.style.display = 'none'; }, 2000);
  }
}
function setStatus(msg) { document.getElementById('status').textContent = msg; }
function formatBytes(n) {
  if (n < 1024)    return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
  return (n / 1048576).toFixed(2) + ' MB';
}

// ── Shared: fetch → Blob → save ───────────────────────────────────────────────
// Used by Stream, Bytes, and Excel — all return a binary file body.
async function downloadBinaryFile(apiUrl, defaultFileName, btnId) {
  uiBusy(btnId, true);
  setStatus('⏳ Generating on server…');
  try {
    const res = await fetch(apiUrl);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    setStatus('⬇ Downloading…');
    const blob = await res.blob();
    const cd   = res.headers.get('Content-Disposition') || '';
    const m    = cd.match(/filename[^;=\n]*=["']?([^"';\n]+)/i);
    triggerDownload(blob, m ? m[1].trim() : defaultFileName);
    setStatus(`✅ Done! (${formatBytes(blob.size)})`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Method 4: streamCsvToResponse ─────────────────────────────────────────────
// Server pipes rows straight to socket — zero intermediate heap buffer.
// Content-Length is NOT set (size unknown until fully written).
function downloadCsvStream() {
  downloadBinaryFile(ENDPOINTS.csvStream, 'users.csv', 'csvStreamBtn');
}

// ── Method 2: generateCsvAsBytes ──────────────────────────────────────────────
// Server builds byte[] in memory then sends it.
// Content-Length IS set → browser shows exact download-progress %.
function downloadCsvBytes() {
  downloadBinaryFile(ENDPOINTS.csvBytes, 'users.csv', 'csvBytesBtn');
}

// ── Method 3: generateCsvAsBase64 ─────────────────────────────────────────────
// Response is JSON: { fileName: "users.csv", data: "<base64>" }
// We decode Base64 in the browser into a Uint8Array Blob — no extra HTTP request.
async function downloadCsvBase64() {
  const btnId = 'csvBase64Btn';
  uiBusy(btnId, true);
  setStatus('⏳ Generating Base64 on server…');
  try {
    const res  = await fetch(ENDPOINTS.csvBase64);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    setStatus('🔄 Decoding Base64 in browser…');
    const json   = await res.json();
    const binary = atob(json.data);                // Base64 → binary string
    const buf    = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) buf[i] = binary.charCodeAt(i);
    const blob   = new Blob([buf], { type: 'text/csv' });
    triggerDownload(blob, json.fileName || 'users.csv');
    setStatus(`✅ Done! (${formatBytes(blob.size)})`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Method 1: generateCsvToFile ───────────────────────────────────────────────
// No file download — triggers a server-side write to disk.
// Response is JSON: { method, message, path }
async function saveCsvToFile() {
  const btnId = 'csvFileBtn';
  uiBusy(btnId, true);
  setStatus('⏳ Writing file on server disk…');
  try {
    const res  = await fetch(ENDPOINTS.csvFile);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    const json = await res.json();
    setStatus(`✅ ${json.message}\n📁 ${json.path}`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Custom Headers: streamCsvToResponse with custom column list ──────────────
// Headers: id, firstName, email, salary, department, createdDate
// salary and department have no matching field in UserRecord → blank columns
function downloadCsvCustomHeaders() {
  downloadBinaryFile(ENDPOINTS.csvCustom, 'users_custom.csv', 'csvCustomBtn');
}

// ── Excel: streamExcelToResponse ──────────────────────────────────────────────
function downloadExcel() {
  downloadBinaryFile(ENDPOINTS.excel, 'users.xlsx', 'excelBtn');
}

// ── Excel: generateExcelAsBytes ───────────────────────────────────────────
function downloadExcelBytes() {
  downloadBinaryFile(ENDPOINTS.excelBytes, 'users.xlsx', 'excelBytesBtn');
}

// ── Excel: generateExcelAsBase64 ──────────────────────────────────────────
async function downloadExcelBase64() {
  const btnId = 'excelBase64Btn';
  uiBusy(btnId, true);
  setStatus('⏳ Generating Base64 Excel on server…');
  try {
    const res  = await fetch(ENDPOINTS.excelBase64);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    setStatus('🔄 Decoding Base64 in browser…');
    const json   = await res.json();
    const binary = atob(json.data);
    const buf    = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) buf[i] = binary.charCodeAt(i);
    const blob   = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    triggerDownload(blob, json.fileName || 'users.xlsx');
    setStatus(`✅ Done! (${formatBytes(blob.size)})`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Excel: generateExcel (save to server disk) ────────────────────────────
async function saveExcelToFile() {
  const btnId = 'excelFileBtn';
  uiBusy(btnId, true);
  setStatus('⏳ Writing Excel file on server disk…');
  try {
    const res  = await fetch(ENDPOINTS.excelFile);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    const json = await res.json();
    setStatus(`✅ ${json.message}\n📁 ${json.path}`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Excel: custom headers ─────────────────────────────────────────────────
function downloadExcelCustomHeaders() {
  downloadBinaryFile(ENDPOINTS.excelCustom, 'users_custom.xlsx', 'excelCustomBtn');
}

// ── CSV custom headers — bytes ────────────────────────────────────────────────
function downloadCsvCustomHeadersBytes() {
  downloadBinaryFile(ENDPOINTS.csvCustomBytes, 'users_custom.csv', 'csvCustomBytesBtn');
}

// ── CSV custom headers — base64 ───────────────────────────────────────────────
async function downloadCsvCustomHeadersBase64() {
  const btnId = 'csvCustomBase64Btn';
  uiBusy(btnId, true);
  setStatus('⏳ Generating Base64 CSV (custom headers) on server…');
  try {
    const res  = await fetch(ENDPOINTS.csvCustomBase64);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    setStatus('🔄 Decoding Base64 in browser…');
    const json   = await res.json();
    const binary = atob(json.data);
    const buf    = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) buf[i] = binary.charCodeAt(i);
    const blob   = new Blob([buf], { type: 'text/csv' });
    triggerDownload(blob, json.fileName || 'users_custom.csv');
    setStatus(`✅ Done! (${formatBytes(blob.size)})`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── CSV custom headers — server file ─────────────────────────────────────────
async function saveCsvCustomHeadersToFile() {
  const btnId = 'csvCustomFileBtn';
  uiBusy(btnId, true);
  setStatus('⏳ Writing CSV (custom headers) to server disk…');
  try {
    const res  = await fetch(ENDPOINTS.csvCustomFile);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    const json = await res.json();
    setStatus(`✅ ${json.message}\n📁 ${json.path}`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Excel custom headers — bytes ──────────────────────────────────────────────
function downloadExcelCustomHeadersBytes() {
  downloadBinaryFile(ENDPOINTS.excelCustomBytes, 'users_custom.xlsx', 'excelCustomBytesBtn');
}

// ── Excel custom headers — base64 ─────────────────────────────────────────────
async function downloadExcelCustomHeadersBase64() {
  const btnId = 'excelCustomBase64Btn';
  uiBusy(btnId, true);
  setStatus('⏳ Generating Base64 Excel (custom headers) on server…');
  try {
    const res  = await fetch(ENDPOINTS.excelCustomBase64);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    setStatus('🔄 Decoding Base64 in browser…');
    const json   = await res.json();
    const binary = atob(json.data);
    const buf    = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) buf[i] = binary.charCodeAt(i);
    const blob   = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
    triggerDownload(blob, json.fileName || 'users_custom.xlsx');
    setStatus(`✅ Done! (${formatBytes(blob.size)})`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Excel custom headers — server file ────────────────────────────────────────
async function saveExcelCustomHeadersToFile() {
  const btnId = 'excelCustomFileBtn';
  uiBusy(btnId, true);
  setStatus('⏳ Writing Excel (custom headers) to server disk…');
  try {
    const res  = await fetch(ENDPOINTS.excelCustomFile);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    const json = await res.json();
    setStatus(`✅ ${json.message}\n📁 ${json.path}`);
    document.getElementById('progress').value = 100;
  } catch (e) {
    setStatus('❌ Error: ' + e.message);
    console.error(e);
  } finally {
    uiBusy(btnId, false);
  }
}

// ── Utility ───────────────────────────────────────────────────────────────────
function triggerDownload(blob, fileName) {
  const url  = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url; link.download = fileName;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url); // free memory immediately
}

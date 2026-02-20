/**
 * download.js
 *
 * Calls GET /api/users/download/csv and triggers a file-save in the browser.
 *
 * Two approaches are shown:
 *
 *  downloadCsv()        — fetch → Blob → URL.createObjectURL  (recommended)
 *  downloadCsvViaLink() — simple anchor tag redirect           (fallback)
 *
 * The Blob approach gives you UI feedback (spinner, done message) and works
 * with any CORS-enabled server.  The anchor approach is one line but gives
 * no progress feedback.
 */

const CSV_API_URL   = '/api/users/download/csv';
const EXCEL_API_URL = '/api/users/download/xlsx';

// ---------------------------------------------------------------------------
// Shared fetch + Blob helper  (used by both CSV and Excel buttons)
// ---------------------------------------------------------------------------
// How it works:
//   1. fetch() opens a streaming HTTP GET to the server.
//   2. response.blob() collects all chunks into a Blob in browser memory.
//   3. URL.createObjectURL() turns the Blob into a temporary local URL.
//   4. A hidden <a> click triggers the browser's native Save dialog.
//   5. The temporary URL is revoked immediately to free memory.

async function downloadFile(apiUrl, defaultFileName, btnId) {
  const btn      = document.getElementById(btnId);
  const status   = document.getElementById('status');
  const progress = document.getElementById('progress');

  btn.disabled           = true;
  progress.style.display = 'block';
  progress.removeAttribute('value'); // indeterminate spinner
  status.textContent     = '⏳ Generating file on server…';

  try {
    const response = await fetch(apiUrl);

    if (!response.ok) {
      throw new Error(`Server error: ${response.status} ${response.statusText}`);
    }

    status.textContent = '⬇ Downloading…';

    // Collect the streamed response into a Blob
    const blob = await response.blob();

    // Build a filename from the Content-Disposition header if present
    const disposition = response.headers.get('Content-Disposition') || '';
    const match       = disposition.match(/filename[^;=\n]*=["']?([^"';\n]+)/i);
    const fileName    = match ? match[1].trim() : defaultFileName;

    triggerDownload(blob, fileName);

    status.textContent = '✅ Download complete!';
    progress.value = 100;

  } catch (err) {
    console.error('Download failed:', err);
    status.textContent = `❌ Error: ${err.message}`;
  } finally {
    btn.disabled = false;
    setTimeout(() => { progress.style.display = 'none'; }, 2000);
  }
}

/** Thin wrappers — one per export type */
function downloadCsv()   { downloadFile(CSV_API_URL,   'users.csv',  'csvBtn');   }
function downloadExcel() { downloadFile(EXCEL_API_URL, 'users.xlsx', 'excelBtn'); }

/** Creates a hidden anchor, clicks it to start the download, then removes it. */
function triggerDownload(blob, fileName) {
  const url  = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href     = url;
  link.download = fileName;
  link.style.display = 'none';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  // Revoke the object URL to release memory
  URL.revokeObjectURL(url);
}

// ---------------------------------------------------------------------------
// Approach 2 — simple anchor redirect  (FALLBACK / alternative)
// ---------------------------------------------------------------------------
// Works for same-origin requests or when CORS allows it.
// The browser handles the download entirely — no JS blob needed.
// No loading feedback, but zero code complexity.
//
//   <a href="/api/users/download/csv" download="users.csv">Download</a>
//
// Or trigger it programmatically:
//
//   function downloadCsvViaLink() {
//     const a = document.createElement('a');
//     a.href     = API_URL;
//     a.download = 'users.csv';
//     a.click();
//   }

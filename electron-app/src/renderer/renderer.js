// @ts-check

let currentRecords = [];
let currentFilename = '';

document.addEventListener('DOMContentLoaded', () => {

// --- Clock ---
function updateClock() {
  const now = new Date();
  const pad = n => String(n).padStart(2, '0');
  const d = `${pad(now.getMonth() + 1)}/${pad(now.getDate())}/${now.getFullYear()}`;
  const t = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
  document.getElementById('clock').textContent = `${d} ${t}`;
}
updateClock();
setInterval(updateClock, 1000);

// --- Drag and Drop ---
const zone = document.getElementById('uploadZone');

zone.addEventListener('dragover', e => {
  e.preventDefault();
  zone.classList.add('dragover');
});
zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
zone.addEventListener('drop', e => {
  e.preventDefault();
  zone.classList.remove('dragover');
  if (e.dataTransfer.files.length) {
    const file = e.dataTransfer.files[0];
    const filePath = window.electronAPI.getPathForFile(file);
    if (filePath) {
      handleFilePath(filePath, file.name);
    }
  }
});

// Click on upload zone opens file dialog
zone.addEventListener('click', (e) => {
  if (e.target.tagName !== 'SPAN') {
    openFile();
  }
});

// --- Native File Dialog ---
async function openFile() {
  const filePath = await window.electronAPI.openFileDialog();
  if (filePath) {
    const name = filePath.split(/[\\/]/).pop() || 'report.pdf';
    handleFilePath(filePath, name);
  }
}

// --- File Processing ---
async function handleFilePath(filePath, filename) {
  if (!filename.toLowerCase().endsWith('.pdf')) {
    showError('PDF files only.');
    return;
  }

  currentFilename = filename;
  hideError();
  document.getElementById('uploadZone').style.display = 'none';
  document.getElementById('loading').classList.add('active');

  try {
    const records = await window.electronAPI.extractPdf(filePath);
    document.getElementById('loading').classList.remove('active');

    if (!records || records.length === 0) {
      showError('No shipment records found in this PDF.');
      document.getElementById('uploadZone').style.display = '';
      return;
    }

    currentRecords = records;
    renderResults(records);
  } catch (err) {
    document.getElementById('loading').classList.remove('active');
    showError('Processing error: ' + err.message);
    document.getElementById('uploadZone').style.display = '';
  }
}

// --- Render Results ---
function renderResults(records) {
  const tbody = document.getElementById('tbody');
  tbody.innerHTML = '';

  records.forEach((r, i) => {
    const tr = document.createElement('tr');
    if (r.Status === 'VOID') tr.classList.add('void');

    tr.innerHTML = `
      <td>${i + 1}</td>
      <td>${esc(r['Package Ref No.1'])}</td>
      <td class="tracking" onclick="copyTracking(this, '${esc(r['Tracking No.'])}')">${esc(r['Tracking No.'])}</td>
      <td class="status-cell">${r.Status === 'VOID' ? 'VOID' : ''}</td>
    `;
    tbody.appendChild(tr);
  });

  const active = records.filter(r => r.Status === 'Active').length;
  const voided = records.filter(r => r.Status === 'VOID').length;

  document.getElementById('fileName').textContent = currentFilename;
  document.getElementById('summary').textContent =
    `${records.length} PKGS | ${active} ACTIVE | ${voided} VOID`;

  document.getElementById('statusLine').classList.add('active');
  document.getElementById('tableWrap').classList.add('active');
  document.getElementById('downloadBar').classList.add('active');
  document.getElementById('btnCsv').style.display = '';
  document.getElementById('btnXlsx').style.display = '';
  document.getElementById('btnCopy').style.display = '';
  document.getElementById('btnReset').style.display = '';
  document.getElementById('btnDbExport').style.display = '';
}

// --- Clipboard ---
window.copyTracking = function(td, text) {
  navigator.clipboard.writeText(text);
  const tip = document.createElement('span');
  tip.className = 'copied-tooltip';
  tip.textContent = 'COPIED';
  td.appendChild(tip);
  setTimeout(() => tip.remove(), 1000);
};

function copyAll() {
  const lines = currentRecords.map(r =>
    `${r['Package Ref No.1']}\t${r['Tracking No.']}\t${r.Status}`
  );
  const header = 'Package Ref No.1\tTracking No.\tStatus';
  navigator.clipboard.writeText(header + '\n' + lines.join('\n'));
  showToast('Copied all records to clipboard', 'success', 2000);
}

// --- Download ---
async function download(fmt) {
  if (!currentRecords.length) return;

  const defaultName = currentFilename.replace('.pdf', '.' + fmt);

  try {
    if (fmt === 'csv') {
      const header = '"Package Ref No.1","Tracking No.","Status"';
      const rows = currentRecords.map(r => {
        const ref = r['Package Ref No.1'].replace(/"/g, '""');
        const tracking = r['Tracking No.'].replace(/"/g, '""');
        const status = r.Status.replace(/"/g, '""');
        return `"${ref}","${tracking}","${status}"`;
      });
      const csv = [header, ...rows].join('\n');
      const saved = await window.electronAPI.saveCsv(csv, defaultName);
      if (saved) showToast('CSV saved: ' + saved.split(/[\\/]/).pop(), 'success');
    } else {
      const saved = await window.electronAPI.saveXlsx(currentRecords, defaultName);
      if (saved) showToast('Excel saved: ' + saved.split(/[\\/]/).pop(), 'success');
    }
  } catch (err) {
    showToast('Save error: ' + err.message, 'error');
  }
}

// --- DB Export ---
async function exportDb() {
  if (!currentRecords.length) return;

  const btn = document.getElementById('btnDbExport');
  btn.classList.add('active');
  btn.textContent = 'F8 EXPORTING...';

  try {
    const data = await window.electronAPI.exportDb(currentRecords, currentFilename);
    if (data.error) {
      showToast(data.error, 'error');
    } else {
      showToast(`DB EXPORT: ${data.inserted} inserted, ${data.skipped} skipped`, 'success');
    }
  } catch (err) {
    showToast('DB connection error: ' + err.message, 'error');
  } finally {
    btn.classList.remove('active');
    btn.textContent = 'F8 DB EXPORT';
  }
}

// --- Reset ---
function resetUpload() {
  document.getElementById('uploadZone').style.display = '';
  document.getElementById('tableWrap').classList.remove('active');
  document.getElementById('statusLine').classList.remove('active');
  document.getElementById('downloadBar').classList.remove('active');
  document.getElementById('loading').classList.remove('active');
  document.getElementById('btnCsv').style.display = 'none';
  document.getElementById('btnXlsx').style.display = 'none';
  document.getElementById('btnCopy').style.display = 'none';
  document.getElementById('btnReset').style.display = 'none';
  document.getElementById('btnDbExport').style.display = 'none';
  hideError();
  currentRecords = [];
}

// --- Helpers ---
function showError(msg) {
  const el = document.getElementById('errorMsg');
  el.textContent = '> ERROR: ' + msg;
  el.classList.add('active');
}

function hideError() {
  document.getElementById('errorMsg').classList.remove('active');
}

function esc(s) {
  const d = document.createElement('div');
  d.textContent = s;
  return d.innerHTML;
}

function showToast(message, type = 'info', duration = 4000) {
  const toast = document.createElement('div');
  toast.className = 'toast ' + type;
  toast.textContent = '> ' + message;
  document.body.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s';
    setTimeout(() => toast.remove(), 300);
  }, duration);
}

// --- Keyboard Shortcuts ---
document.addEventListener('keydown', e => {
  if (e.key === 'F1') { e.preventDefault(); resetUpload(); }
  if (e.key === 'F2' && currentRecords.length) { e.preventDefault(); download('csv'); }
  if (e.key === 'F3' && currentRecords.length) { e.preventDefault(); download('xlsx'); }
  if (e.key === 'F4' && currentRecords.length) { e.preventDefault(); copyAll(); }
  if (e.key === 'F5') { e.preventDefault(); resetUpload(); }
  if (e.key === 'F8' && currentRecords.length) { e.preventDefault(); exportDb(); }
});

// --- Global function bindings for HTML onclick attributes ---
window.openFile = openFile;
window.resetUpload = resetUpload;
window.download = download;
window.copyAll = copyAll;
window.exportDb = exportDb;

}); // end DOMContentLoaded

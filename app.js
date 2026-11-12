/* ============================================
   StockFlow — Application Logic (app.js)
   ============================================ */

// ============ State ============
let products = [];
let stockHistory = []; // เก็บประวัติรอบการนับ (Snapshots)
let countLogs = [];    // เก็บประวัติการนับรายชิ้นอัตโนมัติ (Activity Logs)
let adminPin = CONFIG.DEFAULT_PIN;
let isAdmin = false;
let editingId = null;
let currentTab = 'all';
let searchTerm = '';
let pinBuffer = '';
let pinCallback = null;
let selectMode = false;
let selectedIds = new Set();
let pendingDeleteAction = null;

// History UI State
let historySubTab = 'sessions'; // 'sessions' | 'logs'
let historySearchTerm = '';
let historyDateFilter = '';
let viewingSessionId = null;
const logTimers = {};

// File System Access API
let fileHandle = null;
let fileName = '';
let fileHasPermission = false;

// กำหนดคีย์สำรองกรณีใน config.js ยังไม่ได้เพิ่มคีย์ประวัติ
const HISTORY_STORAGE_KEY = (CONFIG && CONFIG.HISTORY_KEY) || 'stockflow_history_v1';
const LOG_STORAGE_KEY = (CONFIG && CONFIG.LOG_KEY) || 'stockflow_logs_v1';
const CAT_COL = (CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.CATEGORY) || 'หมวดหมู่';
const DEFAULT_CATEGORY = 'ไม่ระบุ';

// ============ Storage ============
function loadData() {
  try { products = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '[]'); } catch (e) { products = []; }
  try { stockHistory = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || '[]'); } catch (e) { stockHistory = []; }
  try { countLogs = JSON.parse(localStorage.getItem(LOG_STORAGE_KEY) || '[]'); } catch (e) { countLogs = []; }
  try { adminPin = localStorage.getItem(CONFIG.ADMIN_KEY) || CONFIG.DEFAULT_PIN; } catch (e) { adminPin = CONFIG.DEFAULT_PIN; }

  // ล้างไฟล์ เวิร์กบุ๊ค1.xlsx หรือไฟล์ที่เคยจำค้างไว้ใน IndexedDB ออกทันที
  clearCachedExcelFile();

  updateFileBar();
  updateAdminUI();
  render();
  renderHistoryPage();
}

function saveDataToCache() {
  localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
}

function saveHistoryToCache() {
  localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(stockHistory));
  localStorage.setItem(LOG_STORAGE_KEY, JSON.stringify(countLogs));
}

// ============ File System Access API ============
function isFileSystemSupported() {
  return 'showOpenFilePicker' in window;
}

async function openExcelFile() {
  if (!isFileSystemSupported()) {
    alert('⚠️ เบราว์เซอร์ของคุณไม่รองรับการบันทึกทับไฟล์โดยตรง\nกรุณาใช้ Chrome หรือ Edge เวอร์ชันใหม่');
    fallbackImportFile();
    return;
  }

  try {
    const [handle] = await window.showOpenFilePicker({
      types: [{
        description: 'Excel Files',
        accept: {
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
          'application/vnd.ms-excel': ['.xls']
        }
      }],
      multiple: false
    });

    fileHandle = handle;
    fileName = handle.name;
    fileHasPermission = true;

    const file = await handle.getFile();
    await parseExcelFile(file);
    updateFileBar();
    showToast(`📂 เชื่อมต่อ ${fileName} แล้ว`);
  } catch (err) {
    if (err.name === 'AbortError') return;
    console.error(err);
    alert('ไม่สามารถเปิดไฟล์ได้: ' + err.message);
  }
}

async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(firstSheet);

        if (jsonData.length === 0) {
          alert('ไม่พบข้อมูลในไฟล์ Excel');
          reject(new Error('No data'));
          return;
        }

        const c = CONFIG.EXCEL_COLUMNS;
        products = jsonData.map((row) => ({
          id: row[c.ID] || genId(),
          name: row[c.NAME] || 'ไม่มีชื่อ',
          sku: row[c.SKU] || '',
          category: String(row[CAT_COL] || '').trim(),
          system: parseInt(row[c.SYSTEM]) || 0,
          counted: (row[c.COUNTED] !== undefined && row[c.COUNTED] !== '') ? parseInt(row[c.COUNTED]) : null
        }));

        saveDataToCache();
        render();
        resolve();
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

async function saveToFile() {
  if (!fileHandle) {
    alert('กรุณาเปิดไฟล์ Excel ก่อนบันทึก');
    return;
  }

  const saveBtn = document.getElementById('saveBtn');
  const originalHtml = saveBtn.innerHTML;
  saveBtn.disabled = true;
  saveBtn.textContent = '⏳ กำลังบันทึก...';

  try {
    const permission = await fileHandle.queryPermission({ mode: 'readwrite' });
    if (permission !== 'granted') {
      const request = await fileHandle.requestPermission({ mode: 'readwrite' });
      if (request !== 'granted') {
        alert('ไม่ได้รับอนุญาตให้บันทึกไฟล์');
        return;
      }
    }

    const c = CONFIG.EXCEL_COLUMNS;
    const exportData = products.map(p => {
      const d = getDiff(p);
      let status = 'ยังไม่นับ';
      if (d === 0) status = 'ตรง';
      else if (d !== null && d < 0) status = 'ขาด ' + Math.abs(d);
      else if (d !== null && d > 0) status = 'เกิน ' + d;

      return {
        [c.ID]: p.id,
        [c.NAME]: p.name,
        [c.SKU]: p.sku || '',
        [CAT_COL]: p.category || '',
        [c.SYSTEM]: p.system,
        [c.COUNTED]: (p.counted === null || p.counted === '') ? '' : p.counted,
        [c.DIFF]: d === null ? '' : d,
        [c.STATUS]: status,
        [c.UPDATED_AT]: new Date().toLocaleString('th-TH')
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, CONFIG.SHEET_NAME);

    const writable = await fileHandle.createWritable();
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    await writable.write(wbout);
    await writable.close();

    // บันทึกเข้าประวัติอัตโนมัติเมื่อกดเซฟลงไฟล์ Excel
    createHistorySession(`บันทึกไฟล์ (${fileName || 'Excel'})`, true);

    const statusEl = document.getElementById('saveStatus');
    statusEl.classList.add('success');
    statusEl.textContent = '✅ บันทึกสำเร็จ ' + new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
    showToast('💾 บันทึกทับไฟล์เดิมและเก็บประวัติเรียบร้อย');
    setTimeout(() => updateSaveBar(), 3000);
  } catch (err) {
    console.error(err);
    alert('ไม่สามารถบันทึกไฟล์ได้: ' + err.message);
  } finally {
    saveBtn.disabled = false;
    saveBtn.innerHTML = originalHtml;
  }
}

function updateFileBar() {
  const fileBar = document.getElementById('fileBar');
  const fileNameEl = document.getElementById('fileName');
  const fileSub = document.getElementById('fileSub');
  const fileIcon = document.getElementById('fileIcon');
  const openBtn = document.getElementById('openFileBtn');
  const saveBar = document.getElementById('saveBar');
  const setFileNameEl = document.getElementById('setFileName');

  if (fileHandle || fileName) {
    if (fileBar) fileBar.style.display = 'flex';
    if (fileNameEl) fileNameEl.textContent = fileName;
    if (setFileNameEl) setFileNameEl.textContent = fileName;
    if (fileSub) {
      fileSub.textContent = 'อัปโหลดสำเร็จ • พร้อมใช้งาน';
      fileSub.classList.add('connected');
    }
    if (fileIcon) {
      fileIcon.classList.add('connected');
      fileIcon.textContent = '✓';
    }
    if (openBtn) openBtn.textContent = '🔄 เปลี่ยนไฟล์';
    if (saveBar) saveBar.style.display = fileHandle ? 'flex' : 'none';
    updateSaveBar();
  } else {
    // ซ่อนแถบไฟล์เมื่อไม่มีการเชื่อมต่อไฟล์ เพื่อไม่ให้มีชื่อไฟล์ค้าง
    if (fileBar) fileBar.style.display = 'none';
    if (fileNameEl) fileNameEl.textContent = 'ยังไม่ได้เชื่อมต่อไฟล์';
    if (setFileNameEl) setFileNameEl.textContent = 'ยังไม่ได้เชื่อมต่อไฟล์';
    if (fileSub) {
      fileSub.textContent = 'กรุณาเลือกไฟล์ Excel เพื่อเริ่มต้น';
      fileSub.classList.remove('connected');
    }
    if (fileIcon) {
      fileIcon.classList.remove('connected');
      fileIcon.textContent = '📄';
    }
    if (openBtn) openBtn.textContent = '📂 เปิดไฟล์';
    if (saveBar) saveBar.style.display = 'none';
  }
}

function updateSaveBar() {
  const status = document.getElementById('saveStatus');
  if (!status) return;
  status.classList.remove('success');
  if (products.length === 0) {
    status.textContent = 'ไม่มีข้อมูล';
  } else {
    const counted = products.filter(p => p.counted !== null && p.counted !== '' && p.counted !== undefined).length;
    status.innerHTML = `นับแล้ว <strong>${counted}</strong> / ${products.length} รายการ`;
  }
}

// ล้างไฟล์ที่จำค้างอยู่ใน IndexedDB ออกทั้งหมด
function clearCachedExcelFile() {
  fileHandle = null;
  fileName = '';
  fileHasPermission = false;
  try {
    if ('indexedDB' in window) {
      indexedDB.deleteDatabase(CONFIG.DB_NAME || 'StockFlowDB');
    }
  } catch (e) {
    console.error('Clear IndexedDB error:', e);
  }
}

// ฟังก์ชันกดปุ่ม ✕ เพื่อเอาไฟล์ Excel ออก
function removeExcelFile() {
  fileHandle = null;
  fileName = '';
  fileHasPermission = false;
  clearCachedExcelFile();
  updateFileBar();
  showToast('🗑️ เอาไฟล์ Excel ออกแล้ว');
}

function fallbackImportFile() {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xls';
  input.onchange = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    fileName = file.name;
    fileHandle = null;
    parseExcelFile(file).then(() => {
      updateFileBar();
      showToast('📥 นำเข้าข้อมูลแล้ว (โหมดสำรอง)');
    });
  };
  input.click();
}

// ============ Utils ============
function genId() { return Date.now().toString(36) + Math.random().toString(36).slice(2, 6); }

function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 2500);
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
}

function getDiff(p) {
  if (p.counted === null || p.counted === undefined || p.counted === '') return null;
  return p.counted - p.system;
}

function getStatusText(d) {
  if (d === null || d === undefined || d === '') return 'ยังไม่นับ';
  if (d === 0) return 'ตรง';
  if (d < 0) return 'ขาด ' + Math.abs(d);
  return 'เกิน ' + d;
}

function formatDateTimeTH(isoStr) {
  if (!isoStr) return '-';
  const d = new Date(isoStr);
  if (isNaN(d.getTime())) return isoStr;
  return

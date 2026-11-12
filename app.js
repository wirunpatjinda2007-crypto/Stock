/* ============================================
   StockFlow — Application Logic (app.js)
   ============================================ */

// ============ State ============
var products = [];
var stockHistory = []; // เก็บประวัติรอบการนับ (Snapshots)
var countLogs = [];    // เก็บประวัติการนับรายชิ้นอัตโนมัติ (Activity Logs)
var adminPin = (typeof CONFIG !== 'undefined' && CONFIG.DEFAULT_PIN) ? CONFIG.DEFAULT_PIN : '1234';
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
var viewingSessionId = null;
const logTimers = {};

// File System Access API
let fileHandle = null;
let fileName = '';
let fileHasPermission = false;

// กำหนดคีย์สำรองกรณีใน config.js ยังไม่ได้เพิ่มคีย์ประวัติ
const HISTORY_STORAGE_KEY = (typeof CONFIG !== 'undefined' && CONFIG.HISTORY_KEY) || 'stockflow_history_v1';
const LOG_STORAGE_KEY = (typeof CONFIG !== 'undefined' && CONFIG.LOG_KEY) || 'stockflow_logs_v1';
const CAT_COL = (typeof CONFIG !== 'undefined' && CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.CATEGORY) || 'หมวดหมู่';
const DEFAULT_CATEGORY = 'ไม่ระบุ';

// ============ Storage ============
function loadData() {
  try { products = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '[]'); } catch (e) { products = []; }
  try { stockHistory = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || '[]'); } catch (e) { stockHistory = []; }
  try { countLogs = JSON.parse(localStorage.getItem(LOG_STORAGE_KEY) || '[]'); } catch (e) { countLogs = []; }
  try { adminPin = localStorage.getItem(CONFIG.ADMIN_KEY) || CONFIG.DEFAULT_PIN; } catch (e) { adminPin = CONFIG.DEFAULT_PIN; }

  window.products = products;

  // ล้างไฟล์ที่เคยจำค้างไว้ใน IndexedDB ออกทันที
  clearCachedExcelFile();

  updateFileBar();
  updateAdminUI();
  render();
  renderHistoryPage();
}

function saveDataToCache() {
  window.products = products;
  localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
}

function saveHistoryToCache() {
  localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(stockHistory));
  localStorage.setItem(LOG_STORAGE_KEY, JSON.stringify(countLogs));
}

function saveAdminPin(newPin) {
  adminPin = newPin;
  localStorage.setItem(CONFIG.ADMIN_KEY, adminPin);
}

// ============ File System Access API & Excel Export ============
function isFileSystemSupported() {
  return 'showOpenFilePicker' in window;
}

async function openExcelFile() {
  if (!isFileSystemSupported()) {
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
        products = jsonData.map((row) => {
          const rawCounted = row[c.COUNTED];
          return {
            id: String(row[c.ID] || genId()),
            name: String(row[c.NAME] || row['รายการ'] || row['Name'] || 'ไม่มีชื่อ').trim(),
            sku: String(row[c.SKU] || row['รหัสสินค้า'] || '').trim(),
            category: String(row[CAT_COL] || row['Category'] || '').trim(),
            system: parseInt(row[c.SYSTEM] ?? row['จำนวน'] ?? 0, 10) || 0,
            counted: (rawCounted !== undefined && rawCounted !== null && String(rawCounted).trim() !== '')
              ? parseInt(rawCounted, 10)
              : null
          };
        });

        saveDataToCache();
        render();
        resolve();
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// สร้างข้อมูล Array สำหรับส่งออกเป็น Excel
function buildStockExportRows(itemList = products) {
  const c = CONFIG.EXCEL_COLUMNS;
  return itemList.map(p => {
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
      [c.COUNTED]: (p.counted === null || p.counted === '' || p.counted === undefined) ? '' : p.counted,
      [c.DIFF]: d === null ? '' : d,
      [c.STATUS]: status,
      [c.UPDATED_AT]: new Date().toLocaleString('th-TH')
    };
  });
}

// บันทึกข้อมูลจากหน้า #check (หากเชื่อมไฟล์ไว้จะเซฟทับไฟล์เดิม หากไม่ได้เชื่อมจะดาวน์โหลดเป็น .xlsx ทันที)
async function saveToFile() {
  if (!products || products.length === 0) {
    showToast('⚠️ ยังไม่มีรายการสินค้าให้บันทึก');
    return;
  }

  // กรณีไม่ได้เปิดไฟล์ผ่าน File System Access API (เช่น ใช้ผ่านเว็บ/มือถือ หรือซิงก์จาก Supabase) ให้ดาวน์โหลดไฟล์ .xlsx ทันที
  if (!fileHandle) {
    exportCurrentStockToExcel(true);
    return;
  }

  const saveBtn = document.getElementById('saveBtn');
  const originalHtml = saveBtn ? saveBtn.innerHTML : '';
  if (saveBtn) {
    saveBtn.disabled = true;
    saveBtn.textContent = '⏳ กำลังบันทึก...';
  }

  try {
    const permission = await fileHandle.queryPermission({ mode: 'readwrite' });
    if (permission !== 'granted') {
      const request = await fileHandle.requestPermission({ mode: 'readwrite' });
      if (request !== 'granted') {
        // ถ้าไม่ได้รับสิทธิ์เซฟทับ ให้ดาวน์โหลดเป็นไฟล์ใหม่แทน
        exportCurrentStockToExcel(true);
        return;
      }
    }

    const exportData = buildStockExportRows(products);
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, CONFIG.SHEET_NAME || 'Stock');

    const writable = await fileHandle.createWritable();
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    await writable.write(wbout);
    await writable.close();

    // บันทึกเข้าประวัติอัตโนมัติเมื่อกดเซฟลงไฟล์ Excel
    createHistorySession(`บันทึกทับไฟล์ (${fileName || 'Excel'})`, true);

    const statusEl = document.getElementById('saveStatus');
    if (statusEl) {
      statusEl.classList.add('success');
      statusEl.textContent = '✅ บันทึกสำเร็จ ' + new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
    }
    showToast('💾 บันทึกทับไฟล์ Excel และเก็บประวัติเรียบร้อย');
    setTimeout(() => updateSaveBar(), 3000);
  } catch (err) {
    console.error(err);
    // กรณีเบราว์เซอร์ติดปัญหาการเขียนทับไฟล์ ให้ดาวน์โหลดเป็นไฟล์ .xlsx แทนทันที
    exportCurrentStockToExcel(true);
  } finally {
    if (saveBtn) {
      saveBtn.disabled = false;
      saveBtn.innerHTML = originalHtml;
    }
  }
}

// ดาวน์โหลดสต๊อกปัจจุบันเป็นไฟล์ Excel (.xlsx) ได้จากทั้งหน้า #check และ #report
function exportCurrentStockToExcel(saveSession = false) {
  if (!products || products.length === 0) {
    showToast('⚠️ ยังไม่มีข้อมูลสินค้าให้ส่งออก');
    return;
  }

  try {
    const exportData = buildStockExportRows(products);
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, CONFIG.SHEET_NAME || 'Stock');

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    const baseName = fileName ? fileName.replace(/\.(xlsx|xls)$/i, '') : 'stockflow';
    const outName = `${baseName}_${stamp}.xlsx`;

    XLSX.writeFile(wb, outName);

    if (saveSession) {
      createHistorySession(`เซฟไฟล์ Excel (${outName})`, true);
    }

    const statusEl = document.getElementById('saveStatus');
    if (statusEl) {
      statusEl.classList.add('success');
      statusEl.textContent = '✅ ดาวน์โหลด Excel สำเร็จ ' + now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
      setTimeout(() => updateSaveBar(), 3000);
    }

    showToast(`📥 ดาวน์โหลดไฟล์ ${outName} เรียบร้อย`);
  } catch (err) {
    console.error(err);
    alert('ไม่สามารถสร้างไฟล์ Excel ได้: ' + err.message);
  }
}

function updateFileBar() {
  const fileBar = document.getElementById('fileBar');
  const fileNameEl = document.getElementById('fileName');
  const fileSub = document.getElementById('fileSub');
  const fileIcon = document.getElementById('fileIcon');
  const openBtn = document.getElementById('openFileBtn');
  const saveBar = document.getElementById('saveBar');
  const saveBtn = document.getElementById('saveBtn');
  const setFileNameEl = document.getElementById('setFileName');

  if (fileHandle || fileName) {
    if (fileBar) fileBar.style.display = 'flex';
    if (fileNameEl) fileNameEl.textContent = fileName;
    if (setFileNameEl) setFileNameEl.textContent = fileName;
    if (fileSub) {
      fileSub.textContent = fileHandle
        ? 'เชื่อมต่อไฟล์แล้ว • กดบันทึกทับไฟล์เดิมได้ทันที'
        : 'นำเข้าข้อมูลแล้ว • กดเซฟเป็นไฟล์ Excel ใหม่ได้';
      fileSub.classList.add('connected');
    }
    if (fileIcon) {
      fileIcon.classList.add('connected');
      fileIcon.textContent = '✓';
    }
    if (openBtn) openBtn.textContent = '🔄 เปลี่ยนไฟล์';
  } else {
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
    if (openBtn) openBtn.textContent = '📂 เปิดไฟล์ Excel';
  }

  // แสดงแถบ Save Bar เสมอเมื่อมีสินค้าหรือมีการเปิดไฟล์ เพื่อให้กดเซฟ Excel ในหน้า #check ได้ตลอดเวลา
  if (saveBar) {
    saveBar.style.display = (products.length > 0 || fileHandle || fileName) ? 'flex' : 'none';
  }
  if (saveBtn) {
    saveBtn.innerHTML = fileHandle
      ? '💾 บันทึกทับไฟล์ Excel'
      : '📥 เซฟเป็น Excel (.xlsx)';
  }
  updateSaveBar();
}

function updateSaveBar() {
  const status = document.getElementById('saveStatus');
  if (!status) return;
  status.classList.remove('success');
  if (products.length === 0) {
    status.textContent = 'ไม่มีข้อมูลสินค้า';
  } else {
    const counted = products.filter(p => p.counted !== null && p.counted !== '' && p.counted !== undefined).length;
    status.innerHTML = `นับแล้ว <strong>${counted}</strong> / ${products.length} รายการ`;
  }
}

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

function removeExcelFile() {
  fileHandle = null;
  fileName = '';
  fileHasPermission = false;
  clearCachedExcelFile();
  updateFileBar();
  showToast('🗑️ ยกเลิกการเชื่อมไฟล์ Excel แล้ว');
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
      showToast('📥 นำเข้าข้อมูลจาก Excel แล้ว');
    });
  };
  input.click();
}

// ============ Utils ============
function genId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
}

function showToast(msg) {
  const t = document.getElementById('toast');
  if (!t) return;
  t.textContent = msg;
  t.classList.add('show');
  clearTimeout(t._timer);
  t._timer = setTimeout(() => t.classList.remove('show'), 2500);
}

function escapeHtml(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  }[c]));
}

function getDiff(p) {
  if (p.counted === null || p.counted === undefined || p.counted === '') return null;
  return Number(p.counted) - Number(p.system || 0);
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
  return d.toLocaleString('th-TH', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
}

// ============ Render & Stock Counting ============
function render() {
  window.products = products;
  const list = document.getElementById('productList');
  if (!list) return;

  const total = products.length;
  const countedItems = products.filter(p => p.counted !== null && p.counted !== undefined && p.counted !== '');
  const pendingItems = products.filter(p => p.counted === null || p.counted === undefined || p.counted === '');
  const diffItems = products.filter(p => {
    const d = getDiff(p);
    return d !== null && d !== 0;
  });

  // Update Home & Tab Stats
  const setTxt = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  };
  setTxt('statTotal', total);
  setTxt('statCounted', countedItems.length);
  setTxt('statDiff', diffItems.length);
  setTxt('tabCountAll', total);
  setTxt('tabCountPending', pendingItems.length);
  setTxt('tabCountDone', countedItems.length);
  setTxt('tabCountDiff', diffItems.length);

  updateFileBar();

  // Filter items
  const q = searchTerm.trim().toLowerCase();
  const catSelect = document.getElementById('catFilter');
  const selectedCat = catSelect ? catSelect.value : 'all';

  const filtered = products.filter(p => {
    if (currentTab === 'pending' && (p.counted !== null && p.counted !== undefined && p.counted !== '')) return false;
    if (currentTab === 'done' && (p.counted === null || p.counted === undefined || p.counted === '')) return false;
    if (currentTab === 'diff') {
      const d = getDiff(p);
      if (d === null || d === 0) return false;
    }
    if (selectedCat && selectedCat !== 'all') {
      const pCat = (p.category || '').trim() || DEFAULT_CATEGORY;
      if (pCat !== selectedCat) return false;
    }
    if (q) {
      const matchName = (p.name || '').toLowerCase().includes(q);
      const matchSku = (p.sku || '').toLowerCase().includes(q);
      const matchCat = (p.category || '').toLowerCase().includes(q);
      if (!matchName && !matchSku && !matchCat) return false;
    }
    return true;
  });

  if (filtered.length === 0) {
    list.innerHTML = `
      <div class="empty-state" style="grid-column: 1 / -1;">
        <div class="icon">📦</div>
        <div class="title">${products.length === 0 ? 'ยังไม่มีรายการสินค้า' : 'ไม่พบสินค้าที่ค้นหา'}</div>
        <div class="desc">${products.length === 0 ? 'กดปุ่ม "เปิดไฟล์ Excel" ด้านบน หรือกดปุ่ม + (โหมดแอดมิน) เพื่อเพิ่มสินค้า' : 'ลองเปลี่ยนคำค้นหาหรือตัวกรองหมวดหมู่'}</div>
      </div>
    `;
    return;
  }

  list.innerHTML = filtered.map(p => {
    const d = getDiff(p);
    let cardState = 'plain';
    let badgeHtml = '';

    if (d === 0) {
      cardState = 'ok';
      badgeHtml = `<span class="diff-badge ok">✓ ครบ</span>`;
    } else if (d !== null && d < 0) {
      cardState = 'short';
      badgeHtml = `<span class="diff-badge short">ขาด ${Math.abs(d)}</span>`;
    } else if (d !== null && d > 0) {
      cardState = 'over';
      badgeHtml = `<span class="diff-badge over">เกิน +${d}</span>`;
    }

    const isSelected = selectedIds.has(String(p.id));
    const catName = (p.category || '').trim() || DEFAULT_CATEGORY;
    const countVal = (p.counted === null || p.counted === undefined || p.counted === '') ? '' : p.counted;

    return `
      <div class="product-card ${cardState} ${isSelected ? 'selected' : ''}"
           data-id="${escapeHtml(p.id)}"
           data-category="${escapeHtml(catName)}"
           onclick="handleCardClick(event, '${escapeHtml(p.id)}')">
        <div class="checkbox"></div>
        ${isAdmin ? `<button class="delete-btn" onclick="event.stopPropagation(); deleteProduct('${escapeHtml(p.id)}')" title="ลบสินค้า">✕</button>` : ''}
        <div class="product-header">
          <div class="product-title">
            <div class="product-name">${escapeHtml(p.name)}</div>
            ${p.sku ? `<span class="product-sku">${escapeHtml(p.sku)}</span>` : ''}
          </div>
          ${badgeHtml}
        </div>
        <div class="product-info">
          <div class="info-item">
            <span class="info-label">ในระบบ:</span>
            <span class="info-value">${p.system}</span>
          </div>
          <div class="info-item">
            <span class="info-label">นับได้:</span>
            <span class="info-value">${countVal === '' ? '-' : countVal}</span>
          </div>
        </div>
        <div class="product-footer" onclick="event.stopPropagation()">
          <div class="count-controls">
            <button type="button" onclick="adjustCount('${escapeHtml(p.id)}', -1)">−</button>
            <input type="number" inputmode="numeric" placeholder="-"
                   value="${countVal}"
                   onchange="setCount('${escapeHtml(p.id)}', this.value)">
            <button type="button" onclick="adjustCount('${escapeHtml(p.id)}', 1)">+</button>
          </div>
          ${isAdmin ? `<button type="button" class="edit-btn" onclick="openEditModal('${escapeHtml(p.id)}')">แก้ไข</button>` : ''}
        </div>
      </div>
    `;
  }).join('');
}

function handleCardClick(event, id) {
  if (!selectMode) return;
  toggleSelectCard(String(id));
}

function adjustCount(id, delta) {
  const p = products.find(item => String(item.id) === String(id));
  if (!p) return;
  const oldVal = p.counted;
  const base = (p.counted === null || p.counted === undefined || p.counted === '') ? p.system : Number(p.counted);
  const next = Math.max(0, base + delta);
  p.counted = next;
  scheduleCountLog(p, oldVal, next);
  saveDataToCache();
  render();
}

function setCount(id, rawValue) {
  const p = products.find(item => String(item.id) === String(id));
  if (!p) return;
  const oldVal = p.counted;
  if (rawValue === '' || rawValue === null || rawValue === undefined) {
    p.counted = null;
  } else {
    const parsed = parseInt(rawValue, 10);
    p.counted = isNaN(parsed) ? null : Math.max(0, parsed);
  }
  scheduleCountLog(p, oldVal, p.counted);
  saveDataToCache();
  render();
}

// ============ History & Activity Logs ============
function scheduleCountLog(product, oldVal, newVal) {
  if (oldVal === newVal) return;
  const key = String(product.id);
  if (logTimers[key]) {
    clearTimeout(logTimers[key].timer);
  } else {
    logTimers[key] = { initialOldVal: oldVal };
  }

  logTimers[key].timer = setTimeout(() => {
    const startVal = logTimers[key].initialOldVal;
    delete logTimers[key];
    if (startVal === product.counted) return;

    const diff = getDiff(product);
    countLogs.unshift({
      id: genId(),
      timestamp: new Date().toISOString(),
      productId: product.id,
      name: product.name,
      sku: product.sku || '',
      category: (product.category || '').trim() || DEFAULT_CATEGORY,
      system: product.system,
      oldCounted: (startVal === null || startVal === undefined || startVal === '') ? null : Number(startVal),
      newCounted: (product.counted === null || product.counted === undefined || product.counted === '') ? null : Number(product.counted),
      diff: diff,
      status: getStatusText(diff)
    });

    if (countLogs.length > 500) countLogs = countLogs.slice(0, 500);
    saveHistoryToCache();
    renderHistoryPage();
  }, 600);
}

function openSaveHistoryModal() {
  if (!products || products.length === 0) {
    showToast('⚠️ ยังไม่มีรายการสินค้าให้บันทึกประวัติ');
    return;
  }
  const counted = products.filter(p => p.counted !== null && p.counted !== undefined && p.counted !== '').length;
  const summaryEl = document.getElementById('historyModalSummary');
  if (summaryEl) {
    summaryEl.textContent = `สินค้าทั้งหมด ${products.length} รายการ • นับแล้ว ${counted} รายการ`;
  }
  const input = document.getElementById('historyNoteInput');
  if (input) {
    const now = new Date();
    input.value = `เช็คสต๊อก ${now.toLocaleDateString('th-TH')} (${now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })})`;
  }
  openModal('saveHistoryModal');
}

function confirmSaveHistorySession() {
  const input = document.getElementById('historyNoteInput');
  const note = (input && input.value.trim()) ? input.value.trim() : `เช็คสต๊อก ${new Date().toLocaleDateString('th-TH')}`;
  createHistorySession(note, false);
  closeModal('saveHistoryModal');
  showToast('📌 บันทึกรอบการนับเข้าประวัติเรียบร้อย');
}

function createHistorySession(title, silent = false) {
  if (!products || products.length === 0) return null;

  let counted = 0, ok = 0, short = 0, over = 0;
  const itemsSnapshot = products.map(p => {
    const d = getDiff(p);
    if (d !== null) {
      counted++;
      if (d === 0) ok++;
      else if (d < 0) short++;
      else if (d > 0) over++;
    }
    return {
      id: p.id,
      name: p.name,
      sku: p.sku || '',
      category: (p.category || '').trim() || DEFAULT_CATEGORY,
      system: p.system,
      counted: (p.counted === null || p.counted === undefined || p.counted === '') ? null : Number(p.counted),
      diff: d,
      status: getStatusText(d)
    };
  });

  const session = {
    id: genId(),
    title: title || 'รอบการนับสต๊อก',
    createdAt: new Date().toISOString(),
    fileName: fileName || '',
    total: products.length,
    counted,
    ok,
    short,
    over,
    items: itemsSnapshot
  };

  stockHistory.unshift(session);
  if (stockHistory.length > 100) stockHistory = stockHistory.slice(0, 100);
  saveHistoryToCache();
  renderHistoryPage();
  return session;
}

function switchHistorySubTab(subtab) {
  historySubTab = subtab;
  document.querySelectorAll('.hist-subtab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.subtab === subtab);
  });
  const sWrap = document.getElementById('histSessionsWrap');
  const lWrap = document.getElementById('histLogsWrap');
  if (sWrap) sWrap.style.display = subtab === 'sessions' ? '' : 'none';
  if (lWrap) lWrap.style.display = subtab === 'logs' ? '' : 'none';
  renderHistoryPage();
}

function renderHistoryPage() {
  const sStat = document.getElementById('histStatSessions');
  const lStat = document.getElementById('histStatLogs');
  const lastStat = document.getElementById('histStatLast');
  const sBadge = document.getElementById('histBadgeSessions');
  const lBadge = document.getElementById('histBadgeLogs');

  if (sStat) sStat.textContent = stockHistory.length;
  if (lStat) lStat.textContent = countLogs.length;
  if (sBadge) sBadge.textContent = stockHistory.length;
  if (lBadge) lBadge.textContent = countLogs.length;

  const latestIso = stockHistory[0]?.createdAt || countLogs[0]?.timestamp || null;
  if (lastStat) lastStat.textContent = latestIso ? formatDateTimeTH(latestIso) : 'ยังไม่มีประวัติ';

  const q = historySearchTerm.trim().toLowerCase();
  const dateFilter = historyDateFilter;

  // Render Sessions
  const sWrap = document.getElementById('histSessionsWrap');
  if (sWrap) {
    const filteredSessions = stockHistory.filter(s => {
      if (dateFilter && (!s.createdAt || !s.createdAt.startsWith(dateFilter))) return false;
      if (q) {
        const matchTitle = (s.title || '').toLowerCase().includes(q);
        const matchItems = Array.isArray(s.items) && s.items.some(it =>
          (it.name || '').toLowerCase().includes(q) || (it.sku || '').toLowerCase().includes(q)
        );
        if (!matchTitle && !matchItems) return false;
      }
      return true;
    });

    if (filteredSessions.length === 0) {
      sWrap.innerHTML = `
        <div class="rpt-empty" style="margin:0;">
          ยังไม่มีประวัติรอบการนับ กดปุ่ม "📌 บันทึกรอบการนับ" หรือกดบันทึกไฟล์ Excel เพื่อเก็บประวัติ
        </div>
      `;
    } else {
      sWrap.innerHTML = `<div class="history-card-list">` + filteredSessions.map(s => `
        <div class="history-card">
          <div class="hc-main">
            <div class="hc-title-row">
              <h4 class="hc-title">${escapeHtml(s.title)}</h4>
              <span class="hc-time">🕒 ${formatDateTimeTH(s.createdAt)}</span>
            </div>
            <div class="hc-pills">
              <span class="hc-pill">ทั้งหมด ${s.total}</span>
              <span class="hc-pill info">นับแล้ว ${s.counted}</span>
              <span class="hc-pill ok">ครบ ${s.ok}</span>
              <span class="hc-pill short">ขาด ${s.short}</span>
              <span class="hc-pill over">เกิน ${s.over}</span>
            </div>
          </div>
          <div class="hc-actions">
            <button class="btn btn-secondary" onclick="openHistoryDetail('${escapeHtml(s.id)}')">🔍 ดูรายละเอียด</button>
            <button class="btn btn-success" onclick="exportSingleSessionToExcel('${escapeHtml(s.id)}')">📥 Excel</button>
            ${isAdmin ? `<button class="btn btn-danger" onclick="deleteHistorySession('${escapeHtml(s.id)}')">🗑️</button>` : ''}
          </div>
        </div>
      `).join('') + `</div>`;
    }
  }

  // Render Activity Logs
  const logBody = document.getElementById('histLogBody');
  const logEmpty = document.getElementById('histLogEmpty');
  const logTableWrap = document.getElementById('histLogTableWrap');
  if (logBody && logEmpty && logTableWrap) {
    const filteredLogs = countLogs.filter(l => {
      if (dateFilter && (!l.timestamp || !l.timestamp.startsWith(dateFilter))) return false;
      if (q) {
        const matchName = (l.name || '').toLowerCase().includes(q);
        const matchSku = (l.sku || '').toLowerCase().includes(q);
        const matchCat = (l.category || '').toLowerCase().includes(q);
        if (!matchName && !matchSku && !matchCat) return false;
      }
      return true;
    });

    if (filteredLogs.length === 0) {
      logTableWrap.style.display = 'none';
      logEmpty.style.display = '';
      logBody.innerHTML = '';
    } else {
      logTableWrap.style.display = '';
      logEmpty.style.display = 'none';
      logBody.innerHTML = filteredLogs.map(l => {
        const d = l.diff;
        let stClass = 'plain';
        if (d === 0) stClass = 'ok';
        else if (d !== null && d < 0) stClass = 'short';
        else if (d !== null && d > 0) stClass = 'over';
        const diffTxt = (d === null || d === undefined) ? '-' : (d > 0 ? `+${d}` : `${d}`);
        return `
          <tr>
            <td>${formatDateTimeTH(l.timestamp)}</td>
            <td>${escapeHtml(l.name)}</td>
            <td>${escapeHtml(l.category || DEFAULT_CATEGORY)}</td>
            <td>${escapeHtml(l.sku || '-')}</td>
            <td class="num">${l.system}</td>
            <td class="num">${l.oldCounted === null || l.oldCounted === undefined ? '-' : l.oldCounted}</td>
            <td class="num">${l.newCounted === null || l.newCounted === undefined ? '-' : l.newCounted}</td>
            <td class="num">${diffTxt}</td>
            <td><span class="st ${stClass}">${escapeHtml(l.status || getStatusText(d))}</span></td>
          </tr>
        `;
      }).join('');
    }
  }
}

function openHistoryDetail(sessionId) {
  const s = stockHistory.find(item => String(item.id) === String(sessionId));
  if (!s) return;
  viewingSessionId = s.id;

  const setTxt = (id, v) => {
    const el = document.getElementById(id);
    if (el) el.textContent = v;
  };
  setTxt('hdTitle', s.title);
  setTxt('hdMeta', `บันทึกเมื่อ ${formatDateTimeTH(s.createdAt)}`);
  setTxt('hdTotal', s.total);
  setTxt('hdCounted', s.counted);
  setTxt('hdOk', s.ok);
  setTxt('hdShort', s.short);
  setTxt('hdOver', s.over);

  const restoreBtn = document.getElementById('hdRestoreBtn');
  if (restoreBtn) restoreBtn.style.display = isAdmin ? '' : 'none';

  const body = document.getElementById('hdBody');
  if (body) {
    body.innerHTML = (s.items || []).map((it, idx) => {
      const d = it.diff;
      let stClass = 'plain';
      if (d === 0) stClass = 'ok';
      else if (d !== null && d < 0) stClass = 'short';
      else if (d !== null && d > 0) stClass = 'over';
      const diffTxt = (d === null || d === undefined) ? '-' : (d > 0 ? `+${d}` : `${d}`);
      return `
        <tr>
          <td>${idx + 1}</td>
          <td>${escapeHtml(it.name)}</td>
          <td>${escapeHtml(it.category || DEFAULT_CATEGORY)}</td>
          <td>${escapeHtml(it.sku || '-')}</td>
          <td class="num">${it.system}</td>
          <td class="num">${it.counted === null || it.counted === undefined ? '-' : it.counted}</td>
          <td class="num">${diffTxt}</td>
          <td><span class="st ${stClass}">${escapeHtml(it.status || getStatusText(d))}</span></td>
        </tr>
      `;
    }).join('');
  }

  openModal('historyDetailModal');
}

function restoreHistorySession(sessionId) {
  const s = stockHistory.find(item => String(item.id) === String(sessionId));
  if (!s || !Array.isArray(s.items)) return;
  if (!confirm(`ต้องการนำข้อมูลจากรอบ "${s.title}" กลับมาใช้ในหน้าเช็คสต็อกหรือไม่?`)) return;

  products = s.items.map(it => ({
    id: it.id || genId(),
    name: it.name,
    sku: it.sku || '',
    category: it.category || '',
    system: Number(it.system) || 0,
    counted: (it.counted === null || it.counted === undefined || it.counted === '') ? null : Number(it.counted)
  }));

  saveDataToCache();
  render();
  closeModal('historyDetailModal');
  showToast('🔄 นำข้อมูลรอบการนับกลับมาใช้แล้ว');
}

function deleteHistorySession(sessionId) {
  if (!isAdmin) return;
  if (!confirm('ต้องการลบประวัติรอบการนับนี้หรือไม่?')) return;
  stockHistory = stockHistory.filter(s => String(s.id) !== String(sessionId));
  saveHistoryToCache();
  renderHistoryPage();
  showToast('🗑️ ลบประวัติรอบการนับแล้ว');
}

function clearAllHistory() {
  if (!isAdmin) return;
  if (!confirm('ต้องการล้างประวัติการนับทั้งหมดหรือไม่?')) return;
  stockHistory = [];
  countLogs = [];
  saveHistoryToCache();
  renderHistoryPage();
  showToast('🗑️ ล้างประวัติทั้งหมดเรียบร้อย');
}

function exportSingleSessionToExcel(sessionId) {
  const s = stockHistory.find(item => String(item.id) === String(sessionId));
  if (!s || !Array.isArray(s.items)) {
    showToast('⚠️ ไม่พบข้อมูลรอบการนับ');
    return;
  }
  const c = CONFIG.EXCEL_COLUMNS;
  const rows = s.items.map(it => ({
    [c.ID]: it.id,
    [c.NAME]: it.name,
    [c.SKU]: it.sku || '',
    [CAT_COL]: it.category || '',
    [c.SYSTEM]: it.system,
    [c.COUNTED]: (it.counted === null || it.counted === undefined) ? '' : it.counted,
    [c.DIFF]: (it.diff === null || it.diff === undefined) ? '' : it.diff,
    [c.STATUS]: it.status || getStatusText(it.diff),
    [c.UPDATED_AT]: formatDateTimeTH(s.createdAt)
  }));

  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'History');
  const safeTitle = (s.title || 'session').replace(/[\\/:*?"<>|]/g, '_');
  XLSX.writeFile(wb, `stockflow_${safeTitle}.xlsx`);
  showToast('📥 ดาวน์โหลดไฟล์ Excel เรียบร้อย');
}

function exportAllHistoryToExcel() {
  if (stockHistory.length === 0 && countLogs.length === 0) {
    showToast('⚠️ ยังไม่มีประวัติการนับให้ส่งออก');
    return;
  }

  const wb = XLSX.utils.book_new();

  // Sheet 1: สรุปรอบการนับ
  const sessionSummaryRows = stockHistory.map(s => ({
    'ชื่อรอบการนับ': s.title,
    'วัน-เวลา': formatDateTimeTH(s.createdAt),
    'สินค้าทั้งหมด': s.total,
    'นับแล้ว': s.counted,
    'ครบ': s.ok,
    'ขาด': s.short,
    'เกิน': s.over
  }));
  const wsSessions = XLSX.utils.json_to_sheet(sessionSummaryRows.length ? sessionSummaryRows : [{ 'สถานะ': 'ไม่มีข้อมูลรอบการนับ' }]);
  XLSX.utils.book_append_sheet(wb, wsSessions, 'สรุปรอบการนับ');

  // Sheet 2: ประวัติการนับรายชิ้น
  const logRows = countLogs.map(l => ({
    'วัน-เวลา': formatDateTimeTH(l.timestamp),
    'สินค้า': l.name,
    'หมวดหมู่': l.category || DEFAULT_CATEGORY,
    'SKU': l.sku || '',
    'ในระบบ': l.system,
    'ค่าเดิม': l.oldCounted === null || l.oldCounted === undefined ? '' : l.oldCounted,
    'นับได้ใหม่': l.newCounted === null || l.newCounted === undefined ? '' : l.newCounted,
    'ผลต่าง': l.diff === null || l.diff === undefined ? '' : l.diff,
    'สถานะ': l.status || getStatusText(l.diff)
  }));
  const wsLogs = XLSX.utils.json_to_sheet(logRows.length ? logRows : [{ 'สถานะ': 'ไม่มีประวัติรายชิ้น' }]);
  XLSX.utils.book_append_sheet(wb, wsLogs, 'ประวัติการนับรายชิ้น');

  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(wb, `stockflow_history_${stamp}.xlsx`);
  showToast('📥 ส่งออกประวัติทั้งหมดเป็น Excel เรียบร้อย');
}

// ============ Admin & PIN ============
function handleAdminClick() {
  if (isAdmin) {
    isAdmin = false;
    if (selectMode) toggleSelectMode();
    updateAdminUI();
    render();
    renderHistoryPage();
    showToast('🔒 ออกจากโหมดแอดมินแล้ว');
  } else {
    pinBuffer = '';
    updatePinDots();
    const errEl = document.getElementById('pinError');
    if (errEl) errEl.textContent = '';
    openModal('pinModal');
  }
}

function updateAdminUI() {
  const badge = document.getElementById('adminBadge');
  const icon = document.getElementById('adminIcon');
  const text = document.getElementById('adminText');
  const banner = document.getElementById('readonlyBanner');
  const fab = document.getElementById('fabAdd');

  if (badge) badge.classList.toggle('unlocked', isAdmin);
  if (icon) icon.textContent = isAdmin ? '🔐' : '🔓';
  if (text) text.textContent = isAdmin ? 'โหมดแอดมิน' : 'โหมดทั่วไป';
  if (banner) banner.style.display = isAdmin ? 'none' : 'block';
  if (fab) fab.classList.toggle('hidden', !isAdmin);

  document.querySelectorAll('.admin-only').forEach(btn => {
    btn.disabled = !isAdmin;
  });
}

function pinInput(digit) {
  if (pinBuffer.length >= (CONFIG.PIN_LENGTH || 4)) return;
  pinBuffer += String(digit);
  updatePinDots();

  if (pinBuffer.length === (CONFIG.PIN_LENGTH || 4)) {
    if (pinBuffer === String(adminPin)) {
      isAdmin = true;
      closeModal('pinModal');
      updateAdminUI();
      render();
      renderHistoryPage();
      showToast('✅ เข้าสู่โหมดแอดมินแล้ว');
    } else {
      const errEl = document.getElementById('pinError');
      if (errEl) errEl.textContent = 'รหัส PIN ไม่ถูกต้อง';
      pinBuffer = '';
      setTimeout(updatePinDots, 250);
    }
  }
}

function pinBackspace() {
  pinBuffer = pinBuffer.slice(0, -1);
  updatePinDots();
}

function updatePinDots() {
  const dots = document.querySelectorAll('#pinDisplay .pin-dot');
  dots.forEach((dot, idx) => {
    dot.classList.toggle('filled', idx < pinBuffer.length);
  });
}

function openChangePinModal() {
  if (!isAdmin) return;
  document.getElementById('oldPin').value = '';
  document.getElementById('newPin').value = '';
  document.getElementById('confirmPin').value = '';
  openModal('changePinModal');
}

function changePin() {
  const oldP = document.getElementById('oldPin').value.trim();
  const newP = document.getElementById('newPin').value.trim();
  const confP = document.getElementById('confirmPin').value.trim();

  if (oldP !== String(adminPin)) {
    alert('รหัส PIN เดิมไม่ถูกต้อง');
    return;
  }
  if (!/^\d{4}$/.test(newP)) {
    alert('กรุณาตั้งรหัส PIN ใหม่เป็นตัวเลข 4 หลัก');
    return;
  }
  if (newP !== confP) {
    alert('รหัส PIN ใหม่และการยืนยันไม่ตรงกัน');
    return;
  }

  saveAdminPin(newP);
  closeModal('changePinModal');
  showToast('🔑 เปลี่ยนรหัส PIN เรียบร้อย');
}

// ============ Product CRUD & Bulk Actions ============
function openAddModal() {
  if (!isAdmin) return;
  editingId = null;
  document.getElementById('modalTitle').textContent = 'เพิ่มสินค้า';
  document.getElementById('inputName').value = '';
  document.getElementById('inputSku').value = '';
  document.getElementById('inputCategory').value = '';
  document.getElementById('inputSystem').value = '';
  openModal('productModal');
}

function openEditModal(id) {
  if (!isAdmin) return;
  const p = products.find(item => String(item.id) === String(id));
  if (!p) return;
  editingId = p.id;
  document.getElementById('modalTitle').textContent = 'แก้ไขสินค้า';
  document.getElementById('inputName').value = p.name || '';
  document.getElementById('inputSku').value = p.sku || '';
  document.getElementById('inputCategory').value = p.category || '';
  document.getElementById('inputSystem').value = p.system ?? 0;
  openModal('productModal');
}

function saveProduct() {
  if (!isAdmin) return;
  const name = document.getElementById('inputName').value.trim();
  const sku = document.getElementById('inputSku').value.trim();
  const category = document.getElementById('inputCategory').value.trim();
  const system = parseInt(document.getElementById('inputSystem').value, 10) || 0;

  if (!name) {
    alert('กรุณากรอกชื่อสินค้า');
    return;
  }

  if (editingId) {
    const p = products.find(item => String(item.id) === String(editingId));
    if (p) {
      p

/* ============================================
   StockFlow — Application Logic
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

// Excel column name for category (falls back to Thai header 'หมวดหมู่')
const CAT_COL = (CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.CATEGORY) || 'หมวดหมู่';
const DEFAULT_CATEGORY = 'ไม่ระบุ';

// ============ Storage ============
function loadData() {
  try { products = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '[]'); } catch (e) { products = []; }
  try { stockHistory = JSON.parse(localStorage.getItem(CONFIG.HISTORY_KEY) || '[]'); } catch (e) { stockHistory = []; }
  try { countLogs = JSON.parse(localStorage.getItem(CONFIG.LOG_KEY) || '[]'); } catch (e) { countLogs = []; }
  try { adminPin = localStorage.getItem(CONFIG.ADMIN_KEY) || CONFIG.DEFAULT_PIN; } catch (e) { adminPin = CONFIG.DEFAULT_PIN; }
  restoreFileHandle().then(() => {
    updateAdminUI();
    render();
    renderHistoryPage();
  });
}

function saveDataToCache() {
  localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
}

function saveHistoryToCache() {
  localStorage.setItem(CONFIG.HISTORY_KEY, JSON.stringify(stockHistory));
  localStorage.setItem(CONFIG.LOG_KEY, JSON.stringify(countLogs));
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
    await saveFileHandleToDB(handle);
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
  const fileNameEl = document.getElementById('fileName');
  const fileSub = document.getElementById('fileSub');
  const fileIcon = document.getElementById('fileIcon');
  const openBtn = document.getElementById('openFileBtn');
  const saveBar = document.getElementById('saveBar');

  if (fileHandle) {
    fileNameEl.textContent = fileName;
    fileSub.textContent = 'อัปโหลดสำเร็จ • พร้อมใช้งาน';
    fileSub.classList.add('connected');
    fileIcon.classList.add('connected');
    fileIcon.textContent = '✓';
    openBtn.textContent = '🔄 เปลี่ยนไฟล์';
    saveBar.style.display = 'flex';
    updateSaveBar();
  } else {
    fileNameEl.textContent = 'ยังไม่ได้เชื่อมต่อไฟล์';
    fileSub.textContent = 'กรุณาเลือกไฟล์ Excel เพื่อเริ่มต้น';
    fileSub.classList.remove('connected');
    fileIcon.classList.remove('connected');
    fileIcon.textContent = '📄';
    openBtn.textContent = '📂 เปิดไฟล์';
    saveBar.style.display = 'none';
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

// ============ IndexedDB ============
function openDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(CONFIG.DB_NAME, 1);
    req.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains('files')) db.createObjectStore('files');
    };
    req.onsuccess = (e) => resolve(e.target.result);
    req.onerror = (e) => reject(e.target.error);
  });
}

async function saveFileHandleToDB(handle) {
  try {
    const db = await openDB();
    const tx = db.transaction('files', 'readwrite');
    tx.objectStore('files').put(handle, 'excelFile');
  } catch (e) { console.error('Save handle error:', e); }
}

async function restoreFileHandle() {
  try {
    const db = await openDB();
    const tx = db.transaction('files', 'readonly');
    const req = tx.objectStore('files').get('excelFile');
    req.onsuccess = async () => {
      if (req.result) {
        fileHandle = req.result;
        fileName = fileHandle.name;
        const perm = await fileHandle.queryPermission({ mode: 'readwrite' });
        fileHasPermission = perm === 'granted';
        updateFileBar();
      }
    };
  } catch (e) { console.error('Restore handle error:', e); }
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
  return d.toLocaleString('th-TH', {
    year: 'numeric', month: 'short', day: 'numeric',
    hour: '2-digit', minute: '2-digit', second: '2-digit'
  });
}

function getLocalDateKey(isoStr) {
  const d = new Date(isoStr);
  if (isNaN(d.getTime())) return '';
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yyyy}-${mm}-${dd}`;
}

// ============ History & Count Log System ============

// บันทึก Log รายชิ้นอัตโนมัติ (มี debounce กรณีกด + / - รัวๆ)
function scheduleCountLog(product, oldVal, newVal) {
  if (oldVal === newVal) return;
  const pid = product.id;

  if (!logTimers[pid]) {
    logTimers[pid] = { initialOldVal: oldVal, timer: null };
  }
  clearTimeout(logTimers[pid].timer);

  logTimers[pid].timer = setTimeout(() => {
    const startVal = logTimers[pid].initialOldVal;
    delete logTimers[pid];
    if (startVal === newVal) return;

    const diff = (newVal === null || newVal === '' || newVal === undefined) ? null : (newVal - product.system);
    const entry = {
      id: genId(),
      timestamp: new Date().toISOString(),
      productId: product.id,
      name: product.name,
      sku: product.sku || '',
      category: (product.category || '').trim() || DEFAULT_CATEGORY,
      system: product.system,
      oldCounted: startVal === undefined ? null : startVal,
      newCounted: newVal === undefined ? null : newVal,
      diff: diff,
      status: getStatusText(diff),
      role: isAdmin ? 'แอดมิน' : 'ผู้ใช้งานทั่วไป'
    };

    countLogs.unshift(entry);
    // เก็บสูงสุด 2,000 รายการล่าสุดเพื่อไม่ให้ localStorage เต็ม
    if (countLogs.length > 2000) countLogs.length = 2000;
    saveHistoryToCache();
    renderHistoryPage();
  }, 600);
}

// เปิด Modal เพื่อบันทึกรอบการนับสต๊อก
function openSaveHistoryModal() {
  if (products.length === 0) {
    showToast('⚠️ ยังไม่มีรายการสินค้าให้บันทึกประวัติ');
    return;
  }
  const countedCount = products.filter(p => p.counted !== null && p.counted !== '' && p.counted !== undefined).length;
  const nowStr = new Date().toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' });
  const input = document.getElementById('historyNoteInput');
  const summaryEl = document.getElementById('historyModalSummary');
  if (input) input.value = `รอบเช็คสต๊อก ${nowStr}`;
  if (summaryEl) {
    summaryEl.textContent = `นับแล้ว ${countedCount} จากทั้งหมด ${products.length} รายการ`;
  }
  document.getElementById('saveHistoryModal').classList.add('show');
  setTimeout(() => { if (input) input.select(); }, 150);
}

function confirmSaveHistorySession() {
  const input = document.getElementById('historyNoteInput');
  const note = (input && input.value.trim()) || `รอบเช็คสต๊อก ${new Date().toLocaleString('th-TH')}`;
  closeModal('saveHistoryModal');
  createHistorySession(note, false);
}

function createHistorySession(note, silent = false) {
  if (products.length === 0) return null;

  const now = new Date().toISOString();
  const items = products.map(p => {
    const d = getDiff(p);
    return {
      id: p.id,
      name: p.name,
      sku: p.sku || '',
      category: (p.category || '').trim() || DEFAULT_CATEGORY,
      system: p.system,
      counted: (p.counted === null || p.counted === undefined || p.counted === '') ? null : p.counted,
      diff: d,
      status: getStatusText(d)
    };
  });

  const total = items.length;
  const counted = items.filter(i => i.counted !== null).length;
  const ok = items.filter(i => i.diff === 0).length;
  const short = items.filter(i => i.diff !== null && i.diff < 0).length;
  const over = items.filter(i => i.diff !== null && i.diff > 0).length;

  const session = {
    id: genId(),
    timestamp: now,
    note: note || 'เช็คสต๊อก',
    fileName: fileName || 'ข้อมูลในระบบ',
    role: isAdmin ? 'แอดมิน' : 'ผู้ใช้งานทั่วไป',
    summary: { total, counted, ok, short, over },
    items: items
  };

  stockHistory.unshift(session);
  if (stockHistory.length > 200) stockHistory.length = 200;
  saveHistoryToCache();
  renderHistoryPage();

  if (!silent) {
    showToast('✅ บันทึกประวัติรอบการนับเรียบร้อย');
  }
  return session;
}

// ดูรายละเอียดรอบการนับย้อนหลัง
function viewHistorySession(sessionId) {
  const s = stockHistory.find(x => x.id === sessionId);
  if (!s) return;
  viewingSessionId = sessionId;

  document.getElementById('hdTitle').textContent = s.note;
  document.getElementById('hdMeta').textContent = `🕒 ${formatDateTimeTH(s.timestamp)} • 📄 ${s.fileName || '-'} • 👤 ${s.role || '-'}`;
  document.getElementById('hdTotal').textContent = s.summary.total;
  document.getElementById('hdCounted').textContent = s.summary.counted;
  document.getElementById('hdOk').textContent = s.summary.ok;
  document.getElementById('hdShort').textContent = s.summary.short;
  document.getElementById('hdOver').textContent = s.summary.over;

  const tbody = document.getElementById('hdBody');
  tbody.innerHTML = s.items.map((item, idx) => {
    let badgeClass = 'plain';
    if (item.diff === 0) badgeClass = 'ok';
    else if (item.diff !== null && item.diff < 0) badgeClass = 'short';
    else if (item.diff !== null && item.diff > 0) badgeClass = 'over';

    const diffStr = item.diff === null ? '-' : (item.diff > 0 ? `+${item.diff}` : `${item.diff}`);
    return `
      <tr>
        <td>${idx + 1}</td>
        <td><strong>${escapeHtml(item.name)}</strong></td>
        <td>${escapeHtml(item.category || DEFAULT_CATEGORY)}</td>
        <td>${escapeHtml(item.sku || '-')}</td>
        <td class="num">${item.system}</td>
        <td class="num">${item.counted === null ? '-' : item.counted}</td>
        <td class="num">${diffStr}</td>
        <td><span class="st ${badgeClass}">${escapeHtml(item.status)}</span></td>
      </tr>
    `;
  }).join('');

  document.getElementById('historyDetailModal').classList.add('show');
}

// โหลดข้อมูลจากประวัติรอบนั้นกลับมาที่หน้าเช็คสต็อก (เฉพาะแอดมิน)
function restoreHistorySession(sessionId) {
  requireAdmin(() => {
    const s = stockHistory.find(x => x.id === sessionId);
    if (!s) return;
    if (!confirm(`ต้องการนำข้อมูลจาก "${s.note}" (${formatDateTimeTH(s.timestamp)}) กลับมาใช้ในหน้าเช็คสต็อกปัจจุบันหรือไม่?`)) return;

    products = s.items.map(item => ({
      id: item.id || genId(),
      name: item.name,
      sku: item.sku === '-' ? '' : item.sku,
      category: item.category === DEFAULT_CATEGORY ? '' : item.category,
      system: item.system,
      counted: item.counted
    }));

    saveDataToCache();
    render();
    closeModal('historyDetailModal');
    showToast('🔄 นำข้อมูลจากประวัติกลับมาใช้แล้ว');
    if (typeof window.go === 'function') window.go('check');
  });
}

function deleteHistorySession(sessionId) {
  requireAdmin(() => {
    const s = stockHistory.find(x => x.id === sessionId);
    if (!s) return;
    pendingDeleteAction = () => {
      stockHistory = stockHistory.filter(x => x.id !== sessionId);
      saveHistoryToCache();
      renderHistoryPage();
      showToast('🗑️ ลบประวัติรอบการนับแล้ว');
    };
    showConfirmDelete('🗑️ ลบประวัติรอบการนับ', `ต้องการลบประวัติ "${s.note}" หรือไม่?`);
  });
}

function clearAllHistory() {
  requireAdmin(() => {
    if (stockHistory.length === 0 && countLogs.length === 0) {
      showToast('ไม่มีประวัติให้ลบ');
      return;
    }
    pendingDeleteAction = () => {
      stockHistory = [];
      countLogs = [];
      saveHistoryToCache();
      renderHistoryPage();
      showToast('🗑️ ล้างประวัติทั้งหมดแล้ว');
    };
    showConfirmDelete('🗑️ ล้างประวัติทั้งหมด', 'ต้องการลบประวัติรอบการนับและประวัติรายชิ้นทั้งหมดหรือไม่? (ไม่สามารถกู้คืนได้)');
  });
}

// สลับแท็บในหน้าประวัติ
function switchHistorySubTab(tabName) {
  historySubTab = tabName;
  document.querySelectorAll('.hist-subtab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.subtab === tabName);
  });
  renderHistoryPage();
}

// แสดงผลหน้าประวัติการนับ (#history)
function renderHistoryPage() {
  const statSessionsEl = document.getElementById('histStatSessions');
  const statLogsEl = document.getElementById('histStatLogs');
  const statLastEl = document.getElementById('histStatLast');
  const badgeSessions = document.getElementById('histBadgeSessions');
  const badgeLogs = document.getElementById('histBadgeLogs');

  if (statSessionsEl) statSessionsEl.textContent = stockHistory.length;
  if (statLogsEl) statLogsEl.textContent = countLogs.length;
  if (badgeSessions) badgeSessions.textContent = stockHistory.length;
  if (badgeLogs) badgeLogs.textContent = countLogs.length;

  if (statLastEl) {
    const latestTime = (stockHistory[0] && stockHistory[0].timestamp) || (countLogs[0] && countLogs[0].timestamp);
    statLastEl.textContent = latestTime ? formatDateTimeTH(latestTime) : 'ยังไม่มีประวัติ';
  }

  const sessionsWrap = document.getElementById('histSessionsWrap');
  const logsWrap = document.getElementById('histLogsWrap');
  if (!sessionsWrap || !logsWrap) return;

  const q = historySearchTerm.trim().toLowerCase();
  const dateFilter = historyDateFilter;

  if (historySubTab === 'sessions') {
    sessionsWrap.style.display = 'block';
    logsWrap.style.display = 'none';

    let list = stockHistory;
    if (dateFilter) {
      list = list.filter(s => getLocalDateKey(s.timestamp) === dateFilter);
    }
    if (q) {
      list = list.filter(s =>
        (s.note || '').toLowerCase().includes(q) ||
        (s.fileName || '').toLowerCase().includes(q) ||
        (s.items || []).some(i =>
          (i.name || '').toLowerCase().includes(q) ||
          (i.sku || '').toLowerCase().includes(q) ||
          (i.category || '').toLowerCase().includes(q)
        )
      );
    }

    if (list.length === 0) {
      sessionsWrap.innerHTML = `
        <div class="rpt-empty">
          <div style="font-size:32px;margin-bottom:8px;">🕒</div>
          <div style="font-weight:600;color:#0f1d36;margin-bottom:4px;">ยังไม่มีประวัติรอบการนับสต๊อก</div>
          <div>กดปุ่ม <strong>"📌 บันทึกรอบการนับปัจจุบัน"</strong> ด้านบน หรือกดบันทึกไฟล์ในหน้าเช็คสต็อกเพื่อเก็บประวัติ</div>
        </div>
      `;
      return;
    }

    sessionsWrap.innerHTML = `<div class="history-card-list">` + list.map(s => `
      <div class="history-card">
        <div class="hc-main">
          <div class="hc-title-row">
            <h4 class="hc-title">${escapeHtml(s.note)}</h4>
            <span class="hc-time">🕒 ${formatDateTimeTH(s.timestamp)}</span>
          </div>
          <div class="hc-sub">
            <span>📄 ไฟล์: ${escapeHtml(s.fileName || '-')}</span>
            <span>👤 โดย: ${escapeHtml(s.role || 'ผู้ใช้งาน')}</span>
          </div>
          <div class="hc-pills">
            <span class="hc-pill">ทั้งหมด <strong>${s.summary.total}</strong></span>
            <span class="hc-pill info">นับแล้ว <strong>${s.summary.counted}</strong></span>
            <span class="hc-pill ok">ครบ <strong>${s.summary.ok}</strong></span>
            <span class="hc-pill short">ขาด <strong>${s.summary.short}</strong></span>
            <span class="hc-pill over">เกิน <strong>${s.summary.over}</strong></span>
          </div>
        </div>
        <div class="hc-actions">
          <button class="btn btn-primary" onclick="viewHistorySession('${s.id}')">👁️ ดูย้อนหลัง</button>
          <button class="btn btn-secondary" onclick="exportSingleSessionToExcel('${s.id}')">📥 Excel</button>
          ${isAdmin ? `<button class="btn btn-danger" onclick="deleteHistorySession('${s.id}')" title="ลบ">🗑️</button>` : ''}
        </div>
      </div>
    `).join('') + `</div>`;
  } else {
    sessionsWrap.style.display = 'none';
    logsWrap.style.display = 'block';

    let logs = countLogs;
    if (dateFilter) {
      logs = logs.filter(l => getLocalDateKey(l.timestamp) === dateFilter);
    }
    if (q) {
      logs = logs.filter(l =>
        (l.name || '').toLowerCase().includes(q) ||
        (l.sku || '').toLowerCase().includes(q) ||
        (l.category || '').toLowerCase().includes(q) ||
        (l.status || '').toLowerCase().includes(q)
      );
    }

    const logBody = document.getElementById('histLogBody');
    const logEmpty = document.getElementById('histLogEmpty');
    const logTableWrap = document.getElementById('histLogTableWrap');

    if (logs.length === 0) {
      logEmpty.style.display = 'block';
      logTableWrap.style.display = 'none';
      logBody.innerHTML = '';
      return;
    }

    logEmpty.style.display = 'none';
    logTableWrap.style.display = 'block';

    logBody.innerHTML = logs.map(l => {
      let badgeClass = 'plain';
      if (l.diff === 0) badgeClass = 'ok';
      else if (l.diff !== null && l.diff < 0) badgeClass = 'short';
      else if (l.diff !== null && l.diff > 0) badgeClass = 'over';

      const oldTxt = (l.oldCounted === null || l.oldCounted === undefined || l.oldCounted === '') ? 'ยังไม่นับ' : l.oldCounted;
      const newTxt = (l.newCounted === null || l.newCounted === undefined || l.newCounted === '') ? 'ยังไม่นับ' : l.newCounted;
      const diffTxt = l.diff === null ? '-' : (l.diff > 0 ? `+${l.diff}` : `${l.diff}`);

      return `
        <tr>
          <td style="white-space:nowrap;">${formatDateTimeTH(l.timestamp)}</td>
          <td><strong>${escapeHtml(l.name)}</strong></td>
          <td>${escapeHtml(l.category || DEFAULT_CATEGORY)}</td>
          <td>${escapeHtml(l.sku || '-')}</td>
          <td class="num">${l.system}</td>
          <td class="num" style="color:#64748b;">${oldTxt}</td>
          <td class="num"><strong>${newTxt}</strong></td>
          <td class="num">${diffTxt}</td>
          <td><span class="st ${badgeClass}">${escapeHtml(l.status)}</span></td>
        </tr>
      `;
    }).join('');
  }
}

// ============ Excel Export Functions (.xlsx) ============

// เซฟประวัติทั้งหมดเป็นไฟล์ Excel (.xlsx) ครบทุกชีต
function exportAllHistoryToExcel() {
  if (stockHistory.length === 0 && countLogs.length === 0) {
    showToast('⚠️ ยังไม่มีข้อมูลประวัติสำหรับส่งออกเป็น Excel');
    return;
  }

  try {
    const wb = XLSX.utils.book_new();

    // 1. ชีตสรุปรอบการนับทั้งหมด
    const summaryData = stockHistory.map((s, idx) => ({
      'ลำดับ': idx + 1,
      'วันที่-เวลา': formatDateTimeTH(s.timestamp),
      'ชื่อรอบการนับ / หมายเหตุ': s.note,
      'ไฟล์อ้างอิง': s.fileName || '-',
      'ผู้บันทึก': s.role || '-',
      'สินค้าทั้งหมด': s.summary.total,
      'นับแล้ว': s.summary.counted,
      'ยังไม่ได้นับ': s.summary.total - s.summary.counted,
      'ครบ (ตรง)': s.summary.ok,
      'ขาด': s.summary.short,
      'เกิน': s.summary.over
    }));
    const wsSummary = XLSX.utils.json_to_sheet(
      summaryData.length ? summaryData : [{ 'ข้อความ': 'ยังไม่มีประวัติรอบการนับ' }]
    );
    XLSX.utils.book_append_sheet(wb, wsSummary, 'สรุปรอบการนับ');

    // 2. ชีตรายละเอียดสินค้าทุกรอบการนับรวมกัน
    const allSessionItems = [];
    stockHistory.forEach((s) => {
      (s.items || []).forEach((item) => {
        allSessionItems.push({
          'วันที่-เวลา': formatDateTimeTH(s.timestamp),
          'ชื่อรอบการนับ': s.note,
          'ID สินค้า': item.id,
          'ชื่อสินค้า': item.name,
          'หมวดหมู่': item.category || DEFAULT_CATEGORY,
          'SKU': item.sku || '',
          'ในระบบ': item.system,
          'นับได้': item.counted === null ? '' : item.counted,
          'ผลต่าง': item.diff === null ? '' : item.diff,
          'สถานะ': item.status
        });
      });
    });
    const wsDetails = XLSX.utils.json_to_sheet(
      allSessionItems.length ? allSessionItems : [{ 'ข้อความ': 'ยังไม่มีรายละเอียดรอบการนับ' }]
    );
    XLSX.utils.book_append_sheet(wb, wsDetails, 'รายละเอียดทุกรอบ');

    // 3. ชีตประวัติการนับรายชิ้นตามเวลาจริง (Activity Logs)
    const logSheetData = countLogs.map((l, idx) => ({
      'ลำดับ': idx + 1,
      'วัน-เวลาที่นับ': formatDateTimeTH(l.timestamp),
      'ชื่อสินค้า': l.name,
      'หมวดหมู่': l.category || DEFAULT_CATEGORY,
      'SKU': l.sku || '',
      'ในระบบ': l.system,
      'ค่าเดิม': (l.oldCounted === null || l.oldCounted === undefined || l.oldCounted === '') ? 'ยังไม่นับ' : l.oldCounted,
      'นับได้ใหม่': (l.newCounted === null || l.newCounted === undefined || l.newCounted === '') ? 'ยังไม่นับ' : l.newCounted,
      'ผลต่าง': l.diff === null ? '' : l.diff,
      'สถานะ': l.status,
      'ผู้ทำรายการ': l.role || '-'
    }));
    const wsLogs = XLSX.utils.json_to_sheet(
      logSheetData.length ? logSheetData : [{ 'ข้อความ': 'ยังไม่มีประวัติการนับรายชิ้น' }]
    );
    XLSX.utils.book_append_sheet(wb, wsLogs, 'Log_การนับรายชิ้น');

    const now = new Date();
    const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `StockFlow_History_${stamp}.xlsx`);
    showToast('📥 ดาวน์โหลดไฟล์ Excel ประวัติทั้งหมดเรียบร้อย');
  } catch (err) {
    console.error(err);
    alert('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: ' + err.message);
  }
}

// เซฟเฉพาะรอบการนับที่เลือกเป็นไฟล์ Excel (.xlsx)
function exportSingleSessionToExcel(sessionId) {
  const targetId = sessionId || viewingSessionId;
  const s = stockHistory.find(x => x.id === targetId);
  if (!s) {
    showToast('ไม่พบรอบการนับที่ต้องการส่งออก');
    return;
  }

  try {
    const wb = XLSX.utils.book_new();
    const rows = s.items.map((item, idx) => ({
      'ลำดับ': idx + 1,
      'ID': item.id,
      'ชื่อสินค้า': item.name,
      'หมวดหมู่': item.category || DEFAULT_CATEGORY,
      'SKU': item.sku || '',
      'ในระบบ': item.system,
      'นับได้': item.counted === null ? '' : item.counted,
      'ผลต่าง': item.diff === null ? '' : item.diff,
      'สถานะ': item.status,
      'รอบการนับ': s.note,
      'วันที่บันทึก': formatDateTimeTH(s.timestamp)
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, 'ประวัติการนับ');

    const d = new Date(s.timestamp);
    const stamp = isNaN(d.getTime())
      ? Date.now()
      : `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}_${String(d.getHours()).padStart(2, '0')}${String(d.getMinutes()).padStart(2, '0')}`;
    XLSX.writeFile(wb, `StockFlow_Session_${stamp}.xlsx`);
    showToast('📥 ดาวน์โหลดไฟล์ Excel รอบการนับเรียบร้อย');
  } catch (err) {
    console.error(err);
    alert('ไม่สามารถส่งออกไฟล์ Excel ได้: ' + err.message);
  }
}

// ส่งออกข้อมูลสต๊อกปัจจุบันเป็น Excel (.xlsx) จากหน้ารายงาน
function exportCurrentStockToExcel() {
  if (products.length === 0) {
    showToast('ยังไม่มีข้อมูลให้ส่งออก');
    return;
  }
  const c = CONFIG.EXCEL_COLUMNS;
  const exportData = products.map(p => {
    const d = getDiff(p);
    return {
      [c.ID]: p.id,
      [c.NAME]: p.name,
      [CAT_COL]: (p.category || '').trim() || DEFAULT_CATEGORY,
      [c.SKU]: p.sku || '',
      [c.SYSTEM]: p.system,
      [c.COUNTED]: (p.counted === null || p.counted === '') ? '' : p.counted,
      [c.DIFF]: d === null ? '' : d,
      [c.STATUS]: getStatusText(d),
      [c.UPDATED_AT]: new Date().toLocaleString('th-TH')
    };
  });
  const ws = XLSX.utils.json_to_sheet(exportData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, CONFIG.SHEET_NAME);
  const now = new Date();
  const stamp = `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
  XLSX.writeFile(wb, `StockFlow_Report_${stamp}.xlsx`);
  showToast('📥 ส่งออกรายงาน Excel เรียบร้อย');
}

// ============ Admin ============
function updateAdminUI() {
  const badge = document.getElementById('adminBadge');
  const text = document.getElementById('adminText');
  const icon = document.getElementById('adminIcon');
  const banner = document.getElementById('readonlyBanner');
  const fab = document.getElementById('fabAdd');
  const adminBtns = document.querySelectorAll('.admin-only');
  const restoreBtn = document.getElementById('hdRestoreBtn');

  if (isAdmin) {
    badge.classList.add('unlocked');
    text.textContent = 'แอดมิน';
    icon.textContent = '🔒';
    banner.style.display = 'none';
    fab.classList.remove('hidden');
    adminBtns.forEach(b => b.disabled = false);
    if (restoreBtn) restoreBtn.style.display = 'inline-flex';
  } else {
    badge.classList.remove('unlocked');
    text.textContent = 'โหมดทั่วไป';
    icon.textContent = '🔓';
    banner.style.display = 'block';
    fab.classList.add('hidden');
    adminBtns.forEach(b => b.disabled = true);
    if (restoreBtn) restoreBtn.style.display = 'none';
    if (selectMode) exitSelectMode();
  }
}

function handleAdminClick() {
  if (isAdmin) {
    if (confirm('ออกจากโหมดแอดมิน?')) {
      isAdmin = false;
      updateAdminUI();
      render();
      renderHistoryPage();
      showToast('ออกจากโหมดแอดมินแล้ว');
    }
  } else {
    openPinModal();
  }
}

function openPinModal(callback) {
  pinBuffer = '';
  pinCallback = callback || null;
  document.getElementById('pinTitle').textContent = callback ? '🔐 ยืนยันตัวตน' : '🔐 เข้าสู่โหมดแอดมิน';
  document.getElementById('pinError').textContent = '';
  renderPinDots();
  document.getElementById('pinModal').classList.add('show');
}

function renderPinDots() {
  document.querySelectorAll('#pinDisplay .pin-dot').forEach((d, i) => {
    d.classList.toggle('filled', i < pinBuffer.length);
  });
}

function pinInput(digit) {
  if (pinBuffer.length >= CONFIG.PIN_LENGTH) return;
  pinBuffer += digit;
  renderPinDots();
  document.getElementById('pinError').textContent = '';
  if (pinBuffer.length === CONFIG.PIN_LENGTH) setTimeout(checkPin, 150);
}

function pinBackspace() {
  pinBuffer = pinBuffer.slice(0, -1);
  renderPinDots();
  document.getElementById('pinError').textContent = '';
}

function checkPin() {
  if (pinBuffer === adminPin) {
    closeModal('pinModal');
    if (pinCallback) {
      const cb = pinCallback; pinCallback = null; cb();
    } else {
      isAdmin = true;
      updateAdminUI();
      render();
      renderHistoryPage();
      showToast('✅ เข้าสู่โหมดแอดมินแล้ว');
    }
    pinBuffer = '';
  } else {
    document.getElementById('pinError').textContent = '❌ PIN ไม่ถูกต้อง';
    pinBuffer = '';
    renderPinDots();
    if (navigator.vibrate) navigator.vibrate(200);
  }
}

function requireAdmin(action) {
  if (isAdmin) action();
  else openPinModal(action);
}

function saveAdminPin(p) {
  adminPin = p;
  localStorage.setItem(CONFIG.ADMIN_KEY, p);
}

// ============ Render ============
function renderStats() {
  const total = products.length;
  const counted = products.filter(p => p.counted !== null && p.counted !== '' && p.counted !== undefined).length;
  const diff = products.filter(p => {
    const d = getDiff(p);
    return d !== null && d !== 0;
  }).length;

  document.getElementById('statTotal').textContent = total;
  document.getElementById('statCounted').textContent = counted;
  document.getElementById('statDiff').textContent = diff;

  const pending = total - counted;
  document.getElementById('tabCountAll').textContent = total;
  document.getElementById('tabCountPending').textContent = pending;
  document.getElementById('tabCountDone').textContent = counted;
  document.getElementById('tabCountDiff').textContent = diff;
}

function renderList() {
  const container = document.getElementById('productList');
  let list = products;

  if (currentTab === 'pending') {
    list = list.filter(p => p.counted === null || p.counted === '' || p.counted === undefined);
  } else if (currentTab === 'done') {
    list = list.filter(p => p.counted !== null && p.counted !== '' && p.counted !== undefined);
  } else if (currentTab === 'diff') {
    list = list.filter(p => {
      const d = getDiff(p);
      return d !== null && d !== 0;
    });
  }

  if (searchTerm) {
    const s = searchTerm.toLowerCase();
    list = list.filter(p =>
      (p.name || '').toLowerCase().includes(s) ||
      (p.sku || '').toLowerCase().includes(s) ||
      (p.category || '').toLowerCase().includes(s)
    );
  }

  if (list.length === 0) {
    container.innerHTML = `
      <div class="empty-state" style="grid-column: 1/-1;">
        <div class="icon">📦</div>
        <div class="title">ไม่พบรายการสินค้า</div>
        <div class="desc">${isAdmin ? 'กดปุ่ม + เพื่อเพิ่มสินค้า หรือเปิดไฟล์ Excel' : 'กรุณาให้แอดมินเพิ่มสินค้า'}</div>
      </div>`;
    return;
  }

  container.innerHTML = list.map(p => {
    const diff = getDiff(p);
    let cardClass = 'plain';
    let diffHtml = '';
    if (diff !== null) {
      if (diff === 0) {
        cardClass = 'ok';
        diffHtml = `<span class="diff-badge ok">✓ ตรง</span>`;
      } else if (diff < 0) {
        cardClass = 'short';
        diffHtml = `<span class="diff-badge short">↓ ขาด ${Math.abs(diff)}</span>`;
      } else {
        cardClass = 'over';
        diffHtml = `<span class="diff-badge over">↑ เกิน ${diff}</span>`;
      }
    }
    const isSelected = selectedIds.has(p.id);
    return `
      <div class="product-card ${cardClass} ${isSelected ? 'selected' : ''}"
           data-id="${p.id}" data-category="${escapeHtml(p.category || '')}"
           onclick="handleCardClick(event, '${p.id}')">
        <div class="checkbox"></div>
        ${isAdmin ? `<button class="delete-btn" onclick="event.stopPropagation(); deleteSingle('${p.id}')" title="ลบ">✕</button>` : ''}
        <div class="product-header">
          <div class="product-title">
            <div class="product-name">${escapeHtml(p.name)}</div>
            ${p.sku ? `<span class="product-sku">${escapeHtml(p.sku)}</span>` : ''}
          </div>
          ${diffHtml}
        </div>
        <div class="product-info">
          <div class="info-item">
            <span class="info-label">ในระบบ</span>
            <span class="info-value">${p.system}</span>
          </div>
          ${p.counted !== null && p.counted !== undefined && p.counted !== '' ? `
          <div class="info-item">
            <span class="info-label">นับได้</span>
            <span class="info-value">${p.counted}</span>
          </div>` : ''}
        </div>
        <div class="product-footer">
          <div class="count-controls">
            <button onclick="event.stopPropagation(); changeCount('${p.id}', -1)" title="ลด">−</button>
            <input type="number"
              value="${p.counted === null || p.counted === undefined || p.counted === '' ? '' : p.counted}"
              placeholder="นับ"
              onclick="event.stopPropagation()"
              onchange="setCount('${p.id}', this.value)"
              onfocus="this.select()">
            <button onclick="event.stopPropagation(); changeCount('${p.id}', 1)" title="เพิ่ม">+</button>
          </div>
          ${isAdmin ? `<button class="edit-btn" onclick="event.stopPropagation(); openEditModal('${p.id}')">แก้ไข</button>` : ''}
        </div>
      </div>
    `;
  }).join('');

  updateSelectionBar();
}

function render() { renderStats(); renderList(); updateSaveBar(); }

// ============ Selection Mode ============
function toggleSelectMode() {
  if (selectMode) exitSelectMode();
  else enterSelectMode();
}

function enterSelectMode() {
  if (!isAdmin) { requireAdmin(() => enterSelectMode()); return; }
  selectMode = true;
  selectedIds.clear();
  document.body.classList.add('select-mode');
  document.getElementById('selectionBar').classList.add('show');
  document.getElementById('selectModeBtn').textContent = '✖ ยกเลิก';
  document.getElementById('fabAdd').classList.add('hidden');
  document.getElementById('saveBar').style.display = 'none';
  renderList();
}

function exitSelectMode() {
  selectMode = false;
  selectedIds.clear();
  document.body.classList.remove('select-mode');
  document.getElementById('selectionBar').classList.remove('show');
  document.getElementById('selectModeBtn').textContent = '☑️ เลือกลบ';
  if (isAdmin) document.getElementById('fabAdd').classList.remove('hidden');
  if (fileHandle) document.getElementById('saveBar').style.display = 'flex';
  renderList();
}

function handleCardClick(event, id) {
  if (!selectMode) return;
  event.preventDefault();
  if (selectedIds.has(id)) selectedIds.delete(id);
  else selectedIds.add(id);
  renderList();
}

function updateSelectionBar() {
  const n = selectedIds.size;
  document.getElementById('selectedCount').textContent = n;
  document.getElementById('deleteSelectedBtn').disabled = n === 0;
}

function deleteSelected() {
  if (selectedIds.size === 0) return;
  const n = selectedIds.size;
  pendingDeleteAction = () => {
    products = products.filter(p => !selectedIds.has(p.id));
    saveDataToCache();
    exitSelectMode();
    render();
    showToast(`🗑️ ลบ ${n} รายการแล้ว`);
  };
  showConfirmDelete('🗑️ ลบสินค้าที่เลือก', `ต้องการลบสินค้า ${n} รายการที่เลือกหรือไม่? (ไม่สามารถกู้คืนได้)`);
}

function deleteSingle(id) {
  const p = products.find(x => x.id === id);
  if (!p) return;
  pendingDeleteAction = () => {
    products = products.filter(x => x.id !== id);
    saveDataToCache();
    render();
    showToast('🗑️ ลบสินค้าแล้ว');
  };
  showConfirmDelete('🗑️ ลบสินค้า', `ต้องการลบ "${p.name}" หรือไม่?`);
}

function showConfirmDelete(title, text) {
  document.getElementById('confirmDeleteTitle').textContent = title;
  document.getElementById('confirmDeleteText').textContent = text;
  document.getElementById('confirmDeleteModal').classList.add('show');
}

function confirmDeleteExecute() {
  closeModal('confirmDeleteModal');
  if (pendingDeleteAction) {
    const fn = pendingDeleteAction;
    pendingDeleteAction = null;
    fn();
  }
}

// ============ Count Actions ============
function changeCount(id, delta) {
  const p = products.find(x => x.id === id);
  if (!p) return;
  const oldVal = (p.counted === null || p.counted === undefined || p.counted === '') ? null : parseInt(p.counted);
  let cur = oldVal === null ? 0 : oldVal;
  cur = Math.max(0, cur + delta);
  p.counted = cur;
  scheduleCountLog(p, oldVal, cur);
  saveDataToCache();
  render();
}

function setCount(id, val) {
  const p = products.find(x => x.id === id);
  if (!p) return;
  const oldVal = (p.counted === null || p.counted === undefined || p.counted === '') ? null : parseInt(p.counted);
  let newVal = null;
  if (val !== '' && val !== null) {
    const n = parseInt(val);
    newVal = isNaN(n) ? null : Math.max(0, n);
  }
  p.counted = newVal;
  scheduleCountLog(p, oldVal, newVal);
  saveDataToCache();
  render();
}

// ============ Admin Actions ============
function fillCategoryList() {
  const list = document.getElementById('catList');
  if (!list) return;
  const cats = [...new Set(products.map(p => (p.category || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
  list.innerHTML = cats.map(c => `<option value="${escapeHtml(c)}"></option>`).join('');
}

function openAddModal() {
  requireAdmin(() => {
    editingId = null;
    document.getElementById('modalTitle').textContent = 'เพิ่มสินค้า';
    document.getElementById('inputName').value = '';
    document.getElementById('inputSku').value = '';
    document.getElementById('inputCategory').value = '';
    document.getElementById('inputSystem').value = '';
    fillCategoryList();
    document.getElementById('productModal').classList.add('show');
    setTimeout(() => document.getElementById('inputName').focus(), 200);
  });
}

function openEditModal(id) {
  requireAdmin(() => {
    const p = products.find(x => x.id === id);
    if (!p) return;
    editingId = id;
    document.getElementById('modalTitle').textContent = 'แก้ไขสินค้า';
    document.getElementById('inputName').value = p.name;
    document.getElementById('inputSku').value = p.sku || '';
    document.getElementById('inputCategory').value = p.category || '';
    document.getElementById('inputSystem').value = p.system;
    fillCategoryList();
    document.getElementById('productModal').classList.add('show');
  });
}

function saveProduct() {
  const name = document.getElementById('inputName').value.trim();
  const sku = document.getElementById('inputSku').value.trim();
  const category = document.getElementById('inputCategory').value.trim();
  const system = parseInt(document.getElementById('inputSystem').value) || 0;
  if (!name) { alert('กรุณากรอกชื่อสินค้า'); return; }

  if (editingId) {
    const p = products.find(x => x.id === editingId);
    if (p) { p.name = name; p.sku = sku; p.category = category; p.system = system; }
    showToast('✅ แก้ไขสินค้าแล้ว');
  } else {
    products.push({ id: genId(), name, sku, category, system, counted: null });
    showToast('✅ เพิ่มสินค้าแล้ว');
  }
  saveDataToCache();
  closeModal('productModal');
  render();
}

function resetCount() {
  requireAdmin(() => {
    const hasCounted = products.some(p => p.counted !== null && p.counted !== undefined && p.counted !== '');
    if (!confirm('ต้องการรีเซ็ตตัวเลขที่นับได้ทั้งหมดหรือไม่?\n(ระบบจะบันทึกประวัติรอบการนับปัจจุบันไว้ให้อัตโนมัติก่อนรีเซ็ต)')) return;
    if (hasCounted) {
      createHistorySession('บันทึกอัตโนมัติก่อนรีเซ็ตการนับ', true);
    }
    products.forEach(p => p.counted = null);
    saveDataToCache();
    render();
    showToast('🔄 บันทึกประวัติและรีเซ็ตการนับแล้ว');
  });
}

function clearAll() {
  requireAdmin(() => {
    if (products.length === 0) { showToast('ไม่มีข้อมูลให้ลบ'); return; }
    pendingDeleteAction = () => {
      products = [];
      saveDataToCache();
      render();
      showToast('🗑️ ลบข้อมูลทั้งหมดแล้ว');
    };
    showConfirmDelete('🗑️ ลบทั้งหมด', 'ต้องการลบสินค้าทั้งหมดหรือไม่? (ไม่สามารถกู้คืนได้)');
  });
}

function openChangePinModal() {
  requireAdmin(() => {
    document.getElementById('oldPin').value = '';
    document.getElementById('newPin').value = '';
    document.getElementById('confirmPin').value = '';
    document.getElementById('changePinModal').classList.add('show');
  });
}

function changePin() {
  const oldP = document.getElementById('oldPin').value.trim();
  const newP = document.getElementById('newPin').value.trim();
  const confP = document.getElementById('confirmPin').value.trim();
  if (oldP !== adminPin) { alert('PIN เดิมไม่ถูกต้อง'); return; }
  if (!/^\d{4}$/.test(newP)) { alert('PIN ใหม่ต้องเป็นตัวเลข 4 หลัก'); return; }
  if (newP !== confP) { alert('PIN ใหม่และยืนยันไม่ตรงกัน'); return; }
  saveAdminPin(newP);
  closeModal('changePinModal');
  showToast('✅ เปลี่ยน PIN แล้ว');
}

// ============ Modal Helpers ============
function closeModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('show');
}

// ============ Events ============
function setupEvents() {
  // Modal backdrop click
  document.querySelectorAll('.modal').forEach(m => {
    m.addEventListener('click', e => { if (e.target === m) closeModal(m.id); });
  });

  // Tabs
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentTab = tab.dataset.tab;
      renderList();
    });
  });

  // Search
  document.getElementById('searchInput').addEventListener('input', e => {
    searchTerm = e.target.value;
    renderList();
  });

  // History Search & Date Filter
  const histSearch = document.getElementById('historySearchInput');
  if (histSearch) {
    histSearch.addEventListener('input', e => {
      historySearchTerm = e.target.value;
      renderHistoryPage();
    });
  }
  const histDate = document.getElementById('historyDateFilter');
  if (histDate) {
    histDate.addEventListener('change', e => {
      historyDateFilter = e.target.value;
      renderHistoryPage();
    });
  }

  // PIN keyboard
  document.addEventListener('keydown', e => {
    if (!document.getElementById('pinModal').classList.contains('show')) return;
    if (/^\d$/.test(e.key)) pinInput(e.key);
    else if (e.key === 'Backspace') pinBackspace();
    else if (e.key === 'Escape') closeModal('pinModal');
  });
}

// ============ Init ============
setupEvents();
loadData();

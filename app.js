/* ============================================
   StockFlow — Application Logic
   ============================================ */

// ============ State ============
let products = [];
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

// File System Access API
let fileHandle = null;
let fileName = '';
let fileHasPermission = false;

// Excel column name for category (falls back to Thai header 'หมวดหมู่')
const CAT_COL = (CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.CATEGORY) || 'หมวดหมู่';
// Excel column name for unit of measure (e.g. ขวด, แพ็ค)
const UNIT_COL = (CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.UNIT) || 'หน่วยนับ';
// Excel column name for free-text details (e.g. แกะแล้ว)
const NOTE_COL = (CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.NOTE) || 'รายละเอียด';
const DEFAULT_CATEGORY = 'ไม่ระบุ';

// ============ Storage ============
function loadData() {
  try { products = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '[]'); } catch (e) { products = []; }
  try { adminPin = localStorage.getItem(CONFIG.ADMIN_KEY) || CONFIG.DEFAULT_PIN; } catch (e) { adminPin = CONFIG.DEFAULT_PIN; }
  restoreFileHandle().then(() => {
    updateAdminUI();
    render();
  });
}

function saveDataToCache() {
  localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
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

        // อ่านคำขอจากชีต "คำขอ" (ถ้ามี) แล้วรวมเข้ากับรายการในเครื่อง โดยไม่ซ้ำตาม id
        if (workbook.Sheets[REQ_SHEET_NAME]) {
          const reqData = XLSX.utils.sheet_to_json(workbook.Sheets[REQ_SHEET_NAME]);
          const local = loadRequests();
          const known = new Set(local.map(r => r.id));
          reqData.forEach(row => {
            const id = String(row[REQ_ID_COL] || '').trim();
            const text = String(row['คอมเม้น'] || '').trim();
            if (!id || !text || known.has(id)) return;
            known.add(id);
            const at = new Date(row['วันที่']);
            local.push({
              id,
              name: String(row['ชื่อ'] || 'ไม่ระบุชื่อ').trim(),
              text,
              at: isNaN(at) ? new Date().toISOString() : at.toISOString()
            });
          });
          local.sort((a, b) => new Date(b.at) - new Date(a.at));
          saveRequests(local);
          renderRequests();
        }

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
          unit: String(row[UNIT_COL] || '').trim(),
          note: String(row[NOTE_COL] || '').trim(),
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
        [UNIT_COL]: p.unit || '',
        [NOTE_COL]: p.note || '',
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

    // ชีตคำขอจากพนักงาน (เขียนทุกครั้งที่บันทึก เพื่อให้มีประวัติครบในไฟล์เดียวกัน)
    const reqRows = loadRequests().map(r => ({
      'ชื่อ': r.name,
      'คอมเม้น': r.text,
      'วันที่': new Date(r.at).toLocaleString('th-TH'),
      [REQ_ID_COL]: r.id
    }));
    const reqWs = XLSX.utils.json_to_sheet(reqRows.length ? reqRows : [{ 'ชื่อ': '', 'คอมเม้น': '', 'วันที่': '', [REQ_ID_COL]: '' }]);
    XLSX.utils.book_append_sheet(wb, reqWs, REQ_SHEET_NAME);

    const writable = await fileHandle.createWritable();
    const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    await writable.write(wbout);
    await writable.close();

    const statusEl = document.getElementById('saveStatus');
    statusEl.classList.add('success');
    statusEl.textContent = '✅ บันทึกสำเร็จ ' + new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
    showToast('💾 บันทึกทับไฟล์เดิมเรียบร้อย');
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

// ============ Admin ============
function updateAdminUI() {
  const badge = document.getElementById('adminBadge');
  const text = document.getElementById('adminText');
  const icon = document.getElementById('adminIcon');
  const banner = document.getElementById('readonlyBanner');
  const fab = document.getElementById('fabAdd');
  const adminBtns = document.querySelectorAll('.admin-only');

  if (isAdmin) {
    badge.classList.add('unlocked');
    text.textContent = 'แอดมิน';
    icon.textContent = '🔒';
    banner.style.display = 'none';
    fab.classList.remove('hidden');
    adminBtns.forEach(b => b.disabled = false);
  } else {
    badge.classList.remove('unlocked');
    text.textContent = 'โหมดทั่วไป';
    icon.textContent = '🔓';
    banner.style.display = 'block';
    fab.classList.add('hidden');
    adminBtns.forEach(b => b.disabled = true);
    if (selectMode) exitSelectMode();
  }
}

function handleAdminClick() {
  if (isAdmin) {
    if (confirm('ออกจากโหมดแอดมิน?')) {
      isAdmin = false;
      updateAdminUI();
      render();
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
// Donut (status) + per-category bars on the home page. Pure SVG/CSS, no libraries.
function renderSummaryCharts() {
  const donutOk = document.getElementById('donutOk');
  if (!donutOk) return;

  const total = products.length;
  let ok = 0, off = 0, pending = 0;
  products.forEach(p => {
    const d = getDiff(p);
    if (d === null) pending++;
    else if (d === 0) ok++;
    else off++;
  });

  const C = 2 * Math.PI * 46; // circumference ≈ 289
  const okLen = total ? (ok / total) * C : 0;
  const offLen = total ? (off / total) * C : 0;
  donutOk.setAttribute('stroke-dasharray', `${okLen} ${C}`);
  document.getElementById('donutDiff').setAttribute('stroke-dasharray', `${offLen} ${C}`);
  document.getElementById('donutDiff').setAttribute('stroke-dashoffset', `${-okLen}`);
  const pct = total ? Math.round((ok + off) / total * 100) : 0;
  document.getElementById('donutPct').textContent = pct + '%';
  document.getElementById('legOk').textContent = ok;
  document.getElementById('legDiff').textContent = off;
  document.getElementById('legPending').textContent = pending;

  // Category bars: total items vs counted items per category
  const groups = {};
  products.forEach(p => {
    const cat = (p.category || '').trim() || DEFAULT_CATEGORY;
    const g = groups[cat] || (groups[cat] = { total: 0, counted: 0 });
    g.total++;
    const d = getDiff(p);
    if (d !== null) g.counted++;
  });
  const box = document.getElementById('catBars');
  const cats = Object.keys(groups).sort((a, b) => a.localeCompare(b, 'th'));
  if (!cats.length) {
    box.innerHTML = '<div class="cat-empty">ยังไม่มีสินค้า</div>';
    return;
  }
  box.innerHTML = cats.map(c => {
    const g = groups[c];
    const pct = g.total ? Math.round(g.counted / g.total * 100) : 0;
    return `
      <div>
        <div class="cat-bar-head"><span>${escapeHtml(c)}</span><span>${g.counted}/${g.total} (${pct}%)</span></div>
        <div class="cat-bar-track"><div class="cat-bar-fill" style="width:${pct}%"></div></div>
      </div>`;
  }).join('');
}

// ============ Requests / messages from staff (home page) ============
// Local copy in localStorage; synced to cloud by sync.js and written to the Excel file as a second sheet.
const REQ_KEY = 'stockflow_requests';
const REQ_SHEET_NAME = 'Sheet1';   // ชีตที่ใช้เก็บคำขอ/คอมเม้น
const REQ_ID_COL = 'รหัส';          // คอลัมน์รหัสสำหรับกันข้อมูลซ้ำ

function loadRequests() {
  try {
    const raw = localStorage.getItem(REQ_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return Array.isArray(arr) ? arr : [];
  } catch (e) { return []; }
}

function saveRequests(list) {
  try { localStorage.setItem(REQ_KEY, JSON.stringify(list)); }
  catch (e) { alert('บันทึกคำขอไม่สำเร็จ (พื้นที่ของเบราว์เซอร์เต็มหรือถูกปิดไว้)'); }
}

function addRequest() {
  const name = document.getElementById('reqName').value.trim() || 'ไม่ระบุชื่อ';
  const text = document.getElementById('reqText').value.trim();
  if (!text) { alert('กรุณากรอกรายการที่ต้องการ'); return; }
  const list = loadRequests();
  list.unshift({ id: Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name, text, at: new Date().toISOString() });
  saveRequests(list);
  document.getElementById('reqText').value = '';
  renderRequests();
  showToast('✅ ส่งคำขอแล้ว');
}

function deleteRequest(id) {
  requireAdmin(() => {
    if (!confirm('ลบคำขอนี้หรือไม่?')) return;
    saveRequests(loadRequests().filter(r => r.id !== id));
    renderRequests();
  });
}

function renderRequests() {
  const box = document.getElementById('reqList');
  if (!box) return;
  const list = loadRequests();
  if (!list.length) {
    box.innerHTML = '<div class="req-empty">ยังไม่มีคำขอ</div>';
    return;
  }
  box.innerHTML = list.map(r => {
    const when = new Date(r.at).toLocaleString('th-TH', { dateStyle: 'short', timeStyle: 'short' });
    return `
      <div class="req-item">
        <div>
          <div class="req-meta">${escapeHtml(r.name)} · ${when}</div>
          <div>${escapeHtml(r.text)}</div>
        </div>
        ${isAdmin ? `<button class="req-del" onclick="deleteRequest('${r.id}')" title="ลบ">✕</button>` : ''}
      </div>`;
  }).join('');
}

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
  renderSummaryCharts();
  renderRequests();

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
    const cat = (p.category || '').trim() || DEFAULT_CATEGORY;
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
            ${p.note ? `<div class="product-note">${escapeHtml(p.note)}</div>` : ''}
          </div>
          ${diffHtml}
        </div>
        <div class="product-info">
          ${p.unit ? `
          <div class="info-item">
            <span class="info-label">หน่วย</span>
            <span class="info-value">${escapeHtml(p.unit)}</span>
          </div>` : ''}
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
  let cur = (p.counted === null || p.counted === undefined || p.counted === '') ? 0 : parseInt(p.counted);
  cur = Math.max(0, cur + delta);
  p.counted = cur;
  saveDataToCache();
  render();
}

function setCount(id, val) {
  const p = products.find(x => x.id === id);
  if (!p) return;
  if (val === '' || val === null) p.counted = null;
  else {
    const n = parseInt(val);
    p.counted = isNaN(n) ? null : Math.max(0, n);
  }
  saveDataToCache();
  render();
}

// ============ Admin Actions ============
// Make sure the category <select> can show a category that is not one of the presets
// (e.g. a value imported from an existing Excel file).
function setCategorySelect(value) {
  const sel = document.getElementById('inputCategory');
  const v = (value || '').trim();
  if (v && ![...sel.options].some(o => o.value === v)) {
    const opt = document.createElement('option');
    opt.value = v;
    opt.textContent = v;
    sel.appendChild(opt);
  }
  sel.value = v;
}

function openAddModal() {
  requireAdmin(() => {
    editingId = null;
    document.getElementById('modalTitle').textContent = 'เพิ่มสินค้า';
    document.getElementById('inputName').value = '';
    setCategorySelect('');
    document.getElementById('inputUnit').value = '';
    document.getElementById('inputNote').value = '';
    document.getElementById('inputSystem').value = '';
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
    setCategorySelect(p.category);
    document.getElementById('inputUnit').value = p.unit || '';
    document.getElementById('inputNote').value = p.note || '';
    document.getElementById('inputSystem').value = p.system;
    document.getElementById('productModal').classList.add('show');
  });
}

function saveProduct() {
  const name = document.getElementById('inputName').value.trim();
  const category = document.getElementById('inputCategory').value.trim();
  const unit = document.getElementById('inputUnit').value.trim();
  const note = document.getElementById('inputNote').value.trim();
  const system = parseInt(document.getElementById('inputSystem').value) || 0;
  if (!name) { alert('กรุณากรอกชื่อสินค้า'); return; }

  if (editingId) {
    // SKU is no longer edited in the modal; existing sku value is kept as-is.
    const p = products.find(x => x.id === editingId);
    if (p) { p.name = name; p.category = category; p.unit = unit; p.note = note; p.system = system; }
    showToast('✅ แก้ไขสินค้าแล้ว');
  } else {
    products.push({ id: genId(), name, sku: '', category, unit, note, system, counted: null });
    showToast('✅ เพิ่มสินค้าแล้ว');
  }
  saveDataToCache();
  closeModal('productModal');
  render();
}

function resetCount() {
  requireAdmin(() => {
    if (!confirm('ต้องการรีเซ็ตตัวเลขที่นับได้ทั้งหมดหรือไม่?')) return;
    products.forEach(p => p.counted = null);
    saveDataToCache();
    render();
    showToast('🔄 รีเซ็ตการนับแล้ว');
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
  document.getElementById(id).classList.remove('show');
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

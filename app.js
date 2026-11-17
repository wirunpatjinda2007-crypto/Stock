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

// หน้าเติมสต็อก
let restockTab = 'none';      // 'none' = ยังไม่ได้เติม, 'waiting' = รอเติม, 'done' = เติมแล้ว
let restockSearch = '';

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
          counted: (row[c.COUNTED] !== undefined && row[c.COUNTED] !== '') ? parseInt(row[c.COUNTED]) : null,
          restocked: parseInt(row[c.RESTOCKED]) || 0,
          restockStatus: parseRestockStatus(row[c.RESTOCK_STATUS], parseInt(row[c.RESTOCKED]) || 0)
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

function parseRestockStatus(text, restockedQty) {
  const t = String(text || '').trim();
  if (t) {
    if (t.includes('ยังไม่')) return 'none';
    if (t.includes('รอ')) return 'waiting';
    if (t.includes('แล้ว')) return 'done';
  }
  return restockedQty > 0 ? 'done' : 'none';
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
        [c.RESTOCK_STATUS]: (RESTOCK_STATUSES.find(m => m.key === getRestockStatus(p)) || {}).short || '',
        [c.RESTOCKED]: p.restocked > 0 ? p.restocked : '',
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

// จำนวนที่ต้องเติม = ในระบบ - นับได้ (เฉพาะรายการที่นับแล้วพบว่าขาด)
function getNeed(p) {
  const d = getDiff(p);
  return (d !== null && d < 0) ? -d : 0;
}

// ============ Admin ============
function updateAdminUI() {
  const badge = document.getElementById('adminBadge');
  const text = document.getElementById('adminText');
  const icon = document.getElementById('adminIcon');
  const adminBtns = document.querySelectorAll('.admin-only');

  if (isAdmin) {
    badge.classList.add('unlocked');
    text.textContent = 'คนเติมสต๊อก';
    icon.textContent = '🔒';
    adminBtns.forEach(b => b.disabled = false);
  } else {
    badge.classList.remove('unlocked');
    text.textContent = 'เข้าโหมดคนเติมสต๊อก';
    icon.textContent = '🔓';
    adminBtns.forEach(b => b.disabled = true);
  }
}

function handleAdminClick() {
  if (isAdmin) {
    if (confirm('ออกจากโหมดคนเติมสต๊อก?')) {
      isAdmin = false;
      updateAdminUI();
      render();
      showToast('ออกจากโหมดคนเติมสต๊อกแล้ว');
    }
  } else {
    openPinModal();
  }
}

function openPinModal(callback) {
  pinBuffer = '';
  pinCallback = callback || null;
  document.getElementById('pinTitle').textContent = callback ? '🔐 ยืนยันตัวตน' : '🔐 เข้าสู่โหมดคนเติมสต๊อก';
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
      showToast('✅ เข้าสู่โหมดคนเติมสต๊อกแล้ว');
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
      (p.category || '').toLowerCase().includes(s)
    );
  }

  if (list.length === 0) {
    container.innerHTML = `
      <div class="empty-state" style="grid-column: 1/-1;">
        <div class="icon">📦</div>
        <div class="title">ไม่พบรายการสินค้า</div>
        <div class="desc">${isAdmin ? 'เปิดไฟล์ Excel ในหน้านี้เพื่อดึงรายการสินค้า' : 'กรุณาเปิดไฟล์ Excel เพื่อดึงรายการสินค้า'}</div>
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
    return `
      <div class="product-card ${cardClass}"
           data-id="${p.id}" data-category="${escapeHtml(p.category || '')}">
        <div class="product-header">
          <div class="product-title">
            <div class="product-name">${escapeHtml(p.name)}</div>
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

}

// ============ Restock (หน้าเติมสต็อก) ============
// รายการที่ต้องเติม = นับแล้วพบว่าขาด (ในระบบ > นับได้)
// แต่ละรายการมี 3 สถานะ: ยังไม่ได้เติม / รอเติม / เติมแล้ว (เก็บใน p.restockStatus ซิงก์ขึ้น Supabase พร้อม products)
const RESTOCK_STATUSES = [
  { key: 'none',    label: 'ยังไม่ได้เติมสต๊อก', short: 'ยังไม่ได้เติม' },
  { key: 'waiting', label: 'รอเติมสต๊อก',        short: 'รอเติม' },
  { key: 'done',    label: 'เติมสต๊อกแล้ว',      short: 'เติมแล้ว' }
];

function getRestockStatus(p) {
  if (p.restockStatus === 'none' || p.restockStatus === 'waiting' || p.restockStatus === 'done') return p.restockStatus;
  return p.restocked > 0 ? 'done' : 'none';   // ข้อมูลเก่าที่มีแต่จำนวนที่เติม
}

function renderRestock() {
  const box = document.getElementById('restockList');
  if (!box) return;

  const items = products.filter(p => getRestockStatus(p) === 'done' || getNeed(p) > 0);
  const groups = { none: [], waiting: [], done: [] };
  items.forEach(p => groups[getRestockStatus(p)].push(p));
  const openQty = groups.none.concat(groups.waiting).reduce((s, p) => s + getNeed(p), 0);
  const uncounted = products.filter(p => getDiff(p) === null).length;

  document.getElementById('rsNone').textContent = groups.none.length;
  document.getElementById('rsWaiting').textContent = groups.waiting.length;
  document.getElementById('rsDone').textContent = groups.done.length;
  document.getElementById('rsQty').textContent = openQty;
  document.getElementById('rsTabCountNone').textContent = groups.none.length;
  document.getElementById('rsTabCountWaiting').textContent = groups.waiting.length;
  document.getElementById('rsTabCountDone').textContent = groups.done.length;

  let list = groups[restockTab] || [];
  if (restockSearch) {
    const s = restockSearch.toLowerCase();
    list = list.filter(p =>
      (p.name || '').toLowerCase().includes(s) ||
      (p.category || '').toLowerCase().includes(s)
    );
  }

  if (list.length === 0) {
    const meta = RESTOCK_STATUSES.find(m => m.key === restockTab);
    let desc = '';
    if (products.length === 0) desc = 'กรุณาเปิดไฟล์ในหน้า "เช็คสต็อก" ก่อน';
    else if (restockTab === 'none' && uncounted > 0) desc = `ยังไม่ได้นับอีก ${uncounted} รายการ`;
    box.innerHTML = `
      <div class="empty-state" style="grid-column: 1/-1;">
        <div class="icon">✅</div>
        <div class="title">ไม่มีรายการที่${meta.label}</div>
        <div class="desc">${desc}</div>
      </div>`;
    return;
  }

  box.innerHTML = list.map(p => {
    const st = getRestockStatus(p);
    const meta = RESTOCK_STATUSES.find(m => m.key === st);
    const need = st === 'done' ? (p.restocked || getNeed(p)) : getNeed(p);
    const unit = p.unit ? ' ' + escapeHtml(p.unit) : '';
    const cardCls = st === 'done' ? 'ok' : (st === 'waiting' ? 'over' : 'short');
    const badgeCls = cardCls;
    return `
      <div class="product-card ${cardCls}" data-id="${p.id}">
        <div class="product-header">
          <div class="product-title">
            <div class="product-name">${escapeHtml(p.name)}</div>
            ${p.note ? `<div class="product-note">${escapeHtml(p.note)}</div>` : ''}
          </div>
          <span class="diff-badge ${badgeCls}">${meta.short} ${need}${unit}</span>
        </div>
        <div class="product-info">
          <div class="info-item">
            <span class="info-label">หมวด</span>
            <span class="info-value">${escapeHtml((p.category || '').trim() || DEFAULT_CATEGORY)}</span>
          </div>
          <div class="info-item">
            <span class="info-label">ในระบบ</span>
            <span class="info-value">${p.system}</span>
          </div>
          <div class="info-item">
            <span class="info-label">นับได้</span>
            <span class="info-value">${p.counted === null || p.counted === undefined || p.counted === '' ? '-' : p.counted}</span>
          </div>
        </div>
        <div class="product-footer">
          <div class="rs-status">
            ${RESTOCK_STATUSES.map(m => `
              <button class="rs-btn ${m.key} ${m.key === st ? 'active' : ''}"
                      onclick="setRestockStatus('${p.id}', '${m.key}')">${m.label}</button>`).join('')}
          </div>
        </div>
      </div>`;
  }).join('');
}

function setRestockStatus(id, key) {
  const p = products.find(x => x.id === id);
  if (!p) return;
  if (getRestockStatus(p) === key) return;
  p.restockStatus = key;
  p.restocked = key === 'done' ? (getNeed(p) || p.restocked || 0) : 0;
  saveDataToCache();
  render();
  const meta = RESTOCK_STATUSES.find(m => m.key === key);
  showToast(`${p.name}: ${meta.label}`);
}

function render() { renderStats(); renderList(); renderRestock(); updateSaveBar(); }

// ============ Count Actions ============
// ถ้านับใหม่ ให้ล้างสถานะ "เติมแล้ว" เพราะจำนวนที่ขาดอาจเปลี่ยนไป
function changeCount(id, delta) {
  const p = products.find(x => x.id === id);
  if (!p) return;
  let cur = (p.counted === null || p.counted === undefined || p.counted === '') ? 0 : parseInt(p.counted);
  cur = Math.max(0, cur + delta);
  p.counted = cur;
  p.restocked = 0; p.restockStatus = 'none';
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
  p.restocked = 0; p.restockStatus = 'none';
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
    document.getElementById('inputRestockStatus').value = getRestockStatus(p);
    document.getElementById('productModal').classList.add('show');
  });
}

function saveProduct() {
  const name = document.getElementById('inputName').value.trim();
  const category = document.getElementById('inputCategory').value.trim();
  const unit = document.getElementById('inputUnit').value.trim();
  const note = document.getElementById('inputNote').value.trim();
  const system = parseInt(document.getElementById('inputSystem').value) || 0;
  const status = document.getElementById('inputRestockStatus').value || 'none';
  if (!name) { alert('กรุณากรอกชื่อสินค้า'); return; }

  if (editingId) {
    // SKU is no longer edited in the modal; existing sku value is kept as-is.
    const p = products.find(x => x.id === editingId);
    if (p) {
      const prevStatus = getRestockStatus(p);
      const systemChanged = p.system !== system;
      p.name = name; p.category = category; p.unit = unit; p.note = note; p.system = system;
      // เปลี่ยนจำนวนในระบบแต่ไม่ได้แตะสถานะ → ถือว่าต้องเติมใหม่
      p.restockStatus = (systemChanged && status === prevStatus) ? 'none' : status;
      p.restocked = p.restockStatus === 'done' ? (getNeed(p) || p.restocked || 0) : 0;
    }
    showToast('✅ แก้ไขสินค้าแล้ว');
  }
  saveDataToCache();
  closeModal('productModal');
  render();
}

function resetCount() {
  requireAdmin(() => {
    if (!confirm('ต้องการรีเซ็ตตัวเลขที่นับได้ทั้งหมดหรือไม่? (สถานะเติมแล้วจะถูกล้างด้วย)')) return;
    products.forEach(p => { p.counted = null; p.restocked = 0; p.restockStatus = 'none'; });
    saveDataToCache();
    render();
    showToast('🔄 รีเซ็ตการนับแล้ว');
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

  // Tabs (หน้าเช็คสต็อก)
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentTab = tab.dataset.tab;
      renderList();
    });
  });

  // Search (หน้าเช็คสต็อก)
  document.getElementById('searchInput').addEventListener('input', e => {
    searchTerm = e.target.value;
    renderList();
  });

  // Tabs + Search (หน้าเติมสต็อก)
  document.querySelectorAll('.rs-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.rs-tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      restockTab = tab.dataset.rsTab;
      renderRestock();
    });
  });
  document.getElementById('restockSearch').addEventListener('input', e => {
    restockSearch = e.target.value;
    renderRestock();
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

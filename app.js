/* ============================================
   StockFlow — Application Logic
   ============================================ */

// ============ State ============
let products = [];
let adminPin = (typeof CONFIG !== 'undefined' && CONFIG.DEFAULT_PIN) ? CONFIG.DEFAULT_PIN : '1234';
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

const CAT_COL = (typeof CONFIG !== 'undefined' && CONFIG.EXCEL_COLUMNS && CONFIG.EXCEL_COLUMNS.CATEGORY) || 'หมวดหมู่';
const DEFAULT_CATEGORY = 'ไม่ระบุ';

// ============ Storage ============
function loadData() {
  try {
    products = JSON.parse(localStorage.getItem(CONFIG.STORAGE_KEY) || '[]');
  } catch (e) { products = []; }

  try {
    adminPin = localStorage.getItem(CONFIG.ADMIN_KEY) || CONFIG.DEFAULT_PIN;
  } catch (e) { adminPin = CONFIG.DEFAULT_PIN; }

  restoreFileHandle().then(() => {
    updateAdminUI();
    render();
  }).catch(() => {
    updateAdminUI();
    render();
  });
}

function saveDataToCache() {
  try {
    localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
  } catch (e) { console.error('localStorage save failed', e); }

  // sync to cloud ถ้า jap.js โหลดแล้ว
  if (typeof window.pushToCloud === 'function') {
    window.pushToCloud();
  }
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
  if (!fileNameEl || !fileSub || !fileIcon || !openBtn || !saveBar) return;

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
        try {
          const perm = await fileHandle.queryPermission({ mode: 'readwrite' });
          fileHasPermission = perm === 'granted';
        } catch (e) { fileHasPermission = false; }
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
    }).catch(err => alert('อ่านไฟล์ไม่สำเร็จ: ' + err.message));
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

// ============ Admin ============
function updateAdminUI() {
  const badge = document.getElementById('adminBadge');
  const text = document.getElementById('adminText');
  const icon = document.getElementById('adminIcon');
  const banner = document.getElementById('readonlyBanner');
  const fab = document.getElementById('fabAdd');
  const adminBtns = document.querySelectorAll('.admin-only');
  if (!badge || !text || !icon || !banner || !fab) return;

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
  const title = document.getElementById('pinTitle');
  const err = document.getElementById('pinError');
  if (title) title.textContent = callback ? '🔐 ยืนยันตัวตน' : '🔐 เข้าสู่โหมดแอดมิน';
  if (err) err.textContent = '';
  renderPinDots();
  const modal = document.getElementById('pinModal');
  if (modal) modal.classList.add('show');
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
  const err = document.getElementById('pinError');
  if (err) err.textContent = '';
  if (pinBuffer.length === CONFIG.PIN_LENGTH) setTimeout(checkPin, 150);
}

function pinBackspace() {
  pinBuffer = pinBuffer.slice(0, -1);
  renderPinDots();
  const err = document.getElementById('pinError');
  if (err) err.textContent = '';
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
    const err = document.getElementById('pinError');
    if (err) err.textContent = '❌ PIN ไม่ถูกต้อง';
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
  // sync to cloud
  if (typeof window.pushToCloud === 'function') window.pushToCloud();
}

// ============ Render ============
function renderStats() {
  const total = products.length;
  const counted = products.filter(p => p.counted !== null && p.counted !== '' && p.counted !== undefined).length;
  const diff = products.filter(p => {
    const d = getDiff(p);
    return d !== null && d !== 0;
  }).length;

  const setText = (id, val) => { const el = document.getElementById(id); if (el) el.textContent = val; };
  setText('statTotal', total);
  setText('statCounted', counted);
  setText('statDiff', diff);
  setText('tabCountAll', total);
  setText('tabCountPending', total - counted);
  setText('tabCountDone', counted);
  setText('tabCountDiff', diff);
}

function renderList() {
  const container = document.getElementById('productList');
  if (!container) return;

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
    updateSelectionBar();
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

function render() {
  renderStats();
  renderList();
  updateSaveBar();
}

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
  const bar = document.getElementById('selectionBar');
  const btn = document.getElementById('selectModeBtn');
  const fab = document.getElementById('fabAdd');
  const saveBar = document.getElementById('saveBar');
  if (bar) bar.classList.add('show');
  if (btn) btn.textContent = '✖ ยกเลิก';
  if (fab) fab.classList.add('hidden');
  if (saveBar) saveBar.style.display = 'none';
  renderList();
}

function exitSelectMode() {
  selectMode = false;
  selectedIds.clear();
  document.body.classList.remove('select-mode');
  const bar = document.getElementById('selectionBar');
  const btn = document.getElementById('selectModeBtn');
  const fab = document.getElementById('fabAdd');
  const saveBar = document.getElementById('saveBar');
  if (bar) bar.classList.remove('show');
  if (btn) btn.textContent = '☑️ เลือกลบ';
  if (isAdmin && fab) fab.classList.remove('hidden');
  if (fileHandle && saveBar) saveBar.style.display = 'flex';
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
  const countEl = document.getElementById('selectedCount');
  const delBtn = document.getElementById('deleteSelectedBtn');
  if (countEl) countEl.textContent = n;
  if (delBtn) delBtn.disabled = n === 0;
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
  const t = document.getElementById('confirmDeleteTitle');
  const x = document.getElementById('confirmDeleteText');
  const m = document.getElementById('confirmDeleteModal');
  if (t) t.textContent = title;
  if (x) x.textContent = text;
  if (m) m.classList.add('show');
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
function fillCategoryList() {
  const list = document.getElementById('catList');
  if (!list) return;
  const cats = [...new Set(products.map(p => (p.category || '').trim()).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'th'));
  list.innerHTML = cats.map(c => `<option value="${escapeHtml(c)}"></option>`).join('');
}

function openAddModal() {
  requireAdmin(() => {
    editingId = null;
    const title = document.getElementById('modalTitle');
    if (title) title.textContent = 'เพิ่มสินค้า';
    const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
    setVal('inputName', '');
    setVal('inputSku', '');
    setVal('inputCategory', '');
    setVal('inputSystem', '');
    fillCategoryList();
    const modal = document.getElementById('productModal');
    if (modal) modal.classList.add('show');
    setTimeout(() => { const n = document.getElementById('inputName'); if (n) n.focus(); }, 200);
  });
}

function openEditModal(id) {
  requireAdmin(() => {
    const p = products.find(x => x.id === id);
    if (!p) return;
    editingId = id;
    const title = document.getElementById('modalTitle');
    if (title) title.textContent = 'แก้ไขสินค้า';
    const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
    setVal('inputName', p.name);
    setVal('inputSku', p.sku || '');
    setVal('inputCategory', p.category || '');
    setVal('inputSystem', p.system);
    fillCategoryList();
    const modal = document.getElementById('productModal');
    if (modal) modal.classList.add('show');
  });
}

function saveProduct() {
  const getVal = (id) => { const el = document.getElementById(id); return el ? el.value.trim() : ''; };
  const name = getVal('inputName');
  const sku = getVal('inputSku');
  const category = getVal('inputCategory');
  const system = parseInt(getVal('inputSystem')) || 0;
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
    const setVal = (id, v) => { const el = document.getElementById(id); if (el) el.value = v; };
    setVal('oldPin', '');
    setVal('newPin', '');
    setVal('confirmPin', '');
    const modal = document.getElementById('changePinModal');
    if (modal) modal.classList.add('show');
  });
}

function changePin() {
  const getVal = (id) => { const el = document.getElementById(id); return el ? el.value.trim() : ''; };
  const oldP = getVal('oldPin');
  const newP = getVal('newPin');
  const confP = getVal('confirmPin');
  if (oldP !== adminPin) { alert('PIN เดิมไม่ถูกต้อง'); return; }
  if (!/^\d{4}$/.test(newP)) { alert('PIN ใหม่ต้องเป็นตัวเลข 4 หลัก'); return; }
  if (newP !== confP) { alert('PIN ใหม่และยืนยันไม่ตรงกัน'); return; }
  saveAdminPin(newP);
  closeModal('changePinModal');
  showToast('✅ เปลี่ยน PIN แล้ว');
}

// ============ Modal Helpers ============
function closeModal(id) {
  const m = document.getElementById(id);
  if (m) m.classList.remove('show');
}

// ============ Events ============
function setupEvents() {
  document.querySelectorAll('.modal').forEach(m => {
    m.addEventListener('click', e => { if (e.target === m) closeModal(m.id); });
  });

  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      tab.classList.add('active');
      currentTab = tab.dataset.tab;
      renderList();
    });
  });

  const searchInput = document.getElementById('searchInput');
  if (searchInput) {
    searchInput.addEventListener('input', e => {
      searchTerm = e.target.value;
      renderList();
    });
  }

  document.addEventListener('keydown', e => {
    const pinModal = document.getElementById('pinModal');
    if (!pinModal || !pinModal.classList.contains('show')) return;
    if (/^\d$/.test(e.key)) pinInput(e.key);
    else if (e.key === 'Backspace') pinBackspace();
    else if (e.key === 'Escape') closeModal('pinModal');
  });
}

// ============ Init ============
// รอ DOM พร้อมก่อน แล้วค่อยทำงาน
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    setupEvents();
    loadData();
  });
} else {
  setupEvents();
  loadData();
}

// เปิดให้ jap.js เรียกใช้ได้
window.render = render;
window.products = products;
window.handleAdminClick = handleAdminClick;
window.openExcelFile = openExcelFile;
window.saveToFile = saveToFile;
window.toggleSelectMode = toggleSelectMode;
window.resetCount = resetCount;
window.clearAll = clearAll;
window.openChangePinModal = openChangePinModal;
window.openAddModal = openAddModal;
window.changePin = changePin;
window.deleteSingle = deleteSingle;
window.openEditModal = openEditModal;
window.saveProduct = saveProduct;
window.changeCount = changeCount;
window.setCount = setCount;
window.handleCardClick = handleCardClick;
window.deleteSelected = deleteSelected;
window.confirmDeleteExecute = confirmDeleteExecute;
window.closeModal = closeModal;
window.pinInput = pinInput;
window.pinBackspace = pinBackspace;
window.requireAdmin = requireAdmin;
window.saveAdminPin = saveAdminPin;
window.showToast = showToast;
window.saveDataToCache = saveDataToCache;

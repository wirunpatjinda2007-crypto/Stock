/* ============================================
   StockFlow — Cloud Sync (jap.js / sync.js)
   ซิงก์ข้อมูลสินค้าและประวัติการนับผ่าน Supabase
   ============================================ */

const SUPABASE_URL = 'https://jjifrsqvftsorwyupowt.supabase.co';
const SUPABASE_KEY = 'sb_publishable_Z9Xpb1A5KwO89l6Z8CosSw_RjSc1MHv';

let isSyncingFromCloud = false;

// 1. ดึงข้อมูลสินค้าและประวัติจาก Supabase มาแสดงบนหน้าเว็บ
async function pullFromCloud() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stock_state?id=in.(main,history)&select=*`, {
      headers: {
        'apikey': SUPABASE_KEY
      }
    });
    if (!res.ok) return;
    const data = await res.json();
    if (Array.isArray(data) && data.length > 0) {
      isSyncingFromCloud = true;

      // ข้อมูลสินค้าหลัก (id = 'main')
      const mainRow = data.find(r => r.id === 'main');
      if (mainRow) {
        products = Array.isArray(mainRow.products) ? mainRow.products : [];
        if (mainRow.admin_pin) {
          adminPin = mainRow.admin_pin;
          localStorage.setItem(CONFIG.ADMIN_KEY, adminPin);
        }
        localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
      }

      // ข้อมูลประวัติการนับ (id = 'history')
      const histRow = data.find(r => r.id === 'history');
      if (histRow && Array.isArray(histRow.products) && histRow.products[0]) {
        const payload = histRow.products[0];
        if (Array.isArray(payload.sessions)) {
          stockHistory = payload.sessions;
          localStorage.setItem(CONFIG.HISTORY_KEY, JSON.stringify(stockHistory));
        }
        if (Array.isArray(payload.logs)) {
          countLogs = payload.logs;
          localStorage.setItem(CONFIG.LOG_KEY, JSON.stringify(countLogs));
        }
      }

      render();
      if (typeof renderHistoryPage === 'function') renderHistoryPage();
      isSyncingFromCloud = false;
    }
  } catch (err) {
    console.error('ดึงข้อมูลออนไลน์ไม่สำเร็จ:', err);
    isSyncingFromCloud = false;
  }
}

// 2. ส่งข้อมูลสินค้าขึ้น Supabase
async function pushToCloud() {
  if (isSyncingFromCloud) return;
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/stock_state?on_conflict=id`, {
      method: 'POST',
      headers: {
        'apikey': SUPABASE_KEY,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates,return=minimal'
      },
      body: JSON.stringify({
        id: 'main',
        products: products,
        admin_pin: adminPin,
        updated_at: new Date().toISOString()
      })
    });
  } catch (err) {
    console.error('บันทึกข้อมูลออนไลน์ไม่สำเร็จ:', err);
  }
}

// 3. ส่งข้อมูลประวัติการนับขึ้น Supabase (เก็บในแถว id = 'history')
async function pushHistoryToCloud() {
  if (isSyncingFromCloud) return;
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/stock_state?on_conflict=id`, {
      method: 'POST',
      headers: {
        'apikey': SUPABASE_KEY,
        'Content-Type': 'application/json',
        'Prefer': 'resolution=merge-duplicates,return=minimal'
      },
      body: JSON.stringify({
        id: 'history',
        products: [{
          sessions: stockHistory.slice(0, 100),
          logs: countLogs.slice(0, 500)
        }],
        admin_pin: adminPin,
        updated_at: new Date().toISOString()
      })
    });
  } catch (err) {
    console.error('บันทึกประวัติออนไลน์ไม่สำเร็จ:', err);
  }
}

// 4. เชื่อมเข้ากับฟังก์ชันเดิมของ app.js อัตโนมัติ
const originalSaveDataToCache = saveDataToCache;
saveDataToCache = function() {
  originalSaveDataToCache();
  pushToCloud();
};

const originalSaveHistoryToCache = saveHistoryToCache;
saveHistoryToCache = function() {
  originalSaveHistoryToCache();
  pushHistoryToCloud();
};

const originalSaveAdminPin = saveAdminPin;
saveAdminPin = function(p) {
  originalSaveAdminPin(p);
  pushToCloud();
};

// โหลดข้อมูลเมื่อเปิดเว็บ และซิงก์ข้อมูลใหม่ทุก 5 วินาที
pullFromCloud();
setInterval(pullFromCloud, 5000);

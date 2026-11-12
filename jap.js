/* ============================================
   StockFlow — Cloud Sync (jap.js)
   ซิงก์ข้อมูลสินค้าและประวัติการนับผ่าน Supabase
   ============================================ */

const SUPABASE_URL = 'https://jjifrsqvftsorwyupowt.supabase.co';
const SUPABASE_KEY = 'sb_publishable_Z9Xpb1A5KwO89l6Z8CosSw_RjSc1MHv';

let isSyncingFromCloud = false;
let isPushingToCloud = false;
let lastCloudUpdatedAt = null;

function setOnlineIndicator(online) {
  const onlineEl = document.querySelector('.side-foot .online');
  if (!onlineEl) return;
  if (online) {
    onlineEl.style.color = '#6ee7b7';
    onlineEl.innerHTML = '<span class="dot" style="background:#10b981;box-shadow:0 0 0 3px rgba(16,185,129,0.25);"></span>ออนไลน์ (ซิงก์แล้ว)';
  } else {
    onlineEl.style.color = '#fca5a5';
    onlineEl.innerHTML = '<span class="dot" style="background:#ef4444;box-shadow:0 0 0 3px rgba(239,68,68,0.25);"></span>ออฟไลน์ (เช็คตาราง Supabase)';
  }
}

// 1. ดึงข้อมูลสินค้าและประวัติจาก Supabase มาแสดงบนหน้าเว็บ
async function pullFromCloud(forceRender = false) {
  if (isPushingToCloud) return; // ป้องกันการดึงข้อมูลเก่ามาทับขณะกำลังส่งข้อมูลใหม่ขึ้น Cloud

  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stock_state?id=in.(main,history)&select=*`, {
      headers: {
        'apikey': SUPABASE_KEY
      },
      cache: 'no-store'
    });

    if (!res.ok) {
      setOnlineIndicator(false);
      return;
    }

    const data = await res.json();
    setOnlineIndicator(true);

    if (Array.isArray(data) && data.length > 0) {
      isSyncingFromCloud = true;
      let hasChanges = false;

      // ข้อมูลสินค้าหลัก (id = 'main')
      const mainRow = data.find(r => r.id === 'main');
      if (mainRow) {
        const cloudProducts = Array.isArray(mainRow.products) ? mainRow.products : [];
        const currentJson = JSON.stringify(products);
        const cloudJson = JSON.stringify(cloudProducts);

        // อัปเดตเฉพาะเมื่อข้อมูลบน Cloud เปลี่ยนแปลงจริงๆ
        if (currentJson !== cloudJson) {
          products = cloudProducts;
          window.products = products;
          localStorage.setItem(CONFIG.STORAGE_KEY, cloudJson);
          hasChanges = true;
        }

        if (mainRow.admin_pin && mainRow.admin_pin !== adminPin) {
          adminPin = mainRow.admin_pin;
          localStorage.setItem(CONFIG.ADMIN_KEY, adminPin);
        }
        lastCloudUpdatedAt = mainRow.updated_at || null;
      }

      // ข้อมูลประวัติการนับ (id = 'history')
      const histRow = data.find(r => r.id === 'history');
      if (histRow && Array.isArray(histRow.products) && histRow.products[0]) {
        const payload = histRow.products[0];
        if (Array.isArray(payload.sessions)) {
          const newHistJson = JSON.stringify(payload.sessions);
          if (JSON.stringify(stockHistory) !== newHistJson) {
            stockHistory = payload.sessions;
            localStorage.setItem(CONFIG.HISTORY_KEY, newHistJson);
            hasChanges = true;
          }
        }
        if (Array.isArray(payload.logs)) {
          const newLogsJson = JSON.stringify(payload.logs);
          if (JSON.stringify(countLogs) !== newLogsJson) {
            countLogs = payload.logs;
            localStorage.setItem(CONFIG.LOG_KEY, newLogsJson);
            hasChanges = true;
          }
        }
      }

      // ไม่สั่ง render ทับถ้าผู้ใช้กำลังพิมพ์ตัวเลขในช่อง input อยู่
      const isTypingNumber = document.activeElement &&
                             document.activeElement.tagName === 'INPUT' &&
                             document.activeElement.type === 'number';

      if ((hasChanges || forceRender) && !isTypingNumber) {
        if (typeof render === 'function') render();
        if (typeof renderHistoryPage === 'function') renderHistoryPage();
      }

      isSyncingFromCloud = false;
    }
  } catch (err) {
    console.error('ดึงข้อมูลออนไลน์ไม่สำเร็จ:', err);
    setOnlineIndicator(false);
    isSyncingFromCloud = false;
  }
}

// 2. ส่งข้อมูลสินค้าขึ้น Supabase
async function pushToCloud() {
  if (isSyncingFromCloud) return;
  isPushingToCloud = true;
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stock_state?on_conflict=id`, {
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
    setOnlineIndicator(res.ok);
    if (!res.ok) {
      const errText = await res.text();
      console.error('Supabase pushToCloud error:', errText);
    }
  } catch (err) {
    console.error('บันทึกข้อมูลออนไลน์ไม่สำเร็จ:', err);
    setOnlineIndicator(false);
  } finally {
    isPushingToCloud = false;
  }
}

// 3. ส่งข้อมูลประวัติการนับขึ้น Supabase (เก็บในแถว id = 'history')
async function pushHistoryToCloud() {
  if (isSyncingFromCloud) return;
  isPushingToCloud = true;
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stock_state?on_conflict=id`, {
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
    setOnlineIndicator(res.ok);
  } catch (err) {
    console.error('บันทึกประวัติออนไลน์ไม่สำเร็จ:', err);
    setOnlineIndicator(false);
  } finally {
    isPushingToCloud = false;
  }
}

// 4. เชื่อมเข้ากับฟังก์ชันของ app.js อัตโนมัติ
if (typeof saveDataToCache === 'function') {
  const originalSaveDataToCache = saveDataToCache;
  saveDataToCache = function() {
    originalSaveDataToCache();
    pushToCloud();
  };
}

if (typeof saveHistoryToCache === 'function') {
  const originalSaveHistoryToCache = saveHistoryToCache;
  saveHistoryToCache = function() {
    originalSaveHistoryToCache();
    pushHistoryToCloud();
  };
}

if (typeof saveAdminPin === 'function') {
  const originalSaveAdminPin = saveAdminPin;
  saveAdminPin = function(p) {
    originalSaveAdminPin(p);
    pushToCloud();
  };
}

// โหลดข้อมูลออนไลน์เมื่อเปิดเว็บ และซิงก์ข้อมูลใหม่ทุก 5 วินาที
window.addEventListener('DOMContentLoaded', () => {
  pullFromCloud(true);
  setInterval(() => pullFromCloud(false), 5000);
});

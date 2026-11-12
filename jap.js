/* ============================================
   StockFlow — Cloud Sync (sync.js)
   ============================================ */

const SUPABASE_URL = 'วาง_PROJECT_URL_ที่นี่';
const SUPABASE_KEY = 'วาง_ANON_PUBLIC_KEY_ที่นี่';

let isSyncingFromCloud = false;

// ดึงข้อมูลจาก Supabase มาแสดงบนหน้าเว็บ
async function pullFromCloud() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stock_state?id=eq.main&select=*`, {
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`
      }
    });
    const data = await res.json();
    if (data && data.length > 0) {
      isSyncingFromCloud = true;
      products = Array.isArray(data[0].products) ? data[0].products : [];
      if (data[0].admin_pin) {
        adminPin = data[0].admin_pin;
        localStorage.setItem(CONFIG.ADMIN_KEY, adminPin);
      }
      localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify(products));
      render();
      isSyncingFromCloud = false;
    }
  } catch (err) {
    console.error('ดึงข้อมูลออนไลน์ไม่สำเร็จ:', err);
  }
}

// ส่งข้อมูลขึ้น Supabase เมื่อแอดมินเพิ่ม/แก้ไข หรือมีการนับสต๊อก
async function pushToCloud() {
  if (isSyncingFromCloud) return;
  try {
    await fetch(`${SUPABASE_URL}/rest/v1/stock_state?id=eq.main`, {
      method: 'PATCH',
      headers: {
        'apikey': SUPABASE_KEY,
        'Authorization': `Bearer ${SUPABASE_KEY}`,
        'Content-Type': 'application/json',
        'Prefer': 'return=minimal'
      },
      body: JSON.stringify({
        products: products,
        admin_pin: adminPin,
        updated_at: new Date().toISOString()
      })
    });
  } catch (err) {
    console.error('บันทึกข้อมูลออนไลน์ไม่สำเร็จ:', err);
  }
}

// เชื่อมเข้ากับฟังก์ชันเดิมของ app.js อัตโนมัติ
const originalSaveDataToCache = saveDataToCache;
saveDataToCache = function() {
  originalSaveDataToCache();
  pushToCloud();
};

const originalSaveAdminPin = saveAdminPin;
saveAdminPin = function(p) {
  originalSaveAdminPin(p);
  pushToCloud();
};

// โหลดข้อมูลเมื่อเปิดเว็บ และซิงก์ข้อมูลใหม่ทุก 5 วินาที
pullFromCloud();
setInterval(pullFromCloud, 5000);

/* ============================================
   StockFlow — Cloud Sync (sync.js)
   ============================================ */

const SUPABASE_URL = 'https://jjifrsqvftsorwyupowt.supabase.co';
const SUPABASE_KEY = 'sb_publishable_Z9Xpb1A5KwO89l6Z8CosSw_RjSc1MHv';

let isSyncingFromCloud = false;

// 1. ดึงข้อมูลจาก Supabase มาแสดงบนหน้าเว็บ
async function pullFromCloud() {
  try {
    const res = await fetch(`${SUPABASE_URL}/rest/v1/stock_state?id=eq.main&select=*`, {
      headers: {
        'apikey': SUPABASE_KEY
      }
    });
    if (!res.ok) return;
    const data = await res.json();
    if (Array.isArray(data) && data.length > 0) {
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

// 2. ส่งข้อมูลขึ้น Supabase เมื่อแอดมินเพิ่ม/แก้ไข หรือมีการนับสต๊อก
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

// 3. เชื่อมเข้ากับฟังก์ชันเดิมของ app.js อัตโนมัติ
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

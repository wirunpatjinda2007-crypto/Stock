/* ============================================
   StockFlow — Cloud Sync (jap.js)
   ============================================ */

(function() {
  'use strict';

  // ตรวจสอบว่า CONFIG พร้อมใช้งาน
  if (typeof CONFIG === 'undefined') {
    console.error('[jap.js] CONFIG not found — ตรวจสอบว่า config.js ถูกโหลดก่อน jap.js');
    return;
  }

  let isSyncingFromCloud = false;
  let isPushingToCloud = false;
  let lastCloudUpdatedAt = null;
  let pushDebounceTimer = null;

  // ============ Online Indicator ============
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

  // ============ Pull from Cloud ============
  async function pullFromCloud(forceRender = false) {
    if (isPushingToCloud) return;

    try {
      const res = await fetch(
        `${CONFIG.SUPABASE_URL}/rest/v1/stock_state?id=in.(main,history)&select=*`,
        {
          headers: { 'apikey': CONFIG.SUPABASE_KEY },
          cache: 'no-store'
        }
      );

      if (!res.ok) {
        console.warn('[jap.js] Pull failed:', res.status);
        setOnlineIndicator(false);
        return;
      }

      const data = await res.json();
      setOnlineIndicator(true);

      if (!Array.isArray(data) || data.length === 0) return;

      isSyncingFromCloud = true;
      let hasChanges = false;

      // ---- Main products ----
      const mainRow = data.find(r => r.id === 'main');
      if (mainRow) {
        const cloudProducts = Array.isArray(mainRow.products) ? mainRow.products : [];
        const currentJson = JSON.stringify(window.products || []);
        const cloudJson = JSON.stringify(cloudProducts);

        if (currentJson !== cloudJson) {
          window.products = cloudProducts;
          localStorage.setItem(CONFIG.STORAGE_KEY, cloudJson);
          hasChanges = true;
        }

        if (mainRow.admin_pin) {
          const currentPin = localStorage.getItem(CONFIG.ADMIN_KEY);
          if (mainRow.admin_pin !== currentPin) {
            localStorage.setItem(CONFIG.ADMIN_KEY, mainRow.admin_pin);
          }
        }
        lastCloudUpdatedAt = mainRow.updated_at || null;
      }

      // ---- History (ถ้ามี) ----
      const histRow = data.find(r => r.id === 'history');
      if (histRow && Array.isArray(histRow.products) && histRow.products[0]) {
        const payload = histRow.products[0];
        if (Array.isArray(payload.sessions) && typeof CONFIG.HISTORY_KEY !== 'undefined') {
          localStorage.setItem(CONFIG.HISTORY_KEY, JSON.stringify(payload.sessions));
        }
        if (Array.isArray(payload.logs) && typeof CONFIG.LOG_KEY !== 'undefined') {
          localStorage.setItem(CONFIG.LOG_KEY, JSON.stringify(payload.logs));
        }
      }

      // ไม่ render ทับถ้าผู้ใช้กำลังพิมพ์อยู่
      const active = document.activeElement;
      const isTyping = active && active.tagName === 'INPUT' && (active.type === 'number' || active.type === 'text');

      if ((hasChanges || forceRender) && !isTyping && typeof window.render === 'function') {
        window.render();
      }

      isSyncingFromCloud = false;
    } catch (err) {
      console.error('[jap.js] pullFromCloud error:', err);
      setOnlineIndicator(false);
      isSyncingFromCloud = false;
    }
  }

  // ============ Push to Cloud ============
  async function pushToCloud() {
    if (isSyncingFromCloud || isPushingToCloud) return;

    // debounce 300ms กันยิงถี่เกินไป
    clearTimeout(pushDebounceTimer);
    pushDebounceTimer = setTimeout(async () => {
      isPushingToCloud = true;
      try {
        const res = await fetch(
          `${CONFIG.SUPABASE_URL}/rest/v1/stock_state?on_conflict=id`,
          {
            method: 'POST',
            headers: {
              'apikey': CONFIG.SUPABASE_KEY,
              'Content-Type': 'application/json',
              'Prefer': 'resolution=merge-duplicates,return=minimal'
            },
            body: JSON.stringify({
              id: 'main',
              products: window.products || [],
              admin_pin: localStorage.getItem(CONFIG.ADMIN_KEY) || CONFIG.DEFAULT_PIN,
              updated_at: new Date().toISOString()
            })
          }
        );
        setOnlineIndicator(res.ok);
        if (!res.ok) {
          const errText = await res.text();
          console.error('[jap.js] pushToCloud failed:', res.status, errText);
        }
      } catch (err) {
        console.error('[jap.js] pushToCloud error:', err);
        setOnlineIndicator(false);
      } finally {
        isPushingToCloud = false;
      }
    }, 300);
  }

  // ============ Push History ============
  async function pushHistoryToCloud() {
    if (isSyncingFromCloud || isPushingToCloud) return;
    try {
      let sessions = [];
      let logs = [];
      try { sessions = JSON.parse(localStorage.getItem(CONFIG.HISTORY_KEY) || '[]'); } catch(e) {}
      try { logs = JSON.parse(localStorage.getItem(CONFIG.LOG_KEY) || '[]'); } catch(e) {}

      const res = await fetch(
        `${CONFIG.SUPABASE_URL}/rest/v1/stock_state?on_conflict=id`,
        {
          method: 'POST',
          headers: {
            'apikey': CONFIG.SUPABASE_KEY,
            'Content-Type': 'application/json',
            'Prefer': 'resolution=merge-duplicates,return=minimal'
          },
          body: JSON.stringify({
            id: 'history',
            products: [{ sessions: sessions.slice(0, 100), logs: logs.slice(0, 500) }],
            updated_at: new Date().toISOString()
          })
        }
      );
      setOnlineIndicator(res.ok);
    } catch (err) {
      console.error('[jap.js] pushHistoryToCloud error:', err);
    }
  }

  // ============ Export to window (ให้ app.js เรียกใช้) ============
  window.pushToCloud = pushToCloud;
  window.pullFromCloud = pullFromCloud;
  window.pushHistoryToCloud = pushHistoryToCloud;

  // ============ Auto Start ============
  function startSync() {
    pullFromCloud(true);
    setInterval(() => pullFromCloud(false), CONFIG.SYNC_INTERVAL_MS || 5000);
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', startSync);
  } else {
    startSync();
  }

})();

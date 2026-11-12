/* ============================================
   StockFlow — Staff login (login.js)
   พนักงานเข้าด้วยรหัสพนักงาน + PIN ที่ตั้งเอง
   ต้องรัน staff_setup.sql ใน Supabase ก่อน
   ============================================ */
(function () {
  const SB_URL = 'https://jjifrsqvftsorwyupowt.supabase.co';
  const SB_KEY = 'sb_publishable_Z9Xpb1A5KwO89l6Z8CosSw_RjSc1MHv';
  const SESSION_KEY = 'stockflow_staff';

  const STYLE = `
    #loginOverlay { position: fixed; inset: 0; background: rgba(10,34,89,.55); display: none;
      align-items: center; justify-content: center; z-index: 9999; padding: 16px; }
    #loginOverlay.show { display: flex; }
    .login-box { background: #fff; border-radius: 18px; padding: 28px 24px; width: 100%; max-width: 360px;
      box-shadow: 0 20px 50px rgba(0,0,0,.25); box-sizing: border-box; }
    .login-box h2 { margin: 0 0 6px; font-size: 20px; color: #0f1d36; }
    .login-box .sub { margin: 0 0 14px; font-size: 14px; color: #64748b; }
    .login-box label { display: block; font-size: 13px; color: #334155; font-weight: 600; margin: 12px 0 6px; }
    .login-box input { width: 100%; padding: 12px 14px; font: inherit; font-size: 16px; border: 1px solid #cbd5e1;
      border-radius: 10px; box-sizing: border-box; }
    .login-box input:focus { outline: none; border-color: #2563eb; box-shadow: 0 0 0 3px rgba(37,99,235,.15); }
    .login-box input:disabled { background: #f1f5f9; color: #64748b; }
    .login-box [hidden] { display: none !important; }
    .login-box .main-btn { width: 100%; margin-top: 18px; padding: 13px; border: none; border-radius: 12px;
      background: #1d4ed8; color: #fff; font: inherit; font-size: 16px; font-weight: 600; cursor: pointer; }
    .login-box .main-btn:disabled { opacity: .6; cursor: wait; }
    .login-box .link-btn { display: block; margin: 12px auto 0; border: none; background: none; color: #2563eb;
      font: inherit; font-size: 13px; cursor: pointer; }
    .login-msg { min-height: 20px; margin-top: 10px; font-size: 13px; color: #dc2626; }
    #reqName { display: none !important; }
  `;

  const MARKUP = `
    <div id="loginOverlay">
      <form class="login-box" id="loginForm" autocomplete="off">
        <h2 id="loginTitle">เข้าสู่ระบบ</h2>
        <p class="sub" id="loginSub">กรอกรหัสพนักงานของคุณ</p>
        <label for="loginCode">รหัสพนักงาน</label>
        <input id="loginCode" type="text" autocapitalize="characters" spellcheck="false">
        <div id="pinWrap" hidden>
          <label for="loginPin">PIN</label>
          <input id="loginPin" type="password" inputmode="numeric" maxlength="6">
        </div>
        <div id="pinNewWrap" hidden>
          <label for="loginPin2">PIN ใหม่ (ตัวเลข 4–6 หลัก)</label>
          <input id="loginPin2" type="password" inputmode="numeric" maxlength="6">
          <label for="loginPin3">ยืนยัน PIN อีกครั้ง</label>
          <input id="loginPin3" type="password" inputmode="numeric" maxlength="6">
        </div>
        <div class="login-msg" id="loginMsg"></div>
        <button class="main-btn" type="submit" id="loginBtn">ถัดไป</button>
        <button class="link-btn" type="button" id="loginBack" hidden>เปลี่ยนรหัสพนักงาน</button>
      </form>
    </div>`;

  const MSG = {
    invalid: 'ไม่พบรหัสพนักงานนี้',
    wrong: 'PIN ไม่ถูกต้อง',
    locked: 'ลองผิดหลายครั้ง กรุณารอ 5 นาทีแล้วลองใหม่',
    bad_format: 'PIN ต้องเป็นตัวเลข 4–6 หลัก',
    already_set: 'รหัสนี้ตั้ง PIN ไปแล้ว กรุณาเข้าสู่ระบบด้วย PIN'
  };

  let step = 'code';          // 'code' | 'pin' | 'setpin'
  let pendingCode = '';
  let pendingId = null;
  let pendingName = '';

  const $ = (id) => document.getElementById(id);

  // ---------- session ----------
  function getStaff() {
    try { return JSON.parse(localStorage.getItem(SESSION_KEY) || 'null'); } catch (e) { return null; }
  }
  function setStaff(s) {
    try { localStorage.setItem(SESSION_KEY, JSON.stringify(s)); } catch (e) { /* ใช้งานต่อได้ แต่ต้องล็อกอินใหม่ตอนรีเฟรช */ }
  }
  function clearStaff() {
    try { localStorage.removeItem(SESSION_KEY); } catch (e) {}
  }

  function applyStaffName() {
    const s = getStaff();
    const label = s ? s.name : 'ยังไม่ได้เข้าสู่ระบบ';
    ['userName', 'sideUserName'].forEach(id => { const el = $(id); if (el) el.textContent = label; });
  }

  // ---------- server calls ----------
  async function rpc(fn, body) {
    const res = await fetch(`${SB_URL}/rest/v1/rpc/${fn}`, {
      method: 'POST',
      headers: { 'apikey': SB_KEY, 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    });
    if (!res.ok) throw new Error('เชื่อมต่อเซิร์ฟเวอร์ไม่สำเร็จ');
    return res.json();
  }

  // ---------- UI ----------
  function setMsg(text) { $('loginMsg').textContent = text || ''; }

  function showStep(s) {
    step = s;
    $('pinWrap').hidden = s !== 'pin';
    $('pinNewWrap').hidden = s !== 'setpin';
    $('loginCode').disabled = s !== 'code';
    $('loginBack').hidden = s === 'code';
    $('loginBtn').textContent = s === 'code' ? 'ถัดไป' : s === 'pin' ? 'เข้าสู่ระบบ' : 'ตั้ง PIN และเข้าสู่ระบบ';
    $('loginTitle').textContent = s === 'setpin' ? 'ตั้ง PIN ของคุณ' : 'เข้าสู่ระบบ';
    $('loginSub').textContent =
      s === 'code' ? 'กรอกรหัสพนักงานของคุณ' :
      s === 'pin' ? `สวัสดี ${pendingName} กรอก PIN ของคุณ` :
      `ยินดีต้อนรับ ${pendingName} ตั้ง PIN ใหม่สำหรับเข้าใช้งานครั้งต่อไป`;
    setMsg('');
    setTimeout(() => {
      const focusEl = s === 'code' ? $('loginCode') : s === 'pin' ? $('loginPin') : $('loginPin2');
      if (focusEl && !focusEl.disabled) focusEl.focus();
    }, 50);
  }

  function showOverlay() {
    $('loginOverlay').classList.add('show');
    $('loginCode').value = '';
    $('loginPin').value = '';
    $('loginPin2').value = '';
    $('loginPin3').value = '';
    pendingCode = ''; pendingId = null; pendingName = '';
    showStep('code');
  }

  function hideOverlay() {
    $('loginOverlay').classList.remove('show');
  }

  function finishLogin(r) {
    setStaff({ id: r.id, name: r.name });
    applyStaffName();
    hideOverlay();
  }

  // ---------- form ----------
  async function onSubmit(e) {
    e.preventDefault();
    const btn = $('loginBtn');
    btn.disabled = true;
    setMsg('');
    try {
      if (step === 'code') {
        const code = $('loginCode').value.trim();
        if (!code) { setMsg('กรุณากรอกรหัสพนักงาน'); return; }
        const r = await rpc('staff_lookup', { p_code: code });
        if (r.status === 'invalid') { setMsg(MSG.invalid); return; }
        if (r.status === 'locked') { setMsg(MSG.locked); return; }
        pendingCode = code;
        pendingId = r.id;
        pendingName = r.name;
        showStep(r.status === 'needs_pin' ? 'setpin' : 'pin');
      } else if (step === 'pin') {
        const pin = $('loginPin').value;
        if (!pin) { setMsg('กรุณากรอก PIN'); return; }
        const r = await rpc('staff_login', { p_code: pendingCode, p_pin: pin });
        if (r.status === 'ok') { finishLogin(r); $('loginPin').value = ''; }
        else setMsg(MSG[r.status] || 'เข้าสู่ระบบไม่สำเร็จ');
      } else if (step === 'setpin') {
        const a = $('loginPin2').value;
        const b = $('loginPin3').value;
        if (a !== b) { setMsg('PIN ทั้งสองช่องไม่ตรงกัน'); return; }
        const set = await rpc('staff_set_pin', { p_id: pendingId, p_pin: a });
        if (set.status !== 'ok') { setMsg(MSG[set.status] || 'ตั้ง PIN ไม่สำเร็จ'); return; }
        const r = await rpc('staff_login', { p_code: pendingCode, p_pin: a });
        if (r.status === 'ok') finishLogin(r);
        else setMsg('เข้าสู่ระบบไม่สำเร็จ ลองใหม่อีกครั้ง');
      }
    } catch (err) {
      setMsg(err.message || 'เกิดข้อผิดพลาด ลองใหม่อีกครั้ง');
    } finally {
      btn.disabled = false;
    }
  }

  // ---------- hook into existing app ----------
  // ชื่อผู้ส่งคำขอ = ชื่อพนักงานที่เข้าสู่ระบบ
  const originalAddRequest = window.addRequest;
  window.addRequest = function () {
    const s = getStaff();
    if (!s) { showOverlay(); return; }
    const el = $('reqName');
    if (el) el.value = s.name;
    originalAddRequest();
  };

  // ออกจากระบบ = ล้างเซสชันพนักงาน และปิดโหมดผู้ดูแลถ้าเปิดอยู่
  window.logout = function () {
    if (!confirm('ต้องการออกจากระบบหรือไม่?')) return;
    if (typeof isAdmin === 'function' && isAdmin()) {
      const badge = $('adminBadge');
      if (badge) badge.click();
    }
    clearStaff();
    applyStaffName();
    showOverlay();
  };

  function init() {
    const st = document.createElement('style');
    st.textContent = STYLE;
    document.head.appendChild(st);
    document.body.insertAdjacentHTML('beforeend', MARKUP);

    $('loginForm').addEventListener('submit', onSubmit);
    $('loginBack').addEventListener('click', () => {
      pendingCode = ''; pendingId = null; pendingName = '';
      $('loginCode').value = '';
      showStep('code');
    });

    applyStaffName();
    if (getStaff()) hideOverlay();
    else showOverlay();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
  else init();
})();

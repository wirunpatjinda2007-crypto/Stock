/* ============================================================
   sync.js — ซิงก์ StockFlow กับ Google Sheets
   เวอร์ชัน: ใช้ no-cors เพื่อเลี่ยงปัญหา CORS ของ Apps Script
   วิธีใช้: แค่ใส่ URL ของ Apps Script ด้านล่าง
   ============================================================ */

(function () {
  "use strict";

  // ==== ใส่ URL ของ Google Apps Script Web App ที่นี่ ====
  // ต้องลงท้ายด้วย /exec และ Deploy แบบ Who has access: Anyone
  const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxJN1g7KnAhKuvfoY4Ba9dyRNJP1eQyMVUNq2MchpZfmXK1AklJNm2ZSWwYO3NG-pPQ/exec";

  // ชื่อฟังก์ชันบันทึกเดิมในแอป (ตาม index.html คือ saveToFile)
  const SAVE_FN_NAME = "saveToFile";

  function showToast(msg, color) {
    const t = document.getElementById("toast");
    if (t) {
      t.textContent = msg;
      t.style.background = color || "#333";
      t.classList.add("show");
      setTimeout(() => t.classList.remove("show"), 2500);
    }
    console.log("[sync]", msg);
  }

  // ดึงข้อมูลสินค้าจาก state ของแอป
  function getProducts() {
    const candidates = ["products", "PRODUCTS", "items", "stockData", "state"];
    for (const name of candidates) {
      const v = window[name];
      if (Array.isArray(v)) return v;
      if (v && Array.isArray(v.products)) return v.products;
    }
    // fallback: อ่านจาก DOM ที่มี data-*
    const nodes = document.querySelectorAll("[data-id]");
    if (nodes.length) {
      return Array.from(nodes).map((n) => ({
        id: n.dataset.id,
        name: n.dataset.name || "",
        sku: n.dataset.sku || "",
        system: Number(n.dataset.system || 0),
        counted: Number(n.dataset.counted || 0),
        diff: Number(n.dataset.diff || 0),
      }));
    }
    return null;
  }

  async function pushToSheet(products) {
    if (!SCRIPT_URL || SCRIPT_URL.indexOf("วาง_URL") === 0) {
      showToast("⚠️ ยังไม่ได้ใส่ URL ใน sync.js", "#f59e0b");
      return false;
    }
    if (!Array.isArray(products)) {
      showToast("⚠️ อ่านข้อมูลสินค้าไม่ได้ (ไม่พบ state)", "#f59e0b");
      return false;
    }

    try {
      // ใช้ no-cors เพื่อเลี่ยงการบล็อกจาก CORS ของ Apps Script
      // ข้อจำกัด: อ่าน response ไม่ได้ จึงไม่รู้ว่า success/error
      // ต้องไปตรวจที่ Google Sheets เอง
      await fetch(SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "syncAll",
          products: products,
          sentAt: new Date().toISOString(),
        }),
      });

      showToast("✅ ส่งข้อมูลแล้ว (ตรวจสอบที่ Google Sheets)", "#16a34a");
      return true;
    } catch (err) {
      showToast("❌ ส่งไม่สำเร็จ: " + err.message, "#dc2626");
      return false;
    }
  }

  // ==== ห่อ (wrap) ฟังก์ชัน saveToFile เดิม ====
  function wrapSaveFunction() {
    const original = window[SAVE_FN_NAME];
    if (typeof original !== "function") {
      console.warn(
        "[sync] ไม่พบฟังก์ชัน " + SAVE_FN_NAME + " — จะผูกกับปุ่ม saveBtn แทน"
      );
      const btn = document.getElementById("saveBtn");
      if (btn) {
        btn.addEventListener("click", () => {
          setTimeout(() => {
            const p = getProducts();
            if (p) pushToSheet(p);
          }, 800);
        });
      }
      return;
    }

    window[SAVE_FN_NAME] = async function () {
      const ret = original.apply(this, arguments);
      const p = getProducts();
      if (p) await pushToSheet(p);
      return ret;
    };
    console.log("[sync] ผูกกับ " + SAVE_FN_NAME + " เรียบร้อย");
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", wrapSaveFunction);
  } else {
    wrapSaveFunction();
  }
})();

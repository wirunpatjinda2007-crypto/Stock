/* ============================================================
   sync.js — ซิงก์ StockFlow กับ Google Sheets
   วิธีใช้: แค่ใส่ URL ของ Apps Script ด้านล่าง
   ไม่ต้องแก้ app.js หรือ index.html ส่วนอื่น
   ============================================================ */

(function () {
  "use strict";

  // ==== ใส่ URL ของ Google Apps Script Web App ที่นี่ ====
  const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbxJN1g7KnAhKuvfoY4Ba9dyRNJP1eQyMVUNq2MchpZfmXK1AklJNm2ZSWwYO3NG-pPQ/exec";

  // ชื่อฟังก์ชันบันทึกเดิมในแอปคุณ (ตาม index.html คือ saveToFile)
  const SAVE_FN_NAME = "saveToFile";

  function showToast(msg, color) {
    const t = document.getElementById("toast");
    if (t) {
      t.textContent = msg;
      t.style.background = color || "#333";
      t.classList.add("show");
      setTimeout(() => t.classList.remove("show"), 2500);
    } else {
      console.log("[sync]", msg);
    }
  }

  // ดึงข้อมูลสินค้าจาก state ของแอป
  // ลองหาตัวแปร global ก่อน ถ้าไม่เจอจะ fallback ไปอ่านจาก DOM
  function getProducts() {
    // ลองชื่อตัวแปรที่พบบ่อยในแอปแนวนี้
    const candidates = ["products", "PRODUCTS", "items", "stockData", "state"];
    for (const name of candidates) {
      const v = window[name];
      if (Array.isArray(v)) return v;
      if (v && Array.isArray(v.products)) return v.products;
    }
    // fallback: อ่านจาก DOM ที่มี data-* (ถ้าแอปคุณใส่ data-id ไว้)
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
      // ส่งเป็น JSON string ผ่าน text/plain เพื่อเลี่ยง CORS preflight
      const res = await fetch(SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({
          action: "syncAll",
          products: products,
          sentAt: new Date().toISOString(),
        }),
      });

      const json = await res.json();
      if (json && json.result === "success") {
        showToast("✅ ซิงก์ขึ้น Google Sheets แล้ว", "#16a34a");
        return true;
      }
      showToast("❌ Sheets ตอบกลับผิดปกติ", "#dc2626");
      return false;
    } catch (err) {
      showToast("❌ เชื่อมต่อ Sheets ไม่ได้: " + err.message, "#dc2626");
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
          // รอให้แอปเดิมบันทึกเสร็จก่อน (หน่วงเล็กน้อย)
          setTimeout(() => {
            const p = getProducts();
            if (p) pushToSheet(p);
          }, 800);
        });
      }
      return;
    }

    window[SAVE_FN_NAME] = async function () {
      // เรียกฟังก์ชันเดิมก่อน (ให้บันทึก Excel ในเครื่องตามปกติ)
      const ret = original.apply(this, arguments);
      // แล้วค่อยซิงก์ขึ้น Sheets
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

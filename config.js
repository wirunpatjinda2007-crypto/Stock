/* ============================================
   StockFlow — Configuration
   ค่าคงที่ทั้งหมดของระบบ อยู่ที่นี่ที่เดียว
   ============================================ */

const CONFIG = {
  // Storage Keys (ใช้ใน localStorage & IndexedDB)
  STORAGE_KEY: 'stockflow_data_v5',
  HISTORY_KEY: 'stockflow_history_v1',
  LOG_KEY: 'stockflow_logs_v1',
  ADMIN_KEY: 'stockflow_admin_v1',
  DB_NAME: 'StockFlowDB',

  // ค่าเริ่มต้น
  DEFAULT_PIN: '1234',
  PIN_LENGTH: 4,

  // ชื่อคอลัมน์ใน Excel (ถ้าเปลี่ยนต้องอัปเดตทั้ง import & export)
  EXCEL_COLUMNS: {
    ID: 'ID',
    NAME: 'ชื่อสินค้า',
    SKU: 'SKU',
    CATEGORY: 'หมวดหมู่',
    SYSTEM: 'ในระบบ',
    COUNTED: 'นับได้',
    DIFF: 'ผลต่าง',
    STATUS: 'สถานะ',
    UPDATED_AT: 'วันที่อัปเดต'
  },

  // ชื่อ Sheet ในไฟล์ Excel
  SHEET_NAME: 'Stock'
};

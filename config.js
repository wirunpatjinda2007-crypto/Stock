/* ============================================
   StockFlow — Configuration
   ============================================ */

const CONFIG = {
  STORAGE_KEY: 'stockflow_data_v5',
  ADMIN_KEY: 'stockflow_admin_v1',
  HISTORY_KEY: 'stockflow_history_v1',
  LOG_KEY: 'stockflow_log_v1',
  DB_NAME: 'StockFlowDB',
  AUDIT_KEY: 'stockflow_audit_log_v1',

  DEFAULT_PIN: '1234',
  PIN_LENGTH: 4,

  EXCEL_COLUMNS: {
    ID: 'ID',
    NAME: 'ชื่อสินค้า',
    SKU: 'SKU',
    SYSTEM: 'ในระบบ',
    COUNTED: 'นับได้',
    DIFF: 'ผลต่าง',
    STATUS: 'สถานะ',
    UPDATED_AT: 'วันที่อัปเดต'
  },

  SHEET_NAME: 'Stock',

  // Supabase
  SUPABASE_URL: 'https://jjifrsqvftsorwyupowt.supabase.co',
  SUPABASE_KEY: 'sb_publishable_Z9Xpb1A5KwO89l6Z8CosSw_RjSc1MHv',
  SYNC_INTERVAL_MS: 5000
};

// กันไม่ให้ตัวแปรหายเมื่อโหลดซ้ำ
window.CONFIG = CONFIG;

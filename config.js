// ==========================================
// config.js — ไฟล์เก็บค่าการเชื่อมต่อ Supabase
// ⚠️ ห้าม commit ไฟล์นี้ขึ้น GitHub สาธารณะ!
// ==========================================

// 1. Project URL — หาได้จาก Supabase → Project Settings → API → Project URL
//    รูปแบบ: https://[ชื่อโปรเจกต์].supabase.co
const SUPABASE_URL = 'https://supabase.com/dashboard/project/tsgraphiopdkygjdgsok/editor/17715?schema=public';

// 2. anon key (public) — หาได้จาก Supabase → Project Settings → API → Project API keys
//    ต้องขึ้นต้นด้วย eyJ... (ยาวๆ)
//    ⚠️ ห้ามใช้ service_role หรือ secret key เด็ดขาด!
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9...';

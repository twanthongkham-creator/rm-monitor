-- ============================================================
-- STOCK MONITORING SYSTEM - SUPABASE SCHEMA
-- ============================================================
-- Run this in Supabase SQL Editor

-- 1. กลุ่มวัตถุดิบ (Product Groups)
CREATE TABLE IF NOT EXISTS product_groups (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  name TEXT NOT NULL UNIQUE,
  sort_order INTEGER DEFAULT 0,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. Master วัตถุดิบ
-- NOTE: ถ้าตารางมีอยู่แล้ว ให้รัน SQL ด้านล่างเพื่อเพิ่มคอลัมน์:
-- ALTER TABLE raw_materials ADD COLUMN IF NOT EXISTS rm_code TEXT;
CREATE TABLE IF NOT EXISTS raw_materials (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  code TEXT NOT NULL,
  rm_code TEXT,                -- รหัสวัตถุดิบ (RM Code) จาก Excel
  product_name TEXT,           -- สูตร/ชื่อผลิตภัณฑ์ที่ใช้วัตถุดิบนี้
  name TEXT,                   -- ชื่อวัตถุดิบ (ใช้ในสูตร)
  recipe_step TEXT,            -- ขั้นตอนในสูตร (Step / RM_Name)
  unit_package TEXT,           -- รายละเอียด package
  unit TEXT,                   -- หน่วย (KG, UN, ขวด)
  unit_per_pack NUMERIC DEFAULT 0, -- จำนวนต่อ Pack (Unit ที่ผลิตได้)
  supplier TEXT,               -- Supplier
  remark TEXT,                 -- หมายเหตุ
  is_monitored BOOLEAN DEFAULT TRUE, -- ใช้เลือกว่าติดตามวัตถุดิบตัวนี้หรือไม่
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Index สำหรับ code
CREATE INDEX IF NOT EXISTS idx_raw_materials_code ON raw_materials(code);
CREATE INDEX IF NOT EXISTS idx_raw_materials_product_name ON raw_materials(product_name);

-- 3. Stock Entries (บันทึก stock รายวัน)
CREATE TABLE IF NOT EXISTS stock_entries (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  product_group TEXT NOT NULL,         -- กลุ่มวัตถุดิบ (sheet name)
  code TEXT NOT NULL,                  -- รหัสสินค้า SAP
  description TEXT,                    -- ชื่อ
  new_description TEXT,                -- ชื่อใหม่
  shelf_life TEXT,                     -- อายุสินค้า (เช่น 1 ปี, 6 เดือน)
  batch TEXT,                          -- Batch number
  received_date DATE,                  -- วันที่รับสินค้า
  production_date DATE,                -- วันที่ผลิต
  expiry_date DATE,                    -- วันหมดอายุ (จากผู้ผลิต)
  extended_expiry_1 DATE,              -- วันหมดอายุจากการต่ออายุ ครั้งที่ 1
  extended_expiry_2 DATE,              -- วันหมดอายุจากการต่ออายุ ครั้งที่ 2
  extended_expiry_1_doc TEXT,          -- ลิงก์/Base64 เอกสารขอต่ออายุ ครั้งที่ 1
  extended_expiry_2_doc TEXT,          -- ลิงก์/Base64 เอกสารขอต่ออายุ ครั้งที่ 2
  quantity NUMERIC,                    -- จำนวน (Units หรือ KG)
  check_date DATE DEFAULT CURRENT_DATE, -- วันที่ตรวจเช็ค
  notes TEXT,                          -- หมายเหตุ
  created_at TIMESTAMPTZ DEFAULT NOW(),
  updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- Indexes
CREATE INDEX IF NOT EXISTS idx_stock_entries_code ON stock_entries(code);
CREATE INDEX IF NOT EXISTS idx_stock_entries_check_date ON stock_entries(check_date);
CREATE INDEX IF NOT EXISTS idx_stock_entries_expiry ON stock_entries(expiry_date);
CREATE INDEX IF NOT EXISTS idx_stock_entries_group ON stock_entries(product_group);
CREATE INDEX IF NOT EXISTS idx_stock_entries_batch ON stock_entries(batch);

-- 4. View: Stock พร้อมคำนวณวันหมดอายุ
CREATE OR REPLACE VIEW stock_with_expiry AS
SELECT
  se.*,
  rm.product_name,
  rm.name AS rm_full_name,
  rm.unit_per_pack,
  rm.supplier,
  rm.is_monitored,
  -- วันหมดอายุที่ใช้งาน (ใช้ extended ถ้ามี)
  COALESCE(se.extended_expiry_2, se.extended_expiry_1, se.expiry_date) AS effective_expiry_date,
  -- วันที่เหลือ (คำนวณจากวันนี้)
  (COALESCE(se.extended_expiry_2, se.extended_expiry_1, se.expiry_date) - CURRENT_DATE) AS days_remaining,
  -- สถานะ
  CASE
    WHEN COALESCE(se.extended_expiry_2, se.extended_expiry_1, se.expiry_date) < CURRENT_DATE THEN 'expired'
    WHEN COALESCE(se.extended_expiry_2, se.extended_expiry_1, se.expiry_date) <= CURRENT_DATE + 30 THEN 'critical'
    WHEN COALESCE(se.extended_expiry_2, se.extended_expiry_1, se.expiry_date) <= CURRENT_DATE + 60 THEN 'warning'
    ELSE 'ok'
  END AS expiry_status,
  -- หน่วยที่ผลิตได้
  COALESCE(se.quantity * rm.unit_per_pack, 0) AS production_units
FROM stock_entries se
LEFT JOIN raw_materials rm ON rm.code = se.code;

-- 5. View: สรุปยอด stock ล่าสุดต่อ code+batch
CREATE OR REPLACE VIEW latest_stock_summary AS
SELECT DISTINCT ON (code, batch)
  code,
  batch,
  description,
  new_description,
  product_group,
  quantity,
  production_date,
  expiry_date,
  extended_expiry_1,
  extended_expiry_2,
  extended_expiry_1_doc,
  extended_expiry_2_doc,
  COALESCE(extended_expiry_2, extended_expiry_1, expiry_date) AS effective_expiry_date,
  (COALESCE(extended_expiry_2, extended_expiry_1, expiry_date) - CURRENT_DATE) AS days_remaining,
  CASE
    WHEN COALESCE(extended_expiry_2, extended_expiry_1, expiry_date) < CURRENT_DATE THEN 'expired'
    WHEN COALESCE(extended_expiry_2, extended_expiry_1, expiry_date) <= CURRENT_DATE + 30 THEN 'critical'
    WHEN COALESCE(extended_expiry_2, extended_expiry_1, expiry_date) <= CURRENT_DATE + 60 THEN 'warning'
    ELSE 'ok'
  END AS expiry_status,
  check_date,
  received_date
FROM stock_entries
ORDER BY code, batch, check_date DESC;

-- 6. Enable Row Level Security (ถ้าต้องการ)
-- ALTER TABLE raw_materials ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE stock_entries ENABLE ROW LEVEL SECURITY;
-- ALTER TABLE product_groups ENABLE ROW LEVEL SECURITY;

-- สำหรับ public access (GitHub Pages ไม่มี auth):
-- CREATE POLICY "allow all" ON raw_materials FOR ALL USING (true);
-- CREATE POLICY "allow all" ON stock_entries FOR ALL USING (true);
-- CREATE POLICY "allow all" ON product_groups FOR ALL USING (true);

-- 7. Insert product groups (จาก sheets)
INSERT INTO product_groups (name, sort_order) VALUES
  ('Cola #521692T', 1),
  ('Cola #050014 1515T', 2),
  ('Cola #1515T_22.50 kg. (1u)', 3),
  ('เกลือบริสุทธิ์ 99.9%', 4),
  ('Co2 SUSTAIN 2501', 5),
  ('Cotton Candy', 6),
  ('Acesulfame& Dry cola5.3', 7),
  ('Cola Sucralose & Dry Cola7.5', 8),
  ('Masking Flavor', 9),
  ('BIOLIGO', 10),
  ('Sugar free Coal', 11),
  ('Strawberry', 12),
  ('SUGAR Free play', 13),
  ('Dry Sarsi', 14),
  ('Citrus Cloud', 15),
  ('Lychee flavor(CSD)', 16),
  ('Grape flavor (CSD) 5.5', 17),
  ('OISHI', 18),
  ('100Plus', 19),
  ('Coconut', 20),
  ('PB 5.4', 21),
  ('Orange 5.4', 22),
  ('Green Soda5.4', 23),
  ('Lemon Lime 5.4', 24),
  ('Tastegam SW Flavor 5.4', 25),
  ('Kamikaze Lime 5.4', 26),
  ('BIB', 27),
  ('NPD สูตรใหม่', 28),
  ('Salty Lemonade_NPD', 29)
ON CONFLICT (name) DO NOTHING;

-- ============================================================
-- FUNCTION: Updated_at trigger
-- ============================================================
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN
  NEW.updated_at = NOW();
  RETURN NEW;
END;
$$ LANGUAGE plpgsql;

CREATE TRIGGER trg_raw_materials_updated
  BEFORE UPDATE ON raw_materials
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

CREATE TRIGGER trg_stock_entries_updated
  BEFORE UPDATE ON stock_entries
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- 8. ตารางผู้ใช้งานระบบ (Users Table)
CREATE TABLE IF NOT EXISTS users (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  email TEXT NOT NULL UNIQUE,
  name TEXT,
  role TEXT DEFAULT 'operator',                 -- บทบาทในอนาคต เช่น admin, operator, viewer
  password TEXT,                                -- รหัสผ่านสำหรับ admin/operator (plain text)
  send_email BOOLEAN DEFAULT TRUE,             -- ติ๊กว่าจะส่ง Email หาคนนี้หรือไม่
  email_type TEXT DEFAULT 'to' CHECK (email_type IN ('to', 'cc')), -- ส่งแบบ To หรือ CC
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- เพิ่ม password column ให้ฐานข้อมูลที่มีอยู่แล้ว (migration)
ALTER TABLE users ADD COLUMN IF NOT EXISTS password TEXT;

-- เพิ่มตัวอย่างข้อมูลผู้ใช้เริ่มต้น
INSERT INTO users (email, name, role, send_email, email_type) VALUES 
  ('manager@company.com', 'ผู้จัดการฝ่ายผลิต', 'admin', true, 'to'),
  ('supervisor@company.com', 'หัวหน้างานคลัง', 'operator', true, 'cc')
ON CONFLICT (email) DO NOTHING;



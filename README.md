# 🧪 StockFlow — ระบบติดตาม Stock หัวเชื้อ

Web App สำหรับบันทึกและ Monitor Stock วัตถุดิบ/หัวเชื้อ พร้อมติดตามวันหมดอายุ  
ใช้ **Supabase** เป็น Backend และรันได้บน **XAMPP** หรือ **GitHub Pages**

---

## 📁 โครงสร้างไฟล์

```
/
├── index.html              ← Dashboard (หน้าหลัก)
├── dashboard.html          ← Dashboard (alias)
├── entry.html              ← บันทึกรับเข้าวัตถุดิบ
├── expiry.html             ← ตารางวันหมดอายุ
├── stock.html              ← ดู Stock ทั้งหมด + แก้ไข
├── master.html             ← Master วัตถุดิบ
├── report.html             ← Aging Report
├── css/
│   └── theme.css           ← Design System (CSS Variables + Components)
├── js/
│   └── script.js           ← Logic ทั้งหมด (Shared across pages)
├── supabase_schema.sql     ← SQL สำหรับสร้าง Database + View
├── import_master.py        ← Script import ข้อมูลจาก Excel
├── requirements/
│   ├── requirements.md     ← System Requirements Specification
│   └── skill.md            ← Technical Stack & Patterns
└── README.md
```

---

## 🚀 ขั้นตอนการตั้งค่า

### 1. สร้าง Database ใน Supabase

1. ไปที่ [Supabase Dashboard](https://app.supabase.com)
2. เข้าโปรเจกต์ของคุณ
3. ไปที่ **SQL Editor**
4. Copy เนื้อหาใน `supabase_schema.sql` แล้ว Run

### 2. ตั้งค่า RLS Policy (สำคัญ!)

ใน Supabase SQL Editor รัน:

```sql
-- อนุญาต public access (สำหรับ XAMPP / GitHub Pages ที่ไม่มี Auth)
ALTER TABLE raw_materials ENABLE ROW LEVEL SECURITY;
ALTER TABLE stock_entries ENABLE ROW LEVEL SECURITY;

CREATE POLICY "public read" ON raw_materials FOR SELECT USING (true);
CREATE POLICY "public write" ON raw_materials FOR ALL USING (true);

CREATE POLICY "public read" ON stock_entries FOR SELECT USING (true);
CREATE POLICY "public write" ON stock_entries FOR ALL USING (true);
```

### 3. ตั้งค่า Supabase Credentials

แก้ไขใน `js/script.js` บรรทัดต้นไฟล์:
```js
const SUPABASE_URL = "https://your-project.supabase.co";
const SUPABASE_KEY = "your-publishable-key";
```

### 4. Import ข้อมูลจาก Excel

```bash
pip install pandas requests openpyxl
python import_master.py
```

> ต้องรันในโฟลเดอร์เดียวกับไฟล์ Excel ทั้งสอง

### 5. เปิดใช้งาน

**บน XAMPP**: วางโฟลเดอร์ใน `c:\xampp\htdocs\` แล้วเปิด `http://localhost/rm-monitor/`

**บน GitHub Pages**:
1. สร้าง Repository ใน GitHub
2. Upload ไฟล์ทั้งหมด
3. ไปที่ Settings → Pages → เลือก Branch: main
4. เข้าใช้งานที่ `https://[username].github.io/[repo-name]`

---

## 🗄️ โครงสร้าง Database

### `raw_materials`
| Column | Type | คำอธิบาย |
|--------|------|----------|
| id | UUID | Primary Key |
| code | TEXT | รหัส SAP |
| product_name | TEXT | สูตร/ผลิตภัณฑ์ที่ใช้วัตถุดิบนี้ |
| name | TEXT | ชื่อวัตถุดิบในสูตร |
| recipe_step | TEXT | ขั้นตอนในสูตร (Step / RM_Name) |
| unit_package | TEXT | รายละเอียด package |
| unit | TEXT | หน่วย (KG, UN, ขวด) |
| unit_per_pack | NUMERIC | จำนวน Unit ที่ผลิตได้ต่อ Pack |
| supplier | TEXT | Supplier |
| remark | TEXT | หมายเหตุ |

### `stock_entries`
| Column | Type | คำอธิบาย |
|--------|------|----------|
| id | UUID | Primary Key |
| product_group | TEXT | กลุ่มวัตถุดิบ (Auto-mapped จาก Master) |
| code | TEXT | รหัส SAP |
| batch | TEXT | Batch Number |
| quantity | NUMERIC | จำนวน |
| received_date | DATE | วันที่รับสินค้า |
| production_date | DATE | วันที่ผลิต |
| expiry_date | DATE | วันหมดอายุ (จากผู้ผลิต) |
| extended_expiry_1 | DATE | ต่ออายุครั้งที่ 1 |
| extended_expiry_2 | DATE | ต่ออายุครั้งที่ 2 |
| shelf_life | TEXT | อายุสินค้า (ตั้งค่าอัตโนมัติ) |
| check_date | DATE | วันที่ตรวจเช็ค (ตั้งค่าอัตโนมัติ) |
| notes | TEXT | หมายเหตุ |

---

## 🎯 ฟีเจอร์

| หน้า | ฟีเจอร์ |
|------|---------|
| **Dashboard** | KPI cards, รายการเร่งด่วน (expired/≤30วัน/≤60วัน) |
| **วันหมดอายุ** | ตารางทั้งหมด + filter + highlight สี + Export CSV |
| **บันทึก Stock** | Form กรอก stock ใหม่ + รายการล่าสุด + Shortcut Date Buttons |
| **Stock ทั้งหมด** | ดูทุก batch + filter + pagination + แก้ไขผ่าน Modal |
| **Master วัตถุดิบ** | ดูข้อมูล raw_materials ทั้งหมด |
| **Aging Report** | วิเคราะห์อายุ stock + คำนวณ Unit ที่ผลิตได้ |

---

## 🎨 สีสถานะ

| สี | ความหมาย |
|----|----------|
| 🔴 แดง | หมดอายุแล้ว (< 0 วัน) |
| 🟠 ส้ม | เหลือ ≤ 30 วัน |
| 🟡 เหลือง | เหลือ 31–60 วัน |
| 🟢 เขียว | ปกติ (> 60 วัน) |

---

## 🔧 Libraries ที่ใช้

| Library | Version | วัตถุประสงค์ |
|---------|---------|------------|
| [@supabase/supabase-js](https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2) | v2 | Database Client |
| [Flatpickr](https://cdn.jsdelivr.net/npm/flatpickr) | latest | Date Picker (dd/mm/yyyy) |
| [IBM Plex Sans Thai](https://fonts.google.com/specimen/IBM+Plex+Sans+Thai) | — | ฟอนต์หลัก |
| [IBM Plex Mono](https://fonts.google.com/specimen/IBM+Plex+Mono) | — | ฟอนต์รหัส/ตัวเลข |

# ระบบประกาศผลการเรียนกลางภาค

ค้นหาผลการเรียนด้วยเลขบัตรประชาชน 13 หลัก แสดงผลเป็นตาราง รองรับมือถือ

## สถาปัตยกรรม

```
GitHub Pages (docs/index.html)  →  Cloudflare Worker (worker/)  →  D1 (midterm-results-db)
```

- ค้นหาฝั่ง server เท่านั้น — browser ได้รับเฉพาะข้อมูลของเลขบัตรที่ค้น ไม่มีทางโหลดข้อมูลทั้งหมด
- ฐานข้อมูล D1: `midterm-results-db` (id: `39cd815e-64a2-4c51-b773-0f69dc34580d`) — สร้างและใส่ schema แล้ว

## ขั้นตอน deploy (ทำครั้งแรก)

```powershell
cd worker
npx wrangler login          # เปิด browser ให้ล็อกอิน Cloudflare
npx wrangler deploy         # จะได้ URL เช่น https://midterm-results-api.xxx.workers.dev
```

จากนั้นแก้ `docs/index.html` บรรทัด `API_BASE` ให้เป็น URL Worker จริง แล้ว commit + push

แนะนำ: หลังได้ URL GitHub Pages แล้ว แก้ `worker/wrangler.toml` ให้
`ALLOWED_ORIGIN = "https://entclick88.github.io"` แล้ว deploy ซ้ำ เพื่อจำกัดไม่ให้เว็บอื่นเรียก API

## นำเข้าข้อมูลคะแนนจากไฟล์ "นำส่งข้อมูลเข้าเว็บ" (แนะนำ)

ใช้กับไฟล์ Excel โครงสร้างของงานวัดผลโดยตรง (ชีต M1–M6, วิชาละ 4 คอลัมน์:
เต็ม/ได้/งานค้าง/%เข้าเรียน) — ต้องมีคอลัมน์ **เลขประจำตัวประชาชน** กรอกครบทุกคน

```powershell
cd import
python convert_xlsx.py "ไฟล์นำส่งข้อมูล.xlsx"   # ตรวจสอบ + สร้าง import.sql
cd ../worker
npx wrangler d1 execute midterm-results-db --remote --file=../import/import.sql
```

ถ้าเลขบัตรยังไม่ครบ/ซ้ำ/ผิด สคริปต์จะรายงานรายคนและไม่สร้าง SQL ให้

## นำเข้าจาก CSV (ทางเลือก)

1. จัดข้อมูลใน Excel ตามคอลัมน์ใน `import/template.csv` (1 แถว = นักเรียน 1 คน × 1 วิชา)
2. Save As → **CSV UTF-8 (Comma delimited)** เช่น `grades.csv` ไว้ในโฟลเดอร์ `import/`
3. รัน:

```powershell
cd import
node import.js grades.csv                 # ได้ไฟล์ import.sql + สรุปจำนวน
cd ../worker
npx wrangler d1 execute midterm-results-db --remote --file=../import/import.sql
```

หมายเหตุ: import จะ**ล้างข้อมูลเดิมทั้งหมด**แล้วใส่ใหม่ (เหมาะกับการอัปเดตคะแนนทั้งชุด)

**ห้าม** commit ไฟล์ csv/sql ที่มีข้อมูลนักเรียนจริง (มี .gitignore กันไว้แล้ว)

## เปิด/ปิดการประกาศผล

```powershell
cd worker
# ปิดชั่วคราว
npx wrangler d1 execute midterm-results-db --remote --command="UPDATE settings SET value='0' WHERE key='announce_open'"
# เปิด
npx wrangler d1 execute midterm-results-db --remote --command="UPDATE settings SET value='1' WHERE key='announce_open'"
# เปลี่ยนชื่อภาคเรียนที่แสดง
npx wrangler d1 execute midterm-results-db --remote --command="UPDATE settings SET value='ผลการเรียนกลางภาค ภาคเรียนที่ 2 ปีการศึกษา 2569' WHERE key='term_label'"
```

## ทดสอบ local

```powershell
cd worker
npx wrangler d1 execute midterm-results-db --local --file=schema.sql --persist-to .wrangler/state
npx wrangler dev --port 8787 --persist-to .wrangler/state
# อีกหน้าต่าง: เปิด docs/index.html ผ่าน http server (หน้าเว็บจะเรียก localhost:8787 อัตโนมัติ)
```

เลขบัตรทดสอบ: `1234567890121` (สมชาย), `1100700256030` (สมหญิง)

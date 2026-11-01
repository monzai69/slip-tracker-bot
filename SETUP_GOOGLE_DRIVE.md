# เชื่อม Slip Tracker กับ Google Drive (ทำครั้งเดียว ~10 นาที)

ระบบจะอ่าน statement ธนาคาร (CSV) และไฟล์ payroll จากโฟลเดอร์ Drive ที่ส่งให้บริษัทบัญชีอยู่แล้ว ไม่ต้องอัปโหลดซ้ำ
สิทธิ์ที่ได้คือ **อ่านอย่างเดียว** และ **เห็นเฉพาะโฟลเดอร์ที่แชร์ให้** เท่านั้น

## 1. สร้าง service account (Google Cloud — ฟรี)
1. เปิด https://console.cloud.google.com → มุมบนซ้าย เลือก/สร้าง project ใหม่ ชื่ออะไรก็ได้ เช่น `slip-tracker`
2. เมนู **APIs & Services → Library** → ค้นหา **Google Drive API** → กด **Enable**
3. เมนู **IAM & Admin → Service Accounts** → **Create service account**
   - ชื่อ เช่น `slip-tracker-drive` → **Create and continue**
   - ขั้น "Grant this service account access" **ข้ามได้** (ไม่ต้องให้ role) → **Done**
4. คลิก service account ที่เพิ่งสร้าง → แท็บ **Keys** → **Add key → Create new key → JSON** → ไฟล์ `.json` จะถูกดาวน์โหลด

> ⚠️ ไฟล์ JSON นี้คือกุญแจ ห้ามส่งในแชท ห้าม commit เข้า GitHub ห้ามใส่ใน Drive

## 2. ใส่กุญแจใน Railway
1. Railway → project → service **slip-tracker-bot** → **Variables** → **New Variable**
2. ชื่อ `GOOGLE_SERVICE_ACCOUNT_JSON` · ค่า = **เปิดไฟล์ JSON ด้วย Notepad/TextEdit แล้วคัดลอกทั้งไฟล์** มาวาง
3. Save → Railway จะ redeploy เอง (~1 นาที)

## 3. แชร์โฟลเดอร์ให้ระบบ
1. เปิด web app → แท็บ **Report** → การ์ด **🏦 Bank statements** → จะเห็นอีเมล เช่น `slip-tracker-drive@slip-tracker.iam.gserviceaccount.com` (กดเลือกคัดลอกได้)
2. Google Drive → คลิกขวาโฟลเดอร์ statement → **Share** → วางอีเมลนั้น → สิทธิ์ **Viewer** → เอาติ๊ก "Notify people" ออก → **Share**
   - ใช้โฟลเดอร์เดิมที่แชร์ให้บริษัทบัญชีได้เลย ระบบอ่านโฟลเดอร์ย่อยด้วย
   - ถ้าโฟลเดอร์มีเอกสารอื่นที่ไม่อยากให้ระบบเข้าถึง ให้แชร์เฉพาะโฟลเดอร์ย่อย เช่น `Bank statements`
   - ถ้าโฟลเดอร์นี้บริษัทบัญชีเป็นเจ้าของ อาจต้องขอให้เขาแชร์ให้ (หรือให้สิทธิ์ Editor กับคุณก่อน)
3. คัดลอกลิงก์โฟลเดอร์ (Share → Copy link) → วางในช่อง **"Add a Drive folder link"** → **＋ Add**
   - ใส่ได้หลายโฟลเดอร์ เช่น โฟลเดอร์รายได้/petty cash และโฟลเดอร์ `statement`
   - ในโฟลเดอร์ `statement/2026/1026/` ชื่อโฟลเดอร์ **MMYY** (1026 = ต.ค. 2026) บอกเดือนของไฟล์ข้างใน
4. กด **🔄 Get new files from Drive** — หลังจากนี้ระบบเช็กเองทุก 6 ชั่วโมง

## ระบบอ่านไฟล์อะไรบ้าง
| ไฟล์ | อ่านได้ |
|---|---|
| KBank statement CSV (K-DEPOSIT STATEMENT) | ✅ |
| SCB statement CSV (HISTSTMT…) | ✅ |
| Payroll CSV จากระบบเงินเดือน (มีคอลัมน์ Branch, NET PAY; ชื่อไฟล์มีเดือน-ปี เช่น `RCG_Payroll_October_2026.csv`) | ✅ |
| รีพอตขาย `.xlsx` (ชีต "รายการขาย") | ✅ |
| รีพอต Petty cash `.xlsx` ทีละสาขา (ย้อน 60 วัน ซ้ำกันได้ ระบบตัดด้วย ID) | ✅ |
| PDF ของ Krungsri / GSB | ⏭ อัปเดตถัดไป (PR B2) |
| ไฟล์อื่น (Excel report, PDF เงินเดือนรายคน ฯลฯ) | ข้ามไปเอง |

ทุกไฟล์ statement ถูก **ตรวจยอดคงเหลือทุกบรรทัด** ก่อนรับเข้า ถ้าตัวเลขไม่ต่อกันแม้ 1 สตางค์ ระบบจะไม่รับและบอกว่าบรรทัดไหน

## ยกเลิกการเชื่อม
- ถอนแชร์โฟลเดอร์ใน Drive → ระบบเข้าไม่ได้ทันที
- หรือ Google Cloud → Service Accounts → Keys → ลบ key
- ไม่มี Drive ก็ยังใช้ปุ่ม **📤 Upload statement / payroll CSV** ได้ตามปกติ

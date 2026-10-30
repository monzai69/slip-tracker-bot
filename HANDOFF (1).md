# RCG Payroll — Handoff สำหรับย้ายไปทำงานใน Claude Code

อัปเดต: 30 ต.ค. 2026 · เจ้าของระบบ: Mon (HR/Owner ของ ReadyCheckGo)
เอกสารนี้สรุปสิ่งที่ทำไปแล้ว สิ่งที่ค้างอยู่ และบทเรียนทางเทคนิค เพื่อให้ session ใหม่ทำต่อได้โดยไม่ต้องถามซ้ำ

---

## 0. เช็กก่อนเริ่ม (สำคัญที่สุด)

ไฟล์ล่าสุดที่ส่งให้ในแชทคือ `index.html` (admin web app) และ `server.js` (LINE server) ถ้ายังไม่ได้ commit ขึ้น GitHub ให้ทำก่อน ไม่งั้น Claude Code จะทำงานบนโค้ดเก่า ซึ่งยังมีบั๊กที่แก้ไปแล้ว (เช่น published flag หาย, payslip PDF ผิด)

- [ ] `rcg-payroll/index.html` ใน GitHub เป็นเวอร์ชันล่าสุด (มี `publish_meta`, conflict guard, Reports page)
- [ ] `rcg-payroll-line/server.js` ใน GitHub เป็นเวอร์ชันล่าสุด (มี `/api/unpublish-salary`, cron ที่อ่าน `publish_meta`)
- [ ] Railway มี env var `ADMIN_LINE_ID` (LINE user ID ของ Mon) ไม่งั้นจะไม่ได้แจ้งเตือนตอนพนักงานกดยืนยันเงินเดือน

---

## 1. โปรเจกต์และการ deploy

| Repo | คืออะไร | Deploy |
|---|---|---|
| `rcg-payroll` | Admin web app ไฟล์เดียว `index.html` (เกือบ 4,000 บรรทัด) | GitHub Pages: `monzai69.github.io/rcg-payroll` |
| `rcg-payroll-line` | Node.js server: LINE bot, LIFF API, สร้าง PDF (puppeteer), cron | Railway: `rcg-payroll-line-production.up.railway.app` |
| `slip-tracker-bot` | โปรเจกต์ที่สาม: LINE bot + web app ติดตามสลิปโอนเงิน/บิล + รายงาน Excel รายเดือน — **มี `CLAUDE.md` ของตัวเองใน repo เป็นแหล่งอ้างอิงหลัก อ่านก่อนแตะ** (สรุปย่อในข้อ 8) | Railway: `slip-tracker-bot-production.up.railway.app` |

- `rcg-payroll-line` มี `index.html` ด้วย แต่เป็น **LIFF Staff Portal** (หน้าให้พนักงานดูเงินเดือนใน LINE) คนละไฟล์กับ admin web app อย่าสับสน
- Dependencies ของ server: `@line/bot-sdk`, `express`, `firebase-admin`, `cors`, `dotenv`, `puppeteer`, `multer`, `node-cron`
- ⚠️ ถ้าจะอัปเวอร์ชัน `@line/bot-sdk`: v10 ขึ้นไป **ตัด `line.Client` (legacy) ทิ้ง** — slip-tracker-bot เคยพังเพราะเรื่องนี้จึง pin `^9.0.0` ไว้ ตรวจว่า server.js ของ payroll ใช้ API แบบไหนก่อนอัป
- Web app โหลด SheetJS (xlsx) และ Chart.js จาก CDN
- Database: Firebase Firestore ต้องอยู่แผน pay-as-you-go (เคยโดน "billing account closed" 403 ตอน free trial หมด) แนะนำตั้ง budget alert ราว $10–15/เดือน

---

## 2. กฎการทำงานของ Mon (ใส่ใน CLAUDE.md ด้วย)

1. **เสนอแผนและรอ confirm ก่อนแก้โค้ดทุกครั้ง** ไม่ลงมือเองก่อน
2. ทำงานจากไฟล์ล่าสุดใน repo เท่านั้น ไม่เดาจากความจำ
3. สถาปัตยกรรมไฟล์เดียวของ admin web app เป็นความตั้งใจ ห้ามแตกเป็นหลายไฟล์ถ้า Mon ไม่ได้สั่ง
4. บั๊กที่เกี่ยวกับข้อมูลจริง: วินิจฉัยด้วย console snippet (อ่านอย่างเดียว) ให้เห็นหลักฐานก่อน แล้วค่อยแก้
5. ทดสอบ logic ด้วยตัวเลขจริงจาก screenshot ของ Mon ก่อนส่งงาน
6. Mon ใช้ทั้งไทยและอังกฤษ ตอบไทยเมื่อเธอเขียนไทย

---

## 3. สถาปัตยกรรมที่ต้องรู้

**ข้อมูลเงินเดือน** อยู่ใน Firestore `config/payroll` เป็น JSON ก้อนเดียว key รูปแบบ `YYYY_M` (เช่น `2026_10`) ข้างในเป็นรายพนักงาน (`entry`) ฟิลด์สำคัญของ entry:
- input: `rev`, `bonus`, `otHrs`, `leaveDays`, `leaveHrs`, `dOth`, `periodic{itemId: ค่า}`, `diligence{}`
- `snap`: snapshot ของ base/rates/contract/items ที่ถูกแช่ไว้ตอน Save All
- `net`, `payRows`: ผลคำนวณที่ cache ไว้ (`payRows` คือรายการแสดงผลเต็ม ใช้โดย slip PDF และหน้า LIFF)

**calcM(s, e, key)** ใน index.html คือที่เดียวที่คำนวณเงินเดือน (ภาษี ภงด.1, SSO, WHT, commission, leave deduction) ห้ามเขียนสูตรซ้ำใน server ให้ใช้ `payRows` ที่ web app cache ไว้

**Snapshot**: ถ้ามี `e.snap` calcM จะใช้ค่าจาก snapshot ไม่ใช้โปรไฟล์พนักงานปัจจุบัน Save All สร้าง snapshot ใหม่จากโปรไฟล์ปัจจุบันทุกครั้ง Save & Copy ก็สร้างใหม่ทุกครั้งแล้ว

**Publish / confirmation** อยู่ใน `config/publish_meta` (เอกสารแยก **ฝั่ง server เขียนอย่างเดียว** web app อ่านอย่างเดียว):
```
{ "2026_10": { published, publishedAt, unpublishedAt?, reminderAt, reminderSent,
               autoConfirmAt, autoConfirmDone,
               confirmations: { staffId: { confirmed, confirmedAt, autoConfirmed? } } } }
```
ระบบเก่าเก็บ flag ที่ `payroll[key]._meta` ยังมี fallback อ่านได้และ migrate เข้าเอกสารใหม่เอง

**Server endpoints**: `GET /api/salary/:staffId/:year/:month` (LIFF, ตรวจ published), `POST /api/publish-salary` (เขียน flag ก่อนส่ง LINE), `POST /api/unpublish-salary`, `POST /api/confirm-salary`, cron ทุก 10 นาที (timezone Asia/Bangkok) ส่งเตือน 2 ชม. และ auto-confirm 5 ชม. เฉพาะพนักงาน Active ที่ลงทะเบียน LINE

**ล็อกงวด**: งวดล็อกอัตโนมัติวันที่ 2 ของเดือนถัดไป ปลดล็อกด้วยปุ่มใน banner (`P[key]._meta.unlocked`) แล้วล็อกกลับเอง

**Navigation**: index.html มี banner คั่นส่วนด้วย `══════` (Dashboard, Payroll, Reports, Staff, Leave ฯลฯ) ค้นหาด้วยข้อความนี้ได้

---

## 4. งานที่ทำเสร็จแล้ว (ส่งมอบแล้ว ต้อง deploy ถ้ายังไม่ได้ทำ)

**Web app**
- แก้บั๊ก 6-month incentive (ปุ่ม Apply แล้วค่า Bonus หายตอน Save)
- Payment History: ปุ่ม Export เป็น Excel เลือกช่วงเดือนได้ คอลัมน์ item ตามชื่อจริงของแต่ละคน
- หน้า **Reports** ใหม่ (Chart.js): salary trend แยก Base / Other fixed / Commission / ค่ามือ, อันดับ Revenue และ ค่ามือ เลือกช่วงเวลาได้, Labor cost % of revenue, ต้นทุนเงินเดือนแยกสาขา
- นำเข้า DF Report (`.xlsx`/`.csv`) รูปแบบใหม่: ชื่อ tab `DF*` → ค่ามือ, `COMM*` (COMM_SALE + COMM_DEPOSIT) → revenue รวมกัน, ข้าม tab `All`, อ่านจากบล็อก "สรุปรายคน", จับ header แบบ prefix (ชื่อ column เคยเปลี่ยน ค่าคอม ↔ ค่ามือ)
- ปลายทางของ ค่ามือ: item periodic/comm_flat_custom ชื่อ "ค่ามือ" ก่อน → ถ้ามี item commission แบบ % ข้ามไป → ไม่งั้นลง Bonus (แบบ replace ไม่ใช่ add ผ่าน `_importedKamue`)
- item type ใหม่ `comm_flat_custom` (Commission % of Custom Value) มีกล่องกรอกฐานคิดของตัวเองบนการ์ดเงินเดือน
- แก้ item ID ไม่ให้เปลี่ยนทุกครั้งที่ Save staff (`data-item-id`) และให้ calcM แก้ ID ใน snapshot เก่าที่ค้างอยู่
- ปุ่มปลดล็อก/ล็อกงวด
- ปุ่ม Publish เป็น toggle (publish / unpublish), badge สถานะ
- Save ลง cloud แบบ debounce 1.2 วินาที (`svP`), Save All/Save & Copy เขียนทันที (`flushP`), **conflict guard**: ปฏิเสธการเซฟถ้ามีแท็บ/อุปกรณ์อื่นเซฟไปแล้วหลังโหลดหน้านี้ (`window._payStamp`)

**Server**
- `buildPaySlipHTML` ใช้ `payRows` (แก้ payslip PDF ที่แสดง Total Income = 0)
- ย้าย published flag และ confirmations ไป `publish_meta`, server เลิกเขียน `config/payroll`
- Cron เตือน/auto-confirm ทำงานบน `publish_meta`
- เขียน flag ก่อนส่ง LINE (พนักงานกดลิงก์ทันทีไม่เจอ "ยังไม่เปิด")
- ประกาศ `ADMIN_LINE_ID` จาก env (เดิมไม่เคยถูกประกาศ จึงไม่เคยแจ้งแอดมินตอนมีคนยืนยัน)

---

## 5. งานค้าง

### 5.1 Dashboard แบ่งต้นทุนสำนักงานใหญ่ (ออกแบบแล้ว ยังไม่ build) ← งานถัดไป
RCG main (สำนักงานใหญ่ ฝ่ายหลังบ้าน/marketing) ดูแลทุกสาขา ต้นทุนเงินเดือนของมันต้องหารไปให้สาขาปฏิบัติการ
- Branch Settings: checkbox "Back-office / shared-cost branch" (ติ๊กที่ RCG main) ไม่ hardcode ชื่อสาขา
- โหมดแบ่ง: **Equal** (ค่าเริ่มต้น หารด้วยจำนวนสาขาปฏิบัติการ ปรับเองเมื่อเปิดสาขา 3, 4, 5) หรือ **Custom %** ต่อสาขา (เตือนถ้ารวมไม่เท่า 100%)
- Dashboard การ์ดสาขาปฏิบัติการ: ยอด net ของสาขาเอง + ส่วนแบ่งต้นทุนหลังบ้าน = **ยอดที่ต้องโอนของสาขานั้น** (ตัวหนา) เช่น Samui 400,000 + (RCG 120,000 ÷ 2) = 460,000
- การ์ด RCG main แสดงยอดพร้อมหมายเหตุ "แบ่งให้ N สาขา"
- Dashboard ใช้ **net pay**, หน้า Reports ใช้ gross (แยกกันตามนี้)
- ถ้าไม่ติ๊กสาขาไหนเป็น back-office หน้าตาเหมือนเดิม
- ขอบเขต: web app อย่างเดียว ไม่แตะการคำนวณเงินเดือน ไม่แตะ server/LIFF

### 5.2 Guaranteed minimum income (เสนอแล้ว Mon ยังไม่ confirm)
สัญญาบางคนรับประกันรายได้ขั้นต่ำช่วง 3 เดือนแรก (ตัวอย่าง 42,500 ฿ = base package + commission + ค่ามือ **ไม่รวม OT**) ตอนนี้ทำมือโดยกรอกส่วนต่างในช่อง Bonus ข้อเสนอ: ฟิลด์ต่อคน "ประกันรายได้ขั้นต่ำ X เดือนแรก" คำนวณ top-up อัตโนมัติและแสดงเป็นบรรทัดแยก WHT/ภาษีคิดบนยอดรวมหลัง top-up

### 5.3 Pro-rata พนักงานเข้าใหม่/ลาออก (ยังไม่มีฟีเจอร์)
ตอนนี้ใช้ทางเลี่ยงด้วยช่อง Leave Days/Hours (หารด้วยจำนวนวันทำงานของแผนก ค่าเริ่มต้น 26) ข้อเสนอ: ช่อง "วันทำงานในงวดนี้" ที่ pro-rate base + fixed items ด้วย วัน ÷ ตัวหาร และแสดง "Pro-rata 6/30 วัน" ในสลิป
- **ยังต้องให้ Mon ตัดสิน**: ตัวหารเป็น 30 (แบบมาตรฐานเงินเดือนรายเดือน) หรือจำนวนวันทำงานแผนก (26) ซึ่งตรงกับการหักลาของพนักงานคนอื่น
- ระวัง: พนักงาน permanent ต้องใช้เงินเดือนเต็มเดือนในการประมาณภาษีรายปี ไม่ใช้ยอด pro-rate

### 5.4 อื่น ๆ
- การ์ดเงินเดือนเตือนเมื่อโปรไฟล์พนักงานต่างจาก snapshot ที่ Save ไว้ ("ตั้งค่าพนักงานเปลี่ยนหลังบันทึก กด Save All เพื่อใช้")
- Leave system Phase 2 (สิทธิ์ลาตามอายุงาน, ลาคลอด/เกณฑ์ทหาร/งานศพ/แต่งงาน/ภัยพิบัติ)

---

## 6. บทเรียนทางเทคนิค (กันบั๊กซ้ำ)

1. **ห้ามให้สองฝั่งเขียนเอกสารเดียวกัน** เหตุการณ์ 29 ต.ค. 2026: แท็บสองอันของ admin เขียนทับกัน ทำให้ published flag และตัวเลขที่แก้หาย แก้โดยแยก `publish_meta` ให้ server เขียนฝั่งเดียว และมี conflict guard ฝั่ง web app
2. **Item ID ต้องคงที่** ค่าในกล่อง periodic/comm_flat_custom ผูกกับ `item.id` ถ้า ID เปลี่ยน ข้อมูลเก่าจะ "หาย" (ยังอยู่แต่หาไม่เจอ)
3. **Snapshot บัง profile**: แก้เงินเดือนหรือสัญญาในหน้า Staff หลัง Save All แล้ว การ์ดไม่เปลี่ยนจนกว่าจะ Save All ใหม่ (หรือ Save & Copy)
4. `gatherE()` (เรียกทุกครั้งที่แก้ช่องในการ์ด) สร้าง entry ใหม่จาก input เท่านั้น จึงทิ้ง `payRows` เก่าไป ต้องกด Save & Copy/Save All เพื่อให้ LIFF และ PDF ได้รายการเต็ม
5. การ์ดแสดงแถวหัก (SSO/ภาษี/WHT) ตาม `s.ct` ของโปรไฟล์ปัจจุบัน แต่ calcM คำนวณตาม `snap.ct` ถ้าสองค่าต่างกัน จะเห็นเลขไม่ตรงกับแถวที่แสดง (เช่น "SSO −0" ทั้งที่ถูกหัก WHT ไป)
6. Labor cost % ใน Reports เป็นค่าประมาณ: gross รวม ÷ revenue ที่กรอกรายคนเท่านั้น ไม่ใช่รายได้คลินิกทั้งหมด
7. ไม่มี test อัตโนมัติใน repo ที่ผ่านมาทดสอบด้วย script node แยกพร้อม mock data และตัวเลขจริงจากหน้าจอ

---

## 7. สิ่งที่ Mon ต้องทำเอง (Claude Code ทำแทนไม่ได้)
- Deploy: merge เข้า branch ที่ GitHub Pages / Railway ดึงไป และตรวจว่า deploy ผ่าน
- ตั้ง env var บน Railway (`ADMIN_LINE_ID`) และดูแล Firebase billing
- ตรวจข้อมูลจริงใน browser (console) เมื่อมีปัญหาที่ผูกกับข้อมูลใน Firestore
- แจ้งพนักงานเองเมื่อแก้เงินเดือนหลัง publish (ระบบยังไม่ส่งข้อความแจ้งเมื่อตัวเลขเปลี่ยน และสถานะ "ยืนยันแล้ว" เดิมยังอยู่)

---

## 8. โปรเจกต์ที่สาม: slip-tracker-bot (อ้างอิงย่อ)

ระบบติดตามสลิปโอนเงินของกลุ่มคลินิก (~70–120 รายการ/เดือน) จับคู่สลิปกับบิล/ใบแจ้งหนี้ และออก Excel รายเดือนแยกบัญชีธนาคารให้บัญชี **รายละเอียดเต็ม กฎ และ data model อยู่ใน `CLAUDE.md` ที่ root ของ repo นั้น — ใช้ไฟล์นั้นเป็นหลัก ไม่ใช่เอกสารนี้**

| หัวข้อ | ค่า |
|---|---|
| Repo | `monzai69/slip-tracker-bot` (ควรเป็น **Private**) |
| ไฟล์หลัก | `index.js` (ทุกอย่างฝั่ง server), `webapp.html` (web app ที่ `/app`), `CLAUDE.md` |
| Deploy | Railway service `slip-tracker-bot` — push GitHub แล้ว redeploy อัตโนมัติ (~1 นาที) |
| ข้อมูลจริง | **Railway volume 5 GB ที่ `/app/data`** — `payments.json`, `inbox.json`, `categories.json`, รูปสลิป/บิลทั้งหมด |
| Env vars | `LINE_CHANNEL_ACCESS_TOKEN` / `LINE_CHANNEL_SECRET` (**คนละ LINE channel กับ payroll**), `ANTHROPIC_API_KEY`, `DASHBOARD_URL`, `WEB_PASSWORD` |
| เวอร์ชันปัจจุบัน | v7 (ต.ค. 2026): evidence note, reply-to-slip ใน LINE, expense categories + vendor-memory learning, Stats tab |

**กฎพิเศษของโปรเจกต์นี้ (เพิ่มจากข้อ 2 ซึ่งใช้ทุกข้อเช่นกัน):**
- Volume = ข้อมูลจริงหนึ่งเดียว เคยหายมาแล้ว 2 เดือน ห้ามเขียน migration ทับ `payments.json` โดยไม่สร้าง backup ก่อน (มี `payments_backup_v4.json` อยู่แล้วหนึ่งชุด)
- ตอบใน LINE ให้สั้น รายละเอียดเต็มเก็บใน database/web app
- ส่งงานเป็นไฟล์ทั้งไฟล์ (Mon แก้ผ่าน GitHub web UI) และพยายามแตะครั้งละไฟล์เดียว

**บทเรียนที่ใช้ร่วมกันทั้งสาม repo:**
1. `@line/bot-sdk` pin `^9` — v10 ตัด `line.Client` (ดูข้อ 1)
2. Webhook ของ slip bot ใช้ `express.json()` ตรง ๆ **ไม่ใช้ LINE signature middleware** — middleware เคยทำให้ bot เงียบทั้งระบบแบบไม่มี error
3. Push GitHub = Railway redeploy ทันที — ปลอดภัยต่อข้อมูล (มี volume) แต่ข้อความ LINE ที่เข้ามาช่วง ~30 วินาทีระหว่าง redeploy อาจหายไป ควร deploy ช่วงที่ไม่มีคนส่งสลิป
4. `ANTHROPIC_API_KEY` เครดิตหมด = การอ่านสลิป/บิล/จัดหมวดล้มทั้งหมดด้วย error 400 "credit balance is too low" — เปิด auto-reload ใน console.anthropic.com กันเงียบตาย (payroll ไม่ใช้ Anthropic จึงไม่กระทบ)
5. ภาพหน้าจอจากแอปธนาคารบางตัวเป็น PNG — ส่งให้ Anthropic ต้องประกาศ media type ตามไฟล์จริง (slip bot มี `sniffMediaType()` แล้ว ถ้าโปรเจกต์อื่นจะส่งรูปเข้า AI ให้ยืมแนวทางนี้)

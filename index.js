"use strict";

const express = require("express");
const line = require("@line/bot-sdk");
const axios = require("axios");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");
const ExcelJS = require("exceljs");
const { format, parseISO } = require("date-fns");

const app = express();
app.use(express.json({ limit: "30mb" }));
const PORT = process.env.PORT || 8080;

const LINE_CONFIG = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN || "",
  channelSecret: process.env.LINE_CHANNEL_SECRET || "",
};
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || "";
const WEB_PASSWORD = process.env.WEB_PASSWORD || "";
const BASE_URL = process.env.DASHBOARD_URL || "https://slip-tracker-bot-production.up.railway.app";

const DATA_DIR   = process.env.DATA_DIR || path.join(__dirname, "data");
const SLIPS_DIR  = path.join(DATA_DIR, "slips");
const BILLS_DIR  = path.join(DATA_DIR, "bills");
const DATA_FILE  = path.join(DATA_DIR, "payments.json");
const INBOX_FILE = path.join(DATA_DIR, "inbox.json");
const CATS_FILE  = path.join(DATA_DIR, "categories.json");
const ACCTS_FILE    = path.join(DATA_DIR, "accounts.json");
const SETTINGS_FILE = path.join(DATA_DIR, "settings.json");
const SCHED_FILE    = path.join(DATA_DIR, "report_schedule.json");

[DATA_DIR, SLIPS_DIR, BILLS_DIR].forEach(function(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

// One-time safety backups (v4: multi-bill migration, v8: branch split / internal transfer fields, v9: date fix / re-numbering, v10: name memory / duplicate flags)
["payments_backup_v4.json", "payments_backup_v8.json", "payments_backup_v9.json", "payments_backup_v10.json"].forEach(function(bn) {
  var bf = path.join(DATA_DIR, bn);
  try {
    if (fs.existsSync(DATA_FILE) && !fs.existsSync(bf)) fs.copyFileSync(DATA_FILE, bf);
  } catch (e) { console.error("Backup error:", e.message); }
});

// ── Storage ────────────────────────────────────────────────────────────────
function loadPayments() {
  try {
    if (!fs.existsSync(DATA_FILE)) return [];
    return JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  } catch (e) { return []; }
}
function savePayments(list) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(list, null, 2));
}
var DEFAULT_CATS = ["ยา/Medicine","เงินเดือน/Salary","โบนัส/Bonus","อาหารพนักงาน/Staff meal","พยาบาล/Nursing fee","หมอ/Doctor fee","ค่าเช่า/Rent","โฆษณา/Advertising","ขนส่ง/Logistics","อุปกรณ์/Equipment","ค่าน้ำค่าไฟ/Utilities","ภาษี-ประกันสังคม/Tax & SSO","อื่นๆ/Other"];
function loadCategories() {
  try {
    if (!fs.existsSync(CATS_FILE)) { fs.writeFileSync(CATS_FILE, JSON.stringify(DEFAULT_CATS, null, 2)); return DEFAULT_CATS.slice(); }
    var c = JSON.parse(fs.readFileSync(CATS_FILE, "utf8"));
    return Array.isArray(c) ? c : DEFAULT_CATS.slice();
  } catch (e) { return DEFAULT_CATS.slice(); }
}
function saveCategories(list) { fs.writeFileSync(CATS_FILE, JSON.stringify(list, null, 2)); }

// ── Category learning: vendor memory + examples from past corrections ──
function vendorKey(name) { return String(name || "").trim().toLowerCase().replace(/\s+/g, " "); }
function vendorCategoryMap() {
  var map = {};
  loadPayments().forEach(function(p) {
    if (p.category && p.recipient_name) map[vendorKey(p.recipient_name)] = p.category; // later entries (newer) win
  });
  return map;
}
// withNames=false for the image reader: it once copied a past recipient onto a slip that never showed that name
function categoryExamples(limit, withNames) {
  var ex = [];
  var l = loadPayments();
  for (var i = l.length - 1; i >= 0 && ex.length < (limit || 15); i--) {
    var p = l[i];
    if (withNames === false) { if (p.category && p.purpose) ex.push({ purpose: p.purpose, category: p.category }); continue; }
    if (p.category && (p.purpose || p.recipient_name)) ex.push({ purpose: p.purpose || "", recipient: p.recipient_name || "", category: p.category });
  }
  return ex;
}

// ── Name memory: corrections you made to AI-read names, and Biller IDs → the vendor's real name ──
function correctedName(data) {
  var byRead = {}, byBiller = {};
  loadPayments().forEach(function(p) {
    if (p.recipient_read && p.recipient_name && p.recipient_read !== p.recipient_name) byRead[vendorKey(p.recipient_read)] = p.recipient_name;
    if (p.biller_id && p.recipient_name) byBiller[String(p.biller_id).replace(/\D/g, "")] = p.recipient_name;
  });
  var biller = String(data.biller_id || "").replace(/\D/g, "");
  if (biller.length >= 6 && byBiller[biller]) return byBiller[biller];
  return byRead[vendorKey(data.recipient_name)] || null;
}

// ── Possible duplicates: same transaction number, or same paying account + amount + date + recipient ──
function refKey(r) { var k = String(r || "").replace(/[^0-9a-z]/gi, "").toUpperCase(); return k.length >= 6 ? k : null; }
// Same paying account + amount + date always required (two branches may pay one biller the same amount);
// then the transaction number decides if both have one (a biller's customer reference repeats every time), else the recipient
function sameSlip(a, b) {
  if (a.id === b.id || !a.transaction_date || a.transaction_date !== b.transaction_date) return false;
  if (Math.abs((Number(a.amount) || 0) - (Number(b.amount) || 0)) >= 0.01) return false;
  if (normalizeBank(a.bank_from) !== normalizeBank(b.bank_from)) return false;
  if (String(a.account_from || "").replace(/\D/g, "").slice(-4) !== String(b.account_from || "").replace(/\D/g, "").slice(-4)) return false;
  var ra = refKey(a.reference_number), rb = refKey(b.reference_number);
  if (ra && rb) return ra === rb;
  return vendorKey(a.recipient_name) === vendorKey(b.recipient_name);
}
// the newer slip of each pair is the suspect; dup_ok = you said "not a duplicate"
function duplicateOf(p, list) {
  if (p.dup_ok) return null;
  return list.find(function(q) { return Number(q.id) < Number(p.id) && sameSlip(p, q); }) || null;
}
// Recipient paid before at an account one digit away (9408 vs 9498) → probably a misread
function similarToAccount(p, list) {
  var mine = String(p.account_to || "").replace(/\D/g, "").slice(-4);
  if (mine.length < 4 || !p.recipient_name) return null;
  var counts = {};
  list.forEach(function(q) {
    if (q.id === p.id || vendorKey(q.recipient_name) !== vendorKey(p.recipient_name) || normalizeBank(q.bank_to) !== normalizeBank(p.bank_to)) return;
    var l = String(q.account_to || "").replace(/\D/g, "").slice(-4);
    if (l.length === 4 && l !== mine) counts[l] = (counts[l] || 0) + 1;
  });
  var hit = null;
  Object.keys(counts).forEach(function(l) {
    var diff = 0;
    for (var i = 0; i < 4; i++) if (l[i] !== mine[i]) diff++;
    if (diff === 1 && (!hit || counts[l] > counts[hit])) hit = l;
  });
  return hit;
}
// Recipient read as our own / the sender's name while the money did not go to our account → reader got confused
function recipientLooksLikeSender(p, reg) {
  if (!p.recipient_name || p.recipient_ok) return false;
  if (classify(p, reg || loadAccounts()).kind === "internal") return false; // our account → our account: same name is right
  var n = normName(p.recipient_name);
  var sn = normName(p.sender_name);
  return !!(sn.length >= 3 && n.length >= 3 && (n === sn || n.indexOf(sn) !== -1 || sn.indexOf(n) !== -1));
}

function loadInbox() {
  try {
    if (!fs.existsSync(INBOX_FILE)) return [];
    return JSON.parse(fs.readFileSync(INBOX_FILE, "utf8"));
  } catch (e) { return []; }
}
function saveInbox(list) {
  fs.writeFileSync(INBOX_FILE, JSON.stringify(list, null, 2));
}

// ── Display ID: yymmNN (e.g. 261001 = 2026 Oct #1) ────────────────────────
function nextDisplayId(dateStr) {
  var d = dateStr ? new Date(dateStr) : new Date();
  if (isNaN(d.getTime())) d = new Date();
  var yy = String(d.getFullYear()).slice(2);
  var mm = String(d.getMonth() + 1).padStart(2, "0");
  var prefix = yy + mm;
  var max = 0;
  loadPayments().forEach(function(p) {
    if (p.display_id && String(p.display_id).indexOf(prefix) === 0) {
      var n = parseInt(String(p.display_id).slice(4));
      if (!isNaN(n) && n > max) max = n;
    }
  });
  return prefix + String(max + 1).padStart(2, "0");
}
function nextNumericId() {
  var max = 0;
  loadPayments().forEach(function(p) { if (Number(p.id) > max) max = Number(p.id); });
  return max + 1;
}
function showId(p) { return p.display_id ? p.display_id : String(p.id); }

// ── Multi-bill support: billFiles[] is the new store; billFile kept for compatibility ──
function getBillFiles(p) {
  if (Array.isArray(p.billFiles)) return p.billFiles;
  return p.billFile ? [p.billFile] : [];
}
function setBillFiles(p, arr) {
  p.billFiles = arr;
  p.billFile = arr.length ? arr[0] : null;
}
function moveToInboxAsItem(file, type, billAmount, source) {
  var inbox = loadInbox();
  var item = { id: "ib" + Date.now() + Math.floor(Math.random()*1000), file: file, type: type || "bill", billAmount: (billAmount === 0 || billAmount) ? billAmount : null, uploadedAt: new Date().toISOString(), source: source || "web" };
  inbox.push(item);
  saveInbox(inbox);
  return item;
}
async function readInboxAmountsInBackground(ids) {
  for (var i = 0; i < ids.length; i++) {
    try {
      var inbox = loadInbox();
      var idx = inbox.findIndex(function(x) { return x.id === ids[i]; });
      if (idx === -1) continue;
      if (inbox[idx].billAmount !== null && inbox[idx].billAmount !== undefined) continue;
      var fp = path.join(BILLS_DIR, inbox[idx].file);
      if (!fs.existsSync(fp)) continue;
      var amt = await readBillAmount(fs.readFileSync(fp).toString("base64"));
      inbox = loadInbox();
      idx = inbox.findIndex(function(x) { return x.id === ids[i]; });
      if (idx === -1) continue;
      inbox[idx].billAmount = (amt === 0 || amt) ? amt : null;
      saveInbox(inbox);
    } catch (e) { console.error("bg amount error:", e.message); }
  }
}

// ── Last slip per user (for purpose note in LINE) ─────────────────────────
const lastSlip = {};
function getLastSlip(uid) {
  var s = lastSlip[uid];
  if (!s) return null;
  if (Date.now() > s.expiresAt) { delete lastSlip[uid]; return null; }
  return s;
}

// ── Detect the REAL image format (JPEG / PNG / GIF / WEBP) from the file bytes ──
function sniffMediaType(b64) {
  try {
    var h = Buffer.from(String(b64).slice(0, 64), "base64");
    if (h[0] === 0x89 && h[1] === 0x50 && h[2] === 0x4E && h[3] === 0x47) return "image/png";
    if (h[0] === 0x47 && h[1] === 0x49 && h[2] === 0x46) return "image/gif";
    if (h[0] === 0x52 && h[1] === 0x49 && h[2] === 0x46 && h[3] === 0x46 && h[8] === 0x57 && h[9] === 0x45 && h[10] === 0x42 && h[11] === 0x50) return "image/webp";
  } catch (e) {}
  return "image/jpeg";
}

function safeJson(raw) {
  var t = String(raw || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  try { return JSON.parse(t); } catch (e) {}
  var m = t.match(/\[[\s\S]*\]|\{[\s\S]*\}/);
  if (m) { try { return JSON.parse(m[0]); } catch (e2) {} }
  return null;
}

// ── Claude: detect image type ──────────────────────────────────────────────
async function detectImageType(imageBase64) {
  try {
    const resp = await axios.post(
      "https://api.anthropic.com/v1/messages",
      {
        model: "claude-sonnet-4-5",
        max_tokens: 50,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: sniffMediaType(imageBase64), data: imageBase64 } },
            { type: "text", text: "Look at this image carefully and decide if it is a bank transfer confirmation slip or a bill/invoice.\n\nIt is a SLIP if it shows ANY of these:\n- 'Transfer Completed' or 'Transfer Successful'\n- 'โอนเงินสำเร็จ' or 'รายการสำเร็จ'\n- 'Transaction No' or 'Reference No' or 'เลขที่รายการ'\n- Sender account AND receiver account with bank names\n- FROM and TO with account numbers\n- Banks: SCB, KBank, KBIZ, GSB, Krungthai, Bangkok Bank, Krungsri, TMB, TTB, PromptPay\n\nIt is a BILL if it shows ANY of these:\n- QR code for payment (even if it shows an amount or bank account number)\n- Invoice or receipt with list of products/services\n- 'Please pay' or 'Amount due' or 'ยอดที่ต้องชำระ'\n- Bill with company name and itemized costs\n- Bank account number for receiving payment (but NO transfer confirmation)\n\nKEY RULE: A QR code or bank account number shown on a bill is NOT a slip. A slip must show transfer confirmation that money has already been sent.\n\nKEY RULE 2: If the image shows NO sender (no From bank/account) and no transfer-success wording, it is a BILL — even when it shows a recipient name and an amount.\n\nReturn ONLY one word: SLIP, BILL, or UNKNOWN" }
          ]
        }]
      },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    const result = resp.data.content[0].text.trim().toUpperCase();
    console.log("Image type:", result);
    if (result.indexOf("SLIP") !== -1) return "SLIP";
    if (result.indexOf("BILL") !== -1) return "BILL";
    return "UNKNOWN";
  } catch (e) {
    console.error("Detect error:", e.message, e.response ? JSON.stringify(e.response.data) : "");
    return "UNKNOWN";
  }
}

// ── Claude: read slip details ──────────────────────────────────────────────
async function readSlip(imageBase64, caption) {
  try {
    const resp = await axios.post(
      "https://api.anthropic.com/v1/messages",
      {
        model: "claude-sonnet-4-5",
        max_tokens: 800,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: sniffMediaType(imageBase64), data: imageBase64 } },
            { type: "text", text: "You are an expert Thai bank payment slip reader. Extract ALL details carefully.\n\nToday is " + format(new Date(), "yyyy-MM-dd") + " — slips are almost always from the last few weeks.\n\nDate formats:\n- DD/MM/YYYY or DD/MM/YY\n- DD MMM YYYY (e.g. 08 พ.ค. 2568)\n- Buddhist year (พ.ศ.) subtract 543 to get AD. 2568=2025, 2569=2026\n- 2-digit years: with an ENGLISH month (e.g. \"28 Oct 26\") the year is AD 20YY → 2026. With a THAI month (e.g. \"28 ต.ค. 69\") it is พ.ศ. 25YY → 2569 → 2026.\n- Time: HH:MM or HH:MM:SS\n\nAccounts: the slip shows FROM (sender = us) and TO (recipient). account_from / account_to = the last 4 DIGITS shown for that account (e.g. xxx-x-x8464-x → 8464). Read each digit carefully (6 vs 8, 3 vs 8, 1 vs 7). Use the digits only (XXX-XXX009-4 → 0094).\n\nNames: copy names EXACTLY as printed, letter by letter, including Thai tone marks and ์ (e.g. เช้นจ์ซี is not เซนซี). NEVER output a name that does not appear on this document — use null if unreadable.\n\nSCB \"ใบแจ้งการชำระเงิน / Payment Advice\" (bulk / payroll transfer): the recipient is \"ชื่อบัญชี / Account Name\" with \"เลขที่บัญชี/พร้อมเพย์\" in the upper block (NOT the bank name, NOT the sender); the sender is the \"Sender Details / รายละเอียดผู้จ่ายเงิน\" box; date = Transaction Date; amount = Grand Total Amount; purpose = \"Customer Transaction Reference / ข้อมูลอ้างอิงรายการ\"; reference_number = Customer Batch Reference.\nBill Payment slips: the TO side has no bank — bank_to = null, put the Biller ID in biller_id.\n\nBanks: SCB, Krungthai (KTB), Bangkok Bank (BBL), Kasikorn (KBank), Krungsri (BAY), TMB, GSB, PromptPay.\nUser note: \"" + (caption || "") + "\"\nOwner's past category choices (match this style): " + JSON.stringify(categoryExamples(15, false)) + "\n\nReturn ONLY valid JSON:\n{\n  \"bank_from\": \"bank name or null\",\n  \"account_from\": \"last 4 digits or null\",\n  \"bank_to\": \"bank name or null\",\n  \"account_to\": \"last 4 digits or null\",\n  \"sender_name\": \"FROM / sender account name or null\",\n  \"recipient_name\": \"TO name exactly as printed, or null\",\n  \"biller_id\": \"Biller ID digits or null\",\n  \"amount\": 0.00,\n  \"transaction_date\": \"YYYY-MM-DD or null\",\n  \"transaction_time\": \"HH:MM or null\",\n  \"reference_number\": \"ref or null\",\n  \"purpose\": \"use user note if given, else the slip memo or what is paid to the TO/recipient — never describe it as a payment to the FROM/sender name\",\n  \"category\": \"exactly one string from this list that best fits the payment, or null: " + JSON.stringify(loadCategories()) + "\",\n  \"slip_type\": \"mobile_banking or internet_banking or prompt_pay\"\n}" }
          ]
        }]
      },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    const raw = resp.data.content[0].text.trim();
    console.log("Claude slip:", raw);
    return safeJson(raw);
  } catch (e) {
    console.error("Claude slip error:", e.message, e.response ? JSON.stringify(e.response.data) : "");
    return null;
  }
}

// ── Claude: read bill amount only ──────────────────────────────────────────
async function readBillAmount(imageBase64) {
  try {
    const resp = await axios.post(
      "https://api.anthropic.com/v1/messages",
      {
        model: "claude-sonnet-4-5",
        max_tokens: 100,
        messages: [{
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: sniffMediaType(imageBase64), data: imageBase64 } },
            { type: "text", text: "Look at this bill or invoice image. Find the TOTAL amount to pay.\nReturn ONLY a JSON object, nothing else:\n{\"amount\": 0.00}\nAmount must be a number. No commas. If you cannot find an amount return {\"amount\": null}" }
          ]
        }]
      },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    const raw = resp.data.content[0].text.trim();
    const parsed = safeJson(raw);
    if (!parsed) { console.error("Bill amount: no JSON in reply:", raw.slice(0,80)); return null; }
    return (parsed.amount === 0 || parsed.amount) ? parsed.amount : null;
  } catch (e) {
    console.error("Bill amount error:", e.message, e.response ? JSON.stringify(e.response.data) : "");
    return null;
  }
}

async function getImage(client, messageId) {
  const stream = await client.getMessageContent(messageId);
  const parts = [];
  for await (const chunk of stream) parts.push(chunk);
  return Buffer.concat(parts);
}

// ── Create records ─────────────────────────────────────────────────────────
// ── Date sanity: a slip dated far from its upload time almost always has a misread year ──
var DATE_PAST_DAYS = 183, DATE_FUTURE_DAYS = 7;
function dateLooksWrong(dateStr, refMs) {
  var t = Date.parse(dateStr || "");
  if (isNaN(t)) return false;
  return t > refMs + DATE_FUTURE_DAYS * 86400000 || t < refMs - DATE_PAST_DAYS * 86400000;
}
// Same day/month, year moved to the one closest to the upload time (null if nothing fits)
function suggestYearFix(dateStr, refMs) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(dateStr || "");
  if (!m) return null;
  var refY = new Date(refMs).getUTCFullYear();
  var best = null;
  [refY - 1, refY, refY + 1].forEach(function(y) {
    var cand = y + "-" + m[2] + "-" + m[3];
    if (cand === dateStr || dateLooksWrong(cand, refMs)) return;
    var d = Math.abs(Date.parse(cand) - refMs);
    if (!best || d < best.d) best = { date: cand, d: d };
  });
  return best ? best.date : null;
}
function suspectDates(list) {
  return list.filter(function(p) {
    return p.transaction_date && !p.date_ok && dateLooksWrong(p.transaction_date, Date.parse(p.savedAt || "") || Date.now());
  }).map(function(p) {
    return { id: p.id, display_id: showId(p), date: p.transaction_date, savedAt: p.savedAt || null, suggested: suggestYearFix(p.transaction_date, Date.parse(p.savedAt || "") || Date.now()) };
  });
}

function createSlipRecord(buf, data, source, lineMsgId) {
  if (data && dateLooksWrong(data.transaction_date, Date.now())) {
    var fixed = suggestYearFix(data.transaction_date, Date.now());
    if (fixed) {
      console.log("Year auto-fixed:", data.transaction_date, "→", fixed);
      data.date_autofixed_from = data.transaction_date;
      data.transaction_date = fixed;
    }
  }
  if (data && data.category) {
    var catsV = loadCategories();
    if (catsV.indexOf(data.category) === -1) data.category = null;
  }
  if (data && (data.recipient_name || data.biller_id)) {
    var cn = correctedName(data);
    if (cn && cn !== data.recipient_name) {
      console.log("Name memory:", data.recipient_name, "→", cn);
      data.recipient_read = data.recipient_name;
      data.recipient_name = cn;
    }
  }
  if (data && data.recipient_name) {
    var vmNew = vendorCategoryMap();
    var vk = vendorKey(data.recipient_name);
    if (vmNew[vk]) data.category = vmNew[vk]; // your past choice beats the AI guess
  }
  var fname = "slip_" + Date.now() + ".jpg";
  fs.writeFileSync(path.join(SLIPS_DIR, fname), buf);
  var list = loadPayments();
  var rec = Object.assign({
    id: nextNumericId(),
    display_id: nextDisplayId(data.transaction_date),
    imageFile: fname,
    billFile: null,
    savedAt: new Date().toISOString(),
    source: source,
    line_msg_ids: lineMsgId ? [lineMsgId] : []
  }, data);
  list.push(rec);
  savePayments(list);
  var reg = loadAccounts();
  if (registerFromAccount(reg, rec)) saveAccounts(reg);
  setImmediate(function() {
    flagInternalCandidates([rec.id], false).catch(function(e) { console.error("internal flag error:", e.message); });
  });
  return rec;
}

function createInboxItem(buf, type, billAmount, source) {
  var fname = "img_" + Date.now() + "_" + Math.floor(Math.random()*1000) + ".jpg";
  fs.writeFileSync(path.join(BILLS_DIR, fname), buf);
  var inbox = loadInbox();
  var item = {
    id: "ib" + Date.now() + Math.floor(Math.random()*1000),
    file: fname,
    type: type || "unknown",      // "bill" | "unknown"
    billAmount: (billAmount === 0 || billAmount) ? billAmount : null,
    uploadedAt: new Date().toISOString(),
    source: source
  };
  inbox.push(item);
  saveInbox(inbox);
  return item;
}

function slipReply(rec) {
  var amt = rec.amount ? "฿" + Number(rec.amount).toLocaleString("th-TH", { minimumFractionDigits: 2 }) : "?";
  var reg = loadAccounts();
  var c = classify(rec, reg);
  var lines = [
    "✅ #" + showId(rec) + "  " + (rec.bank_from || "?") + "  " + amt,
    "→ " + (rec.recipient_name || rec.bank_to || "?") + (rec.purpose ? "  (" + rec.purpose + ")" : "") + (rec.category ? "  🏷" + rec.category.split("/")[0] : "")
  ];
  if (c.kind === "internal") lines.push("🔁 โอนภายใน — ไม่นับเป็นค่าใช้จ่าย");
  if (c.kind === "petty_cash") lines.push("💵 เงินสดย่อย → " + (c.toBranch || "?"));
  if (c.from && !c.from.branch && c.from.firstSeenSlip === rec.id) {
    var sim = similarAccount(reg, c.from);
    lines.push(sim
      ? "🆕 " + c.from.bank + " ****" + c.from.last4 + " — หรืออ่านผิดจาก ****" + sim.last4 + "? ถ้าผิด แก้เลขบัญชีใน /app (Edit)"
      : "🆕 บัญชีใหม่ " + c.from.bank + " ****" + c.from.last4 + " — ตั้งสาขาที่ /app → Report");
  }
  var allR = loadPayments();
  var dup = duplicateOf(rec, allR);
  if (dup) lines.push("⚠️ อาจซ้ำกับ #" + showId(dup) + " (" + (dup.transaction_date || "") + " " + baht(dup.amount) + ") — เช็กใน /app → Report");
  var simTo = similarToAccount(rec, allR);
  if (simTo) lines.push("⚠️ บัญชีปลายทาง ****" + String(rec.account_to).replace(/\D/g, "").slice(-4) + " — เคยโอนให้ผู้รับนี้ที่ ****" + simTo + " อ่านผิด?");
  if (recipientLooksLikeSender(rec, reg)) lines.push("⚠️ ผู้รับไม่ชัด (อ่านได้ชื่อเดียวกับผู้โอน) — แก้ชื่อผู้รับใน /app");
  if (rec.recipient_read) lines.push("✏️ ชื่อผู้รับแก้ตามที่เคยแก้ไว้: " + rec.recipient_read + " → " + rec.recipient_name);
  if (rec.date_autofixed_from) lines.push("📅 ปีบนสลิปอ่านได้ " + rec.date_autofixed_from + " → แก้เป็น " + rec.transaction_date + " (เช็กได้ใน /app)");
  lines.push("📎 Match bills at " + BASE_URL + "/app");
  return lines.join("\n");
}

// ── Excel helpers ──────────────────────────────────────────────────────────
function styleHdr(row, color) {
  row.eachCell(function(c) {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color || "FF1565C0" } };
    c.alignment = { horizontal: "center", vertical: "middle" };
  });
  row.height = 22;
}
function normalizeBank(b) {
  var s = (b || "Unknown").trim();
  var l = s.toLowerCase();
  if (/kasikorn|kbank|kbiz|กสิกร/.test(l)) return "KBank";
  if (/government\s*savings|gsb|ออมสิน/.test(l)) return "GSB";
  if (/ayudhya|krungsri|(^|[^a-z])bay([^a-z]|$)|กรุงศรี/.test(l)) return "Krungsri";
  if (/krung\s*thai|ktb|กรุงไทย/.test(l)) return "Krungthai";
  if (/bangkok\s*bank|bbl|กรุงเทพ/.test(l)) return "Bangkok Bank";
  if (/siam\s*commercial|scb|ไทยพาณิชย์/.test(l)) return "SCB";
  if (/tmb|ttb|thanachart|ทหารไทย|ธนชาต/.test(l)) return "TTB";
  if (/baac|เพื่อการเกษตร/.test(l)) return "BAAC";
  if (/prompt\s*pay|พร้อมเพย์/.test(l)) return "PromptPay";
  if (/uob/.test(l)) return "UOB";
  return s;
}
function getBankColor(bank) {
  var map = { "SCB":"FF4E2E8C","Kasikorn":"FF1A8C2E","KBank":"FF1A8C2E","Krungthai":"FF009FDA","KTB":"FF009FDA","Bangkok":"FF0050A0","BBL":"FF0050A0","Krungsri":"FFD4A017","BAY":"FFD4A017","GSB":"FF8B0000","TMB":"FF0056A6","TTB":"FF0056A6","BAAC":"FF0B7A3B","UOB":"FF003087","PromptPay":"FF2D7DD2" };
  var found = Object.keys(map).find(function(k) { return (bank||"").toLowerCase().indexOf(k.toLowerCase()) !== -1; });
  return found ? map[found] : "FF1565C0";
}

// ── Accounts registry: which bank account belongs to which branch ─────────
// accounts.json = { branches: [name], aliases: [our company names], accounts: [{bank,last4,alt[],branch,type,label,firstSeenSlip}] }
// Branch / internal status is worked out at report time from this registry, so editing it re-classifies past months too.
var ACCT_TYPES = ["company", "petty_cash", "personal", "external"];
var PETTY_CAT = "เงินสดย่อย/Petty cash";
var NO_BRANCH = "❓ ยังไม่ระบุสาขา";

function loadJsonFile(f, def) {
  try {
    if (!fs.existsSync(f)) return def;
    return JSON.parse(fs.readFileSync(f, "utf8"));
  } catch (e) { console.error("Read " + path.basename(f) + ":", e.message); return def; }
}
function loadAccounts() {
  var a = loadJsonFile(ACCTS_FILE, null) || {};
  return {
    branches: Array.isArray(a.branches) ? a.branches : [],
    aliases: Array.isArray(a.aliases) ? a.aliases : [],
    accounts: Array.isArray(a.accounts) ? a.accounts : []
  };
}
function saveAccounts(a) { fs.writeFileSync(ACCTS_FILE, JSON.stringify(a, null, 2)); }
function loadSettings() { return loadJsonFile(SETTINGS_FILE, null) || {}; }
function saveSettings(s) { fs.writeFileSync(SETTINGS_FILE, JSON.stringify(s, null, 2)); }

function last4(acct) {
  var d = String(acct || "").replace(/\D/g, "");
  return d.length >= 3 ? d.slice(-4) : null;
}
// Banks show different slices of one number (168-3-58464-5 → "8464" on slips, "X4645" / "X584645" on statements);
// with the full number saved, any 4+ digit slice of it matches
function findAccount(reg, bank, acct) {
  var l = last4(acct);
  if (!l) return null;
  var nb = normalizeBank(bank);
  var digits = String(acct || "").replace(/\D/g, "");
  return reg.accounts.find(function(a) { return a.bank === nb && (a.last4 === l || (a.alt || []).indexOf(l) !== -1); }) ||
    (digits.length >= 4 ? reg.accounts.find(function(a) { return a.bank === nb && a.full && a.full.indexOf(digits) !== -1; }) : null) || null;
}
// Paying (from) accounts are discovered automatically from slips
function registerFromAccount(reg, p) {
  if (!last4(p.account_from) || findAccount(reg, p.bank_from, p.account_from)) return false;
  reg.accounts.push({ bank: normalizeBank(p.bank_from), last4: last4(p.account_from), alt: [], branch: null, type: "company", label: "", firstSeenSlip: p.id, firstSeenAt: new Date().toISOString() });
  return true;
}
// Same bank, same length, exactly one digit different → probably a misread (6284 vs 8284)
function similarAccount(reg, acct) {
  return reg.accounts.find(function(a) {
    if (a === acct || a.bank !== acct.bank || a.last4.length !== acct.last4.length || a.type === "external") return false;
    var diff = 0;
    for (var i = 0; i < a.last4.length; i++) if (a.last4[i] !== acct.last4[i]) diff++;
    return diff === 1;
  }) || null;
}
function syncAccounts() {
  var reg = loadAccounts();
  var added = 0;
  loadPayments().forEach(function(p) { if (registerFromAccount(reg, p)) added++; });
  if (added) saveAccounts(reg);
  return reg;
}

// Split amounts across branches; whatever is not split stays with the home branch
function allocate(amount, home, split) {
  var out = {}, used = 0;
  (Array.isArray(split) ? split : []).forEach(function(s) {
    var a = Number(s.amount) || 0;
    if (!s.branch || a <= 0) return;
    out[s.branch] = (out[s.branch] || 0) + a;
    used += a;
  });
  var rest = Math.round((amount - used) * 100) / 100;
  if (rest > 0.009) out[home] = (out[home] || 0) + rest;
  return Object.keys(out).map(function(k) { return { branch: k, amount: Math.round(out[k] * 100) / 100 }; });
}

// kind: "expense" | "internal" (not counted) | "petty_cash" (counted as expense of the petty-cash branch)
function classify(p, reg) {
  var from = findAccount(reg, p.bank_from, p.account_from);
  var to = findAccount(reg, p.bank_to, p.account_to);
  if (from && from.type === "external") from = null;
  if (to && to.type === "external") to = null;
  if (to && to.type === "personal") to = null; // paying a person (salary / reimbursement) is a real expense
  var kind = "expense", via = null;
  if (p.internal === true) { kind = "internal"; via = "manual"; }
  else if (to && to.type === "petty_cash") { kind = "petty_cash"; via = "account"; }
  else if (p.internal === false) { kind = "expense"; via = "manual"; }
  else if (to) { kind = "internal"; via = "account"; }
  var fromBranch = (from && from.branch) || null;
  var toBranch = (to && to.branch) || null;
  var amount = Number(p.amount) || 0;
  var home = (kind === "petty_cash" ? toBranch : fromBranch) || NO_BRANCH;
  return {
    kind: kind, via: via, from: from, to: to, fromBranch: fromBranch, toBranch: toBranch,
    category: p.category || (kind === "petty_cash" ? PETTY_CAT : null),
    alloc: kind === "internal" ? [] : allocate(amount, home, p.split),
    pending: kind === "expense" && p.internal == null && p.internal_suggest === "pending",
    needsBill: kind === "expense"
  };
}

// Internal transfer B→A that is a share of an expense A paid (e.g. 50,000 of a 100,000 ad bill)
var NICE_RATIOS = [1, 0.5, 1/3, 2/3, 0.25, 0.75, 0.2, 0.4, 0.6, 0.8];
function linkSuggestions(t, all, reg) {
  var ct = classify(t, reg);
  if (ct.kind !== "internal" || t.linked_to || !t.transaction_date) return [];
  if (!ct.fromBranch || !ct.toBranch || ct.fromBranch === ct.toBranch) return [];
  var ta = Number(t.amount) || 0;
  if (!ta) return [];
  var td = parseISO(t.transaction_date).getTime();
  var out = [];
  all.forEach(function(e) {
    if (e.id === t.id || !e.transaction_date) return;
    if ((t.link_dismissed || []).indexOf(e.id) !== -1) return;
    var ce = classify(e, reg);
    if (ce.kind !== "expense" || ce.fromBranch !== ct.toBranch) return;
    var ea = Number(e.amount) || 0;
    if (ea < ta - 1) return;
    var days = Math.abs(parseISO(e.transaction_date).getTime() - td) / 86400000;
    if (days > 30) return;
    var r = ta / ea;
    var off = NICE_RATIOS.reduce(function(b, n) { return Math.min(b, Math.abs(r - n)); }, 1);
    if (off > 0.005) return;
    out.push({ id: e.id, display_id: showId(e), amount: ea, date: e.transaction_date, purpose: e.purpose || "", recipient: e.recipient_name || "", pct: Math.round(r * 1000) / 10, score: days + off * 1000 });
  });
  return out.sort(function(a, b) { return a.score - b.score; }).slice(0, 3);
}

// ── Our-company name matching (aliases) → "is this an internal transfer?" suggestions ──
function normName(s) {
  return String(s || "").toLowerCase()
    .replace(/บริษัท|บจก\.?|จำกัด|\(มหาชน\)|มหาชน|หจก\.?|ห้างหุ้นส่วน/g, "")
    .replace(/\b(company|limited|co|ltd|inc|plc|public)\b\.?/g, "")
    .replace(/[^a-z0-9฀-๿]/g, "");
}
function nameMatchesAlias(name, aliases) {
  var n = normName(name);
  if (n.length < 3) return false;
  return aliases.some(function(a) {
    var x = normName(a);
    if (x.length < 3) return false;
    return n.indexOf(x) !== -1 || (n.length >= 4 && x.indexOf(n) !== -1);
  });
}
async function aiOurNames(names, aliases) {
  var found = [];
  for (var off = 0; off < names.length; off += 80) {
    var chunk = names.slice(off, off + 80);
    var resp = await axios.post(
      "https://api.anthropic.com/v1/messages",
      { model: "claude-sonnet-4-5", max_tokens: 2000, messages: [{ role: "user", content:
        "Our company group's names (Thai/English, any of our legal entities or branches): " + JSON.stringify(aliases) +
        "\nBelow are recipient names read from bank transfer slips. Which ones are OUR OWN company (same entity — Thai/English versions, abbreviations, or names cut off on the slip count)? Employees, people and outside vendors are NOT ours." +
        "\nNames: " + JSON.stringify(chunk) +
        "\nReturn ONLY a JSON array of the matching names copied exactly as given, or [] if none." }] },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    var arr = safeJson(resp.data.content[0].text);
    if (Array.isArray(arr)) arr.forEach(function(n) { if (chunk.indexOf(n) !== -1) found.push(n); });
  }
  return found;
}
// Marks slips as internal_suggest "pending" (asks the owner) — never changes amounts or the internal flag itself.
// force = re-check slips already checked (used by the back-scan after aliases change)
async function flagInternalCandidates(ids, force) {
  var reg = loadAccounts();
  if (!reg.aliases.length) return { checked: 0, flagged: 0 };
  var idSet = {};
  ids.forEach(function(id) { idSet[String(id)] = true; });
  var targets = loadPayments().filter(function(p) {
    if (!idSet[String(p.id)] || !p.recipient_name || p.internal != null) return false;
    if (p.internal_suggest === "pending") return false;
    if (p.internal_suggest && !force) return false;
    return classify(p, reg).kind === "expense";
  });
  if (!targets.length) return { checked: 0, flagged: 0 };
  var ours = {};
  var unsure = [];
  targets.forEach(function(p) {
    if (nameMatchesAlias(p.recipient_name, reg.aliases)) ours[p.recipient_name] = true;
    else if (unsure.indexOf(p.recipient_name) === -1) unsure.push(p.recipient_name);
  });
  if (unsure.length && ANTHROPIC_KEY) {
    (await aiOurNames(unsure, reg.aliases)).forEach(function(n) { ours[n] = true; });
  }
  var flagged = 0;
  var fresh = loadPayments();
  targets.forEach(function(t) {
    var i = fresh.findIndex(function(p) { return p.id === t.id; });
    if (i === -1 || fresh[i].internal != null) return;
    fresh[i].internal_suggest = ours[t.recipient_name] ? "pending" : "no";
    if (ours[t.recipient_name]) flagged++;
  });
  savePayments(fresh);
  return { checked: targets.length, flagged: flagged };
}

// ── Bank statements: import, balance check, match with slips ──────────────
// The statement is the bank's own record; slips and bills are the evidence attached to it.
// statements.json = { files: {key: info}, lines: [line], drive: {fileId: modifiedTime} }
const STMT_FILE = path.join(DATA_DIR, "statements.json");
const STMT_DIR = path.join(DATA_DIR, "statements"); // original files, kept for audit
const PAYROLL_DIR = path.join(DATA_DIR, "payroll");
function loadStatements() {
  var d = loadJsonFile(STMT_FILE, null) || {};
  if (!d.files) d.files = {};
  if (!Array.isArray(d.lines)) d.lines = [];
  if (!d.drive) d.drive = {};
  return d;
}
function saveStatements(d) { fs.writeFileSync(STMT_FILE, JSON.stringify(d)); }
function round2(n) { return Math.round(n * 100) / 100; }
function sha(s) { return crypto.createHash("sha1").update(s).digest("hex"); }

// RFC-4180 CSV (quoted fields may contain commas and new lines)
function parseCsv(text) {
  var rows = [], row = [], cell = "", q = false;
  for (var i = 0; i < text.length; i++) {
    var ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else q = false; }
      else cell += ch;
    } else if (ch === '"') q = true;
    else if (ch === ",") { row.push(cell); cell = ""; }
    else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(cell); rows.push(row); row = []; cell = "";
    } else cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows;
}
// UTF-8, or Thai Windows-874 when the bytes are not valid UTF-8
function decodeText(buf) {
  var t = buf.toString("utf8");
  if (t.indexOf("�") !== -1) { try { t = new TextDecoder("windows-874").decode(buf); } catch (e) {} }
  return t.replace(/^﻿/, "");
}
function toNum(v) {
  var s = String(v == null ? "" : v).replace(/[,\s"]/g, "");
  if (!s) return null;
  var neg = /^\(.*\)$/.test(s);
  var n = Number(s.replace(/[()]/g, ""));
  return isNaN(n) ? null : (neg ? -n : n);
}
function dmyToIso(d) { var m = /^(\d{2})\/(\d{2})\/(\d{4})$/.exec(String(d || "").trim()); return m ? m[3] + "-" + m[2] + "-" + m[1] : null; }

function detectFileKind(text) {
  var first = text.split(/\r?\n/)[0] || "";
  if (/K-DEPOSIT STATEMENT/i.test(text.slice(0, 600))) return "kbank_csv";
  if (/^Account Number,Account Name,Account Type/i.test(first)) return "scb_csv";
  if (/(^|,)Branch(,|$)/i.test(first) && /NET PAY/i.test(first)) return "payroll_csv";
  return null;
}

// "โอนไป SCB X6710 บริษัท …", "รับโอนจาก KBANK x4645 …", "PromptPay to x2484 MR …", "เพื่อชำระ Ref X5042 บจก. บัตรกรุงศรีอยุธยา"
function parseCounterparty(desc, ownBank) {
  var d = String(desc || "").replace(/\+\+\s*$/, "").trim();
  var bill = /เพื่อชำระ\s+(?:([A-Z]{2,6})\s+)?Ref\s+\S+\s+(.*)$/i.exec(d);
  if (bill) return { bank: null, digits: null, name: bill[2].trim(), bill: true };
  var m = /(?:^|\s)(?:([A-Z]{2,6})\s+)?[xX](\d{3,})\s*(.*)$/.exec(d);
  if (!m) return { bank: null, digits: null, name: d };
  var bank = m[1] && !/^(TO|REF)$/i.test(m[1]) ? normalizeBank(m[1]) : null;
  if (/พร้อมเพย์|promptpay/i.test(d)) bank = "PromptPay";
  return { bank: bank || ownBank, digits: m[2], name: (m[3] || "").replace(/\+\+$/, "").trim() };
}

function parseKBankCsv(text) {
  var rows = parseCsv(text), acct = null, period = null, opening = null, lines = [];
  rows.forEach(function(r) {
    var c = r.map(function(x) { return String(x || "").trim(); });
    var i = c.indexOf("เลขที่บัญชีเงินฝาก");
    if (i !== -1) acct = c.slice(i + 1).find(function(x) { return /\d{3}-\d/.test(x); }) || acct;
    i = c.indexOf("รอบระหว่างวันที่");
    if (i !== -1) period = c.slice(i + 1).find(function(x) { return /\d{2}\/\d{2}\/\d{4}/.test(x); }) || period;
    var dm = /^(\d{2})-(\d{2})-(\d{2})$/.exec(c[1] || "");
    if (!dm) return;
    if (c[3] === "ยอดยกมา") { opening = toNum(c[8]); return; }
    lines.push({ date: "20" + dm[3] + "-" + dm[2] + "-" + dm[1], time: c[2] || "", type: c[3], out: toNum(c[4]), in: toNum(c[6]), balance: toNum(c[8]), channel: c[10] || "", desc: c[12] || "" });
  });
  var pm = /(\d{2}\/\d{2}\/\d{4})\s*-\s*(\d{2}\/\d{2}\/\d{4})/.exec(period || "");
  return { bank: "KBank", full: String(acct || "").replace(/\D/g, ""), period: pm ? { from: dmyToIso(pm[1]), to: dmyToIso(pm[2]) } : null, opening: opening, lines: lines };
}

function parseScbCsv(text) {
  var rows = parseCsv(text), lines = [], full = null;
  rows.slice(1).forEach(function(r) {
    if (r.length < 14 || !dmyToIso(r[5])) return;
    full = full || String(r[0]).replace(/\D/g, "");
    lines.push({ date: dmyToIso(r[5]), time: String(r[6] || "").trim(), type: String(r[8] || "").trim(), out: toNum(r[11]), in: toNum(r[12]), balance: toNum(r[13]), channel: String(r[9] || "").trim(), desc: String(r[14] || "").trim() });
  });
  if (!lines.length) return { bank: "SCB", full: full, period: null, opening: null, lines: [] };
  var f = lines[0];
  var opening = round2(f.balance + (f.out || 0) - (f.in || 0)); // export has no opening row
  return { bank: "SCB", full: full, period: { from: lines[0].date, to: lines[lines.length - 1].date }, opening: opening, lines: lines };
}

// every running balance must follow from the previous one — catches a bad file before it is trusted
function checkBalances(opening, lines) {
  if (opening === null || opening === undefined) return { ok: false, error: "no opening balance found" };
  var bal = opening;
  for (var i = 0; i < lines.length; i++) {
    bal = round2(bal - (lines[i].out || 0) + (lines[i].in || 0));
    if (lines[i].balance !== null && Math.abs(bal - lines[i].balance) > 0.005) {
      return { ok: false, error: "balance does not add up at line " + (i + 1) + " (" + lines[i].date + " " + lines[i].time + "): expected " + bal.toFixed(2) + ", file says " + lines[i].balance.toFixed(2) };
    }
  }
  return { ok: true, closing: bal };
}

// find (or add) the registry entry for a statement's account; remembers the full number
function accountForStatement(reg, bank, full) {
  var a = reg.accounts.find(function(x) { return x.bank === bank && x.full === full; }) ||
    reg.accounts.find(function(x) { return x.bank === bank && !x.full && [x.last4].concat(x.alt || []).some(function(l) { return l && full.indexOf(l) !== -1; }); });
  if (a) { a.full = full; return a; }
  a = { bank: bank, last4: full.slice(-4), alt: [], full: full, branch: null, type: "company", label: "", firstSeenSlip: null, firstSeenAt: new Date().toISOString() };
  reg.accounts.push(a);
  return a;
}
function acctLabelOf(a) { return a ? a.bank + " ****" + a.last4 + (a.label ? " " + a.label : "") + (a.branch ? " · " + a.branch : "") : "?"; }

function importStatement(text, info) {
  var kind = detectFileKind(text);
  if (kind === "payroll_csv") return importPayroll(text, info);
  if (!kind) return { ok: false, skipped: true, error: "not a bank statement this app can read (KBank CSV, SCB CSV, payroll CSV)" };
  var P = kind === "kbank_csv" ? parseKBankCsv(text) : parseScbCsv(text);
  // SCB exports carry no date range: an account with few payments (e.g. payroll only) would look like "part of the month".
  // When every line is in the month of its folder (1026) or the month chosen for the upload, the file covers that month.
  var hint = info.folderMonth || info.month;
  if (kind === "scb_csv" && hint && P.lines.length && P.lines.every(function(l) { return l.date.slice(0, 7) === hint; })) {
    P.period = { from: hint + "-01", to: hint + "-" + String(new Date(+hint.slice(0, 4), +hint.slice(5), 0).getDate()).padStart(2, "0") };
  }
  if (!P.full || P.full.length < 8) return { ok: false, error: "account number not found in file" };
  if (!P.lines.length) return { ok: false, error: "no transactions in file" };
  var bc = checkBalances(P.opening, P.lines);
  if (!bc.ok) return { ok: false, error: bc.error };
  var reg = loadAccounts();
  var acct = accountForStatement(reg, P.bank, P.full);
  saveAccounts(reg);
  var st = loadStatements();
  var key = sha(P.bank + P.full + P.period.from + P.period.to).slice(0, 16);
  // keep your earlier decisions on lines that come back in a re-import
  var old = {};
  st.lines.forEach(function(l) { old[l.id] = l; });
  st.lines = st.lines.filter(function(l) { return l.file !== key; });
  var have = {};
  st.lines.forEach(function(l) { have[l.id] = true; });
  var added = 0;
  P.lines.forEach(function(l) {
    var id = sha([P.full, l.date, l.time, l.out, l.in, l.balance].join("|")).slice(0, 12);
    if (have[id]) return; // same line already imported from an overlapping export
    var cp = parseCounterparty(l.desc, P.bank);
    var rec = Object.assign({ id: id, file: key, bank: P.bank, account: P.full, counterparty: cp }, l);
    if (old[id]) ["slipId", "resolved"].forEach(function(f) { if (old[id][f] !== undefined) rec[f] = old[id][f]; });
    st.lines.push(rec);
    have[id] = true;
    added++;
  });
  st.lines.sort(function(a, b) { return (a.date + a.time).localeCompare(b.date + b.time); });
  st.files[key] = { key: key, name: info.name || "", source: info.source || "upload", driveId: info.driveId || null, bank: P.bank, account: P.full,
    period: P.period, opening: P.opening, closing: bc.closing, count: P.lines.length, importedAt: new Date().toISOString() };
  saveStatements(st);
  try {
    if (!fs.existsSync(STMT_DIR)) fs.mkdirSync(STMT_DIR, { recursive: true });
    fs.writeFileSync(path.join(STMT_DIR, key + ".csv"), text);
  } catch (e) { console.error("statement copy:", e.message); }
  return { ok: true, kind: kind, account: acctLabelOf(acct), period: P.period, lines: P.lines.length, added: added, closing: bc.closing };
}

// ── Matching statement lines ↔ slips ──
function lineAccount(reg, l) { return reg.accounts.find(function(a) { return a.bank === l.bank && a.full === l.account; }) || null; }
function counterpartyAccount(reg, l) {
  var cp = l.counterparty || {};
  if (!cp.digits) return null;
  var a = findAccount(reg, cp.bank || l.bank, cp.digits);
  return a && a.type !== "external" && a.type !== "personal" ? a : null;
}
function daysApart(a, b) { return Math.abs(Date.parse(a) - Date.parse(b)) / 86400000; }
function minutesOf(t) { var m = /^(\d{1,2}):(\d{2})/.exec(t || ""); return m ? +m[1] * 60 + +m[2] : null; }

// one-to-one: each money-out line gets at most one slip of the same account, amount and ±2 days
function matchLines(lines, slips, reg) {
  var used = {}, byLine = {};
  lines.forEach(function(l) {
    var sid = l.slipId || (l.resolved && l.resolved.slipId);
    if (sid) { byLine[l.id] = sid; used[sid] = true; }
  });
  var pairs = [];
  lines.forEach(function(l) {
    if (byLine[l.id] || l.resolved || !(l.out > 0)) return;
    var la = lineAccount(reg, l);
    slips.forEach(function(p) {
      if (used[p.id] || !p.transaction_date) return;
      if (Math.abs((Number(p.amount) || 0) - l.out) >= 0.01) return;
      if (findAccount(reg, p.bank_from, p.account_from) !== la) return;
      var dd = daysApart(p.transaction_date, l.date);
      if (dd > 2) return;
      var tm = minutesOf(p.transaction_time), lm = minutesOf(l.time);
      pairs.push({ line: l.id, slip: p.id, score: dd * 1440 + (tm !== null && lm !== null ? Math.abs(tm - lm) : 720) });
    });
  });
  pairs.sort(function(a, b) { return a.score - b.score; });
  pairs.forEach(function(x) { if (!byLine[x.line] && !used[x.slip]) { byLine[x.line] = x.slip; used[x.slip] = true; } });
  return { byLine: byLine, used: used };
}

function statementCoverage(st, full, ym) {
  var first = ym + "-01", last = ym + "-" + String(new Date(+ym.slice(0, 4), +ym.slice(5), 0).getDate()).padStart(2, "0");
  var files = Object.values(st.files).filter(function(f) { return f.account === full && f.period && f.period.from <= last && f.period.to >= first; });
  if (!files.length) return "none";
  var from = files.map(function(f) { return f.period.from; }).sort()[0], to = files.map(function(f) { return f.period.to; }).sort().pop();
  return from <= first && to >= last ? "full" : "partial";
}
function lineCovered(st, full, date) {
  return Object.values(st.files).some(function(f) { return f.account === full && f.period && f.period.from <= date && f.period.to >= date; });
}

function reconMonth(y, m) {
  var ym = ymKey(y, m);
  var st = loadStatements(), reg = loadAccounts(), all = loadPayments();
  var near = function(d) { return d && d.slice(0, 7) >= ymKey(prevMonthOf(y, m).y, prevMonthOf(y, m).m) && d.slice(0, 7) <= ymKey(m === 12 ? y + 1 : y, m === 12 ? 1 : m + 1); };
  var lines = st.lines.filter(function(l) { return near(l.date); });
  var slips = all.filter(function(p) { return near(p.transaction_date); });
  var mt = matchLines(lines, slips, reg);
  var byId = {};
  all.forEach(function(p) { byId[p.id] = p; });
  var out = { ym: ym, accounts: [], noSlip: [], notInStatement: [], missingStatements: [], counts: { matched: 0, internal: 0, noSlip: 0, resolved: 0 } };
  var monthLines = lines.filter(function(l) { return l.date.slice(0, 7) === ym; });
  monthLines.forEach(function(l) {
    var sid = mt.byLine[l.id];
    if (l.out > 0) {
      if (l.resolved && l.resolved.how === "ignored") { l.status = "ignored"; out.counts.resolved++; }
      else if (sid) { l.status = "matched"; l.slip = byId[sid] ? { id: sid, display_id: showId(byId[sid]) } : { id: sid }; out.counts.matched++; }
      else if (counterpartyAccount(reg, l)) { l.status = "internal"; out.counts.internal++; }
      else {
        l.status = "noSlip"; out.counts.noSlip++;
        // likely slips: same amount ±10 days, any account (helps when the slip's account was misread)
        l.candidates = all.filter(function(p) { return !mt.used[p.id] && p.transaction_date && Math.abs((Number(p.amount) || 0) - l.out) < 0.01 && daysApart(p.transaction_date, l.date) <= 10; })
          .slice(0, 5).map(function(p) { return { id: p.id, display_id: showId(p), date: p.transaction_date, from: (p.bank_from || "") + " ****" + (p.account_from || ""), recipient: p.recipient_name || "" }; });
        out.noSlip.push(l);
      }
    } else l.status = "in";
  });
  // slips whose account has a statement for that day but no line matched them
  all.forEach(function(p) {
    if (!p.transaction_date || p.transaction_date.slice(0, 7) !== ym || mt.used[p.id] || p.stmt_ok) return;
    var a = findAccount(reg, p.bank_from, p.account_from);
    if (!a || !a.full || !lineCovered(st, a.full, p.transaction_date)) return;
    out.notInStatement.push(p.id);
  });
  // accounts that use statements must have one for this month
  var withFiles = {};
  Object.values(st.files).forEach(function(f) { if (!withFiles[f.account] || f.period.from < withFiles[f.account]) withFiles[f.account] = f.period.from; });
  reg.accounts.forEach(function(a) {
    if (!a.full || !withFiles[a.full]) return;
    var cov = statementCoverage(st, a.full, ym);
    var ls = monthLines.filter(function(l) { return l.account === a.full && l.bank === a.bank; });
    var sum = function(f) { return round2(ls.reduce(function(s, l) { return s + (f(l) || 0); }, 0)); };
    var info = { label: acctLabelOf(a), bank: a.bank, full: a.full, branch: a.branch, coverage: cov, lines: ls.length,
      moneyOut: sum(function(l) { return l.out; }), moneyIn: sum(function(l) { return l.in; }),
      sales: sum(function(l) { return l.in && /ขาย|sale/i.test(l.type) ? l.in : 0; }),
      matched: ls.filter(function(l) { return l.status === "matched"; }).length,
      internal: ls.filter(function(l) { return l.status === "internal"; }).length,
      noSlip: ls.filter(function(l) { return l.status === "noSlip"; }).length };
    out.accounts.push(info);
    if (cov !== "full" && withFiles[a.full] <= ym + "-31") out.missingStatements.push(info.label + (cov === "partial" ? " (part of the month only)" : ""));
  });
  return out;
}

// ── Payroll file → split the payroll batch by branch ──
var MONTHS_EN = ["january","february","march","april","may","june","july","august","september","october","november","december"];
function payrollMonthFromName(name) {
  var m = new RegExp("(" + MONTHS_EN.join("|") + ")[^0-9]{0,3}(\\d{4})", "i").exec(name || "");
  return m ? ymKey(+m[2], MONTHS_EN.indexOf(m[1].toLowerCase()) + 1) : null;
}
function importPayroll(text, info) {
  var nameYm = payrollMonthFromName(info.name);
  var ym = info.folderMonth || nameYm || info.month;
  if (!ym) return { ok: false, error: "payroll month unknown — put it in the month folder (e.g. 1026) or name it like RCG_Payroll_October_2026.csv" };
  var note = info.folderMonth && nameYm && nameYm !== info.folderMonth ? "file name says " + nameYm + " but it is in folder " + info.folderMonth.slice(5) + info.folderMonth.slice(2, 4) + " → used " + ym : null;
  var rows = parseCsv(text);
  var h = rows[0].map(function(x) { return String(x).trim().toLowerCase(); });
  var ix = { name: h.indexOf("name"), nick: h.indexOf("nickname"), branch: h.indexOf("branch"), net: h.indexOf("net pay") };
  if (ix.branch === -1 || ix.net === -1) return { ok: false, error: "payroll file needs Branch and NET PAY columns" };
  var people = rows.slice(1).filter(function(r) { return r.length > ix.net; }).map(function(r) {
    return { name: String(r[ix.name] || "").trim(), nick: ix.nick !== -1 ? String(r[ix.nick] || "").trim() : "", branchRaw: String(r[ix.branch] || "").trim(), net: toNum(r[ix.net]) || 0 };
  }).filter(function(p) { return p.net > 0; });
  if (!fs.existsSync(PAYROLL_DIR)) fs.mkdirSync(PAYROLL_DIR, { recursive: true });
  fs.writeFileSync(path.join(PAYROLL_DIR, ym + ".json"), JSON.stringify({ ym: ym, file: info.name || "", driveId: info.driveId || null, importedAt: new Date().toISOString(), people: people }, null, 2));
  return { ok: true, kind: "payroll_csv", month: ym, note: note, people: people.length, total: round2(people.reduce(function(s, p) { return s + p.net; }, 0)) };
}
function payrollBranch(raw, reg, pnl) {
  var map = pnl.payrollBranchMap || {};
  if (map[raw] && reg.branches.indexOf(map[raw]) !== -1) return map[raw];
  var hit = reg.branches.find(function(b) { return raw.toLowerCase().indexOf(b.toLowerCase()) !== -1; });
  if (hit) return hit;
  if (/สำนักงานใหญ่|head ?office|\bhq\b/i.test(raw) && pnl.hqBranches.length) return pnl.hqBranches[0];
  return null;
}
function isSalary(p) { return /salary|เงินเดือน|payroll/i.test((p.category || "") + " " + (p.purpose || "")); }
// individual transfers = exact NET PAY of one person; the biggest remaining salary payment = the batch
function payrollPlan(y, m) {
  var ym = ymKey(y, m);
  var f = path.join(PAYROLL_DIR, ym + ".json");
  if (!fs.existsSync(f)) return { ym: ym, imported: false };
  var pr = JSON.parse(fs.readFileSync(f, "utf8"));
  var reg = loadAccounts(), pnl = loadPnl();
  var people = pr.people.map(function(p) { return Object.assign({}, p, { branch: payrollBranch(p.branchRaw, reg, pnl) }); });
  var unmapped = {};
  people.forEach(function(p) { if (!p.branch) unmapped[p.branchRaw] = true; });
  var pays = loadPayments().filter(function(p) { return costYm(p) === ym && classify(p, reg).kind === "expense" && isSalary(p); });
  var left = people.slice(), individual = [];
  pays.slice().sort(function(a, b) { return a.amount - b.amount; }).forEach(function(p) {
    var hits = left.filter(function(x) { return Math.abs(x.net - (Number(p.amount) || 0)) < 0.01; });
    if (hits.length === 1) { individual.push({ slipId: p.id, display_id: showId(p), amount: Number(p.amount), person: hits[0].nick || hits[0].name, branch: hits[0].branch }); left.splice(left.indexOf(hits[0]), 1); }
  });
  var rest = pays.filter(function(p) { return !individual.some(function(i) { return i.slipId === p.id; }); }).sort(function(a, b) { return b.amount - a.amount; });
  var batch = null;
  if (rest.length) {
    var b = rest[0], netLeft = left.reduce(function(s, p) { return s + p.net; }, 0);
    var byB = {};
    left.forEach(function(p) { var k = p.branch || "?"; byB[k] = (byB[k] || 0) + p.net; });
    var scale = netLeft > 0 ? Number(b.amount) / netLeft : 0;
    batch = { slipId: b.id, display_id: showId(b), amount: Number(b.amount), netLeft: round2(netLeft), diff: round2(Number(b.amount) - netLeft),
      alloc: Object.keys(byB).map(function(k) { return { branch: k, net: round2(byB[k]), amount: round2(byB[k] * scale) }; }) };
  }
  var byBranch = {};
  people.forEach(function(p) { var k = p.branch || "?"; byBranch[k] = round2((byBranch[k] || 0) + p.net); });
  return { ym: ym, imported: true, file: pr.file, importedAt: pr.importedAt, people: people.length, total: round2(people.reduce(function(s, p) { return s + p.net; }, 0)),
    byBranch: byBranch, unmapped: Object.keys(unmapped), individual: individual, batch: batch, applied: pays.some(function(p) { return p.payroll_ym === ym; }), branches: reg.branches };
}
function applyPayroll(y, m) {
  var plan = payrollPlan(y, m);
  if (!plan.imported) return { ok: false, error: "no payroll file for this month" };
  if (plan.unmapped.length) return { ok: false, error: "match these payroll branches first: " + plan.unmapped.join(", ") };
  var reg = loadAccounts(), list = loadPayments();
  var setSplit = function(id, alloc) {
    var p = list.find(function(x) { return x.id === id; });
    if (!p) return;
    var home = classify(p, reg).fromBranch || NO_BRANCH;
    var split = alloc.filter(function(a) { return a.branch !== home && a.amount > 0.009; }).map(function(a) { return { branch: a.branch, amount: a.amount }; });
    if (split.length) p.split = split; else delete p.split;
    p.payroll_ym = plan.ym;
  };
  plan.individual.forEach(function(i) { setSplit(i.slipId, [{ branch: i.branch, amount: i.amount }]); });
  if (plan.batch) {
    setSplit(plan.batch.slipId, plan.batch.alloc);
    // links from internal transfers moved shares of this slip before — the payroll file replaces them
    list.forEach(function(t) { if (String(t.linked_to) === String(plan.batch.slipId)) { t.linked_to_replaced = t.linked_to; delete t.linked_to; } });
  }
  savePayments(list);
  return { ok: true, plan: payrollPlan(y, m) };
}

// ── Google Drive (service account, read-only, only folders shared with it) ──
function driveCreds() { try { return JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON || ""); } catch (e) { return null; } }
function driveConfigError() {
  var v = process.env.GOOGLE_SERVICE_ACCOUNT_JSON;
  if (!v) return "GOOGLE_SERVICE_ACCOUNT_JSON is not set on Railway (or Railway has not redeployed since)";
  var c;
  try { c = JSON.parse(v); } catch (e) { return "GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON (" + e.message.slice(0, 80) + ") — paste the whole key file, plain quotes"; }
  if (!c.client_email || !c.private_key) return "the JSON has no client_email / private_key — use the service-account key file";
  return null;
}
var DRIVE_API = process.env.DRIVE_API_URL || "https://www.googleapis.com/drive/v3";
var driveTok = null;
function b64url(x) { return Buffer.from(x).toString("base64").replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_"); }
async function driveToken() {
  if (driveTok && driveTok.exp > Date.now() + 60000) return driveTok.token;
  var c = driveCreds();
  if (!c || !c.client_email || !c.private_key) throw new Error("GOOGLE_SERVICE_ACCOUNT_JSON is not set on the server");
  var now = Math.floor(Date.now() / 1000), aud = c.token_uri || "https://oauth2.googleapis.com/token";
  var unsigned = b64url(JSON.stringify({ alg: "RS256", typ: "JWT" })) + "." + b64url(JSON.stringify({ iss: c.client_email, scope: "https://www.googleapis.com/auth/drive.readonly", aud: aud, iat: now, exp: now + 3600 }));
  var sig = crypto.createSign("RSA-SHA256").update(unsigned).sign(c.private_key, "base64").replace(/=+$/, "").replace(/\+/g, "-").replace(/\//g, "_");
  var r = await axios.post(aud, "grant_type=" + encodeURIComponent("urn:ietf:params:oauth:grant-type:jwt-bearer") + "&assertion=" + unsigned + "." + sig, { headers: { "content-type": "application/x-www-form-urlencoded" } });
  driveTok = { token: r.data.access_token, exp: Date.now() + (r.data.expires_in || 3600) * 1000 };
  return driveTok.token;
}
async function driveListAll(folderId) {
  var tok = await driveToken(), files = [], queue = [{ id: folderId, path: "" }], guard = 0;
  while (queue.length && guard++ < 200) {
    var f = queue.shift(), pageToken = null;
    do {
      var r = await axios.get(DRIVE_API + "/files", { headers: { Authorization: "Bearer " + tok }, params: {
        q: "'" + f.id + "' in parents and trashed=false", fields: "nextPageToken,files(id,name,mimeType,modifiedTime,size)",
        pageSize: 200, pageToken: pageToken || undefined, supportsAllDrives: true, includeItemsFromAllDrives: true } });
      (r.data.files || []).forEach(function(x) {
        if (x.mimeType === "application/vnd.google-apps.folder") queue.push({ id: x.id, path: f.path + x.name + "/" });
        else files.push(Object.assign({ path: f.path + x.name }, x));
      });
      pageToken = r.data.nextPageToken;
    } while (pageToken);
  }
  return files;
}
async function driveDownload(id) {
  var r = await axios.get(DRIVE_API + "/files/" + id, { headers: { Authorization: "Bearer " + await driveToken() }, params: { alt: "media", supportsAllDrives: true }, responseType: "arraybuffer" });
  return Buffer.from(r.data);
}
function driveFolderId(link) { var m = /folders\/([\w-]{10,})/.exec(link || "") || /^([\w-]{10,})$/.exec(String(link || "").trim()); return m ? m[1] : null; }
// last 60 files read (Drive or upload) with their outcome — shown on the Report tab
function importMessage(r) {
  if (r.ok) {
    if (r.kind === "sales_xlsx") return "sales " + r.months.join(", ") + " · " + r.bills + " bills";
    if (r.kind === "petty_xlsx") return "petty cash " + r.branch + " " + r.period.from + " → " + r.period.to + " · " + r.entries + " items";
    if (r.kind === "payroll_csv") return "payroll " + r.month + " · " + r.people + " people · " + baht(r.total) + (r.note ? " · ⚠️ " + r.note : "");
    return r.account + " " + r.period.from + " → " + r.period.to + " · " + r.lines + " lines";
  }
  return r.error || "not read";
}
function logImport(file, source, r) {
  var s = loadSettings();
  s.importLog = [{ at: new Date().toISOString(), file: file, source: source, ok: !!r.ok, skipped: !!r.skipped, msg: importMessage(r) }].concat(s.importLog || []).slice(0, 60);
  saveSettings(s);
}
function pdfSkip(name) {
  return /payroll|เงินเดือน/i.test(name) ? "payroll PDF (per person) — not needed, the payroll CSV is used" : "PDF statements (Krungsri / GSB): coming in the next update";
}
var driveSyncing = false;
// full = read every file again (after a fix, or to re-check)
async function driveSync(full) {
  if (driveSyncing) return { ok: false, error: "sync already running" };
  var folders = driveFolders(loadSettings());
  if (!folders.length) return { ok: false, error: "no Drive folder set" };
  driveSyncing = true;
  var results = [];
  try {
    var files = [];
    for (var fi = 0; fi < folders.length; fi++) files = files.concat(await driveListAll(folders[fi]));
    var st = loadStatements();
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      if (!full && st.drive[f.id] === f.modifiedTime) continue; // unchanged since last sync
      var isCsv = /\.csv$/i.test(f.name) || f.mimeType === "text/csv";
      var isXlsx = /\.xlsx$/i.test(f.name) || f.mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
      var r;
      if ((isCsv || isXlsx) && Number(f.size || 0) < 15e6) {
        try { r = await importAny(await driveDownload(f.id), { name: f.name, source: "drive", driveId: f.id, folderMonth: monthFromPath(f.path) }); }
        catch (e) { r = { ok: false, error: e.message }; }
      } else if (/\.pdf$/i.test(f.name) || f.mimeType === "application/pdf") r = { ok: false, skipped: true, error: pdfSkip(f.name) };
      else if (f.mimeType === "application/vnd.google-apps.spreadsheet") r = { ok: false, skipped: true, error: "Google Sheets file — keep the original .xlsx / .csv (turn off Drive's 'Convert uploads')" };
      else r = { ok: false, skipped: true, error: "not a statement / report file" };
      results.push(Object.assign({ file: f.path }, r));
      logImport(f.path, "drive", r);
      st = loadStatements();
      if (r.ok || r.skipped) { st.drive[f.id] = f.modifiedTime; saveStatements(st); }
    }
    // a payroll file deleted from Drive (e.g. put in the wrong month) takes its month's payroll with it
    var alive = {};
    files.forEach(function(f) { alive[f.id] = true; });
    if (fs.existsSync(PAYROLL_DIR)) fs.readdirSync(PAYROLL_DIR).forEach(function(fn) {
      try {
        var pr = JSON.parse(fs.readFileSync(path.join(PAYROLL_DIR, fn), "utf8"));
        if (pr.driveId && !alive[pr.driveId]) { fs.unlinkSync(path.join(PAYROLL_DIR, fn)); results.push({ file: pr.file, ok: true, removed: true, msg: "payroll " + pr.ym + " removed (file no longer in Drive)" }); logImport(pr.file, "drive", { ok: false, skipped: true, error: "payroll " + pr.ym + " removed — the file is no longer in Drive" }); }
      } catch (e) {}
    });
    var s2 = loadSettings();
    s2.driveLastSync = new Date().toISOString();
    saveSettings(s2);
    return { ok: true, checked: files.length, results: results };
  } catch (e) {
    var msg = e.response && e.response.data ? JSON.stringify(e.response.data).slice(0, 300) : e.message;
    console.error("Drive sync:", msg);
    return { ok: false, error: msg };
  } finally { driveSyncing = false; }
}
// folder names like "1026" (MM + YY) say which month the files inside belong to
function monthFromPath(p) {
  var parts = String(p || "").split("/").slice(0, -1).reverse();
  for (var i = 0; i < parts.length; i++) { var m = /^(0[1-9]|1[0-2])(\d{2})$/.exec(parts[i].trim()); if (m) return "20" + m[2] + "-" + m[1]; }
  return null;
}
function driveFolders(s) { return (s.driveFolders || []).concat(s.driveFolderId && (s.driveFolders || []).indexOf(s.driveFolderId) === -1 ? [s.driveFolderId] : []); }
function driveAutoSync() {
  var s = loadSettings();
  if (!driveCreds() || !driveFolders(s).length) return;
  if (s.driveLastSync && Date.now() - Date.parse(s.driveLastSync) < 6 * 3600 * 1000) return;
  driveSync().catch(function() {});
}

// ── POS sales report (xlsx) and petty-cash report (xlsx) ───────────────────
// Only totals are kept (branch × month / day × channel) — no customer names.
const SALES_DIR = path.join(DATA_DIR, "sales");
const PETTY_FILE = path.join(DATA_DIR, "petty.json");
function xv(v) {
  if (v === null || v === undefined) return null;
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") { if (v.result !== undefined) return v.result; if (v.richText) return v.richText.map(function(x) { return x.text; }).join(""); if (v.text !== undefined) return v.text; }
  return v;
}
function xnum(v) { v = xv(v); return typeof v === "number" ? v : (toNum(v) || 0); }
function rowVals(ws, i) { return ws.getRow(i).values.slice(1).map(xv); }
function headerIndex(vals) { var ix = {}; vals.forEach(function(h, c) { if (h !== null && h !== undefined) ix[String(h).trim()] = c; }); return ix; }
// "01/10/2026 10:26" or a real date → "2026-10-01"
function anyDateIso(v) {
  v = xv(v);
  if (!v) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(v)) return String(v).slice(0, 10);
  var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(String(v).trim());
  return m ? m[3] + "-" + m[2] + "-" + m[1] : null;
}
// POS / petty-cash branch names → our branches (same name, contained, or chosen once by the owner)
function mapPosBranch(raw, reg, pnl) {
  raw = String(raw || "").trim();
  var map = pnl.posBranchMap || {};
  if (map[raw] && reg.branches.indexOf(map[raw]) !== -1) return map[raw];
  var l = raw.toLowerCase();
  return reg.branches.find(function(b) { return b.toLowerCase() === l; }) ||
    reg.branches.find(function(b) { var x = b.toLowerCase(); return l.indexOf(x) !== -1 || x.indexOf(l) !== -1; }) || null;
}

function importSalesWb(wb, info) {
  var ws = wb.getWorksheet("รายการขาย");
  var hr = 0;
  for (var i = 1; i <= Math.min(ws.rowCount, 6); i++) if (rowVals(ws, i).indexOf("วันที่เปิดบิล") !== -1) { hr = i; break; }
  if (!hr) return { ok: false, error: "sales report: header row not found" };
  var ix = headerIndex(rowVals(ws, hr));
  var need = ["วันที่เปิดบิล", "สาขา", "ยอดรวม (฿)", "รับรู้รายได้ (฿)"];
  var miss = need.filter(function(h) { return ix[h] === undefined; });
  if (miss.length) return { ok: false, error: "sales report: missing columns " + miss.join(", ") };
  var col = function(v, h) { return ix[h] === undefined ? 0 : xnum(v[ix[h]]); };
  var months = {}, rows = 0;
  for (var r = hr + 1; r <= ws.rowCount; r++) {
    var v = rowVals(ws, r);
    var br = v[ix["สาขา"]], d = anyDateIso(v[ix["วันที่เปิดบิล"]]);
    if (!br || !d) continue; // summary rows at the bottom have no branch
    if (ix["สถานะ"] !== undefined && /void|cancel|ยกเลิก/i.test(String(v[ix["สถานะ"]] || ""))) continue;
    var ym = d.slice(0, 7);
    if (!months[ym]) months[ym] = {};
    var B = months[ym][br] || (months[ym][br] = { bills: 0, gross: 0, discBill: 0, discCard: 0, discPoints: 0, netSales: 0, sales: 0, recognized: 0,
      cash: 0, transfer: 0, card: 0, wallet: 0, online: 0, voucher: 0, walletTopup: 0, deferredAdded: 0, deferredUsed: 0, byDay: {} });
    var g = col(v, "ยอดรวม (฿)"), db = col(v, "ส่วนลดท้ายบิล (฿)"), dc = col(v, "ส่วนลดบัตร (฿)");
    B.bills++; B.gross += g; B.discBill += db; B.discCard += dc; B.discPoints += col(v, "ส่วนลดใช้แต้ม (฿)");
    B.netSales += col(v, "ยอดสุทธิ (฿)"); B.sales += g - db - dc; B.recognized += col(v, "รับรู้รายได้ (฿)");
    var ch = { cash: col(v, "เงินสด (฿)"), transfer: col(v, "โอนเงิน (฿)"), card: col(v, "บัตรเครดิต (฿)") };
    B.cash += ch.cash; B.transfer += ch.transfer; B.card += ch.card; B.wallet += col(v, "Wallet (฿)"); B.online += col(v, "ซื้อออนไลน์ (฿)"); B.voucher += col(v, "Voucher (฿)");
    B.walletTopup += col(v, "เติม Wallet (฿)"); B.deferredAdded += col(v, "ตั้งหนี้สินเพิ่ม (฿)"); B.deferredUsed += col(v, "รับรู้จากหนี้สินเดิม (฿)");
    var day = B.byDay[d] || (B.byDay[d] = { cash: 0, transfer: 0, card: 0 });
    day.cash += ch.cash; day.transfer += ch.transfer; day.card += ch.card;
    rows++;
  }
  if (!rows) return { ok: false, error: "sales report: no bills found" };
  var pm = /(\d{4}-\d{2}-\d{2})_to_(\d{4}-\d{2}-\d{2})/.exec(info.name || "");
  if (!fs.existsSync(SALES_DIR)) fs.mkdirSync(SALES_DIR, { recursive: true });
  var saved = [];
  Object.keys(months).forEach(function(ym) {
    var last = ym + "-" + String(new Date(+ym.slice(0, 4), +ym.slice(5), 0).getDate()).padStart(2, "0");
    var days = [].concat.apply([], Object.values(months[ym]).map(function(b) { return Object.keys(b.byDay); })).sort();
    var full = pm ? (pm[1] <= ym + "-01" && pm[2] >= last) : (days[0] <= ym + "-03" && days[days.length - 1] >= ym + "-" + String(+last.slice(8) - 2).padStart(2, "0"));
    Object.values(months[ym]).forEach(function(b) { Object.keys(b).forEach(function(k) { if (typeof b[k] === "number" && k !== "bills") b[k] = round2(b[k]); }); });
    fs.writeFileSync(path.join(SALES_DIR, ym + ".json"), JSON.stringify({ ym: ym, file: info.name || "", importedAt: new Date().toISOString(), fullMonth: !!full, branches: months[ym] }, null, 2));
    saved.push(ym + (full ? "" : " (part)"));
  });
  return { ok: true, kind: "sales_xlsx", months: saved, bills: rows,
    summary: Object.keys(months).map(function(ym) { return ym + ": " + Object.keys(months[ym]).map(function(b) { return b + " " + baht(months[ym][b].recognized); }).join(", "); }).join(" · ") };
}
function loadSales(ym) { var f = path.join(SALES_DIR, ym + ".json"); return fs.existsSync(f) ? JSON.parse(fs.readFileSync(f, "utf8")) : null; }
// sales of one month keyed by our branch names
function salesByBranch(ym, reg, pnl) {
  var s = loadSales(ym), out = { data: {}, unmapped: [], file: s ? s.file : null, fullMonth: s ? s.fullMonth : false };
  if (!s) return out;
  Object.keys(s.branches).forEach(function(raw) {
    var b = mapPosBranch(raw, reg, pnl);
    if (!b) { out.unmapped.push(raw); return; }
    out.data[b] = s.branches[raw];
  });
  return out;
}

function loadPetty() { var d = loadJsonFile(PETTY_FILE, null) || {}; if (!d.branches) d.branches = {}; return d; }
function savePetty(d) { fs.writeFileSync(PETTY_FILE, JSON.stringify(d)); }
function importPettyWb(wb, info) {
  var ws = wb.getWorksheet("รายการ");
  var head = {};
  for (var i = 1; i <= 7; i++) { var v = rowVals(ws, i); if (v[0]) head[String(v[0]).trim()] = v[1]; }
  var branch = String(head["สาขา"] || "").trim();
  var pm = /(\d{2}\/\d{2}\/\d{4})\s*[–-]\s*(\d{2}\/\d{2}\/\d{4})/.exec(String(head["ช่วงวันที่"] || ""));
  if (!branch || !pm) return { ok: false, error: "petty cash: branch or date range not found at the top of the sheet" };
  var hr = 0;
  for (i = 1; i <= 12; i++) if (rowVals(ws, i).indexOf("ID") !== -1) { hr = i; break; }
  if (!hr) return { ok: false, error: "petty cash: header row not found" };
  var ix = headerIndex(rowVals(ws, hr));
  var entries = [], bal = xnum(head["ยอดยกมา"]);
  for (var r = hr + 1; r <= ws.rowCount; r++) {
    var v = rowVals(ws, r);
    var id = v[ix["ID"]];
    if (id === null || id === undefined || id === "") continue;
    var e = { id: String(id), date: anyDateIso(v[ix["วันที่รายการ"]]) || anyDateIso(v[ix["วันที่บันทึก/อนุมัติ"]]), type: String(v[ix["ประเภท"]] || ""),
      category: String(v[ix["รายการ"]] || "").trim(), detail: String(v[ix["รายละเอียด"]] || "").trim(),
      in: xnum(v[ix["เข้า"]]), out: xnum(v[ix["ออก"]]), other: xnum(v[ix["รายรับอื่น (เข้าลิ้นชัก)"]]), status: String(v[ix["สถานะ"]] || "") };
    var balance = v[ix["ยอดคงเหลือ"]];
    // every row moves the balance as recorded: a cancelled item was deducted when entered and its
    // "กลับรายการ" row adds it back later — only "สำเร็จ" items count as spending (see pettyCosts)
    bal = round2(bal + e.in - e.out);
    if (typeof balance === "number" && Math.abs(bal - balance) > 0.005) return { ok: false, error: "petty cash " + branch + ": balance does not add up at ID " + e.id + " (expected " + bal.toFixed(2) + ", file says " + balance.toFixed(2) + ")" };
    entries.push(e);
  }
  var closing = typeof xv(head["ยอดปิดช่วง"]) === "number" ? xv(head["ยอดปิดช่วง"]) : null;
  if (closing !== null && Math.abs(bal - closing) > 0.005) return { ok: false, error: "petty cash " + branch + ": closing balance " + closing.toFixed(2) + " ≠ " + bal.toFixed(2) + " from the items" };
  var d = loadPetty();
  var B = d.branches[branch] || (d.branches[branch] = { periods: [], entries: {} });
  entries.forEach(function(e) { B.entries[e.id] = e; }); // exports overlap (60-day window) → the ID keeps one copy
  var per = { from: dmyToIso(pm[1]), to: dmyToIso(pm[2]), file: info.name || "", importedAt: new Date().toISOString() };
  B.periods = B.periods.filter(function(p) { return !(p.from === per.from && p.to === per.to); }).concat([per]);
  savePetty(d);
  return { ok: true, kind: "petty_xlsx", branch: branch, period: { from: per.from, to: per.to }, entries: entries.length };
}
function pettyCoverage(raw, ym) {
  var B = loadPetty().branches[raw];
  if (!B) return "none";
  var first = ym + "-01", last = ym + "-" + String(new Date(+ym.slice(0, 4), +ym.slice(5), 0).getDate()).padStart(2, "0");
  // periods may chain (each export covers 60 days): walk from the 1st and see if they reach the last day
  var ps = B.periods.slice().sort(function(a, b) { return a.from.localeCompare(b.from); }), reach = null;
  var dayAfter = function(d) { return new Date(Date.parse(d + "T00:00:00Z") + 86400000).toISOString().slice(0, 10); };
  ps.forEach(function(p) { if (p.from <= first && (reach === null || p.to > reach)) reach = p.to; else if (reach !== null && p.from <= dayAfter(reach) && p.to > reach) reach = p.to; });
  if (reach !== null && reach >= last) return "full";
  return ps.some(function(p) { return p.from <= last && p.to >= first; }) ? "partial" : "none";
}
// petty-cash spending of one month for P&L: [{branch, category (ours), pettyCategory, amount}]
function pettyCosts(ym, reg, pnl) {
  var out = [], d = loadPetty(), map = pnl.pettyCatMap || {};
  Object.keys(d.branches).forEach(function(raw) {
    var b = mapPosBranch(raw, reg, pnl);
    Object.values(d.branches[raw].entries).forEach(function(e) {
      if (e.status !== "สำเร็จ" || !(e.out > 0) || !e.date || e.date.slice(0, 7) !== ym || !/รายจ่าย/.test(e.type)) return;
      out.push({ branch: b, rawBranch: raw, pettyCategory: e.category || "อื่นๆ", category: map[e.category] || PETTY_CAT, amount: e.out });
    });
  });
  return out;
}

// read any uploaded / Drive file: xlsx (sales, petty cash) or CSV (statements, payroll)
async function importAny(buf, info) {
  if (buf[0] === 0x50 && buf[1] === 0x4B) { // "PK" = xlsx
    var wb = new ExcelJS.Workbook();
    try { await wb.xlsx.load(buf); } catch (e) { return { ok: false, error: "cannot open Excel file: " + e.message }; }
    if (wb.getWorksheet("รายการขาย")) return importSalesWb(wb, info);
    var pc = wb.getWorksheet("รายการ");
    if (pc && /petty cash/i.test(String(xv(pc.getCell("A1").value) || ""))) return importPettyWb(wb, info);
    return { ok: false, skipped: true, error: "Excel file is not a sales or petty-cash report" };
  }
  return importStatement(decodeText(buf), info);
}

// ── Report 2 (5th of the month): overview / P&L, once every document is in ──
function opBranchesOf(reg, pnl) { return reg.branches.filter(function(b) { return pnl.hqBranches.indexOf(b) === -1; }); }
function pnlReadiness(y, m) {
  var ym = ymKey(y, m), reg = loadAccounts(), pnl = loadPnl();
  var op = opBranchesOf(reg, pnl);
  var r = { year: y, month: m, ym: ym };
  var sched = loadSchedule().months[ym] || {};
  r.expenseDone = !!sched.sentAt || checkMonth(y, m).ready;
  var S = salesByBranch(ym, reg, pnl);
  r.salesFile = S.file;
  r.salesMissing = S.file ? op.filter(function(b) { return !S.data[b]; }) : op.slice();
  r.salesPartial = !!S.file && !S.fullMonth;
  r.unmappedBranches = S.unmapped.slice();
  var petty = loadPetty();
  Object.keys(petty.branches).forEach(function(raw) { if (!mapPosBranch(raw, reg, pnl) && r.unmappedBranches.indexOf(raw) === -1) r.unmappedBranches.push(raw); });
  r.pettyMissing = op.filter(function(b) {
    var raws = Object.keys(petty.branches).filter(function(raw) { return mapPosBranch(raw, reg, pnl) === b; });
    return !raws.some(function(raw) { return pettyCoverage(raw, ym) === "full"; });
  });
  var pl = payrollPlan(y, m);
  r.payrollMissing = !pl.imported;
  r.payrollNotApplied = !!pl.imported && !pl.applied;
  r.ready = r.expenseDone && !r.salesMissing.length && !r.salesPartial && !r.unmappedBranches.length && !r.pettyMissing.length && !r.payrollMissing && !r.payrollNotApplied;
  return r;
}
function pnlReadinessText(r) {
  if (r.ready) return "✅ Overview " + thMonthRange(r.year, r.month) + " — เอกสารครบแล้ว";
  var L = ["⏳ Overview / กำไร " + thMonthRange(r.year, r.month) + " ยังออกไม่ได้", "ต้องมีให้ครบก่อน:"];
  if (!r.expenseDone) L.push("🧾 รีพอตค่าใช้จ่ายยังไม่ผ่าน (/check)");
  if (r.salesMissing.length) L.push("📊 รีพอตขาย: " + r.salesMissing.join(", "));
  if (r.salesPartial) L.push("📊 รีพอตขายยังไม่ครบทั้งเดือน");
  if (r.pettyMissing.length) L.push("💵 Petty cash: " + r.pettyMissing.join(", "));
  if (r.payrollMissing) L.push("👥 ไฟล์ payroll");
  if (r.payrollNotApplied) L.push("👥 ยังไม่กด Apply แบ่งเงินเดือนตามสาขา (แท็บ P&L)");
  if (r.unmappedBranches.length) L.push("🏢 ชื่อสาขาในรีพอตที่ยังไม่ได้จับคู่: " + r.unmappedBranches.join(", "));
  L.push("", "👉 " + BASE_URL + "/app (ใส่ไฟล์ใน Drive แล้วกด 🔄 หรืออัปโหลด)", "ครบแล้ว overview ออกเองภายใน ~10 นาที", "/overview send = ออกเลยแม้ยังไม่ครบ");
  return L.join("\n");
}
function overviewText(y, m, forced) {
  var P = pnlMonth(y, m);
  var line = function(b) { return "  • " + b.branch + ": รายได้ " + baht(b.revenue) + " · ต้นทุน " + baht(b.variable + b.fixed + b.unset) + " · กำไร " + baht(b.profit) + (b.profitPct !== null ? " (" + (b.profitPct * 100).toFixed(1) + "%)" : ""); };
  return [
    (forced ? "📊" : "📈") + " Overview " + thMonthRange(y, m) + " พร้อมแล้ว" + (forced ? " (ออกก่อนเอกสารครบ)" : ""),
    "รายได้ = รายได้ที่รับรู้ (บริการที่ทำแล้ว)"
  ].concat(P.branches.map(line)).concat([
    "🏦 รวม: รายได้ " + baht(P.company.revenue) + " · กำไร " + baht(P.company.profit) + (P.company.profitPct !== null ? " (" + (P.company.profitPct * 100).toFixed(1) + "%)" : ""),
    "", "📊 Dashboard: " + BASE_URL + "/app (แท็บ 💹 P&L)", "⬇️ Excel: " + reportLink(y, m)
  ]).join("\n");
}
async function sendOverviewReport(to, y, m, forced) {
  var pn = loadPnl(); snapshotAlloc(pn, ymKey(y, m)); savePnl(pn);
  await makeReport(y, m);
  await pushText(to, overviewText(y, m, forced));
}

// ── Month helpers / readiness check ────────────────────────────────────────
var TH_MONTHS = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
function thMonth(y, m) { return TH_MONTHS[m - 1] + " " + y; }
function thMonthRange(y, m) { return thMonth(y, m) + " (1–" + new Date(y, m, 0).getDate() + ")"; }
function prevMonthOf(y, m) { return m === 1 ? { y: y - 1, m: 12 } : { y: y, m: m - 1 }; }
function inMonth(p, y, m) {
  if (!p.transaction_date) return false;
  var d = parseISO(p.transaction_date);
  return d.getFullYear() === y && d.getMonth() + 1 === m;
}
function baht(n) { return "฿" + Number(n || 0).toLocaleString("th-TH", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

function checkMonth(y, m) {
  var reg = syncAccounts();
  var all = loadPayments();
  var rows = all.filter(function(p) { return inMonth(p, y, m); });
  var r = { year: y, month: m, count: rows.length, missingBill: [], noCategory: [], pendingInternal: [], unassigned: [], noDate: [], inbox: loadInbox().length };
  var ua = {};
  rows.forEach(function(p) {
    var c = classify(p, reg);
    if (c.pending) r.pendingInternal.push(showId(p));
    if (c.kind === "internal") return;
    if (c.needsBill && !getBillFiles(p).length && !p.evidence_note && !p.no_bill_expected) r.missingBill.push(showId(p));
    if (!c.category) r.noCategory.push(showId(p));
    var acct = c.kind === "petty_cash" ? c.to : c.from;
    if (acct && !acct.branch) ua[acct.bank + " ****" + acct.last4] = true;
    else if (!acct && c.kind === "expense" && !(p.split && p.split.length)) ua["#" + showId(p) + " (ไม่รู้บัญชีต้นทาง)"] = true;
  });
  r.unassigned = Object.keys(ua);
  r.noDate = all.filter(function(p) { return !p.transaction_date; }).map(showId);
  r.suspectDate = suspectDates(all);
  r.duplicates = [];
  r.recipientUnclear = [];
  rows.forEach(function(p) {
    var d = duplicateOf(p, all);
    if (d) r.duplicates.push({ id: p.id, display_id: showId(p), of: showId(d), date: p.transaction_date, amount: Number(p.amount) || 0, recipient: p.recipient_name || "" });
    if (recipientLooksLikeSender(p, reg)) r.recipientUnclear.push(showId(p));
  });
  // bank statements: money out with no slip, slips the bank never shows, accounts without this month's statement
  var rc = reconMonth(y, m);
  r.stmtNoSlip = rc.noSlip.map(function(l) { return { id: l.id, date: l.date, time: l.time, amount: l.out, account: acctLabelOf(lineAccount(reg, l)), desc: l.desc, counterparty: l.counterparty, candidates: l.candidates }; });
  r.notInStatement = rc.notInStatement.map(function(id) { var p = all.find(function(x) { return x.id === id; }); return p ? showId(p) : String(id); });
  r.missingStatements = rc.missingStatements;
  r.ready = !r.missingBill.length && !r.noCategory.length && !r.pendingInternal.length && !r.unassigned.length && !r.noDate.length && !r.suspectDate.length && !r.duplicates.length && !r.recipientUnclear.length &&
    !r.stmtNoSlip.length && !r.notInStatement.length && !r.missingStatements.length;
  // numeric ids per problem + the records themselves (web page shows picture / details and fixes in place)
  var byShow = {};
  all.forEach(function(p) { byShow[showId(p)] = p; });
  var toIds = function(arr) { return arr.map(function(d) { return byShow[d] ? byShow[d].id : null; }).filter(function(x) { return x !== null; }); };
  r.items = {
    missingBill: toIds(r.missingBill), noCategory: toIds(r.noCategory), pendingInternal: toIds(r.pendingInternal),
    noDate: toIds(r.noDate), recipientUnclear: toIds(r.recipientUnclear), notInStatement: toIds(r.notInStatement),
    duplicates: r.duplicates.map(function(d) { return [d.id, byShow[d.of] ? byShow[d.of].id : null]; })
  };
  var need = {};
  Object.keys(r.items).forEach(function(k) { r.items[k].forEach(function(x) { [].concat(x).forEach(function(id) { if (id !== null) need[id] = true; }); }); });
  r.suspectDate.forEach(function(x) { need[x.id] = true; });
  r.slips = {};
  all.forEach(function(p) { if (need[p.id]) r.slips[p.id] = p; });
  return r;
}

function monthTotals(y, m) {
  var reg = loadAccounts();
  var rows = loadPayments().filter(function(p) { return inMonth(p, y, m); });
  var t = { count: rows.length, expense: 0, internal: 0, internalCount: 0, byBranch: {} };
  rows.forEach(function(p) {
    var c = classify(p, reg);
    if (c.kind === "internal") { t.internal += Number(p.amount) || 0; t.internalCount++; return; }
    c.alloc.forEach(function(a) { t.byBranch[a.branch] = (t.byBranch[a.branch] || 0) + a.amount; t.expense += a.amount; });
  });
  return t;
}

// ── P&L: revenue, fixed / variable costs, HQ cost sharing ─────────────────
// pnl.json = { costTypes: {category: type}, hqBranches: [name], months: { "YYYY-MM": { revenue: {branch: amount}, alloc: {mode, pct} } } }
// Each month keeps its own HQ-sharing setting, so changing it later never rewrites earlier months.
const PNL_FILE = path.join(DATA_DIR, "pnl.json");
var COST_TYPES = ["fixed", "variable", "capex", "unset"];
function loadPnl() {
  var d = loadJsonFile(PNL_FILE, null) || {};
  if (!d.costTypes) d.costTypes = {};
  if (!d.months) d.months = {};
  if (!Array.isArray(d.hqBranches)) {
    d.hqBranches = loadAccounts().branches.filter(function(b) { return /main|\bhq\b|head ?office|ส่วนกลาง|สำนักงานใหญ่/i.test(b); });
  }
  return d;
}
function savePnl(d) { fs.writeFileSync(PNL_FILE, JSON.stringify(d, null, 2)); }
// suggested type for the standard categories; anything else stays "unset" until the owner picks
function defaultCostType(cat) {
  var c = String(cat || "").toLowerCase();
  if (/rent|ค่าเช่า|salary|เงินเดือน|tax|sso|ประกันสังคม|utilit|ค่าน้ำ|ค่าไฟ|accounting|ค่าบัญชี/.test(c)) return "fixed";
  if (/medicine|ยา|doctor|หมอ|nursing|พยาบาล|bonus|โบนัส|commission|ค่าคอม|ค่ามือ|logistic|ขนส่ง/.test(c)) return "variable";
  if (/equipment|อุปกรณ์/.test(c)) return "capex";
  return "unset";
}
function costTypeOf(cat, pnl) {
  if (cat && pnl.costTypes[cat]) return pnl.costTypes[cat];
  return defaultCostType(cat);
}
function costYm(p) { return p.cost_month || (p.transaction_date || "").slice(0, 7); }
// HQ share % per operating branch for one month (revenue mode = by that month's sales)
function allocFor(ym, pnl, opBranches, revenue) {
  var md = pnl.months[ym] || {};
  var a = md.alloc || { mode: "revenue" };
  var pct = {};
  if (a.mode === "custom" && a.pct) {
    opBranches.forEach(function(b) { pct[b] = Number(a.pct[b]) || 0; });
  } else {
    var rev = revenue || md.revenue || {};
    var tot = opBranches.reduce(function(s, b) { return s + (Number(rev[b]) || 0); }, 0);
    var mode = (a.mode === "equal" || tot <= 0) ? "equal" : "revenue";
    opBranches.forEach(function(b) { pct[b] = opBranches.length ? (mode === "equal" ? 100 / opBranches.length : (Number(rev[b]) || 0) / tot * 100) : 0; });
    a = { mode: a.mode === "equal" ? "equal" : "revenue", fallbackEqual: mode === "equal" && a.mode !== "equal" };
  }
  return { mode: a.mode, pct: pct, fallbackEqual: !!a.fallbackEqual, saved: !!md.alloc };
}
// freeze this month's sharing setting (called when revenue / setting for the month is saved)
function snapshotAlloc(pnl, ym) {
  if (!pnl.months[ym]) pnl.months[ym] = {};
  if (pnl.months[ym].alloc) return;
  var prev = Object.keys(pnl.months).filter(function(k) { return k < ym && pnl.months[k].alloc; }).sort().pop();
  var pa = prev ? pnl.months[prev].alloc : null;
  pnl.months[ym].alloc = pa && pa.mode === "custom" ? { mode: "custom", pct: Object.assign({}, pa.pct) } : { mode: pa ? pa.mode : "revenue" };
}

function pnlMonth(y, m) {
  var ym = ymKey(y, m);
  var reg = loadAccounts(), pnl = loadPnl();
  var hq = pnl.hqBranches.filter(function(b) { return reg.branches.indexOf(b) !== -1; });
  var op = reg.branches.filter(function(b) { return hq.indexOf(b) === -1; });
  var manualRev = (pnl.months[ym] || {}).revenue || {};
  var S = salesByBranch(ym, reg, pnl);
  // revenue = recognised revenue from the POS report; typed-in sales only when there is no report
  var rev = {}, revSource = {};
  op.forEach(function(b) {
    if (S.data[b]) { rev[b] = S.data[b].recognized; revSource[b] = "sales report"; }
    else if (Number(manualRev[b]) > 0) { rev[b] = Number(manualRev[b]); revSource[b] = "typed in"; }
  });
  var al = allocFor(ym, pnl, op, rev);
  function blank() { return { variable: 0, fixed: 0, unset: 0, capex: 0, cats: {} }; }
  var direct = {}, hqCost = blank(), unassigned = blank();
  op.forEach(function(b) { direct[b] = blank(); });
  function add(bucket, type, cat, amt) {
    bucket[type] += amt;
    var k = cat || "Uncategorized";
    if (!bucket.cats[k]) bucket.cats[k] = { amount: 0, type: type };
    bucket.cats[k].amount += amt;
  }
  // a branch whose petty-cash report is in: its spending counts from the report, so a bank top-up of its petty cash is just moving money
  var pettyIn = {};
  Object.keys(loadPetty().branches).forEach(function(raw) { var b = mapPosBranch(raw, reg, pnl); if (b && pettyCoverage(raw, ym) !== "none") pettyIn[b] = true; });
  loadPayments().forEach(function(p) {
    if (costYm(p) !== ym) return;
    var c = classify(p, reg);
    if (c.kind === "internal") return;
    if (c.kind === "petty_cash" && pettyIn[c.toBranch]) return;
    var type = (p.cost_type && COST_TYPES.indexOf(p.cost_type) !== -1) ? p.cost_type : costTypeOf(c.category, pnl);
    c.alloc.forEach(function(a) {
      var bucket = direct[a.branch] || (hq.indexOf(a.branch) !== -1 ? hqCost : unassigned);
      add(bucket, type, c.category, a.amount);
    });
  });
  // cash spent from the shop's petty cash (no slips)
  pettyCosts(ym, reg, pnl).forEach(function(e) {
    var bucket = (e.branch && direct[e.branch]) || (e.branch && hq.indexOf(e.branch) !== -1 ? hqCost : unassigned);
    add(bucket, costTypeOf(e.category, pnl), e.category, e.amount);
  });
  var r2 = function(n) { return Math.round(n * 100) / 100; };
  var branches = op.map(function(b) {
    var d = direct[b], share = (al.pct[b] || 0) / 100;
    var t = { variable: d.variable + hqCost.variable * share, fixed: d.fixed + hqCost.fixed * share, unset: d.unset + hqCost.unset * share, capex: d.capex + hqCost.capex * share };
    var cats = {};
    [[d.cats, 1], [hqCost.cats, share]].forEach(function(x) {
      Object.keys(x[0]).forEach(function(k) {
        if (!cats[k]) cats[k] = { amount: 0, type: x[0][k].type };
        cats[k].amount += x[0][k].amount * x[1];
      });
    });
    var revenue = Number(rev[b]) || 0;
    var sd = S.data[b];
    var cm = revenue - t.variable;
    var cmPct = revenue > 0 ? cm / revenue : null;
    var fixedAll = t.fixed + t.unset;
    var profit = cm - fixedAll;
    var be = cmPct && cmPct > 0 ? fixedAll / cmPct : null;
    return {
      branch: b, revenue: revenue, revenueSource: revSource[b] || null,
      billed: sd ? sd.sales : null, received: sd ? { cash: sd.cash, transfer: sd.transfer, card: sd.card, wallet: sd.wallet, online: sd.online } : null,
      deferredAdded: sd ? sd.deferredAdded : null, deferredUsed: sd ? sd.deferredUsed : null, bills: sd ? sd.bills : null,
      hqSharePct: r2(al.pct[b] || 0),
      hqShare: r2((hqCost.variable + hqCost.fixed + hqCost.unset) * share),
      variable: r2(t.variable), fixed: r2(t.fixed), unset: r2(t.unset), capex: r2(t.capex),
      cm: r2(cm), cmPct: cmPct, profit: r2(profit), profitPct: revenue > 0 ? profit / revenue : null,
      breakeven: be === null ? null : r2(be), mos: be !== null && revenue > 0 ? (revenue - be) / revenue : null,
      cats: Object.keys(cats).map(function(k) { return { category: k, type: cats[k].type, amount: r2(cats[k].amount), pctOfRevenue: revenue > 0 ? cats[k].amount / revenue : null }; })
        .sort(function(a, b2) { return b2.amount - a.amount; })
    };
  });
  var sum = function(f) { return r2(branches.reduce(function(s, b) { return s + b[f]; }, 0)); };
  var company = { revenue: sum("revenue"), variable: sum("variable"), fixed: sum("fixed"), unset: sum("unset"), capex: sum("capex"),
    billed: branches.some(function(b) { return b.billed !== null; }) ? r2(branches.reduce(function(s, b) { return s + (b.billed || 0); }, 0)) : null };
  company.profit = r2(company.revenue - company.variable - company.fixed - company.unset);
  company.profitPct = company.revenue > 0 ? company.profit / company.revenue : null;
  var warnings = [];
  var missingRev = op.filter(function(b) { return !(Number(rev[b]) > 0); });
  if (missingRev.length) warnings.push("ยังไม่มีรายได้ (รีพอตขาย): " + missingRev.join(", "));
  if (S.unmapped.length) warnings.push("ชื่อสาขาในรีพอตขายที่ยังไม่ได้จับคู่: " + S.unmapped.join(", "));
  if (!op.length) warnings.push("ยังไม่มีสาขาที่มียอดขาย — ติ๊กสาขาส่วนกลาง (HQ) ใน Dashboard → Settings");
  if (company.unset > 0) warnings.push("มีค่าใช้จ่ายที่ยังไม่ได้เลือก Fixed/Variable " + baht(company.unset) + " (นับเป็น fixed ไปก่อน)");
  var unassignedTotal = unassigned.variable + unassigned.fixed + unassigned.unset + unassigned.capex;
  if (unassignedTotal > 0) warnings.push("ค่าใช้จ่ายที่ยังไม่รู้สาขา " + baht(unassignedTotal) + " ไม่ได้รวมในสาขาใด");
  if (al.fallbackEqual) warnings.push("ยังไม่มียอดขาย จึงหารต้นทุนส่วนกลางเท่า ๆ กันไปก่อน");
  return {
    ym: ym, year: y, month: m, branches: branches, company: company, hqBranches: hq, alloc: al,
    hqCost: { variable: r2(hqCost.variable), fixed: r2(hqCost.fixed), unset: r2(hqCost.unset), capex: r2(hqCost.capex) },
    unassigned: r2(unassignedTotal), warnings: warnings,
    salesFile: S.file,
    caveat: "รายได้ = รายได้ที่รับรู้จากรีพอตขาย · ต้นทุน = สลิป/statement + petty cash — ยังไม่รวมรายการบัตรเครดิตรายตัว"
  };
}
function idList(ids) { return ids.slice(0, 8).map(function(i) { return "#" + i; }).join(" ") + (ids.length > 8 ? " +" + (ids.length - 8) : ""); }
function readinessText(r) {
  if (r.ready) return "✅ " + thMonthRange(r.year, r.month) + " ครบแล้ว — " + r.count + " รายการ" + (r.inbox ? "\n📥 Inbox ค้าง " + r.inbox + " รูป (ไม่บล็อก)" : "");
  var lines = ["⏳ Report ค่าใช้จ่าย " + thMonthRange(r.year, r.month) + " ยังออกไม่ได้", "ต้องเคลียร์ก่อน:"];
  if (r.missingBill.length) lines.push("⚠️ ขาดบิล/note: " + r.missingBill.length + " — " + idList(r.missingBill));
  if (r.noCategory.length) lines.push("🏷 ยังไม่จัดหมวด: " + r.noCategory.length + " — " + idList(r.noCategory));
  if (r.noDate.length) lines.push("📅 ไม่มีวันที่: " + r.noDate.length + " — " + idList(r.noDate));
  if (r.duplicates.length) lines.push("♊ อาจซ้ำ: " + r.duplicates.length + " — " + r.duplicates.slice(0, 6).map(function(d) { return "#" + d.display_id + "=#" + d.of; }).join(" ") + (r.duplicates.length > 6 ? " +" + (r.duplicates.length - 6) : ""));
  if (r.recipientUnclear.length) lines.push("❓ ผู้รับไม่ชัด: " + r.recipientUnclear.length + " — " + idList(r.recipientUnclear));
  if (r.stmtNoSlip.length) lines.push("🏦 มีใน statement แต่ไม่มีสลิป: " + r.stmtNoSlip.length + " รายการ (" + baht(r.stmtNoSlip.reduce(function(s, l) { return s + l.amount; }, 0)) + ")");
  if (r.notInStatement.length) lines.push("🔍 สลิปที่ไม่เจอใน statement: " + r.notInStatement.length + " — " + idList(r.notInStatement));
  if (r.missingStatements.length) lines.push("📄 ยังไม่มี statement: " + r.missingStatements.join(", "));
  if (r.suspectDate.length) lines.push("📅 ปีน่าสงสัย (ห่างจากวันอัปโหลด): " + r.suspectDate.length + " — " + idList(r.suspectDate.map(function(x) { return x.display_id; })));
  if (r.unassigned.length) lines.push("🆕 บัญชียังไม่ระบุสาขา: " + r.unassigned.slice(0, 5).join(", ") + (r.unassigned.length > 5 ? " +" + (r.unassigned.length - 5) : ""));
  if (r.pendingInternal.length) lines.push("🔁 รอยืนยันโอนภายใน: " + r.pendingInternal.length + " — " + idList(r.pendingInternal));
  if (r.inbox) lines.push("📥 Inbox ค้าง " + r.inbox + " รูป (ไม่บล็อก)");
  lines.push("", "👉 " + BASE_URL + "/app", "ทำครบแล้ว report ออกเองภายใน ~10 นาที", "/report send = ออกเลยแม้ยังไม่ครบ");
  return lines.join("\n");
}
function reportReadyText(y, m, forced) {
  var t = monthTotals(y, m);
  var br = Object.entries(t.byBranch).sort(function(a, b) { return b[1] - a[1]; })
    .map(function(e) { return "  • " + e[0] + ": " + baht(e[1]); });
  return [
    (forced ? "📊" : "✅") + " Report ค่าใช้จ่าย " + thMonthRange(y, m) + " พร้อมแล้ว" + (forced ? " (ออกก่อนเคลียร์ครบ)" : ""),
    "📋 " + t.count + " รายการ",
    "💰 ค่าใช้จ่ายจริง: " + baht(t.expense),
    t.internalCount ? "🔁 โอนภายใน (ไม่นับ): " + t.internalCount + " รายการ " + baht(t.internal) : null,
    br.length ? "🏢 แยกสาขา:\n" + br.join("\n") : null,
    "",
    "⬇️ Download Excel:",
    reportLink(y, m),
    "(ลิงก์ใช้ได้ " + LINK_DAYS_REPORT + " วัน · หลังจากนั้นโหลดได้ที่ /app → Report)"
  ].filter(function(l) { return l !== null; }).join("\n");
}

async function makeReport(year, month) {
  const all = loadPayments();
  const rows = all.filter(function(p) {
    if (!p.transaction_date) return false;
    const d = parseISO(p.transaction_date);
    return d.getFullYear() === year && d.getMonth() + 1 === month;
  });
  const wb = new ExcelJS.Workbook();
  const label = format(new Date(year, month - 1), "MMMM yyyy");
  const grandTotal = rows.reduce(function(s,p) { return s+(Number(p.amount)||0); }, 0);
  const baseUrl = BASE_URL;
  const reg = loadAccounts();
  const cls = {};
  rows.forEach(function(p) { cls[p.id] = classify(p, reg); });
  const internalRows = rows.filter(function(p) { return cls[p.id].kind === "internal"; });
  const expenseRows = rows.filter(function(p) { return cls[p.id].kind !== "internal"; });
  const expenseTotal = expenseRows.reduce(function(s,p) { return s+(Number(p.amount)||0); }, 0);
  const internalTotal = grandTotal - expenseTotal;
  function acctLabel(a) { return a ? (a.bank+" ****"+a.last4+(a.label?" "+a.label:"")) : ""; }
  function kindLabel(c) { return c.kind === "internal" ? "🔁 Internal" : (c.kind === "petty_cash" ? "💵 Petty cash" : ""); }
  function branchText(c) { return c.kind === "internal" ? "" : c.alloc.map(function(a){ return c.alloc.length>1 ? a.branch+" "+a.amount.toLocaleString("th-TH") : a.branch; }).join(" / "); }
  var TYPE_LABEL = { company: "Company", petty_cash: "Petty cash", personal: "Personal", external: "External" };

  const s1 = wb.addWorksheet("All Transactions");
  s1.mergeCells("A1:Q1");
  var t1 = s1.getCell("A1");
  t1.value = "Payment Report — " + label;
  t1.font = { size: 14, bold: true, color: { argb: "FFFFFFFF" } };
  t1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1565C0" } };
  t1.alignment = { horizontal: "center" };
  s1.getRow(1).height = 30;
  styleHdr(s1.addRow(["#","Date","Time","Amount (฿)","From Bank","Acct","To Bank","Acct","Recipient","Purpose","Category","Ref","Type","Evidence","Branch","Acct type","Internal"]), "FF0D47A1");
  rows.forEach(function(p, i) {
    var c = cls[p.id];
    var ev = c.kind === "internal" ? "—" : (getBillFiles(p).length?("✅ "+getBillFiles(p).length):(p.evidence_note?"📝":(p.no_bill_expected?"➖":(c.kind==="petty_cash"?"—":"⚠️"))));
    var row = s1.addRow([showId(p),p.transaction_date||"",p.transaction_time||"",Number(p.amount)||0,p.bank_from||"",p.account_from||"",p.bank_to||"",p.account_to||"",p.recipient_name||"",p.purpose||"",c.category||"",p.reference_number||"",(p.slip_type||"").replace(/_/g," "),ev,branchText(c),c.from?TYPE_LABEL[c.from.type]||"":"",kindLabel(c)]);
    if (i%2===0) row.eachCell(function(cc){cc.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF3F8FF"}};});
    if (c.kind === "internal") row.eachCell(function(cc){cc.font={color:{argb:"FF9E9E9E"},italic:true};});
    row.getCell(4).numFmt="#,##0.00"; row.getCell(4).font={bold:true,color:{argb:c.kind==="internal"?"FF9E9E9E":"FF1565C0"}}; row.height=20;
  });
  var tr=s1.addRow(["","TOTAL EXPENSE","",expenseTotal,"(excl. internal transfers)"]);
  tr.getCell(4).numFmt="#,##0.00"; tr.getCell(4).font={bold:true,size:13};
  tr.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
  if (internalRows.length) {
    var tri=s1.addRow(["","Internal (not counted)","",internalTotal,internalRows.length+" transfer(s)"]);
    tri.getCell(4).numFmt="#,##0.00"; tri.eachCell(function(c){c.font={italic:true,color:{argb:"FF757575"}};});
    var trg=s1.addRow(["","All money out","",grandTotal]);
    trg.getCell(4).numFmt="#,##0.00"; trg.eachCell(function(c){c.font={italic:true,color:{argb:"FF757575"}};});
  }
  s1.columns=[{width:9},{width:12},{width:7},{width:14},{width:15},{width:11},{width:15},{width:11},{width:20},{width:28},{width:20},{width:16},{width:15},{width:10},{width:22},{width:12},{width:12}];
  s1.views=[{state:"frozen",ySplit:2}];

  const s2 = wb.addWorksheet("Summary");
  s2.mergeCells("A1:D1");
  var t2=s2.getCell("A1"); t2.value="Summary — "+label;
  t2.font={size:13,bold:true,color:{argb:"FFFFFFFF"}}; t2.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1565C0"}}; t2.alignment={horizontal:"center"}; s2.getRow(1).height=28;
  s2.addRow([]);
  var k1=s2.addRow(["💰 Real expense (excl. internal transfers)",expenseRows.length,expenseTotal,""]);
  k1.getCell(3).numFmt="#,##0.00"; k1.eachCell(function(c){c.font={bold:true,size:12};c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
  var k2=s2.addRow(["🔁 Internal transfers (not counted)",internalRows.length,internalTotal,""]);
  k2.getCell(3).numFmt="#,##0.00"; k2.eachCell(function(c){c.font={italic:true,color:{argb:"FF757575"}};});
  s2.addRow([]);
  function summaryTable(title, map, total) {
    styleHdr(s2.addRow([title,"Transactions","Total (฿)","% of Total"]),"FF0D47A1");
    Object.entries(map).sort(function(a,b){return b[1].sum-a[1].sum;}).forEach(function(e,i){
      var pct=total>0?(e[1].sum/total*100).toFixed(1)+"%":"0%";
      var row=s2.addRow([e[0],e[1].count,e[1].sum,pct]);
      row.getCell(3).numFmt="#,##0.00"; row.getCell(3).font={bold:true,color:{argb:"FF1565C0"}};
      if(i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};});
    });
    return total;
  }
  function totalRow(count, total) {
    var r=s2.addRow(["TOTAL",count,total,"100%"]);
    r.getCell(3).numFmt="#,##0.00"; r.eachCell(function(c){c.font={bold:true};c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
    s2.addRow([]); s2.addRow([]);
  }
  var brMap={};
  expenseRows.forEach(function(p){ cls[p.id].alloc.forEach(function(a){ if(!brMap[a.branch])brMap[a.branch]={count:0,sum:0}; brMap[a.branch].count++; brMap[a.branch].sum+=a.amount; }); });
  summaryTable("Branch (expense)", brMap, expenseTotal); totalRow(expenseRows.length, expenseTotal);
  var purMap={};
  expenseRows.forEach(function(p){var k=cls[p.id].category||"Uncategorized"; if(!purMap[k])purMap[k]={count:0,sum:0}; purMap[k].count++; purMap[k].sum+=Number(p.amount)||0;});
  summaryTable("Category (expense)", purMap, expenseTotal); totalRow(expenseRows.length, expenseTotal);
  var acctMap={};
  rows.forEach(function(p){var k=normalizeBank(p.bank_from)+" ****"+(p.account_from||"????"); if(!acctMap[k])acctMap[k]={count:0,sum:0}; acctMap[k].count++; acctMap[k].sum+=Number(p.amount)||0;});
  summaryTable("Bank Account (all money out, incl. internal)", acctMap, grandTotal); totalRow(rows.length, grandTotal);
  s2.columns=[{width:42},{width:14},{width:18},{width:12}];

  // Branch × Category: which branch spends what
  const s3 = wb.addWorksheet("Branch x Category");
  var brNames = Object.keys(brMap).sort(function(a,b){ return brMap[b].sum-brMap[a].sum; });
  var grid = {};
  expenseRows.forEach(function(p){
    var cat = cls[p.id].category || "Uncategorized";
    if (!grid[cat]) grid[cat] = {};
    cls[p.id].alloc.forEach(function(a){ grid[cat][a.branch] = (grid[cat][a.branch]||0) + a.amount; });
  });
  var lastCol = String.fromCharCode(65 + Math.min(brNames.length + 1, 25));
  s3.mergeCells("A1:" + lastCol + "1");
  var t3=s3.getCell("A1"); t3.value="Branch × Category — "+label+" (excl. internal transfers)";
  t3.font={size:13,bold:true,color:{argb:"FFFFFFFF"}}; t3.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1565C0"}}; t3.alignment={horizontal:"center"}; s3.getRow(1).height=28;
  s3.addRow([]);
  styleHdr(s3.addRow(["Category"].concat(brNames).concat(["Total (฿)"])),"FF0D47A1");
  var cats = Object.keys(grid).sort(function(a,b){ return (purMap[b]?purMap[b].sum:0)-(purMap[a]?purMap[a].sum:0); });
  cats.forEach(function(cat, i){
    var vals = brNames.map(function(b){ return grid[cat][b] || 0; });
    var row = s3.addRow([cat].concat(vals).concat([vals.reduce(function(s,v){return s+v;},0)]));
    for (var ci=2; ci<=brNames.length+2; ci++) row.getCell(ci).numFmt="#,##0.00";
    row.getCell(brNames.length+2).font={bold:true};
    if(i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};});
  });
  var bt = s3.addRow(["TOTAL"].concat(brNames.map(function(b){ return brMap[b].sum; })).concat([expenseTotal]));
  for (var bci=2; bci<=brNames.length+2; bci++) bt.getCell(bci).numFmt="#,##0.00";
  bt.eachCell(function(c){c.font={bold:true};c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
  s3.addRow([]);
  styleHdr(s3.addRow(["% of branch"].concat(brNames).concat(["% of total"])),"FF0D47A1");
  cats.forEach(function(cat){
    var tot = brNames.reduce(function(s,b){ return s+(grid[cat][b]||0); },0);
    s3.addRow([cat].concat(brNames.map(function(b){ return brMap[b].sum>0 ? ((grid[cat][b]||0)/brMap[b].sum*100).toFixed(1)+"%" : "-"; })).concat([expenseTotal>0?(tot/expenseTotal*100).toFixed(1)+"%":"-"]));
  });
  s3.columns=[{width:32}].concat(brNames.map(function(){ return {width:18}; })).concat([{width:18}]);
  s3.views=[{state:"frozen",xSplit:1,ySplit:3}];

  // Internal transfers: shown for audit, not counted
  const s4 = wb.addWorksheet("Internal Transfers");
  s4.mergeCells("A1:I1");
  var t4=s4.getCell("A1"); t4.value="🔁 Internal Transfers (not counted as expense) — "+label;
  t4.font={size:13,bold:true,color:{argb:"FFFFFFFF"}}; t4.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF546E7A"}}; t4.alignment={horizontal:"center"}; s4.getRow(1).height=28;
  s4.addRow([]);
  styleHdr(s4.addRow(["#","Date","Amount (฿)","From account","From branch","To account","To branch","Purpose","Share of expense"]),"FF546E7A");
  internalRows.forEach(function(p){
    var c = cls[p.id];
    var linked = p.linked_to ? all.find(function(q){ return String(q.id)===String(p.linked_to); }) : null;
    var row = s4.addRow([showId(p),p.transaction_date||"",Number(p.amount)||0,acctLabel(c.from)||((p.bank_from||"")+" ****"+(p.account_from||"????")),c.fromBranch||"",acctLabel(c.to)||(p.recipient_name||""),c.toBranch||"",p.purpose||"",linked?("#"+showId(linked)+" ("+(linked.purpose||linked.recipient_name||"")+")"):""]);
    row.getCell(3).numFmt="#,##0.00";
  });
  var it = s4.addRow(["","TOTAL",internalTotal]);
  it.getCell(3).numFmt="#,##0.00"; it.eachCell(function(c){c.font={bold:true};});
  s4.columns=[{width:9},{width:12},{width:14},{width:24},{width:16},{width:24},{width:16},{width:28},{width:30}];

  // P&L by branch (by cost month; HQ costs shared by this month's frozen setting)
  const P = pnlMonth(year, month);
  if (P.branches.length) {
    const s5 = wb.addWorksheet("P&L by Branch");
    var lastP = String.fromCharCode(65 + Math.min(P.branches.length + 1, 25));
    s5.mergeCells("A1:" + lastP + "1");
    var t5 = s5.getCell("A1"); t5.value = "P&L by Branch — " + label;
    t5.font={size:13,bold:true,color:{argb:"FFFFFFFF"}}; t5.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF2E7D32"}}; t5.alignment={horizontal:"center"}; s5.getRow(1).height=28;
    s5.addRow([P.caveat]).getCell(1).font = { italic: true, color: { argb: "FF757575" } };
    styleHdr(s5.addRow([""].concat(P.branches.map(function(b){ return b.branch; })).concat(["Company"])), "FF2E7D32");
    var money = function(lbl, f, comp, bold) {
      var r = s5.addRow([lbl].concat(P.branches.map(function(b){ return b[f]; })).concat([comp]));
      for (var ci = 2; ci <= P.branches.length + 2; ci++) r.getCell(ci).numFmt = "#,##0.00";
      if (bold) r.eachCell(function(c){ c.font = { bold: true }; c.fill = { type:"pattern", pattern:"solid", fgColor:{ argb:"FFF1F8E9" } }; });
    };
    var pctRow = function(lbl, f) {
      s5.addRow([lbl].concat(P.branches.map(function(b){ return b[f] === null ? "-" : (b[f] * 100).toFixed(1) + "%"; })).concat([""]));
    };
    var C = P.company;
    money("Revenue (ยอดขาย)", "revenue", C.revenue, true);
    if (C.billed !== null) {
      money("  Billed sales (after discounts)", "billed", C.billed);
      var recRow = s5.addRow(["  Money received (cash + transfer + card + online)"].concat(P.branches.map(function(b) { return b.received ? round2(b.received.cash + b.received.transfer + b.received.card + b.received.online) : null; })).concat([""]));
      for (var rci = 2; rci <= P.branches.length + 1; rci++) recRow.getCell(rci).numFmt = "#,##0.00";
    }
    money("Variable cost", "variable", C.variable);
    money("Contribution margin", "cm", C.revenue - C.variable, true);
    pctRow("CM %", "cmPct");
    money("Fixed cost", "fixed", C.fixed);
    money("Not classified (counted as fixed)", "unset", C.unset);
    money("  of which HQ share", "hqShare", "");
    pctRow("  HQ share %", "hqSharePct");
    money("Operating profit", "profit", C.profit, true);
    pctRow("Profit %", "profitPct");
    money("Break-even revenue", "breakeven", "");
    pctRow("Margin of safety", "mos");
    money("Investment (capex, not in profit)", "capex", C.capex);
    s5.addRow([]);
    styleHdr(s5.addRow(["Cost as % of revenue"].concat(P.branches.map(function(b){ return b.branch; })).concat(["Type"])), "FF2E7D32");
    var catNames = {};
    P.branches.forEach(function(b){ b.cats.forEach(function(c){ catNames[c.category] = c.type; }); });
    Object.keys(catNames).forEach(function(cat) {
      s5.addRow([cat].concat(P.branches.map(function(b){
        var c = b.cats.find(function(x){ return x.category === cat; });
        return c && c.pctOfRevenue !== null ? (c.pctOfRevenue * 100).toFixed(1) + "%" : (c ? c.amount.toLocaleString("th-TH") : "-");
      })).concat([catNames[cat]]));
    });
    if (P.warnings.length) { s5.addRow([]); P.warnings.forEach(function(w){ s5.addRow(["⚠️ " + w]).getCell(1).font = { color: { argb: "FFBF360C" } }; }); }
    s5.columns = [{ width: 36 }].concat(P.branches.map(function(){ return { width: 18 }; })).concat([{ width: 18 }]);
  }

  var accountGroups={};
  rows.forEach(function(p){
    var nb=normalizeBank(p.bank_from);
    var key=nb+"_"+(p.account_from||"????");
    if(!accountGroups[key])accountGroups[key]={bank:nb,acct:p.account_from||"????",payments:[]};
    accountGroups[key].payments.push(p);
  });

  var accountKeys=Object.keys(accountGroups).sort();
  for(var ai=0;ai<accountKeys.length;ai++){
    var grp=accountGroups[accountKeys[ai]];
    var acctTotal=grp.payments.reduce(function(s,p){return s+(Number(p.amount)||0);},0);
    var color=getBankColor(grp.bank);
    var sheetName=(grp.bank+" ("+grp.acct+")").replace(/[*?:\\/\[\]]/g,"-").substring(0,31);
    var sa=wb.addWorksheet(sheetName);
    sa.mergeCells("A1:F1");
    var ta=sa.getCell("A1"); ta.value=grp.bank+" (****"+grp.acct+") — "+label;
    ta.font={size:13,bold:true,color:{argb:"FFFFFFFF"}}; ta.fill={type:"pattern",pattern:"solid",fgColor:{argb:color}}; ta.alignment={horizontal:"center"}; sa.getRow(1).height=28;
    sa.addRow([]);
    styleHdr(sa.addRow(["Purpose","Count","Total (฿)","% of Account"]),color);
    var apMap={};
    grp.payments.forEach(function(p){var k=p.purpose||"Unspecified"; if(!apMap[k])apMap[k]={count:0,sum:0}; apMap[k].count++; apMap[k].sum+=Number(p.amount)||0;});
    Object.entries(apMap).sort(function(a,b){return b[1].sum-a[1].sum;}).forEach(function(e,i){
      var pct=acctTotal>0?(e[1].sum/acctTotal*100).toFixed(1)+"%":"0%";
      var row=sa.addRow([e[0],e[1].count,e[1].sum,pct]);
      row.getCell(3).numFmt="#,##0.00"; row.getCell(3).font={bold:true};
      if(i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};});
    });
    var atot=sa.addRow(["TOTAL",grp.payments.length,acctTotal,"100%"]);
    atot.getCell(3).numFmt="#,##0.00"; atot.eachCell(function(c){c.font={bold:true};c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
    sa.addRow([]); sa.addRow([]);
    sa.mergeCells("A"+(sa.rowCount)+":F"+(sa.rowCount));
    var evT=sa.getCell("A"+sa.rowCount); evT.value="📎 Payment Evidence";
    evT.font={bold:true,size:12,color:{argb:"FFFFFFFF"}}; evT.fill={type:"pattern",pattern:"solid",fgColor:{argb:color}}; evT.alignment={horizontal:"center"}; sa.getRow(sa.rowCount).height=24;
    sa.addRow([]);
    var maxBills=1;
    grp.payments.forEach(function(p){ var n=getBillFiles(p).length; if(n>maxBills) maxBills=n; });
    var evHdr=["#","Date","Amount (฿)","Purpose","💳 Payment Slip"];
    for(var bi=1;bi<=maxBills;bi++) evHdr.push(maxBills>1 ? ("📄 Bill "+bi) : "📄 Bill / Invoice");
    styleHdr(sa.addRow(evHdr),color);
    var evCols=[{width:9},{width:12},{width:14},{width:30},{width:28}];
    for(bi=0;bi<maxBills;bi++) evCols.push({width:28});
    sa.columns=evCols;
    var ri=sa.rowCount+1;
    for(var pi=0;pi<grp.payments.length;pi++){
      var p=grp.payments[pi];
      var row=sa.getRow(ri); row.height=22;
      row.getCell(1).value="#"+showId(p);
      row.getCell(2).value=p.transaction_date||"";
      row.getCell(3).value=Number(p.amount)||0; row.getCell(3).numFmt="#,##0.00"; row.getCell(3).font={bold:true,color:{argb:"FF1565C0"}};
      row.getCell(4).value=p.purpose||""; row.getCell(4).alignment={wrapText:true,vertical:"middle"};
      if(p.imageFile){
        row.getCell(5).value={text:"🔗 View Slip",hyperlink:signedUrl("/slips/"+p.imageFile, LINK_DAYS_FILES)};
        row.getCell(5).font={color:{argb:"FF1565C0"},underline:true,bold:true};
      } else {
        row.getCell(5).value="No slip image";
        row.getCell(5).font={italic:true,color:{argb:"FF9E9E9E"}};
      }
      var bills=getBillFiles(p);
      for(bi=0;bi<maxBills;bi++){
        var cell=row.getCell(6+bi);
        if(bills[bi]){
          cell.value={text:"🔗 View Bill"+(maxBills>1?(" "+(bi+1)):""),hyperlink:signedUrl("/bills/"+bills[bi], LINK_DAYS_FILES)};
          cell.font={color:{argb:"FF2E7D32"},underline:true,bold:true};
        } else if(bi===0){
          if(p.evidence_note){
            cell.value="📝 "+p.evidence_note;
            cell.font={italic:true,color:{argb:"FF37474F"}};
            cell.alignment={wrapText:true,vertical:"top"};
            var noteLines=Math.ceil(String(p.evidence_note).length/26)+(String(p.evidence_note).split("\n").length-1);
            var nh=Math.max(22,Math.min(140,(noteLines+1)*13));
            if(nh>row.height) row.height=nh;
          } else if(p.no_bill_expected){
            cell.value="➖ No bill expected";
            cell.font={italic:true,color:{argb:"FF9E9E9E"}};
          } else {
            cell.value="⚠️ No bill";
            cell.font={italic:true,color:{argb:"FFBF360C"}};
          }
        }
      }
      for(var ci=1;ci<=5+maxBills;ci++){
        row.getCell(ci).border={bottom:{style:"thin",color:{argb:"FFE0E0E0"}}};
        if(!row.getCell(ci).alignment) row.getCell(ci).alignment={vertical:"middle"};
      }
      if(pi%2===0){for(ci=1;ci<=4;ci++){row.getCell(ci).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};}}
      ri++;
    }
  }

  var outPath=path.join(DATA_DIR,"report_"+year+"_"+String(month).padStart(2,"0")+".xlsx");
  await wb.xlsx.writeFile(outPath);
  return outPath;
}

// ── LINE bot ───────────────────────────────────────────────────────────────
const client = new line.Client(LINE_CONFIG);

// ── Auto monthly report: 1st of month 09:00 (Bangkok), reminders 13/17/21 until ready ──
// State lives on the volume so a redeploy never re-sends or forgets.
var REMIND_SLOTS = [9, 13, 17, 21];
function reportRecipient() { return process.env.REPORT_LINE_ID || loadSettings().reportTo || null; }
function loadSchedule() { var s = loadJsonFile(SCHED_FILE, null) || {}; if (!s.months) s.months = {}; return s; }
function saveSchedule(s) { fs.writeFileSync(SCHED_FILE, JSON.stringify(s, null, 2)); }
function ymKey(y, m) { return y + "-" + String(m).padStart(2, "0"); }
function bkkNow(nowMs) {
  var d = new Date((nowMs || Date.now()) + 7 * 3600 * 1000);
  return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, day: d.getUTCDate(), h: d.getUTCHours(), date: d.toISOString().slice(0, 10) };
}
function remindSlot(h) {
  if (h < REMIND_SLOTS[0] || h >= 22) return null; // quiet 22:00–08:59
  var s = null;
  REMIND_SLOTS.forEach(function(x) { if (h >= x) s = x; });
  return s;
}
function pushText(to, text) { return client.pushMessage(to, { type: "text", text: text }); }

async function sendMonthlyReport(to, y, m, forced) {
  var pn = loadPnl(); snapshotAlloc(pn, ymKey(y, m)); savePnl(pn); // freeze this month's HQ sharing
  await makeReport(y, m); // fail here (not in the user's download) if something is broken
  await pushText(to, reportReadyText(y, m, forced));
}
function markReportSent(y, m, forced) {
  var st = loadSchedule();
  var k = ymKey(y, m);
  st.months[k] = Object.assign(st.months[k] || {}, { sentAt: new Date().toISOString(), forced: !!forced });
  saveSchedule(st);
}

var ticking = false;
// report 1 (1st): expenses match slips / bills / statements · report 2 (5th): overview / P&L once every document is in
var REPORT_JOBS = [
  { state: "months", start: "startReportMonth", day: 1, check: function(y, m) { return checkMonth(y, m); }, text: function(r) { return readinessText(r); },
    send: function(to, y, m) { return sendMonthlyReport(to, y, m, false); } },
  { state: "pnl", start: "pnlStartMonth", day: 5, check: function(y, m) { return pnlReadiness(y, m); }, text: function(r) { return pnlReadinessText(r); },
    send: function(to, y, m) { return sendOverviewReport(to, y, m, false); } }
];
async function scheduleTick(nowMs) {
  if (ticking) return;
  ticking = true;
  try {
    var to = reportRecipient();
    if (!to) return;
    var t = bkkNow(nowMs);
    var cur = ymKey(t.y, t.m);
    var slot = remindSlot(t.h);
    for (var j = 0; j < REPORT_JOBS.length; j++) {
      var job = REPORT_JOBS[j];
      var st = loadSchedule();
      if (!st[job.state]) st[job.state] = {};
      if (!st[job.start]) {
        // first run: don't chase months whose report day already passed before this feature existed
        var nx = t.m === 12 ? { y: t.y + 1, m: 1 } : { y: t.y, m: t.m + 1 };
        st[job.start] = t.day <= job.day ? cur : ymKey(nx.y, nx.m);
        saveSchedule(st);
      }
      var ry = +st[job.start].slice(0, 4), rm = +st[job.start].slice(5);
      while (ymKey(ry, rm) <= cur) {
        var due = ymKey(ry, rm) < cur || t.day >= job.day;
        var tg = prevMonthOf(ry, rm);
        var k = ymKey(tg.y, tg.m);
        var ms = st[job.state][k] || {};
        if (due && !ms.sentAt && (ms.firstAt || slot !== null)) {
          var chk = job.check(tg.y, tg.m);
          if (chk.ready) {
            await job.send(to, tg.y, tg.m);
            ms.sentAt = new Date().toISOString();
          } else if (slot !== null && ms.lastSlot !== t.date + "_" + slot) {
            await pushText(to, job.text(chk));
            ms.lastSlot = t.date + "_" + slot;
            ms.reminders = (ms.reminders || 0) + 1;
          }
          if (!ms.firstAt) ms.firstAt = new Date().toISOString();
          st = loadSchedule();
          if (!st[job.state]) st[job.state] = {};
          st[job.state][k] = ms;
          saveSchedule(st);
        }
        if (rm === 12) { ry++; rm = 1; } else rm++;
      }
    }
  } catch (e) {
    console.error("Schedule tick error:", e.message);
  } finally {
    ticking = false;
  }
}

async function handleEvent(event) {
  if (event.type !== "message") return;
  var uid = event.source.userId;
  var gid = event.source.groupId || event.source.roomId || uid;

  // ═══ IMAGE: simple save mode ═══
  if (event.message.type === "image") {
    var buf = await getImage(client, event.message.id);
    var base64 = buf.toString("base64");
    var type = await detectImageType(base64);

    if (type === "SLIP") {
      await client.replyMessage(event.replyToken, { type: "text", text: "🔍 Reading slip..." });
      var data = await readSlip(base64, "");
      if (!data) {
        // Could not read — keep the picture anyway in inbox
        var item = createInboxItem(buf, "unknown", null, "line");
        return client.pushMessage(gid, { type: "text", text: "⚠️ Could not read slip — image saved to inbox.\nOrganize at " + BASE_URL + "/app" });
      }
      var rec = createSlipRecord(buf, data, "line", event.message.id);
      lastSlip[uid] = { paymentId: rec.id, expiresAt: Date.now() + 5 * 60 * 1000 };
      try {
        var pr = await client.pushMessage(gid, { type: "text", text: slipReply(rec) });
        if (pr && pr.sentMessages && pr.sentMessages[0] && pr.sentMessages[0].id) {
          var lc = loadPayments();
          var li = lc.findIndex(function(p) { return p.id === rec.id; });
          if (li !== -1) { lc[li].line_msg_ids = (lc[li].line_msg_ids || []).concat([pr.sentMessages[0].id]); savePayments(lc); }
        }
      } catch (e) { console.error("confirm push error:", e.message); }
      return;
    }

    // BILL or UNKNOWN → just keep the picture, read amount in background for matching
    var billAmount = null;
    if (type === "BILL") billAmount = await readBillAmount(base64);
    var item = createInboxItem(buf, type === "BILL" ? "bill" : "unknown", billAmount, "line");
    var amtTxt = (billAmount === 0 || billAmount) ? " (฿" + Number(billAmount).toLocaleString("th-TH") + ")" : "";
    return client.replyMessage(event.replyToken, { type: "text", text: "📸 Saved to inbox" + amtTxt + "\nOrganize at " + BASE_URL + "/app" });
  }

  // ═══ TEXT ═══
  if (event.message.type === "text") {
    var txt = event.message.text.trim();
    var tl = txt.toLowerCase();
    var last = getLastSlip(uid);

    // ── Reply to a slip (or its bot confirmation) = evidence note, any time later ──
    var qid = event.message.quotedMessageId;
    if (qid && txt && tl.indexOf("/") !== 0) {
      var listQ = loadPayments();
      var qi = listQ.findIndex(function(p) { return (p.line_msg_ids || []).indexOf(qid) !== -1; });
      if (qi !== -1) {
        var noteTxtQ = txt.replace(/^note\s*:/i, "").trim();
        if (noteTxtQ) {
          listQ[qi].evidence_note = noteTxtQ;
          savePayments(listQ);
          return client.replyMessage(event.replyToken, { type: "text", text: "📝 Evidence note saved for #" + showId(listQ[qi]) });
        }
      } else if (/^note\b/i.test(txt)) {
        return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ Can't link this reply to a recorded slip (reply-link works for slips recorded after this update).\nUse instead: note #261015: your text\n(the # number is in the bot's reply, or /list)" });
      }
      // otherwise fall through — normal chat reply, not a note
    }

    // ── note #id: <text> — attach an evidence note to any slip, any age ──
    var noteIdM = txt.match(/^note\s*#?\s*(\d+)\s*:\s*([\s\S]+)/i);
    if (noteIdM) {
      var listN = loadPayments();
      var ni = listN.findIndex(function(p) { return String(p.display_id) === noteIdM[1] || String(p.id) === noteIdM[1] || (p.old_display_ids || []).indexOf(noteIdM[1]) !== -1; });
      if (ni === -1) return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ No transaction #" + noteIdM[1] + " found. Check /list or the web app." });
      listN[ni].evidence_note = noteIdM[2].trim();
      savePayments(listN);
      return client.replyMessage(event.replyToken, { type: "text", text: "📝 Evidence note saved for #" + showId(listN[ni]) });
    }

    // ── /notify here — this private chat receives the monthly report & reminders ──
    if (tl === "/notify" || tl === "/notify here" || tl === "/notify off") {
      if (process.env.REPORT_LINE_ID) return client.replyMessage(event.replyToken, { type: "text", text: "ℹ️ ผู้รับ report ถูกตั้งไว้ใน env REPORT_LINE_ID แล้ว — แก้ได้ที่ Railway" });
      if (event.source.type !== "user") return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ ใช้ได้เฉพาะแชทส่วนตัวกับ bot (ไม่รับในกลุ่ม เพื่อประหยัดโควตาข้อความ)" });
      var stN = loadSettings();
      if (stN.reportTo && stN.reportTo !== uid) return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ มีผู้รับ report ตั้งไว้แล้ว — เปลี่ยนได้จากแชทของผู้รับเดิมเท่านั้น (/notify off)" });
      if (tl === "/notify off") {
        delete stN.reportTo; saveSettings(stN);
        return client.replyMessage(event.replyToken, { type: "text", text: "🔕 ปิดการส่ง report อัตโนมัติแล้ว" });
      }
      stN.reportTo = uid; saveSettings(stN);
      return client.replyMessage(event.replyToken, { type: "text", text: "🔔 ตั้งแชทนี้เป็นผู้รับ report แล้ว\nทุกวันที่ 1 เวลา 09:00 จะเช็กเดือนก่อน — ถ้ายังไม่ครบจะเตือน 13:00 / 17:00 / 21:00 จนครบ\n/check = เช็กตอนนี้" });
    }

    // ── /check [YYYY MM] — readiness of last month (default) ──
    if (tl.indexOf("/check") === 0) {
      var cp = txt.split(/\s+/);
      var pm = prevMonthOf(bkkNow().y, bkkNow().m);
      var cy = parseInt(cp[1]) || pm.y, cm = parseInt(cp[2]) || pm.m;
      var chkR = checkMonth(cy, cm);
      await client.replyMessage(event.replyToken, { type: "text", text: readinessText(chkR) });
      if (chkR.ready && uid === reportRecipient()) scheduleTick().catch(function() {});
      return;
    }

    // ── /report send [YYYY MM] — release last month's report now, even if not complete ──
    if (tl.indexOf("/report send") === 0) {
      var sp = txt.split(/\s+/);
      var spm = prevMonthOf(bkkNow().y, bkkNow().m);
      var sy = parseInt(sp[2]) || spm.y, sm = parseInt(sp[3]) || spm.m;
      var chkS = checkMonth(sy, sm);
      try {
        await sendMonthlyReport(gid, sy, sm, !chkS.ready);
        markReportSent(sy, sm, !chkS.ready);
      } catch (err) {
        console.error("Report send error:", err.message);
        return client.pushMessage(gid, { type: "text", text: "❌ Report error: " + err.message });
      }
      return;
    }

    if (tl === "/help" || tl === "help") {
      return client.replyMessage(event.replyToken, {
        type: "text",
        text: [
          "💳 Slip Tracker Bot",
          "━━━━━━━━━━━━━━━━━━",
          "📸 Send slip → auto-read & recorded",
          "📸 Send bill/other → saved to inbox",
          "💬 Text after slip → sets purpose",
          "📝 Reply to a slip → saves evidence note",
          "📝 note #261015: text → note for any slip",
          "",
          "🌐 Web App (upload, match, edit):",
          BASE_URL + "/app",
          "",
          "/summary — this month totals",
          "/list — last 5 payments",
          "/report — Excel for this month",
          "/report YYYY MM — specific month",
          "",
          "🗓 Auto report (1st of month):",
          "/notify here — send it to this chat",
          "/check — is last month ready?",
          "/report send — release last month now",
          "/overview — is the 5th-of-month overview ready?",
          "/overview send — release the overview now",
          "/revenue 2026 10 Samui 2010869 — sales (if no sales report)",
          "/help — this menu"
        ].join("\n")
      });
    }

    if (tl === "/summary" || tl === "summary") {
      var list = loadPayments(); var now = new Date();
      var thisMonth = list.filter(function(p) { if (!p.transaction_date) return false; var d = parseISO(p.transaction_date); return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth(); });
      var regS = loadAccounts();
      var total = thisMonth.reduce(function(s,p) { return classify(p, regS).kind === "internal" ? s : s+(Number(p.amount)||0); }, 0);
      var withBill = thisMonth.filter(function(p) { return getBillFiles(p).length > 0; }).length;
      var inboxCount = loadInbox().length;
      var aMap = {}; thisMonth.forEach(function(p) { var k=normalizeBank(p.bank_from)+" ****"+(p.account_from||"????"); aMap[k]=(aMap[k]||0)+(Number(p.amount)||0); });
      var aLines = Object.entries(aMap).map(function(e) { return "  • "+e[0]+": ฿"+e[1].toLocaleString("th-TH",{minimumFractionDigits:2}); }).join("\n");
      return client.replyMessage(event.replyToken, { type: "text", text: ["📊 "+format(now,"MMMM yyyy")+" Summary","━━━━━━━━━━━━━━━━━━","📋 Transactions: "+thisMonth.length,"💰 Expense (excl. internal): ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2}),"📎 With bill: "+withBill+"/"+thisMonth.length,"📥 Inbox (unorganized): "+inboxCount,"","By Account:",aLines||"  (none yet)","","Organize: "+BASE_URL+"/app"].join("\n") });
    }

    if (tl === "/list" || tl === "list") {
      var list = loadPayments();
      if (!list.length) return client.replyMessage(event.replyToken, { type: "text", text: "No payments recorded yet." });
      var recent = list.slice(-5).reverse();
      var lines = recent.map(function(p) { return ["#"+showId(p)+"  "+(p.transaction_date||"?")+"  ฿"+Number(p.amount||0).toLocaleString("th-TH"),"  "+(p.bank_from||"?")+" ****"+(p.account_from||"????"),"  "+(p.purpose||"-")+"  "+(getBillFiles(p).length?"📎✅":"⚠️")].join("\n"); });
      return client.replyMessage(event.replyToken, { type: "text", text: "📋 Last 5:\n\n" + lines.join("\n\n") });
    }

    if (tl.indexOf("/report") === 0 || tl === "report") {
      var parts = txt.split(" "); var now = new Date();
      var year = parseInt(parts[1]) || now.getFullYear();
      var month = parseInt(parts[2]) || now.getMonth() + 1;
      await client.replyMessage(event.replyToken, { type: "text", text: "📊 Generating report..." });
      try {
        await makeReport(year, month);
        var list = loadPayments();
        var filtered = list.filter(function(p) { if (!p.transaction_date) return false; var d = parseISO(p.transaction_date); return d.getFullYear() === year && d.getMonth()+1 === month; });
        var total = monthTotals(year, month).expense;
        var withBill = filtered.filter(function(p){return getBillFiles(p).length > 0;}).length;
        var accts = {}; filtered.forEach(function(p){accts[normalizeBank(p.bank_from)+" ****"+(p.account_from||"????")] = true;});
        return client.pushMessage(gid, { type: "text", text: ["✅ Report ready — "+format(new Date(year,month-1),"MMMM yyyy"),"📋 "+filtered.length+" transactions","💰 ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2})+" (excl. internal)","🏦 "+Object.keys(accts).length+" account sheet(s)","📎 Bill evidence: "+withBill+"/"+filtered.length,"","⬇️ Download Excel (link valid "+LINK_DAYS_REPORT+" days):",reportLink(year, month)].join("\n") });
      } catch(err) {
        console.error("Report error:", err.message);
        return client.pushMessage(gid, { type: "text", text: "❌ Report error: " + err.message });
      }
    }

    // ── /overview [send] [YYYY MM] — report 2 (P&L) readiness, or release it now ──
    if (tl.indexOf("/overview") === 0) {
      var op2 = txt.split(/\s+/), force = /^send$/i.test(op2[1] || "");
      var opm = prevMonthOf(bkkNow().y, bkkNow().m);
      var oy = parseInt(op2[force ? 2 : 1]) || opm.y, om = parseInt(op2[force ? 3 : 2]) || opm.m;
      var orr = pnlReadiness(oy, om);
      if (!force) {
        await client.replyMessage(event.replyToken, { type: "text", text: pnlReadinessText(orr) });
        if (orr.ready && uid === reportRecipient()) scheduleTick().catch(function() {});
        return;
      }
      try {
        await sendOverviewReport(gid, oy, om, !orr.ready);
        var sto = loadSchedule(); if (!sto.pnl) sto.pnl = {};
        sto.pnl[ymKey(oy, om)] = Object.assign(sto.pnl[ymKey(oy, om)] || {}, { sentAt: new Date().toISOString(), forced: !orr.ready });
        saveSchedule(sto);
      } catch (err) { return client.pushMessage(gid, { type: "text", text: "❌ Overview error: " + err.message }); }
      return;
    }

    // ── /revenue YYYY MM <branch> <amount> — monthly sales for the P&L ──
    if (tl.indexOf("/revenue") === 0) {
      var rv = txt.match(/^\/revenue\s+(\d{4})\s+(\d{1,2})\s+(.+?)\s+([\d,\.]+)\s*$/i);
      var brs = loadAccounts().branches;
      if (!rv) return client.replyMessage(event.replyToken, { type: "text", text: "พิมพ์: /revenue 2026 10 Samui 2010869\nสาขา: " + (brs.join(", ") || "(ยังไม่มี)") });
      var bName = brs.find(function(b) { return b.toLowerCase() === rv[3].trim().toLowerCase(); });
      if (!bName) return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ ไม่รู้จักสาขา \"" + rv[3] + "\" — มี: " + brs.join(", ") });
      var rym = ymKey(+rv[1], +rv[2]);
      var pn = loadPnl();
      snapshotAlloc(pn, rym);
      if (!pn.months[rym].revenue) pn.months[rym].revenue = {};
      pn.months[rym].revenue[bName] = Number(rv[4].replace(/,/g, ""));
      savePnl(pn);
      return client.replyMessage(event.replyToken, { type: "text", text: "💰 ยอดขาย " + bName + " " + thMonth(+rv[1], +rv[2]) + ": " + baht(pn.months[rym].revenue[bName]) });
    }

    if (tl.charAt(0) === "/") {
      return client.replyMessage(event.replyToken, { type: "text", text: "❓ ไม่รู้จักคำสั่ง " + txt.split(/\s+/)[0] + " — พิมพ์ /help ดูรายการคำสั่ง" });
    }

    // Text after a slip: "note:" = evidence note, plain = purpose (window stays open 5 min)
    if (last) {
      var list = loadPayments();
      var idx = list.findIndex(function(p) { return p.id === last.paymentId; });
      if (idx === -1) { delete lastSlip[uid]; return; }
      var pid = showId(list[idx]);
      if (/^note\s*:/i.test(txt)) {
        var noteTxt = txt.replace(/^note\s*:/i, "").trim();
        if (noteTxt) {
          list[idx].evidence_note = noteTxt;
          savePayments(list);
          return client.replyMessage(event.replyToken, { type: "text", text: "📝 Evidence note saved for #" + pid });
        }
        return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ Empty note — type: note: <your text>" });
      }
      list[idx].purpose = txt;
      savePayments(list);
      return client.replyMessage(event.replyToken, { type: "text", text: "✅ Purpose saved for #" + pid + ": " + txt + "\n(You can also send: note: <evidence text>)" });
    }
  }
}

// ── Web auth ───────────────────────────────────────────────────────────────
// Sessions live on the volume so a redeploy doesn't log everyone out.
// The token also goes in an HttpOnly cookie so <img> tags and downloads (which can't send headers) are authorised.
const SESS_FILE = path.join(DATA_DIR, "sessions.json");
const SESSION_MS = 12 * 3600 * 1000;
const sessions = loadJsonFile(SESS_FILE, null) || {};
var sessSaveTimer = null;
function saveSessionsSoon() {
  if (sessSaveTimer) return;
  sessSaveTimer = setTimeout(function() {
    sessSaveTimer = null;
    var now = Date.now();
    Object.keys(sessions).forEach(function(k) { if (sessions[k] < now) delete sessions[k]; });
    try { fs.writeFileSync(SESS_FILE, JSON.stringify(sessions)); } catch (e) { console.error("Session save:", e.message); }
  }, 2000);
}
function cookieToken(req) {
  var m = /(?:^|;\s*)st_auth=([a-f0-9]+)/.exec(req.headers.cookie || "");
  return m ? m[1] : null;
}
function sessionOk(req) {
  var t = req.headers["x-auth"] || cookieToken(req);
  if (!t || !sessions[t] || sessions[t] < Date.now()) return null;
  sessions[t] = Date.now() + SESSION_MS;
  saveSessionsSoon();
  return t;
}
function setAuthCookie(req, res, t) {
  var secure = req.headers["x-forwarded-proto"] === "https" ? "; Secure" : "";
  res.setHeader("Set-Cookie", "st_auth=" + (t || "") + "; Path=/; HttpOnly; SameSite=Lax" + secure + "; Max-Age=" + (t ? SESSION_MS / 1000 : 0));
}
function auth(req, res, next) {
  if (sessionOk(req)) return next();
  res.status(401).json({ error: "unauthorized" });
}

// ── Signed, expiring links for LINE messages and Excel files (no login needed while valid) ──
function linkSecret() {
  if (process.env.LINK_SECRET) return process.env.LINK_SECRET;
  var st = loadSettings();
  if (!st.linkSecret) { st.linkSecret = crypto.randomBytes(32).toString("hex"); saveSettings(st); }
  return st.linkSecret;
}
function linkSig(pathname, year, month, exp) {
  return crypto.createHmac("sha256", linkSecret()).update(pathname + "|" + (year || "") + "|" + (month || "") + "|" + exp).digest("hex").slice(0, 32);
}
var LINK_DAYS_REPORT = 7, LINK_DAYS_FILES = 90;
function signedUrl(pathname, days, year, month) {
  var exp = Math.floor(Date.now() / 1000) + days * 86400;
  var q = (year ? "year=" + year + "&month=" + month + "&" : "") + "exp=" + exp + "&sig=" + linkSig(pathname, year, month, exp);
  return BASE_URL + pathname + "?" + q;
}
function reportLink(y, m) { return signedUrl("/api/report", LINK_DAYS_REPORT, y, m); }
function signatureOk(req) {
  var exp = parseInt(req.query.exp), sig = String(req.query.sig || "");
  if (!exp || !sig || exp * 1000 < Date.now()) return false;
  var want = linkSig(req.originalUrl.split("?")[0], req.query.year, req.query.month, exp);
  return sig.length === want.length && crypto.timingSafeEqual(Buffer.from(sig), Buffer.from(want));
}
function authOrSigned(req, res, next) {
  if (signatureOk(req) || sessionOk(req)) return next();
  res.status(401).send("🔒 This link has expired or needs login. Open " + BASE_URL + "/app (log in), then download again.");
}

app.post("/api/login", function(req, res) {
  if (!WEB_PASSWORD) return res.status(500).json({ error: "WEB_PASSWORD not set on server" });
  if ((req.body && req.body.password) === WEB_PASSWORD) {
    var token = crypto.randomBytes(24).toString("hex");
    sessions[token] = Date.now() + SESSION_MS;
    saveSessionsSoon();
    setAuthCookie(req, res, token);
    return res.json({ token: token });
  }
  res.status(401).json({ error: "wrong password" });
});
// re-issue the cookie for a browser that is already logged in (header token)
app.get("/api/session", function(req, res) {
  var t = sessionOk(req);
  if (!t) return res.status(401).json({ error: "unauthorized" });
  setAuthCookie(req, res, t);
  res.json({ ok: true });
});
app.post("/api/logout", function(req, res) {
  var t = req.headers["x-auth"] || cookieToken(req);
  if (t) { delete sessions[t]; saveSessionsSoon(); }
  setAuthCookie(req, res, null);
  res.json({ ok: true });
});

// ── Web API ────────────────────────────────────────────────────────────────
app.get("/api/inbox", auth, function(req, res) { res.json(loadInbox()); });

app.get("/api/slips", auth, function(req, res) {
  var list = loadPayments();
  var year = parseInt(req.query.year), month = parseInt(req.query.month);
  if (year && month) {
    list = list.filter(function(p) {
      if (!p.transaction_date) return req.query.nodate === "1";
      var d = parseISO(p.transaction_date);
      return d.getFullYear() === year && d.getMonth()+1 === month;
    });
  }
  // _cls = branch / internal classification for display (computed, never stored)
  var reg = syncAccounts();
  var all = loadPayments();
  res.json(list.map(function(p) {
    var c = classify(p, reg);
    return Object.assign({}, p, { _cls: {
      kind: c.kind, via: c.via, fromBranch: c.fromBranch, toBranch: c.toBranch, category: c.category,
      alloc: c.alloc, pending: c.pending, needsBill: c.needsBill,
      newAcct: !!(c.from && !c.from.branch),
      noFromAcct: c.kind !== "internal" && !last4(p.account_from),
      suggestions: c.kind === "internal" ? linkSuggestions(p, all, reg) : [],
      duplicateOf: (function() { var d = duplicateOf(p, all); return d ? showId(d) : null; })(),
      toSimilar: similarToAccount(p, all),
      recipientUnclear: recipientLooksLikeSender(p, reg),
      linkedDisplay: p.linked_to ? (function() { var e = all.find(function(q) { return String(q.id) === String(p.linked_to); }); return e ? showId(e) : null; })() : null
    } });
  }));
});

// ── Bank statements & payroll ──
// upload one file (CSV) from the web app: { name, data: base64, month?: "YYYY-MM" for payroll }
app.post("/api/statements/upload", auth, function(req, res) {
  try {
    var b = req.body || {};
    var buf = Buffer.from(String(b.data || "").replace(/^data:[^,]*,/, ""), "base64");
    if (!buf.length) return res.status(400).json({ error: "empty file" });
    if (/\.pdf$/i.test(b.name || "")) { var sk = { ok: false, skipped: true, error: pdfSkip(b.name) }; logImport(b.name, "upload", sk); return res.json(sk); }
    importAny(buf, { name: b.name, source: "upload", month: b.month }).then(function(r) { logImport(b.name, "upload", r); res.json(r); })
      .catch(function(e) { logImport(b.name, "upload", { ok: false, error: e.message }); res.status(500).json({ error: e.message }); });
  } catch (e) { res.status(500).json({ error: e.message }); }
});
app.get("/api/statements", auth, function(req, res) {
  var pm = prevMonthOf(bkkNow().y, bkkNow().m);
  var y = parseInt(req.query.year) || pm.y, m = parseInt(req.query.month) || pm.m;
  var rc = reconMonth(y, m), st = loadStatements();
  res.json({ recon: { ym: rc.ym, accounts: rc.accounts, counts: rc.counts, missingStatements: rc.missingStatements },
    files: Object.values(st.files).sort(function(a, b) { return (b.importedAt || "").localeCompare(a.importedAt || ""); }).slice(0, 40) });
});
// money-out line without a slip: record it (creates a transaction), point it to a slip, or set it aside with a note
app.post("/api/statements/line/:id", auth, function(req, res) {
  var st = loadStatements();
  var l = st.lines.find(function(x) { return x.id === req.params.id; });
  if (!l) return res.status(404).json({ error: "line not found" });
  var a = req.body.action;
  if (a === "record") {
    var reg = loadAccounts(), acct = lineAccount(reg, l), cp = l.counterparty || {};
    var list = loadPayments();
    var rec = { id: nextNumericId(), display_id: nextDisplayId(l.date), imageFile: null, billFile: null, savedAt: new Date().toISOString(), source: "statement", stmt_line: l.id,
      bank_from: acct ? acct.bank : l.bank, account_from: l.account, transaction_date: l.date, transaction_time: l.time, amount: l.out,
      recipient_name: cp.name || l.desc, bank_to: cp.bill ? null : (cp.bank || null), account_to: cp.digits || null, purpose: null, category: null };
    var vm = vendorCategoryMap()[vendorKey(rec.recipient_name)];
    if (vm) rec.category = vm;
    list.push(rec);
    savePayments(list);
    l.resolved = { how: "recorded", slipId: rec.id, at: new Date().toISOString() };
    saveStatements(st);
    return res.json({ ok: true, record: rec });
  }
  if (a === "match") {
    var p = loadPayments().find(function(x) { return String(x.id) === String(req.body.slipId); });
    if (!p) return res.status(404).json({ error: "slip not found" });
    l.slipId = p.id;
    delete l.resolved;
    saveStatements(st);
    return res.json({ ok: true });
  }
  if (a === "ignore") {
    var note = String(req.body.note || "").trim();
    if (!note) return res.status(400).json({ error: "say why (e.g. refund, moved between own accounts)" });
    l.resolved = { how: "ignored", note: note, at: new Date().toISOString() };
    saveStatements(st);
    return res.json({ ok: true });
  }
  if (a === "reset") { delete l.slipId; delete l.resolved; saveStatements(st); return res.json({ ok: true }); }
  res.status(400).json({ error: "unknown action" });
});
app.get("/api/payroll", auth, function(req, res) {
  var pm = prevMonthOf(bkkNow().y, bkkNow().m);
  res.json(payrollPlan(parseInt(req.query.year) || pm.y, parseInt(req.query.month) || pm.m));
});
app.post("/api/payroll/apply", auth, function(req, res) { res.json(applyPayroll(parseInt(req.body.year), parseInt(req.body.month))); });
app.post("/api/payroll/branch-map", auth, function(req, res) {
  var pnl = loadPnl();
  var key = req.body.kind === "pos" ? "posBranchMap" : "payrollBranchMap"; // pos = sales & petty-cash reports
  if (!pnl[key]) pnl[key] = {};
  if (req.body.branch) pnl[key][String(req.body.raw)] = String(req.body.branch); else delete pnl[key][String(req.body.raw)];
  savePnl(pnl);
  res.json({ ok: true });
});
app.get("/api/drive", auth, function(req, res) {
  var c = driveCreds(), s = loadSettings();
  res.json({ configured: !!(c && c.client_email), email: c ? c.client_email : null, configError: driveConfigError(), folders: driveFolders(s), lastSync: s.driveLastSync || null, log: s.importLog || [] });
});
// several folders (e.g. sales / petty cash, and statement/2026/MMYY); remove: true takes one out
app.post("/api/drive/folder", auth, function(req, res) {
  var id = driveFolderId(req.body.link);
  if (!id) return res.status(400).json({ error: "paste the folder link from Google Drive" });
  var s = loadSettings();
  var list = driveFolders(s).filter(function(x) { return x !== id; });
  if (!req.body.remove) list.push(id);
  s.driveFolders = list; delete s.driveFolderId; saveSettings(s);
  res.json({ ok: true, folders: list });
});
app.get("/api/overview/check", auth, function(req, res) {
  var pm = prevMonthOf(bkkNow().y, bkkNow().m);
  var r = pnlReadiness(parseInt(req.query.year) || pm.y, parseInt(req.query.month) || pm.m);
  r.sentAt = ((loadSchedule().pnl || {})[r.ym] || {}).sentAt || null;
  res.json(r);
});
// petty-cash categories seen so far → which of our categories they count as
app.get("/api/petty/categories", auth, function(req, res) {
  var d = loadPetty(), pnl = loadPnl(), tot = {};
  Object.values(d.branches).forEach(function(B) { Object.values(B.entries).forEach(function(e) { if (e.status === "สำเร็จ" && e.out > 0 && /รายจ่าย/.test(e.type)) tot[e.category] = (tot[e.category] || 0) + e.out; }); });
  res.json({ categories: Object.keys(tot).sort(function(a, b) { return tot[b] - tot[a]; }).map(function(c) { return { category: c, total: round2(tot[c]), mapped: (pnl.pettyCatMap || {})[c] || null }; }),
    branches: Object.keys(d.branches).map(function(raw) { var B = d.branches[raw]; return { name: raw, periods: B.periods.map(function(p) { return p.from + " → " + p.to; }) }; }) });
});
app.post("/api/drive/sync", auth, async function(req, res) { res.json(await driveSync(!!(req.body && req.body.full))); });
// remove one month's payroll (wrong file)
app.post("/api/payroll/remove", auth, function(req, res) {
  var f = path.join(PAYROLL_DIR, ymKey(parseInt(req.body.year), parseInt(req.body.month)) + ".json");
  if (fs.existsSync(f)) fs.unlinkSync(f);
  res.json({ ok: true });
});

// ── P&L dashboard ──
app.get("/api/pnl", auth, function(req, res) {
  var pm = prevMonthOf(bkkNow().y, bkkNow().m);
  var y = parseInt(req.query.year) || pm.y, m = parseInt(req.query.month) || pm.m;
  var cur = pnlMonth(y, m);
  var trend = [];
  var ty = y, tm = m;
  for (var i = 0; i < 6; i++) {
    var P = i === 0 ? cur : pnlMonth(ty, tm);
    trend.unshift({ ym: P.ym, company: P.company, branches: P.branches.map(function(b) { return { branch: b.branch, revenue: b.revenue, cost: b.variable + b.fixed + b.unset, profit: b.profit }; }) });
    var pv = prevMonthOf(ty, tm); ty = pv.y; tm = pv.m;
  }
  var pnl = loadPnl();
  res.json({ pnl: cur, trend: trend, revenue: (pnl.months[cur.ym] || {}).revenue || {}, salesFile: cur.salesFile });
});
app.get("/api/pnl/settings", auth, function(req, res) {
  var pnl = loadPnl();
  var types = {};
  loadCategories().forEach(function(c) { types[c] = costTypeOf(c, pnl); });
  res.json({ costTypes: types, hqBranches: pnl.hqBranches, branches: loadAccounts().branches });
});
app.post("/api/pnl/settings", auth, function(req, res) {
  var pnl = loadPnl();
  var b = req.body || {};
  if (b.costTypes && typeof b.costTypes === "object") {
    Object.keys(b.costTypes).forEach(function(c) { if (COST_TYPES.indexOf(b.costTypes[c]) !== -1) pnl.costTypes[c] = b.costTypes[c]; });
  }
  if (b.pettyCatMap && typeof b.pettyCatMap === "object") {
    if (!pnl.pettyCatMap) pnl.pettyCatMap = {};
    var cats0 = loadCategories();
    Object.keys(b.pettyCatMap).forEach(function(k) { var v = b.pettyCatMap[k]; if (!v) delete pnl.pettyCatMap[k]; else if (cats0.indexOf(v) !== -1 || v === PETTY_CAT) pnl.pettyCatMap[k] = v; });
  }
  if (Array.isArray(b.hqBranches)) {
    var brs = loadAccounts().branches;
    pnl.hqBranches = b.hqBranches.filter(function(x) { return brs.indexOf(x) !== -1; });
  }
  savePnl(pnl);
  res.json({ ok: true });
});
// one month: sales per branch and how HQ costs are shared (frozen for that month)
app.post("/api/pnl/month", auth, function(req, res) {
  var y = parseInt(req.body.year), m = parseInt(req.body.month);
  if (!y || !m) return res.status(400).json({ error: "year/month required" });
  var ym = ymKey(y, m);
  var pnl = loadPnl();
  snapshotAlloc(pnl, ym);
  var brs = loadAccounts().branches;
  if (req.body.revenue && typeof req.body.revenue === "object") {
    var rv = {};
    Object.keys(req.body.revenue).forEach(function(k) {
      var v = Number(String(req.body.revenue[k]).replace(/,/g, ""));
      if (brs.indexOf(k) !== -1 && v > 0) rv[k] = v;
    });
    pnl.months[ym].revenue = rv;
  }
  var a = req.body.alloc;
  if (a && ["revenue", "equal", "custom"].indexOf(a.mode) !== -1) {
    if (a.mode === "custom") {
      var pct = {}, tot = 0;
      Object.keys(a.pct || {}).forEach(function(k) { if (brs.indexOf(k) !== -1) { pct[k] = Number(a.pct[k]) || 0; tot += pct[k]; } });
      if (Math.abs(tot - 100) > 0.5) return res.status(400).json({ error: "custom % must add up to 100 (now " + tot.toFixed(1) + ")" });
      pnl.months[ym].alloc = { mode: "custom", pct: pct };
    } else pnl.months[ym].alloc = { mode: a.mode };
  }
  savePnl(pnl);
  res.json({ ok: true, month: pnl.months[ym] });
});

// ── Accounts / branches / aliases ──
app.get("/api/accounts", auth, function(req, res) {
  var reg = syncAccounts();
  var counts = {};
  loadPayments().forEach(function(p) { var a = findAccount(reg, p.bank_from, p.account_from); if (a) counts[a.bank + "_" + a.last4] = (counts[a.bank + "_" + a.last4] || 0) + 1; });
  res.json(Object.assign({}, reg, { accounts: reg.accounts.map(function(a) {
    var sim = similarAccount(reg, a);
    return Object.assign({}, a, { _slips: counts[a.bank + "_" + a.last4] || 0, _similar: sim ? sim.last4 : null });
  }) }));
});
app.post("/api/accounts", auth, function(req, res) {
  var b = req.body || {};
  var clean = function(arr) {
    var out = [];
    (Array.isArray(arr) ? arr : []).forEach(function(x) { x = String(x || "").trim(); if (x && out.indexOf(x) === -1) out.push(x); });
    return out;
  };
  var branches = clean(b.branches), aliases = clean(b.aliases);
  var seen = {};
  var accounts = (Array.isArray(b.accounts) ? b.accounts : []).map(function(a) {
    var l4 = last4(a.last4);
    if (!l4) return null;
    var bank = normalizeBank(a.bank);
    if (seen[bank + "_" + l4]) return null;
    seen[bank + "_" + l4] = true;
    return {
      bank: bank, last4: l4,
      alt: clean(a.alt).map(last4).filter(Boolean),
      full: String(a.full || "").replace(/\D/g, "").length >= 8 ? String(a.full).replace(/\D/g, "") : null,
      branch: branches.indexOf(a.branch) !== -1 ? a.branch : null,
      type: ACCT_TYPES.indexOf(a.type) !== -1 ? a.type : "company",
      label: String(a.label || "").trim(),
      firstSeenSlip: a.firstSeenSlip || null,
      firstSeenAt: a.firstSeenAt || new Date().toISOString()
    };
  }).filter(Boolean);
  var reg = { branches: branches, aliases: aliases, accounts: accounts };
  saveAccounts(reg);
  res.json(reg);
});
// Rename a branch everywhere (accounts + saved splits)
app.post("/api/branches/rename", auth, function(req, res) {
  var from = String(req.body.from || "").trim(), to = String(req.body.to || "").trim();
  if (!from || !to) return res.status(400).json({ error: "from/to required" });
  var reg = loadAccounts();
  if (reg.branches.indexOf(from) === -1) return res.status(404).json({ error: "branch not found" });
  if (reg.branches.indexOf(to) !== -1) return res.status(400).json({ error: "name already used" });
  reg.branches[reg.branches.indexOf(from)] = to;
  reg.accounts.forEach(function(a) { if (a.branch === from) a.branch = to; });
  saveAccounts(reg);
  var list = loadPayments(), changed = 0;
  list.forEach(function(p) { (p.split || []).forEach(function(s) { if (s.branch === from) { s.branch = to; changed++; } }); });
  if (changed) savePayments(list);
  res.json({ ok: true, accounts: reg, splitsUpdated: changed });
});
// Back-scan: ask AI which recipients are our own companies (suggestions only)
app.post("/api/accounts/scan", auth, async function(req, res) {
  try {
    var y = parseInt(req.body.year), m = parseInt(req.body.month);
    if (!y || !m) return res.status(400).json({ error: "year/month required" });
    if (!loadAccounts().aliases.length) return res.status(400).json({ error: "add your company names first" });
    var from = ymKey(y, m) + "-01";
    var ids = loadPayments().filter(function(p) { return p.transaction_date && p.transaction_date >= from; }).map(function(p) { return p.id; });
    var r = await flagInternalCandidates(ids, true);
    res.json(Object.assign({ ok: true }, r));
  } catch (e) {
    console.error("Scan error:", e.message, e.response ? JSON.stringify(e.response.data) : "");
    res.status(500).json({ error: e.message });
  }
});
app.get("/api/check", auth, function(req, res) {
  var pm = prevMonthOf(bkkNow().y, bkkNow().m);
  var r = checkMonth(parseInt(req.query.year) || pm.y, parseInt(req.query.month) || pm.m);
  var st = loadSchedule().months[ymKey(r.year, r.month)] || {};
  r.sentAt = st.sentAt || null;
  r.recipientSet = !!reportRecipient();
  res.json(r);
});

// Internal transfer: true = internal, false = real expense, null = automatic
app.post("/api/slips/:id/internal", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var v = req.body.value;
  if (v === true || v === false) list[idx].internal = v; else delete list[idx].internal;
  if (list[idx].internal_suggest === "pending" && list[idx].internal != null) list[idx].internal_suggest = "done";
  savePayments(list);
  res.json({ ok: true, record: list[idx] });
});
// Branch split of one expense: [{branch, amount}] — remainder stays with the paying account's branch
app.post("/api/slips/:id/split", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var branches = loadAccounts().branches;
  var split = (Array.isArray(req.body.split) ? req.body.split : []).map(function(s) {
    return { branch: String(s.branch || ""), amount: Math.round((Number(s.amount) || 0) * 100) / 100 };
  }).filter(function(s) { return s.amount > 0; });
  if (split.some(function(s) { return branches.indexOf(s.branch) === -1; })) return res.status(400).json({ error: "unknown branch" });
  var sum = split.reduce(function(s, x) { return s + x.amount; }, 0);
  if (sum > (Number(list[idx].amount) || 0) + 1) return res.status(400).json({ error: "split is more than the slip amount" });
  if (split.length) list[idx].split = split; else delete list[idx].split;
  savePayments(list);
  res.json({ ok: true, record: list[idx] });
});
// "No, not a share" — hide one link suggestion for good
app.post("/api/slips/:id/link-dismiss", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var eid = Number(req.body.expenseId);
  if (!eid) return res.status(400).json({ error: "expenseId required" });
  list[idx].link_dismissed = (list[idx].link_dismissed || []).filter(function(x) { return x !== eid; }).concat([eid]);
  savePayments(list);
  res.json({ ok: true });
});
// Link internal transfer → the expense it pays a share of; moves that share to the sender's branch
app.post("/api/slips/:id/link", auth, function(req, res) {
  var list = loadPayments();
  var reg = loadAccounts();
  var ti = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (ti === -1) return res.status(404).json({ error: "not found" });
  var t = list[ti];
  var ct = classify(t, reg);
  var share = Number(t.amount) || 0;
  // move `amt` of expense e from branch fromB to toB; the home branch keeps the remainder implicitly
  function shift(e, fromB, toB, amt) {
    var home = classify(e, reg).fromBranch || NO_BRANCH;
    var map = {};
    classify(e, reg).alloc.forEach(function(a) { map[a.branch] = a.amount; });
    var move = Math.min(amt, map[fromB] || 0);
    map[fromB] = (map[fromB] || 0) - move;
    map[toB] = (map[toB] || 0) + move;
    var split = Object.keys(map).filter(function(b) { return b !== home && map[b] > 0.009; })
      .map(function(b) { return { branch: b, amount: Math.round(map[b] * 100) / 100 }; });
    if (split.length) e.split = split; else delete e.split;
  }
  var expenseId = req.body.expenseId;
  if (expenseId == null) {
    if (!t.linked_to) return res.json({ ok: true });
    var oi = list.findIndex(function(p) { return String(p.id) === String(t.linked_to); });
    if (oi !== -1 && ct.fromBranch) shift(list[oi], ct.fromBranch, classify(list[oi], reg).fromBranch || NO_BRANCH, share);
    delete t.linked_to;
    savePayments(list);
    return res.json({ ok: true });
  }
  if (ct.kind !== "internal") return res.status(400).json({ error: "not an internal transfer" });
  if (!ct.fromBranch) return res.status(400).json({ error: "set the branch of the sending account first" });
  if (t.linked_to) return res.status(400).json({ error: "already linked — unlink first" });
  var ei = list.findIndex(function(p) { return String(p.id) === String(expenseId); });
  if (ei === -1) return res.status(404).json({ error: "expense not found" });
  var ce = classify(list[ei], reg);
  if (ce.kind !== "expense") return res.status(400).json({ error: "target must be a normal expense" });
  if (!ce.fromBranch) return res.status(400).json({ error: "set the branch of the paying account first" });
  shift(list[ei], ce.fromBranch, ct.fromBranch, share);
  t.linked_to = list[ei].id;
  savePayments(list);
  res.json({ ok: true, record: t, expense: list[ei] });
});

// Upload one image from the web app (base64). Auto-detect & process.
app.post("/api/upload", auth, async function(req, res) {
  try {
    var b64 = (req.body && req.body.imageBase64) || "";
    b64 = b64.replace(/^data:image\/\w+;base64,/, "");
    if (!b64) return res.status(400).json({ error: "no image" });
    var buf = Buffer.from(b64, "base64");
    var type = await detectImageType(b64);
    if (type === "SLIP") {
      var data = await readSlip(b64, "");
      if (data) {
        var rec = createSlipRecord(buf, data, "web");
        return res.json({ result: "slip", record: rec });
      }
      var item = createInboxItem(buf, "unknown", null, "web");
      return res.json({ result: "inbox", item: item, note: "slip unreadable" });
    }
    var billAmount = (type === "BILL") ? await readBillAmount(b64) : null;
    var item = createInboxItem(buf, type === "BILL" ? "bill" : "unknown", billAmount, "web");
    res.json({ result: "inbox", item: item });
  } catch (e) {
    console.error("Upload error:", e.message);
    res.status(500).json({ error: e.message });
  }
});

// Inbox item → treat as SLIP (read it, create payment)
app.post("/api/inbox/:id/as-slip", auth, async function(req, res) {
  var inbox = loadInbox();
  var idx = inbox.findIndex(function(x) { return x.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var item = inbox[idx];
  var fpath = path.join(BILLS_DIR, item.file);
  if (!fs.existsSync(fpath)) return res.status(404).json({ error: "file missing" });
  var buf = fs.readFileSync(fpath);
  var data = await readSlip(buf.toString("base64"), "");
  if (!data) return res.status(422).json({ error: "could not read as slip" });
  var rec = createSlipRecord(buf, data, item.source || "web");
  fs.unlinkSync(fpath);
  inbox.splice(idx, 1);
  saveInbox(inbox);
  res.json({ record: rec });
});

// Inbox item → mark as BILL (read amount if missing)
app.post("/api/inbox/:id/as-bill", auth, async function(req, res) {
  var inbox = loadInbox();
  var idx = inbox.findIndex(function(x) { return x.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  inbox[idx].type = "bill";
  if (inbox[idx].billAmount === null || inbox[idx].billAmount === undefined) {
    var fpath = path.join(BILLS_DIR, inbox[idx].file);
    if (fs.existsSync(fpath)) {
      inbox[idx].billAmount = await readBillAmount(fs.readFileSync(fpath).toString("base64"));
    }
  }
  saveInbox(inbox);
  res.json({ item: inbox[idx] });
});

// Match inbox bill → slip (Option B: instant amount compare, warn on mismatch)
app.post("/api/match", auth, function(req, res) {
  var slipId = req.body.slipId, inboxId = req.body.inboxId, force = !!req.body.force;
  var list = loadPayments();
  var pIdx = list.findIndex(function(p) { return String(p.id) === String(slipId); });
  if (pIdx === -1) return res.status(404).json({ error: "slip not found" });
  var inbox = loadInbox();
  var iIdx = inbox.findIndex(function(x) { return x.id === inboxId; });
  if (iIdx === -1) return res.status(404).json({ error: "inbox item not found" });
  var slipAmt = Number(list[pIdx].amount) || null;
  var billAmt = (inbox[iIdx].billAmount === 0 || inbox[iIdx].billAmount) ? Number(inbox[iIdx].billAmount) : null;
  if (!force && slipAmt !== null && billAmt !== null && Math.abs(slipAmt - billAmt) > 1) {
    return res.json({ warning: true, slipAmount: slipAmt, billAmount: billAmt });
  }
  var bills = getBillFiles(list[pIdx]);
  bills.push(inbox[iIdx].file);
  setBillFiles(list[pIdx], bills);
  list[pIdx].no_bill_expected = false;
  savePayments(list);
  inbox.splice(iIdx, 1);
  saveInbox(inbox);
  res.json({ matched: true, record: list[pIdx] });
});

// Unmatch: detach bill from slip, put it back in inbox
app.post("/api/unmatch", auth, function(req, res) {
  var list = loadPayments();
  var pIdx = list.findIndex(function(p) { return String(p.id) === String(req.body.slipId); });
  if (pIdx === -1) return res.status(404).json({ error: "slip not found" });
  var bills = getBillFiles(list[pIdx]);
  if (!bills.length) return res.status(400).json({ error: "no bill attached" });
  bills.forEach(function(f) { moveToInboxAsItem(f, "bill", null, "unmatch"); });
  setBillFiles(list[pIdx], []);
  savePayments(list);
  res.json({ ok: true });
});

// Remove ONE bill from a slip → back to inbox
app.post("/api/slips/:id/remove-bill", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var bills = getBillFiles(list[idx]);
  var f = req.body.file;
  var bIdx = bills.indexOf(f);
  if (bIdx === -1) return res.status(404).json({ error: "bill not on this slip" });
  bills.splice(bIdx, 1);
  setBillFiles(list[idx], bills);
  savePayments(list);
  moveToInboxAsItem(f, "bill", null, "removed");
  res.json({ ok: true, record: list[idx] });
});

// Add a bill directly to a slip by upload (no AI — user chose the slip)
app.post("/api/slips/:id/bill-upload", auth, function(req, res) {
  try {
    var list = loadPayments();
    var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
    if (idx === -1) return res.status(404).json({ error: "not found" });
    var b64 = (req.body && req.body.imageBase64) || "";
    b64 = b64.replace(/^data:image\/\w+;base64,/, "");
    if (!b64) return res.status(400).json({ error: "no image" });
    var buf = Buffer.from(b64, "base64");
    var fname = "img_" + Date.now() + "_" + Math.floor(Math.random()*1000) + ".jpg";
    fs.writeFileSync(path.join(BILLS_DIR, fname), buf);
    var bills = getBillFiles(list[idx]);
    bills.push(fname);
    setBillFiles(list[idx], bills);
    list[idx].no_bill_expected = false;
    savePayments(list);
    res.json({ ok: true, record: list[idx] });
  } catch (e) { res.status(500).json({ error: e.message }); }
});

// Categories
app.get("/api/categories", auth, function(req, res) { res.json(loadCategories()); });
app.post("/api/categories", auth, function(req, res) {
  var cats = req.body && req.body.categories;
  if (!Array.isArray(cats)) return res.status(400).json({ error: "categories must be an array" });
  cats = cats.map(function(c){ return String(c).trim(); }).filter(function(c){ return c.length > 0; });
  if (!cats.length) return res.status(400).json({ error: "list cannot be empty" });
  saveCategories(cats);
  res.json({ ok: true, categories: cats });
});

// Bulk auto-categorize one month (text-only AI call, chunks of 60)
app.post("/api/categorize", auth, async function(req, res) {
  try {
    var year = parseInt(req.body.year), month = parseInt(req.body.month);
    if (!year || !month) return res.status(400).json({ error: "year/month required" });
    var cats = loadCategories();
    var list = loadPayments();
    var targets = list.filter(function(p) {
      if (p.category) return false;
      if (!p.transaction_date) return false;
      var d = parseISO(p.transaction_date);
      return d.getFullYear() === year && d.getMonth() + 1 === month;
    });
    var totalTargets = targets.length;
    if (!totalTargets) return res.json({ ok: true, categorized: 0, of: 0 });
    var updated = 0;
    // Pass 1: vendor memory (no AI, free)
    var vmB = vendorCategoryMap();
    var freshPre = loadPayments();
    targets = targets.filter(function(p) {
      var k = vendorKey(p.recipient_name);
      if (p.recipient_name && vmB[k]) {
        var iP = freshPre.findIndex(function(q) { return q.id === p.id; });
        if (iP !== -1 && !freshPre[iP].category) { freshPre[iP].category = vmB[k]; updated++; }
        return false;
      }
      return true;
    });
    savePayments(freshPre);
    for (var off = 0; off < targets.length; off += 60) {
      var chunk = targets.slice(off, off + 60).map(function(p) {
        return { id: p.id, purpose: p.purpose || "", recipient: p.recipient_name || "", amount: Number(p.amount) || 0 };
      });
      var resp = await axios.post(
        "https://api.anthropic.com/v1/messages",
        { model: "claude-sonnet-4-5", max_tokens: 3000, messages: [{ role: "user", content:
          "You are classifying clinic expense transactions into categories.\nCategories (use the exact strings): " + JSON.stringify(cats) +
          "\nThe owner's past choices (match this style): " + JSON.stringify(categoryExamples(20)) +
          "\nTransactions: " + JSON.stringify(chunk) +
          "\nReturn ONLY a JSON array like [{\"id\":1,\"category\":\"...\"}]. Use null when nothing fits well." }] },
        { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
      );
      var mapping = safeJson(resp.data.content[0].text);
      if (!Array.isArray(mapping)) { console.error("Categorize: no JSON array in reply"); continue; }
      var fresh = loadPayments();
      mapping.forEach(function(m) {
        if (!m || m.category == null) return;
        if (cats.indexOf(m.category) === -1) return;
        var i2 = fresh.findIndex(function(p) { return p.id === m.id; });
        if (i2 !== -1 && !fresh[i2].category) { fresh[i2].category = m.category; updated++; }
      });
      savePayments(fresh);
    }
    res.json({ ok: true, categorized: updated, of: totalTargets });
  } catch (e) {
    console.error("Categorize error:", e.message, e.response ? JSON.stringify(e.response.data) : "");
    res.status(500).json({ error: e.message });
  }
});

// Mark / unmark a slip as "no bill expected"
app.post("/api/slips/:id/no-bill", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  list[idx].no_bill_expected = !!req.body.expected;
  savePayments(list);
  res.json({ ok: true, record: list[idx] });
});

// Send a recorded slip back to the inbox (deletes the transaction; bills return too)
app.post("/api/slips/:id/to-inbox", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var p = list[idx];
  getBillFiles(p).forEach(function(f) { moveToInboxAsItem(f, "bill", null, "back"); });
  if (p.imageFile) {
    var oldPath = path.join(SLIPS_DIR, p.imageFile);
    var newName = "img_" + Date.now() + "_" + Math.floor(Math.random()*1000) + ".jpg";
    try {
      if (fs.existsSync(oldPath)) fs.renameSync(oldPath, path.join(BILLS_DIR, newName));
      moveToInboxAsItem(newName, "unknown", null, "back");
    } catch (e) { console.error("to-inbox move error:", e.message); }
  }
  list.splice(idx, 1);
  savePayments(list);
  res.json({ ok: true });
});

// Batch inbox actions: mark as bill (instant, amounts read in background) or delete
app.post("/api/inbox/batch", auth, function(req, res) {
  var ids = Array.isArray(req.body.ids) ? req.body.ids : [];
  var action = req.body.action;
  if (!ids.length) return res.status(400).json({ error: "no ids" });
  var inbox = loadInbox();
  if (action === "bill") {
    var needAmount = [];
    ids.forEach(function(id) {
      var idx = inbox.findIndex(function(x) { return x.id === id; });
      if (idx === -1) return;
      inbox[idx].type = "bill";
      if (inbox[idx].billAmount === null || inbox[idx].billAmount === undefined) needAmount.push(id);
    });
    saveInbox(inbox);
    if (needAmount.length) readInboxAmountsInBackground(needAmount);
    return res.json({ ok: true, marked: ids.length, readingAmounts: needAmount.length });
  }
  if (action === "delete") {
    var deleted = 0;
    ids.forEach(function(id) {
      var idx = inbox.findIndex(function(x) { return x.id === id; });
      if (idx === -1) return;
      var fp = path.join(BILLS_DIR, inbox[idx].file);
      if (fs.existsSync(fp)) { try { fs.unlinkSync(fp); } catch(e) {} }
      inbox.splice(idx, 1);
      deleted++;
    });
    saveInbox(inbox);
    return res.json({ ok: true, deleted: deleted });
  }
  res.status(400).json({ error: "unknown action" });
});

// Edit slip fields
app.patch("/api/slips/:id", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var allowed = ["amount","transaction_date","transaction_time","bank_from","account_from","bank_to","account_to","recipient_name","purpose","reference_number","category","evidence_note","date_ok","dup_ok","recipient_ok","cost_type","cost_month","stmt_ok"];
  var newName = req.body.recipient_name;
  if (newName !== undefined && newName && list[idx].recipient_name && newName !== list[idx].recipient_name && !list[idx].recipient_read) {
    list[idx].recipient_read = list[idx].recipient_name; // what the AI read — used to auto-correct the same misread next time
  }
  allowed.forEach(function(f) {
    if (req.body[f] !== undefined) list[idx][f] = req.body[f];
  });
  if (list[idx].cost_type && COST_TYPES.indexOf(list[idx].cost_type) === -1) delete list[idx].cost_type;
  if (!list[idx].cost_type) delete list[idx].cost_type;
  if (list[idx].cost_month && !/^\d{4}-\d{2}$/.test(list[idx].cost_month)) delete list[idx].cost_month;
  if (!list[idx].cost_month) delete list[idx].cost_month;
  if (list[idx].category) {
    var catsP = loadCategories();
    if (catsP.indexOf(list[idx].category) === -1) list[idx].category = null;
  }
  // display id = yymmNN of the transaction month — re-number when the date moves to another month
  var td = list[idx].transaction_date;
  var wantPrefix = td && /^\d{4}-\d{2}/.test(td) ? td.slice(2, 4) + td.slice(5, 7) : null;
  if (wantPrefix && (!list[idx].display_id || String(list[idx].display_id).indexOf(wantPrefix) !== 0)) {
    if (list[idx].display_id) list[idx].old_display_ids = (list[idx].old_display_ids || []).concat([String(list[idx].display_id)]);
    list[idx].display_id = nextDisplayId(td);
  }
  savePayments(list);
  res.json(list[idx]);
});

app.delete("/api/slips/:id", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  list.splice(idx, 1);
  savePayments(list);
  res.json({ ok: true });
});

app.delete("/api/inbox/:id", auth, function(req, res) {
  var inbox = loadInbox();
  var idx = inbox.findIndex(function(x) { return x.id === req.params.id; });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var fpath = path.join(BILLS_DIR, inbox[idx].file);
  if (fs.existsSync(fpath)) { try { fs.unlinkSync(fpath); } catch(e) {} }
  inbox.splice(idx, 1);
  saveInbox(inbox);
  res.json({ ok: true });
});

// ── Routes ─────────────────────────────────────────────────────────────────
app.get("/", function(req, res) { res.send("✅ Slip Tracker Bot is running!"); });
app.get("/app", function(req, res) { res.sendFile(path.join(__dirname, "webapp.html")); });
app.use("/slips", authOrSigned, express.static(SLIPS_DIR));
app.use("/bills", authOrSigned, express.static(BILLS_DIR));
app.use("/dashboard", authOrSigned, express.static(path.join(__dirname, "dashboard")));
app.get("/api/report", authOrSigned, async function(req, res) {
  var now = new Date(); var year = parseInt(req.query.year) || now.getFullYear(); var month = parseInt(req.query.month) || now.getMonth() + 1;
  try { var file = await makeReport(year, month); res.download(file, "Payment_Report_"+year+"_"+String(month).padStart(2,"0")+".xlsx"); }
  catch(err) { res.status(500).send("Error: " + err.message); }
});

app.post("/webhook", function(req, res) {
  console.log("WEBHOOK HIT!");
  res.status(200).end();
  if (!req.body || !req.body.events) return;
  req.body.events.forEach(function(event) {
    handleEvent(event).catch(function(err) { console.error("Event error:", err.message); });
  });
});

if (require.main === module) {
  app.listen(PORT, function() {
    console.log("✅ Slip Tracker v8 running on port " + PORT);
    setTimeout(function() { scheduleTick(); }, 30 * 1000);
    setInterval(function() { scheduleTick(); driveAutoSync(); }, 10 * 60 * 1000);
  });
}
// exported for offline test scripts only
module.exports = { app: app, client: client, classify: classify, loadAccounts: loadAccounts, checkMonth: checkMonth, monthTotals: monthTotals, makeReport: makeReport, scheduleTick: scheduleTick, linkSuggestions: linkSuggestions, nameMatchesAlias: nameMatchesAlias, readinessText: readinessText, reportReadyText: reportReadyText, suggestYearFix: suggestYearFix, suspectDates: suspectDates, duplicateOf: duplicateOf, similarToAccount: similarToAccount, correctedName: correctedName, pnlMonth: pnlMonth, loadPnl: loadPnl,
  importStatement: importStatement, importAny: importAny, pnlReadiness: pnlReadiness, overviewText: overviewText, reconMonth: reconMonth, payrollPlan: payrollPlan, applyPayroll: applyPayroll, driveSync: driveSync, parseCounterparty: parseCounterparty, loadStatements: loadStatements };

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

// One-time safety backups (v4: multi-bill migration, v8: branch split / internal transfer fields, v9: date fix / re-numbering)
["payments_backup_v4.json", "payments_backup_v8.json", "payments_backup_v9.json"].forEach(function(bn) {
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
function categoryExamples(limit) {
  var ex = [];
  var l = loadPayments();
  for (var i = l.length - 1; i >= 0 && ex.length < (limit || 15); i--) {
    var p = l[i];
    if (p.category && (p.purpose || p.recipient_name)) ex.push({ purpose: p.purpose || "", recipient: p.recipient_name || "", category: p.category });
  }
  return ex;
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
            { type: "text", text: "You are an expert Thai bank payment slip reader. Extract ALL details carefully.\n\nToday is " + format(new Date(), "yyyy-MM-dd") + " — slips are almost always from the last few weeks.\n\nDate formats:\n- DD/MM/YYYY or DD/MM/YY\n- DD MMM YYYY (e.g. 08 พ.ค. 2568)\n- Buddhist year (พ.ศ.) subtract 543 to get AD. 2568=2025, 2569=2026\n- 2-digit years: with an ENGLISH month (e.g. \"28 Oct 26\") the year is AD 20YY → 2026. With a THAI month (e.g. \"28 ต.ค. 69\") it is พ.ศ. 25YY → 2569 → 2026.\n- Time: HH:MM or HH:MM:SS\n\nAccounts: the slip shows FROM (sender = us) and TO (recipient). account_from / account_to = the last 4 DIGITS shown for that account (e.g. xxx-x-x8464-x → 8464). Read each digit carefully (6 vs 8, 3 vs 8, 1 vs 7).\n\nBanks: SCB, Krungthai (KTB), Bangkok Bank (BBL), Kasikorn (KBank), Krungsri (BAY), TMB, GSB, PromptPay.\nUser note: \"" + (caption || "") + "\"\nOwner's past category choices (match this style): " + JSON.stringify(categoryExamples(15)) + "\n\nReturn ONLY valid JSON:\n{\n  \"bank_from\": \"bank name or null\",\n  \"account_from\": \"last 4 digits or null\",\n  \"bank_to\": \"bank name or null\",\n  \"account_to\": \"last 4 digits or null\",\n  \"recipient_name\": \"name or null\",\n  \"amount\": 0.00,\n  \"transaction_date\": \"YYYY-MM-DD or null\",\n  \"transaction_time\": \"HH:MM or null\",\n  \"reference_number\": \"ref or null\",\n  \"purpose\": \"use user note if given, else the slip memo or what is paid to the TO/recipient — never describe it as a payment to the FROM/sender name\",\n  \"category\": \"exactly one string from this list that best fits the payment, or null: " + JSON.stringify(loadCategories()) + "\",\n  \"slip_type\": \"mobile_banking or internet_banking or prompt_pay\"\n}" }
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
  if (/kasikorn|kbank|kbiz/.test(l)) return "KBank";
  if (/government\s*savings|gsb|ออมสิน/.test(l)) return "GSB";
  if (/ayudhya|krungsri|(^|[^a-z])bay([^a-z]|$)/.test(l)) return "Krungsri";
  if (/krung\s*thai|ktb/.test(l)) return "Krungthai";
  if (/bangkok\s*bank|bbl/.test(l)) return "Bangkok Bank";
  if (/siam\s*commercial|scb/.test(l)) return "SCB";
  if (/tmb|ttb|thanachart/.test(l)) return "TTB";
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
function findAccount(reg, bank, acct) {
  var l = last4(acct);
  if (!l) return null;
  var nb = normalizeBank(bank);
  return reg.accounts.find(function(a) { return a.bank === nb && (a.last4 === l || (a.alt || []).indexOf(l) !== -1); }) || null;
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
  r.ready = !r.missingBill.length && !r.noCategory.length && !r.pendingInternal.length && !r.unassigned.length && !r.noDate.length && !r.suspectDate.length;
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

function idList(ids) { return ids.slice(0, 8).map(function(i) { return "#" + i; }).join(" ") + (ids.length > 8 ? " +" + (ids.length - 8) : ""); }
function readinessText(r) {
  if (r.ready) return "✅ " + thMonthRange(r.year, r.month) + " ครบแล้ว — " + r.count + " รายการ" + (r.inbox ? "\n📥 Inbox ค้าง " + r.inbox + " รูป (ไม่บล็อก)" : "");
  var lines = ["⏳ Report ค่าใช้จ่าย " + thMonthRange(r.year, r.month) + " ยังออกไม่ได้", "ต้องเคลียร์ก่อน:"];
  if (r.missingBill.length) lines.push("⚠️ ขาดบิล/note: " + r.missingBill.length + " — " + idList(r.missingBill));
  if (r.noCategory.length) lines.push("🏷 ยังไม่จัดหมวด: " + r.noCategory.length + " — " + idList(r.noCategory));
  if (r.noDate.length) lines.push("📅 ไม่มีวันที่: " + r.noDate.length + " — " + idList(r.noDate));
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
    BASE_URL + "/api/report?year=" + y + "&month=" + m
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
        row.getCell(5).value={text:"🔗 View Slip",hyperlink:baseUrl+"/slips/"+p.imageFile};
        row.getCell(5).font={color:{argb:"FF1565C0"},underline:true,bold:true};
      } else {
        row.getCell(5).value="No slip image";
        row.getCell(5).font={italic:true,color:{argb:"FF9E9E9E"}};
      }
      var bills=getBillFiles(p);
      for(bi=0;bi<maxBills;bi++){
        var cell=row.getCell(6+bi);
        if(bills[bi]){
          cell.value={text:"🔗 View Bill"+(maxBills>1?(" "+(bi+1)):""),hyperlink:baseUrl+"/bills/"+bills[bi]};
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
async function scheduleTick(nowMs) {
  if (ticking) return;
  ticking = true;
  try {
    var to = reportRecipient();
    if (!to) return;
    var t = bkkNow(nowMs);
    var st = loadSchedule();
    if (!st.startReportMonth) {
      // first run: don't chase months whose report day already passed before this feature existed
      var nx = t.m === 12 ? { y: t.y + 1, m: 1 } : { y: t.y, m: t.m + 1 };
      st.startReportMonth = t.day === 1 ? ymKey(t.y, t.m) : ymKey(nx.y, nx.m);
      saveSchedule(st);
    }
    var cur = ymKey(t.y, t.m);
    var ry = +st.startReportMonth.slice(0, 4), rm = +st.startReportMonth.slice(5);
    while (ymKey(ry, rm) <= cur) {
      var tg = prevMonthOf(ry, rm);
      var k = ymKey(tg.y, tg.m);
      var ms = st.months[k] || {};
      var slot = remindSlot(t.h);
      if (!ms.sentAt && (ms.firstAt || slot !== null)) {
        var chk = checkMonth(tg.y, tg.m);
        if (chk.ready) {
          await sendMonthlyReport(to, tg.y, tg.m, false);
          ms.sentAt = new Date().toISOString();
        } else if (slot !== null && ms.lastSlot !== t.date + "_" + slot) {
          await pushText(to, readinessText(chk));
          ms.lastSlot = t.date + "_" + slot;
          ms.reminders = (ms.reminders || 0) + 1;
        }
        if (!ms.firstAt) ms.firstAt = new Date().toISOString();
        st.months[k] = ms;
        saveSchedule(st);
      }
      if (rm === 12) { ry++; rm = 1; } else rm++;
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
    if (tl === "/notify here" || tl === "/notify off") {
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
        return client.pushMessage(gid, { type: "text", text: ["✅ Report ready — "+format(new Date(year,month-1),"MMMM yyyy"),"📋 "+filtered.length+" transactions","💰 ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2})+" (excl. internal)","🏦 "+Object.keys(accts).length+" account sheet(s)","📎 Bill evidence: "+withBill+"/"+filtered.length,"","⬇️ Download Excel:",BASE_URL+"/api/report?year="+year+"&month="+month].join("\n") });
      } catch(err) {
        console.error("Report error:", err.message);
        return client.pushMessage(gid, { type: "text", text: "❌ Report error: " + err.message });
      }
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
const sessions = {};
function auth(req, res, next) {
  var t = req.headers["x-auth"];
  if (t && sessions[t] && sessions[t] > Date.now()) {
    sessions[t] = Date.now() + 12 * 3600 * 1000;
    return next();
  }
  res.status(401).json({ error: "unauthorized" });
}

app.post("/api/login", function(req, res) {
  if (!WEB_PASSWORD) return res.status(500).json({ error: "WEB_PASSWORD not set on server" });
  if ((req.body && req.body.password) === WEB_PASSWORD) {
    var token = crypto.randomBytes(24).toString("hex");
    sessions[token] = Date.now() + 12 * 3600 * 1000;
    return res.json({ token: token });
  }
  res.status(401).json({ error: "wrong password" });
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
      suggestions: c.kind === "internal" ? linkSuggestions(p, all, reg) : [],
      linkedDisplay: p.linked_to ? (function() { var e = all.find(function(q) { return String(q.id) === String(p.linked_to); }); return e ? showId(e) : null; })() : null
    } });
  }));
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
  var allowed = ["amount","transaction_date","transaction_time","bank_from","account_from","bank_to","account_to","recipient_name","purpose","reference_number","category","evidence_note","date_ok"];
  allowed.forEach(function(f) {
    if (req.body[f] !== undefined) list[idx][f] = req.body[f];
  });
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
app.use("/slips", express.static(SLIPS_DIR));
app.use("/bills", express.static(BILLS_DIR));
app.use("/dashboard", express.static(path.join(__dirname, "dashboard")));
app.get("/api/report", async function(req, res) {
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
    setInterval(function() { scheduleTick(); }, 10 * 60 * 1000);
  });
}
// exported for offline test scripts only
module.exports = { app: app, client: client, classify: classify, loadAccounts: loadAccounts, checkMonth: checkMonth, monthTotals: monthTotals, makeReport: makeReport, scheduleTick: scheduleTick, linkSuggestions: linkSuggestions, nameMatchesAlias: nameMatchesAlias, readinessText: readinessText, reportReadyText: reportReadyText, suggestYearFix: suggestYearFix, suspectDates: suspectDates };

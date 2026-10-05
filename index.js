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

const DATA_DIR   = path.join(__dirname, "data");
const SLIPS_DIR  = path.join(DATA_DIR, "slips");
const BILLS_DIR  = path.join(DATA_DIR, "bills");
const DATA_FILE  = path.join(DATA_DIR, "payments.json");
const INBOX_FILE = path.join(DATA_DIR, "inbox.json");

[DATA_DIR, SLIPS_DIR, BILLS_DIR].forEach(function(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
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

// ── Last slip per user (for purpose note in LINE) ─────────────────────────
const lastSlip = {};
function getLastSlip(uid) {
  var s = lastSlip[uid];
  if (!s) return null;
  if (Date.now() > s.expiresAt) { delete lastSlip[uid]; return null; }
  return s;
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
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text", text: "Look at this image carefully and decide if it is a bank transfer confirmation slip or a bill/invoice.\n\nIt is a SLIP if it shows ANY of these:\n- 'Transfer Completed' or 'Transfer Successful'\n- 'โอนเงินสำเร็จ' or 'รายการสำเร็จ'\n- 'Transaction No' or 'Reference No' or 'เลขที่รายการ'\n- Sender account AND receiver account with bank names\n- FROM and TO with account numbers\n- Banks: SCB, KBank, KBIZ, GSB, Krungthai, Bangkok Bank, Krungsri, TMB, TTB, PromptPay\n\nIt is a BILL if it shows ANY of these:\n- QR code for payment (even if it shows an amount or bank account number)\n- Invoice or receipt with list of products/services\n- 'Please pay' or 'Amount due' or 'ยอดที่ต้องชำระ'\n- Bill with company name and itemized costs\n- Bank account number for receiving payment (but NO transfer confirmation)\n\nKEY RULE: A QR code or bank account number shown on a bill is NOT a slip. A slip must show transfer confirmation that money has already been sent.\n\nReturn ONLY one word: SLIP, BILL, or UNKNOWN" }
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
    console.error("Detect error:", e.message);
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
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text", text: "You are an expert Thai bank payment slip reader. Extract ALL details carefully.\n\nThai date formats:\n- DD/MM/YYYY or DD/MM/YY\n- DD MMM YYYY (e.g. 08 พ.ค. 2568)\n- Buddhist year (พ.ศ.) subtract 543 to get AD. 2568=2025, 2569=2026\n- Time: HH:MM or HH:MM:SS\n\nBanks: SCB, Krungthai (KTB), Bangkok Bank (BBL), Kasikorn (KBank), Krungsri (BAY), TMB, GSB, PromptPay.\nUser note: \"" + (caption || "") + "\"\n\nReturn ONLY valid JSON:\n{\n  \"bank_from\": \"bank name or null\",\n  \"account_from\": \"last 4 digits or null\",\n  \"bank_to\": \"bank name or null\",\n  \"account_to\": \"last 4 digits or null\",\n  \"recipient_name\": \"name or null\",\n  \"amount\": 0.00,\n  \"transaction_date\": \"YYYY-MM-DD or null\",\n  \"transaction_time\": \"HH:MM or null\",\n  \"reference_number\": \"ref or null\",\n  \"purpose\": \"use user note if given, else infer from recipient\",\n  \"slip_type\": \"mobile_banking or internet_banking or prompt_pay\"\n}" }
          ]
        }]
      },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    const raw = resp.data.content[0].text.trim();
    console.log("Claude slip:", raw);
    return JSON.parse(raw.replace(/```json/g,"").replace(/```/g,"").trim());
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
            { type: "image", source: { type: "base64", media_type: "image/jpeg", data: imageBase64 } },
            { type: "text", text: "Look at this bill or invoice image. Find the TOTAL amount to pay.\nReturn ONLY a JSON object, nothing else:\n{\"amount\": 0.00}\nAmount must be a number. No commas. If you cannot find an amount return {\"amount\": null}" }
          ]
        }]
      },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    const raw = resp.data.content[0].text.trim();
    const parsed = JSON.parse(raw.replace(/```json/g,"").replace(/```/g,"").trim());
    return (parsed.amount === 0 || parsed.amount) ? parsed.amount : null;
  } catch (e) {
    console.error("Bill amount error:", e.message);
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
function createSlipRecord(buf, data, source) {
  var fname = "slip_" + Date.now() + ".jpg";
  fs.writeFileSync(path.join(SLIPS_DIR, fname), buf);
  var list = loadPayments();
  var rec = Object.assign({
    id: nextNumericId(),
    display_id: nextDisplayId(data.transaction_date),
    imageFile: fname,
    billFile: null,
    savedAt: new Date().toISOString(),
    source: source
  }, data);
  list.push(rec);
  savePayments(list);
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
  return [
    "✅ #" + showId(rec) + "  " + (rec.bank_from || "?") + "  " + amt,
    "→ " + (rec.recipient_name || rec.bank_to || "?") + (rec.purpose ? "  (" + rec.purpose + ")" : ""),
    "📎 Match bills at " + BASE_URL + "/app"
  ].join("\n");
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
function getBankColor(bank) {
  var map = { "SCB":"FF4E2E8C","Kasikorn":"FF1A8C2E","KBank":"FF1A8C2E","Krungthai":"FF009FDA","KTB":"FF009FDA","Bangkok":"FF0050A0","BBL":"FF0050A0","Krungsri":"FFD4A017","BAY":"FFD4A017","GSB":"FF8B0000","TMB":"FF0056A6","PromptPay":"FF2D7DD2" };
  var found = Object.keys(map).find(function(k) { return (bank||"").toLowerCase().indexOf(k.toLowerCase()) !== -1; });
  return found ? map[found] : "FF1565C0";
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

  const s1 = wb.addWorksheet("All Transactions");
  s1.mergeCells("A1:M1");
  var t1 = s1.getCell("A1");
  t1.value = "Payment Report — " + label;
  t1.font = { size: 14, bold: true, color: { argb: "FFFFFFFF" } };
  t1.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1565C0" } };
  t1.alignment = { horizontal: "center" };
  s1.getRow(1).height = 30;
  styleHdr(s1.addRow(["#","Date","Time","Amount (฿)","From Bank","Acct","To Bank","Acct","Recipient","Purpose","Ref","Type","Evidence"]), "FF0D47A1");
  rows.forEach(function(p, i) {
    var row = s1.addRow([showId(p),p.transaction_date||"",p.transaction_time||"",Number(p.amount)||0,p.bank_from||"",p.account_from||"",p.bank_to||"",p.account_to||"",p.recipient_name||"",p.purpose||"",p.reference_number||"",(p.slip_type||"").replace(/_/g," "),p.billFile?"✅":"⚠️"]);
    if (i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF3F8FF"}};});
    row.getCell(4).numFmt="#,##0.00"; row.getCell(4).font={bold:true,color:{argb:"FF1565C0"}}; row.height=20;
  });
  var tr=s1.addRow(["","TOTAL","",grandTotal]);
  tr.getCell(4).numFmt="#,##0.00"; tr.getCell(4).font={bold:true,size:13};
  tr.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
  s1.columns=[{width:9},{width:12},{width:7},{width:14},{width:15},{width:11},{width:15},{width:11},{width:20},{width:28},{width:16},{width:15},{width:10}];
  s1.views=[{state:"frozen",ySplit:2}];

  const s2 = wb.addWorksheet("Summary");
  s2.mergeCells("A1:D1");
  var t2=s2.getCell("A1"); t2.value="Summary — "+label;
  t2.font={size:13,bold:true,color:{argb:"FFFFFFFF"}}; t2.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF1565C0"}}; t2.alignment={horizontal:"center"}; s2.getRow(1).height=28;
  s2.addRow([]);
  styleHdr(s2.addRow(["Bank Account","Transactions","Total (฿)","% of Total"]),"FF0D47A1");
  var acctMap={};
  rows.forEach(function(p){var k=(p.bank_from||"Unknown")+" ****"+(p.account_from||"????"); if(!acctMap[k])acctMap[k]={count:0,sum:0}; acctMap[k].count++; acctMap[k].sum+=Number(p.amount)||0;});
  Object.entries(acctMap).sort(function(a,b){return b[1].sum-a[1].sum;}).forEach(function(e,i){
    var pct=grandTotal>0?(e[1].sum/grandTotal*100).toFixed(1)+"%":"0%";
    var row=s2.addRow([e[0],e[1].count,e[1].sum,pct]);
    row.getCell(3).numFmt="#,##0.00"; row.getCell(3).font={bold:true,color:{argb:"FF1565C0"}};
    if(i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};});
  });
  var s2tot=s2.addRow(["TOTAL",rows.length,grandTotal,"100%"]);
  s2tot.getCell(3).numFmt="#,##0.00"; s2tot.eachCell(function(c){c.font={bold:true};c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
  s2.addRow([]); s2.addRow([]);
  styleHdr(s2.addRow(["Purpose / Category","Transactions","Total (฿)","% of Total"]),"FF0D47A1");
  var purMap={};
  rows.forEach(function(p){var k=p.purpose||"Unspecified"; if(!purMap[k])purMap[k]={count:0,sum:0}; purMap[k].count++; purMap[k].sum+=Number(p.amount)||0;});
  Object.entries(purMap).sort(function(a,b){return b[1].sum-a[1].sum;}).forEach(function(e,i){
    var pct=grandTotal>0?(e[1].sum/grandTotal*100).toFixed(1)+"%":"0%";
    var row=s2.addRow([e[0],e[1].count,e[1].sum,pct]);
    row.getCell(3).numFmt="#,##0.00"; row.getCell(3).font={bold:true,color:{argb:"FF1565C0"}};
    if(i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};});
  });
  s2.columns=[{width:35},{width:14},{width:18},{width:12}];

  var accountGroups={};
  rows.forEach(function(p){
    var key=(p.bank_from||"Unknown").replace(/kasikorn|kbank/gi,"KBank")+"_"+(p.account_from||"????");
    if(!accountGroups[key])accountGroups[key]={bank:p.bank_from||"Unknown",acct:p.account_from||"????",payments:[]};
    accountGroups[key].payments.push(p);
  });

  var accountKeys=Object.keys(accountGroups).sort();
  for(var ai=0;ai<accountKeys.length;ai++){
    var grp=accountGroups[accountKeys[ai]];
    var acctTotal=grp.payments.reduce(function(s,p){return s+(Number(p.amount)||0);},0);
    var color=getBankColor(grp.bank);
    var sheetName=(grp.bank.replace(/kasikorn|kbank/gi,"KBank")+" ("+grp.acct+")").replace(/[*?:\\/\[\]]/g,"-").substring(0,31);
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
    styleHdr(sa.addRow(["#","Date","Amount (฿)","Purpose","💳 Payment Slip","📄 Bill / Invoice"]),color);
    sa.columns=[{width:9},{width:12},{width:14},{width:30},{width:28},{width:28}];
    var ri=sa.rowCount+1;
    for(var pi=0;pi<grp.payments.length;pi++){
      var p=grp.payments[pi];
      sa.getRow(ri).height=22;
      sa.getCell("A"+ri).value="#"+showId(p);
      sa.getCell("B"+ri).value=p.transaction_date||"";
      sa.getCell("C"+ri).value=Number(p.amount)||0; sa.getCell("C"+ri).numFmt="#,##0.00"; sa.getCell("C"+ri).font={bold:true,color:{argb:"FF1565C0"}};
      sa.getCell("D"+ri).value=p.purpose||""; sa.getCell("D"+ri).alignment={wrapText:true,vertical:"middle"};
      if(p.imageFile){
        sa.getCell("E"+ri).value={text:"🔗 View Slip",hyperlink:baseUrl+"/slips/"+p.imageFile};
        sa.getCell("E"+ri).font={color:{argb:"FF1565C0"},underline:true,bold:true};
      } else {
        sa.getCell("E"+ri).value="No slip image";
        sa.getCell("E"+ri).font={italic:true,color:{argb:"FF9E9E9E"}};
      }
      if(p.billFile){
        sa.getCell("F"+ri).value={text:"🔗 View Bill",hyperlink:baseUrl+"/bills/"+p.billFile};
        sa.getCell("F"+ri).font={color:{argb:"FF2E7D32"},underline:true,bold:true};
      } else {
        sa.getCell("F"+ri).value="⚠️ No bill";
        sa.getCell("F"+ri).font={italic:true,color:{argb:"FFBF360C"}};
      }
      ["A","B","C","D","E","F"].forEach(function(col){
        sa.getCell(col+ri).border={bottom:{style:"thin",color:{argb:"FFE0E0E0"}}};
        if(!sa.getCell(col+ri).alignment) sa.getCell(col+ri).alignment={vertical:"middle"};
      });
      if(pi%2===0){["A","B","C","D"].forEach(function(col){sa.getCell(col+ri).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF8FBFF"}};});}
      ri++;
    }
  }

  var outPath=path.join(DATA_DIR,"report_"+year+"_"+String(month).padStart(2,"0")+".xlsx");
  await wb.xlsx.writeFile(outPath);
  return outPath;
}

// ── LINE bot ───────────────────────────────────────────────────────────────
const client = new line.Client(LINE_CONFIG);

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
      var rec = createSlipRecord(buf, data, "line");
      lastSlip[uid] = { paymentId: rec.id, expiresAt: Date.now() + 5 * 60 * 1000 };
      return client.pushMessage(gid, { type: "text", text: slipReply(rec) });
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

    if (tl === "/help" || tl === "help") {
      return client.replyMessage(event.replyToken, {
        type: "text",
        text: [
          "💳 Slip Tracker Bot",
          "━━━━━━━━━━━━━━━━━━",
          "📸 Send slip → auto-read & recorded",
          "📸 Send bill/other → saved to inbox",
          "💬 Text after slip → sets purpose",
          "",
          "🌐 Web App (upload, match, edit):",
          BASE_URL + "/app",
          "",
          "/summary — this month totals",
          "/list — last 5 payments",
          "/report — Excel for this month",
          "/report YYYY MM — specific month",
          "/help — this menu"
        ].join("\n")
      });
    }

    if (tl === "/summary" || tl === "summary") {
      var list = loadPayments(); var now = new Date();
      var thisMonth = list.filter(function(p) { if (!p.transaction_date) return false; var d = parseISO(p.transaction_date); return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth(); });
      var total = thisMonth.reduce(function(s,p) { return s+(Number(p.amount)||0); }, 0);
      var withBill = thisMonth.filter(function(p) { return p.billFile; }).length;
      var inboxCount = loadInbox().length;
      var aMap = {}; thisMonth.forEach(function(p) { var k=(p.bank_from||"Unknown")+" ****"+(p.account_from||"????"); aMap[k]=(aMap[k]||0)+(Number(p.amount)||0); });
      var aLines = Object.entries(aMap).map(function(e) { return "  • "+e[0]+": ฿"+e[1].toLocaleString("th-TH",{minimumFractionDigits:2}); }).join("\n");
      return client.replyMessage(event.replyToken, { type: "text", text: ["📊 "+format(now,"MMMM yyyy")+" Summary","━━━━━━━━━━━━━━━━━━","📋 Transactions: "+thisMonth.length,"💰 Total: ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2}),"📎 With bill: "+withBill+"/"+thisMonth.length,"📥 Inbox (unorganized): "+inboxCount,"","By Account:",aLines||"  (none yet)","","Organize: "+BASE_URL+"/app"].join("\n") });
    }

    if (tl === "/list" || tl === "list") {
      var list = loadPayments();
      if (!list.length) return client.replyMessage(event.replyToken, { type: "text", text: "No payments recorded yet." });
      var recent = list.slice(-5).reverse();
      var lines = recent.map(function(p) { return ["#"+showId(p)+"  "+(p.transaction_date||"?")+"  ฿"+Number(p.amount||0).toLocaleString("th-TH"),"  "+(p.bank_from||"?")+" ****"+(p.account_from||"????"),"  "+(p.purpose||"-")+"  "+(p.billFile?"📎✅":"⚠️")].join("\n"); });
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
        var total = filtered.reduce(function(s,p){return s+(Number(p.amount)||0);},0);
        var withBill = filtered.filter(function(p){return p.billFile;}).length;
        var accts = {}; filtered.forEach(function(p){accts[(p.bank_from||"?")+" ****"+(p.account_from||"????")] = true;});
        return client.pushMessage(gid, { type: "text", text: ["✅ Report ready — "+format(new Date(year,month-1),"MMMM yyyy"),"📋 "+filtered.length+" transactions","💰 ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2}),"🏦 "+Object.keys(accts).length+" account sheet(s)","📎 Bill evidence: "+withBill+"/"+filtered.length,"","⬇️ Download Excel:",BASE_URL+"/api/report?year="+year+"&month="+month].join("\n") });
      } catch(err) {
        console.error("Report error:", err.message);
        return client.pushMessage(gid, { type: "text", text: "❌ Report error: " + err.message });
      }
    }

    // Text after a slip = purpose note
    if (last) {
      var list = loadPayments();
      var idx = list.findIndex(function(p) { return p.id === last.paymentId; });
      if (idx !== -1) { list[idx].purpose = txt; savePayments(list); }
      delete lastSlip[uid];
      var pid = idx !== -1 ? showId(list[idx]) : "?";
      return client.replyMessage(event.replyToken, { type: "text", text: "✅ Purpose saved for #" + pid + ": " + txt });
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
  res.json(list);
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
  list[pIdx].billFile = inbox[iIdx].file;
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
  if (!list[pIdx].billFile) return res.status(400).json({ error: "no bill attached" });
  var inbox = loadInbox();
  inbox.push({ id: "ib" + Date.now(), file: list[pIdx].billFile, type: "bill", billAmount: null, uploadedAt: new Date().toISOString(), source: "unmatch" });
  saveInbox(inbox);
  list[pIdx].billFile = null;
  savePayments(list);
  res.json({ ok: true });
});

// Edit slip fields
app.patch("/api/slips/:id", auth, function(req, res) {
  var list = loadPayments();
  var idx = list.findIndex(function(p) { return String(p.id) === String(req.params.id); });
  if (idx === -1) return res.status(404).json({ error: "not found" });
  var allowed = ["amount","transaction_date","transaction_time","bank_from","account_from","bank_to","account_to","recipient_name","purpose","reference_number"];
  allowed.forEach(function(f) {
    if (req.body[f] !== undefined) list[idx][f] = req.body[f];
  });
  if (!list[idx].display_id && list[idx].transaction_date) {
    list[idx].display_id = nextDisplayId(list[idx].transaction_date);
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

app.listen(PORT, function() { console.log("✅ Slip Tracker v3 running on port " + PORT); });

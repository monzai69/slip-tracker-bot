"use strict";

const express = require("express");
const line = require("@line/bot-sdk");
const axios = require("axios");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { format, parseISO } = require("date-fns");

const app = express();
const PORT = process.env.PORT || 8080;

const LINE_CONFIG = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN || "",
  channelSecret: process.env.LINE_CHANNEL_SECRET || "",
};
const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || "";

const DATA_DIR  = path.join(__dirname, "data");
const SLIPS_DIR = path.join(DATA_DIR, "slips");
const BILLS_DIR = path.join(DATA_DIR, "bills");
const DATA_FILE = path.join(DATA_DIR, "payments.json");

[DATA_DIR, SLIPS_DIR, BILLS_DIR].forEach(function(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

function loadPayments() {
  try {
    if (!fs.existsSync(DATA_FILE)) return [];
    return JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  } catch (e) { return []; }
}

function savePayments(list) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(list, null, 2));
}

// ── State: normal mode pending bill ───────────────────────────────────────
const pendingBill = {};
function getPending(userId) {
  const s = pendingBill[userId];
  if (!s) return null;
  if (Date.now() > s.expiresAt) { delete pendingBill[userId]; return null; }
  return s;
}

// ── State: batch mode ─────────────────────────────────────────────────────
const batchState = {};
function getBatch(userId) {
  const s = batchState[userId];
  if (!s) return null;
  if (Date.now() > s.expiresAt) { delete batchState[userId]; return null; }
  return s;
}

// ── State: waiting for 1 or 2 answer (unknown image) ─────────────────────
// unknownState[userId] = { imageBuffer, groupId, expiresAt, batchContext }
const unknownState = {};
function getUnknown(userId) {
  const s = unknownState[userId];
  if (!s) return null;
  if (Date.now() > s.expiresAt) { delete unknownState[userId]; return null; }
  return s;
}

// ── Claude: detect image type ─────────────────────────────────────────────
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
{ type: "text", text: "Look at this image carefully and decide if it is a bank transfer confirmation slip or a bill/invoice.\n\nIt is a SLIP if it shows ANY of these:\n- 'Transfer Completed' or 'Transfer Successful'\n- 'โอนเงินสำเร็จ' or 'รายการสำเร็จ'\n- 'Transaction No' or 'Reference No' or 'เลขที่รายการ'\n- Sender account AND receiver account with bank names\n- FROM and TO with account numbers\n- Banks: SCB, KBank, KBIZ, GSB, Krungthai, Bangkok Bank, Krungsri, TMB, TTB, PromptPay\n\nIt is a BILL if it shows ANY of these:\n- QR code for payment (even if it shows an amount or bank account number)\n- Invoice or receipt with list of products/services\n- 'Please pay' or 'Amount due' or 'ยอดที่ต้องชำระ'\n- Bill with company name and itemized costs\n- Bank account number for receiving payment (but NO transfer confirmation)\n\nKEY RULE: A QR code or bank account number shown on a bill is NOT a slip. A slip must show transfer confirmation that money has already been sent.\n\nReturn ONLY one word: SLIP, BILL, or UNKNOWN" }          ]
        }]
      },
      { headers: { "x-api-key": ANTHROPIC_KEY, "anthropic-version": "2023-06-01", "content-type": "application/json" } }
    );
    const result = resp.data.content[0].text.trim().toUpperCase();
    console.log("Image type detected:", result);
    if (result.includes("SLIP")) return "SLIP";
    if (result.includes("BILL")) return "BILL";
    return "UNKNOWN";
  } catch (e) {
    console.error("Detect error:", e.message);
    return "UNKNOWN";
  }
}

// ── Claude: read slip details ─────────────────────────────────────────────
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

// ── Claude: read bill amount only ─────────────────────────────────────────
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
    return parsed.amount || null;
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

// ── Short reply after slip recorded ──────────────────────────────────────
function slipReply(data, id) {
  const amt = data.amount ? "฿" + Number(data.amount).toLocaleString("th-TH", { minimumFractionDigits: 2 }) : "Unknown";
  return [
    "✅ #" + id + "  " + (data.bank_from || "?") + "  " + amt,
    "→ " + (data.recipient_name || data.bank_to || "?") + (data.purpose ? "  (" + data.purpose + ")" : ""),
    "📎 Send bill or /skip"
  ].join("\n");
}

// ── Save slip to disk + payments.json ────────────────────────────────────
function saveSlip(buf, data, groupId, userId) {
  const fname = "slip_" + Date.now() + ".jpg";
  fs.writeFileSync(path.join(SLIPS_DIR, fname), buf);
  const list = loadPayments();
  const newId = list.length + 1;
  list.push(Object.assign({ id: newId, imageFile: fname, billFile: null, savedAt: new Date().toISOString(), groupId: groupId, userId: userId }, data));
  savePayments(list);
  return newId;
}

// ── Save bill to disk + link to payment ──────────────────────────────────
function saveBill(buf, paymentId) {
  const fname = "bill_" + Date.now() + ".jpg";
  fs.writeFileSync(path.join(BILLS_DIR, fname), buf);
  const list = loadPayments();
  const idx = list.findIndex(function(p) { return p.id === paymentId; });
  if (idx !== -1) { list[idx].billFile = fname; savePayments(list); }
  return fname;
}

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
  const baseUrl = process.env.DASHBOARD_URL || "https://slip-tracker-bot-production.up.railway.app";

  // ── Sheet 1: All Transactions ──
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
    var row = s1.addRow([p.id,p.transaction_date||"",p.transaction_time||"",Number(p.amount)||0,p.bank_from||"",p.account_from||"",p.bank_to||"",p.account_to||"",p.recipient_name||"",p.purpose||"",p.reference_number||"",(p.slip_type||"").replace(/_/g," "),p.billFile?"✅":"⚠️"]);
    if (i%2===0) row.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFF3F8FF"}};});
    row.getCell(4).numFmt="#,##0.00"; row.getCell(4).font={bold:true,color:{argb:"FF1565C0"}}; row.height=20;
  });
  var tr=s1.addRow(["","TOTAL","",grandTotal]);
  tr.getCell(4).numFmt="#,##0.00"; tr.getCell(4).font={bold:true,size:13};
  tr.eachCell(function(c){c.fill={type:"pattern",pattern:"solid",fgColor:{argb:"FFFFF9C4"}};});
  s1.columns=[{width:5},{width:12},{width:7},{width:14},{width:15},{width:11},{width:15},{width:11},{width:20},{width:28},{width:16},{width:15},{width:10}];
  s1.views=[{state:"frozen",ySplit:2}];

  // ── Sheet 2: Summary ──
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

  // ── Per-Account Sheets ──
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
    sa.columns=[{width:6},{width:12},{width:14},{width:30},{width:28},{width:28}];
    var ri=sa.rowCount+1;
    for(var pi=0;pi<grp.payments.length;pi++){
      var p=grp.payments[pi];
      sa.getRow(ri).height=22;
      sa.getCell("A"+ri).value="#"+p.id;
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

// ── LINE client ───────────────────────────────────────────────────────────
const client = new line.Client(LINE_CONFIG);

// ── Process image as SLIP ─────────────────────────────────────────────────
async function processAsSlip(buf, gid, uid, replyToken) {
  try {
    const data = await readSlip(buf.toString("base64"), "");
    if (!data) {
      return client.replyMessage(replyToken, { type: "text", text: "❌ Could not read slip. Please send a clearer image." });
    }
    const newId = saveSlip(buf, data, gid, uid);
    pendingBill[uid] = { paymentId: newId, expiresAt: Date.now() + 5 * 60 * 1000 };
    return client.pushMessage(gid, { type: "text", text: slipReply(data, newId) });
  } catch(err) {
    console.error("Slip error:", err.message);
    return client.pushMessage(gid, { type: "text", text: "❌ Error reading slip. Please try again." });
  }
}

// ── Process image as BILL (normal mode) ───────────────────────────────────
async function processAsBill(buf, paymentId, gid, replyToken) {
  saveBill(buf, paymentId);
  delete pendingBill[arguments[3] || ""];
  return client.replyMessage(replyToken, { type: "text", text: "📎 Bill saved for #" + paymentId + " ✅" });
}

// ── Process image as BILL (batch mode) ────────────────────────────────────
async function processAsBillBatch(buf, gid, uid, replyToken, batch) {
  const billAmount = await readBillAmount(buf.toString("base64"));
  if (!billAmount) {
    return client.replyMessage(replyToken, { type: "text", text: "⚠️ Could not read amount from bill — skipped." });
  }
  const list = loadPayments();
  const TOLERANCE = 1;
  var matchIdx = -1;
  for (var i = 0; i < batch.slipIds.length; i++) {
    var slipId = batch.slipIds[i];
    var slipIdx = list.findIndex(function(p) { return p.id === slipId; });
    if (slipIdx === -1 || list[slipIdx].billFile) continue;
    if (Math.abs((Number(list[slipIdx].amount)||0) - Number(billAmount)) <= TOLERANCE) { matchIdx = slipIdx; break; }
  }
  if (matchIdx === -1) {
    var billAmt = "฿" + Number(billAmount).toLocaleString("th-TH", { minimumFractionDigits: 2 });
    return client.replyMessage(replyToken, { type: "text", text: "⚠️ Bill " + billAmt + " — no matching slip found, skipped.\n\nSend more bills or type /done" });
  }
  var fname = "bill_" + Date.now() + ".jpg";
  fs.writeFileSync(path.join(BILLS_DIR, fname), buf);
  var matchedId = list[matchIdx].id;
  var matchedAmt = "฿" + Number(list[matchIdx].amount).toLocaleString("th-TH", { minimumFractionDigits: 2 });
  list[matchIdx].billFile = fname;
  savePayments(list);
  batchState[uid].expiresAt = Date.now() + 30 * 60 * 1000;
  return client.replyMessage(replyToken, { type: "text", text: "✅ Bill matched → #" + matchedId + " " + matchedAmt + "\n\nSend more or /done" });
}

async function handleEvent(event) {
  if (event.type !== "message") return;
  var uid = event.source.userId;
  var gid = event.source.groupId || event.source.roomId || uid;
  var pending = getPending(uid);
  var batch = getBatch(uid);
  var unknown = getUnknown(uid);

  // ════════════════════════════════════════════════
  // TEXT
  // ════════════════════════════════════════════════
  if (event.message.type === "text") {
    var txt = event.message.text.trim();
    var tl = txt.toLowerCase();

    // ── Answer to unknown image prompt (1 or 2) ──
    if (unknown) {
      if (txt === "1") {
        // User says it's a slip
        delete unknownState[uid];
        await client.replyMessage(event.replyToken, { type: "text", text: "🔍 Reading as slip..." });
        return processAsSlip(unknown.imageBuffer, gid, uid, event.replyToken);
      }
      if (txt === "2") {
        // User says it's a bill
        delete unknownState[uid];
        if (unknown.batchContext) {
          // In batch mode
          return processAsBillBatch(unknown.imageBuffer, gid, uid, event.replyToken, unknown.batchContext);
        } else if (pending) {
          // Normal mode
          saveBill(unknown.imageBuffer, pending.paymentId);
          delete pendingBill[uid];
          return client.replyMessage(event.replyToken, { type: "text", text: "📎 Bill saved for #" + pending.paymentId + " ✅" });
        } else {
          return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ No recent slip to attach this bill to. Send a slip first." });
        }
      }
      // If they typed something else, ignore the unknown state and continue
      delete unknownState[uid];
    }

    // ── /done in batch mode ──
    if (tl === "/done" && batch) {
      if (batch.step === "slips") {
        if (batch.slipIds.length === 0) {
          delete batchState[uid];
          return client.replyMessage(event.replyToken, { type: "text", text: "❌ No slips recorded. Batch cancelled." });
        }
        batchState[uid].step = "bills";
        batchState[uid].expiresAt = Date.now() + 30 * 60 * 1000;
        var list = loadPayments();
        var slipLines = batch.slipIds.map(function(id) {
          var p = list.find(function(x) { return x.id === id; });
          if (!p) return "#" + id + " ?";
          return "#" + id + "  ฿" + Number(p.amount||0).toLocaleString("th-TH") + "  " + (p.recipient_name || p.bank_to || "?");
        });
        return client.replyMessage(event.replyToken, {
          type: "text",
          text: ["✅ " + batch.slipIds.length + " slips recorded:", slipLines.join("\n"), "", "📎 Now send all BILL / INVOICE photos", "Bot will auto-match by amount", "Type /done when finished"].join("\n")
        });
      }
      if (batch.step === "bills") {
        var list = loadPayments();
        var matched = batch.slipIds.filter(function(id) { var p = list.find(function(x){return x.id===id;}); return p && p.billFile; });
        var unmatched = batch.slipIds.filter(function(id) { var p = list.find(function(x){return x.id===id;}); return p && !p.billFile; });
        var unmatchedLines = unmatched.map(function(id) {
          var p = list.find(function(x){return x.id===id;});
          return "  ⚠️ #" + id + " ฿" + Number(p?p.amount:0).toLocaleString("th-TH") + " — " + (p?(p.recipient_name||p.bank_to||"?"):"?");
        });
        delete batchState[uid];
        return client.replyMessage(event.replyToken, {
          type: "text",
          text: ["🎉 Batch complete!", "━━━━━━━━━━━━━━━━━━", "✅ Matched: " + matched.length + "/" + batch.slipIds.length, unmatched.length > 0 ? "⚠️ No bill (" + unmatched.length + "):\n" + unmatchedLines.join("\n") : "✅ All slips have evidence!", "", "Type /report to generate Excel 📊"].join("\n")
        });
      }
    }

    // ── /cancel batch ──
    if (tl === "/cancel" && batch) {
      delete batchState[uid];
      return client.replyMessage(event.replyToken, { type: "text", text: "❌ Batch cancelled." });
    }

    // ── Normal mode: text note after slip ──
    if (pending && !batch) {
      if (tl === "/skip") {
        delete pendingBill[uid];
        return client.replyMessage(event.replyToken, { type: "text", text: "⏭️ Skipped. You can add the bill later." });
      }
      var list = loadPayments();
      var idx = list.findIndex(function(p) { return p.id === pending.paymentId; });
      if (idx !== -1) { list[idx].purpose = txt; savePayments(list); }
      delete pendingBill[uid];
      return client.replyMessage(event.replyToken, { type: "text", text: "✅ Purpose saved: " + txt + "\n\nSend bill photo now, or /skip." });
    }

    // ── /batch ──
    if (tl === "/batch") {
      if (batch) return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ Already in batch mode! Send slips or type /done or /cancel." });
      batchState[uid] = { step: "slips", slipIds: [], expiresAt: Date.now() + 30 * 60 * 1000 };
      return client.replyMessage(event.replyToken, {
        type: "text",
        text: ["📦 Batch Mode ON", "━━━━━━━━━━━━━━━━━━", "📸 Send all SLIP photos now", "Type /done when all slips sent", "Type /cancel to exit", "", "⏱ Session expires in 30 min"].join("\n")
      });
    }

    // ── /help ──
    if (tl === "/help" || tl === "help") {
      return client.replyMessage(event.replyToken, {
        type: "text",
        text: ["💳 Slip Tracker Bot", "━━━━━━━━━━━━━━━━━━", "NORMAL MODE:", "📸 Send slip → bot auto-detects & records", "📸 Send bill → bot auto-detects & saves", "💬 Type note → sets purpose", "/skip → skip bill", "", "BATCH MODE:", "/batch → start batch", "📸📸📸 Send all slips", "/done → switch to bills", "📸📸📸 Send all bills (any order)", "/done → finish", "/cancel → exit batch", "", "/summary — this month totals", "/list — last 5 payments", "/report — Excel for this month", "/report YYYY MM — specific month", "/help — this menu"].join("\n")
      });
    }

    // ── /summary ──
    if (tl === "/summary" || tl === "summary") {
      var list = loadPayments(); var now = new Date();
      var thisMonth = list.filter(function(p) { if (!p.transaction_date) return false; var d = parseISO(p.transaction_date); return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth(); });
      var total = thisMonth.reduce(function(s,p) { return s+(Number(p.amount)||0); }, 0);
      var withBill = thisMonth.filter(function(p) { return p.billFile; }).length;
      var aMap = {}; thisMonth.forEach(function(p) { var k=(p.bank_from||"Unknown")+" ****"+(p.account_from||"????"); aMap[k]=(aMap[k]||0)+(Number(p.amount)||0); });
      var aLines = Object.entries(aMap).map(function(e) { return "  • "+e[0]+": ฿"+e[1].toLocaleString("th-TH",{minimumFractionDigits:2}); }).join("\n");
      return client.replyMessage(event.replyToken, { type: "text", text: ["📊 "+format(now,"MMMM yyyy")+" Summary","━━━━━━━━━━━━━━━━━━","📋 Transactions: "+thisMonth.length,"💰 Total: ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2}),"📎 With bill: "+withBill+"/"+thisMonth.length,"","By Account:",aLines||"  (none yet)","","Type /report to get Excel"].join("\n") });
    }

    // ── /list ──
    if (tl === "/list" || tl === "list") {
      var list = loadPayments();
      if (!list.length) return client.replyMessage(event.replyToken, { type: "text", text: "No payments recorded yet." });
      var recent = list.slice(-5).reverse();
      var lines = recent.map(function(p) { return ["#"+p.id+"  "+(p.transaction_date||"?")+"  ฿"+Number(p.amount||0).toLocaleString("th-TH"),"  "+(p.bank_from||"?")+" ****"+(p.account_from||"????"),"  "+(p.purpose||"-")+"  "+(p.billFile?"📎✅":"⚠️")].join("\n"); });
      return client.replyMessage(event.replyToken, { type: "text", text: "📋 Last 5:\n\n" + lines.join("\n\n") });
    }

    // ── /report ──
    if (tl.startsWith("/report") || tl === "report") {
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
        var baseUrl = process.env.DASHBOARD_URL || "https://slip-tracker-bot-production.up.railway.app";
        return client.pushMessage(gid, { type: "text", text: ["✅ Report ready — "+format(new Date(year,month-1),"MMMM yyyy"),"📋 "+filtered.length+" transactions","💰 ฿"+total.toLocaleString("th-TH",{minimumFractionDigits:2}),"🏦 "+Object.keys(accts).length+" account sheet(s)","📎 Bill evidence: "+withBill+"/"+filtered.length,"","⬇️ Download Excel:",baseUrl+"/api/report?year="+year+"&month="+month].join("\n") });
      } catch(err) {
        console.error("Report error:", err.message);
        return client.pushMessage(gid, { type: "text", text: "❌ Report error: " + err.message });
      }
    }
  }

  // ════════════════════════════════════════════════
  // IMAGE
  // ════════════════════════════════════════════════
  if (event.message.type === "image") {
    var buf = await getImage(client, event.message.id);
    var base64 = buf.toString("base64");

    // ── BATCH MODE ──
    if (batch) {
      if (batch.step === "slips") {
        // In batch slip collection — auto detect
        var type = await detectImageType(base64);
        if (type === "SLIP") {
          await client.replyMessage(event.replyToken, { type: "text", text: "🔍 Reading slip " + (batch.slipIds.length + 1) + "..." });
          try {
            var data = await readSlip(base64, "");
            if (!data) { client.pushMessage(gid, { type: "text", text: "⚠️ Could not read slip — skipped." }); return; }
            var newId = saveSlip(buf, data, gid, uid);
            batch.slipIds.push(newId);
            batchState[uid].expiresAt = Date.now() + 30 * 60 * 1000;
            var amt = data.amount ? "฿" + Number(data.amount).toLocaleString("th-TH") : "?";
            client.pushMessage(gid, { type: "text", text: "✅ Slip #" + newId + " — " + amt + " → " + (data.recipient_name || data.bank_to || "?") + "\n\nSend more slips or /done" });
          } catch(err) { client.pushMessage(gid, { type: "text", text: "⚠️ Error reading slip — try again." }); }
        } else if (type === "BILL") {
          // Bill sent during slip collection phase — save for later or notify
          client.replyMessage(event.replyToken, { type: "text", text: "⚠️ This looks like a bill — but we're still collecting slips!\nType /done first, then send your bills." });
        } else {
          // Unknown — ask
          unknownState[uid] = { imageBuffer: buf, groupId: gid, expiresAt: Date.now() + 2 * 60 * 1000, batchContext: batch };
          client.replyMessage(event.replyToken, { type: "text", text: "❓ Can't identify this image\nIs this a:\n1️⃣ Payment Slip\n2️⃣ Bill / Invoice\n\nReply 1 or 2" });
        }
        return;
      }

      if (batch.step === "bills") {
        // In batch bill collection — auto detect
        var type = await detectImageType(base64);
        if (type === "BILL" || type === "SLIP") {
          // Treat slip sent during bill phase as a bill attempt (read amount)
          await client.replyMessage(event.replyToken, { type: "text", text: "🔍 Reading bill amount..." });
          return processAsBillBatch(buf, gid, uid, event.replyToken, batch);
        } else {
          unknownState[uid] = { imageBuffer: buf, groupId: gid, expiresAt: Date.now() + 2 * 60 * 1000, batchContext: batch };
          client.replyMessage(event.replyToken, { type: "text", text: "❓ Can't identify this image\nIs this a:\n1️⃣ Payment Slip\n2️⃣ Bill / Invoice\n\nReply 1 or 2" });
        }
        return;
      }
    }

    // ── NORMAL MODE — auto detect ──
    var type = await detectImageType(base64);

    if (type === "SLIP") {
      await client.replyMessage(event.replyToken, { type: "text", text: "🔍 Reading slip..." });
      return processAsSlip(buf, gid, uid, event.replyToken);
    }

    if (type === "BILL") {
      if (pending) {
        saveBill(buf, pending.paymentId);
        delete pendingBill[uid];
        return client.replyMessage(event.replyToken, { type: "text", text: "📎 Bill saved for #" + pending.paymentId + " ✅" });
      } else {
        return client.replyMessage(event.replyToken, { type: "text", text: "⚠️ No recent slip to attach this bill to.\nSend your payment slip first!" });
      }
    }

    // Unknown image
    unknownState[uid] = { imageBuffer: buf, groupId: gid, expiresAt: Date.now() + 2 * 60 * 1000, batchContext: null };
    return client.replyMessage(event.replyToken, { type: "text", text: "❓ Can't identify this image\nIs this a:\n1️⃣ Payment Slip\n2️⃣ Bill / Invoice\n\nReply 1 or 2" });
  }
}

app.get("/", function(req, res) { res.send("✅ Slip Tracker Bot is running!"); });
app.use("/slips", express.static(SLIPS_DIR));
app.use("/bills", express.static(BILLS_DIR));
app.use("/dashboard", express.static(path.join(__dirname, "dashboard")));
app.get("/api/payments", function(req, res) { res.json(loadPayments()); });
app.get("/api/report", async function(req, res) {
  var now = new Date(); var year = parseInt(req.query.year) || now.getFullYear(); var month = parseInt(req.query.month) || now.getMonth() + 1;
  try { var file = await makeReport(year, month); res.download(file, "Payment_Report_"+year+"_"+String(month).padStart(2,"0")+".xlsx"); }
  catch(err) { res.status(500).send("Error: " + err.message); }
});

app.post("/webhook", express.json(), function(req, res) {
  console.log("WEBHOOK HIT!");
  res.status(200).end();
  if (!req.body || !req.body.events) return;
  req.body.events.forEach(function(event) {
    handleEvent(event).catch(function(err) { console.error("Event error:", err.message); });
  });
});

app.listen(PORT, function() { console.log("✅ Slip Tracker running on port " + PORT); });

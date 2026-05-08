"use strict";

const express = require("express");
const line = require("@line/bot-sdk");
const axios = require("axios");
const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const { format, parseISO } = require("date-fns");

// ── App setup ──────────────────────────────────────────────────────────────
const app = express();
const PORT = process.env.PORT || 8080;

const LINE_CONFIG = {
  channelAccessToken: process.env.LINE_CHANNEL_ACCESS_TOKEN || "",
  channelSecret: process.env.LINE_CHANNEL_SECRET || "",
};

const ANTHROPIC_KEY = process.env.ANTHROPIC_API_KEY || "";

// ── Data folders ───────────────────────────────────────────────────────────
const DATA_DIR   = path.join(__dirname, "data");
const SLIPS_DIR  = path.join(DATA_DIR, "slips");
const BILLS_DIR  = path.join(DATA_DIR, "bills");
const DATA_FILE  = path.join(DATA_DIR, "payments.json");

[DATA_DIR, SLIPS_DIR, BILLS_DIR].forEach(function(d) {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

// ── Payment storage ────────────────────────────────────────────────────────
function loadPayments() {
  try {
    if (!fs.existsSync(DATA_FILE)) return [];
    return JSON.parse(fs.readFileSync(DATA_FILE, "utf8"));
  } catch (e) {
    return [];
  }
}

function savePayments(list) {
  fs.writeFileSync(DATA_FILE, JSON.stringify(list, null, 2));
}

// ── Per-user state: waiting for bill image after slip ──────────────────────
const pendingBill = {};  // userId -> { paymentId, expiresAt }

function getPending(userId) {
  const s = pendingBill[userId];
  if (!s) return null;
  if (Date.now() > s.expiresAt) { delete pendingBill[userId]; return null; }
  return s;
}

// ── Claude: read the slip image ────────────────────────────────────────────
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
            {
              type: "image",
              source: { type: "base64", media_type: "image/jpeg", data: imageBase64 }
            },
            {
              type: "text",
text: "You are an expert Thai bank payment slip reader. Extract ALL details carefully.\n\nThai date formats to watch for:\n- DD/MM/YYYY or DD/MM/YY\n- DD MMM YYYY (e.g. 08 พ.ค. 2568 or 08 May 2026)\n- Buddhist year (พ.ศ.) — subtract 543 to get AD year. Example: 2568 = 2025, 2569 = 2026\n- Time is usually HH:MM or HH:MM:SS\n\nBanks: SCB, Krungthai (KTB), Bangkok Bank (BBL), Kasikorn (KBank), Krungsri (BAY), TMB, GSB, PromptPay.\n\nUser note: \"" + (caption || "") + "\"\n\nLook carefully at every part of the image for date and time. Return ONLY valid JSON:\n{\n  \"bank_from\": \"bank name or null\",\n  \"account_from\": \"last 4 digits or null\",\n  \"bank_to\": \"bank name or null\",\n  \"account_to\": \"last 4 digits or null\",\n  \"recipient_name\": \"name or null\",\n  \"amount\": 0.00,\n  \"transaction_date\": \"YYYY-MM-DD or null\",\n  \"transaction_time\": \"HH:MM or null\",\n  \"reference_number\": \"ref or null\",\n  \"purpose\": \"use user note if given, else infer from recipient name\",\n  \"slip_type\": \"mobile_banking or internet_banking or prompt_pay\"\n}"            }
          ]
        }]
      },
      {
        headers: {
          "x-api-key": ANTHROPIC_KEY,
          "anthropic-version": "2023-06-01",
          "content-type": "application/json"
        }
      }
    );
    const raw = resp.data.content[0].text.trim();
    console.log("Claude raw response:", raw);
    const clean = raw.replace(/```json/g, "").replace(/```/g, "").trim();
    return JSON.parse(clean);
  } catch (e) {
    console.error("Claude error:", e.message);
    console.error("Claude details:", e.response ? JSON.stringify(e.response.data) : "no response data");
    return null;
  }
}

// ── Download image from LINE ───────────────────────────────────────────────
async function getImage(client, messageId) {
  const stream = await client.getMessageContent(messageId);
  const parts = [];
  for await (const chunk of stream) parts.push(chunk);
  return Buffer.concat(parts);
}

// ── Format bot reply after reading slip ───────────────────────────────────
function slipReply(data, id) {
  const amt = data.amount
    ? "฿" + Number(data.amount).toLocaleString("th-TH", { minimumFractionDigits: 2 })
    : "Unknown";
  return [
    "✅ Slip Recorded  #" + id,
    "━━━━━━━━━━━━━━━━━━",
    "📅 " + (data.transaction_date || "-") + "  " + (data.transaction_time || ""),
    "💰 " + amt,
    "📤 " + (data.bank_from || "?") + " (" + (data.account_from || "****") + ")",
    "📥 " + (data.bank_to   || "?") + " (" + (data.account_to   || "****") + ")",
    "👤 " + (data.recipient_name || "-"),
    "📝 " + (data.purpose || "-"),
    "🔖 " + (data.reference_number || "-"),
    "",
    "📎 Send the BILL or INVOICE photo next",
    "   (or type a note, or send /skip)"
  ].join("\n");
}

// ── Generate Excel report ──────────────────────────────────────────────────
async function makeReport(year, month) {
  const all = loadPayments();
  const rows = all.filter(function(p) {
    if (!p.transaction_date) return false;
    const d = parseISO(p.transaction_date);
    return d.getFullYear() === year && d.getMonth() + 1 === month;
  });

  const wb = new ExcelJS.Workbook();
  const label = format(new Date(year, month - 1), "MMMM yyyy");

  // Sheet 1 – Transactions
  const s1 = wb.addWorksheet("Transactions");
  s1.mergeCells("A1:M1");
  const hdr = s1.getCell("A1");
  hdr.value = "Payment Report — " + label;
  hdr.font  = { size: 14, bold: true, color: { argb: "FFFFFFFF" } };
  hdr.fill  = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1565C0" } };
  hdr.alignment = { horizontal: "center" };
  s1.getRow(1).height = 30;

  const cols = ["#","Date","Time","Amount (฿)","From Bank","Acct","To Bank","Acct","Recipient","Purpose","Ref","Type","Evidence"];
  const hr = s1.addRow(cols);
  hr.eachCell(function(c) {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0D47A1" } };
    c.alignment = { horizontal: "center" };
  });

  var total = 0;
  rows.forEach(function(p, i) {
    const row = s1.addRow([
      p.id,
      p.transaction_date || "",
      p.transaction_time || "",
      Number(p.amount) || 0,
      p.bank_from || "",
      p.account_from || "",
      p.bank_to || "",
      p.account_to || "",
      p.recipient_name || "",
      p.purpose || "",
      p.reference_number || "",
      (p.slip_type || "").replace(/_/g, " "),
      p.billFile ? "✅ Sheet 4" : "⚠️ Missing"
    ]);
    total += Number(p.amount) || 0;
    if (i % 2 === 0) row.eachCell(function(c) {
      c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF3F8FF" } };
    });
    row.getCell(4).numFmt = "#,##0.00";
    row.height = 20;
  });

  const tr = s1.addRow(["", "TOTAL", "", total]);
  tr.getCell(4).numFmt = "#,##0.00";
  tr.getCell(4).font = { bold: true };
  tr.eachCell(function(c) {
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF9C4" } };
  });

  s1.columns = [
    {width:5},{width:12},{width:7},{width:14},{width:15},{width:11},
    {width:15},{width:11},{width:20},{width:28},{width:16},{width:15},{width:13}
  ];

  // Sheet 2 – By Bank
  const s2 = wb.addWorksheet("By Bank");
  s2.addRow(["Bank", "Count", "Total (฿)", "% of Total"]).eachCell(function(c) {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1565C0" } };
  });
  const bankMap = {};
  rows.forEach(function(p) {
    const k = p.bank_from || "Unknown";
    if (!bankMap[k]) bankMap[k] = { count: 0, sum: 0 };
    bankMap[k].count++;
    bankMap[k].sum += Number(p.amount) || 0;
  });
  Object.entries(bankMap).sort(function(a,b){ return b[1].sum - a[1].sum; }).forEach(function(e) {
    const r = s2.addRow([e[0], e[1].count, e[1].sum, total > 0 ? (e[1].sum/total*100).toFixed(1)+"%" : "0%"]);
    r.getCell(3).numFmt = "#,##0.00";
  });
  s2.columns = [{width:22},{width:10},{width:18},{width:12}];

  // Sheet 3 – By Purpose
  const s3 = wb.addWorksheet("By Purpose");
  s3.addRow(["Purpose", "Count", "Total (฿)", "% of Total"]).eachCell(function(c) {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1565C0" } };
  });
  const purMap = {};
  rows.forEach(function(p) {
    const k = p.purpose || "Unspecified";
    if (!purMap[k]) purMap[k] = { count: 0, sum: 0 };
    purMap[k].count++;
    purMap[k].sum += Number(p.amount) || 0;
  });
  Object.entries(purMap).sort(function(a,b){ return b[1].sum - a[1].sum; }).forEach(function(e) {
    const r = s3.addRow([e[0], e[1].count, e[1].sum, total > 0 ? (e[1].sum/total*100).toFixed(1)+"%" : "0%"]);
    r.getCell(3).numFmt = "#,##0.00";
  });
  s3.columns = [{width:35},{width:10},{width:18},{width:12}];

  // Sheet 4 – Evidence Gallery
  const s4 = wb.addWorksheet("Evidence Gallery");
  s4.addRow(["ID","Date","Amount","Purpose","Payment Slip","Bill / Invoice"]).eachCell(function(c) {
    c.font = { bold: true, color: { argb: "FFFFFFFF" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF0D47A1" } };
    c.alignment = { horizontal: "center" };
  });
  s4.columns = [{width:6},{width:12},{width:14},{width:30},{width:24},{width:24}];

  var ri = 2;
  for (var i = 0; i < rows.length; i++) {
    const p = rows[i];
    s4.getRow(ri).height = 98;
    s4.getCell("A"+ri).value = "#" + p.id;
    s4.getCell("B"+ri).value = p.transaction_date || "";
    s4.getCell("C"+ri).value = Number(p.amount) || 0;
    s4.getCell("C"+ri).numFmt = "#,##0.00";
    s4.getCell("D"+ri).value = p.purpose || "";

    const sp = p.imageFile ? path.join(SLIPS_DIR, p.imageFile) : null;
    if (sp && fs.existsSync(sp)) {
      try {
        s4.addImage(wb.addImage({ filename: sp, extension: "jpeg" }), {
          tl: { col: 4, row: ri-1 }, br: { col: 5, row: ri }, editAs: "oneCell"
        });
      } catch(e) { s4.getCell("E"+ri).value = "Image error"; }
    } else {
      s4.getCell("E"+ri).value = "No slip saved";
    }

    const bp = p.billFile ? path.join(BILLS_DIR, p.billFile) : null;
    if (bp && fs.existsSync(bp)) {
      try {
        s4.addImage(wb.addImage({ filename: bp, extension: "jpeg" }), {
          tl: { col: 5, row: ri-1 }, br: { col: 6, row: ri }, editAs: "oneCell"
        });
      } catch(e) { s4.getCell("F"+ri).value = "Image error"; }
    } else {
      s4.getCell("F"+ri).value = "No bill uploaded";
      s4.getCell("F"+ri).font = { italic: true, color: { argb: "FFBF360C" } };
    }
    ri++;
  }

  const outPath = path.join(DATA_DIR, "report_" + year + "_" + String(month).padStart(2,"0") + ".xlsx");
  await wb.xlsx.writeFile(outPath);
  return outPath;
}

// ── LINE client & event handler ────────────────────────────────────────────
const client = new line.Client(LINE_CONFIG);

async function handleEvent(event) {
  if (event.type !== "message") return;

  const uid    = event.source.userId;
  const gid    = event.source.groupId || event.source.roomId || uid;
  const pending = getPending(uid);

  // ── IMAGE ──────────────────────────────────────────────────────────────
  if (event.message.type === "image") {

    // Second image = bill evidence
    if (pending) {
      const buf = await getImage(client, event.message.id);
      const fname = "bill_" + Date.now() + ".jpg";
      fs.writeFileSync(path.join(BILLS_DIR, fname), buf);

      const list = loadPayments();
      const idx = list.findIndex(function(p) { return p.id === pending.paymentId; });
      if (idx !== -1) { list[idx].billFile = fname; savePayments(list); }
      delete pendingBill[uid];

      return client.replyMessage(event.replyToken, {
        type: "text",
        text: "📎 Bill saved for #" + pending.paymentId + "\n✅ Done! Both images will appear in the report."
      });
    }

    // First image = payment slip
    await client.replyMessage(event.replyToken, {
      type: "text", text: "🔍 Reading your slip..."
    });

    try {
      const buf  = await getImage(client, event.message.id);
      const data = await readSlip(buf.toString("base64"), "");

      const fname = "slip_" + Date.now() + ".jpg";
      fs.writeFileSync(path.join(SLIPS_DIR, fname), buf);

      const list = loadPayments();
      const newId = list.length + 1;
      list.push(Object.assign({ id: newId, imageFile: fname, billFile: null, savedAt: new Date().toISOString() }, data));
      savePayments(list);

      pendingBill[uid] = { paymentId: newId, expiresAt: Date.now() + 5 * 60 * 1000 };

      return client.pushMessage(gid, {
        type: "text", text: slipReply(data, newId)
      });
    } catch (err) {
      console.error("Slip error:", err.message);
      return client.pushMessage(gid, {
        type: "text", text: "❌ Could not read slip. Please send a clearer image."
      });
    }
  }

  // ── TEXT ───────────────────────────────────────────────────────────────
  if (event.message.type === "text") {
    const txt = event.message.text.trim();

    // Update purpose or skip while waiting for bill
    if (pending) {
      if (txt === "/skip") {
        delete pendingBill[uid];
        return client.replyMessage(event.replyToken, {
          type: "text", text: "⏭️ Skipped. You can add the bill later from the dashboard."
        });
      }
      const list = loadPayments();
      const idx = list.findIndex(function(p) { return p.id === pending.paymentId; });
      if (idx !== -1) { list[idx].purpose = txt; savePayments(list); }
      delete pendingBill[uid];
      return client.replyMessage(event.replyToken, {
        type: "text", text: "✅ Purpose saved: " + txt + "\n\nSend bill photo now, or /skip."
      });
    }

    // Commands
    if (txt === "/help" || txt === "help") {
      return client.replyMessage(event.replyToken, {
        type: "text",
        text: [
          "💳 Slip Tracker Bot",
          "━━━━━━━━━━━━━━━━━━",
          "📸 Send slip photo → bot reads it",
          "📸 Send bill photo → saved as evidence",
          "💬 Type note after slip → sets purpose",
          "/skip → skip bill upload",
          "",
          "/summary  — this month totals",
          "/list     — last 5 payments",
          "/report   — Excel for this month",
          "/help     — this menu"
        ].join("\n")
      });
    }

    if (txt === "/summary") {
      const list = loadPayments();
      const now  = new Date();
      const month = list.filter(function(p) {
        if (!p.transaction_date) return false;
        const d = parseISO(p.transaction_date);
        return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
      });
      const total = month.reduce(function(s,p) { return s + (Number(p.amount)||0); }, 0);
      const withBill = month.filter(function(p) { return p.billFile; }).length;
      const byBank = {};
      month.forEach(function(p) {
        const k = p.bank_from || "Unknown";
        byBank[k] = (byBank[k] || 0) + (Number(p.amount) || 0);
      });
      const bankLines = Object.entries(byBank)
        .map(function(e) { return "  • " + e[0] + ": ฿" + e[1].toLocaleString("th-TH",{minimumFractionDigits:2}); })
        .join("\n");

      return client.replyMessage(event.replyToken, {
        type: "text",
        text: [
          "📊 " + format(now, "MMMM yyyy") + " Summary",
          "━━━━━━━━━━━━━━━━━━",
          "📋 Transactions: " + month.length,
          "💰 Total: ฿" + total.toLocaleString("th-TH",{minimumFractionDigits:2}),
          "📎 With bill: " + withBill + "/" + month.length,
          "",
          "By Bank:",
          bankLines || "  (none yet)",
          "",
          "Type /report to get Excel file"
        ].join("\n")
      });
    }

    if (txt === "/list") {
      const list = loadPayments();
      if (!list.length) {
        return client.replyMessage(event.replyToken, { type: "text", text: "No payments recorded yet." });
      }
      const recent = list.slice(-5).reverse();
      const lines = recent.map(function(p) {
        return "#" + p.id + "  " + (p.transaction_date||"?") + "  ฿" + Number(p.amount||0).toLocaleString("th-TH") + "\n  " + (p.purpose||"-") + "  " + (p.billFile ? "📎✅" : "⚠️");
      });
      return client.replyMessage(event.replyToken, {
        type: "text", text: "📋 Last 5:\n\n" + lines.join("\n\n")
      });
    }

    if (txt.startsWith("/report")) {
      const parts = txt.split(" ");
      const now   = new Date();
      const year  = parseInt(parts[1]) || now.getFullYear();
      const month = parseInt(parts[2]) || now.getMonth() + 1;

      await client.replyMessage(event.replyToken, {
        type: "text", text: "📊 Generating report..."
      });

      try {
        await makeReport(year, month);
        const list = loadPayments();
        const filtered = list.filter(function(p) {
          if (!p.transaction_date) return false;
          const d = parseISO(p.transaction_date);
          return d.getFullYear() === year && d.getMonth()+1 === month;
        });
        const total = filtered.reduce(function(s,p){ return s+(Number(p.amount)||0); },0);
        const withBill = filtered.filter(function(p){ return p.billFile; }).length;
        const baseUrl = process.env.DASHBOARD_URL || "https://slip-tracker-bot-production.up.railway.app";
        return client.pushMessage(gid, {
          type: "text",
          text: [
            "✅ Report ready — " + format(new Date(year, month-1), "MMMM yyyy"),
            "📋 " + filtered.length + " transactions",
            "💰 ฿" + total.toLocaleString("th-TH",{minimumFractionDigits:2}),
            "📎 Bill evidence: " + withBill + "/" + filtered.length,
            "",
            "Download Excel:",
            baseUrl + "/api/report?year=" + year + "&month=" + month
          ].join("\n")
        });
      } catch (err) {
        console.error("Report error:", err.message);
        return client.pushMessage(gid, { type: "text", text: "❌ Report error: " + err.message });
      }
    }
  }
}

// ── Routes ─────────────────────────────────────────────────────────────────
app.get("/", function(req, res) {
  res.send("✅ Slip Tracker Bot is running!");
});

app.post("/test", express.json(), function(req, res) {
  console.log("TEST HIT:", JSON.stringify(req.body));
  res.send("ok");
});

app.use("/slips",     express.static(SLIPS_DIR));
app.use("/bills",     express.static(BILLS_DIR));
app.use("/dashboard", express.static(path.join(__dirname, "dashboard")));

app.get("/api/payments", function(req, res) {
  res.json(loadPayments());
});

app.get("/api/report", async function(req, res) {
  const now   = new Date();
  const year  = parseInt(req.query.year)  || now.getFullYear();
  const month = parseInt(req.query.month) || now.getMonth() + 1;
  try {
    const file = await makeReport(year, month);
    res.download(file, "Payment_Report_" + year + "_" + String(month).padStart(2,"0") + ".xlsx");
  } catch (err) {
    res.status(500).send("Error: " + err.message);
  }
});

// Webhook — must use raw body for LINE signature check
app.post(
  "/webhook",
  express.json(),
  function(req, res) {
    console.log("WEBHOOK HIT!", JSON.stringify(req.body));
    res.status(200).end();
    if (!req.body.events) return;
    req.body.events.forEach(function(event) {
      handleEvent(event).catch(function(err) {
        console.error("Event error:", err.message);
      });
    });
  }
);

// ── Start ──────────────────────────────────────────────────────────────────
app.listen(PORT, function() {
  console.log("✅ Slip Tracker running on port " + PORT);
});

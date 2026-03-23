/**
 * @OnlyCurrentDoc
 * Mail Merge with Open Tracking, Bounce Detection, and Quota Management
 */

// ─── Column names — change these to match your sheet headers ─────────────────
const RECIPIENT_COL  = "Email";
const EMAIL_SENT_COL = "Merge Status";
const OPEN_COUNT_COL = "Opens";
const LAST_OPEN_COL  = "Last Opened";
const BOUNCE_COL     = "Bounced";

// ─── Paste your Web App URL here after deploying ─────────────────────────────
// Deploy > New Deployment > Web App > Execute as Me > Access Anyone > Deploy
const TRACKING_BASE_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";      // Update your ID here

// ─── Required columns in order ───────────────────────────────────────────────
const REQUIRED_COLS = [
  RECIPIENT_COL,   // "Email"
  EMAIL_SENT_COL,  // "Merge Status"
  OPEN_COUNT_COL,  // "Opens"
  LAST_OPEN_COL,   // "Last Opened"
  BOUNCE_COL,      // "Bounced"
];

// ─── Auto-setup missing columns ──────────────────────────────────────────────
function setupColumns_() {
  const sheet    = SpreadsheetApp.getActiveSheet();
  const lastCol  = sheet.getLastColumn();

  // If sheet is completely empty, write all headers from scratch in row 1
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, REQUIRED_COLS.length).setValues([REQUIRED_COLS]);
    formatHeaders_(sheet, 1, REQUIRED_COLS.length);
    return;
  }

  // Read existing headers
  const heads = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  // Add only the ones that are missing, appended to the right
  let nextCol = lastCol + 1;
  REQUIRED_COLS.forEach(col => {
    if (!heads.includes(col)) {
      sheet.getRange(1, nextCol).setValue(col);
      formatHeaders_(sheet, nextCol, 1);
      nextCol++;
    }
  });

  formatBounceColumn_(sheet, heads);
}

// ─── Style the header cells ───────────────────────────────────────────────────
function formatHeaders_(sheet, startCol, count) {
  const range = sheet.getRange(1, startCol, 1, count);
  range
    .setBackground("#1a73e8")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontSize(11)
    .setBorder(true, true, true, true, false, false, "#1558b0",
      SpreadsheetApp.BorderStyle.SOLID);

  // Set sensible column widths per header
  for (let i = 0; i < count; i++) {
    const col  = startCol + i;
    const name = sheet.getRange(1, col).getValue();
    const widthMap = {
      [RECIPIENT_COL]:  220,
      [EMAIL_SENT_COL]: 160,
      [OPEN_COUNT_COL]:  80,
      [LAST_OPEN_COL]:  170,
      [BOUNCE_COL]:     100,
    };
    sheet.setColumnWidth(col, widthMap[name] || 150);
  }
}

// ─── Menu ─────────────────────────────────────────────────────────────────────
function onOpen() {
  setupColumns_();  // ← auto-creates missing columns on every open

  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Mail Merge")
    .addItem("Send Emails",        "sendEmails")
    .addItem("Check Bounces",      "checkBounces")
    .addItem("Check Quota",        "showQuotaStatus")
    .addItem("Debug Placeholders", "debugTemplatePlaceholders")
    .addItem("Campaign Summary",   "showCampaignSummary")
    .addToUi();
}



// ─── Web App entry point (tracking pixel) ────────────────────────────────────
function doGet(e) {
  try {
    const uid         = e.parameter.uid;
    const sheet       = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const heads       = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const openCol     = heads.indexOf(OPEN_COUNT_COL) + 1;
    const lastOpenCol = heads.indexOf(LAST_OPEN_COL) + 1;
    const dataRow     = parseInt(uid) + 1;

    if (openCol > 0) {
      const current = sheet.getRange(dataRow, openCol).getValue() || 0;
      sheet.getRange(dataRow, openCol).setValue(current + 1);
    }

    if (lastOpenCol > 0) {
      const now       = new Date();
      const formatted = Utilities.formatDate(
        now,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm:ss"
      );
      sheet.getRange(dataRow, lastOpenCol).setValue(formatted);
    }
  } catch (err) {
    console.error("Tracking error:", err);
  }

  return _servePixel();
}

function _servePixel() {
  const pixel = Utilities.base64Decode(
    "R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"
  );
  return ContentService.createTextOutput(
    String.fromCharCode(...pixel)
  ).setMimeType(ContentService.MimeType.GIF);
}


// ─── Main send function ───────────────────────────────────────────────────────
function sendEmails(subjectLine, sheet = SpreadsheetApp.getActiveSheet()) {
  const ui = SpreadsheetApp.getUi();

  let processedSubjectLine = subjectLine;
  if (!processedSubjectLine) {
    processedSubjectLine = Browser.inputBox(
      "Mail Merge",
      "Type or copy/paste the subject line of the Gmail draft:",
      Browser.Buttons.OK_CANCEL
    );
    if (processedSubjectLine === "cancel" || processedSubjectLine === "") return;
  }

  if (TRACKING_BASE_URL === "YOUR_WEB_APP_URL_HERE") {
    ui.alert(
      "⚠️ Missing Tracking URL",
      "Please deploy this script as a Web App first and paste the URL into TRACKING_BASE_URL.\n\n" +
      "Steps:\n" +
      "1. Click Deploy > New Deployment\n" +
      "2. Type: Web App\n" +
      "3. Execute as: Me | Access: Anyone\n" +
      "4. Copy the URL and paste it into the TRACKING_BASE_URL constant.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Show quota before sending
  const quotaInfo = getRemainingQuota_();
  const proceed   = ui.alert(
    "📊 Daily Quota Status",
    `Emails sent today : ${quotaInfo.sent}\n` +
    `Remaining today   : ${quotaInfo.remaining}\n` +
    `Daily limit       : ${quotaInfo.limit}\n\n` +
    (quotaInfo.remaining === 0
      ? "❌ You've hit today's limit. Try again tomorrow."
      : `✅ Ready to send. Click OK to continue.`),
    ui.ButtonSet.OK_CANCEL
  );

  if (proceed !== ui.Button.OK) return;
  if (quotaInfo.remaining === 0) return;

  const emailTemplate   = getGmailTemplateFromDrafts_(processedSubjectLine);
  const dataRange       = sheet.getDataRange();
  const data            = dataRange.getDisplayValues();
  const heads           = data.shift();

  const emailSentColIdx = heads.indexOf(EMAIL_SENT_COL);
  const openColIdx      = heads.indexOf(OPEN_COUNT_COL);

  if (emailSentColIdx === -1) {
    ui.alert("⚠️ Column Not Found", `Could not find a column named "${EMAIL_SENT_COL}" in your sheet.`, ui.ButtonSet.OK);
    return;
  }

  const obj = data.map((r) =>
    heads.reduce((o, k, i) => { o[k] = r[i] || ""; return o; }, {})
  );

  const out         = [];
  let sentThisRun   = 0;
  let skippedCount  = 0;
  let remainingRows = 0;

  obj.forEach((row, rowIdx) => {
    if (row[EMAIL_SENT_COL] === "") {

      // Stop sending if quota reached — leave cell blank for tomorrow
      if (sentThisRun >= quotaInfo.remaining) {
        out.push([""]);
        remainingRows++;
        return;
      }

      try {
        const msgObj = fillInTemplateFromObject_(emailTemplate.message, row);

        // Inject tracking pixel
        const trackingPixel =
          `<img src="${TRACKING_BASE_URL}?uid=${rowIdx + 1}" ` +
          `width="1" height="1" style="display:none" alt="" />`;
        const htmlWithPixel = (msgObj.html || "") + trackingPixel;

        GmailApp.sendEmail(row[RECIPIENT_COL], msgObj.subject, msgObj.text, {
          htmlBody:     htmlWithPixel,
          attachments:  emailTemplate.attachments,
          inlineImages: emailTemplate.inlineImages,
        });

        out.push([`Sent: ${new Date().toLocaleDateString()}`]);
        sentThisRun++;

      } catch (e) {
        if (e.message.toLowerCase().includes("quota") ||
            e.message.toLowerCase().includes("limit")) {
          // Quota hit mid-run — leave blank so it retries tomorrow
          out.push([""]);
          remainingRows++;
        } else {
          // Real error (bad address etc.) — log it so you can investigate
          out.push([`Error: ${e.message}`]);
        }
      }

    } else {
      out.push([row[EMAIL_SENT_COL]]);
      skippedCount++;
    }
  });

  sheet.getRange(2, emailSentColIdx + 1, out.length).setValues(out);

  // Initialise Opens column to 0 for newly sent rows
  if (openColIdx > -1) {
    const openRange = sheet.getRange(2, openColIdx + 1, out.length);
    const existing  = openRange.getValues();
    const init      = existing.map(([v]) => [v === "" ? 0 : v]);
    openRange.setValues(init);
  }

  ui.alert(
    "✅ Mail Merge Complete",
    `Sent this run      : ${sentThisRun}\n` +
    `Already sent before: ${skippedCount}\n` +
    `Pending (tomorrow) : ${remainingRows}\n\n` +
    (remainingRows > 0
      ? `⏳ Run again tomorrow to send the remaining ${remainingRows} emails.`
      : "🎉 All emails have been sent!"),
    ui.ButtonSet.OK
  );
}
// ─── Campaign Summary ───────────────────────────────────────────────────────
function showCampaignSummary() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  const runBounceCheck = ui.alert(
    "🔄 Refresh bounce data?",
    "Click YES to check for bounces first (recommended), or NO to use existing data.",
    ui.ButtonSet.YES_NO
  );

  if (runBounceCheck === ui.Button.YES) {
    checkBounces();
    // Force Sheets to flush all pending writes before we read data back
    SpreadsheetApp.flush();
  }

  // Read data AFTER flush so we always get the latest values
  const data  = sheet.getDataRange().getValues();
  const heads = data.shift();

  const emailIdx  = heads.indexOf(RECIPIENT_COL);
  const statusIdx = heads.indexOf(EMAIL_SENT_COL);
  const opensIdx  = heads.indexOf(OPEN_COUNT_COL);
  const lastIdx   = heads.indexOf(LAST_OPEN_COL);
  const bounceIdx = heads.indexOf(BOUNCE_COL);

  const rows = data
    .filter(r => (r[emailIdx] || "").toString().trim() !== "")
    .map(r => {
      // Normalize bounce value — Sheets may return boolean true/false or string "TRUE"/"FALSE"
      const rawBounce = r[bounceIdx];
      let bounceVal   = "";
      if (rawBounce === true  || String(rawBounce).trim().toUpperCase() === "TRUE")  bounceVal = "TRUE";
      if (rawBounce === false || String(rawBounce).trim().toUpperCase() === "FALSE") bounceVal = "FALSE";

      return {
        email:   (r[emailIdx]  || "").toString(),
        status:  (r[statusIdx] || "").toString(),
        opens:    r[opensIdx]  || 0,
        last:    (r[lastIdx]   || "").toString(),
        bounced:  bounceVal,   // always "", "TRUE", or "FALSE" — never raw boolean
      };
    });

  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head><base target="_top">
    <style>
      * { box-sizing:border-box; margin:0; padding:0; font-family:Arial,sans-serif; }
      body { background:#fff; color:#1a1a1a; font-size:13px; }
      .dashboard { padding:20px; }
      .header h2 { font-size:16px; font-weight:600; margin-bottom:4px; }
      .header p  { font-size:12px; color:#666; margin-bottom:16px; }
      .metrics { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; margin-bottom:20px; }
      .metric { background:#f5f5f5; border-radius:8px; padding:12px 14px; }
      .metric-label { font-size:11px; color:#888; margin-bottom:4px; }
      .metric-value { font-size:22px; font-weight:600; }
      .metric-sub   { font-size:11px; color:#aaa; margin-top:2px; }
      .metric.sent    .metric-value { color:#185FA5; }
      .metric.opened  .metric-value { color:#0F6E56; }
      .metric.bounced .metric-value { color:#993C1D; }
      .metric.pending .metric-value { color:#854F0B; }
      .bars { margin-bottom:20px; display:flex; flex-direction:column; gap:12px; }
      .bar-row { display:flex; flex-direction:column; gap:4px; }
      .bar-meta { display:flex; justify-content:space-between; font-size:12px; color:#555; }
      .bar-track { height:10px; background:#eee; border-radius:99px; overflow:hidden; }
      .bar-fill  { height:100%; border-radius:99px; transition:width 0.8s cubic-bezier(.4,0,.2,1); }
      .bar-fill.sent    { background:#185FA5; }
      .bar-fill.opened  { background:#0F6E56; }
      .bar-fill.bounced { background:#993C1D; }
      .bar-fill.pending { background:#BA7517; }
      .warn-banner { background:#FFF8E1; border:1px solid #FFD54F; border-radius:6px; padding:10px 14px; margin-bottom:16px; font-size:12px; color:#5D4037; display:none; }
      .table-wrap { border:1px solid #e5e5e5; border-radius:8px; overflow:hidden; }
      table { width:100%; border-collapse:collapse; font-size:12px; table-layout:fixed; }
      thead th { background:#f9f9f9; padding:8px 12px; text-align:left; font-weight:600; color:#666; border-bottom:1px solid #e5e5e5; }
      tbody td { padding:8px 12px; border-bottom:1px solid #f0f0f0; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
      tbody tr:last-child td { border-bottom:none; }
      tbody tr:hover td { background:#fafafa; }
      .badge { display:inline-block; padding:2px 8px; border-radius:99px; font-size:11px; font-weight:600; }
      .badge.sent    { background:#E6F1FB; color:#0C447C; }
      .badge.opened  { background:#EAF3DE; color:#3B6D11; }
      .badge.bounced { background:#FAECE7; color:#993C1D; }
      .badge.error   { background:#FAEEDA; color:#854F0B; }
      .badge.pending { background:#f0f0f0; color:#888; }
    </style></head><body>
    <div class="dashboard">
      <div id="warn-banner" class="warn-banner">
        ⚠️ Some sent rows have not been bounce-checked yet. Run <strong>Check Bounces</strong> from the Mail Merge menu for accurate results.
      </div>
      <div class="header">
        <h2>Campaign summary</h2>
        <p id="subtitle"></p>
      </div>
      <div class="metrics">
        <div class="metric sent">
          <div class="metric-label">Delivered</div>
          <div class="metric-value" id="m-sent">—</div>
          <div class="metric-sub" id="m-sent-sub"></div>
        </div>
        <div class="metric opened">
          <div class="metric-label">Opened</div>
          <div class="metric-value" id="m-opened">—</div>
          <div class="metric-sub" id="m-open-sub"></div>
        </div>
        <div class="metric bounced">
          <div class="metric-label">Bounced</div>
          <div class="metric-value" id="m-bounced">—</div>
          <div class="metric-sub" id="m-bounce-sub"></div>
        </div>
        <div class="metric pending">
          <div class="metric-label">Pending</div>
          <div class="metric-value" id="m-pending">—</div>
          <div class="metric-sub">not yet sent</div>
        </div>
      </div>
      <div class="bars">
        <div class="bar-row">
          <div class="bar-meta"><span>Delivered (no bounce)</span><span id="pct-sent">0%</span></div>
          <div class="bar-track"><div class="bar-fill sent"    id="bar-sent"    style="width:0%"></div></div>
        </div>
        <div class="bar-row">
          <div class="bar-meta"><span>Opened (of delivered)</span><span id="pct-opened">0%</span></div>
          <div class="bar-track"><div class="bar-fill opened"  id="bar-opened"  style="width:0%"></div></div>
        </div>
        <div class="bar-row">
          <div class="bar-meta"><span>Bounced (of attempted)</span><span id="pct-bounced">0%</span></div>
          <div class="bar-track"><div class="bar-fill bounced" id="bar-bounced" style="width:0%"></div></div>
        </div>
        <div class="bar-row">
          <div class="bar-meta"><span>Pending</span><span id="pct-pending">0%</span></div>
          <div class="bar-track"><div class="bar-fill pending" id="bar-pending" style="width:0%"></div></div>
        </div>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr>
            <th style="width:35%">Email</th>
            <th style="width:17%">Status</th>
            <th style="width:10%">Opens</th>
            <th style="width:22%">Last opened</th>
            <th style="width:16%">Bounced</th>
          </tr></thead>
          <tbody id="tbl-body"></tbody>
        </table>
      </div>
    </div>
    <script>
      const rows = ${JSON.stringify(rows)};

      function classify(r) {
        const b = String(r.bounced).trim().toUpperCase();
        if (b === "TRUE")  return "bounced";
        if (!r.status)     return "pending";
        if (r.status.startsWith("Error")) return "error";
        if (r.status.startsWith("Sent"))  return r.opens > 0 ? "opened" : "sent";
        return "pending";
      }

      const total     = rows.length;
      const bounced   = rows.filter(r => classify(r) === "bounced").length;
      const opened    = rows.filter(r => classify(r) === "opened").length;
      const sent      = rows.filter(r => ["sent","opened"].includes(classify(r))).length;
      const pending   = rows.filter(r => classify(r) === "pending").length;
      const attempted = sent + bounced;

      // Warn if any sent row still has blank bounce value
      const sentRows         = rows.filter(r => r.status.startsWith("Sent"));
      const uncheckedBounces = sentRows.filter(r => r.bounced === "").length;
      if (uncheckedBounces > 0) {
        document.getElementById("warn-banner").style.display = "block";
      }

      const pSent    = total     ? Math.round(sent    / total     * 100) : 0;
      const pOpened  = sent      ? Math.round(opened  / sent      * 100) : 0;
      const pBounced = attempted ? Math.round(bounced / attempted * 100) : 0;
      const pPending = total     ? Math.round(pending / total     * 100) : 0;

      document.getElementById("subtitle").textContent     = total + " recipients · " + attempted + " attempted · " + pending + " pending";
      document.getElementById("m-sent").textContent       = sent;
      document.getElementById("m-opened").textContent     = opened;
      document.getElementById("m-bounced").textContent    = bounced;
      document.getElementById("m-pending").textContent    = pending;
      document.getElementById("m-sent-sub").textContent   = pSent    + "% delivered";
      document.getElementById("m-open-sub").textContent   = pOpened  + "% of delivered";
      document.getElementById("m-bounce-sub").textContent = pBounced + "% of attempted";
      document.getElementById("pct-sent").textContent     = pSent    + "%";
      document.getElementById("pct-opened").textContent   = pOpened  + "%";
      document.getElementById("pct-bounced").textContent  = pBounced + "%";
      document.getElementById("pct-pending").textContent  = pPending + "%";

      setTimeout(() => {
        document.getElementById("bar-sent").style.width    = pSent    + "%";
        document.getElementById("bar-opened").style.width  = pOpened  + "%";
        document.getElementById("bar-bounced").style.width = pBounced + "%";
        document.getElementById("bar-pending").style.width = pPending + "%";
      }, 100);

      const labels = { sent:"Sent", opened:"Opened", bounced:"Bounced", error:"Error", pending:"Pending" };
      document.getElementById("tbl-body").innerHTML = rows.map(r => {
        const type = classify(r);
        const b    = String(r.bounced).trim().toUpperCase();
        const bounceDisplay =
          b === "TRUE"  ? "<span style='color:#c0392b;font-weight:600'>Yes</span>" :
          b === "FALSE" ? "<span style='color:#2e7d32'>No</span>" :
                          "<span style='color:#f0a500;font-style:italic'>Not checked</span>";
        return "<tr>"
          + "<td title='" + r.email + "'>" + r.email + "</td>"
          + "<td><span class='badge " + type + "'>" + labels[type] + "</span></td>"
          + "<td>" + (r.opens || 0) + "</td>"
          + "<td style='color:#888'>" + (r.last || "—") + "</td>"
          + "<td>" + bounceDisplay + "</td>"
          + "</tr>";
      }).join("");
    <\/script>
    </body></html>
  `)
  .setWidth(720)
  .setHeight(580)
  .setTitle("Campaign Summary");

  SpreadsheetApp.getUi().showModalDialog(html, "Campaign Summary");
}
// ─── Debuggin if placeholders are correct ──────────────────────────────────────────────
function debugTemplatePlaceholders() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  // Get sheet headers
  const heads = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Get first data row
  const firstRow = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];

  // Ask for subject line of draft
  const subjectLine = Browser.inputBox(
    "Debug",
    "Enter the exact subject line of your Gmail draft:",
    Browser.Buttons.OK_CANCEL
  );
  if (subjectLine === "cancel" || subjectLine === "") return;

  // Get draft body
  const drafts  = GmailApp.getDrafts();
  const draft   = drafts.filter(d => d.getMessage().getSubject() === subjectLine)[0];
  if (!draft) {
    ui.alert("❌ Draft not found. Check the subject line.", "", ui.ButtonSet.OK);
    return;
  }

  const htmlBody = draft.getMessage().getBody();

  // Extract all {{placeholders}} from draft
  const placeholderMatches = [...htmlBody.matchAll(/{{([^{}]+)}}/g)];
  const placeholders       = placeholderMatches.map(m => m[1]);

  if (placeholders.length === 0) {
    ui.alert(
      "⚠️ No Placeholders Found",
      "No {{...}} placeholders found in your draft.\n\n" +
      "Gmail may have converted {{ to curly/smart quotes.\n" +
      "Retype the {{ and }} manually in your draft and try again.",
      ui.ButtonSet.OK
    );
    return;
  }

  // Compare placeholders vs sheet headers
  let report      = "=== PLACEHOLDER MATCH REPORT ===\n\n";
  let hasProblems = false;

  placeholders.forEach(p => {
    // Exact match
    if (heads.includes(p)) {
      const colIdx = heads.indexOf(p);
      const value  = firstRow[colIdx];
      report += `✅ {{${p}}} → matched → value: "${value}"\n`;
    } else {
      hasProblems = true;
      // Try to find a close match
      const closeMatch = heads.find(
        h => h.trim().toLowerCase() === p.trim().toLowerCase()
      );
      if (closeMatch) {
        report += `⚠️ {{${p}}} → NO exact match, but close match found: "${closeMatch}"\n`;
        report += `   Fix: Rename sheet column to exactly "${p}" OR change draft to {{${closeMatch}}}\n\n`;
      } else {
        report += `❌ {{${p}}} → NOT FOUND in sheet headers\n`;
        report += `   Sheet headers are: ${heads.map(h => `"${h}"`).join(", ")}\n\n`;
      }
    }
  });

  report += "\n=== SHEET HEADERS (exact) ===\n";
  heads.forEach(h => {
    report += `"${h}" (${h.length} chars)\n`;
  });

  Logger.log(report);
  ui.alert(
    hasProblems ? "⚠️ Issues Found" : "✅ All Matched",
    report,
    ui.ButtonSet.OK
  );
}
// ─── Quota helpers ────────────────────────────────────────────────────────────
function getRemainingQuota_() {
  const remaining = MailApp.getRemainingDailyQuota();
  const limit     = remaining > 100 ? 1500 : 100;
  const sent      = limit - remaining;
  return { sent, remaining, limit };
}

function showQuotaStatus() {
  const ui = SpreadsheetApp.getUi();
  const q  = getRemainingQuota_();
  ui.alert(
    "📊 Today's Email Quota",
    `Emails sent today : ${q.sent}\n` +
    `Remaining today   : ${q.remaining}\n` +
    `Daily limit       : ${q.limit}\n\n` +
    (q.remaining === 0
      ? "❌ Limit reached. Quota resets at midnight Pacific Time."
      : `✅ You can send ${q.remaining} more email(s) today.`),
    ui.ButtonSet.OK
  );
}


// ─── Bounce checker ───────────────────────────────────────────────────────────
function checkBounces() {
  const ui    = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();
  const data  = sheet.getDataRange().getValues();
  const heads = data.shift();

  const recipColIdx  = heads.indexOf(RECIPIENT_COL);
  const bounceColIdx = heads.indexOf(BOUNCE_COL);
  const statusColIdx = heads.indexOf(EMAIL_SENT_COL);

  if (bounceColIdx === -1) {
    ui.alert("⚠️ Column Not Found",
      `Add a column called "${BOUNCE_COL}" to your sheet first.`,
      ui.ButtonSet.OK);
    return;
  }

  const bouncedAddresses = new Set();

  const searchQueries = [
    'subject:"Delivery Status Notification (Failure)"',
    'subject:"Mail delivery failed"',
    'subject:"Undeliverable:"',
    'subject:"Delivery has failed"',
    'subject:"Failed Delivery"',
    'from:(mailer-daemon OR postmaster)',
  ];

  searchQueries.forEach(query => {
    try {
      const threads = GmailApp.search(query, 0, 50);
      threads.forEach(thread => {
        thread.getMessages().forEach(msg => {
          const body = msg.getPlainBody();

          const finalRecipient = body.match(
            /Final-Recipient\s*:\s*rfc822\s*;\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i
          );
          if (finalRecipient) { bouncedAddresses.add(finalRecipient[1].toLowerCase()); return; }

          const originalRecipient = body.match(
            /Original-Recipient\s*:\s*rfc822\s*;\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i
          );
          if (originalRecipient) { bouncedAddresses.add(originalRecipient[1].toLowerCase()); return; }

          const deliveryFail = body.match(
            /delivery has failed to these recipients?[^\n]*\n+\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i
          );
          if (deliveryFail) { bouncedAddresses.add(deliveryFail[1].toLowerCase()); return; }

          const failurePatterns = [
            /does not exist[^\n]*\n\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i,
            /no such user[^\n]*\n\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i,
            /user unknown[^\n]*\n\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i,
            /address not found[^\n]*\n\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i,
            /account.*?does not exist[^\n]*\n\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i,
            /invalid.*?address[^\n]*\n\s*([\w.+\-]+@[\w\-]+\.[a-z.]{2,})/i,
          ];
          for (const pattern of failurePatterns) {
            const match = body.match(pattern);
            if (match) { bouncedAddresses.add(match[1].toLowerCase()); break; }
          }
        });
      });
    } catch (e) {
      console.error("Search error for query: " + query, e);
    }
  });

  let bounceCount = 0;

  data.forEach((row, i) => {
    const email    = (row[recipColIdx]  || "").toLowerCase().trim();
    const status   = (row[statusColIdx] || "").toString();
    const wasSent  = status.startsWith("Sent");
    const cellRef  = sheet.getRange(i + 2, bounceColIdx + 1);

    if (!email || !wasSent) {
      // Not sent yet — leave bounce cell blank, don't touch it
      return;
    }

    if (bouncedAddresses.has(email)) {
      // ✅ Confirmed bounce
      cellRef.setValue("TRUE");
      bounceCount++;
    } else {
      // ✅ Sent and NOT bounced — explicitly mark as FALSE so summary knows it was checked
      cellRef.setValue("FALSE");
    }
  });

  formatBounceColumn_(sheet, heads);

  ui.alert(
    "🔍 Bounce Check Complete",
    `Checked all sent rows.\nBounced: ${bounceCount}\nDelivered: no bounce detected.`,
    ui.ButtonSet.OK
  );
}

// ─── Apply subtle color formatting to the Bounced column ─────────────────────
function formatBounceColumn_(sheet, heads) {
  const bounceColIdx = heads.indexOf(BOUNCE_COL);
  if (bounceColIdx === -1) return;

  const lastRow  = sheet.getLastRow();
  if (lastRow < 2) return;

  const col      = bounceColIdx + 1;
  const dataRows = lastRow - 1;
  const range    = sheet.getRange(2, col, dataRows, 1);
  const values   = range.getValues();

  values.forEach((row, i) => {
    const val  = String(row[0]).trim().toUpperCase();
    const cell = sheet.getRange(i + 2, col);

    if (val === "TRUE") {
      // Soft red — muted, not alarming
      cell.setBackground("#fce8e6")
          .setFontColor("#c0392b")
          .setFontWeight("bold");
    } else if (val === "FALSE") {
      // Soft green — calm confirmation
      cell.setBackground("#e8f5e9")
          .setFontColor("#2e7d32")
          .setFontWeight("normal");
    } else {
      // Blank / not checked — neutral
      cell.setBackground(null)
          .setFontColor(null)
          .setFontWeight("normal");
    }
  });
}

// ─── Gmail template helpers ───────────────────────────────────────────────────
function getGmailTemplateFromDrafts_(subject_line) {
  try {
    const drafts = GmailApp.getDrafts();
    const draft  = drafts.filter(
      el => el.getMessage().getSubject() === subject_line
    )[0];

    if (!draft) throw new Error("Draft not found");

    const msg             = draft.getMessage();
    const allInlineImages = msg.getAttachments({ includeInlineImages: true,  includeAttachments: false });
    const attachments     = msg.getAttachments({ includeInlineImages: false });
    const htmlBody        = msg.getBody();

    const img_obj = allInlineImages.reduce((obj, i) => {
      obj[i.getName()] = i;
      return obj;
    }, {});

    const imgexp          = /<img.*?src="cid:(.*?)".*?alt="(.*?)"[^\>]+>/g;
    const matches         = [...htmlBody.matchAll(imgexp)];
    const inlineImagesObj = {};
    for (const match of matches) inlineImagesObj[match[1]] = img_obj[match[2]];

    // Strip any HTML tags that Gmail may have injected inside {{ }} placeholders
    // e.g. {{<span style="...">Description</span>}} → {{Description}}
    const cleanedHtmlBody = htmlBody.replace(
      /{{(.*?)}}/gs,
      (match, inner) => {
        const stripped = inner
          .replace(/<[^>]+>/g, "")   // remove all HTML tags inside {{ }}
          .replace(/&nbsp;/g, " ")   // replace &nbsp; with space
          .replace(/\s+/g, " ")      // collapse multiple spaces
          .trim();
        return `{{${stripped}}}`;
      }
    );

    return {
      message: {
        subject: subject_line,
        text:    msg.getPlainBody(),
        html:    cleanedHtmlBody,     // ← use cleaned version
      },
      attachments:  attachments,
      inlineImages: inlineImagesObj,
    };
  } catch (e) {
    throw new Error("Oops - can't find Gmail draft");
  }
}

function fillInTemplateFromObject_(template, data) {
  let template_string = JSON.stringify(template);
  template_string = template_string.replace(/{{[^{}]+}}/g, (key) =>
    escapeData_(data[key.replace(/[{}]+/g, "")] || "")
  );
  return JSON.parse(template_string);
}

function escapeData_(str) {
  return str
    .replace(/[\\]/g, "\\\\")
    .replace(/[\"]/g, '\\"')
    .replace(/[\/]/g, "\\/")
    .replace(/[\b]/g, "\\b")
    .replace(/[\f]/g, "\\f")
    .replace(/[\n]/g, "\\n")
    .replace(/[\r]/g, "\\r")
    .replace(/[\t]/g, "\\t");
}

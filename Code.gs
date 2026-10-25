const SPREADSHEET_ID = "1hz8AsbfyxGodqdyFL9-3eOzIrdsdxcmgKMekQgF-YUE";
const SHEET_NAME = "Web app dev";
const EMAIL_SHEET_NAME = "Mail ID's";
const SEARCH_ITEM_SHEET_NAME = "Search Item";
const STOCK_SHEET_NAME = "Automatic Spare Tracker";

// =========================================================
// PORTAL AUTHENTICATION SETTINGS
// =========================================================

const USER_SHEET_NAME = "User Management";
const LOGIN_HISTORY_SHEET_NAME = "Login History";
const BREAKDOWN_SHEET_NAME = "breakdown management";
const SCHEDULED_DOWNTIME_SHEET_NAME = "Scheduled Downtime";
const RECEIVED_SPARE_SHEET_NAME = "Received Spare";
const SUPPORT_ORDER_SHEET_NAME = "Support order";
const PVC_BELT_RECEIVED_SHEET_NAME = "PVC Belt Received";
const PVC_BELT_CONSUMED_SHEET_NAME = "PVC Belt Consumed";
const PVC_BELT_TYPES = ["PVC belt", "PVC grub belt"];
const PVC_BELT_SIZES = ["1000mm", "1200mm", "1500mm"];
const PVC_BELT_USED_IN = ["CBS PC", "Nido Bagging"];
const SUPPORT_ORDER_TOOL_NO = "3004";
const RECOVERY_ADMIN_EMAIL = "ramjivishwakarma.service.1746@gmail.com";
const TEMP_ADMIN_PASSWORD = "0000";
const TEMP_ADMIN_MINUTES = 5;

const SESSION_PREFIX = "SESSION_";
const SESSION_HOURS = 8;

// =========================================================
// STATUS SETTINGS
// =========================================================

const STATUS_REQUESTING_ORDER = "Requesting Order";
const STATUS_ACCEPTED = "Accepted";
const STATUS_ORDERED = "Ordered";
const STATUS_NO_NEED = "No Need"; // legacy status kept for old records only
const STATUS_ACCEPTED_ORDERED = "Ordered"; // legacy alias
const STATUS_COMPLETED = "Completed";
const STATUS_CLOSED = "Closed";

// =========================================================
// OPEN PORTAL
// =========================================================

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("portal")
    .setTitle("Material Request Dashboard");
}

// SUBMIT MATERIAL REQUEST
function submitRequest(data) {
  const session = requirePortalPermission_(
    data && data.sessionId,
    "spareParts",
    "create"
  );

  try {
    const portalRequestType = String(
      data && data.requestType || "New Material"
    ).trim();

    if (!data || !data.email || !data.name || !data.mobile || !data.stationId ||
        (!isNewSpareRequest && !data.ticketId)) {
      return { success: false, message: "Please fill all required details." };
    }

    if (!Array.isArray(data.items) || data.items.length === 0) {
      return { success: false, message: "Please add at least one item." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      return { success: false, message: "Sheet not found: " + SHEET_NAME };
    }

    ensureWebAppDevColumns_(sheet);

    const rows = [];

    if (portalRequestType !== "New Material" && portalRequestType !== "Consumed Material" && portalRequestType !== "New Spare Request") {
      return { success: false, message: "Invalid Request Type." };
    }

    if (portalRequestType === "Consumed Material" && !String(data.ticketId || "").trim()) {
      return { success: false, message: "Ticket ID is required for Consumed Material requests." };
    }

    const requestId =
      "MR-" +
      Utilities.formatDate(
        new Date(),
        Session.getScriptTimeZone() || "Asia/Kolkata",
        "yyyyMMdd-HHmmss"
      ) +
      "-" +
      Utilities.getUuid().slice(0, 6).toUpperCase();

    data.items.forEach(function(item) {
      const itemId = String(item.itemId || "").trim();
      const itemName = String(item.itemName || item.partName || "").trim();
      const qty = Number(item.qty || 0);
      // Operator only raises the request. Status always starts at Requesting Order.
      const requestStatus = STATUS_REQUESTING_ORDER;
      let requestType = String(item.requestType || "").trim();

      // Top-level Request Type = Consumed Material controls the mandatory
      // Breakdown Details, while the actual Item Type remains Consumable.
      if (portalRequestType === "Consumed Material") {
        requestType = "Consumable";
      }

      if (!itemId || qty <= 0 || !requestStatus) {
        throw new Error("Please enter Item ID, valid Qty and Request Status for every item.");
      }

      if (portalRequestType === "Consumed Material") {
        if (!String(item.requestDetail || "").trim()) {
          throw new Error("Consumed Material details are required.");
        }
      } else if (
        requestType &&
        requestType !== "Scrap" &&
        requestType !== "Consumable" &&
        requestType !== "Returnable"
      ) {
        throw new Error("Invalid Material Type. Select Scrap, Consumable or Returnable.");
      }

      if (portalRequestType === "New Spare Request") {
        requestType = "Replacement";
      }

      // Actual Web app dev columns:
      // A Timestamp
      // B Email
      // C Mobile
      // D Name
      // E Item Name
      // F Item QTY
      // G Item ID
      // H Date
      // I Status
      // J Application Area
      // K Remark
      // L Closing Remark
      // M Request rejection remark
      // N Spare request status
      // O Tool No
      // P Ticket No
      // Q Mail Notification
      // R Request ID
      // S Additional mail IDs
      // T Mail thread
      // U Last mail ID
      // V Request Type
      // W Request Detail
      rows.push([
        new Date(),
        String(data.email).trim(),
        String(data.mobile).trim(),
        String(data.name).trim(),
        itemName,
        qty,
        itemId,
        new Date(),
        requestStatus,
        String(data.stationId).trim(),
        "",
        "",
        "",
        "",
        String(item.toolNo || "").trim(),
        String(data.ticketId).trim(),
        "",
        requestId,
        "",
        "",
        "",
        requestType,
        String(item.requestDetail || "").trim()
      ]);
    });

    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, rows.length, 23).setValues(rows);
    if (portalRequestType === "Consumed Material") {
      const breakdownResult = saveConsumedBreakdownFromMaterialRequest_(data);
      if (!breakdownResult.success) {
        return { success: false, message: breakdownResult.message || "Breakdown details could not be saved." };
      }
    }

    // Keep the existing email notification behavior.
    const emailSheet = ss.getSheetByName(EMAIL_SHEET_NAME);

    if (!emailSheet) {
      return {
        success: false,
        message: "Request saved, but 'Mail ID\'s' sheet was not found for email."
      };
    }

    const lastEmailRow = emailSheet.getLastRow();
    let toEmails = [];
    let ccEmails = [];

    if (lastEmailRow >= 2) {
      const emailData = emailSheet.getRange(2, 2, lastEmailRow - 1, 2).getValues();

      emailData.forEach(function(row) {
        const to = String(row[0] || "").trim();
        const cc = String(row[1] || "").trim();
        if (to) toEmails.push(to);
        if (cc) ccEmails.push(cc);
      });
    }

    toEmails = [...new Set(toEmails)];
    ccEmails = [...new Set(ccEmails)];

    if (toEmails.length > 0) {
      const subjectRef = String(data.ticketId || requestId).trim();
      const subject = String(data.ticketId || "").trim()
        ? "Material Request - " + portalRequestType + " - Ticket ID: " + String(data.ticketId).trim()
        : "Material Request - " + portalRequestType + " - Ref: " + subjectRef;
      let itemTable = "";

      data.items.forEach(function(item, index) {
        itemTable +=
          "<tr>" +
          "<td style='padding:8px;border:1px solid #ddd;'>" + (index + 1) + "</td>" +
          "<td style='padding:8px;border:1px solid #ddd;'>" + escapeHtml(item.itemId) + "</td>" +
          "<td style='padding:8px;border:1px solid #ddd;'>" + escapeHtml(item.itemName || item.partName || "") + "</td>" +
          "<td style='padding:8px;border:1px solid #ddd;text-align:center;'>" + Number(item.qty) + "</td>" +
          "<td style='padding:8px;border:1px solid #ddd;'>" + escapeHtml(item.requestType || "Replacement") + "</td>" +
          "<td style='padding:8px;border:1px solid #ddd;'>" + escapeHtml(item.requestStatus) + "</td>" +
          "</tr>";
      });

      const htmlBody =
        "<div style='font-family:Calibri,Arial,sans-serif;font-size:14px;'>" +
        "<h2 style='color:#1F4E79;'>Material Request</h2>" +
        "<p>Dear Team,</p>" +
        "<p>A new material request has been submitted with the following details:</p>" +
        "<table style='border-collapse:collapse;margin-bottom:15px;'>" +
        "<tr><td style='padding:6px 12px;font-weight:bold;'>Name</td><td style='padding:6px 12px;'>" + escapeHtml(data.name) + "</td></tr>" +
        "<tr><td style='padding:6px 12px;font-weight:bold;'>Email</td><td style='padding:6px 12px;'>" + escapeHtml(data.email) + "</td></tr>" +
        "<tr><td style='padding:6px 12px;font-weight:bold;'>Mobile</td><td style='padding:6px 12px;'>" + escapeHtml(data.mobile) + "</td></tr>" +
        "<tr><td style='padding:6px 12px;font-weight:bold;'>Ticket ID</td><td style='padding:6px 12px;'>" + escapeHtml(data.ticketId || "") + "</td></tr>" +
        "<tr><td style='padding:6px 12px;font-weight:bold;'>Station / Machine ID</td><td style='padding:6px 12px;'>" + escapeHtml(data.stationId) + "</td></tr>" +
        "<tr><td style='padding:6px 12px;font-weight:bold;'>Request Type</td><td style='padding:6px 12px;'>" + escapeHtml(portalRequestType) + "</td></tr>" +
        "</table>" +
        (portalRequestType === "Consumed Material" ? "<h3 style='color:#1F4E79;'>Consumed Material / Breakdown Details</h3><pre style='background:#f8f9fa;padding:10px;border:1px solid #ddd;white-space:pre-wrap;'>" + escapeHtml(data.items[0] && data.items[0].requestDetail ? data.items[0].requestDetail : "") + "</pre>" : "") +
        "<h3 style='color:#1F4E79;'>Item Details</h3>" +
        "<table style='border-collapse:collapse;width:100%;max-width:850px;'>" +
        "<tr style='background:#f2f2f2;'>" +
        "<th style='padding:8px;border:1px solid #ddd;'>S.No.</th>" +
        "<th style='padding:8px;border:1px solid #ddd;'>Item ID</th>" +
        "<th style='padding:8px;border:1px solid #ddd;'>Item Name</th>" +
        "<th style='padding:8px;border:1px solid #ddd;'>Qty</th>" +
        "<th style='padding:8px;border:1px solid #ddd;'>Type</th>" +
        "<th style='padding:8px;border:1px solid #ddd;'>Status</th>" +
        "</tr>" + itemTable +
        "</table>" +
        "<p style='margin-top:20px;'>Please review the above material request and proceed accordingly.</p>" +
        "<p>Regards,<br>Material Request System</p>" +
        "</div>";

      const mailOptions = { htmlBody: htmlBody };
      if (ccEmails.length > 0) mailOptions.cc = ccEmails.join(",");

      GmailApp.sendEmail(
        toEmails.join(","),
        subject,
        "A new material request has been submitted.",
        mailOptions
      );

      let threadId = "";
      let messageId = "";

      try {
        Utilities.sleep(1000);

        const query =
          'subject:"' +
          subject.replace(/"/g, "") +
          '" newer_than:1d';

        const threads = GmailApp.search(query, 0, 20);

        if (threads && threads.length) {
          let selectedThread = threads[0];

          threads.forEach(function(thread) {
            if (
              thread.getLastMessageDate() >
              selectedThread.getLastMessageDate()
            ) {
              selectedThread = thread;
            }
          });

          threadId = selectedThread.getId();

          const messages = selectedThread.getMessages();
          if (messages && messages.length) {
            messageId = messages[messages.length - 1].getId();
          }
        }
      } catch (threadError) {
        Logger.log("Initial mail thread capture failed: " + threadError);
      }

      if (threadId) {
        sheet.getRange(
          startRow,
          COL_MAIL_THREAD,
          rows.length,
          1
        ).setValues(
          rows.map(function() { return [threadId]; })
        );
      }

      if (messageId) {
        sheet.getRange(
          startRow,
          COL_LAST_MAIL_ID,
          rows.length,
          1
        ).setValues(
          rows.map(function() { return [messageId]; })
        );
      }
    }

    // Initial portal status was already communicated by the initial
    // material-request email. Mark it as sent so the same status can
    // never generate a duplicate status email later.
    for (let r = startRow; r < startRow + rows.length; r++) {
      const initialStatus = String(
        sheet.getRange(r, COL_STATUS).getDisplayValue() || ""
      ).trim();

      const initialThreadId = String(
        sheet.getRange(r, COL_MAIL_THREAD).getDisplayValue() || ""
      ).trim();

      if (
        initialThreadId &&
        (
          initialStatus === STATUS_REQUESTING_ORDER ||
          initialStatus === STATUS_NO_NEED
        )
      ) {
        markInitialStatusMailSent_(
          sheet,
          r,
          initialStatus
        );
      }
    }

    return {
      success: true,
      message: "Request submitted successfully! " + data.items.length + " item(s) saved."
    };

  } catch (error) {
    return {
      success: false,
      message: error && error.message ? error.message : String(error)
    };
  }
}


function saveConsumedBreakdownFromMaterialRequest_(data) {
  try {
    const item = data && data.items && data.items.length ? data.items[0] : null;
    if (!item || !String(item.requestDetail || "").trim()) return {success:false,message:"Breakdown details are required."};
    const d = JSON.parse(String(item.requestDetail || ""));
    return submitBreakdownFromMaterialRequest_legacyAdapter_(data, d);
  } catch(e) {
    return {success:false,message:e && e.message ? e.message : String(e)};
  }
}

function submitBreakdownFromMaterialRequest_legacyAdapter_(data, d) {
  d = d || {};
  const actualTpt = Number(d.actualTpt != null ? d.actualTpt : (d.tptDuringBreakdown != null ? d.tptDuringBreakdown : d.tptPerHour || 0));
  const payload = {
    email: data.email, name: data.name, mobile: data.mobile,
    ticketId: data.ticketId, breakdownType: d.breakdownType,
    issueOwner: d.issueOwner, startTime: d.startTime, endTime: d.endTime,
    actualTpt: actualTpt, caseNo: d.caseNo, spareUsed: d.spareUsed,
    spares: d.spares || [], orderSpare: d.orderSpare || "No", orderSpares: d.orderSpares || []
  };
  return submitBreakdownFromMaterialRequest(payload);
}

function submitBreakdownFromMaterialRequest(data) {
  const session = requirePortalPermission_(data && data.sessionId, "maintenance", "create");

  try {
    data = data || {};
    const email = String(data.email || "").trim();
    const name = String(data.name || "").trim();
    const mobile = String(data.mobile || "").trim();
    const ticketId = String(data.ticketId || "").trim();
    const type = String(data.breakdownType || "").trim();
    const owner = String(data.issueOwner || "").trim();
    const sv = String(data.startTime || "").trim();
    const ev = String(data.endTime || "").trim();
    const actualTpt = Number(data.actualTpt || 0);
    const caseNo = String(data.caseNo || "").trim();
    const spareUsed = String(data.spareUsed || "").trim();
    const spares = Array.isArray(data.spares) ? data.spares : [];
    let orderSpare = String(data.orderSpare || "No").trim();
    const orderSpares = Array.isArray(data.orderSpares) ? data.orderSpares : [];

    if (!email || !name || !mobile) return { success:false, message:"Email, Name and Mobile are required." };
    if (!ticketId) return { success:false, message:"Ticket ID is required." };
    if (["Partial Breakdown","Planned Downtime","Complete Breakdown"].indexOf(type) < 0) return { success:false, message:"Invalid Breakdown Type." };
    if (["FAPL","Delhivery"].indexOf(owner) < 0) return { success:false, message:"Issue Owner must be FAPL or Delhivery." };
    if (!sv || !ev) return { success:false, message:"Start Time and End Time are required." };
    const start = new Date(sv), end = new Date(ev);
    if (isNaN(start) || isNaN(end) || end <= start) return { success:false, message:"Invalid breakdown Start Time / End Time." };
    if (!isFinite(actualTpt) || actualTpt < 0) return { success:false, message:"Please enter valid TPT During Breakdown." };
    if (spareUsed !== "Yes" && spareUsed !== "No") return { success:false, message:"Please select whether spare was consumed." };

    let spareQty = 0, spareIds = [];
    if (spareUsed === "Yes") {
      if (!spares.length) return { success:false, message:"Please add at least one consumed spare." };
      spares.forEach(function(x) {
        const id = String(x.itemId || "").trim();
        const qty = Number(x.qty || 0);
        if (!id || qty <= 0) throw new Error("Please enter valid Spare Item ID and Qty for every consumed spare.");
        spareQty += qty;
        spareIds.push(id + ": " + qty);
      });
      if (orderSpare !== "Yes" && orderSpare !== "No") return { success:false, message:"Please select whether the consumed spare should be ordered." };
      if (orderSpare === "Yes") {
        if (!orderSpares.length) return { success:false, message:"Please add at least one spare to order." };
        orderSpares.forEach(function(x) {
          const id = String(x.itemId || "").trim();
          const qty = Number(x.qty || 0);
          if (!id || qty <= 0) throw new Error("Please enter valid Item ID and Qty for every spare request.");
        });
      }
    } else {
      orderSpare = "No";
    }

    const hours = (end - start) / 3600000;
    const expectedTpt = hours * 12000;
    const tptLoss = Math.max(0, expectedTpt - actualTpt);
    const percent = (hours / 24) * 100;
    const sheet = ensureBreakdownSheet_();
    sheet.getRange(sheet.getLastRow()+1,1,1,22).setValues([[
      new Date(), email, name, ticketId, type, start, end, hours, percent,
      owner, 12000, tptLoss, actualTpt, "", "", "Closed", spareUsed,
      spareQty, spareIds.join(", "), caseNo, "", mobile
    ]]);

    let orderedRequestId = "";
    if (orderSpare === "Yes") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const materialSheet = ss.getSheetByName(SHEET_NAME);
      if (!materialSheet) throw new Error("Sheet not found: " + SHEET_NAME);
      ensureWebAppDevColumns_(materialSheet);
      orderedRequestId = "MR-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Kolkata", "yyyyMMdd-HHmmss") + "-" + Utilities.getUuid().slice(0,6).toUpperCase();
      const rows = orderSpares.map(function(item) {
        return [new Date(), email, mobile, name, "", Number(item.qty), String(item.itemId).trim(), new Date(), STATUS_REQUESTING_ORDER, "", "", "", "", "", "", ticketId, "", orderedRequestId, "", "", "", "Replacement", "Breakdown consumed spare order"];
      });
      const startRow = Math.max(materialSheet.getLastRow()+1,2);
      materialSheet.getRange(startRow,1,rows.length,23).setValues(rows);
      sendNewMaterialRequestMail_(email,name,mobile,ticketId,"Breakdown Spare Order",rows.map(function(r){return {itemId:r[6],qty:r[5]};}),orderedRequestId);
    }

    sendBreakdownSubmissionMail_(data, hours, expectedTpt, actualTpt, tptLoss, spareIds, orderSpare, orderedRequestId);
    SpreadsheetApp.flush();
    return {success:true,message: orderSpare === "Yes" ? "Breakdown submitted and spare order request raised successfully." : "Breakdown submitted successfully."};
  } catch(e) {
    return {success:false,message:e && e.message ? e.message : String(e)};
  }
}

function submitConsumedSpareDetails(data) {
  const session = requirePortalPermission_(data && data.sessionId, "spareParts", "create");

  try {
    data = data || {};
    const email = String(data.email || "").trim();
    const name = String(data.name || "").trim();
    const mobile = String(data.mobile || "").trim();
    const ticketId = String(data.ticketId || "").trim();
    if (!email || !name || !mobile) return {success:false,message:"Email, Name and Mobile are required."};

    let items = Array.isArray(data.items) ? data.items : [];
    if (!items.length && data.itemId) {
      items = [{itemId:data.itemId, qty:data.qty}];
    }
    if (!items.length) return {success:false,message:"Please add at least one consumed spare."};

    const cleanItems = [];
    for (let i = 0; i < items.length; i++) {
      const itemId = String(items[i].itemId || "").trim();
      const qty = Number(items[i].qty || 0);
      if (!itemId || qty <= 0) return {success:false,message:"Please enter valid Spare Item ID and Qty for every spare row."};
      cleanItems.push({itemId:itemId, qty:qty});
    }

    const sheet = ensureBreakdownSheet_();
    const rows = cleanItems.map(function(item) {
      return [new Date(),email,name,ticketId,"", "", "", "", "", "", 0, 0, 0, "", "", "Consumed Spare Recorded", "Yes", item.qty, item.itemId + ": " + item.qty, "", "", mobile];
    });
    sheet.getRange(sheet.getLastRow()+1,1,rows.length,22).setValues(rows);
    SpreadsheetApp.flush();
    return {success:true,message:cleanItems.length + " consumed spare detail" + (cleanItems.length === 1 ? "" : "s") + " saved successfully."};
  } catch(e) {
    return {success:false,message:e && e.message ? e.message : String(e)};
  }
}

function sendBreakdownSubmissionMail_(data, hours, expectedTpt, actualTpt, tptLoss, spareIds, orderSpare, requestId) {
  try {
    const recipients = getConfiguredEmailRecipients_("problemObservation");
    if (!recipients.to.length) return;
    const subject = "Breakdown Submitted - Ticket ID: " + String(data.ticketId || "").trim();
    const html = "<div style='font-family:Calibri,Arial,sans-serif;font-size:14px;'><h2 style='color:#1F4E79;'>Breakdown Submission</h2>" +
      "<p>A new breakdown entry has been submitted.</p><table style='border-collapse:collapse;'>" +
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Name</td><td style='padding:6px 12px;'>"+escapeHtml(data.name)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Ticket ID</td><td style='padding:6px 12px;'>"+escapeHtml(data.ticketId)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Breakdown Type</td><td style='padding:6px 12px;'>"+escapeHtml(data.breakdownType)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Issue Owner</td><td style='padding:6px 12px;'>"+escapeHtml(data.issueOwner)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Start Time</td><td style='padding:6px 12px;'>"+escapeHtml(data.startTime)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>End Time</td><td style='padding:6px 12px;'>"+escapeHtml(data.endTime)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>TPT During Breakdown</td><td style='padding:6px 12px;'>₹"+Number(actualTpt).toLocaleString("en-IN",{minimumFractionDigits:2})+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>TPT Loss</td><td style='padding:6px 12px;'>₹"+Number(tptLoss).toLocaleString("en-IN",{minimumFractionDigits:2})+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Spare Consumed</td><td style='padding:6px 12px;'>"+escapeHtml(data.spareUsed)+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Consumed Spare</td><td style='padding:6px 12px;'>"+escapeHtml(spareIds.join(", "))+"</td></tr>"+
      "<tr><td style='padding:6px 12px;font-weight:bold;'>Order Spare</td><td style='padding:6px 12px;'>"+escapeHtml(orderSpare)+"</td></tr>"+
      "</table><p>Regards,<br>Material Request System</p></div>";
    const options={htmlBody:html}; if(recipients.cc.length) options.cc=recipients.cc.join(",");
    GmailApp.sendEmail(recipients.to.join(","),subject,"Breakdown submitted for Ticket ID: "+data.ticketId,options);
  } catch(e) { Logger.log("Breakdown submission mail failed: "+e); }
}

function sendNewMaterialRequestMail_(email,name,mobile,ticketId,requestType,items,requestId) {
  try {
    const recipients = getConfiguredEmailRecipients_("newMaterialRequest");
    if (!recipients.to.length) return;
    let rows=""; (items||[]).forEach(function(item,i){rows += "<tr><td style='padding:8px;border:1px solid #ddd;'>"+(i+1)+"</td><td style='padding:8px;border:1px solid #ddd;'>"+escapeHtml(item.itemId)+"</td><td style='padding:8px;border:1px solid #ddd;'>"+Number(item.qty)+"</td></tr>";});
    const subject="Material Request - "+requestType+" - Ref: "+String(ticketId||requestId);
    const html="<div style='font-family:Calibri,Arial,sans-serif;font-size:14px;'><h2 style='color:#1F4E79;'>Material Request</h2><p>A new spare request has been submitted.</p><p><b>Name:</b> "+escapeHtml(name)+"<br><b>Email:</b> "+escapeHtml(email)+"<br><b>Mobile:</b> "+escapeHtml(mobile)+"<br><b>Reference:</b> "+escapeHtml(ticketId||requestId)+"</p><table style='border-collapse:collapse;'><tr><th style='padding:8px;border:1px solid #ddd;'>S.No.</th><th style='padding:8px;border:1px solid #ddd;'>Item ID</th><th style='padding:8px;border:1px solid #ddd;'>Qty</th></tr>"+rows+"</table></div>";
    const options={htmlBody:html}; if(recipients.cc.length) options.cc=recipients.cc.join(",");
    GmailApp.sendEmail(recipients.to.join(","),subject,"New spare request submitted.",options);
  } catch(e) { Logger.log("New spare request mail failed: "+e); }
}


// =========================================================
// STATUS UPDATE EMAIL / REPLY IN SAME GMAIL THREAD
// =========================================================

function setupStatusUpdateTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    throw new Error("Web app dev sheet not found.");
  }

  ensureWebAppDevColumns_(sheet);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "onWebAppDevEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onWebAppDevEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  return {
    success: true,
    message: "Status update email trigger created successfully."
  };
}

function onWebAppDevEdit(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    ensureWebAppDevColumns_(sheet);

    const row = e.range.getRow();
    const col = e.range.getColumn();

    // Web app dev: Status = Column I.
    if (row <= 1 || col !== COL_STATUS) return;

    sendStatusUpdateReply_(sheet, row, e.oldValue || "");

  } catch (error) {
    Logger.log("Status update trigger failed: " + error);
  }
}

function sendStatusUpdateReply_(sheet, row, oldStatusFromEvent) {
  if (!isSpareRequestRow_(sheet, row)) {
    return;
  }

  const status = String(
    sheet.getRange(row, COL_STATUS).getDisplayValue() || ""
  ).trim();

  const canonicalStatus = canonicalRequestStatus_(status);
  if (!canonicalStatus) return;

  const statusMailColumn = getStatusMailColumn_(canonicalStatus);
  if (!statusMailColumn) return;

  // One-time status rule: if this status has already generated a mail,
  // it can never generate another mail for this request.
  if (sheet.getRange(row, statusMailColumn).getValue()) {
    return;
  }

  const previousStatus = String(oldStatusFromEvent || "").trim();

  if (!isAllowedStatusTransition_(previousStatus, canonicalStatus)) {
    Logger.log(
      "Invalid status transition row " +
      row +
      ": " +
      previousStatus +
      " -> " +
      canonicalStatus
    );

    if (previousStatus) {
      sheet.getRange(row, COL_STATUS).setValue(previousStatus);
    }

    return;
  }

  const person = String(
    sheet.getRange(row, COL_NAME).getDisplayValue() || ""
  ).trim();

  const item = String(
    sheet.getRange(row, COL_ITEM_NAME).getDisplayValue() || ""
  ).trim();

  const qty = sheet.getRange(row, COL_QTY).getDisplayValue();

  const itemId = String(
    sheet.getRange(row, COL_ITEM_ID).getDisplayValue() || ""
  ).trim();

  const usedDate = sheet.getRange(row, COL_DATE).getDisplayValue();

  const area = String(
    sheet.getRange(row, COL_AREA).getDisplayValue() || ""
  ).trim();

  const closingRemark = String(
    sheet.getRange(row, COL_CLOSING_REMARK).getDisplayValue() || ""
  ).trim();

  const rejectionReason = String(
    sheet.getRange(row, COL_REJECTION_REMARK).getDisplayValue() || ""
  ).trim();

  if (canonicalStatus === "Rejected" && !rejectionReason) {
    Logger.log("Rejected status requires a rejection reason for row " + row);
    if (previousStatus) {
      sheet.getRange(row, COL_STATUS).setValue(previousStatus);
    }
    return;
  }

  const ticket = String(
    sheet.getRange(row, COL_TICKET_ID).getDisplayValue() || ""
  ).trim();

  const requestId = String(
    sheet.getRange(row, COL_REQUEST_ID).getDisplayValue() || ""
  ).trim();

  let threadId = String(
    sheet.getRange(row, COL_MAIL_THREAD).getDisplayValue() || ""
  ).trim();

  const referenceNo = ticket || requestId;
  if (!referenceNo) {
    Logger.log("Ticket/Request reference missing for row " + row);
    return;
  }

  // Recover the original Gmail thread for older rows where T is blank.
  if (!threadId) {
    try {
      const subject = "Material Request - Reference: " + referenceNo;

      const threads = GmailApp.search(
        'from:me subject:"' +
        subject.replace(/"/g, "") +
        '" newer_than:365d',
        0,
        20
      );

      if (threads && threads.length) {
        let selectedThread = threads[0];

        threads.forEach(function(thread) {
          if (
            thread.getLastMessageDate() >
            selectedThread.getLastMessageDate()
          ) {
            selectedThread = thread;
          }
        });

        threadId = selectedThread.getId();
        sheet.getRange(row, COL_MAIL_THREAD).setValue(threadId);

        const messages = selectedThread.getMessages();
        if (messages && messages.length) {
          sheet
            .getRange(row, COL_LAST_MAIL_ID)
            .setValue(messages[messages.length - 1].getId());
        }
      }
    } catch (error) {
      Logger.log(
        "Existing Gmail thread lookup failed: " +
        (error && error.stack ? error.stack : error)
      );
    }
  }

  // Status notifications are sent to the configured Settings recipients.
  // The old Gmail thread is retained only for historical traceability.

  let message = "";
  let statusColor = "#1F4E79";

  if (canonicalStatus === STATUS_ACCEPTED) {
    message =
      "<p>Your material request against <b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> has been accepted.</p>" +
      "<p>The responsible team has accepted the request and will proceed with the next step.</p>";
    statusColor = "#2563EB";

  // Status update for Ordered.
  } else if (canonicalStatus === STATUS_ORDERED) {
    message =
      "<p>Your spare part request against " +
      "<b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> has been approved.</p>" +

      "<p>The requested spare part has been ordered " +
      "and will be dispatched by the dispatch team " +
      "as per availability.</p>";

    statusColor = "#F39C12";

  // EXACT wording from the supplied status-mail script for Rejected.
  } else if (canonicalStatus === "Rejected") {
    message =
      "<p>We regret to inform you that your spare " +
      "part request against <b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> has been rejected.</p>" +

      "<p><b>Reason:</b> " +
      escapeHtml(rejectionReason) +
      "</p>";

    statusColor = "#E74C3C";

  // EXACT wording from the supplied status-mail script for Closed.
  } else if (canonicalStatus === STATUS_CLOSED) {
    message =
      "<p>Your requested spare part against " +
      "<b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> has been successfully delivered " +
      "to the site.</p>" +

      "<p>Your spare part request has been successfully delivered to the site team..</p>";

    statusColor = "#28A745";

  } else if (canonicalStatus === STATUS_COMPLETED) {
    message =
      "<p>Your spare part request against " +
      "<b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> has been completed.</p>";

    statusColor = "#28A745";

  } else if (canonicalStatus === STATUS_NO_NEED) {
    message =
      "<p>Your spare part request against " +
      "<b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> has been marked as <b>No Need</b>.</p>";

    statusColor = "#6B7280";

  } else if (canonicalStatus === STATUS_REQUESTING_ORDER) {
    message =
      "<p>Your spare part request against " +
      "<b>Ticket No. " +
      escapeHtml(referenceNo) +
      "</b> is currently under processing.</p>";
  }

  let html =
    "<div style='font-family:Calibri,Arial,sans-serif;font-size:14px'>" +

    "<h2 style='color:#1F4E79'>" +
    "Spare Part Request Status Update" +
    "</h2>" +

    "<p>Dear <b>" +
    escapeHtml(person) +
    "</b>,</p>" +

    message +

    "<br>" +

    "<table border='1' cellpadding='8' cellspacing='0' " +
    "style='border-collapse:collapse;width:500px'>" +

    "<tr>" +
    "<td><b>Current Status</b></td>" +
    "<td style='color:" +
    statusColor +
    "'><b>" +
    escapeHtml(canonicalStatus) +
    "</b></td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Request ID</b></td>" +
    "<td>" +
    escapeHtml(requestId) +
    "</td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Reference</b></td>" +
    "<td>" +
    escapeHtml(referenceNo) +
    "</td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Item Name</b></td>" +
    "<td>" +
    escapeHtml(item) +
    "</td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Item Quantity</b></td>" +
    "<td>" +
    escapeHtml(qty) +
    "</td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Item ID</b></td>" +
    "<td>" +
    escapeHtml(itemId) +
    "</td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Date of Spare Usage</b></td>" +
    "<td>" +
    escapeHtml(usedDate) +
    "</td>" +
    "</tr>" +

    "<tr>" +
    "<td><b>Used In (Machine/Station)</b></td>" +
    "<td>" +
    escapeHtml(area) +
    "</td>" +
    "</tr>";

  if (canonicalStatus === "Rejected") {
    html +=
      "<tr>" +
      "<td><b>Rejection Reason</b></td>" +
      "<td>" +
      escapeHtml(rejectionReason) +
      "</td>" +
      "</tr>";
  }

  html +=
    "</table>" +

    "<br>" +

    "<p>For any questions, please contact with the respected team.</p>" +

    "<br>" +

    "<p>Regards,<br>" +
    "<b>Service & Support</b><br>" +
    "Falcon Autotech Pvt. Ltd.</p>" +

    "<hr>" +

    "<p style='font-size:11px;color:#777'>" +
    "This is an auto-generated email. Please do not reply." +
    "</p>" +

    "</div>";

  try {
    const recipients = getConfiguredEmailRecipients_("materialRequestStatus");
    if (!recipients.to.length) {
      Logger.log("No To recipients configured in Mail ID's sheet.");
      return;
    }

    const subject = "Material Request Status - " + canonicalStatus + " - Reference: " + referenceNo;
    const options = { htmlBody: html };
    if (recipients.cc.length) options.cc = recipients.cc.join(",");

    GmailApp.sendEmail(
      recipients.to.join(","),
      subject,
      "Material Request status updated to " + canonicalStatus + ". Reference: " + referenceNo,
      options
    );

    Utilities.sleep(500);

    try {
      const sentThreads = GmailApp.search('from:me subject:"' + subject.replace(/"/g, "") + '" newer_than:1d', 0, 10);
      if (sentThreads && sentThreads.length) {
        const messages = sentThreads[0].getMessages();
        const latest = messages[messages.length - 1];
        if (latest) sheet.getRange(row, COL_LAST_MAIL_ID).setValue(latest.getId());
      }
    } catch (mailLookupError) {
      Logger.log("Status mail lookup failed: " + mailLookupError);
    }

    // Q = latest mailed status (kept for compatibility).
    sheet
      .getRange(row, COL_MAIL_NOTIFICATION)
      .setValue(canonicalStatus);

    // X:AC = permanent one-time markers.
    sheet
      .getRange(row, statusMailColumn)
      .setValue(new Date());

    Logger.log(
      "Status mail sent successfully. Row=" +
      row +
      ", Status=" +
      canonicalStatus
    );

  } catch (error) {
    Logger.log(
      "Status mail failed: " +
      (error && error.stack ? error.stack : error)
    );
  }
}

// =========================================================
// HTML ESCAPE
// =========================================================

function escapeHtml(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

// =========================================================
// ENSURE REMARK COLUMN
// Sheet1 J = Remark
// =========================================================

function ensureRemarkColumn_(sheet) {
  if (!sheet) {
    return;
  }

  const header = String(
    sheet.getRange(1, 10).getValue() || ""
  ).trim();

  if (!header) {
    sheet.getRange(1, 10).setValue("Remark");
  }
}

// =========================================================
// WEB APP DEV COLUMN MAP
// =========================================================
// A Timestamp
// B Email Address
// C Mo. No.
// D Name
// E Item Name
// F Item QTY.
// G Item ID
// H Date
// I Status
// J Application Area
// K Remark
// L Closing Remark
// M Request rejection remark
// N Spare request status
// O Tool No.
// P Ticket No.
// Q Mail Notification
// R Request ID
// S additional mail ID's
// T Mail thread
// U last mail ID
// V Request Type (portal)
// W Request Detail (portal)
// =========================================================

const COL_TIMESTAMP = 1;
const COL_EMAIL = 2;
const COL_MOBILE = 3;
const COL_NAME = 4;
const COL_ITEM_NAME = 5;
const COL_QTY = 6;
const COL_ITEM_ID = 7;
const COL_DATE = 8;
const COL_STATUS = 9;
const COL_AREA = 10;
const COL_REMARK = 11;
const COL_CLOSING_REMARK = 12;
const COL_REJECTION_REMARK = 13;
const COL_SPARE_STATUS = 14;
const COL_TOOL_NO = 15;
const COL_TICKET_ID = 16;
const COL_MAIL_NOTIFICATION = 17;
const COL_REQUEST_ID = 18;
const COL_ADDITIONAL_MAIL = 19;
const COL_MAIL_THREAD = 20;
const COL_LAST_MAIL_ID = 21;
const COL_REQUEST_TYPE = 22;
const COL_REQUEST_DETAIL = 23;

// Automatically added after existing Web app dev columns.
const COL_STATUS_REQUESTING_ORDER_MAIL = 24; // X
const COL_STATUS_NO_NEED_MAIL = 25;          // Y (legacy)
const COL_STATUS_REJECTED_MAIL = 26;         // Z (legacy)
const COL_STATUS_ORDERED_MAIL = 27;          // AA
const COL_STATUS_COMPLETED_MAIL = 28;        // AB
const COL_STATUS_CLOSED_MAIL = 29;           // AC
const COL_STATUS_ACCEPTED_MAIL = 30;         // AD

function ensureWebAppDevColumns_(sheet) {
  if (!sheet) return;

  const headers = [
    "Timestamp",
    "Email Address",
    "Mo. No.",
    "Name",
    "Item Name",
    "Item QTY.",
    "Item ID",
    "Date",
    "Status",
    "Application Area (Machine/Station Name)",
    "Remark",
    "closing remark",
    "Request rejection remark",
    "Spare request status",
    "Tool No.",
    "Ticket No.",
    "Mail Notification",
    "Request ID",
    "additional mail ID's",
    "Mail thread",
    "last mail ID",
    "Request Type",
    "Request Detail",
    "Requesting Order Mail Sent",
    "No Need Mail Sent",
    "Rejected Mail Sent",
    "Ordered Mail Sent",
    "Completed Mail Sent",
    "Closed Mail Sent",
    "Accepted Mail Sent"
  ];

  const currentLastColumn = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, currentLastColumn).getValues()[0];

  for (let i = 0; i < headers.length; i++) {
    if (!String(current[i] || "").trim()) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }
}

function normalizeRequestType_(value) {
  return String(value || "").trim().toLowerCase();
}

function isOpenRequestStatus_(status) {
  const s = normalizeStatus_(status);
  return s !== "closed" && s !== "no need";
}

function getRequestQty_(value) {
  if (value instanceof Date) return 0;

  const n = Number(value);
  if (!isFinite(n) || n <= 0) return 0;

  // Protect report from accidental date/timestamp values.
  if (n > 1000000000) return 0;

  return n;
}

function getSheetDateString_(value) {
  const timezone = Session.getScriptTimeZone() || "Asia/Kolkata";
  if (value instanceof Date) {
    return Utilities.formatDate(value, timezone, "dd-MM-yyyy HH:mm");
  }
  return String(value || "");
}


// =========================================================
// NORMALIZE STATUS
// =========================================================

function normalizeStatus_(status) {
  return String(status || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

// =========================================================
// CHECK ACCEPTED / ORDERED STATUS
// =========================================================

function isAcceptedOrderedStatus_(status) {
  const normalized = normalizeStatus_(status)
    .replace(/\s*\/\s*/g, "/");

  return (
    normalized === "accepted/ordered" ||
    normalized === "accepted" ||
    normalized === "ordered"
  );
}

// =========================================================
// CHECK VALID STATUS
// =========================================================

function canonicalRequestStatus_(status) {
  const normalized = normalizeStatus_(status).replace(/\s*\/\s*/g, "/");

  if (normalized === "requesting order") return STATUS_REQUESTING_ORDER;
  if (normalized === "accepted") return STATUS_ACCEPTED;
  if (normalized === "ordered") return STATUS_ORDERED;
  if (normalized === "accepted/ordered") return STATUS_ORDERED; // legacy rows
  if (normalized === "no need") return STATUS_NO_NEED;
  if (normalized === "rejected") return "Rejected";
  if (normalized === "completed") return STATUS_COMPLETED;
  if (normalized === "closed") return STATUS_CLOSED;

  return "";
}

function isValidRequestStatus_(status) {
  return !!canonicalRequestStatus_(status);
}

function isSpareRequestRow_(sheet, row) {
  const requestType = normalizeRequestType_(
    sheet.getRange(row, COL_REQUEST_TYPE).getDisplayValue()
  );

  const requestId = String(
    sheet.getRange(row, COL_REQUEST_ID).getDisplayValue() || ""
  ).trim();
  return !!requestId;
}

function getStatusMailColumn_(status) {
  const canonical = canonicalRequestStatus_(status);

  if (canonical === STATUS_REQUESTING_ORDER) return COL_STATUS_REQUESTING_ORDER_MAIL;
  if (canonical === STATUS_NO_NEED) return COL_STATUS_NO_NEED_MAIL;
  if (canonical === "Rejected") return COL_STATUS_REJECTED_MAIL;
  if (canonical === STATUS_ACCEPTED) return COL_STATUS_ACCEPTED_MAIL;
  if (canonical === STATUS_ORDERED) return COL_STATUS_ORDERED_MAIL;
  if (canonical === STATUS_COMPLETED) return COL_STATUS_COMPLETED_MAIL;
  if (canonical === STATUS_CLOSED) return COL_STATUS_CLOSED_MAIL;

  return 0;
}

/**
 * Status flow:
 * Requesting Order
 *       ↙       ↘
 * Accepted     Rejected
 *    ↓
 * Ordered
 *    ↓
 * Completed
 *    ↓
 * Closed
 *
 * Rejected and Closed are terminal.
 */
function isAllowedStatusTransition_(oldStatus, newStatus) {
  const oldCanonical = canonicalRequestStatus_(oldStatus);
  const newCanonical = canonicalRequestStatus_(newStatus);

  if (!newCanonical) return false;

  // New requests always start at Requesting Order.
  if (!oldCanonical) return newCanonical === STATUS_REQUESTING_ORDER;

  if (oldCanonical === newCanonical) return false;

  if (oldCanonical === STATUS_REQUESTING_ORDER) {
    // First decision must be either Accepted or Rejected.
    return newCanonical === STATUS_ACCEPTED || newCanonical === "Rejected";
  }

  if (oldCanonical === STATUS_ACCEPTED) {
    return newCanonical === STATUS_ORDERED;
  }

  if (oldCanonical === STATUS_ORDERED) {
    return newCanonical === STATUS_COMPLETED;
  }

  if (oldCanonical === STATUS_COMPLETED) {
    return newCanonical === STATUS_CLOSED;
  }

  return false;
}
function markInitialStatusMailSent_(sheet, row, status) {
  const col = getStatusMailColumn_(status);
  if (!col) return;

  if (!sheet.getRange(row, col).getValue()) {
    sheet.getRange(row, col).setValue(new Date());
  }

  sheet.getRange(row, COL_MAIL_NOTIFICATION).setValue(
    canonicalRequestStatus_(status)
  );
}



// =========================================================
// DASHBOARD
// =========================================================

function getProblemObservationDashboardCounts_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROBLEM_OBSERVATION_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return {open:0,inProgress:0,closed:0,totalRaised:0,totalClosed:0};
  const values=sheet.getRange(2,12,sheet.getLastRow()-1,1).getValues(); let open=0,inProgress=0,closed=0;
  values.forEach(function(r){const s=String(r[0]||"Open").trim().toLowerCase();if(s==="open")open++;else if(s==="in progress")inProgress++;else if(s==="closed")closed++;});
  return {open:open,inProgress:inProgress,closed:closed,totalRaised:values.length,totalClosed:closed};
}
function getDashboardObservationRecent_(fromDate,toDate){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PROBLEM_OBSERVATION_SHEET_NAME);if(!sheet||sheet.getLastRow()<2)return[];
  const tz=Session.getScriptTimeZone()||"Asia/Kolkata";const start=new Date(fromDate.getFullYear(),fromDate.getMonth(),fromDate.getDate());const end=new Date(toDate.getFullYear(),toDate.getMonth(),toDate.getDate()+1);const out=[];
  sheet.getRange(2,1,sheet.getLastRow()-1,19).getValues().forEach(function(r){const d=r[0] instanceof Date?r[0]:null;if(!d||d<start||d>=end||!r[1])return;out.push({timestamp:Utilities.formatDate(d,tz,"dd-MM-yyyy HH:mm"),observationId:String(r[1]||""),email:String(r[2]||""),name:String(r[3]||""),stationId:String(r[5]||""),ticketId:String(r[6]||""),category:String(r[7]||""),title:String(r[8]||""),priority:String(r[10]||""),status:String(r[11]||"Open"),mediaUrl:String(r[14]||"")});});
  out.reverse();return out.slice(0,10);
}


function getSupportOrderHeaderRow_(values) {
  if (!values || !values.length) return -1;
  const maxRows = Math.min(values.length, 10);
  const normalize = function(value) {
    return String(value == null ? "" : value)
      .replace(/<[^>]*>/g, " ")
      .replace(/[\u00A0]/g, " ")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  };
  const required = ["item id", "tool no.", "order date"];
  for (let r = 0; r < maxRows; r++) {
    const row = values[r] || [];
    const headers = row.map(normalize);
    if (required.every(function(name) { return headers.indexOf(name) >= 0; })) return r;
  }
  return -1;
}

function getSupportOrderDashboardData_(toolNo, filterStart, filterEnd) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUPPORT_ORDER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const requestedTool = String(toolNo || SUPPORT_ORDER_TOOL_NO).trim().toLowerCase();
  const values = sheet.getDataRange().getDisplayValues();
  if (!values.length) return [];

  const headerRow = getSupportOrderHeaderRow_(values);
  if (headerRow < 0) return [];

  const headers = values[headerRow].map(function(h) {
    return String(h || "").replace(/<[^>]*>/g, " ").replace(/[\u00A0]/g, " ").trim().toLowerCase().replace(/\s+/g, " ");
  });

  function findHeader_(name) {
    const target = String(name).trim().toLowerCase().replace(/\s+/g, " ");
    return headers.indexOf(target);
  }

  const idx = {
    orderNo: findHeader_("Order Number"),
    deliveryLocation: findHeader_("Delivery Location"),
    status: findHeader_("Status"),
    orderQty: findHeader_("Order Qty"),
    pickedQty: findHeader_("Picked Qty"),
    itemId: findHeader_("Item ID"),
    toolNo: findHeader_("Tool No."),
    station: findHeader_("Station"),
    positionNo: findHeader_("Position No"),
    reworkNo: findHeader_("Rework No"),
    make: findHeader_("Make"),
    model: findHeader_("Model"),
    wt: findHeader_("Wt."),
    orderType: findHeader_("Order Type"),
    partName: findHeader_("Part Name"),
    placedBy: findHeader_("Placed By"),
    orderDate: findHeader_("Order Date"),
    picker: findHeader_("Picker"),
    challanType: findHeader_("Challan Type"),
    challanNo: findHeader_("Challan No."),
    dispatchDate: findHeader_("Dispatch Date"),
    courier: findHeader_("Courier Service Name"),
    docketNo: findHeader_("Docket No."),
    courierStatus: findHeader_("Courier Status"),
    deliveryDate: findHeader_("Delivery Date"),
    machineNo: findHeader_("Machine No."),
    requestedBy: findHeader_("Requested By"),
    remarks: findHeader_("Remarks"),
    check: findHeader_("Check")
  };

  if (idx.toolNo < 0) return [];

  // Apply the same dashboard date/time filter to Support Order rows.
  // If no filter is supplied, the existing all-records behaviour is preserved.
  let filterStartDate = filterStart instanceof Date && !isNaN(filterStart.getTime()) ? new Date(filterStart.getTime()) : null;
  let filterEndDate = filterEnd instanceof Date && !isNaN(filterEnd.getTime()) ? new Date(filterEnd.getTime()) : null;

  function parseSupportDashboardDate_(value) {
    if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
    const raw = String(value == null ? "" : value).replace(/[\u00A0]/g, " ").trim();
    if (!raw) return null;

    let d = new Date(raw);
    if (!isNaN(d.getTime())) return d;

    let m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
      if (!isNaN(d.getTime())) return d;
    }

    m = raw.match(/^(\d{1,2})[-\/]([A-Za-z]{3,9})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      const months = {jan:0,january:0,feb:1,february:1,mar:2,march:2,apr:3,april:3,may:4,jun:5,june:5,jul:6,july:6,aug:7,august:7,sep:8,sept:8,september:8,oct:9,october:9,nov:10,november:10,dec:11,december:11};
      const month = months[String(m[2]).toLowerCase()];
      if (month != null) {
        d = new Date(Number(m[3]), month, Number(m[1]), Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0));
        if (!isNaN(d.getTime())) return d;
      }
    }
    return null;
  }

  const out = [];
  for (let i = headerRow + 1; i < values.length; i++) {
    const row = values[i];
    const rowTool = idx.toolNo >= 0 ? String(row[idx.toolNo] || "").trim().toLowerCase() : "";
    const rowMachine = idx.machineNo >= 0 ? String(row[idx.machineNo] || "").trim().toLowerCase() : "";
    // Tracking table includes records where either Tool No. OR Machine No. is 3004.
    if (rowTool !== requestedTool && rowMachine !== requestedTool) continue;

    if (filterStartDate && filterEndDate && idx.orderDate >= 0) {
      const orderDate = parseSupportDashboardDate_(row[idx.orderDate]);
      if (!orderDate) continue;
      if (orderDate < filterStartDate || orderDate > filterEndDate) continue;
    }

    out.push({
      rowNumber: i + 1,
      orderNumber: idx.orderNo >= 0 ? row[idx.orderNo] : "",
      deliveryLocation: idx.deliveryLocation >= 0 ? row[idx.deliveryLocation] : "",
      status: idx.status >= 0 ? row[idx.status] : "",
      orderQty: idx.orderQty >= 0 ? row[idx.orderQty] : "",
      pickedQty: idx.pickedQty >= 0 ? row[idx.pickedQty] : "",
      itemId: idx.itemId >= 0 ? row[idx.itemId] : "",
      toolNo: idx.toolNo >= 0 ? row[idx.toolNo] : "",
      station: idx.station >= 0 ? row[idx.station] : "",
      positionNo: idx.positionNo >= 0 ? row[idx.positionNo] : "",
      reworkNo: idx.reworkNo >= 0 ? row[idx.reworkNo] : "",
      make: idx.make >= 0 ? row[idx.make] : "",
      model: idx.model >= 0 ? row[idx.model] : "",
      wt: idx.wt >= 0 ? row[idx.wt] : "",
      orderType: idx.orderType >= 0 ? row[idx.orderType] : "",
      partName: idx.partName >= 0 ? row[idx.partName] : "",
      placedBy: idx.placedBy >= 0 ? row[idx.placedBy] : "",
      orderDate: idx.orderDate >= 0 ? row[idx.orderDate] : "",
      picker: idx.picker >= 0 ? row[idx.picker] : "",
      challanType: idx.challanType >= 0 ? row[idx.challanType] : "",
      challanNo: idx.challanNo >= 0 ? row[idx.challanNo] : "",
      dispatchDate: idx.dispatchDate >= 0 ? row[idx.dispatchDate] : "",
      courier: idx.courier >= 0 ? row[idx.courier] : "",
      docketNo: idx.docketNo >= 0 ? row[idx.docketNo] : "",
      courierStatus: idx.courierStatus >= 0 ? row[idx.courierStatus] : "",
      deliveryDate: idx.deliveryDate >= 0 ? row[idx.deliveryDate] : "",
      machineNo: idx.machineNo >= 0 ? row[idx.machineNo] : "",
      requestedBy: idx.requestedBy >= 0 ? row[idx.requestedBy] : "",
      remarks: idx.remarks >= 0 ? row[idx.remarks] : "",
      check: idx.check >= 0 ? row[idx.check] : ""
    });
  }

  return out.slice(-50).reverse();
}


// =========================================================
// SUPPORT ORDER MATCH FOR A MATERIAL REQUEST
// Tool No. 3004 + Item ID + Order Date >= Request Raised Date
// Returns the first matching order after the request date.
// =========================================================
function getSupportOrderMatchForRequest(sessionId, itemId, requestRaisedDate) {
  const session = validateSession_(sessionId);
  if (!session) return { success: false, message: "Session expired. Please login again." };

  itemId = String(itemId || "").trim();
  if (!itemId) return { success: false, message: "Item ID is required.", found: false };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUPPORT_ORDER_SHEET_NAME);
  if (!sheet || sheet.getLastRow() < 2) return { success: true, found: false };

  const values = sheet.getDataRange().getValues();
  if (!values.length) return { success: true, found: false };

  const headerRow = getSupportOrderHeaderRow_(values);
  if (headerRow < 0) return { success: true, found: false, message: "Support order header row not found." };

  const headers = values[headerRow].map(function(h) {
    return String(h || "").replace(/<[^>]*>/g, " ").replace(/[\u00A0]/g, " ").trim().toLowerCase().replace(/\s+/g, " ");
  });
  function findHeader_(name) {
    return headers.indexOf(String(name).trim().toLowerCase().replace(/\s+/g, " "));
  }

  const idx = {
    orderNo: findHeader_("Order Number"),
    status: findHeader_("Status"),
    itemId: findHeader_("Item ID"),
    toolNo: findHeader_("Tool No."),
    orderDate: findHeader_("Order Date"),
    dispatchDate: findHeader_("Dispatch Date"),
    courier: findHeader_("Courier Service Name"),
    docketNo: findHeader_("Docket No."),
    courierStatus: findHeader_("Courier Status"),
    deliveryDate: findHeader_("Delivery Date"),
    remarks: findHeader_("Remarks"),
    deliveryLocation: findHeader_("Delivery Location"),
    partName: findHeader_("Part Name")
  };

  if (idx.itemId < 0 || idx.toolNo < 0 || idx.orderDate < 0) {
    return { success: true, found: false, message: "Support order columns not found." };
  }

  const requestedTool = normalizeToolNo_(SUPPORT_ORDER_TOOL_NO);
  const requestedItem = String(itemId).trim().toLowerCase();
  let requestDate = requestRaisedDate instanceof Date ? new Date(requestRaisedDate.getTime()) : new Date(String(requestRaisedDate || ""));
  if (isNaN(requestDate.getTime())) requestDate = null;
  if (requestDate) requestDate.setHours(0, 0, 0, 0);

  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const matches = [];

  for (let i = headerRow + 1; i < values.length; i++) {
    const row = values[i];
    const rowTool = normalizeToolNo_(row[idx.toolNo]);
    const rowItem = String(row[idx.itemId] == null ? "" : row[idx.itemId]).trim().toLowerCase();
    if (rowTool !== requestedTool || rowItem !== requestedItem) continue;

    let orderDate = row[idx.orderDate] instanceof Date ? new Date(row[idx.orderDate].getTime()) : new Date(String(row[idx.orderDate] || ""));
    if (isNaN(orderDate.getTime())) continue;
    if (requestDate) {
      const compareOrder = new Date(orderDate.getFullYear(), orderDate.getMonth(), orderDate.getDate());
      if (compareOrder < requestDate) continue;
    }

    matches.push({
      rowNumber: i + 1,
      orderDateObj: orderDate,
      orderNumber: idx.orderNo >= 0 ? String(row[idx.orderNo] || "") : "",
      status: idx.status >= 0 ? String(row[idx.status] || "") : "",
      itemId: idx.itemId >= 0 ? String(row[idx.itemId] || "") : "",
      toolNo: idx.toolNo >= 0 ? String(row[idx.toolNo] || "") : "",
      orderDate: Utilities.formatDate(orderDate, tz, "dd-MM-yyyy HH:mm"),
      dispatchDate: idx.dispatchDate >= 0 ? String(row[idx.dispatchDate] || "") : "",
      courier: idx.courier >= 0 ? String(row[idx.courier] || "") : "",
      docketNo: idx.docketNo >= 0 ? String(row[idx.docketNo] || "") : "",
      courierStatus: idx.courierStatus >= 0 ? String(row[idx.courierStatus] || "") : "",
      deliveryDate: idx.deliveryDate >= 0 ? String(row[idx.deliveryDate] || "") : "",
      remarks: idx.remarks >= 0 ? String(row[idx.remarks] || "") : "",
      deliveryLocation: idx.deliveryLocation >= 0 ? String(row[idx.deliveryLocation] || "") : "",
      partName: idx.partName >= 0 ? String(row[idx.partName] || "") : ""
    });
  }

  if (!matches.length) return { success: true, found: false };

  matches.sort(function(a, b) { return a.orderDateObj.getTime() - b.orderDateObj.getTime(); });
  const match = matches[0];
  delete match.orderDateObj;
  return { success: true, found: true, match: match };
}


function getSupportOrderMatchData_(itemId, requestRaisedDate, values, headers) {
  // Support Order matching is intentionally limited to:
  // Tool No. 3004 + same Item ID + Support Order Order Date >= Web app dev request date.
  // This function is isolated so no other dashboard/workflow logic is changed.
  itemId = String(itemId == null ? "" : itemId).trim().toLowerCase();
  if (!itemId || !values || values.length < 2) return null;

  function normalizeHeader_(value) {
    return String(value == null ? "" : value)
      .replace(/<[^>]*>/g, " ")
      .replace(/[\u00A0]/g, " ")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }

  function findHeader_(name) {
    const wanted = normalizeHeader_(name);
    for (let h = 0; h < headers.length; h++) {
      if (normalizeHeader_(headers[h]) === wanted) return h;
    }
    return -1;
  }

  // Handles both real Google Sheet Date objects and text dates such as:
  // 30-Sep-2026, 30/09/2026, 30-09-2026, 2026-09-30 and ISO timestamps.
  function parseSupportDate_(value) {
    if (value instanceof Date && !isNaN(value.getTime())) {
      return new Date(value.getTime());
    }

    const raw = String(value == null ? "" : value)
      .replace(/[\u00A0]/g, " ")
      .trim();
    if (!raw) return null;

    let d = new Date(raw);
    if (!isNaN(d.getTime())) return d;

    let m = raw.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      d = new Date(
        Number(m[3]), Number(m[2]) - 1, Number(m[1]),
        Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0), 0
      );
      if (!isNaN(d.getTime())) return d;
    }

    m = raw.match(/^(\d{1,2})[-\/]([A-Za-z]{3,9})[-\/](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      const months = {
        jan:0, january:0, feb:1, february:1, mar:2, march:2,
        apr:3, april:3, may:4, jun:5, june:5, jul:6, july:6,
        aug:7, august:7, sep:8, sept:8, september:8, oct:9,
        october:9, nov:10, november:10, dec:11, december:11
      };
      const month = months[String(m[2]).toLowerCase()];
      if (month != null) {
        d = new Date(
          Number(m[3]), month, Number(m[1]),
          Number(m[4] || 0), Number(m[5] || 0), Number(m[6] || 0), 0
        );
        if (!isNaN(d.getTime())) return d;
      }
    }

    return null;
  }

  function normalizeItemId_(value) {
    return String(value == null ? "" : value)
      .replace(/[\u00A0]/g, " ")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "");
  }

  const detectedHeaderRow = getSupportOrderHeaderRow_(values);
  if (detectedHeaderRow < 0) return null;
  headers = values[detectedHeaderRow].map(function(h) {
    return normalizeHeader_(h);
  });

  const idx = {
    orderNo: findHeader_("Order Number"),
    status: findHeader_("Status"),
    itemId: findHeader_("Item ID"),
    toolNo: findHeader_("Tool No."),
    orderDate: findHeader_("Order Date"),
    dispatchDate: findHeader_("Dispatch Date"),
    courier: findHeader_("Courier Service Name"),
    docketNo: findHeader_("Docket No."),
    courierStatus: findHeader_("Courier Status"),
    deliveryDate: findHeader_("Delivery Date"),
    deliveryLocation: findHeader_("Delivery Location"),
    partName: findHeader_("Part Name"),
    orderQty: findHeader_("Order Qty"),
    pickedQty: findHeader_("Picked Qty"),
    challanNo: findHeader_("Challan No."),
    machineNo: findHeader_("Machine No."),
    requestedBy: findHeader_("Requested By"),
    remarks: findHeader_("Remarks")
  };

  if (idx.itemId < 0 || idx.toolNo < 0 || idx.orderDate < 0) return null;

  const requestDateParsed = parseSupportDate_(requestRaisedDate);
  let requestDate = requestDateParsed ? new Date(requestDateParsed.getTime()) : null;
  if (requestDate) {
    requestDate = new Date(
      requestDate.getFullYear(), requestDate.getMonth(), requestDate.getDate(), 0, 0, 0, 0
    );
  }

  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const requestedTool = normalizeToolNo_(SUPPORT_ORDER_TOOL_NO);
  const matches = [];

  for (let i = detectedHeaderRow + 1; i < values.length; i++) {
    const row = values[i];

    if (normalizeToolNo_(row[idx.toolNo]) !== requestedTool) continue;
    if (normalizeItemId_(row[idx.itemId]) !== normalizeItemId_(itemId)) continue;

    const orderDate = parseSupportDate_(row[idx.orderDate]);
    if (!orderDate) continue;

    const compareOrder = new Date(
      orderDate.getFullYear(), orderDate.getMonth(), orderDate.getDate(), 0, 0, 0, 0
    );
    if (requestDate && compareOrder.getTime() < requestDate.getTime()) continue;

    const text = function(ix) {
      return ix >= 0 ? String(row[ix] == null ? "" : row[ix]).trim() : "";
    };

    matches.push({
      rowNumber: i + 1,
      orderDateObj: orderDate,
      orderNumber: text(idx.orderNo),
      status: text(idx.status),
      itemId: text(idx.itemId),
      toolNo: text(idx.toolNo),
      orderDate: Utilities.formatDate(orderDate, tz, "dd-MM-yyyy HH:mm"),
      orderQty: text(idx.orderQty),
      pickedQty: text(idx.pickedQty),
      dispatchDate: text(idx.dispatchDate),
      courier: text(idx.courier),
      docketNo: text(idx.docketNo),
      courierStatus: text(idx.courierStatus),
      deliveryDate: text(idx.deliveryDate),
      challanNo: text(idx.challanNo),
      deliveryLocation: text(idx.deliveryLocation),
      partName: text(idx.partName),
      machineNo: text(idx.machineNo),
      requestedBy: text(idx.requestedBy),
      remarks: text(idx.remarks),
      dispatched: !!(text(idx.dispatchDate) || text(idx.docketNo) || text(idx.courier))
    });
  }

  if (!matches.length) return null;

  // If multiple Support Orders exist for the same Item ID, use the earliest
  // order created on/after the Web app dev request date.
  matches.sort(function(a, b) {
    return a.orderDateObj.getTime() - b.orderDateObj.getTime();
  });

  const match = matches[0];
  delete match.orderDateObj;
  return match;
}

function getSupportOrderMatchesForRequests_(recent) {
  if (!recent || !recent.length) return recent || [];
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SUPPORT_ORDER_SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) return recent;
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(function(h){ return String(h || "").trim().toLowerCase().replace(/\s+/g, " "); });
  recent.forEach(function(r){
    const match = getSupportOrderMatchData_(r.itemId, r.requestDateObj || r.requestDateIso || r.timestamp, values, headers);
    r.supportOrder = match || null;
  });
  return recent;
}

function getDashboardData(fromDate,toDate,fromTime,toTime){
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if(!sheet)throw new Error("Web app dev sheet not found.");
  ensureWebAppDevColumns_(sheet);
  const tz=Session.getScriptTimeZone()||"Asia/Kolkata";
  let rs=null,re=null;
  const f=String(fromDate||"").trim(),t=String(toDate||"").trim();
  const ft=String(fromTime||"00:00").trim()||"00:00";
  const tt=String(toTime||"23:59").trim()||"23:59";

  if(f&&t){
    const a=f.split("-"),b=t.split("-");
    if(a.length===3&&b.length===3){
      const fp=ft.split(":"),tp=tt.split(":");
      rs=new Date(+a[0],+a[1]-1,+a[2],Number(fp[0]||0),Number(fp[1]||0),0,0);
      re=new Date(+b[0],+b[1]-1,+b[2],Number(tp[0]||23),Number(tp[1]||59),59,999);
    }
  }
  if(!rs||!re||re<rs){
    const d=new Date();
    rs=new Date(d.getFullYear(),d.getMonth(),d.getDate()-6,0,0,0,0);
    re=new Date(d.getFullYear(),d.getMonth(),d.getDate(),23,59,59,999);
  }

  let total=0,today=0,requestingOrder=0,acceptedOrdered=0,noNeed=0,completed=0,closed=0,closedQty=0;
  const recent=[];
  const todayText=Utilities.formatDate(new Date(),tz,"yyyy-MM-dd");

  if(sheet.getLastRow()>=2){
    sheet.getRange(2,1,sheet.getLastRow()-1,32).getValues().forEach(function(row,i){
      if(!row[COL_TIMESTAMP-1]&&!row[COL_ITEM_ID-1]&&!row[COL_TICKET_ID-1])return;
      total++;
      const stamp=row[COL_TIMESTAMP-1],status=String(row[COL_STATUS-1]||"").trim(),ns=normalizeStatus_(status),qty=getRequestQty_(row[COL_QTY-1]);
      if(stamp instanceof Date){
        if(Utilities.formatDate(stamp,tz,"yyyy-MM-dd")===todayText)today++;
        if(stamp>=rs&&stamp<=re){
          recent.push({
            rowNumber:i+2,timestamp:getSheetDateString_(stamp),date:getSheetDateString_(stamp),requestDateIso:stamp.toISOString(),
            email:String(row[COL_EMAIL-1]||""),name:String(row[COL_NAME-1]||""),
            mobile:String(row[COL_MOBILE-1]||""),ticketId:String(row[COL_TICKET_ID-1]||""),
            itemId:String(row[COL_ITEM_ID-1]||""),itemName:String(row[COL_ITEM_NAME-1]||""),
            stationId:String(row[COL_AREA-1]||""),qty:String(row[COL_QTY-1]||""),
            status:status,requestStatus:status,remark:String(row[COL_REMARK-1]||""),
            requestType:String(row[COL_REQUEST_TYPE-1]||""),requestDetail:String(row[COL_REQUEST_DETAIL-1]||"")
          });
        }
      }
      if(ns==="requesting order")requestingOrder++;
      if(isAcceptedOrderedStatus_(status))acceptedOrdered++;
      if(ns==="no need")noNeed++;
      if(ns==="completed")completed++;
      if(ns==="closed"){closed++;closedQty+=qty;}
    });
  }

  recent.reverse();
  getSupportOrderMatchesForRequests_(recent);
  const oc=getProblemObservationDashboardCounts_();
  return {
    total:total,today:today,requestingOrder:requestingOrder,acceptedOrdered:acceptedOrdered,
    noNeed:noNeed,completed:completed,closed:closed,closedQty:closedQty,
    recent:recent.slice(0,10),
    supportOrders:getSupportOrderDashboardData_(SUPPORT_ORDER_TOOL_NO, rs, re),
    observationRecent:getDashboardObservationRecent_(rs,re),
    observationOpen:oc.open,observationInProgress:oc.inProgress,observationClosed:oc.closed,
    observationTotalRaised:oc.totalRaised,observationTotalClosed:oc.totalClosed,
    recentFrom:Utilities.formatDate(rs,tz,"dd-MM-yyyy"),
    recentTo:Utilities.formatDate(re,tz,"dd-MM-yyyy")
  };
}

// =========================================================
// UPDATE DASHBOARD REQUEST
// MATCH:
// Ticket ID = Column P
// Item ID   = Column G
// UPDATE:
// Status = Column I
// Remark = Column K
// =========================================================


// =========================================================
// ACTIVITY-WISE EMAIL ROUTING (Settings -> Email Notifications)
// Stored in Script Properties so the existing Mail ID's sheet stays compatible.
// =========================================================
const NOTIFICATION_CONFIG_KEY_ = "FALCON_ACTIVITY_NOTIFICATION_CONFIG_V1";
const NOTIFICATION_ACTIVITY_DEFAULTS_ = {
  newMaterialRequest: {label:"New Material Request", enabled:true, mode:"external", internalEmails:"", externalEmails:""},
  materialRequestStatus: {label:"Material Request Status", enabled:true, mode:"external", internalEmails:"", externalEmails:""},
  receivedSpare: {label:"Received Spare Update", enabled:true, mode:"external", internalEmails:"", externalEmails:""},
  pvcBelt: {label:"PVC Belt Received / Consumed", enabled:true, mode:"external", internalEmails:"", externalEmails:""},
  pmWork: {label:"PM Work (legacy combined)", enabled:true, mode:"both", internalEmails:"", externalEmails:""},
  pmWorkCompleted: {label:"PM Work Completed", enabled:true, mode:"both", internalEmails:"", externalEmails:""},
  pmWorkOpen: {label:"PM Work Open / Pending", enabled:true, mode:"both", internalEmails:"", externalEmails:""},
  pmWorkOverdue: {label:"PM Work Overdue", enabled:true, mode:"both", internalEmails:"", externalEmails:""},
  problemObservation: {label:"Problem Observation", enabled:true, mode:"external", internalEmails:"", externalEmails:""},
  scheduledDowntime: {label:"Scheduled Downtime Activities", enabled:true, mode:"external", internalEmails:"", externalEmails:""},
  shiftRoster: {label:"Shift Roster Changes", enabled:true, mode:"both", internalEmails:"", externalEmails:""}
};
function getActivityNotificationConfig_(sessionId) {
  if (sessionId) requireAdminSession_(sessionId);
  const saved = PropertiesService.getScriptProperties().getProperty(NOTIFICATION_CONFIG_KEY_);
  let parsed = {};
  try { parsed = saved ? JSON.parse(saved) : {}; } catch (e) { parsed = {}; }
  const result = {};
  Object.keys(NOTIFICATION_ACTIVITY_DEFAULTS_).forEach(function(key) {
    result[key] = Object.assign({}, NOTIFICATION_ACTIVITY_DEFAULTS_[key], parsed[key] || {});
  });
  return result;
}
function saveActivityNotificationConfig(sessionId, config) {
  requireAdminSession_(sessionId);
  if (!config || typeof config !== "object") throw new Error("Notification configuration missing.");
  const cleaned = {};
  Object.keys(NOTIFICATION_ACTIVITY_DEFAULTS_).forEach(function(key) {
    const d = NOTIFICATION_ACTIVITY_DEFAULTS_[key];
    const x = config[key] || {};
    const mode = ["off","internal","external","both"].indexOf(String(x.mode || d.mode).toLowerCase()) >= 0 ? String(x.mode || d.mode).toLowerCase() : d.mode;
    const internalEmails = parseEmailList_(x.internalEmails || "");
    const externalEmails = parseEmailList_(x.externalEmails || "");
    internalEmails.concat(externalEmails).forEach(function(email) { if (!isValidEmail_(email)) throw new Error("Invalid email address: " + email); });
    cleaned[key] = {label:d.label, enabled:x.enabled === true || String(x.enabled) === "true", mode:mode, internalEmails:internalEmails.join(","), externalEmails:externalEmails.join(",")};
  });
  PropertiesService.getScriptProperties().setProperty(NOTIFICATION_CONFIG_KEY_, JSON.stringify(cleaned));
  return {success:true, message:"Activity-wise email preferences saved."};
}
function getConfiguredEmailRecipients_(activityKey) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const emailSheet = ss.getSheetByName(EMAIL_SHEET_NAME);
  let legacyTo = [], legacyCc = [];
  if (emailSheet && emailSheet.getLastRow() >= 2) {
    const rows = emailSheet.getRange(2, 2, emailSheet.getLastRow() - 1, 2).getValues();
    rows.forEach(function(r) { const a=String(r[0]||"").trim(); const b=String(r[1]||"").trim(); if(a)legacyTo.push(a); if(b)legacyCc.push(b); });
  }
  // Calls without a key preserve legacy routing until mapped to a specific event.
  if (!activityKey || !NOTIFICATION_ACTIVITY_DEFAULTS_[activityKey]) return {to:[...new Set(legacyTo)],cc:[...new Set(legacyCc)]};
  const cfg = getActivityNotificationConfig_()[activityKey];
  if (!cfg || !cfg.enabled || cfg.mode === "off") return {to:[],cc:[]};
  const internal = parseEmailList_(cfg.internalEmails || "");
  const external = parseEmailList_(cfg.externalEmails || "");
  const to = [];
  if (cfg.mode === "internal" || cfg.mode === "both") to.push.apply(to, internal);
  if (cfg.mode === "external" || cfg.mode === "both") to.push.apply(to, external.length ? external : legacyTo);
  const cc = (cfg.mode === "external" || cfg.mode === "both") ? legacyCc : [];
  return {to:[...new Set(to)],cc:[...new Set(cc.filter(function(e){return to.indexOf(e)<0;}))]};
}



function responsiblePersonForDepartment_(department) {
  const d = String(department || "").trim().toLowerCase();
  if (d === "electrical") return ["ramji"];
  if (d === "mechanical") return ["akshit", "akshith"];
  if (d === "safety") return ["ramji", "akshit", "akshith"];
  return [];
}

function canManageMaterialRequest_(session, department) {
  if (!session) return false;
  if (String(session.role || "").trim().toLowerCase() === "admin") return true;
  const name = String(session.name || "").trim().toLowerCase();
  const email = String(session.email || "").trim().toLowerCase();
  const keys = responsiblePersonForDepartment_(department);
  return keys.some(function(k) { return name.indexOf(k) >= 0 || email.indexOf(k) >= 0; });
}

function saveCompletedReceivedSpare_(sheet, rowNumber, receivedData, materialType) {
  receivedData = receivedData || {};
  const receivedDate = String(receivedData.receivedDate || "").trim();
  const receivedTime = String(receivedData.receivedTime || "").trim();
  const awbNo = String(receivedData.awbNo || "").trim();
  const courier = String(receivedData.courier || "").trim();
  const challanNo = String(receivedData.challanNo || "").trim();
  const lastDate = String(receivedData.lastDate || "").trim();
  if (!receivedDate || !receivedTime || !awbNo || !challanNo) {
    return { success: false, message: "Received Date, Time, AWB No. and Challan No. are required for Completed." };
  }
  const type = String(materialType || "").trim();
  if (["Scrap", "Consumable", "Returnable"].indexOf(type) < 0) {
    return { success: false, message: "Select Scrap, Consumable or Returnable before completing the request." };
  }
  if ((type === "Scrap" || type === "Returnable") && !lastDate) {
    return { success: false, message: "Last Date is required for Scrap / Returnable." };
  }
  const itemId = String(sheet.getRange(rowNumber, COL_ITEM_ID).getDisplayValue() || "").trim();
  const itemName = String(sheet.getRange(rowNumber, COL_ITEM_NAME).getDisplayValue() || "").trim();
  const qty = Number(sheet.getRange(rowNumber, COL_QTY).getValue() || 0);
  const ticketId = String(sheet.getRange(rowNumber, COL_TICKET_ID).getDisplayValue() || "").trim();
  const operatorName = String(sheet.getRange(rowNumber, COL_NAME).getDisplayValue() || "").trim();
  const sh = ensureReceivedSpareSheet_();
  const dt = new Date(receivedDate + "T" + receivedTime);
  if (isNaN(dt.getTime())) return { success: false, message: "Invalid Received Date / Time." };
  const row = [
    new Date(), new Date(receivedDate + "T00:00:00"), dt, "Material Request", ticketId,
    awbNo, courier, challanNo, lastDate, itemId, itemName, qty, "", operatorName,
    operatorName, "NA", "", "", String(sheet.getRange(rowNumber, COL_EMAIL).getDisplayValue() || "").trim(),
    String(sheet.getRange(rowNumber, COL_NAME).getDisplayValue() || "").trim()
  ];
  sh.getRange(sh.getLastRow() + 1, 1, 1, 20).setValues([row]);
  SpreadsheetApp.flush();
  return { success: true };
}

function updateDashboardRequest(sessionId, ticketId, itemId, status, remark, rowNumber, materialType, receivedData) {
  requirePortalPermission_(sessionId, "spareParts", "approveReject");

  const session = validateSession_(sessionId);
  if (!session) return { success: false, message: "Session expired. Please login again." };

  ticketId = String(ticketId || "").trim();
  itemId = String(itemId || "").trim();
  status = String(status || "").trim();
  remark = String(remark || "").trim();
  materialType = String(materialType || "").trim();
  rowNumber = Number(rowNumber || 0);

  if (!ticketId) return { success: false, message: "Ticket ID is required." };
  if (!itemId) return { success: false, message: "Item ID is required." };
  if (!status) return { success: false, message: "New status is required." };

  const canonicalStatus = canonicalRequestStatus_(status);

  if (!canonicalStatus) {
    return { success: false, message: "Invalid request status." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Web app dev sheet not found." };
  }

  ensureWebAppDevColumns_(sheet);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: false,
      message: "No material request records found."
    };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 29).getValues();
  let matchedRow = -1;

  if (rowNumber >= 2 && rowNumber <= lastRow) {
    const selected = sheet
      .getRange(rowNumber, 1, 1, 29)
      .getValues()[0];

    const selectedTicketId =
      String(selected[COL_TICKET_ID - 1] || "").trim();

    const selectedItemId =
      String(selected[COL_ITEM_ID - 1] || "").trim();

    if (
      selectedTicketId.toLowerCase() === ticketId.toLowerCase() &&
      selectedItemId.toLowerCase() === itemId.toLowerCase()
    ) {
      matchedRow = rowNumber;
    }
  }

  if (matchedRow === -1) {
    for (let i = 0; i < values.length; i++) {
      const rowTicketId =
        String(values[i][COL_TICKET_ID - 1] || "").trim();

      const rowItemId =
        String(values[i][COL_ITEM_ID - 1] || "").trim();

      if (
        rowTicketId.toLowerCase() === ticketId.toLowerCase() &&
        rowItemId.toLowerCase() === itemId.toLowerCase()
      ) {
        matchedRow = i + 2;
        break;
      }
    }
  }

  if (matchedRow === -1) {
    return {
      success: false,
      message: "No matching request found for Ticket ID + Item ID."
    };
  }

  const currentStatus =
    String(
      sheet.getRange(matchedRow, COL_STATUS).getDisplayValue() || ""
    ).trim();

  const currentCanonical = canonicalRequestStatus_(currentStatus);

  const statusMailColumn = getStatusMailColumn_(canonicalStatus);

  if (
    statusMailColumn &&
    sheet.getRange(matchedRow, statusMailColumn).getValue()
  ) {
    return {
      success: false,
      message:
        "Status '" +
        canonicalStatus +
        "' has already been updated for this request."
    };
  }

  if (!isAllowedStatusTransition_(currentCanonical, canonicalStatus)) {
    return {
      success: false,
      message:
        "Invalid status flow. Required flow: Requesting Order → Accepted / Rejected → Ordered → Completed → Closed. Current status is '" +
        (currentCanonical || currentStatus || "blank") + "'."
    };
  }

  if (canonicalStatus === "Rejected") {
    if (!remark) {
      return {
        success: false,
        message: "Rejection reason is mandatory. Please enter the reason in Remark."
      };
    }
  }

  if (canonicalStatus === STATUS_ACCEPTED || canonicalStatus === STATUS_ORDERED) {
    if (["Scrap", "Consumable", "Returnable"].indexOf(materialType) < 0) {
      return { success: false, message: "Select Material Type: Scrap, Consumable or Returnable." };
    }
    sheet.getRange(matchedRow, COL_REQUEST_TYPE).setValue(materialType);
  }

  if (canonicalStatus === STATUS_COMPLETED) {
    const existingType = String(sheet.getRange(matchedRow, COL_REQUEST_TYPE).getDisplayValue() || "").trim();
    const completionType = materialType || existingType;
    const receivedResult = saveCompletedReceivedSpare_(sheet, matchedRow, receivedData, completionType);
    if (!receivedResult.success) return receivedResult;
    sheet.getRange(matchedRow, COL_REQUEST_TYPE).setValue(completionType);
  }

  if (canonicalStatus === "Rejected") {
    // Keep the reason in the main Remark column as requested.
    sheet.getRange(matchedRow, COL_REMARK).setValue(remark);
    // Preserve the existing rejection-email logic as well.
    sheet.getRange(matchedRow, COL_REJECTION_REMARK).setValue(remark);
    sheet.getRange(matchedRow, COL_CLOSING_REMARK).setValue("");
  } else {
    // Existing Accepted / Ordered / Completed / Closed behavior remains unchanged.
    sheet.getRange(matchedRow, COL_REMARK).setValue("");
    sheet.getRange(matchedRow, COL_CLOSING_REMARK).setValue("");
  }
  sheet.getRange(matchedRow, COL_STATUS).setValue(canonicalStatus);

  SpreadsheetApp.flush();

  // Portal update sends the status reply immediately.
  sendStatusUpdateReply_(
    sheet,
    matchedRow,
    currentCanonical || currentStatus
  );

  const sent = statusMailColumn && !!sheet.getRange(matchedRow, statusMailColumn).getValue();
  return {
    success: true,
    message: sent
      ? "Request updated to '" + canonicalStatus + "' and notification sent successfully."
      : "Request updated to '" + canonicalStatus + "'."
  };
}

// =========================================================
// CLOSE DASHBOARD REQUEST
// =========================================================

function closeDashboardRequest(
  sessionId,
  ticketId,
  itemId,
  remark,
  rowNumber
) {
  requirePortalPermission_(sessionId, "spareParts", "delete");

  return updateDashboardRequest(
    sessionId,
    ticketId,
    itemId,
    STATUS_CLOSED,
    remark,
    rowNumber
  );
}

// =========================================================
// SEARCH REQUESTS
// =========================================================


function searchRequests(sessionId, filters) {
  const session = validateSession_(sessionId);

  if (!session) {
    return { success: false, message: "Session expired. Please login again.", records: [] };
  }

  filters = filters || {};

  const ticketId = String(filters.ticketId || "").trim().toLowerCase();
  const itemId = String(filters.itemId || "").trim().toLowerCase();
  const email = String(filters.email || "").trim().toLowerCase();
  const status = String(filters.status || "").trim().toLowerCase();
  const requestType = String(filters.requestType || "").trim().toLowerCase();

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Web app dev sheet not found.", records: [] };
  }

  ensureWebAppDevColumns_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 23).getValues();
  const timezone = Session.getScriptTimeZone() || "Asia/Kolkata";
  const records = [];

  values.forEach(function(row) {
    if (!row[COL_TIMESTAMP - 1] && !row[COL_ITEM_ID - 1] && !row[COL_TICKET_ID - 1]) return;

    const rowTicketId = String(row[COL_TICKET_ID - 1] || "").trim().toLowerCase();
    const rowItemId = String(row[COL_ITEM_ID - 1] || "").trim().toLowerCase();
    const rowEmail = String(row[COL_EMAIL - 1] || "").trim().toLowerCase();
    const rowStatus = String(row[COL_STATUS - 1] || "").trim();
    const rowStatusLower = rowStatus.toLowerCase();
    const rowType = String(row[COL_REQUEST_TYPE - 1] || "").trim();

    if (ticketId && rowTicketId.indexOf(ticketId) === -1) return;
    if (itemId && rowItemId.indexOf(itemId) === -1) return;
    if (email && rowEmail.indexOf(email) === -1) return;

    if (requestType && rowType.toLowerCase() !== requestType) return;

    if (status) {
      let statusMatch = false;

      if (status === "accepted/ordered" ||
          status === "accepted / ordered" ||
          status === "accepted" ||
          status === "ordered") {
        statusMatch = isAcceptedOrderedStatus_(rowStatus);
      } else {
        statusMatch = rowStatusLower === status;
      }

      if (!statusMatch) return;
    }

    records.push({
      rowNumber: values.indexOf(row) + 2,
      date: row[COL_TIMESTAMP - 1] instanceof Date
        ? Utilities.formatDate(row[COL_TIMESTAMP - 1], timezone, "dd-MM-yyyy HH:mm")
        : String(row[COL_TIMESTAMP - 1] || ""),
      email: String(row[COL_EMAIL - 1] || ""),
      name: String(row[COL_NAME - 1] || ""),
      mobile: String(row[COL_MOBILE - 1] || ""),
      itemName: String(row[COL_ITEM_NAME - 1] || ""),
      ticketId: String(row[COL_TICKET_ID - 1] || ""),
      itemId: String(row[COL_ITEM_ID - 1] || ""),
      stationId: String(row[COL_AREA - 1] || ""),
      qty: String(row[COL_QTY - 1] || ""),
      status: rowStatus,
      remark: String(row[COL_REMARK - 1] || ""),
      closingRemark: String(row[COL_CLOSING_REMARK - 1] || ""),
      rejectionRemark: String(row[COL_REJECTION_REMARK - 1] || ""),
      spareRequestStatus: String(row[COL_SPARE_STATUS - 1] || ""),
      toolNo: String(row[COL_TOOL_NO - 1] || ""),
      requestId: String(row[COL_REQUEST_ID - 1] || ""),
      requestType: rowType,
      requestDetail: String(row[COL_REQUEST_DETAIL - 1] || "")
    });
  });

  records.reverse();

  let summary = null;
  if (filters.fromDate && filters.toDate) {
    const summaryResult = getBreakdownDateRangeSummary(
      sessionId,
      filters.fromDate,
      filters.toDate
    );
    if (summaryResult && summaryResult.success) {
      summary = summaryResult;
    }
  }

  return {
    success: true,
    records: records.slice(0, 500),
    summary: summary
  };
}

// =========================================================
// NORMALIZE ITEM MASTER ID
// =========================================================

function normalizeItemMasterId_(value) {
  return String(
    value == null ? "" : value
  )
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

// =========================================================
// GET AUTOMATIC SPARE TRACKER STOCK DATA
//
// Automatic Spare Tracker columns:
// C = Item Master ID
// D = Current Stock
// J = Purchased Qty
//
// Item Search matches Search Item column E (Item Master ID)
// with Automatic Spare Tracker column C.
// Current Stock comes ONLY from Automatic Spare Tracker D.
// Purchased Qty comes ONLY from Automatic Spare Tracker J.
// =========================================================

function getSearchItemStockMap_(sheet) {
  const stockMap = {};

  const tracker = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(STOCK_SHEET_NAME);

  if (!tracker) {
    return stockMap;
  }

  const lastRow = tracker.getLastRow();

  if (lastRow < 2) {
    return stockMap;
  }

  // Read A:J from Automatic Spare Tracker.
  // C = Item Master ID, D = Current Stock, J = Purchased Qty.
  const values = tracker.getRange(
    2,
    1,
    lastRow - 1,
    10
  ).getDisplayValues();

  values.forEach(function(row) {
    const itemMasterId = normalizeItemMasterId_(row[2]);

    if (!itemMasterId) {
      return;
    }

    const currentStock = String(row[3] == null ? "" : row[3]).trim();
    const purchasedQty = String(row[9] == null ? "" : row[9]).trim();

    const existing = stockMap[itemMasterId];

    // If an Item Master ID occurs more than once, prefer the row that
    // contains actual stock/purchase values instead of an empty row.
    if (!existing) {
      stockMap[itemMasterId] = {
        currentStock: currentStock,
        purchasedQty: purchasedQty
      };
      return;
    }

    if (existing.currentStock === "" && currentStock !== "") {
      existing.currentStock = currentStock;
    }

    if (existing.purchasedQty === "" && purchasedQty !== "") {
      existing.purchasedQty = purchasedQty;
    }
  });

  return stockMap;
}

// =========================================================
// SEARCH ITEM MASTER
//
// Search Item columns:
//
// A = Tool No
// B = Station
// C = Part Name
// D = Model No
// E = Item Master ID
// F = Matl
// G = Size
// H = Make
//
// STOCK MATCH:
//
// Search Item Item Master ID (E)
// -> Automatic Spare Tracker Item Master ID (C)
//
// Purchased Qty:
// Automatic Spare Tracker J
//
// Current Stock:
// Automatic Spare Tracker D
//
// Duplicate Item Master IDs:
// Only one result is returned.
// =========================================================

function searchItemMaster(
  sessionId,
  query
) {
  const session =
    validateSession_(sessionId);

  if (!session) {
    return {
      success: false,
      message:
        "Session expired. Please login again.",
      items: []
    };
  }

  query =
    String(query || "")
      .trim()
      .toLowerCase();

  if (!query) {
    return {
      success: true,
      items: []
    };
  }

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SEARCH_ITEM_SHEET_NAME
      );

  if (!sheet) {
    return {
      success: false,
      message:
        "Sheet not found: " +
        SEARCH_ITEM_SHEET_NAME,
      items: []
    };
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {
    return {
      success: true,
      items: []
    };
  }

  // =======================================================
  // READ A:J
  // =======================================================

  const values =
    sheet.getRange(
      2,
      1,
      lastRow - 1,
      10
    ).getDisplayValues();

  // =======================================================
  // AUTOMATIC SPARE TRACKER STOCK MAP
  // C = Item Master ID
  // D = Current Stock
  // J = Purchased Qty
  // =======================================================

  const stockMap =
    getSearchItemStockMap_(sheet);

  const results = [];

  // Prevent duplicate Item Master IDs
  const seenItemIds = {};

  values.forEach(function(row) {

    const toolNo =
      String(row[0] || "").trim();

    const station =
      String(row[1] || "").trim();

    const partName =
      String(row[2] || "").trim();

    const modelNo =
      String(row[3] || "").trim();

    const itemMasterId =
      String(row[4] || "").trim();

    const matl =
      String(row[5] || "").trim();

    const size =
      String(row[6] || "").trim();

    const make =
      String(row[7] || "").trim();

    const searchableText = [
      toolNo,
      station,
      partName,
      modelNo,
      itemMasterId,
      matl,
      size,
      make
    ]
      .map(function(value) {
        return String(value || "")
          .toLowerCase();
      })
      .join(" ");

    if (
      searchableText.indexOf(query) === -1
    ) {
      return;
    }

    // =====================================================
    // DUPLICATE ITEM MASTER ID HANDLING
    // =====================================================

    const normalizedItemId =
      normalizeItemMasterId_(
        itemMasterId
      );

    let uniqueKey;

    if (normalizedItemId) {
      uniqueKey =
        "ITEMID|" +
        normalizedItemId;
    } else {
      uniqueKey =
        "BLANK|" +
        normalizeToolNo_(toolNo) +
        "|" +
        modelNo.toLowerCase() +
        "|" +
        partName.toLowerCase();
    }

    if (seenItemIds[uniqueKey]) {
      return;
    }

    seenItemIds[uniqueKey] = true;

    // =====================================================
    // SEARCH ITEM STOCK MATCH
    //
    // Item Master ID E
    // ->
    // Item Master ID E
    // =====================================================

    const stock =
      stockMap[
        normalizeItemMasterId_(
          itemMasterId
        )
      ];

    let purchasedQty = "-";
    let currentStock = "-";

    if (stock) {

      purchasedQty =
        stock.purchasedQty !== ""
          ? stock.purchasedQty
          : "-";

      currentStock =
        stock.currentStock !== ""
          ? stock.currentStock
          : "-";
    }

    results.push({

      toolNo:
        toolNo,

      station:
        station,

      partName:
        partName,

      modelNo:
        modelNo,

      itemMasterId:
        itemMasterId,

      matl:
        matl,

      size:
        size,

      make:
        make,

      purchasedQty:
        purchasedQty,

      currentStock:
        currentStock
    });
  });

  return {
    success: true,
    items:
      results.slice(0, 1000)
  };
}


// =========================================================
// SEARCH ITEM MASTER - MULTI CELL SEARCH
// =========================================================
function searchItemMasterFilters(sessionId, filters) {
  const session = validateSession_(sessionId);

  if (!session) {
    return { success: false, message: "Session expired. Please login again.", items: [] };
  }

  filters = filters || {};

  const f = {
    toolNo: String(filters.toolNo || "").trim().toLowerCase(),
    station: String(filters.station || "").trim().toLowerCase(),
    partName: String(filters.partName || "").trim().toLowerCase(),
    make: String(filters.make || "").trim().toLowerCase(),
    size: String(filters.size || "").trim().toLowerCase(),
    itemMasterId: String(filters.itemMasterId || "").trim().toLowerCase(),
    query: String(filters.query || "").trim().toLowerCase()
  };

  if (!Object.keys(f).some(function(k) { return !!f[k]; })) {
    return { success: true, items: [] };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SEARCH_ITEM_SHEET_NAME);

  if (!sheet) {
    return { success: false, message: "Sheet not found: " + SEARCH_ITEM_SHEET_NAME, items: [] };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, items: [] };

  const values = sheet.getRange(2, 1, lastRow - 1, 10).getDisplayValues();
  const stockMap = getSearchItemStockMap_(sheet);
  const seen = {};
  const results = [];

  function matches(value, needle) {
    return !needle || String(value || "").toLowerCase().indexOf(needle) !== -1;
  }

  values.forEach(function(row) {
    const item = {
      toolNo: String(row[0] || "").trim(),
      station: String(row[1] || "").trim(),
      partName: String(row[2] || "").trim(),
      modelNo: String(row[3] || "").trim(),
      itemMasterId: String(row[4] || "").trim(),
      matl: String(row[5] || "").trim(),
      size: String(row[6] || "").trim(),
      make: String(row[7] || "").trim()
    };

    const searchable = [
      item.toolNo, item.station, item.partName, item.modelNo,
      item.itemMasterId, item.matl, item.size, item.make
    ].join(" ").toLowerCase();

    if (!matches(item.toolNo, f.toolNo) ||
        !matches(item.station, f.station) ||
        !matches(item.partName, f.partName) ||
        !matches(item.make, f.make) ||
        !matches(item.size, f.size) ||
        !matches(item.itemMasterId, f.itemMasterId) ||
        !matches(searchable, f.query)) {
      return;
    }

    const key = normalizeItemMasterId_(item.itemMasterId) ||
      ("BLANK|" + normalizeToolNo_(item.toolNo) + "|" +
       item.modelNo.toLowerCase() + "|" + item.partName.toLowerCase());

    if (seen[key]) return;
    seen[key] = true;

    const stock = stockMap[normalizeItemMasterId_(item.itemMasterId)];

    results.push({
      toolNo: item.toolNo,
      station: item.station,
      partName: item.partName,
      modelNo: item.modelNo,
      itemMasterId: item.itemMasterId,
      matl: item.matl,
      size: item.size,
      make: item.make,
      purchasedQty: stock && stock.purchasedQty !== "" ? stock.purchasedQty : "-",
      currentStock: stock && stock.currentStock !== "" ? stock.currentStock : "-"
    });
  });

  return { success: true, items: results.slice(0, 1000) };
}

// =========================================================
// REPORT SUMMARY
// =========================================================


function parseReportDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return new Date(value.getTime());
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number" && isFinite(value)) {
    // Google Sheets serial date fallback.
    var dNum = new Date(Math.round((value - 25569) * 86400000));
    return isNaN(dNum.getTime()) ? null : dNum;
  }
  var s = String(value).trim();
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;

  // dd-MM-yyyy / dd/MM/yyyy with optional time.
  var m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    var hh = Number(m[4] || 0), mm = Number(m[5] || 0), ss = Number(m[6] || 0);
    var parsed = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), hh, mm, ss);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function parseReportRange_(fromDate, toDate) {
  var a = String(fromDate || "").trim().split("-"),
      b = String(toDate || "").trim().split("-");
  if (a.length === 3 && b.length === 3) {
    var rs = new Date(Number(a[0]), Number(a[1]) - 1, Number(a[2]));
    var re = new Date(Number(b[0]), Number(b[1]) - 1, Number(b[2]) + 1);
    if (!isNaN(rs.getTime()) && !isNaN(re.getTime()) && re > rs) {
      return { start: rs, end: re };
    }
  }
  var now = new Date();
  return {
    start: new Date(now.getFullYear(), now.getMonth(), 1),
    end: new Date(now.getFullYear(), now.getMonth() + 1, 1)
  };
}

function getReportSummary(sessionId, fromDate, toDate) {
  requireAdminSession_(sessionId);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Web app dev sheet not found.");
  ensureWebAppDevColumns_(sheet);

  var tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  var range = parseReportRange_(fromDate, toDate);
  var rs = range.start, re = range.end;

  var total = 0, requestingOrder = 0, acceptedOrdered = 0, completed = 0,
      pendingScrapQty = 0, pendingReplacementQty = 0, closedReplacementScrapQty = 0;

  if (sheet.getLastRow() >= 2) {
    var lastCol = Math.max(23, sheet.getLastColumn());
    sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues().forEach(function(row) {
      var d = parseReportDate_(row[COL_TIMESTAMP - 1]);
      if (!d || d < rs || d >= re) return;
      total++;
      var qty = getRequestQty_(row[COL_QTY - 1]),
          status = String(row[COL_STATUS - 1] || "").trim(),
          ns = normalizeStatus_(status),
          type = normalizeRequestType_(row[COL_REQUEST_TYPE - 1]);
      if (ns === "requesting order") requestingOrder++;
      if (isAcceptedOrderedStatus_(status)) acceptedOrdered++;
      if (ns === "completed") completed++;
      if (type === "scrap" && isOpenRequestStatus_(status)) pendingScrapQty += qty;
      if (type === "replacement" && isOpenRequestStatus_(status)) pendingReplacementQty += qty;
      if ((type === "scrap" || type === "replacement") && ns === "closed") closedReplacementScrapQty += qty;
    });
  }

  var from = Utilities.formatDate(rs, tz, "yyyy-MM-dd");
  var to = Utilities.formatDate(new Date(re.getTime() - 1), tz, "yyyy-MM-dd");
  var ops = getBreakdownDateRangeSummary(sessionId, from, to);
  var owners = getBreakdownOwnerSummary_(sessionId, from, to);

  return {
    success: true,
    total: total,
    requestingOrder: requestingOrder,
    acceptedOrdered: acceptedOrdered,
    completed: completed,
    pendingScrapQty: pendingScrapQty,
    pendingReplacementQty: pendingReplacementQty,
    closedReplacementScrapQty: closedReplacementScrapQty,
    operations: ops && ops.success ? ops : {},
    ownerBreakdown: owners,
    fromDate: from,
    toDate: to
  };
}

function getSettings(sessionId) {
  requireAdminSession_(sessionId);

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      EMAIL_SHEET_NAME
    );

  let toEmails = [];
  let ccEmails = [];

  if (sheet) {

    const lastRow =
      sheet.getLastRow();

    if (lastRow >= 2) {

      const values =
        sheet.getRange(
          2,
          2,
          lastRow - 1,
          2
        ).getValues();

      values.forEach(function(row) {

        const to =
          String(row[0] || "").trim();

        const cc =
          String(row[1] || "").trim();

        if (to) {
          toEmails.push(to);
        }

        if (cc) {
          ccEmails.push(cc);
        }
      });
    }
  }

  return {

    toEmails:
      [...new Set(toEmails)],

    ccEmails:
      [...new Set(ccEmails)],
    notificationConfig: getActivityNotificationConfig_()
  };
}

// =========================================================
// SAVE SETTINGS
// =========================================================

function saveSettings(sessionId, data) {
  requireAdminSession_(sessionId);

  if (!data) {
    throw new Error(
      "Settings data missing."
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      EMAIL_SHEET_NAME
    );

  if (!sheet) {
    throw new Error(
      "Mail ID\'s sheet not found."
    );
  }

  const toEmails =
    parseEmailList_(
      data.toEmails
    );

  const ccEmails =
    parseEmailList_(
      data.ccEmails
    );

  const allEmails =
    toEmails.concat(
      ccEmails
    );

  allEmails.forEach(function(email) {

    if (!isValidEmail_(email)) {
      throw new Error(
        "Invalid email address: " +
        email
      );
    }

  });

  const maxRows =
    sheet.getMaxRows();

  if (maxRows >= 2) {

    sheet
      .getRange(
        2,
        2,
        maxRows - 1,
        2
      )
      .clearContent();

  }

  const maxLength =
    Math.max(
      toEmails.length,
      ccEmails.length
    );

  if (maxLength > 0) {

    const output = [];

    for (
      let i = 0;
      i < maxLength;
      i++
    ) {

      output.push([
        toEmails[i] || "",
        ccEmails[i] || ""
      ]);

    }

    sheet
      .getRange(
        2,
        2,
        output.length,
        2
      )
      .setValues(
        output
      );
  }

  if (data.notificationConfig) saveActivityNotificationConfig(sessionId, data.notificationConfig);

  return {
    success: true,
    message:
      "Settings saved successfully."
  };
}

// =========================================================
// PARSE EMAIL LIST
// =========================================================

function parseEmailList_(value) {

  if (Array.isArray(value)) {

    return [
      ...new Set(
        value
          .map(function(item) {
            return String(
              item || ""
            ).trim();
          })
          .filter(Boolean)
      )
    ];
  }

  return [
    ...new Set(
      String(value || "")
        .split(/[\n,;]+/)
        .map(function(item) {
          return item.trim();
        })
        .filter(Boolean)
    )
  ];
}

// =========================================================
// PASSWORD HASH
// =========================================================

function hashPassword_(
  password,
  salt
) {

  const raw =
    String(salt) +
    "|" +
    String(password);

  const digest =
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      raw,
      Utilities.Charset.UTF_8
    );

  return digest
    .map(function(byte) {

      return (
        byte < 0
          ? byte + 256
          : byte
      )
        .toString(16)
        .padStart(2, "0");

    })
    .join("");
}

// =========================================================
// LOGIN
// =========================================================

function loginUser(
  email,
  password
) {

  email =
    String(email || "")
      .trim()
      .toLowerCase();

  password =
    String(password || "");

  const temporary =
    getTemporaryAdminCredential_(email);

  if (!email || !password) {

    logLoginActivity_(
      email,
      "",
      "FAILED",
      "Email and password are required.",
      ""
    );

    return {
      success: false,
      message:
        "Email and password are required."
    };
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      USER_SHEET_NAME
    );

  if (!sheet) {
    throw new Error(
      "User Management sheet not found."
    );
  }

  const lastRow =
    sheet.getLastRow();

  if (lastRow < 2) {

    logLoginActivity_(
      email,
      "",
      "FAILED",
      "No users configured.",
      ""
    );

    return {
      success: false,
      message:
        "No users configured."
    };
  }

  const data =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  let user = null;

  for (
    let i = data.length - 1;
    i >= 0;
    i--
  ) {

    const userEmail =
      String(
        data[i][0] || ""
      )
        .trim()
        .toLowerCase();

    if (
      userEmail === email
    ) {

      user = {

        email:
          userEmail,

        name:
          String(
            data[i][1] || ""
          ),

        role:
          String(
            data[i][2] || ""
          ),

        passwordHash:
          String(
            data[i][3] || ""
          ),

        passwordSalt:
          String(
            data[i][4] || ""
          ),

        status:
          String(
            data[i][5] || ""
          )
      };

      break;
    }
  }

  if (!user) {

    logLoginActivity_(
      email,
      "",
      "FAILED",
      "Invalid User ID or Password.",
      ""
    );

    return {
      success: false,
      message:
        "Invalid User ID or Password."
    };
  }

  if (
    user.status
      .trim()
      .toLowerCase() !==
    "active"
  ) {

    logLoginActivity_(
      user.email,
      user.role,
      "FAILED",
      "User account is inactive.",
      ""
    );

    return {
      success: false,
      message:
        "Your account is inactive."
    };
  }

  const isRecoveryAdmin =
    user.email === RECOVERY_ADMIN_EMAIL;

  const usingTemporaryPassword =
    !!temporary &&
    user.email === RECOVERY_ADMIN_EMAIL &&
    password === TEMP_ADMIN_PASSWORD;

  if (
    !usingTemporaryPassword &&
    (!user.passwordHash || !user.passwordSalt)
  ) {

    logLoginActivity_(
      user.email,
      user.role,
      "FAILED",
      "Password is not configured.",
      ""
    );

    return {
      success: false,
      message:
        "Password is not configured for this user."
    };
  }

  if (!usingTemporaryPassword) {

    const enteredHash =
      hashPassword_(
        password,
        user.passwordSalt
      );

    if (
      enteredHash !==
      user.passwordHash
    ) {

      logLoginActivity_(
        user.email,
        user.role,
        "FAILED",
        "Invalid User ID or Password.",
        ""
      );

      return {
        success: false,
        message:
          "Invalid User ID or Password."
      };
    }
  }

  const sessionId =
    Utilities.getUuid();

  const loginTime =
    Date.now();

  const expires =
    loginTime +
    (
      SESSION_HOURS *
      60 *
      60 *
      1000
    );

  const sessionData = {

    email:
      user.email,

    name:
      user.name,

    role:
      user.role,

    loginTime:
      loginTime,

    expires:
      expires
  };

  PropertiesService
    .getScriptProperties()
    .setProperty(
      SESSION_PREFIX +
      sessionId,
      JSON.stringify(
        sessionData
      )
    );

  logLoginActivity_(
    user.email,
    user.role,
    "SUCCESS",
    usingTemporaryPassword
      ? "Login successful - temporary recovery password."
      : "Login successful.",
    sessionId
  );

  return {

    success: true,

    sessionId:
      sessionId,

    user: {

      email:
        user.email,

      name:
        user.name,

      role:
        user.role
    },

    expires:
      expires
  };
}

// =========================================================
// ADMIN PASSWORD RECOVERY - PRE-LOGIN
// =========================================================

function requestAdminRecovery(email) {

  email = String(email || "").trim().toLowerCase();

  if (email !== RECOVERY_ADMIN_EMAIL) {
    return {
      success: false,
      message: "Password recovery is available only for the authorized Admin account."
    };
  }

  const user = findUserByEmail_(email);

  if (!user) {
    return {
      success: false,
      message: "Authorized Admin account was not found in User Management."
    };
  }

  if (String(user.role || "").trim().toLowerCase() !== "admin") {
    return {
      success: false,
      message: "The authorized recovery account is not configured as Admin."
    };
  }

  if (String(user.status || "").trim().toLowerCase() !== "active") {
    return {
      success: false,
      message: "The authorized Admin account is inactive."
    };
  }

  const otp = String(Math.floor(100000 + Math.random() * 900000));
  const expires = Date.now() + (10 * 60 * 1000);

  PropertiesService.getScriptProperties().setProperty(
    "ADMIN_RECOVERY_OTP",
    JSON.stringify({ email: email, otp: otp, expires: expires })
  );

  MailApp.sendEmail({
    to: email,
    subject: "Material Request Portal - Admin Password Recovery OTP",
    body:
      "Your Admin password recovery OTP is: " + otp +
      "\n\nThis OTP is valid for 10 minutes.\n" +
      "After verification, temporary password 0000 will be activated for 5 minutes.\n" +
      "Your permanent password will not be changed."
  });

  return {
    success: true,
    message: "OTP sent to the authorized Admin email. OTP is valid for 10 minutes."
  };
}

function verifyAdminRecovery(email, otp) {

  email = String(email || "").trim().toLowerCase();
  otp = String(otp || "").trim();

  if (email !== RECOVERY_ADMIN_EMAIL) {
    return { success: false, message: "Unauthorized recovery account." };
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("ADMIN_RECOVERY_OTP");

  if (!raw) {
    return { success: false, message: "OTP not found or expired. Please request a new OTP." };
  }

  let data;
  try {
    data = JSON.parse(raw);
  } catch (e) {
    props.deleteProperty("ADMIN_RECOVERY_OTP");
    return { success: false, message: "Invalid recovery session. Please request a new OTP." };
  }

  if (
    data.email !== email ||
    data.otp !== otp ||
    !data.expires ||
    Date.now() > Number(data.expires)
  ) {
    if (data.expires && Date.now() > Number(data.expires)) {
      props.deleteProperty("ADMIN_RECOVERY_OTP");
    }
    return { success: false, message: "Invalid or expired OTP." };
  }

  const user = findUserByEmail_(email);

  if (!user || String(user.role || "").trim().toLowerCase() !== "admin") {
    props.deleteProperty("ADMIN_RECOVERY_OTP");
    return { success: false, message: "Authorized Admin account was not found." };
  }

  if (String(user.status || "").trim().toLowerCase() !== "active") {
    props.deleteProperty("ADMIN_RECOVERY_OTP");
    return { success: false, message: "Admin account is inactive." };
  }

  const tempExpires = Date.now() + (TEMP_ADMIN_MINUTES * 60 * 1000);

  props.setProperty(
    "ADMIN_TEMP_0000",
    JSON.stringify({
      email: RECOVERY_ADMIN_EMAIL,
      password: TEMP_ADMIN_PASSWORD,
      expires: tempExpires
    })
  );

  props.deleteProperty("ADMIN_RECOVERY_OTP");

  return {
    success: true,
    expires: tempExpires,
    message: "OTP verified. Temporary password 0000 is active for 5 minutes. Permanent password is unchanged."
  };
}

function getTemporaryAdminCredential_(email) {

  email = String(email || "").trim().toLowerCase();

  if (email !== RECOVERY_ADMIN_EMAIL) {
    return null;
  }

  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("ADMIN_TEMP_0000");

  if (!raw) {
    return null;
  }

  try {
    const data = JSON.parse(raw);

    if (
      data.email !== RECOVERY_ADMIN_EMAIL ||
      data.password !== TEMP_ADMIN_PASSWORD ||
      !data.expires ||
      Date.now() > Number(data.expires)
    ) {
      props.deleteProperty("ADMIN_TEMP_0000");
      return null;
    }

    return data;
  } catch (e) {
    props.deleteProperty("ADMIN_TEMP_0000");
    return null;
  }
}

function getTemporaryAdminPasswordStatus(email) {

  email = String(email || "").trim().toLowerCase();
  const credential = getTemporaryAdminCredential_(email);

  if (!credential) {
    return { active: false, remainingSeconds: 0 };
  }

  return {
    active: true,
    remainingSeconds: Math.max(0, Math.ceil((Number(credential.expires) - Date.now()) / 1000)),
    expires: Number(credential.expires)
  };
}

function findUserByEmail_(email) {

  email = String(email || "").trim().toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USER_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const rowEmail = String(values[i][0] || "").trim().toLowerCase();

    if (rowEmail === email) {
      return {
        email: rowEmail,
        name: String(values[i][1] || "").trim(),
        role: String(values[i][2] || "").trim(),
        passwordHash: String(values[i][3] || "").trim(),
        passwordSalt: String(values[i][4] || "").trim(),
        status: String(values[i][5] || "").trim()
      };
    }
  }

  return null;
}

// =========================================================
// ROLE ACCESS HELPERS
// =========================================================

function requireAdminSession_(sessionId) {
  const session = validateSession_(sessionId);

  if (!session) {
    throw new Error("Session expired. Please login again.");
  }

  if (
    String(session.role || "").trim().toLowerCase() !==
    "admin"
  ) {
    throw new Error("Access denied. Admin only.");
  }

  return session;
}

// =========================================================
// VALIDATE SESSION
// =========================================================

function validateSession_(
  sessionId
) {

  if (!sessionId) {
    return null;
  }

  const key =
    SESSION_PREFIX +
    sessionId;

  const props =
    PropertiesService
      .getScriptProperties();

  const sessionJson =
    props.getProperty(
      key
    );

  if (!sessionJson) {
    return null;
  }

  let session;

  try {

    session =
      JSON.parse(
        sessionJson
      );

  } catch (error) {

    props.deleteProperty(
      key
    );

    return null;
  }

  if (
    !session.expires ||
    Date.now() >
      Number(session.expires)
  ) {

    props.deleteProperty(
      key
    );

    return null;
  }

  return session;
}

// =========================================================
// GET CURRENT USER
// =========================================================

function getCurrentUser(
  sessionId
) {

  const session =
    validateSession_(
      sessionId
    );

  if (!session) {

    return {
      success: false,
      loggedIn: false
    };
  }

  return {

    success: true,
    loggedIn: true,

    user: {

      email:
        session.email,

      name:
        session.name,

      role:
        session.role,

      loginTime:
        session.loginTime
    }
  };
}

// =========================================================
// LOGOUT
// =========================================================

function logoutUser(
  sessionId
) {

  if (sessionId) {

    PropertiesService
      .getScriptProperties()
      .deleteProperty(
        SESSION_PREFIX +
        sessionId
      );
  }

  return {
    success: true
  };
}

// =========================================================
// SETUP LOGIN HISTORY SHEET
// =========================================================

function setupLoginHistorySheet_() {

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      LOGIN_HISTORY_SHEET_NAME
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        LOGIN_HISTORY_SHEET_NAME
      );
  }

  if (
    sheet.getLastRow() === 0
  ) {

    sheet.appendRow([
      "Timestamp",
      "User ID / Email",
      "Role",
      "Login Status",
      "Reason",
      "Session ID"
    ]);
  }

  return sheet;
}

// =========================================================
// LOG LOGIN ACTIVITY
// =========================================================

function logLoginActivity_(
  email,
  role,
  status,
  reason,
  sessionId
) {

  const sheet =
    setupLoginHistorySheet_();

  sheet.appendRow([
    new Date(),
    String(email || ""),
    String(role || ""),
    String(status || ""),
    String(reason || ""),
    String(sessionId || "")
  ]);

  maintainLoginHistory_(
    sheet
  );
}

// =========================================================
// KEEP LAST 10 SUCCESS + LAST 10 FAILED
// =========================================================

function maintainLoginHistory_(
  sheet
) {

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow <= 1
  ) {
    return;
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  const successRows = [];
  const failedRows = [];

  values.forEach(function(row) {

    const status =
      String(
        row[3] || ""
      )
        .trim()
        .toUpperCase();

    if (
      status ===
      "SUCCESS"
    ) {
      successRows.push(
        row
      );
    }

    if (
      status ===
      "FAILED"
    ) {
      failedRows.push(
        row
      );
    }
  });

  successRows.reverse();
  failedRows.reverse();

  const finalRows =
    successRows
      .slice(0, 10)
      .concat(
        failedRows
          .slice(0, 10)
      );

  sheet
    .getRange(
      2,
      1,
      lastRow - 1,
      6
    )
    .clearContent();

  if (
    finalRows.length > 0
  ) {

    sheet
      .getRange(
        2,
        1,
        finalRows.length,
        6
      )
      .setValues(
        finalRows
      );
  }
}

// =========================================================
// GET LOGIN HISTORY
// ADMIN ONLY
// =========================================================

function getLoginHistory(
  sessionId
) {

  const session =
    validateSession_(
      sessionId
    );

  if (!session) {

    throw new Error(
      "Session expired. Please login again."
    );
  }

  if (
    String(session.role || "")
      .trim()
      .toLowerCase() !==
    "admin"
  ) {

    throw new Error(
      "Access denied. Admin only."
    );
  }

  const sheet =
    setupLoginHistorySheet_();

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {

    return {
      success: true,
      records: []
    };
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  const timezone =
    Session.getScriptTimeZone() ||
    "Asia/Kolkata";

  const records =
    values
      .reverse()
      .map(function(row) {

        return {

          timestamp:
            row[0] instanceof Date
              ? Utilities.formatDate(
                  row[0],
                  timezone,
                  "dd-MM-yyyy HH:mm:ss"
                )
              : String(
                  row[0] || ""
                ),

          email:
            String(
              row[1] || ""
            ),

          role:
            String(
              row[2] || ""
            ),

          status:
            String(
              row[3] || ""
            ),

          reason:
            String(
              row[4] || ""
            ),

          sessionId:
            String(
              row[5] || ""
            )
        };
      });

  return {
    success: true,
    records:
      records
  };
}

// =========================================================
// OPTIONAL MANUAL RECOVERY ACTIVATION
// Run from Apps Script editor only if OTP email cannot be used.
// It does NOT change the permanent password.
// =========================================================

function setRecoveryAdminPasswordTo0000_() {

  const user = findUserByEmail_(RECOVERY_ADMIN_EMAIL);

  if (!user) {
    throw new Error("Authorized Admin account was not found in User Management.");
  }

  if (String(user.role || "").trim().toLowerCase() !== "admin") {
    throw new Error("Recovery account must have Admin role.");
  }

  if (String(user.status || "").trim().toLowerCase() !== "active") {
    throw new Error("Recovery Admin account must be Active.");
  }

  const expires = Date.now() + (TEMP_ADMIN_MINUTES * 60 * 1000);

  PropertiesService.getScriptProperties().setProperty(
    "ADMIN_TEMP_0000",
    JSON.stringify({
      email: RECOVERY_ADMIN_EMAIL,
      password: TEMP_ADMIN_PASSWORD,
      expires: expires
    })
  );

  return "Temporary password 0000 activated for 5 minutes. Permanent password unchanged.";
}

// =========================================================
// MANUAL PASSWORD HASH GENERATOR
// =========================================================

function generatePasswordHash(
  password
) {

  if (!password) {
    throw new Error(
      "Password required."
    );
  }

  const salt =
    Utilities.getUuid();

  const hash =
    hashPassword_(
      password,
      salt
    );

  return {

    hash:
      hash,

    salt:
      salt
  };
}

// =========================================================
// USER MANAGEMENT
// =========================================================

function isValidEmail_(
  email
) {

  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/
    .test(
      String(
        email || ""
      ).trim()
    );
}

// =========================================================
// GET USERS
// =========================================================

function getUserManagement(sessionId) {
  requireAdminSession_(sessionId);

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      USER_SHEET_NAME
    );

  if (!sheet) {
    return [];
  }

  const lastRow =
    sheet.getLastRow();

  if (
    lastRow < 2
  ) {
    return [];
  }

  const values =
    sheet
      .getRange(
        2,
        1,
        lastRow - 1,
        6
      )
      .getValues();

  return values
    .filter(function(row) {

      return String(
        row[0] || ""
      )
        .trim() !== "";

    })
    .map(function(row) {

      return {

        email:
          String(
            row[0] || ""
          ).trim(),

        name:
          String(
            row[1] || ""
          ).trim(),

        role:
          String(
            row[2] || ""
          ).trim(),

        active:
          String(
            row[5] || ""
          )
            .trim()
            .toLowerCase() ===
          "active",

        status:
          String(
            row[5] || ""
          ).trim()
      };
    });
}

// =========================================================
// SAVE USERS
// =========================================================

function saveUserManagement(
  sessionId,
  users
) {
  requireAdminSession_(sessionId);

  if (
    !Array.isArray(users) ||
    users.length === 0
  ) {

    throw new Error(
      "Please add at least one user."
    );
  }

  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();

  const sheet =
    ss.getSheetByName(
      USER_SHEET_NAME
    );

  if (!sheet) {

    throw new Error(
      "User Management sheet not found."
    );
  }

  const rows = [];

  users.forEach(function(
    user,
    index
  ) {

    const email =
      String(
        user.email || ""
      )
        .trim()
        .toLowerCase();

    const name =
      String(
        user.name || ""
      ).trim();

    const role =
      String(
        user.role || ""
      ).trim();

    const password =
      String(
        user.password || ""
      );

    const status =
      String(
        user.status || ""
      )
        .trim()
        .toLowerCase() ===
      "active"
        ? "Active"
        : "Inactive";

    if (!email) {

      throw new Error(
        "User " +
        (index + 1) +
        ": Email is required."
      );
    }

    if (!isValidEmail_(email)) {

      throw new Error(
        "User " +
        (index + 1) +
        ": Invalid email address."
      );
    }

    if (!name) {

      throw new Error(
        "User " +
        (index + 1) +
        ": Name is required."
      );
    }

    if (
      role !== "Admin" &&
      role !== "Operator"
    ) {

      throw new Error(
        "User " +
        (index + 1) +
        ": Invalid role."
      );
    }

    if (!password) {

      throw new Error(
        "User " +
        (index + 1) +
        ": Password is required."
      );
    }

    if (
      password.length < 6
    ) {

      throw new Error(
        "User " +
        (index + 1) +
        ": Password must contain at least 6 characters."
      );
    }

    const salt =
      Utilities.getUuid();

    const hash =
      hashPassword_(
        password,
        salt
      );

    rows.push([
      email,
      name,
      role,
      hash,
      salt,
      status
    ]);
  });

  const startRow =
    Math.max(
      sheet.getLastRow() + 1,
      2
    );

  sheet
    .getRange(
      startRow,
      1,
      rows.length,
      6
    )
    .setValues(rows);

  return {

    success: true,

    message:
      rows.length +
      " user(s) saved successfully."
  };
}

// =========================================================
// PROBLEM OBSERVATION & TRACKING
// =========================================================

const PROBLEM_OBSERVATION_SHEET_NAME = "Problem Observation Tracking";
const PROBLEM_MEDIA_FOLDER_PROPERTY = "PROBLEM_OBSERVATION_MEDIA_FOLDER_ID";
const PROBLEM_MEDIA_MAX_BYTES = 50 * 1024 * 1024;

function ensureProblemObservationSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(PROBLEM_OBSERVATION_SHEET_NAME);
  const headers = [
    "Timestamp", "Observation ID", "Email Address", "Name", "Role",
    "Station / Machine ID", "Ticket ID", "Problem Category", "Problem Title",
    "Observation", "Priority", "Status", "Admin Remark", "Media File Name",
    "Media URL", "Media Type", "Media Size (MB)", "Updated By", "Updated At"
  ];

  if (!sheet) sheet = ss.insertSheet(PROBLEM_OBSERVATION_SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    const width = Math.max(sheet.getLastColumn(), headers.length);
    const current = sheet.getRange(1, 1, 1, width).getValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (!String(current[i] || "").trim()) {
        sheet.getRange(1, i + 1).setValue(headers[i]);
      }
    }
  }

  return sheet;
}

function getProblemMediaFolder_() {
  const props = PropertiesService.getScriptProperties();
  const savedId = String(props.getProperty(PROBLEM_MEDIA_FOLDER_PROPERTY) || "").trim();

  if (savedId) {
    try {
      return DriveApp.getFolderById(savedId);
    } catch (e) {
      props.deleteProperty(PROBLEM_MEDIA_FOLDER_PROPERTY);
    }
  }

  const folder = DriveApp.createFolder("Problem Observation Media");
  props.setProperty(PROBLEM_MEDIA_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function sanitizeFileName_(name) {
  return String(name || "media")
    .replace(/[\\/:*?"<>|#%{}]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120) || "media";
}

function saveProblemObservation(formObject) {
  requirePortalPermission_(String(formObject && formObject.sessionId || ""), "maintenance", "create");

  try {
    if (!formObject) return { success: false, message: "Observation form data is required." };

    const session = validateSession_(String(formObject.sessionId || "").trim());
    if (!session) return { success: false, message: "Session expired. Please login again." };

    const stationId = String(formObject.stationId || "").trim();
    // Ticket ID is generated automatically by the system for every observation.
    const category = String(formObject.category || "").trim();
    const title = String(formObject.title || "").trim();
    const observation = String(formObject.observation || "").trim();
    const priority = String(formObject.priority || "Medium").trim();

    if (!stationId || !category || !title || !observation) {
      return { success: false, message: "Station, Category, Problem Title and Observation are required." };
    }

    /*
     * Accept the new mobile-safe payload (base64) and also keep backward
     * compatibility with the old HTML-form Blob payload.
     */
    let media;
    let contentType = "";
    let bytes;

    if (formObject.mediaBase64) {
      contentType = String(formObject.mediaType || "").toLowerCase();

      if (contentType.indexOf("image/") !== 0 && contentType.indexOf("video/") !== 0) {
        return { success: false, message: "Only image/video files are allowed." };
      }

      bytes = Utilities.base64Decode(String(formObject.mediaBase64));
      if (!bytes || !bytes.length) {
        return { success: false, message: "Selected media file is empty or could not be read." };
      }

      media = Utilities.newBlob(
        bytes,
        contentType,
        sanitizeFileName_(formObject.mediaName || "problem-media")
      );

    } else {
      media = formObject.mediaFile;

      if (!media || typeof media.getBytes !== "function") {
        return { success: false, message: "Realtime photo/video is mandatory." };
      }

      contentType = String(media.getContentType() || "").toLowerCase();

      if (contentType.indexOf("image/") !== 0 && contentType.indexOf("video/") !== 0) {
        return { success: false, message: "Only image/video files are allowed." };
      }

      bytes = media.getBytes();
    }

    if (bytes.length > PROBLEM_MEDIA_MAX_BYTES) {
      return { success: false, message: "Maximum media size is 50 MB." };
    }

    const sheet = ensureProblemObservationSheet_();
    const timezone = Session.getScriptTimeZone() || "Asia/Kolkata";
    const now = new Date();
    // Ticket ID is generated automatically and includes the Station / Machine ID.
    // Example: OBS-STN01-20260919-172530-A7K92F
    const stationCode = String(stationId)
      .trim()
      .toUpperCase()
      .replace(/[^A-Z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 24) || "STATION";

    const uniqueCode = Utilities.getUuid().replace(/-/g, "").slice(0, 6).toUpperCase();
    const generatedTicketId =
      "OBS-" +
      stationCode +
      "-" +
      Utilities.formatDate(now, timezone, "yyyyMMdd-HHmmss") +
      "-" +
      uniqueCode;
    const observationId = generatedTicketId;

    const folder = getProblemMediaFolder_();
    const originalName = sanitizeFileName_(formObject.mediaName || media.getName() || "problem-media");
    const fileName = observationId + " - " + originalName;
    const blob = media.copyBlob().setName(fileName);
    const file = folder.createFile(blob);

    let mediaUrl = file.getUrl();
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      mediaUrl = file.getUrl();
    } catch (shareError) {
      Logger.log("Problem observation media sharing could not be changed: " + shareError);
    }

    const sizeMb = bytes.length / (1024 * 1024);

    sheet.appendRow([
      now,
      observationId,
      String(session.email || "").trim(),
      String(session.name || "").trim(),
      String(session.role || "").trim(),
      stationId,
      generatedTicketId,
      category,
      title,
      observation,
      priority,
      "Open",
      "",
      fileName,
      mediaUrl,
      contentType,
      Number(sizeMb.toFixed(2)),
      "",
      ""
    ]);

    // Route observation notifications through activity-wise settings.
    try {
      sendConfiguredPortalNotice_("problemObservation","Problem Observation - " + generatedTicketId,
        "A new problem observation has been raised. Ticket ID: " + generatedTicketId +
        ". Category: " + category + ". Title: " + title + ". Priority: " + priority +
        ". Raised by: " + String(session.name || session.email) + ". Details: " + observation +
        ". Media: " + mediaUrl);
    } catch (mailError) { Logger.log("Problem observation notification failed: " + mailError.message); }

    return {
      success: true,
      ticketId: generatedTicketId,
      observationId: observationId,
      mediaUrl: mediaUrl,
      message: "Problem observation raised successfully. Ticket ID: " + generatedTicketId
    };
  } catch (error) {
    return {
      success: false,
      message: error && error.message ? error.message : String(error)
    };
  }
}

function buildProblemObservationMailHtml_(ticketId, session, stationId, category, title, observation, priority, mediaUrl) {
  return
    "<div style='font-family:Calibri,Arial,sans-serif;font-size:14px;'>" +
    "<h2 style='color:#1F4E79;'>Problem Observation Raised</h2>" +
    "<table style='border-collapse:collapse;'>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Ticket ID</td><td style='padding:6px 12px;'>" + escapeHtml(ticketId) + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Raised By</td><td style='padding:6px 12px;'>" + escapeHtml(session.name || "") + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Email</td><td style='padding:6px 12px;'>" + escapeHtml(session.email || "") + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Station / Machine ID</td><td style='padding:6px 12px;'>" + escapeHtml(stationId) + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Problem Category</td><td style='padding:6px 12px;'>" + escapeHtml(category) + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Problem Title</td><td style='padding:6px 12px;'>" + escapeHtml(title) + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Priority</td><td style='padding:6px 12px;'>" + escapeHtml(priority) + "</td></tr>" +
    "<tr><td style='padding:6px 12px;font-weight:bold;'>Observation</td><td style='padding:6px 12px;white-space:pre-wrap;'>" + escapeHtml(observation) + "</td></tr>" +
    "</table>" +
    "<p><a href='" + escapeHtml(mediaUrl) + "'>View Observation Media</a></p>" +
    "</div>";
}


function getProblemObservations(sessionId) {
  const session = validateSession_(sessionId);
  if (!session) return { success: false, message: "Session expired. Please login again.", records: [] };

  const sheet = ensureProblemObservationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, records: [] };

  const timezone = Session.getScriptTimeZone() || "Asia/Kolkata";
  const values = sheet.getRange(2, 1, lastRow - 1, 19).getValues();
  const records = [];

  values.forEach(function(row) {
    if (!row[1]) return;
    records.push({
      timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], timezone, "dd-MM-yyyy HH:mm") : String(row[0] || ""),
      observationId: String(row[1] || ""),
      email: String(row[2] || ""),
      name: String(row[3] || ""),
      role: String(row[4] || ""),
      stationId: String(row[5] || ""),
      ticketId: String(row[6] || ""),
      category: String(row[7] || ""),
      title: String(row[8] || ""),
      observation: String(row[9] || ""),
      priority: String(row[10] || ""),
      status: String(row[11] || "Open"),
      adminRemark: String(row[12] || ""),
      mediaFileName: String(row[13] || ""),
      mediaUrl: String(row[14] || ""),
      mediaType: String(row[15] || ""),
      mediaSizeMb: String(row[16] || ""),
      updatedBy: String(row[17] || ""),
      updatedAt: row[18] instanceof Date ? Utilities.formatDate(row[18], timezone, "dd-MM-yyyy HH:mm") : String(row[18] || "")
    });
  });

  records.reverse();
  return { success: true, records: records.slice(0, 200) };
}

function trashProblemObservationMedia_(mediaUrl) {
  const url = String(mediaUrl || "").trim();
  if (!url) return;
  try {
    const match = url.match(/[-\\w]{25,}/);
    if (match && match[0]) {
      DriveApp.getFileById(match[0]).setTrashed(true);
    }
  } catch (e) {
    Logger.log("Problem observation media could not be moved to trash: " + e);
  }
}

function updateProblemObservation(sessionId, observationId, status, remark) {
  const session = validateSession_(sessionId);
  if (!session) return { success: false, message: "Session expired. Please login again." };

  if (!hasPortalPermission_(session, "maintenance", "edit")) return { success:false, message:"Access denied: maintenance edit permission is not enabled." };

  observationId = String(observationId || "").trim();
  status = String(status || "").trim();
  remark = String(remark || "").trim();

  const allowed = ["Open", "In Progress", "Resolved", "Closed", "Rejected"];
  if (!observationId) return { success: false, message: "Observation ID is required." };
  if (allowed.indexOf(status) === -1) return { success: false, message: "Invalid observation status." };

  const sheet = ensureProblemObservationSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No problem observations found." };

  const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  let matchedRow = -1;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || "").trim() === observationId) {
      matchedRow = i + 2;
      break;
    }
  }

  if (matchedRow === -1) return { success: false, message: "Observation not found." };

  const currentStatus = String(sheet.getRange(matchedRow, 12).getDisplayValue() || "").trim();
  if (currentStatus.toLowerCase() === "closed") {
    return { success: false, message: "Closed observation cannot be updated or reopened." };
  }

  const now = new Date();

  if (status === "Closed") {
    const mediaUrl = String(sheet.getRange(matchedRow, 15).getDisplayValue() || "").trim();
    trashProblemObservationMedia_(mediaUrl);
    sheet.getRange(matchedRow, 14, 1, 3).clearContent();
  }

  sheet.getRange(matchedRow, 12).setValue(status);
  sheet.getRange(matchedRow, 13).setValue(remark);
  sheet.getRange(matchedRow, 18).setValue(String(session.email || "").trim());
  sheet.getRange(matchedRow, 19).setValue(now);
  SpreadsheetApp.flush();
  try { sendConfiguredPortalNotice_("problemObservation","Problem Observation Status Update - "+observationId,"Problem observation "+observationId+" status changed to "+status+" by "+String(session.name||session.email)+". Remark: "+remark); } catch(mailError) { Logger.log("Problem observation status notification failed: "+mailError.message); }

  return {
    success: true,
    message: status === "Closed"
      ? "Problem observation closed successfully. Uploaded media was moved to trash and the status is now locked."
      : "Problem observation updated successfully."
  };
}

// =========================================================
// LEGACY TOOL NORMALIZER
// Kept so existing functionality does not break.
// =========================================================


// =========================================================
// USER MANAGEMENT - ADMIN ADD/REMOVE HELPERS
// =========================================================

function addUserManagement(sessionId, user) {
  const session = validateSession_(sessionId);

  if (!session || String(session.role || "").trim().toLowerCase() !== "admin") {
    return { success: false, message: "Admin access required." };
  }

  return saveUserManagement(sessionId, [user]);
}

function removeUserManagement(sessionId, email) {
  const session = validateSession_(sessionId);

  if (!session || String(session.role || "").trim().toLowerCase() !== "admin") {
    return { success: false, message: "Admin access required." };
  }

  email = String(email || "").trim().toLowerCase();

  if (!email) return { success: false, message: "User email is required." };

  if (email === RECOVERY_ADMIN_EMAIL.toLowerCase()) {
    return { success: false, message: "The authorized recovery Admin cannot be removed." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(USER_SHEET_NAME);
  if (!sheet) return { success: false, message: "User Management sheet not found." };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: "No saved users found." };

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (let i = values.length - 1; i >= 0; i--) {
    const rowEmail = String(values[i][0] || "").trim().toLowerCase();

    if (rowEmail === email) {
      sheet.deleteRow(i + 2);
      return { success: true, message: "User removed successfully." };
    }
  }

  return { success: false, message: "User not found." };
}

function normalizeToolNo_(value) {
  return String(
    value == null ? "" : value
  )
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}



// =========================================================
// TICKET MANAGER / BREAKDOWN MANAGEMENT
// =========================================================


function ensureBreakdownSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BREAKDOWN_SHEET_NAME);

  // FIXED BREAKDOWN SCHEMA - A:V. Never append/rename headers per submission.
  // Every required column is created once, in a fixed position.
  const headers = [
    "Timestamp",                         // A
    "Email Address",                     // B
    "Name",                              // C
    "Ticket ID",                         // D
    "Breakdown Type",                    // E
    "Start Time",                        // F
    "End Time",                          // G
    "Total Taken Time (hrs)",            // H
    "Downtime % of 24hrs",               // I
    "Issue Owner",                       // J
    "TPT per Hour",                      // K
    "TPT Loss",                          // L
    "Actual TPT During Breakdown",       // M
    "Description",                       // N
    "Resolution",                        // O
    "Status",                            // P
    "Spare Used",                        // Q
    "Spare Consumed Qty",                // R
    "Spare Item IDs",                    // S
    "Case No.",                          // T
    "Remark",                            // U
    "Mobile"                             // V
  ];

  if (!sheet) {
    sheet = ss.insertSheet(BREAKDOWN_SHEET_NAME);
  }

  // Lock the schema every time. This does NOT add a new header row and does
  // NOT write anything into data rows. Blank fields remain blank.
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);

  // Keep exactly the required A:V schema visible. Do not create additional
  // unnamed/duplicate columns during submissions.
  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  return sheet;
}


function saveBreakdownEntry(sessionId,data){
  const session=validateSession_(sessionId);if(!session)return{success:false,message:"Session expired. Please login again."};if(String(session.role||"").trim().toLowerCase()!=="admin")return{success:false,message:"Admin access required."};if(!data)return{success:false,message:"Breakdown data is required."};
  const ticketId=String(data.ticketId||"").trim(),type=String(data.breakdownType||"").trim(),sv=String(data.startTime||"").trim(),ev=String(data.endTime||"").trim(),owner=String(data.issueOwner||"").trim(),spareUsed=String(data.spareUsed||"").trim(),caseNo=String(data.caseNo||"").trim(),actualTpt=Number(data.actualTpt!=null?data.actualTpt:(data.tptPerHour||0));
  if(!ticketId)return{success:false,message:"Ticket ID is required."};if(["Planned Downtime","Partial Breakdown","Complete Breakdown"].indexOf(type)<0)return{success:false,message:"Invalid Breakdown Type."};if(!sv||!ev)return{success:false,message:"Start Time and End Time are required."};const start=new Date(sv),end=new Date(ev);if(isNaN(start)||isNaN(end)||end<=start)return{success:false,message:"Invalid start/end time."};if(type!=="Planned Downtime"&&!owner)return{success:false,message:"Issue Owner is required for Breakdown."};if(actualTpt<0||!isFinite(actualTpt))return{success:false,message:"Actual TPT During Breakdown must be zero or greater."};if(spareUsed!=="Yes"&&spareUsed!=="No")return{success:false,message:"Please select whether spare was used."};const qty=Number(data.spareConsumedQty||0),ids=String(data.spareItemIds||"").trim();if(spareUsed==="Yes"&&(qty<=0||!ids))return{success:false,message:"Spare quantity and spare item ID(s) are required when Spare Used is Yes."};
  const hours=(end-start)/3600000,percent=(hours/24)*100,expected=hours*12000,loss=Math.max(0,expected-actualTpt),sheet=ensureBreakdownSheet_();sheet.getRange(sheet.getLastRow()+1,1,1,22).setValues([[new Date(),session.email||"",session.name||"",ticketId,type,start,end,hours,percent,owner,12000,loss,actualTpt,"","","Closed",spareUsed,spareUsed==="Yes"?qty:0,spareUsed==="Yes"?ids:"",caseNo,"",session.mobile||""]]);SpreadsheetApp.flush();return{success:true,message:type+" saved successfully."};
}

/**
 * Returns breakdown percentage for a selected date range.
 for a selected date range.
 *
 * Formula:
 *   Breakdown % = total breakdown hours in the selected range
 *                 / (22 hours x number of calendar days) x 100
 *
 * Only the portion of a breakdown that overlaps the selected range is counted.
 */
function getBreakdownDateRangeSummary(sessionId, fromDate, toDate) {
  var s = validateSession_(sessionId);
  if (!s) return {success:false, message:"Session expired. Please login again."};
  if (String(s.role || "").toLowerCase() !== "admin") return {success:false, message:"Admin access required."};

  var tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  var range = parseReportRange_(fromDate, toDate), rs = range.start, re = range.end;
  if (re <= rs) return {success:false, message:"To Date must be on or after From Date."};

  var breakdownHours = 0, plannedHours = 0, tptLoss = 0, records = 0, spare = 0, scheduledHours = 0;
  var sh = ensureBreakdownSheet_();

  if (sh.getLastRow() >= 2) {
    sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(21, sh.getLastColumn())).getValues().forEach(function(r) {
      var st = parseReportDate_(r[5]), en = parseReportDate_(r[6]);
      if (!st || !en || en <= st) return;
      var os = Math.max(st.getTime(), rs.getTime()), oe = Math.min(en.getTime(), re.getTime());
      if (oe <= os) return;
      var h = (oe - os) / 3600000, type = String(r[4] || "").trim();
      if (type.toLowerCase() === "planned downtime") {
        plannedHours += h;
      } else {
        breakdownHours += h;
        records++;
        var storedH = Number(r[7] || 0), storedLoss = Number(r[11] || 0);
        tptLoss += storedH > 0 ? Math.max(0, storedLoss / storedH * h) : 0;
        spare += Number(r[17] || 0);
      }
    });
  }

  var sd = ensureScheduledDowntimeSheet_();
  if (sd.getLastRow() >= 2) {
    sd.getRange(2, 1, sd.getLastRow() - 1, Math.max(8, sd.getLastColumn())).getValues().forEach(function(r) {
      var st = parseReportDate_(r[2]), en = parseReportDate_(r[3]);
      if (!st || !en || en <= st) return;
      var os = Math.max(st.getTime(), rs.getTime()), oe = Math.min(en.getTime(), re.getTime());
      if (oe > os) scheduledHours += (oe - os) / 3600000;
    });
  }

  var expectedBreakdownTpt = breakdownHours * 12000;
  var lossPercent = expectedBreakdownTpt > 0 ? (tptLoss / expectedBreakdownTpt * 100) : 0;

  return {
    success:true,
    fromDate:Utilities.formatDate(rs,tz,"dd-MM-yyyy"),
    toDate:Utilities.formatDate(new Date(re.getTime()-1),tz,"dd-MM-yyyy"),
    days:Math.round((re-rs)/86400000),
    recordCount:records,
    totalBreakdownHours:+breakdownHours.toFixed(2),
    plannedDowntimeHours:+plannedHours.toFixed(2),
    scheduledDowntimeHours:+scheduledHours.toFixed(2),
    availableHours:+Math.max(0,Math.round((re-rs)/86400000)*24-scheduledHours-plannedHours).toFixed(2),
    breakdownPercent:+lossPercent.toFixed(2),
    tptLoss:+tptLoss.toFixed(2),
    spareConsumedQty:+spare.toFixed(2)
  };
}

function getBreakdownRecords(sessionId,filters){
  const s=validateSession_(sessionId);if(!s)return{success:false,message:"Session expired. Please login again.",records:[]};if(String(s.role||"").toLowerCase()!=="admin")return{success:false,message:"Admin access required.",records:[]};filters=filters||{};const tf=String(filters.ticketId||"").toLowerCase(),sf=String(filters.status||"").toLowerCase();let rs=null,re=null;if(filters.fromDate&&filters.toDate){const a=String(filters.fromDate).split("-"),b=String(filters.toDate).split("-");if(a.length===3&&b.length===3){rs=new Date(+a[0],+a[1]-1,+a[2]);re=new Date(+b[0],+b[1]-1,+b[2]+1);}}
  const sh=ensureBreakdownSheet_();if(sh.getLastRow()<2)return{success:true,records:[]};const tz=Session.getScriptTimeZone()||"Asia/Kolkata",out=[];sh.getRange(2,1,sh.getLastRow()-1,Math.max(21,sh.getLastColumn())).getValues().forEach(function(r){if(!r[0]&&!r[3])return;const ticket=String(r[3]||"").trim(),status=String(r[15]||"").trim(),st=r[5] instanceof Date?r[5]:null,en=r[6] instanceof Date?r[6]:null;if(tf&&ticket.toLowerCase().indexOf(tf)<0)return;if(sf&&status.toLowerCase()!==sf)return;if(rs&&re&&st&&en&&(en<=rs||st>=re))return;const f=v=>v instanceof Date?Utilities.formatDate(v,tz,"dd-MM-yyyy HH:mm"):String(v||"");out.push({timestamp:f(r[0]),email:String(r[1]||""),name:String(r[2]||""),ticketId:ticket,breakdownType:String(r[4]||""),startTime:f(r[5]),endTime:f(r[6]),totalTakenTime:Number(r[7]||0).toFixed(2),downtimePercent:Number(r[8]||0).toFixed(2)+"%",issueOwner:String(r[9]||""),tptPerHour:"₹"+Number(r[10]||0).toLocaleString("en-IN"),tptImpact:"₹"+Number(r[11]||0).toLocaleString("en-IN",{minimumFractionDigits:2}),tptAtSubmission:"₹"+Number(r[12]||0).toLocaleString("en-IN",{minimumFractionDigits:2}),status:status,spareUsed:String(r[16]||""),spareConsumedQty:String(r[17]||""),spareItemIds:String(r[18]||""),caseNo:String(r[19]||"")});});out.reverse();return{success:true,records:out.slice(0,500)};
}

function getReportDateValue_(v){
  if(v instanceof Date && !isNaN(v.getTime())) return v;
  if(v===null || v===undefined || v==="") return null;
  if(typeof v === "number") { var dn=new Date(v); if(!isNaN(dn.getTime())) return dn; }
  var s=String(v).trim(); if(!s) return null;
  var d=new Date(s); if(!isNaN(d.getTime())) return d;
  var m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
  if(m){ var d2=new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0),+(m[6]||0)); if(!isNaN(d2.getTime())) return d2; }
  return null;
}

function getBreakdownOwnerSummary_(sessionId,fromDate,toDate){
  const s=validateSession_(sessionId);
  if(!s || String(s.role||"").toLowerCase()!=="admin") return [];
  const sh=ensureBreakdownSheet_();
  if(sh.getLastRow()<2) return [];
  const a=String(fromDate||"").split("-"),b=String(toDate||"").split("-");
  if(a.length!==3||b.length!==3) return [];
  const rs=new Date(+a[0],+a[1]-1,+a[2]),re=new Date(+b[0],+b[1]-1,+b[2]+1);
  if(re<=rs) return [];
  const map={};
  sh.getRange(2,1,sh.getLastRow()-1,Math.max(21,sh.getLastColumn())).getValues().forEach(function(r){
    const st=getReportDateValue_(r[5]),en=getReportDateValue_(r[6]);
    if(!st||!en||en<=st) return;
    const os=new Date(Math.max(st.getTime(),rs.getTime())),oe=new Date(Math.min(en.getTime(),re.getTime()));
    if(oe<=os) return;
    const h=(oe-os)/3600000;
    let owner=String(r[9]||"").trim().replace(/\s+/g," ");
    if(!owner) owner="Unspecified";
    const storedH=Number(r[7]||0),storedLoss=Number(r[11]||0);
    const loss=storedH>0?Math.max(0,storedLoss/storedH*h):Math.max(0,storedLoss);
    if(!map[owner]) map[owner]={hours:0,loss:0,records:0};
    map[owner].hours+=h; map[owner].loss+=loss; map[owner].records++;
  });
  return Object.keys(map).sort().map(function(k){
    const expected=map[k].hours*12000;
    return {owner:k,hours:+map[k].hours.toFixed(2),percent:expected>0?+(map[k].loss/expected*100).toFixed(2):0,tptLoss:+map[k].loss.toFixed(2),records:map[k].records};
  });
}

function ensureReceivedSpareSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sh=ss.getSheetByName(RECEIVED_SPARE_SHEET_NAME);
  const headers=[
    "Timestamp","Received Date","Received Time","Received Against","Request / Ticket ID",
    "AWB No.","Courier / Transporter","Challan No.","Last Date","Item ID","Item Name","Quantity","Unit",
    "Received By / Operator","Handover To","Storage Location","Package / Material Condition",
    "Remarks","Submitted By Email","Submitted By Name"
  ];
  if(!sh)sh=ss.insertSheet(RECEIVED_SPARE_SHEET_NAME);
  if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);
  else{
    const current=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach(function(v,i){if(!String(current[i]||"").trim())sh.getRange(1,i+1).setValue(v);});
  }
  sh.setFrozenRows(1);
  return sh;
}

function saveReceivedSpare(sessionId,data){
  requirePortalPermission_(sessionId, "spareParts", "create");

  const s=validateSession_(sessionId);
  if(!s)return{success:false,message:"Session expired. Please login again."};
  data=data||{};
  const receivedDate=String(data.receivedDate||"").trim();
  const receivedTime=String(data.receivedTime||"").trim();
  const receivedAgainst=String(data.receivedAgainst||"").trim();
  const ticketId=String(data.ticketId||"").trim();
  const awbNo=String(data.awbNo||"").trim();
  const courier=String(data.courier||"").trim();
  const challanNo=String(data.challanNo||"").trim();
  const lastDate=String(data.lastDate||"").trim();
  const receivedBy=String(data.receivedBy||"").trim();
  const handoverTo=String(data.handoverTo||"").trim();
  const location=String(data.location||"").trim();
  const condition=String(data.condition||"").trim();
  const remarks=String(data.remarks||"").trim();

  let items=Array.isArray(data.items)?data.items:[];
  // Backward compatibility with the previous single-item payload.
  if(!items.length && (data.itemId||data.itemName||data.qty)){
    items=[{itemId:data.itemId,itemName:data.itemName,qty:data.qty,unit:data.unit}];
  }
  items=items.map(function(x){return{
    itemId:String(x&&x.itemId||"").trim(),
    itemName:String(x&&x.itemName||"").trim(),
    qty:Number(x&&x.qty||0),
    unit:String(x&&x.unit||"").trim()
  };});

  if(!receivedDate||!receivedTime||!receivedAgainst||!awbNo||!receivedBy||!handoverTo||!location){
    return{success:false,message:"All required Received Spare details must be filled."};
  }
  if(["Material Request","Without Request"].indexOf(receivedAgainst)<0){
    return{success:false,message:"Invalid Received Against selection."};
  }
  if(receivedAgainst==="Material Request"&&!ticketId){
    return{success:false,message:"Request / Ticket ID is required for Material Request."};
  }
  if(!items.length||items.some(function(x){return !x.itemId||!x.itemName||x.qty<=0;})){
    return{success:false,message:"Please enter valid Item ID, Item Name and Quantity for every item."};
  }

  const dt=new Date(receivedDate+"T"+receivedTime);
  if(isNaN(dt.getTime()))return{success:false,message:"Invalid Received Date / Time."};

  const sh=ensureReceivedSpareSheet_();
  const rows=items.map(function(item){return[
    new Date(),new Date(receivedDate+"T00:00:00"),dt,receivedAgainst,ticketId,awbNo,courier,challanNo,lastDate,
    item.itemId,item.itemName,item.qty,item.unit,receivedBy,handoverTo,location,condition,remarks,
    String(s.email||"").trim(),String(s.name||"").trim()
  ];});
  sh.getRange(sh.getLastRow()+1,1,rows.length,20).setValues(rows);
  SpreadsheetApp.flush();
  try { sendConfiguredPortalNotice_("receivedSpare","Received Spare Update","Received spare entry submitted by "+String(s.name||s.email)+". Items: "+items.map(function(x){return x.itemName+" ("+x.qty+")";}).join(", ")+". Received against: "+receivedAgainst+". Ticket: "+ticketId+"."); } catch(mailError) { Logger.log("Received Spare notification failed: "+mailError.message); }
  return{success:true,message:items.length+" received spare item(s) submitted successfully."};
}

function getReceivedSpareRecords(sessionId,fromDate,toDate){
  const s=validateSession_(sessionId);
  if(!s)return{success:false,message:"Session expired. Please login again.",records:[]};
  const sh=ensureReceivedSpareSheet_();
  if(sh.getLastRow()<2)return{success:true,records:[]};
  let rs=null,re=null;
  if(fromDate&&toDate){
    const a=String(fromDate).split("-"),b=String(toDate).split("-");
    if(a.length===3&&b.length===3){
      rs=new Date(+a[0],+a[1]-1,+a[2]);
      re=new Date(+b[0],+b[1]-1,+b[2]+1);
    }
  }
  const tz=Session.getScriptTimeZone()||"Asia/Kolkata";
  const out=[];
  sh.getRange(2,1,sh.getLastRow()-1,20).getValues().forEach(function(r){
    const d=r[1] instanceof Date?r[1]:null;
    if(!d)return;
    if(rs&&re&&(d<rs||d>=re))return;
    out.push({
      receivedDateTime:r[2] instanceof Date?Utilities.formatDate(r[2],tz,"dd-MM-yyyy HH:mm"):String(r[2]||""),
      receivedAgainst:String(r[3]||""),
      ticketId:String(r[4]||""),
      awbNo:String(r[5]||""),
      courier:String(r[6]||""),
      challanNo:String(r[7]||""),
      lastDate:String(r[8]||""),
      itemId:String(r[9]||""),
      itemName:String(r[10]||""),
      qty:Number(r[11]||0),
      unit:String(r[12]||""),
      receivedBy:String(r[13]||""),
      handoverTo:String(r[14]||""),
      location:String(r[15]||""),
      condition:String(r[16]||""),
      remarks:String(r[17]||""),
      submittedByEmail:String(r[18]||""),
      submittedByName:String(r[19]||"")
    });
  });
  out.reverse();
  return{success:true,records:out.slice(0,500)};
}


// =========================================================
// PVC BELT RECEIVED / CONSUMED
// The two data sheets are created automatically when this
// page is opened or when the first record is saved.
// =========================================================

function getPVCBeltSpreadsheet_() {
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensurePVCBeltReceivedSheet_() {
  const ss = getPVCBeltSpreadsheet_();
  let sh = ss.getSheetByName(PVC_BELT_RECEIVED_SHEET_NAME);
  const headers = ["Timestamp", "Received Date", "Belt Type", "Size", "Quantity", "Recorded By Email", "Recorded By Name"];
  if (!sh) sh = ss.insertSheet(PVC_BELT_RECEIVED_SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach(function (h, i) {
      if (!String(current[i] || "").trim()) sh.getRange(1, i + 1).setValue(h);
    });
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensurePVCBeltConsumedSheet_() {
  const ss = getPVCBeltSpreadsheet_();
  let sh = ss.getSheetByName(PVC_BELT_CONSUMED_SHEET_NAME);
  const headers = ["Timestamp", "Consumed Date", "Used In", "Belt Type", "Station No.", "Ticket ID", "Quantity", "Recorded By Email", "Recorded By Name"];
  if (!sh) sh = ss.insertSheet(PVC_BELT_CONSUMED_SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  } else {
    const lastCol = Math.max(sh.getLastColumn(), headers.length);
    let current = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v || "").trim(); });
    // If this sheet already existed with the earlier 8-column format,
    // insert the required department field after Consumed Date.
    if (current.indexOf("Used In") < 0 && current[1] === "Consumed Date") {
      sh.insertColumnAfter(2);
      sh.getRange(1, 3).setValue("Used In");
      current = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0].map(function (v) { return String(v || "").trim(); });
    }
    headers.forEach(function (h, i) {
      if (!current[i]) sh.getRange(1, i + 1).setValue(h);
    });
    sh.setFrozenRows(1);
  }
  return sh;
}

function savePVCBeltReceived(sessionId, data) {
  requirePortalPermission_(sessionId, "spareParts", "create");

  const user = validateSession_(sessionId);
  if (!user) return { success: false, message: "Session expired. Please login again." };
  data = data || {};
  const date = String(data.date || "").trim();
  const beltType = String(data.beltType || "").trim();
  const size = String(data.size || "").trim();
  const qty = Number(data.qty || 0);
  if (!date || PVC_BELT_TYPES.indexOf(beltType) < 0 || PVC_BELT_SIZES.indexOf(size) < 0 || !Number.isFinite(qty) || qty <= 0) {
    return { success: false, message: "Please select date, belt type, size and enter a valid quantity." };
  }
  const dt = new Date(date + "T00:00:00");
  if (isNaN(dt.getTime())) return { success: false, message: "Invalid received date." };
  const sh = ensurePVCBeltReceivedSheet_();
  sh.appendRow([new Date(), dt, beltType, size, qty, String(user.email || ""), String(user.name || "")]);
  SpreadsheetApp.flush();
  try { sendConfiguredPortalNotice_("pvcBelt","PVC Belt Received","PVC belt received entry: "+beltType+", "+size+", Qty "+qty+", Date "+date+", by "+String(user.name||user.email)+"."); } catch(mailError) { Logger.log("PVC belt received notification failed: "+mailError.message); }
  return { success: true, message: "PVC belt received entry saved successfully." };
}

function savePVCBeltConsumed(sessionId, data) {
  requirePortalPermission_(sessionId, "spareParts", "create");

  const user = validateSession_(sessionId);
  if (!user) return { success: false, message: "Session expired. Please login again." };
  data = data || {};
  const date = String(data.date || "").trim();
  const usedIn = String(data.usedIn || "").trim();
  const beltType = String(data.beltType || "").trim();
  const stationNo = String(data.stationNo || "").trim();
  const ticketId = String(data.ticketId || "").trim();
  const qty = Number(data.qty || 0);
  if (!date || PVC_BELT_USED_IN.indexOf(usedIn) < 0 || PVC_BELT_TYPES.indexOf(beltType) < 0 ||
      !stationNo || !ticketId || !Number.isFinite(qty) || qty <= 0) {
    return { success: false, message: "First select CBS PC / Nido Bagging, then fill all consumed belt details and a valid quantity." };
  }
  const dt = new Date(date + "T00:00:00");
  if (isNaN(dt.getTime())) return { success: false, message: "Invalid consumed date." };
  const sh = ensurePVCBeltConsumedSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) { return String(v || "").trim(); });
  const row = [];
  row[headers.indexOf("Timestamp")] = new Date();
  row[headers.indexOf("Consumed Date")] = dt;
  row[headers.indexOf("Used In")] = usedIn;
  row[headers.indexOf("Belt Type")] = beltType;
  row[headers.indexOf("Station No.")] = stationNo;
  row[headers.indexOf("Ticket ID")] = ticketId;
  row[headers.indexOf("Quantity")] = qty;
  row[headers.indexOf("Recorded By Email")] = String(user.email || "");
  row[headers.indexOf("Recorded By Name")] = String(user.name || "");
  sh.appendRow(headers.map(function (_, i) { return row[i] === undefined ? "" : row[i]; }));
  SpreadsheetApp.flush();
  try { sendConfiguredPortalNotice_("pvcBelt","PVC Belt Consumed","PVC belt consumed entry: "+beltType+", Station "+stationNo+", "+usedIn+", Qty "+qty+", Ticket "+ticketId+", Date "+date+", by "+String(user.name||user.email)+"."); } catch(mailError) { Logger.log("PVC belt consumed notification failed: "+mailError.message); }
  return { success: true, message: "PVC belt consumed entry saved successfully." };
}

function getPVCBeltRecords(sessionId, months) {
  const user = validateSession_(sessionId);
  if (!user) return { success: false, message: "Session expired. Please login again.", records: [] };

  // Opening the page creates both data sheets automatically if they do not exist.
  const receivedSheet = ensurePVCBeltReceivedSheet_();
  const consumedSheet = ensurePVCBeltConsumedSheet_();
  const monthCount = Math.max(1, Math.min(120, Number(months) || 1));
  const now = new Date();
  const fromDate = new Date(now.getFullYear(), now.getMonth() - monthCount + 1, 1);
  const toDate = new Date(now.getFullYear(), now.getMonth() + 1, 1);
  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  const records = [];

  if (receivedSheet.getLastRow() > 1) {
    receivedSheet.getRange(2, 1, receivedSheet.getLastRow() - 1, Math.max(7, receivedSheet.getLastColumn())).getValues().forEach(function (r) {
      const d = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (isNaN(d.getTime()) || d < fromDate || d >= toDate) return;
      records.push({
        entryType: "Received",
        date: Utilities.formatDate(d, tz, "yyyy-MM-dd"),
        usedIn: "",
        beltType: String(r[2] || ""),
        size: String(r[3] || ""),
        stationNo: "",
        ticketId: "",
        qty: Number(r[4] || 0),
        timestamp: r[0] instanceof Date ? r[0].getTime() : 0
      });
    });
  }

  if (consumedSheet.getLastRow() > 1) {
    const lastCol = consumedSheet.getLastColumn();
    const headers = consumedSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v || "").trim(); });
    const values = consumedSheet.getRange(2, 1, consumedSheet.getLastRow() - 1, lastCol).getValues();
    values.forEach(function (r) {
      const dateIndex = headers.indexOf("Consumed Date");
      const d = dateIndex >= 0 ? (r[dateIndex] instanceof Date ? r[dateIndex] : new Date(r[dateIndex])) : null;
      if (!d || isNaN(d.getTime()) || d < fromDate || d >= toDate) return;
      function value(name) { const i = headers.indexOf(name); return i >= 0 ? r[i] : ""; }
      const stamp = value("Timestamp");
      records.push({
        entryType: "Consumed",
        date: Utilities.formatDate(d, tz, "yyyy-MM-dd"),
        usedIn: String(value("Used In") || ""),
        beltType: String(value("Belt Type") || ""),
        size: "",
        stationNo: String(value("Station No.") || ""),
        ticketId: String(value("Ticket ID") || ""),
        qty: Number(value("Quantity") || 0),
        timestamp: stamp instanceof Date ? stamp.getTime() : 0
      });
    });
  }

  records.sort(function (a, b) { return b.date.localeCompare(a.date) || b.timestamp - a.timestamp; });
  const receivedTotal = records.filter(function (r) { return r.entryType === "Received"; }).reduce(function (sum, r) { return sum + Number(r.qty || 0); }, 0);
  const consumedTotal = records.filter(function (r) { return r.entryType === "Consumed"; }).reduce(function (sum, r) { return sum + Number(r.qty || 0); }, 0);
  const cbsTotal = records.filter(function (r) { return r.entryType === "Consumed" && r.usedIn === "CBS PC"; }).reduce(function (sum, r) { return sum + Number(r.qty || 0); }, 0);
  const nidoTotal = records.filter(function (r) { return r.entryType === "Consumed" && r.usedIn === "Nido Bagging"; }).reduce(function (sum, r) { return sum + Number(r.qty || 0); }, 0);
  return {
    success: true,
    months: monthCount,
    records: records.slice(0, 3000),
    totals: { received: receivedTotal, consumed: consumedTotal, net: receivedTotal - consumedTotal, cbsPC: cbsTotal, nidoBagging: nidoTotal }
  };
}

function ensureScheduledDowntimeSheet_(){const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(SCHEDULED_DOWNTIME_SHEET_NAME);const h=["Timestamp","Date","From Time","To Time","Ticket No.","Work Details","Work Done By","Duration (hrs)"];if(!sh)sh=ss.insertSheet(SCHEDULED_DOWNTIME_SHEET_NAME);if(sh.getLastRow()===0)sh.getRange(1,1,1,h.length).setValues([h]);else{const cur=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),h.length)).getValues()[0];h.forEach(function(v,i){if(!String(cur[i]||"").trim())sh.getRange(1,i+1).setValue(v);});}return sh;}
function saveScheduledDowntime(sessionId,data){
  requirePortalPermission_(sessionId, "maintenance", "create");
const s=validateSession_(sessionId);if(!s)return{success:false,message:"Session expired. Please login again."};data=data||{};const d=String(data.date||"").trim(),ticket=String(data.ticketNo||"").trim(),ft=String(data.fromTime||"").trim(),tt=String(data.toTime||"").trim(),work=String(data.workDetails||"").trim(),by=String(data.workDoneBy||"").trim();if(!d||!ticket||!ft||!tt||!work||!by)return{success:false,message:"Date, Ticket No., From Time, To Time, Work Details and Work Done By are required."};const st=new Date(d+"T"+ft),en=new Date(d+"T"+tt);if(isNaN(st)||isNaN(en)||en<=st)return{success:false,message:"To Time must be after From Time."};const h=(en-st)/3600000;ensureScheduledDowntimeSheet_().getRange(ensureScheduledDowntimeSheet_().getLastRow()+1,1,1,8).setValues([[new Date(),new Date(d+"T00:00:00"),st,en,ticket,work,by,h]]);try{sendConfiguredPortalNotice_("scheduledDowntime","Scheduled Downtime Activity","Scheduled downtime saved by "+String(s.name||s.email)+". Date: "+d+", Ticket: "+ticket+", Work: "+work+", Duration: "+h+" hours.");}catch(mailError){Logger.log("Scheduled downtime notification failed: "+mailError.message);}return{success:true,message:"Scheduled downtime saved successfully."};}
function getScheduledDowntimeRecords(sessionId,fromDate,toDate){const s=validateSession_(sessionId);if(!s)return{success:false,message:"Session expired. Please login again.",records:[]};const sh=ensureScheduledDowntimeSheet_();if(sh.getLastRow()<2)return{success:true,records:[]};let rs=null,re=null;if(fromDate&&toDate){const a=String(fromDate).split("-"),b=String(toDate).split("-");if(a.length===3&&b.length===3){rs=new Date(+a[0],+a[1]-1,+a[2]);re=new Date(+b[0],+b[1]-1,+b[2]+1);}}const tz=Session.getScriptTimeZone()||"Asia/Kolkata",out=[];sh.getRange(2,1,sh.getLastRow()-1,8).getValues().forEach(function(r){const d=r[1] instanceof Date?r[1]:null;if(!d)return;if(rs&&re&&(d<rs||d>=re))return;out.push({date:Utilities.formatDate(d,tz,"dd-MM-yyyy"),ticketNo:String(r[4]||""),fromTime:r[2]instanceof Date?Utilities.formatDate(r[2],tz,"HH:mm"):String(r[2]||""),toTime:r[3]instanceof Date?Utilities.formatDate(r[3],tz,"HH:mm"):String(r[3]||""),duration:Number(r[7]||0).toFixed(2),workDoneBy:String(r[6]||""),workDetails:String(r[5]||"")});});out.reverse();return{success:true,records:out.slice(0,500)};}
// =========================================================
// NETLIFY API GATEWAY
// =========================================================

function doPost(e) {
  try {
    const request = JSON.parse(
      (e && e.postData && e.postData.contents) || "{}"
    );

    const fn = String(request.fn || "").trim();
    const args = Array.isArray(request.args) ? request.args : [];

    const handlers = {
      addUserManagement: addUserManagement,
      getBreakdownDateRangeSummary: getBreakdownDateRangeSummary,
      getBreakdownRecords: getBreakdownRecords,
      getCurrentUser: getCurrentUser,
      getDashboardData: getDashboardData,
      getLoginHistory: getLoginHistory,
      getPMChecklist: getPMChecklist,
      getCurrentPortalPermissions: getCurrentPortalPermissions,
      savePMChecklistRow: savePMChecklistRow,
      getUserManagement: getUserManagement,
      loginUser: loginUser,
      logoutUser: logoutUser,
      removeUserManagement: removeUserManagement,
      requestAdminRecovery: requestAdminRecovery,
      saveBreakdownEntry: saveBreakdownEntry,
      saveProblemObservation: saveProblemObservation,
      saveReceivedSpare: saveReceivedSpare,
      saveScheduledDowntime: saveScheduledDowntime,
      saveSettings: saveSettings,
      searchItemMaster: searchItemMaster,
      searchItemMasterFilters: searchItemMasterFilters,
      searchRequests: searchRequests,
      submitBreakdownFromMaterialRequest: submitBreakdownFromMaterialRequest,
      submitConsumedSpareDetails: submitConsumedSpareDetails,
      submitRequest: submitRequest,
      updateDashboardRequest: updateDashboardRequest,
      updateProblemObservation: updateProblemObservation,
      verifyAdminRecovery: verifyAdminRecovery
    };

    if (!handlers[fn]) {
      return apiJson_({
        success: false,
        message: "API function is not allowed: " + fn
      });
    }

    const result = handlers[fn].apply(null, args);

    return apiJson_(
      result == null ? null : result
    );

  } catch (error) {
    return apiJson_({
      success: false,
      message:
        error && error.message
          ? error.message
          : String(error)
    });
  }
}


function apiJson_(payload) {
  return ContentService
    .createTextOutput(
      JSON.stringify(payload)
    )
    .setMimeType(
      ContentService.MimeType.JSON
    );
}

// =========================================================
// PM CHECKLIST SOURCE INTEGRATION + T-3004 ROSTER
// =========================================================
const PM_CHECKLIST_SPREADSHEET_ID_ = "1T6ZAPjDngICXtm2Cus-iihj1RzyM4FH041-j2ndLeRA";
const PM_CHECKLIST_SHEET_GID_URL_ = "https://docs.google.com/spreadsheets/d/1T6ZAPjDngICXtm2Cus-iihj1RzyM4FH041-j2ndLeRA/edit?usp=sharing";
const T3004_ROSTER_SPREADSHEET_ID_ = "11MzzsLiYf6L2CIrRU-FgBC5Y-wKlGoPrVZZbnvn5Okc";
const T3004_ROSTER_GID_ = 1314673857;

function portalRequireSession_(sessionId) {
  const session = validateSession_(sessionId);
  if (!session) throw new Error("Session expired. Please login again.");
  return session;
}
function pmNormHeader_(value) {
  return String(value == null ? "" : value).toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}
function pmFindHeader_(headers, candidates) {
  const norm = headers.map(pmNormHeader_);
  for (let i=0;i<candidates.length;i++) {
    const wanted=pmNormHeader_(candidates[i]);
    let idx=norm.indexOf(wanted);
    if(idx>=0)return idx;
    idx=norm.findIndex(function(h){return h && (h.indexOf(wanted)>=0 || wanted.indexOf(h)>=0);});
    if(idx>=0)return idx;
  }
  return -1;
}
function pmDateISO_(value) {
  if (value === null || value === undefined || value === "") return "";
  const tz = Session.getScriptTimeZone() || "Asia/Kolkata";
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, "yyyy-MM-dd");
  }
  const s = String(value).trim();
  let m, y, mo, d;
  // Accept a date anywhere in a tab name, including extra labels/spaces:
  // "11 Jan 2026", "11 Jan 2026 - PM", "PM - 11 Jan 2026".
  m = s.match(/(?:^|[^0-9])(\d{1,2})[\s,._-]+([A-Za-z]{3,9})[\s,._-]+(\d{4})(?:$|[^0-9])/);
  if (m) {
    const months = {jan:0,feb:1,mar:2,apr:3,may:4,jun:5,jul:6,aug:7,sep:8,oct:9,nov:10,dec:11};
    const key = m[2].toLowerCase().slice(0,3);
    if (!(key in months)) return "";
    d = Number(m[1]); mo = months[key]; y = Number(m[3]);
  } else {
    m = s.match(/(?:^|[^0-9])(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:$|[^0-9])/);
    if (m) { y=Number(m[1]); mo=Number(m[2])-1; d=Number(m[3]); }
    else {
      m = s.match(/(?:^|[^0-9])(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:$|[^0-9])/);
      if (!m) return "";
      d=Number(m[1]); mo=Number(m[2])-1; y=Number(m[3]);
    }
  }
  const dt = new Date(y, mo, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo || dt.getDate() !== d) return "";
  return Utilities.formatDate(dt, tz, "yyyy-MM-dd");
}
function getPMChecklist(sessionId) {
  requirePortalPermission_(sessionId,"maintenance","view");
  const ss=SpreadsheetApp.openById(PM_CHECKLIST_SPREADSHEET_ID_);
  const tz=Session.getScriptTimeZone()||"Asia/Kolkata";
  const today=Utilities.formatDate(new Date(),tz,"yyyy-MM-dd");
  const rows=[];
  let foundDateTabs=0;
  ss.getSheets().forEach(function(sh){
    // Each date has its own tab, e.g. "11 Jan 2026" or "09 Jan 2026".
    const tabDate=pmDateISO_(sh.getName());
    if(!tabDate) return;
    foundDateTabs++;
    const values=sh.getDataRange().getValues();
    let headerIndex=-1, headers=[];
    for(let r=0;r<Math.min(values.length,30);r++){
      const h=values[r].map(pmNormHeader_);
      if(h.indexOf("module")>=0 && h.indexOf("component")>=0 && (h.indexOf("checkpoint")>=0||h.indexOf("check point")>=0)){
        headerIndex=r;headers=values[r];break;
      }
    }
    if(headerIndex<0) return;
    const ix={
      module:pmFindHeader_(headers,["Module"]),
      component:pmFindHeader_(headers,["Component"]),
      checkpoint:pmFindHeader_(headers,["CheckPoint","Check Point"]),
      frequency:pmFindHeader_(headers,["Frequency"]),
      correction:pmFindHeader_(headers,["Correction method","Issue / Correction","Issue Description"]),
      issue:pmFindHeader_(headers,["Issue Description"]),
      // Per user's sheet, Condition is Column P (1-based).
      condition:15,
      eta:pmFindHeader_(headers,["ETA","Next Planned Date","Next Planned","Planned Date"])
    };
    for(let r=headerIndex+1;r<values.length;r++){
      const row=values[r];
      const module=ix.module>=0?row[ix.module]:"";
      const component=ix.component>=0?row[ix.component]:"";
      const checkpoint=ix.checkpoint>=0?row[ix.checkpoint]:"";

      // Exclude navigation/link rows such as "Go to Main Sheet".
      // Search the entire row because the label may be in Module or a hyperlink cell,
      // and the visible text can include an icon/emoji.
      const navRowText = row.map(function(v) {
        return String(v == null ? "" : v).replace(/\s+/g, " ").trim().toLowerCase();
      }).join(" ");
      const taskFieldsBlank =
        !String(module || "").trim() &&
        !String(component || "").trim() &&
        !String(checkpoint || "").trim();
      const isNavigationRow =
        /go\s+to\s+main\s+sheet/.test(navRowText) ||
        /click\s+here\s+to\s+go\s+to\s+main\s+sheet/.test(navRowText) ||
        /back\s+to\s+main\s+sheet/.test(navRowText) ||
        (taskFieldsBlank && /\b(main sheet|home|back)\b/.test(navRowText));
      if (isNavigationRow) continue;

      // A real checklist point must have at least one task-identifying field.
      if (taskFieldsBlank) continue;
      const condition=String(row[ix.condition]||"").trim();
      const status=condition.toLowerCase()==="ok"?"Completed":"Open";
      const eta=ix.eta>=0?row[ix.eta]:"";
      rows.push({
        sheetId:sh.getSheetId(),sheetName:sh.getName(),rowNumber:r+1,
        module:String(module||""),component:String(component||""),checkpoint:String(checkpoint||""),
        frequency:ix.frequency>=0?String(row[ix.frequency]||"").trim():"",
        dueDate:tabDate,dueDates:[tabDate],isDueToday:tabDate===today,
        condition:condition,status:status,
        issueDescription:ix.issue>=0?String(row[ix.issue]||""):"",
        correctionMethod:ix.correction>=0?String(row[ix.correction]||""):"",
        eta:eta instanceof Date?Utilities.formatDate(eta,tz,"dd-MMM-yyyy"):String(eta||""),
        etaISO:pmDateISO_(eta)
      });
    }
  });
  if(!foundDateTabs) return {success:false,message:"No date-named tabs were detected. Expected names like 11 Jan 2026. Check the exact tab names and confirm this deployment is connected to the correct spreadsheet."};
  rows.sort(function(a,b){return a.dueDate.localeCompare(b.dueDate)||a.rowNumber-b.rowNumber;});
  return {success:true,rows:rows,today:today,dateTabsFound:foundDateTabs,message:"Detected "+foundDateTabs+" date tabs; loaded "+rows.length+" checklist points."};
}
function savePMChecklistRow(sessionId,sheetIdOrRowNumber,rowNumberOrCondition,conditionOrEta,maybeEta) {
  const session=requirePortalPermission_(sessionId,"maintenance","edit");
  // Supports both the current frontend signature (sessionId, sheetId, rowNumber, condition, eta)
  // and the older signature (sessionId, rowNumber, condition, eta).
  let targetSheetId=0,rowNum,cleanCondition,cleanEta;
  if(arguments.length>=5){
    targetSheetId=Number(sheetIdOrRowNumber)||0;
    rowNum=Number(rowNumberOrCondition);
    cleanCondition=String(conditionOrEta||"").trim();
    cleanEta=String(maybeEta||"").trim();
  } else {
    rowNum=Number(sheetIdOrRowNumber);
    cleanCondition=String(rowNumberOrCondition||"").trim();
    cleanEta=String(conditionOrEta||"").trim();
  }
  if(!Number.isInteger(rowNum)||rowNum<2)throw new Error("Invalid PM checklist row.");
  if(["","OK","Not OK"].indexOf(cleanCondition)<0)throw new Error("Condition must be OK or Not OK.");
  if(cleanCondition!=="OK"&&!cleanEta)throw new Error("ETA / next planned date is required while the task is open.");
  const ss=SpreadsheetApp.openById(PM_CHECKLIST_SPREADSHEET_ID_);
  let sh=null, headerRow=0, headers=null;
  if(targetSheetId){
    sh=ss.getSheets().find(function(s){return Number(s.getSheetId())===targetSheetId;})||null;
  }
  if(!sh){
    ss.getSheets().some(function(s){const data=s.getDataRange().getValues();for(let r=0;r<Math.min(data.length,30);r++){const h=data[r].map(pmNormHeader_);if(h.indexOf("module")>=0&&h.indexOf("component")>=0&&(h.indexOf("checkpoint")>=0||h.indexOf("check point")>=0)){sh=s;headerRow=r+1;headers=data[r];return true;}}return false;});
  } else {
    const data=sh.getDataRange().getValues();
    for(let r=0;r<Math.min(data.length,30);r++){
      const h=data[r].map(pmNormHeader_);
      if(h.indexOf("module")>=0&&h.indexOf("component")>=0&&(h.indexOf("checkpoint")>=0||h.indexOf("check point")>=0)){headerRow=r+1;headers=data[r];break;}
    }
  }
  if(!sh)throw new Error("PM header row not found.");
  if(!headers||!headerRow)throw new Error("PM header row not found in selected sheet.");
  if(rowNum<=headerRow)throw new Error("Selected row is not a task row.");
  // Condition is explicitly stored in Column P of the source sheet.
  sh.getRange(rowNum,16).setValue(cleanCondition);
  let ei=pmFindHeader_(headers,["ETA","Next Planned Date","Next Planned","Planned Date"]);
  if(ei<0){ei=headers.length;sh.getRange(headerRow,ei+1).setValue("ETA / Next Planned Date");}
  sh.getRange(rowNum,ei+1).setValue(cleanEta?new Date(cleanEta+"T12:00:00"): "");
  try { const pmEventKey=cleanCondition==="OK"?"pmWorkCompleted":"pmWorkOpen"; sendConfiguredPortalNotice_(pmEventKey,cleanCondition==="OK"?"PM Work Completed":"PM Work Open / Pending","Task row "+rowNum+" updated by "+String(session.name||session.email)+". Condition: "+(cleanCondition||"Open")+". ETA: "+(cleanEta||"Not set")+"."); } catch(mailError) { Logger.log("PM notification failed: "+mailError.message); }
  return {success:true,message:"PM checklist updated in source sheet."};
}
function getT3004ShiftRoster(sessionId) {
  requirePortalPermission_(sessionId,"shiftRoster","view");
  const ss=SpreadsheetApp.openById(T3004_ROSTER_SPREADSHEET_ID_);
  let sh=null;
  ss.getSheets().some(function(s){if(String(s.getName()).toLowerCase()==="t3004"||String(s.getName()).toLowerCase()==="t-3004"){sh=s;return true;}return false;});
  if(!sh) sh=ss.getSheets().find(function(s){return s.getSheetId()===T3004_ROSTER_GID_;});
  if(!sh) return {success:false,message:"T3004 roster tab not found. Check sheet tab name or configured GID."};
  const data=sh.getDataRange().getDisplayValues();
  let hr=-1, empIdx=-1, nameIdx=-1, streamIdx=-1;
  for(let r=0;r<Math.min(data.length,25);r++){
    const row=data[r].map(pmNormHeader_);
    const e=row.findIndex(function(x){return x==="emp id"||x==="employee id"||x==="emp no"||x==="employee no";});
    const n=row.findIndex(function(x){return x==="name"||x==="employee name";});
    const st=row.findIndex(function(x){return x==="stream"||x==="department";});
    if(e>=0&&n>=0){hr=r;empIdx=e;nameIdx=n;streamIdx=st;break;}
  }
  if(hr<0)return {success:false,message:"Roster header row not found (expected EMP ID and Name)."};
  const headers=data[hr].map(function(v){return String(v||"").trim();});
  // Date/month labels may occupy several rows above the employee header; resolve them dynamically.
  const headerStart=Math.max(0,hr-8);
  let detectedMonth=null, detectedYear=null;
  for(let r=headerStart;r<=hr;r++){
    for(let c=0;c<data[r].length;c++){
      const cell=String(data[r][c]||"").trim();
      const parsed=new Date(cell);
      if(cell && isNaN(Number(cell)) && !isNaN(parsed.getTime()) && /[A-Za-z]/.test(cell) && parsed.getFullYear()>2000){
        detectedMonth=parsed.getMonth();detectedYear=parsed.getFullYear();
      }
    }
  }
  if(detectedMonth===null){detectedMonth=new Date().getMonth();detectedYear=new Date().getFullYear();}
  for(let c=0;c<headers.length;c++){
    if(c<=Math.max(empIdx,nameIdx,streamIdx))continue;
    let candidate=headers[c];
    if(!candidate){
      for(let r=hr-1;r>=headerStart;r--){const v=String(data[r][c]||"").trim();if(v){candidate=v;break;}}
    }
    if(/^\\d{1,2}$/.test(candidate)){
      const day=Number(candidate);
      if(day>=1&&day<=31)headers[c]=Utilities.formatDate(new Date(detectedYear,detectedMonth,day),Session.getScriptTimeZone()||"Asia/Kolkata","dd-MMM-yyyy");
    } else if(candidate && /^\\d{1,2}[\\/-]\\d{1,2}([\\/-]\\d{2,4})?$/.test(candidate)){
      const parsedDate=new Date(candidate);
      if(!isNaN(parsedDate.getTime()))headers[c]=Utilities.formatDate(parsedDate,Session.getScriptTimeZone()||"Asia/Kolkata","dd-MMM-yyyy");
    } else if(candidate && !headers[c]) headers[c]=candidate;
  }
  const rows=[];
  for(let r=hr+1;r<data.length;r++){
    const row=data[r];
    if(!String(row[empIdx]||"").trim()&&!String(row[nameIdx]||"").trim())continue;
    if(!String(row[nameIdx]||"").trim())continue;
    if(streamIdx>=0 && !String(row[streamIdx]||"").trim())continue;
    rows.push(row);
  }
  const monthLabel=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||"Asia/Kolkata","MMMM yyyy");
  return {success:true,monthLabel:monthLabel,headers:headers,rows:rows};
}
function sendConfiguredPortalNotice_(activityKey,subject,body) {
  const cfg=getActivityNotificationConfig_()[activityKey];
  if(!cfg||!cfg.enabled||cfg.mode==="off")return;
  const recipients=getConfiguredEmailRecipients_(activityKey);
  if(!recipients.to||!recipients.to.length)return;
  MailApp.sendEmail({to:recipients.to.join(","),cc:(recipients.cc||[]).join(","),subject:subject,body:body});
}


// =========================================================
// ROLE DEFAULTS + PER-USER PERMISSION OVERRIDES
// =========================================================
const PORTAL_PERMISSION_KEY_ = "PORTAL_PERMISSION_MATRIX_V1";
const PORTAL_PERMISSION_MODULES_ = ["dashboard","maintenance","spareParts","shiftRoster","settings"];
const PORTAL_PERMISSION_ACTIONS_ = ["view","create","edit","delete","approveReject","adminSettings"];
function portalDefaultPermissionMatrix_() {
  const admin={}, operator={};
  PORTAL_PERMISSION_MODULES_.forEach(function(m){
    admin[m]={};operator[m]={};
    PORTAL_PERMISSION_ACTIONS_.forEach(function(a){
      admin[m][a]=true;
      operator[m][a]=(a==="view"||a==="create"||a==="edit") && m!=="settings";
    });
  });
  return {roles:{Admin:admin,Operator:operator},users:{}};
}
function getPortalPermissionMatrix(sessionId) {
  requireAdminSession_(sessionId);
  const raw=PropertiesService.getScriptProperties().getProperty(PORTAL_PERMISSION_KEY_);
  let saved={};
  try{saved=raw?JSON.parse(raw):{};}catch(e){saved={};}
  const defaults=portalDefaultPermissionMatrix_();
  const result={roles:defaults.roles,users:{}};
  PORTAL_PERMISSION_MODULES_.forEach(function(m){
    ["Admin","Operator"].forEach(function(role){
      if(saved.roles&&saved.roles[role]&&saved.roles[role][m]){
        PORTAL_PERMISSION_ACTIONS_.forEach(function(a){
          if(typeof saved.roles[role][m][a]==="boolean")result.roles[role][m][a]=saved.roles[role][m][a];
        });
      }
    });
  });
  Object.keys(saved.users||{}).forEach(function(email){
    const cleanEmail=String(email).trim().toLowerCase();
    if(!cleanEmail)return;
    result.users[cleanEmail]={};
    PORTAL_PERMISSION_MODULES_.forEach(function(m){
      if(saved.users[email]&&saved.users[email][m]){
        result.users[cleanEmail][m]={};
        PORTAL_PERMISSION_ACTIONS_.forEach(function(a){
          if(typeof saved.users[email][m][a]==="boolean")result.users[cleanEmail][m][a]=saved.users[email][m][a];
        });
      }
    });
  });
  return result;
}
function savePortalPermissionMatrix(sessionId,matrix) {
  requireAdminSession_(sessionId);
  if(!matrix||typeof matrix!=="object"||!matrix.roles)throw new Error("Permission matrix is required.");
  const clean=portalDefaultPermissionMatrix_();
  ["Admin","Operator"].forEach(function(role){
    PORTAL_PERMISSION_MODULES_.forEach(function(m){
      PORTAL_PERMISSION_ACTIONS_.forEach(function(a){
        if(matrix.roles[role]&&matrix.roles[role][m]&&typeof matrix.roles[role][m][a]==="boolean")clean.roles[role][m][a]=matrix.roles[role][m][a];
      });
    });
  });
  Object.keys(matrix.users||{}).forEach(function(email){
    const e=String(email||"").trim().toLowerCase();
    if(!isValidEmail_(e))throw new Error("Invalid override email: "+e);
    clean.users[e]={};
    PORTAL_PERMISSION_MODULES_.forEach(function(m){
      if(!matrix.users[email]||!matrix.users[email][m])return;
      clean.users[e][m]={};
      PORTAL_PERMISSION_ACTIONS_.forEach(function(a){
        if(typeof matrix.users[email][m][a]==="boolean")clean.users[e][m][a]=matrix.users[email][m][a];
      });
    });
  });
  PropertiesService.getScriptProperties().setProperty(PORTAL_PERMISSION_KEY_,JSON.stringify(clean));
  return {success:true,message:"Role defaults and per-user overrides saved."};
}
function hasPortalPermission_(session,moduleName,action) {
  if(!session)return false;
  const email=String(session.email||"").trim().toLowerCase();
  const role=String(session.role||"Operator").trim();
  const matrix=portalDefaultPermissionMatrix_();
  const raw=PropertiesService.getScriptProperties().getProperty(PORTAL_PERMISSION_KEY_);
  if(raw){try{const saved=JSON.parse(raw);if(saved.roles)matrix.roles=saved.roles;if(saved.users)matrix.users=saved.users;}catch(e){}}
  const userOverride=matrix.users&&matrix.users[email]&&matrix.users[email][moduleName];
  if(userOverride&&typeof userOverride[action]==="boolean")return userOverride[action];
  const roleMap=matrix.roles&&matrix.roles[role]&&matrix.roles[role][moduleName];
  if(roleMap&&typeof roleMap[action]==="boolean")return roleMap[action];
  return false;
}
function requirePortalPermission_(sessionId,moduleName,action) {
  const session=portalRequireSession_(sessionId);
  if(!hasPortalPermission_(session,moduleName,action))throw new Error("Access denied: permission '"+action+"' is not enabled for "+moduleName+".");
  return session;
}

function getCurrentPortalPermissions(sessionId) {
  const session=portalRequireSession_(sessionId);
  const result={};
  PORTAL_PERMISSION_MODULES_.forEach(function(m){
    result[m]={};
    PORTAL_PERMISSION_ACTIONS_.forEach(function(a){result[m][a]=hasPortalPermission_(session,m,a);});
  });
  return result;
}

function setupT3004RosterNotificationTrigger(sessionId) {
  requireAdminSession_(sessionId);
  const triggers=ScriptApp.getProjectTriggers();
  const exists=triggers.some(function(t){return t.getHandlerFunction()==="onT3004RosterEdit";});
  if(!exists)ScriptApp.newTrigger("onT3004RosterEdit").forSpreadsheet(T3004_ROSTER_SPREADSHEET_ID_).onEdit().create();
  return {success:true,message:exists?"T-3004 roster email trigger is already installed.":"T-3004 roster edit email trigger installed."};
}
function onT3004RosterEdit(e) {
  try {
    if(!e||!e.range)return;
    const sh=e.range.getSheet();
    if(sh.getSheetId()!==T3004_ROSTER_GID_)return;
    const cfg=getActivityNotificationConfig_()["shiftRoster"];
    if(!cfg||!cfg.enabled||cfg.mode==="off")return;
    sendConfiguredPortalNotice_("shiftRoster","T-3004 Shift Roster Changed",
      "The T-3004 shift roster was edited in Google Sheets. Sheet: "+sh.getName()+
      ", Cell/range: "+e.range.getA1Notation()+", Editor: "+(Session.getActiveUser().getEmail()||"Google Sheets editor")+
      (e.oldValue!==undefined?" , Previous value: "+e.oldValue:"")+
      (e.value!==undefined?" , New value: "+e.value:"")+".");
  } catch(err) { Logger.log("T-3004 roster edit notification failed: "+err.message); }
}

function setupPMOverdueNotificationTrigger(sessionId) {
  requireAdminSession_(sessionId);
  const triggers=ScriptApp.getProjectTriggers();
  const exists=triggers.some(function(t){return t.getHandlerFunction()==="sendPMOverdueNotifications";});
  if(!exists)ScriptApp.newTrigger("sendPMOverdueNotifications").timeBased().everyDays(1).atHour(8).create();
  return {success:true,message:exists?"Daily PM overdue notification trigger is already installed.":"Daily PM overdue notification trigger installed (around 8 AM script timezone)."};
}
function sendPMOverdueNotifications() {
  try {
    const ss=SpreadsheetApp.openById(PM_CHECKLIST_SPREADSHEET_ID_);
    let found=null;
    ss.getSheets().some(function(sh){const data=sh.getDataRange().getValues();for(let r=0;r<Math.min(data.length,30);r++){const h=data[r].map(pmNormHeader_);if(h.indexOf("module")>=0&&h.indexOf("component")>=0&&(h.indexOf("checkpoint")>=0||h.indexOf("check point")>=0)){found={sheet:sh,row:r+1,values:data,headers:data[r]};return true;}}return false;});
    if(!found)return;
    const ixCondition=pmFindHeader_(found.headers,["Condition"]),ixEta=pmFindHeader_(found.headers,["ETA","Next Planned Date","Next Planned","Planned Date"]),ixModule=pmFindHeader_(found.headers,["Module"]),ixComponent=pmFindHeader_(found.headers,["Component"]),ixCheckpoint=pmFindHeader_(found.headers,["CheckPoint","Check Point"]);
    if(ixEta<0)return;
    const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||"Asia/Kolkata","yyyy-MM-dd");
    const overdue=[];
    for(let r=found.row;r<found.values.length;r++){
      const row=found.values[r], condition=ixCondition>=0?String(row[ixCondition]||"").trim().toLowerCase():"";
      if(condition==="ok")continue;
      const eta=pmDateISO_(row[ixEta]);
      if(eta && eta<today)overdue.push("Row "+(r+1)+": "+String(row[ixModule]||"")+" / "+String(row[ixComponent]||"")+" / "+String(row[ixCheckpoint]||"")+" — ETA "+eta);
    }
    if(overdue.length)sendConfiguredPortalNotice_("pmWorkOverdue","PM Work Overdue ("+overdue.length+")",overdue.join("\n"));
  } catch(err) { Logger.log("PM overdue notification failed: "+err.message); }
}

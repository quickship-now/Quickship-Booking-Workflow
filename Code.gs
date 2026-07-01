var BOOKING_SHEET = "WorkflowBookings";
var USERS_SHEET = "WorkflowUsers";

var BOOKING_HEADERS = [
  "Record ID", "Created At", "Last Updated",
  "Pickup Type", "Billing Party Name", "Customer Name",
  "Shipper Phone No", "Consignee Phone No", "Delivery Address", "Pickup Address",
  "Customer Phone Number", "Customer Email ID", "Weight in KG", "Boxes in Number",
  "Item Package Description", "Preferred Pickup Date", "Special Instructions",
  "Approved By", "Approver Phone No", "Approver Email", "Approver Status",
  "Pickup Assign To", "Assign Date", "Assignee Mail ID", "Pickup Assign Status",
  "Pickup Date", "Dispatch Date", "Actual Weight", "MAWB No", "India Status"
];

var USER_HEADERS = ["Email", "Password", "Role", "Name"];
var DEFAULT_USERS = [
  ["india@quickshipnow.com", "india123", "india", "India Office"],
  ["sahib@quickshipnow.com", "sahib123", "approver", "Approver"],
  ["approver@quickshipnow.com", "approver123", "approver", "Approver"],
  ["pickup@quickshipnow.com", "pickup123", "pickup", "Pickup Assign"]
];

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var page = String(params.page || "public").toLowerCase();
  var recordId = String(params.recordId || "").trim();
  var validPages = ["public", "login", "dashboard", "form"];
  if (validPages.indexOf(page) === -1) page = "public";

  var recordData = {};
  if (page === "form" && recordId) {
    var rec = getOfficeRecord(recordId);
    if (rec && rec.found) recordData = rec.data;
  }

  var tmpl = HtmlService.createTemplateFromFile("Index");
  tmpl.APP_INIT_JSON = JSON.stringify({
    page: page,
    recordId: recordId,
    data: recordData
  });

  return tmpl.evaluate()
    .setTitle("Quickship Booking Workflow")
    .addMetaTag("viewport", "width=device-width,initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function loginOffice(email, password) {
  try {
    var sh = getOrCreateUsersSheet();
    var rows = sh.getDataRange().getValues();
    var em = String(email || "").trim().toLowerCase();
    var pass = String(password || "").trim();

    for (var i = 1; i < rows.length; i++) {
      var rowEmail = String(rows[i][0] || "").trim().toLowerCase();
      var rowPass = String(rows[i][1] || "").trim();
      var role = String(rows[i][2] || "").trim().toLowerCase();
      if (role === "sahib") role = "approver";
      if (rowEmail === em && rowPass === pass && (role === "india" || role === "approver" || role === "pickup")) {
        return {
          success: true,
          user: {
            email: String(rows[i][0] || "").trim(),
            role: role,
            name: String(rows[i][3] || "").trim() || role.toUpperCase()
          }
        };
      }
    }
    return { success: false, error: "Invalid office login credentials." };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function submitPublicForm(fields) {
  try {
    fields = fields || {};
    var required = [
      "Pickup Type",
      "Shipper Phone No",
      "Consignee Phone No",
      "Delivery Address",
      "Pickup Address",
      "Customer Phone Number",
      "Customer Email ID",
      "Preferred Pickup Date",
      "Special Instructions"
    ];

    for (var i = 0; i < required.length; i++) {
      var key = required[i];
      if (!String(fields[key] || "").trim()) {
        return { success: false, error: "Please fill required field: " + key };
      }
    }

    var sh = getOrCreateBookingSheet();
    var recordId = "BK-" + new Date().getTime();
    var now = formatNow_();
    var row = new Array(BOOKING_HEADERS.length).fill("");

    row[bci("Record ID")] = recordId;
    row[bci("Created At")] = now;
    row[bci("Last Updated")] = now;

    BOOKING_HEADERS.forEach(function(h) {
      if (fields[h] !== undefined && fields[h] !== null) {
        row[bci(h)] = String(fields[h]).trim();
      }
    });

    row[bci("Approver Status")] = "Pending";
    row[bci("Pickup Assign Status")] = "Pending";
    row[bci("India Status")] = "Pending";

    sh.appendRow(row);

    var emailResult = sendSubmissionEmail_(fields, recordId, now);

    return {
      success: true,
      recordId: recordId,
      emailSent: emailResult.sent,
      emailError: emailResult.error || ""
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function getOfficeDashboard() {
  try {
    var sh = getOrCreateBookingSheet();
    var rows = sh.getDataRange().getValues();
    if (rows.length < 2) return { success: true, rows: [] };

    var out = [];
    for (var i = rows.length - 1; i >= 1; i--) {
      if (!String(rows[i][0] || "").trim()) continue;
      var obj = {};
      BOOKING_HEADERS.forEach(function(h, idx) {
        obj[h] = (rows[i][idx] !== undefined && rows[i][idx] !== null) ? String(rows[i][idx]) : "";
      });
      out.push(obj);
    }

    return { success: true, rows: out };
  } catch (err) {
    return { success: false, error: err.message, rows: [] };
  }
}

function getOfficeRecord(recordId) {
  try {
    var rid = String(recordId || "").trim();
    if (!rid) return { found: false, data: {} };

    var sh = getOrCreateBookingSheet();
    var rowNum = findRow_(sh, rid);
    if (rowNum < 0) return { found: false, data: {} };

    var vals = sh.getRange(rowNum, 1, 1, BOOKING_HEADERS.length).getValues()[0];
    var obj = {};
    BOOKING_HEADERS.forEach(function(h, i) {
      obj[h] = (vals[i] !== undefined && vals[i] !== null) ? String(vals[i]) : "";
    });

    return { found: true, data: obj };
  } catch (err) {
    return { found: false, data: {}, error: err.message };
  }
}

function saveOfficeSection(payload) {
  try {
    payload = payload || {};
    var section = String(payload.section || "").trim().toLowerCase();
    var rid = String(payload.recordId || "").trim();
    var fields = payload.fields || {};

    if (section === "sahib") section = "approver";
    if (section !== "india" && section !== "approver" && section !== "pickup") {
      return { success: false, error: "Invalid section." };
    }
    if (!rid) {
      return { success: false, error: "Record ID is required." };
    }

    var sh = getOrCreateBookingSheet();
    var rowNum = findRow_(sh, rid);
    if (rowNum < 0) {
      return { success: false, error: "Record not found." };
    }

    if (section === "india") {
      sh.getRange(rowNum, bci("Pickup Date") + 1).setValue(String(fields["Pickup Date"] || "").trim());
      sh.getRange(rowNum, bci("Dispatch Date") + 1).setValue(String(fields["Dispatch Date"] || "").trim());
      sh.getRange(rowNum, bci("Actual Weight") + 1).setValue(String(fields["Actual Weight"] || "").trim());
      sh.getRange(rowNum, bci("MAWB No") + 1).setValue(String(fields["MAWB No"] || "").trim());
      sh.getRange(rowNum, bci("India Status") + 1).setValue("Submitted");
    } else if (section === "approver") {
      sh.getRange(rowNum, bci("Approved By") + 1).setValue(String(fields["Approved By"] || "").trim());
      sh.getRange(rowNum, bci("Approver Phone No") + 1).setValue(String(fields["Approver Phone No"] || "").trim());
      sh.getRange(rowNum, bci("Approver Email") + 1).setValue(String(fields["Approver Email"] || "").trim());
      sh.getRange(rowNum, bci("Approver Status") + 1).setValue("Submitted");
    } else if (section === "pickup") {
      sh.getRange(rowNum, bci("Pickup Assign To") + 1).setValue(String(fields["Pickup Assign To"] || "").trim());
      sh.getRange(rowNum, bci("Assign Date") + 1).setValue(String(fields["Assign Date"] || "").trim());
      sh.getRange(rowNum, bci("Assignee Mail ID") + 1).setValue(String(fields["Assignee Mail ID"] || "").trim());
      sh.getRange(rowNum, bci("Pickup Assign Status") + 1).setValue("Submitted");
    }

    sh.getRange(rowNum, bci("Last Updated") + 1).setValue(formatNow_());

    return { success: true, recordId: rid };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function sendSubmissionEmail_(fields, recordId, submittedAt) {
  var to = String(fields["Customer Email ID"] || "").trim();
  if (!to) return { sent: false, error: "Recipient email is missing." };

  var subject = "Booking Submitted - " + recordId;
  var lines = [
    "Booking submitted successfully.",
    "",
    "Record ID: " + recordId,
    "Submitted At: " + submittedAt,
    "",
    "Form Details:",
    "Pickup Type: " + String(fields["Pickup Type"] || ""),
    "Billing Party Name: " + String(fields["Billing Party Name"] || ""),
    "Customer Name: " + String(fields["Customer Name"] || ""),
    "Shipper Phone No: " + String(fields["Shipper Phone No"] || ""),
    "Consignee Phone No: " + String(fields["Consignee Phone No"] || ""),
    "Delivery Address: " + String(fields["Delivery Address"] || ""),
    "Pickup Address: " + String(fields["Pickup Address"] || ""),
    "Customer Phone Number: " + String(fields["Customer Phone Number"] || ""),
    "Customer Email ID: " + to,
    "Weight in KG: " + String(fields["Weight in KG"] || ""),
    "Boxes in Number: " + String(fields["Boxes in Number"] || ""),
    "Item Package Description: " + String(fields["Item Package Description"] || ""),
    "Preferred Pickup Date: " + String(fields["Preferred Pickup Date"] || ""),
    "Special Instructions: " + String(fields["Special Instructions"] || "")
  ];

  try {
    MailApp.sendEmail({
      to: to,
      subject: subject,
      body: lines.join("\n"),
      name: "Quickship Booking Workflow"
    });
    return { sent: true };
  } catch (err) {
    return { sent: false, error: err.message };
  }
}

function bci(header) {
  return BOOKING_HEADERS.indexOf(header);
}

function findRow_(sh, recordId) {
  var data = sh.getDataRange().getValues();
  var rid = String(recordId || "").trim();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === rid) return i + 1;
  }
  return -1;
}

function formatNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
}

function getOrCreateBookingSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(BOOKING_SHEET);
  if (!sh) {
    sh = ss.insertSheet(BOOKING_SHEET);
    sh.appendRow(BOOKING_HEADERS);
    sh.setFrozenRows(1);
  } else {
    syncBookingHeaders_(sh);
  }

  sh.getRange(1, 1, 1, BOOKING_HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#1f2937")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center")
    .setWrap(false);
  for (var c = 1; c <= BOOKING_HEADERS.length; c++) {
    sh.setColumnWidth(c, 150);
  }
  return sh;
}

function syncBookingHeaders_(sh) {
  var range = sh.getDataRange();
  var data = range.getValues();
  if (!data.length) {
    sh.appendRow(BOOKING_HEADERS);
    return;
  }

  var currentHeaders = data[0].map(function(h) { return String(h || "").trim(); });
  var alreadySynced = currentHeaders.length === BOOKING_HEADERS.length &&
    BOOKING_HEADERS.every(function(h, i) { return currentHeaders[i] === h; });
  if (alreadySynced) return;

  var rebuilt = [BOOKING_HEADERS.slice()];
  var legacyHeaderMap = {
    "Customer Name": ["Booking By", "Shipper Name"],
    "Customer Email ID": ["Your Email Address", "Consignee Email ID", "Shipper Email"],
    "Shipper Phone No": ["Shipper Phone"],
    "Consignee Phone No": ["Consignee Phone"],
    "Delivery Address": ["Consignee Address"],
    "Pickup Address": ["Shipper Address"],
    "Preferred Pickup Date": ["Preferred Delivery Date", "Booking Date"],
    "Pickup Date": ["Flight Date", "Preferred Pickup Date"],
    "Actual Weight": ["Weight in KG"],
    "Approver Status": ["Sahib Khan Status"]
  };
  for (var r = 1; r < data.length; r++) {
    var nextRow = new Array(BOOKING_HEADERS.length).fill("");
    BOOKING_HEADERS.forEach(function(header, newIdx) {
      var oldIdx = currentHeaders.indexOf(header);
      if (oldIdx >= 0) nextRow[newIdx] = data[r][oldIdx];
      if (oldIdx < 0 && legacyHeaderMap[header]) {
        for (var m = 0; m < legacyHeaderMap[header].length; m++) {
          var legacyIdx = currentHeaders.indexOf(legacyHeaderMap[header][m]);
          if (legacyIdx >= 0 && data[r][legacyIdx] !== "") {
            nextRow[newIdx] = data[r][legacyIdx];
            break;
          }
        }
      }
    });
    rebuilt.push(nextRow);
  }

  range.clearContent();
  sh.getRange(1, 1, rebuilt.length, BOOKING_HEADERS.length).setValues(rebuilt);
}

function getOrCreateUsersSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(USERS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(USERS_SHEET);
    sh.appendRow(USER_HEADERS);
    sh.getRange(1, 1, 1, USER_HEADERS.length)
      .setFontWeight("bold")
      .setBackground("#111827")
      .setFontColor("#ffffff")
      .setHorizontalAlignment("center")
      .setWrap(false);
  }

  ensureDefaultUsers_(sh);

  sh.setColumnWidth(1, 230);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 180);

  return sh;
}

function ensureDefaultUsers_(sh) {
  var rows = sh.getDataRange().getValues();
  var existing = {};
  for (var i = 1; i < rows.length; i++) {
    existing[String(rows[i][0] || "").trim().toLowerCase()] = true;
  }

  DEFAULT_USERS.forEach(function(user) {
    var email = String(user[0] || "").trim().toLowerCase();
    if (!existing[email]) {
      sh.appendRow(user);
      existing[email] = true;
    }
  });
}

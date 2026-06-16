var BOOKING_SHEET = "WorkflowBookings";
var USERS_SHEET = "WorkflowUsers";

var BOOKING_HEADERS = [
  "Record ID", "Created At", "Last Updated",
  "Booking By", "Your Email Address", "Booking Date",
  "Shipper Name", "Shipper Email", "Shipper Phone", "Shipper Address", "Shipper KYC Type",
  "Consignee Name", "Consignee Email ID", "Consignee Phone", "Consignee Address", "Consignee KYC Type",
  "Preferred Delivery Date",
  "Dispatch Date", "MAWB No", "India Status",
  "Flight Date", "Airline Type", "Expected Arrival Date Time", "Sahib Khan Status",
  "Arrival Date", "Custom Date", "Dispatch To Consignee Date", "CS Status"
];

var USER_HEADERS = ["Email", "Password", "Role", "Name"];
var DEFAULT_USERS = [
  ["india@quickshipnow.com", "india123", "india", "India Office"],
  ["sahib@quickshipnow.com", "sahib123", "sahib", "Sahib Khan"],
  ["cs@quickshipnow.com", "cs123", "cs", "CS Team"]
];

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  if (params.action) {
    return handleApiGet_(params);
  }

  var page = String(params.page || "public").toLowerCase();
  var recordId = String(params.recordId || "").trim();
  var validPages = ["public", "login", "dashboard", "form"];
  if (validPages.indexOf(page) === -1) page = "public";

  var recordData = {};
  if (page === "form" && recordId) {
    var rec = getOfficeRecord(recordId);
    if (rec && rec.found) recordData = rec.data;
  }

  try {
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
  } catch (err) {
    return ContentService
      .createTextOutput("Quickship Booking Workflow API is running.")
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

function handleApiGet_(params) {
  var callback = String(params.callback || "").trim();
  var action = String(params.action || "").trim();
  var args = [];

  try {
    if (params.args) {
      args = JSON.parse(String(params.args));
      if (!Array.isArray(args)) args = [];
    }

    var result;
    if (action === "submitPublicForm") {
      result = submitPublicForm(args[0] || {});
    } else if (action === "loginOffice") {
      result = loginOffice(args[0] || "", args[1] || "");
    } else if (action === "getOfficeDashboard") {
      result = getOfficeDashboard();
    } else if (action === "getOfficeRecord") {
      result = getOfficeRecord(args[0] || "");
    } else if (action === "saveOfficeSection") {
      result = saveOfficeSection(args[0] || {});
    } else {
      result = { success: false, error: "Unknown API action." };
    }

    return apiResponse_(result, callback);
  } catch (err) {
    return apiResponse_({ success: false, error: err.message }, callback);
  }
}

function apiResponse_(payload, callback) {
  var json = JSON.stringify(payload || {});
  if (callback && /^[a-zA-Z_$][0-9a-zA-Z_$]*(\.[a-zA-Z_$][0-9a-zA-Z_$]*)*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
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
      if (rowEmail === em && rowPass === pass && (role === "india" || role === "sahib" || role === "cs")) {
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
      "Booking By", "Your Email Address", "Booking Date",
      "Shipper Name", "Shipper Email", "Shipper Phone", "Shipper Address", "Shipper KYC Type",
      "Consignee Name", "Consignee Email ID", "Consignee Phone", "Consignee Address", "Consignee KYC Type",
      "Preferred Delivery Date"
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

    row[bci("India Status")] = "Pending";
    row[bci("Sahib Khan Status")] = "Pending";
    row[bci("CS Status")] = "Pending";

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

    if (section !== "india" && section !== "sahib" && section !== "cs") {
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
      sh.getRange(rowNum, bci("Dispatch Date") + 1).setValue(String(fields["Dispatch Date"] || "").trim());
      sh.getRange(rowNum, bci("MAWB No") + 1).setValue(String(fields["MAWB No"] || "").trim());
      sh.getRange(rowNum, bci("India Status") + 1).setValue("Submitted");
    } else if (section === "sahib") {
      sh.getRange(rowNum, bci("Flight Date") + 1).setValue(String(fields["Flight Date"] || "").trim());
      sh.getRange(rowNum, bci("Airline Type") + 1).setValue(String(fields["Airline Type"] || "").trim());
      sh.getRange(rowNum, bci("Expected Arrival Date Time") + 1).setValue(String(fields["Expected Arrival Date Time"] || "").trim());
      sh.getRange(rowNum, bci("Sahib Khan Status") + 1).setValue("Submitted");
    } else {
      sh.getRange(rowNum, bci("Arrival Date") + 1).setValue(String(fields["Arrival Date"] || "").trim());
      sh.getRange(rowNum, bci("Custom Date") + 1).setValue(String(fields["Custom Date"] || "").trim());
      sh.getRange(rowNum, bci("Dispatch To Consignee Date") + 1).setValue(String(fields["Dispatch To Consignee Date"] || "").trim());
      sh.getRange(rowNum, bci("CS Status") + 1).setValue("Submitted");
    }

    sh.getRange(rowNum, bci("Last Updated") + 1).setValue(formatNow_());

    return { success: true, recordId: rid };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function sendSubmissionEmail_(fields, recordId, submittedAt) {
  var to = String(fields["Your Email Address"] || "").trim();
  if (!to) return { sent: false, error: "Recipient email is missing." };

  var subject = "Booking Submitted - " + recordId;
  var lines = [
    "Booking submitted successfully.",
    "",
    "Record ID: " + recordId,
    "Submitted At: " + submittedAt,
    "",
    "Form Details:",
    "Booking By: " + String(fields["Booking By"] || ""),
    "Your Email Address: " + to,
    "Booking Date: " + String(fields["Booking Date"] || ""),
    "Shipper Name: " + String(fields["Shipper Name"] || ""),
    "Shipper Email: " + String(fields["Shipper Email"] || ""),
    "Shipper Phone: " + String(fields["Shipper Phone"] || ""),
    "Shipper Address: " + String(fields["Shipper Address"] || ""),
    "Shipper KYC Type: " + String(fields["Shipper KYC Type"] || ""),
    "Consignee Name: " + String(fields["Consignee Name"] || ""),
    "Consignee Email ID: " + String(fields["Consignee Email ID"] || ""),
    "Consignee Phone: " + String(fields["Consignee Phone"] || ""),
    "Consignee Address: " + String(fields["Consignee Address"] || ""),
    "Consignee KYC Type: " + String(fields["Consignee KYC Type"] || ""),
    "Preferred Delivery Date: " + String(fields["Preferred Delivery Date"] || "")
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
  for (var r = 1; r < data.length; r++) {
    var nextRow = new Array(BOOKING_HEADERS.length).fill("");
    BOOKING_HEADERS.forEach(function(header, newIdx) {
      var oldIdx = currentHeaders.indexOf(header);
      if (oldIdx >= 0) nextRow[newIdx] = data[r][oldIdx];
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

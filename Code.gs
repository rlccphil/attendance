/**
 * RLCC Attendance — Apps Script back end (JSON API)
 *
 * The front end now lives on GitHub Pages (index.html + checkin.html).
 * This script only reads/writes the Google Sheet and returns JSON.
 *
 * Deploy: Deploy > New deployment > Web app
 *   Execute as: Me      Who has access: Anyone
 *
 * Settings sheet keys (column A = key, column B = value):
 *   REFRESH_INTERVAL, ATTENDANCE_TAB, APP_TITLE, ADMIN_EMAIL, ADMIN_NAME,
 *   TOTAL_ENROLLED, SHOW_LIVE_LIST, FORM_FIELD_1, FORM_FIELD_2, ...
 *   ADMIN_PIN  (optional — if filled in, the staff page asks for it)
 */

// ---------- Helpers ----------

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml(value) {
  return String(value === null || value === undefined ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatTime(ss, value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, ss.getSpreadsheetTimeZone(), "h:mm a");
  }
  return String(value);
}

// ---------- Settings ----------

function getSheetSettings(ss) {
  var settings = {
    interval: 15,
    tabName: "Attendance",
    appTitle: "Attendance System",
    adminEmail: "",
    adminName: "",
    adminPin: "",
    totalEnrolled: 0,
    showLiveList: true,
    formFields: ["Name"]
  };

  var settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet) {
    var data = settingsSheet.getDataRange().getValues();
    var tempFields = [];

    for (var i = 0; i < data.length; i++) {
      var key = String(data[i][0]).trim();
      var value = String(data[i][1]).trim();

      if (key === "REFRESH_INTERVAL") settings.interval = parseInt(value, 10) || 15;
      else if (key === "ATTENDANCE_TAB" && value) settings.tabName = value;
      else if (key === "APP_TITLE" && value) settings.appTitle = value;
      else if (key === "ADMIN_EMAIL" && value) settings.adminEmail = value;
      else if (key === "ADMIN_NAME" && value) settings.adminName = value;
      else if (key === "ADMIN_PIN" && value) settings.adminPin = value;
      else if (key === "TOTAL_ENROLLED") settings.totalEnrolled = parseInt(value, 10) || 0;
      else if (key === "SHOW_LIVE_LIST") {
        settings.showLiveList = !(value.toUpperCase() === "NO" || value.toUpperCase() === "FALSE");
      }
      else if (key.indexOf("FORM_FIELD_") === 0 && value !== "") {
        tempFields.push(value);
      }
    }

    if (tempFields.length > 0) settings.formFields = tempFields;
  }
  return settings;
}

function getTargetSheet(ss) {
  var settings = getSheetSettings(ss);
  return ss.getSheetByName(settings.tabName) || ss.getSheetByName("Log") || ss.getSheets()[0];
}

function ensureSheetHeaders(sheet, fields) {
  var expectedHeaders = ["Timestamp"].concat(fields).concat(["Session ID"]);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(expectedHeaders);
    sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold").setBackground("#edf2f7");
  } else {
    var currentHeaders = sheet.getRange(1, 1, 1, expectedHeaders.length).getValues()[0];
    var isMatching = true;
    for (var i = 0; i < expectedHeaders.length; i++) {
      if (currentHeaders[i] !== expectedHeaders[i]) {
        isMatching = false;
        break;
      }
    }
    if (!isMatching) {
      sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
      sheet.getRange(1, 1, 1, expectedHeaders.length).setFontWeight("bold").setBackground("#edf2f7");
    }
  }
}

function formatSessionDate(sessionId) {
  try {
    var rawDateStr = sessionId.replace("SESSION_", "").slice(0, 8);
    var year = rawDateStr.slice(0, 4);
    var month = rawDateStr.slice(4, 6);
    var day = rawDateStr.slice(6, 8);
    var months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
    var monthName = months[parseInt(month, 10) - 1];
    if (year && monthName && day) return parseInt(day, 10) + " " + monthName + " " + year;
  } catch (e) {}
  return sessionId;
}

// Staff actions need the PIN only when ADMIN_PIN is set in Settings.
function isStaffAuthorized(e, settings) {
  if (!settings.adminPin) return true;
  return String(e.parameter.pin || "") === settings.adminPin;
}

function getSessionAttendees(ss, settings, targetSessionId) {
  var sheet = getTargetSheet(ss);
  var data = sheet.getDataRange().getValues();
  var fields = settings.formFields;
  var sessionColIndex = fields.length + 1;
  var attendees = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][sessionColIndex] === targetSessionId) {
      var attendeeObj = { time: data[i][0] };
      for (var f = 0; f < fields.length; f++) {
        attendeeObj[fields[f]] = data[i][f + 1];
      }
      attendees.push(attendeeObj);
    }
  }
  return attendees;
}

// ---------- Router ----------

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || "";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var props = PropertiesService.getScriptProperties();
  var settings = getSheetSettings(ss);

  // Public: what the check-in page and staff page need to draw themselves
  if (action === "config") {
    return jsonOut({
      status: "success",
      appTitle: settings.appTitle,
      formFields: settings.formFields,
      showLiveList: settings.showLiveList,
      interval: settings.interval,
      pinRequired: !!settings.adminPin
    });
  }

  // Public: members submit their attendance (guarded by the rotating QR id)
  if (action === "submit") {
    var submittedSessionId = e.parameter.sessionId;
    var submittedQrId = e.parameter.qrId;
    var activeQrId = props.getProperty("ACTIVE_QR_ID");

    if (!activeQrId || submittedQrId !== activeQrId) {
      return jsonOut({ status: "expired", message: "Attendance is closed or this QR code has expired. Please scan the current code." });
    }

    var sheet = getTargetSheet(ss);
    ensureSheetHeaders(sheet, settings.formFields);

    var rowData = [new Date()];
    settings.formFields.forEach(function (field) {
      rowData.push(e.parameter[field] || "");
    });
    rowData.push(submittedSessionId);

    sheet.appendRow(rowData);
    return jsonOut({ status: "success" });
  }

  // Everything below is for staff only
  if (["auth", "newSession", "stopSession", "getAttendees", "sendReport"].indexOf(action) !== -1) {
    if (!isStaffAuthorized(e, settings)) {
      return jsonOut({ status: "denied", message: "Incorrect PIN." });
    }
  }

  if (action === "auth") {
    return jsonOut({ status: "success" });
  }

  if (action === "newSession") {
    var newQrId = "QR_" + Math.random().toString(36).substring(2, 9).toUpperCase();
    props.setProperty("ACTIVE_QR_ID", newQrId);

    var tgt = getTargetSheet(ss);
    ensureSheetHeaders(tgt, settings.formFields);

    return jsonOut({ status: "success", qrId: newQrId, interval: settings.interval });
  }

  if (action === "stopSession") {
    props.deleteProperty("ACTIVE_QR_ID");
    return jsonOut({ status: "stopped" });
  }

  if (action === "getAttendees") {
    var attendees = getSessionAttendees(ss, settings, e.parameter.sessionId);
    var list = attendees.map(function (a) {
      var copy = {};
      for (var k in a) copy[k] = a[k];
      copy.time = formatTime(ss, a.time);
      return copy;
    });
    return jsonOut({
      status: "success",
      count: list.length,
      fields: settings.formFields,
      attendees: list
    });
  }

  if (action === "sendReport") {
    var targetSessionId = e.parameter.sessionId;
    if (!settings.adminEmail) return jsonOut({ status: "error", message: "No admin email set" });

    var fields = settings.formFields;
    var formattedDate = formatSessionDate(targetSessionId);
    var rows = getSessionAttendees(ss, settings, targetSessionId);
    var presentCount = rows.length;

    var boxTitle, boxValue, boxSubtext, subjectTag;
    if (settings.totalEnrolled > 0) {
      var pct = Math.round((presentCount / settings.totalEnrolled) * 100);
      boxTitle = "Attendance Rate";
      boxValue = pct + "%";
      boxSubtext = presentCount + " out of " + settings.totalEnrolled + " present";
      subjectTag = "(" + boxValue + ")";
    } else {
      boxTitle = "Total Present";
      boxValue = presentCount;
      boxSubtext = "Total checked-in attendees";
      subjectTag = "(" + presentCount + " Present)";
    }

    var tableHeaders = '<th style="padding: 12px; border-bottom: 2px solid #ddd;">Time</th>';
    fields.forEach(function (f) {
      tableHeaders += '<th style="padding: 12px; border-bottom: 2px solid #ddd;">' + escapeHtml(f) + "</th>";
    });

    var htmlBody =
      '<div style="font-family: Arial, sans-serif; background-color: #f3f6f2; padding: 30px; color: #1d2a1c;">' +
        '<div style="max-width: 600px; margin: auto; background: #fff; padding: 30px; border-radius: 10px; border-top: 6px solid #365233;">' +
          '<h2 style="color: #365233; text-align: center; margin-bottom: 5px;">' + escapeHtml(settings.appTitle) + ' Report</h2>' +
          '<p style="text-align: center; color: #6b7a69; margin-top: 0;">Date: <strong>' + escapeHtml(formattedDate) + '</strong></p>' +
          '<div style="text-align: center; margin: 30px 0;">' +
            '<div style="display: inline-block; background: #e6ede4; border: 2px solid #365233; border-radius: 12px; padding: 20px 40px;">' +
              '<div style="font-size: 14px; color: #365233; text-transform: uppercase; font-weight: bold; margin-bottom: 5px;">' + boxTitle + '</div>' +
              '<div style="font-size: 42px; font-weight: bold; color: #365233; margin: 0;">' + boxValue + '</div>' +
              '<div style="font-size: 13px; color: #6b7a69; margin-top: 5px;">' + boxSubtext + '</div>' +
            '</div>' +
          '</div>' +
          '<h3 style="color: #365233; border-bottom: 2px solid #e6ede4; padding-bottom: 10px;">Attendee List</h3>' +
          '<table style="width: 100%; border-collapse: collapse; margin-top: 15px;">' +
            '<thead><tr style="background-color: #f3f6f2; color: #555; text-align: left;">' + tableHeaders + '</tr></thead>' +
            '<tbody>';

    if (rows.length === 0) {
      htmlBody += '<tr><td colspan="' + (fields.length + 1) + '" style="padding: 15px; text-align: center; color: #999;">No attendees logged for this session.</td></tr>';
    } else {
      rows.forEach(function (row) {
        htmlBody += '<tr><td style="padding: 10px; border-bottom: 1px solid #eee;">' + escapeHtml(formatTime(ss, row.time)) + "</td>";
        fields.forEach(function (f) {
          htmlBody += '<td style="padding: 10px; border-bottom: 1px solid #eee;">' + escapeHtml(row[f]) + "</td>";
        });
        htmlBody += "</tr>";
      });
    }

    htmlBody += "</tbody></table></div></div>";

    var mailOptions = {
      to: settings.adminEmail,
      subject: settings.appTitle + " - Report for " + formattedDate + " " + subjectTag,
      htmlBody: htmlBody
    };
    if (settings.adminName) mailOptions.name = settings.adminName;

    MailApp.sendEmail(mailOptions);
    return jsonOut({ status: "success" });
  }

  // Default: simple health check (the app itself is hosted on GitHub Pages)
  return jsonOut({ status: "ok", app: settings.appTitle, message: "RLCC Attendance API is running." });
}

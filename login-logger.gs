// =====================================================
// EduMongol — Нэвтрэлт + Бүртгэл (Google Apps Script)
// Spreadsheet: 1-Xim2280wnxTyBh13WWahsR6VbMUWquyNOVOm0HhTtQ
// Extensions > Apps Script > Code.gs дотор хуулна
// Дахин Deploy хийхээ мартуузай!
// =====================================================

var SPREADSHEET_ID = "1-Xim2280wnxTyBh13WWahsR6VbMUWquyNOVOm0HhTtQ";

function doGet(e) {
  try {
    var action = e.parameter.action || "login";

    if (action === "register") {
      return handleRegister(e);
    } else {
      return handleLogin(e);
    }

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", msg: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── НЭВТРЭЛТ бүртгэх ──────────────────────────────
function handleLogin(e) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, "Нэвтрэлт",
    ["Огноо, цаг", "Имэйл", "Нэр", "Эрх", "Хөтөч / Төхөөрөмж"],
    ["#2e7d4f", 160, 200, 110, 100, 260]);

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Ulaanbaatar", "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    timestamp,
    e.parameter.email || "—",
    e.parameter.name  || "—",
    e.parameter.role  || "—",
    e.parameter.agent || "—"
  ]);
  styleLastRow(sheet, 5);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", time: timestamp }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── БҮРТГЭЛ хадгалах ──────────────────────────────
function handleRegister(e) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = getOrCreateSheet(ss, "Бүртгэл",
    ["Бүртгэлийн огноо", "Нэр", "Имэйл", "Нууц үг", "Хөтөч / Төхөөрөмж"],
    ["#1a5276", 160, 130, 200, 120, 260]);

  var now       = new Date();
  var timestamp = Utilities.formatDate(now, "Asia/Ulaanbaatar", "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    timestamp,
    e.parameter.name     || "—",
    e.parameter.email    || "—",
    e.parameter.password || "—",
    e.parameter.agent    || "—"
  ]);
  styleLastRow(sheet, 5);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", time: timestamp }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Тусламжийн функцүүд ───────────────────────────
function getOrCreateSheet(ss, name, headers, colors) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setFontWeight("bold");
    hRange.setBackground(colors[0]);
    hRange.setFontColor("#ffffff");
    hRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    for (var i = 0; i < headers.length; i++) {
      sheet.setColumnWidth(i + 1, colors[i + 1]);
    }
  }
  return sheet;
}

function styleLastRow(sheet, colCount) {
  var lastRow = sheet.getLastRow();
  if (lastRow % 2 === 0) {
    sheet.getRange(lastRow, 1, 1, colCount).setBackground("#eaf4fb");
  }
}

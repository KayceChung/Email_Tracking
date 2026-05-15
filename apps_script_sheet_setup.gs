/**
 * Gmail Email Logger - Google Sheets Structure Setup
 *
 * Huong dan:
 * 1) Mo Google Sheet muc tieu
 * 2) Extensions > Apps Script
 * 3) Dan toan bo file nay vao Code.gs
 * 4) Chay ham setupEmailLoggerSheets()
 */

var TARGET_SPREADSHEET_ID = "10DpSr-N4jOU5lhNp-8m0F_0KH5iIKBuTlsAZ3wOtNq0";

function setupEmailLoggerSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  setupEmailLoggerSheetsCore_(ss);
}

/**
 * Chay ham nay neu ban muon setup truc tiep theo Spreadsheet ID.
 */
function setupEmailLoggerSheetsById() {
  var ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  setupEmailLoggerSheetsCore_(ss);
}

function setupEmailLoggerSheetsCore_(ss) {
  var bookingHeaders = [
    "NỀN_TẢNG",
    "PHÂN_LOẠI",
    "SUBJECT",
    "BODY",
    "LINK",
    "DATE",
    "WEEKNUM",
    "DATE_TYPE",
    "ID_ĐH",
    "EMAIL_PHẢN_HỒI",
    "P.I",
    "P.I_TIMES",
    "LATEST_UPDATE_BY",
    "LATEST_UPDATES",
    "Change_counter",
    "STATUS_TIME_CHANGE",
    "DURATION",
    "CHUYẾN",
    "TRẠNG_THÁI_ĐH",
    "Message_id",
    "ĐẢM_NHẬN_HỖ_TRỢ",
    "ĐIỂM ĐI",
    "ĐIỂM ĐÓN",
    "NỘI_DUNG_CONFIRM",
    "THỜI_GIAN_ĐẢM_NHẬN H_TRỢ",
    "CONFIRM_TỪ_NHÂN_VIÊN"
  ];

  var confirmedCancelHeaders = [
    "NỀN_TẢNG",
    "PHÂN_LOẠI",
    "SUBJECT",
    "BODY",
    "LINK",
    "DATE",
    "WEEKNUM",
    "DATE_TYPE",
    "ID_ĐH",
    "P.I",
    "P.I_TIME",
    "LATEST_UPDATE_BY",
    "LATEST_UPDATES",
    "CONFIRM_TIME"
  ];

  var sheetSchemas = [
    { name: "EMAIL_BOOKING", headers: bookingHeaders, cols: bookingHeaders.length },
    { name: "EMAIL_XÁC_NHẬN", headers: confirmedCancelHeaders, cols: confirmedCancelHeaders.length },
    { name: "EMAIL_HỦY", headers: confirmedCancelHeaders, cols: confirmedCancelHeaders.length },
    { name: "EMAIL_CTRIP", headers: confirmedCancelHeaders, cols: confirmedCancelHeaders.length },
    { name: "EMAIL_SEATOS", headers: confirmedCancelHeaders, cols: confirmedCancelHeaders.length },
    { name: "EMAIL_REVIEW", headers: confirmedCancelHeaders, cols: confirmedCancelHeaders.length },
    { name: "EMAIL_KLOOK_SPECIAL", headers: confirmedCancelHeaders, cols: confirmedCancelHeaders.length }
  ];

  sheetSchemas.forEach(function (schema) {
    withRetry_("upsert " + schema.name, function () {
      upsertSheetWithHeaders_(ss, schema.name, schema.headers, schema.cols);
    });
  });

  SpreadsheetApp.flush();
  Logger.log("Done: Da tao/cap nhat toan bo sheet va cot cho: " + ss.getId());
}

/**
 * Optional: xoa du lieu (giu lai header) de reset nhanh.
 */
function clearEmailLoggerDataKeepHeaders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var names = [
    "EMAIL_BOOKING",
    "EMAIL_XÁC_NHẬN",
    "EMAIL_HỦY",
    "EMAIL_CTRIP",
    "EMAIL_SEATOS",
    "EMAIL_REVIEW",
    "EMAIL_KLOOK_SPECIAL"
  ];

  names.forEach(function (name) {
    var ws = ss.getSheetByName(name);
    if (!ws) return;

    var lastRow = ws.getLastRow();
    if (lastRow > 1) {
      ws.getRange(2, 1, lastRow - 1, ws.getMaxColumns()).clearContent();
    }
  });

  SpreadsheetApp.flush();
  Logger.log("Done: Da xoa data, giu nguyen header.");
}

function upsertSheetWithHeaders_(ss, sheetName, headers, requiredCols) {
  var ws = withRetry_("getSheetByName " + sheetName, function () {
    return ss.getSheetByName(sheetName);
  });
  var isNewSheet = false;
  if (!ws) {
    ws = withRetry_("insertSheet " + sheetName, function () {
      return ss.insertSheet(sheetName);
    });
    isNewSheet = true;
  }

  // Dam bao so cot toi thieu.
  if (ws.getMaxColumns() < requiredCols) {
    withRetry_("insertColumnsAfter " + sheetName, function () {
      ws.insertColumnsAfter(ws.getMaxColumns(), requiredCols - ws.getMaxColumns());
    });
  }

  // Chi ghi header khi sheet moi duoc tao, tranh anh huong AppSheet.
  if (isNewSheet) {
    withRetry_("set header " + sheetName, function () {
      ws.getRange(1, 1, 1, headers.length).setValues([headers]);
    });

    // Chi style cho sheet moi tao de tranh timeout tren file lon.
    withRetry_("style header " + sheetName, function () {
      ws.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#e8f0fe")
        .setHorizontalAlignment("center")
        .setWrap(false);
    });

    withRetry_("set frozen rows " + sheetName, function () {
      ws.setFrozenRows(1);
    });

    // Dat width cot quan trong, bo autoResize vi de timeout tren workbook lon.
    withRetry_("set column widths " + sheetName, function () {
      ws.setColumnWidth(4, 420); // BODY
      ws.setColumnWidth(3, 320); // SUBJECT
      ws.setColumnWidth(5, 260); // LINK
    });
  }
}

function withRetry_(name, fn) {
  var maxAttempts = 4;
  var baseSleepMs = 1200;
  var lastErr = null;

  for (var i = 1; i <= maxAttempts; i++) {
    try {
      return fn();
    } catch (err) {
      lastErr = err;
      var msg = String(err);
      var isRetryable =
        msg.indexOf("timed out") !== -1 ||
        msg.indexOf("Service Spreadsheets") !== -1 ||
        msg.indexOf("try again") !== -1;

      if (!isRetryable || i === maxAttempts) {
        throw err;
      }

      var sleepMs = baseSleepMs * Math.pow(2, i - 1);
      Logger.log("Retry " + i + "/" + maxAttempts + " for " + name + " after error: " + msg);
      Utilities.sleep(sleepMs);
    }
  }

  throw lastErr;
}

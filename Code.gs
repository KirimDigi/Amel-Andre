// Code.gs - Google Apps Script untuk RSVP Amel & Andre
// Spreadsheet: https://docs.google.com/spreadsheets/d/179cusuqNesQBSMpjrlZfeB0mznAd0aehxGIwt2cBnYc/edit
// Sheet: Sheet1 | Kolom: timestamp | nama tamu | ucapan | konfirmasi kehadiran | jumlah tamu
// Deploy: Publish -> Deploy as web app -> Anyone, even anonymous -> salin URL exec

var SPREADSHEET_ID = "179cusuqNesQBSMpjrlZfeB0mznAd0aehxGIwt2cBnYc";
var SHEET_NAME = "Sheet1";
var HEADER = ["timestamp", "nama tamu", "ucapan", "konfirmasi kehadiran", "jumlah tamu"];

function getSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  // pastikan header
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADER);
    sheet.getRange(1,1,1,HEADER.length).setFontWeight("bold").setBackground("#a58f6e").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  } else {
    var firstRow = sheet.getRange(1,1,1,HEADER.length).getValues()[0].join("|").toLowerCase();
    if (firstRow.indexOf("timestamp") === -1) {
      sheet.insertRowBefore(1);
      sheet.getRange(1,1,1,HEADER.length).setValues([HEADER]);
      sheet.getRange(1,1,1,HEADER.length).setFontWeight("bold").setBackground("#a58f6e").setFontColor("#ffffff");
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) ? e.parameter.action.toLowerCase() : "read";
    if (action === "read" || action === "get" || action === "") {
      var sheet = getSheet_();
      var lastRow = sheet.getLastRow();
      if (lastRow <= 1) {
        return jsonResponse_({ status: "success", data: [] });
      }
      var values = sheet.getRange(2, 1, lastRow - 1, HEADER.length).getValues();
      var data = [];
      for (var i = values.length - 1; i >= 0; i--) { // terbaru dulu (DESC)
        var r = values[i];
        // r[0]=timestamp, r[1]=nama, r[2]=ucapan, r[3]=kehadiran, r[4]=jumlah tamu
        if (!r[1] && !r[2]) continue;
        data.push({
          timestamp: r[0] ? Utilities.formatDate(new Date(r[0]), Session.getScriptTimeZone() || "Asia/Jakarta", "dd/MM/yyyy HH:mm") : "",
          rawTimestamp: r[0] ? new Date(r[0]).toISOString() : "",
          nama: String(r[1] || "").trim(),
          ucapan: String(r[2] || "").trim(),
          kehadiran: String(r[3] || "").trim(),
          jumlahTamu: String(r[4] || "").trim()
        });
      }
      return jsonResponse_({ status: "success", data: data, count: data.length });
    }
    return jsonResponse_({ status: "error", message: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

function doPost(e) {
  try {
    var payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        // fallback form-urlencoded
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    var nama = String(payload.nama || payload.name || payload.author || "").trim();
    var ucapan = String(payload.ucapan || payload.comment || payload.message || "").trim();
    var kehadiranRaw = String(payload.kehadiran || payload.attendance || payload.konfirmasi || "").trim().toLowerCase();
    var jumlahTamu = String(payload.jumlahTamu || payload.guest || payload.jumlah || "").trim();

    // validasi
    if (nama.length < 2) return jsonResponse_({ status: "error", message: "Nama minimal 2 karakter" });
    if (ucapan.length < 2) return jsonResponse_({ status: "error", message: "Ucapan minimal 2 karakter" });
    if (!kehadiranRaw) return jsonResponse_({ status: "error", message: "Konfirmasi kehadiran wajib" });

    // normalisasi kehadiran
    var kehadiranMap = {
      "present": "Hadir",
      "hadir": "Hadir",
      "notpresent": "Tidak Hadir",
      "tidak hadir": "Tidak Hadir",
      "tidakhadir": "Tidak Hadir",
      "notsure": "Masih Ragu",
      "masih ragu": "Masih Ragu",
      "ragu": "Masih Ragu"
    };
    var kehadiran = kehadiranMap[kehadiranRaw] || payload.kehadiran || kehadiranRaw;

    if (kehadiran === "Hadir") {
      if (!jumlahTamu) jumlahTamu = "1";
    } else {
      // jika tidak hadir / ragu, jumlah tamu boleh kosong, tapi simpan "-"
      if (!jumlahTamu) jumlahTamu = "-";
    }

    var timestamp = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy HH:mm:ss");
    var sheet = getSheet_();
    sheet.appendRow([timestamp, nama, ucapan, kehadiran, jumlahTamu]);

    return jsonResponse_({ status: "success", message: "Terima kasih atas ucapan Anda!" });
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

// handle CORS preflight jika diperlukan
function doOptions(e) {
  return jsonResponse_({ status: "success" });
}

function jsonResponse_(obj) {
  var out = ContentService.createTextOutput(JSON.stringify(obj));
  out.setMimeType(ContentService.MimeType.JSON);
  return out;
}

// Untuk test di editor GAS: Jalankan testRead / testWrite
function testRead() {
  Logger.log(doGet({ parameter: { action: "read" } }).getContent());
}
function testWrite() {
  var res = doPost({ postData: { contents: JSON.stringify({ nama: "Tes Tamu", ucapan: "Selamat ya Amel & Andre!", kehadiran: "Hadir", jumlahTamu: "2" }) } });
  Logger.log(res.getContent());
}

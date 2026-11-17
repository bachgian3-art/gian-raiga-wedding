const SHEET_NAME = "RSVP";

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Nama", "Attendance", "Message", "Emoji"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const callback = (e && e.parameter && e.parameter.callback) || "";
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const comments = rows.slice(1).filter(row => row[1]).map(row => ({
    date: row[0] instanceof Date
      ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "dd MMM yyyy, HH:mm")
      : String(row[0] || ""),
    name: String(row[1] || ""),
    attendance: String(row[2] || ""),
    message: String(row[3] || ""),
    emoji: String(row[4] || "❤️")
  }));

  const result = JSON.stringify({ ok: true, comments: comments });
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + "(" + result + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const data = JSON.parse(raw);
    const name = String(data.name || "").trim().slice(0, 60);
    const attendance = String(data.attendance || "").trim().slice(0, 40);
    const message = String(data.message || "").trim().slice(0, 500);
    const allowedAttendance = ["Hadir", "Tidak dapat hadir", "Masih tentatif"];
    const allowedEmoji = ["❤️", "🥰", "💐", "🎉", "✨", "🤍"];
    const emoji = allowedEmoji.includes(data.emoji) ? data.emoji : "❤️";

    if (!name || !allowedAttendance.includes(attendance) || !message) {
      throw new Error("Data RSVP tidak lengkap atau tidak valid.");
    }

    getSheet_().appendRow([new Date(), name, attendance, message, emoji]);
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

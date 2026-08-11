const SHEET_NAME = 'Pendaftaran';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
    }

    const data = e.parameter || {};
    const fieldOrder = data._fieldOrder ? JSON.parse(data._fieldOrder) : [];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Waktu Submit', ...fieldOrder]);
    }

    const row = [new Date(), ...fieldOrder.map((key) => data[key] || '')];
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: 'Data berhasil disimpan'
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

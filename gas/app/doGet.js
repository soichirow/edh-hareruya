const DB_SHEET_NAME = 'データベース';
const LOG_SHEET_NAME = 'アクセスログ';

/**
 * JSON APIエンドポイント
 * GitHub Pagesから ?format=json&limit=3000 で呼ばれる
 * アクセスログをスプレッドシートに記録
 */
function doGet(e) {
  logAccess_();
  const limit = (e && e.parameter && e.parameter.limit) || 3000;
  const json = fetchDatabaseJson(limit);
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

/** アクセスログを記録 */
function logAccess_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.appendRow(['日時']);
    }
    sheet.appendRow([new Date()]);
  } catch (err) {
    console.log('logAccess_ error: ' + err);
  }
}

/** シート→JSON文字列（Map→Objectに変換） */
function fetchDatabaseJson(limit) {
  const max = Math.max(1, Math.min(Number(limit || 1000), 3000));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = new Sheet(ss.getSheetByName(DB_SHEET_NAME));
  const dicts = sheet.getAsDicts();
  const rows = dicts.slice(0, max).map(m => Object.fromEntries(m));
  return JSON.stringify(rows);
}

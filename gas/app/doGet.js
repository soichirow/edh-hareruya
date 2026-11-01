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

/**
 * アクセス数を1日1行に集計して記録する。
 * 匿名のGETごとに行を足すとシートが際限なく増えるので、行数は日数までに抑える。
 */
function logAccess_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(LOG_SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(LOG_SHEET_NAME);
      sheet.appendRow(['日時']);
    }
    // シートが日付に自動変換しないよう、接尾辞付きの文字列をキーにする。
    const JST_OFFSET_MS = 9 * 60 * 60 * 1000;
    const today = new Date(Date.now() + JST_OFFSET_MS).toISOString().slice(0, 10) + ' JST';
    const lastRow = sheet.getLastRow();
    if (lastRow > 1 && sheet.getRange(lastRow, 1).getValue() === today) {
      const count = sheet.getRange(lastRow, 2);
      count.setValue(Number(count.getValue()) + 1);
    } else {
      sheet.appendRow([today, 1]);
    }
  } catch (err) {
    console.log('logAccess_ error: ' + err);
  }
}

/** シート→JSON文字列 */
function fetchDatabaseJson(limit) {
  const max = Math.max(1, Math.min(Number(limit || 1000), 3000));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const [headers, ...values] = ss.getSheetByName(DB_SHEET_NAME).getDataRange().getValues();
  const rows = values.slice(0, max).map(row =>
    Object.fromEntries(headers.map((header, i) => [header, row[i]]))
  );
  return JSON.stringify(rows);
}

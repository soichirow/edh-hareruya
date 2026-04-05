const DB_SHEET_NAME = 'データベース';
const LOG_SHEET_NAME = 'アクセスログ';
const SEARCH_LOG_SHEET = '検索ログ';

/** HTMLまたはJSON APIを返す */
function doGet(e) {
  const format = (e && e.parameter && e.parameter.format) || '';

  // ?format=json → JSON API（GitHub Pages等の外部サイトから利用）
  if (format === 'json') {
    const limit = e.parameter.limit || 3000;
    const json = fetchDatabaseJson(limit);
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  // デフォルト: HTML（従来のWebアプリ）
  logAccess_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('EDHオタクカード一覧')
    .addMetaTag('viewport', 'width=device-width,initial-scale=1')
    .setFaviconUrl("https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/72x72/1f913.png");
}

/** シート→JSON文字列（Map→Objectに変換） */
function fetchDatabaseJson(limit) {
  const max = Math.max(1, Math.min(Number(limit || 1000), 3000));
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = new Sheet(ss.getSheetByName(DB_SHEET_NAME)); // ← あなたの class Sheet
  const dicts = sheet.getAsDicts();                           // Map[]
  const rows = dicts.slice(0, max).map(m => Object.fromEntries(m));
  return JSON.stringify(rows);
}

function logAccess_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
    sheet.appendRow(['日時']);
  }
  sheet.appendRow([new Date()]);
}



/** 検索ログを最小情報で1行追記（非同期呼び出し想定） */
function logSearch(term, meta = {}) {
  try {
    term = String(term || '').trim();
    if (!term || term.length < 2) return;        // 短すぎ・空は捨てる
    if (term.length > 100) term = term.slice(0, 100); // 長すぎ防止

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SEARCH_LOG_SHEET);
    if (!sh) {
      sh = ss.insertSheet(SEARCH_LOG_SHEET);
      sh.appendRow(['日時', '検索語', 'プレゼンター', '色(選択)', '備考']); // ヘッダ
      sh.setFrozenRows(1);
    }

    const row = [
      new Date(),
      term,
      meta.presenter || '',
      (meta.colors || []).join('') || '',        // 例: "WU" や "R"
      meta.note || ''
    ];
    sh.appendRow(row);
  } catch (err) {
    // UIに影響させない
    console.log('logSearch error: ' + err);
  }
}

/**
 * Scryfall APIからカードデータを取得
 * スプレッドシートの「データベース」シートに書き込む
 */

/**
 * スプレッドシートのカード名からScryfall APIでデータ取得・書き込み
 * L列（12列目）のカード名を読み、R列以降（18列目〜）にデータを書き込む
 * 既にR列にデータがあるカードはスキップ（中断後の再実行でレジューム可能）
 *
 * 中断ガード:
 * - 連続3カード失敗で中断（consecutive-failures）
 * - 開始から5分超過で中断（time-budget）
 * 中断しても、R列が埋まった行はスキップされるため再実行すれば続きから処理される
 *
 * @return {{processed: number, succeeded: number, skipped: number, failed: number, aborted: (string|null)}} 実行サマリ
 */
function fetchMtgCardDataJa() {
  const TIME_BUDGET_MS = 5 * 60 * 1000;
  const MAX_CONSECUTIVE_FAILURES = 3;

  const summary = { processed: 0, succeeded: 0, skipped: 0, failed: 0, aborted: null };
  const start = Date.now();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('データベース');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return summary;

  const sourceRows = sheet.getRange(2, 12, lastRow - 1, 8).getValues();

  const fetchOptions = {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'MtgSheetFetcher/1.0 (contact: sowatanabe@bushiroad-gp.com)',
      'Accept': 'application/json',
    },
  };

  let consecutiveFailures = 0;

  for (let i = 0; i < sourceRows.length; i++) {
    // 時間予算ガード: 5分超過で中断（GASの6分実行制限到達前に安全に止める）
    if (Date.now() - start > TIME_BUDGET_MS) {
      summary.aborted = 'time-budget';
      console.log('時間予算（5分）を超過したため中断します。再実行すれば続きから処理されます。');
      break;
    }

    const row = i + 2;
    const cardNameRaw = (sourceRows[i][0] || '').toString().trim();
    if (!cardNameRaw) {
      summary.skipped++;
      continue;
    }

    // 紹介カード名と取得済みの日英名が一致する行だけスキップする。
    const normalizedName = cardNameRaw.toLowerCase();
    const already = sourceRows[i].slice(6, 8).some(function(name) {
      return name && name.toString().trim().toLowerCase() === normalizedName;
    });
    if (already) {
      summary.skipped++;
      continue;
    }

    summary.processed++;
    try {
      const card = findCardPreferJa(cardNameRaw, fetchOptions);
      const imageCard = findImageCardPreferEn(card, fetchOptions);

      const nameJa = card.printed_name || '';
      const nameEn = card.name || '';

      const manaCost = getJoined(card, 'mana_cost', 'mana_cost');
      const typeLine = getJoined(card, 'type_line', 'printed_type_line');
      const oracleText = getJoined(card, 'oracle_text', 'printed_text');

      const power = getJoined(card, 'power', 'power');
      const toughness = getJoined(card, 'toughness', 'toughness');

      const colors = (card.colors || []).join(',');
      const colorIdentity = (card.color_identity || []).join(',');
      const imageUri = getImageNormal(imageCard) || getImageNormal(card);
      const scryfallUri = card.scryfall_uri || '';
      const cmc = (card.cmc !== null && card.cmc !== undefined) ? card.cmc : '';

      // API待機中に行が追加・移動された場合、別カードの行へ書き込まない。
      const currentCardName = (sheet.getRange(row, 12).getValue() || '').toString().trim();
      if (currentCardName !== cardNameRaw) {
        summary.aborted = 'source-rows-changed';
        console.error('紹介カードの行が処理中に変更されたため中断します。再実行してください。');
        break;
      }

      sheet.getRange(row, 18, 1, 12).setValues([[
        nameJa, nameEn, manaCost, typeLine, oracleText,
        power, toughness, colors, colorIdentity, imageUri, scryfallUri, cmc
      ]]);

      summary.succeeded++;
      consecutiveFailures = 0;
      Utilities.sleep(150);

    } catch (e) {
      console.error('カード取得失敗: ' + cardNameRaw + ' - ' + (e && e.message ? e.message : e));
      summary.failed++;
      consecutiveFailures++;
      Utilities.sleep(1000);

      // 連続失敗ブレーカー: 3カード連続失敗で中断（レート制限・障害時のカスケード防止）
      if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
        summary.aborted = 'consecutive-failures';
        console.error('3カード連続で失敗したため中断します。時間をおいて再実行してください（処理済み行はスキップされます）。');
        break;
      }
    }
  }

  console.log(
    '完了サマリ: 処理 ' + summary.processed +
    ' / 成功 ' + summary.succeeded +
    ' / スキップ ' + summary.skipped +
    ' / 失敗 ' + summary.failed +
    (summary.aborted ? ' / 中断理由: ' + summary.aborted : '')
  );
  return summary;
}

/**
 * 日本語印刷を優先して1枚取得。なければ曖昧検索でフォールバック
 * @param {string} cardName - カード名
 * @param {Object} fetchOptions - UrlFetchAppのオプション
 * @return {Object} Scryfallカードオブジェクト
 */
function findCardPreferJa(cardName, fetchOptions) {
  const escaped = cardName.replace(/"/g, '\\"');
  const qJa = '!"' + escaped + '" lang:ja';
  const urlJa = 'https://api.scryfall.com/cards/search?q=' + encodeURIComponent(qJa) + '&unique=prints&order=released&dir=desc';
  const ja = fetchJsonWithRetry_(urlJa, fetchOptions);

  if (ja && ja.object === 'list' && Array.isArray(ja.data) && ja.data.length > 0) {
    return ja.data[0];
  }

  const urlFuzzy = 'https://api.scryfall.com/cards/named?fuzzy=' + encodeURIComponent(cardName);
  const fuzzy = fetchJsonWithRetry_(urlFuzzy, fetchOptions);

  if (fuzzy && fuzzy.object === 'card') return fuzzy;

  const detail = (ja && ja.details) ? ja.details : (fuzzy && fuzzy.details) ? fuzzy.details : 'not found';
  throw new Error(detail);
}

/**
 * JSON取得（リトライ付き、指数バックオフ）
 * @param {string} url - リクエストURL
 * @param {Object} fetchOptions - UrlFetchAppのオプション
 * @param {number} [maxRetries=5] - 最大リトライ回数
 * @return {Object} パース済みJSON
 */
function fetchJsonWithRetry_(url, fetchOptions, maxRetries) {
  const retries = (maxRetries === null || maxRetries === undefined) ? 5 : maxRetries;

  for (let attempt = 0; attempt <= retries; attempt++) {
    const res = UrlFetchApp.fetch(url, fetchOptions);
    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code >= 200 && code < 300) return JSON.parse(text);

    if (code === 429 && attempt < retries) {
      const waitMs = Math.min(8000, 500 * Math.pow(2, attempt));
      Utilities.sleep(waitMs);
      continue;
    }

    let payload;
    try { payload = JSON.parse(text); } catch (_e) { payload = { details: text }; }
    const msg = payload && payload.details ? payload.details : 'HTTP ' + code;
    throw new Error(msg);
  }

  throw new Error('retry exceeded');
}

/**
 * 多面カードは face ごとに結合して返す（日本語優先、なければ英語）
 * @param {Object} card - Scryfallカードオブジェクト
 * @param {string} baseKey - 英語フィールド名
 * @param {string} printedKey - 日本語フィールド名
 * @return {string}
 */
function getJoined(card, baseKey, printedKey) {
  if (Array.isArray(card.card_faces) && card.card_faces.length > 0) {
    const parts = card.card_faces.map(function(f) { return (f[printedKey] || f[baseKey] || '').toString(); });
    return parts.join(' // ');
  }
  return (card[printedKey] || card[baseKey] || '').toString();
}

/**
 * カード画像URL取得（normal サイズ）
 * @param {Object} card - Scryfallカードオブジェクト
 * @return {string} 画像URL
 */
function getImageNormal(card) {
  if (card.image_uris && card.image_uris.normal) return card.image_uris.normal;
  if (Array.isArray(card.card_faces) && card.card_faces[0] && card.card_faces[0].image_uris && card.card_faces[0].image_uris.normal) {
    return card.card_faces[0].image_uris.normal;
  }
  return '';
}

function findImageCardPreferEn(card, fetchOptions) {
  const name = (card && card.name ? card.name : '').toString().trim();
  if (name) {
    const url = 'https://api.scryfall.com/cards/named?exact=' + encodeURIComponent(name);
    try {
      const en = fetchJsonWithRetry_(url, fetchOptions);
      if (en && en.object === 'card' && en.lang !== 'ja' && getImageNormal(en)) return en;
    } catch (e) {
      console.error('English image lookup fallback: ' + name + ' - ' + (e && e.message ? e.message : e));
    }
  }
  return card || {};
}

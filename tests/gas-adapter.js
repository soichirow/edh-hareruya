/**
 * GASコードをテスト用にimportするアダプター
 * GASのグローバル変数をモックに差し替えて関数を返す
 */

// ===== 純粋関数（カードデータGet.js） =====

// getJoined: 多面カードのテキスト結合
export function getJoined(card, baseKey, printedKey) {
  if (Array.isArray(card.card_faces) && card.card_faces.length > 0) {
    const parts = card.card_faces.map(f => (f[printedKey] || f[baseKey] || '').toString());
    return parts.join(' // ');
  }
  return (card[printedKey] || card[baseKey] || '').toString();
}

// getImageNormal: カード画像URL取得
export function getImageNormal(card) {
  if (card.image_uris && card.image_uris.normal) return card.image_uris.normal;
  if (Array.isArray(card.card_faces) && card.card_faces[0] && card.card_faces[0].image_uris && card.card_faces[0].image_uris.normal) {
    return card.card_faces[0].image_uris.normal;
  }
  return '';
}

export function findImageCardPreferEn(card, fetchOptions, fetchJsonWithRetry_) {
  const name = (card && card.name ? card.name : '').toString().trim();
  if (name && typeof fetchJsonWithRetry_ === 'function') {
    const url = 'https://api.scryfall.com/cards/named?exact=' + encodeURIComponent(name);
    try {
      const en = fetchJsonWithRetry_(url, fetchOptions);
      if (en && en.object === 'card' && en.lang !== 'ja' && getImageNormal(en)) return en;
    } catch {
      // Fall back to the matched metadata card.
    }
  }
  return card || {};
}

// convertISO8601ToTime: ISO8601 → HH:mm:ss
export function convertISO8601ToTime(duration) {
  const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return '';

  const hours = parseInt(match[1] || 0, 10);
  const minutes = parseInt(match[2] || 0, 10);
  const seconds = parseInt(match[3] || 0, 10);

  const hh = hours > 0 ? String(hours).padStart(2, '0') : '00';
  const mm = String(minutes).padStart(2, '0');
  const ss = String(seconds).padStart(2, '0');

  return `${hh}:${mm}:${ss}`;
}

// ===== YouTube動画フィルタ（純粋関数） =====
export function isOtakuCardVideo(item) {
  const title = item.snippet?.title || '';
  if (!title.includes('オタクカード')) return false;
  if (item.snippet?.liveBroadcastContent && item.snippet.liveBroadcastContent !== 'none') return false;
  return true;
}

export function extractVideoData(item, Utilities) {
  const title = item.snippet.title;
  const videoId = item.id;
  const publishedAtUTC = item.snippet.publishedAt;
  const publishedDateJST = Utilities.formatDate(new Date(publishedAtUTC), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const url = 'https://www.youtube.com/watch?v=' + videoId;
  const thumbnailUrl = item.snippet.thumbnails?.high?.url || item.snippet.thumbnails?.default?.url || '';
  const durationISO = item.contentDetails?.duration || '';
  const duration = convertISO8601ToTime(durationISO);
  return [title, videoId, publishedAtUTC, publishedDateJST, url, thumbnailUrl, duration];
}

// ===== 最新動画の差分更新 + データベース連携 =====
export function createUpdateLatestVideos(deps) {
  const { SpreadsheetApp, YouTube, Utilities, Logger } = deps;
  const CHANNEL_ID = 'UC1l7GtlvAmCOXRlxjImbWvw';
  const SHEET_NAME = '動画自動取得';
  const HEADER = ['タイトル', '動画ID', '公開日(UTC)', '公開日(JST)', 'URL', 'サムネイルURL', '再生時間'];
  const VIDEO_ID_COL = 1; // 0-indexed: 動画ID は2列目

  function updateLatestVideos() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    const isNew = !sheet;
    if (isNew) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADER);
    }

    // 既存の動画IDを収集
    const existingIds = new Set();
    if (!isNew) {
      const lastRow = sheet.getLastRow();
      if (lastRow >= 2) {
        const ids = sheet.getRange(2, VIDEO_ID_COL + 1, lastRow - 1, 1).getValues();
        ids.forEach(function(r) { if (r[0]) existingIds.add(String(r[0])); });
      }
    }

    // YouTubeから最新50件取得
    const channelResponse = YouTube.Channels.list('contentDetails', { id: CHANNEL_ID });
    const uploadsPlaylistId = channelResponse.items[0].contentDetails.relatedPlaylists.uploads;

    const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
      playlistId: uploadsPlaylistId, maxResults: 50, pageToken: '',
    });

    if (!playlistResponse.items) return 0;

    const videoIds = playlistResponse.items.map(function(i) { return i.contentDetails.videoId; }).join(',');
    const videoDetails = YouTube.Videos.list('contentDetails,snippet,liveStreamingDetails', { id: videoIds });

    const newRows = [];
    videoDetails.items.forEach(function(item) {
      if (!isOtakuCardVideo(item)) return;
      if (existingIds.has(String(item.id))) return;
      newRows.push(extractVideoData(item, Utilities));
    });

    if (newRows.length > 0) {
      // 動画自動取得シートに挿入（上部）
      sheet.insertRowsAfter(1, newRows.length);
      sheet.getRange(2, 1, newRows.length, HEADER.length).setValues(newRows);

      // データベースシートにも4行×新動画数を挿入
      const dbSheet = ss.getSheetByName('データベース');
      if (dbSheet) {
        const ROWS_PER_VIDEO = 4;
        const totalDbRows = newRows.length * ROWS_PER_VIDEO;
        dbSheet.insertRowsAfter(1, totalDbRows);

        const dbData = [];
        for (let v = 0; v < newRows.length; v++) {
          const videoRow = newRows[v];
          // [タイトル, 動画ID, 公開日UTC, 公開日JST, URL, サムネイルURL, 再生時間]
          const title = videoRow[0];
          const episode = extractEpisodeNumber(title);

          for (let r = 1; r <= ROWS_PER_VIDEO; r++) {
            // 29列: 動画タイトル,ID,公開日UTC,公開日JST,動画URL,サムネイルURL,再生時間,話数,紹介順, 以降空欄
            const row = new Array(29).fill('');
            row[0] = videoRow[0]; // 動画タイトル
            row[1] = videoRow[1]; // ID
            row[2] = videoRow[2]; // 公開日UTC
            row[3] = videoRow[3]; // 公開日JST
            row[4] = videoRow[4]; // 動画URL
            row[5] = videoRow[5]; // サムネイルURL
            row[6] = videoRow[6]; // 再生時間
            row[7] = episode;     // 話数
            row[8] = r;           // 紹介順
            dbData.push(row);
          }
        }
        dbSheet.getRange(2, 1, totalDbRows, 29).setValues(dbData);
      }
    }

    if (Logger) Logger.log('新規追加: ' + newRows.length + '件');
    return newRows.length;
  }

  return { updateLatestVideos };
}

/** タイトルから話数を抽出（例: "EDHオタクカード184" → 184） */
export function extractEpisodeNumber(title) {
  const match = title.match(/オタクカード(\d+)/);
  return match ? parseInt(match[1], 10) : '';
}

// ===== fetchJsonWithRetry_: リトライ付きfetch =====

/** 429時の待機ms: Retry-Afterヘッダー優先（cap 60s）、なければ指数バックオフ 1000*2^attempt（cap 30s） */
function resolveRateLimitWaitMs(res, attempt) {
  if (res && typeof res.getAllHeaders === 'function') {
    const headers = res.getAllHeaders() || {};
    const raw = headers['Retry-After'] !== undefined ? headers['Retry-After'] : headers['retry-after'];
    if (raw !== undefined && raw !== null && raw !== '') {
      const sec = Number(raw);
      if (isFinite(sec) && sec >= 0) return Math.min(60000, sec * 1000);
    }
  }
  return Math.min(30000, 1000 * Math.pow(2, attempt));
}

/** 5xx / fetch例外時の待機ms: 500*2^attempt（cap 8s） */
function resolveServerErrorWaitMs(attempt) {
  return Math.min(8000, 500 * Math.pow(2, attempt));
}

export function createFetchJsonWithRetry(UrlFetchApp, Utilities) {
  return function fetchJsonWithRetry_(url, fetchOptions, maxRetries) {
    const retries = (maxRetries === null || maxRetries === undefined) ? 5 : maxRetries;

    for (let attempt = 0; attempt <= retries; attempt++) {
      let res;
      try {
        res = UrlFetchApp.fetch(url, fetchOptions);
      } catch (e) {
        // GASのネットワーク/DNS/タイムアウト例外もリトライ対象
        if (attempt < retries) {
          Utilities.sleep(resolveServerErrorWaitMs(attempt));
          continue;
        }
        throw e;
      }

      const code = res.getResponseCode();
      const text = res.getContentText();

      if (code >= 200 && code < 300) return JSON.parse(text);

      if (code === 429 && attempt < retries) {
        Utilities.sleep(resolveRateLimitWaitMs(res, attempt));
        continue;
      }

      if (code >= 500 && attempt < retries) {
        Utilities.sleep(resolveServerErrorWaitMs(attempt));
        continue;
      }

      // リトライ対象外の4xx、またはリトライ上限到達
      let payload;
      try { payload = JSON.parse(text); } catch { payload = { details: text }; }
      const msg = payload && payload.details ? payload.details : `HTTP ${code}`;
      throw new Error(msg);
    }

    throw new Error('retry exceeded');
  };
}

// ===== fetchMtgCardDataJa: Scryfallカードデータ一括取得（中断ガード付き） =====
export function createFetchMtgCardDataJa(deps) {
  const { SpreadsheetApp, UrlFetchApp, Utilities } = deps;
  const now = deps.now || Date.now;
  const fetchJsonWithRetry_ = createFetchJsonWithRetry(UrlFetchApp, Utilities);

  const TIME_BUDGET_MS = 5 * 60 * 1000;
  const MAX_CONSECUTIVE_FAILURES = 3;

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

  return function fetchMtgCardDataJa() {
    const summary = { processed: 0, succeeded: 0, skipped: 0, failed: 0, aborted: null };
    const start = now();

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
      // 時間予算ガード: 5分超過で中断（残り行は未記入のまま再実行でレジューム可能）
      if (now() - start > TIME_BUDGET_MS) {
        summary.aborted = 'time-budget';
        break;
      }

      const row = i + 2;
      const cardNameRaw = (sourceRows[i][0] || '').toString().trim();
      if (!cardNameRaw) {
        summary.skipped++;
        continue;
      }

      const normalizedName = cardNameRaw.toLowerCase();
      const already = sourceRows[i].slice(6, 8).some(name =>
        name && name.toString().trim().toLowerCase() === normalizedName
      );
      if (already) {
        summary.skipped++;
        continue;
      }

      summary.processed++;
      try {
        const card = findCardPreferJa(cardNameRaw, fetchOptions);
        const imageCard = findImageCardPreferEn(card, fetchOptions, fetchJsonWithRetry_);

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

        const currentCardName = (sheet.getRange(row, 12).getValue() || '').toString().trim();
        if (currentCardName !== cardNameRaw) {
          summary.aborted = 'source-rows-changed';
          break;
        }

        sheet.getRange(row, 18, 1, 12).setValues([[
          nameJa, nameEn, manaCost, typeLine, oracleText,
          power, toughness, colors, colorIdentity, imageUri, scryfallUri, cmc,
        ]]);

        summary.succeeded++;
        consecutiveFailures = 0;
        Utilities.sleep(150);
      } catch {
        summary.failed++;
        consecutiveFailures++;
        Utilities.sleep(1000);

        // 連続失敗ブレーカー: 3カード連続失敗で中断
        if (consecutiveFailures >= MAX_CONSECUTIVE_FAILURES) {
          summary.aborted = 'consecutive-failures';
          break;
        }
      }
    }

    return summary;
  };
}

// ===== doGet: JSON APIエンドポイント =====
export function createDoGet(deps) {
  const { SpreadsheetApp, ContentService, DB_SHEET_NAME } = deps;

  function fetchDatabaseJson(limit) {
    const max = Math.max(1, Math.min(Number(limit || 1000), 3000));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const [headers, ...values] = ss.getSheetByName(DB_SHEET_NAME).getDataRange().getValues();
    const rows = values.slice(0, max).map(row =>
      Object.fromEntries(headers.map((header, i) => [header, row[i]]))
    );
    return JSON.stringify(rows);
  }

  function doGet(e) {
    const limit = (e && e.parameter && e.parameter.limit) || 3000;
    const json = fetchDatabaseJson(limit);
    return ContentService
      .createTextOutput(json)
      .setMimeType(ContentService.MimeType.JSON);
  }

  return { doGet, fetchDatabaseJson };
}

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

// ===== YouTube動画取得（モック対応） =====
export function createFetchOtakuVideos(deps) {
  const { SpreadsheetApp, YouTube, Utilities, Logger } = deps;
  const CHANNEL_ID = 'UC1l7GtlvAmCOXRlxjImbWvw';
  const HEADER = ['タイトル', '動画ID', '公開日(UTC)', '公開日(JST)', 'URL', 'サムネイルURL', '再生時間'];

  return function fetchOtakuVideos(sheetName, allPages) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clearContents(); }
    sheet.appendRow(HEADER);

    const channelResponse = YouTube.Channels.list('contentDetails', { id: CHANNEL_ID });
    const uploadsPlaylistId = channelResponse.items[0].contentDetails.relatedPlaylists.uploads;

    const videoData = [];
    let nextPageToken = '';

    do {
      const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
        playlistId: uploadsPlaylistId, maxResults: 50, pageToken: nextPageToken,
      });
      if (!playlistResponse.items) break;

      const videoIds = playlistResponse.items.map(i => i.contentDetails.videoId).join(',');
      const videoDetails = YouTube.Videos.list('contentDetails,snippet,liveStreamingDetails', { id: videoIds });

      videoDetails.items.forEach(item => {
        if (isOtakuCardVideo(item)) {
          videoData.push(extractVideoData(item, Utilities));
        }
      });

      nextPageToken = allPages ? playlistResponse.nextPageToken : '';
    } while (nextPageToken);

    if (videoData.length > 0) {
      sheet.getRange(2, 1, videoData.length, 7).setValues(videoData);
    }
    if (Logger) Logger.log('「オタクカード」を含む動画 ' + videoData.length + '件を「' + sheetName + '」に出力しました。');
    return videoData.length;
  };
}

// ===== 最新動画の差分更新 =====
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
export function createFetchJsonWithRetry(UrlFetchApp, Utilities) {
  return function fetchJsonWithRetry_(url, fetchOptions, maxRetries) {
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
      const msg = payload && payload.details ? payload.details : `HTTP ${code}`;
      throw new Error(msg);
    }

    throw new Error('retry exceeded');
  };
}

// ===== doGet: JSON APIエンドポイント =====
export function createDoGet(deps) {
  const { SpreadsheetApp, ContentService, Sheet, DB_SHEET_NAME } = deps;

  function fetchDatabaseJson(limit) {
    const max = Math.max(1, Math.min(Number(limit || 1000), 3000));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = new Sheet(ss.getSheetByName(DB_SHEET_NAME));
    const dicts = sheet.getAsDicts();
    const rows = dicts.slice(0, max).map(m => Object.fromEntries(m));
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

// ===== Sheet class (テスト用エクスポート) =====
export function createSheetClass(SpreadsheetApp) {
  class Sheet {
    constructor(sheet = SpreadsheetApp.getActiveSheet(), headerRows = 1, headerIndex = headerRows - 1) {
      this.sheet = sheet;
      this.headerRows = headerRows;
      this.headerIndex = headerIndex;
    }

    getDataRange() { return this.sheet.getDataRange(); }
    getRange(...args) { return this.sheet.getRange(...args); }
    getLastRow() { return this.sheet.getLastRow(); }
    getLastColumn() { return this.sheet.getLastColumn(); }
    getName() { return this.sheet.getName(); }

    getDataRangeValues() {
      if (this.dataRangeValues_ !== undefined) return this.dataRangeValues_;
      this.dataRangeValues_ = this.getDataRange().getValues();
      return this.dataRangeValues_;
    }

    getHeaders() {
      if (this.headers_ !== undefined) return this.headers_;
      const headerValues = this.getHeaderValues();
      this.headers_ = headerValues[this.headerIndex];
      return this.headers_;
    }

    getHeaderValues() {
      if (this.headerValues_ !== undefined) return this.headerValues_;
      const values = this.getDataRangeValues();
      this.headerValues_ = values.filter((_, i) => i < this.headerRows);
      return this.headerValues_;
    }

    getDataValues() {
      if (this.dataValues_ !== undefined) return this.dataValues_;
      const values = this.getDataRangeValues();
      this.dataValues_ = values.filter((_, i) => i >= this.headerRows);
      return this.dataValues_;
    }

    getColumnByHeaderName(headerName) {
      return this.getColumnIndexByHeaderName(headerName) + 1;
    }

    getColumnIndexByHeaderName(headerName) {
      const headers = this.getHeaders();
      const columnIndex = headers.indexOf(headerName);
      if (columnIndex === -1) throw new Error('The value "' + headerName + '" does not exist in the header row of sheet "' + this.getName() + '".');
      return columnIndex;
    }

    getAsDicts() {
      if (this.dicts_ !== undefined) return this.dicts_;
      const headers = this.getHeaders();
      const values = this.getDataValues();
      this.dicts_ = values.map((record) =>
        record.reduce((acc, cur, j) => acc.set(headers[j], cur), new Map())
      );
      return this.dicts_;
    }

    getFieldValues(headerName, isAddHeader = false) {
      return this.select([headerName], isAddHeader).flat();
    }

    select(headerNames, isAddHeaders = false) {
      const dicts = this.getAsDicts();
      const records = dicts.map(dict => headerNames.map(key => dict.get(key)));
      return isAddHeaders ? [headerNames, ...records] : records;
    }

    hasValueInField(headerName, value) {
      return this.getFieldValues(headerName).includes(value);
    }

    filterDicts(headerName, value, isSameValue = true) {
      const dicts = this.getAsDicts();
      return isSameValue
        ? dicts.filter(dict => dict.get(headerName) === value)
        : dicts.filter(dict => dict.get(headerName) !== value);
    }

    findDict(headerName, value) {
      const dicts = this.getAsDicts();
      const dict = dicts.find(dict => dict.get(headerName) === value);
      if (dict === undefined) throw new Error('The value "' + value + '" does not exist in the "' + headerName + '" column of sheet ' + this.getName() + '.');
      return dict;
    }

    findDictIndex(headerName, dict) {
      const dicts = this.getAsDicts();
      return dicts.findIndex(record => record.get(headerName) === dict.get(headerName));
    }

    appendRows(values) {
      if (values.length === 0) return;
      this.getRange(this.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
      return this;
    }

    appendDicts(dicts) {
      const headerNames = this.getHeaders();
      const records = dicts.map(dict => headerNames.map(key => dict.get(key)));
      this.appendRows(records);
      return this;
    }
  }

  return Sheet;
}

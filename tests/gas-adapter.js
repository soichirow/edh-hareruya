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

// ===== doGet: JSON/HTMLモード分岐 =====
export function createDoGet(deps) {
  const { SpreadsheetApp, ContentService, HtmlService, Sheet, logAccess_, DB_SHEET_NAME } = deps;

  function fetchDatabaseJson(limit) {
    const max = Math.max(1, Math.min(Number(limit || 1000), 3000));
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = new Sheet(ss.getSheetByName(DB_SHEET_NAME));
    const dicts = sheet.getAsDicts();
    const rows = dicts.slice(0, max).map(m => Object.fromEntries(m));
    return JSON.stringify(rows);
  }

  function doGet(e) {
    const format = (e && e.parameter && e.parameter.format) || '';

    if (format === 'json') {
      const limit = e.parameter.limit || 3000;
      const json = fetchDatabaseJson(limit);
      return ContentService
        .createTextOutput(json)
        .setMimeType(ContentService.MimeType.JSON);
    }

    logAccess_();
    return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('EDHオタクカード一覧')
      .addMetaTag('viewport', 'width=device-width,initial-scale=1')
      .setFaviconUrl('https://cdnjs.cloudflare.com/ajax/libs/twemoji/14.0.2/72x72/1f913.png');
  }

  return { doGet, fetchDatabaseJson };
}

// ===== logSearch =====
export function createLogSearch(SpreadsheetApp, SEARCH_LOG_SHEET) {
  return function logSearch(term, meta = {}) {
    try {
      term = String(term || '').trim();
      if (!term || term.length < 2) return;
      if (term.length > 100) term = term.slice(0, 100);

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sh = ss.getSheetByName(SEARCH_LOG_SHEET);
      if (!sh) {
        sh = ss.insertSheet(SEARCH_LOG_SHEET);
        sh.appendRow(['日時', '検索語', 'プレゼンター', '色(選択)', '備考']);
        sh.setFrozenRows(1);
      }

      const row = [
        new Date(),
        term,
        meta.presenter || '',
        (meta.colors || []).join('') || '',
        meta.note || '',
      ];
      sh.appendRow(row);
    } catch (err) {
      console.log('logSearch error: ' + err);
    }
  };
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

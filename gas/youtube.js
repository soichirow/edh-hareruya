/**
 * YouTube動画取得
 * 晴れる屋MTGチャンネルから「オタクカード」を含む動画を取得
 */

const HARERUYA_CHANNEL_ID = 'UC1l7GtlvAmCOXRlxjImbWvw';
const VIDEO_HEADER = ['タイトル', '動画ID', '公開日(UTC)', '公開日(JST)', 'URL', 'サムネイルURL', '再生時間'];

/**
 * オタクカード動画かどうかを判定
 * @param {Object} item - YouTube Videos.list のアイテム
 * @return {boolean}
 */
function isOtakuCardVideo_(item) {
  const title = item.snippet && item.snippet.title ? item.snippet.title : '';
  if (!title.includes('オタクカード')) return false;
  if (item.snippet.liveBroadcastContent && item.snippet.liveBroadcastContent !== 'none') return false;
  return true;
}

/**
 * 動画アイテムからシート用の行データを生成
 * @param {Object} item - YouTube Videos.list のアイテム
 * @return {Array}
 */
function extractVideoRow_(item) {
  const title = item.snippet.title;
  const videoId = item.id;
  const publishedAtUTC = item.snippet.publishedAt;
  const publishedDateJST = Utilities.formatDate(new Date(publishedAtUTC), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
  const url = 'https://www.youtube.com/watch?v=' + videoId;
  const thumbnailUrl = (item.snippet.thumbnails && item.snippet.thumbnails.high)
    ? item.snippet.thumbnails.high.url
    : ((item.snippet.thumbnails && item.snippet.thumbnails.default) ? item.snippet.thumbnails.default.url : '');
  const durationISO = (item.contentDetails && item.contentDetails.duration) ? item.contentDetails.duration : '';
  const duration = convertISO8601ToTime(durationISO);
  return [title, videoId, publishedAtUTC, publishedDateJST, url, thumbnailUrl, duration];
}

/**
 * ISO8601の再生時間を "HH:mm:ss" に変換
 * @param {string} duration - ISO8601形式（例: "PT1H30M45S"）
 * @return {string} "HH:mm:ss" 形式
 */
function convertISO8601ToTime(duration) {
  const match = duration.match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return '';
  const hours = parseInt(match[1] || 0, 10);
  const minutes = parseInt(match[2] || 0, 10);
  const seconds = parseInt(match[3] || 0, 10);
  const hh = hours > 0 ? String(hours).padStart(2, '0') : '00';
  const mm = String(minutes).padStart(2, '0');
  const ss = String(seconds).padStart(2, '0');
  return hh + ':' + mm + ':' + ss;
}

/**
 * オタクカード動画を取得してシートに書き込む（共通処理）
 * @param {string} sheetName - 出力先シート名
 * @param {boolean} allPages - true: 全ページ取得、false: 最新50件のみ
 * @return {number} 取得件数
 */
function fetchOtakuVideos_(sheetName, allPages) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }
  sheet.appendRow(VIDEO_HEADER);

  const channelResponse = YouTube.Channels.list('contentDetails', { id: HARERUYA_CHANNEL_ID });
  const uploadsPlaylistId = channelResponse.items[0].contentDetails.relatedPlaylists.uploads;

  const videoData = [];
  let nextPageToken = '';

  do {
    const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
      playlistId: uploadsPlaylistId,
      maxResults: 50,
      pageToken: nextPageToken,
    });

    if (!playlistResponse.items) break;

    const videoIds = playlistResponse.items.map(function(item) { return item.contentDetails.videoId; }).join(',');
    const videoDetails = YouTube.Videos.list('contentDetails,snippet,liveStreamingDetails', { id: videoIds });

    videoDetails.items.forEach(function(item) {
      if (isOtakuCardVideo_(item)) {
        videoData.push(extractVideoRow_(item));
      }
    });

    nextPageToken = allPages ? playlistResponse.nextPageToken : '';
  } while (nextPageToken);

  if (videoData.length > 0) {
    sheet.getRange(2, 1, videoData.length, 7).setValues(videoData);
  }

  Logger.log('「オタクカード」を含む動画 ' + videoData.length + '件を「' + sheetName + '」に出力しました。');
  return videoData.length;
}

/**
 * 全オタクカード動画を取得（メニュー: 動画全件取得）
 */
function searchOtakuCardVideosFromHareluya() {
  fetchOtakuVideos_('動画自動取得', true);
}

/**
 * 最新動画の差分更新（メニュー: 最新動画取得）
 * 「動画自動取得」シートに、まだ存在しない動画だけを追加する
 */
function newUpdate() {
  const SHEET_NAME = '動画自動取得';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  const isNew = !sheet;

  if (isNew) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(VIDEO_HEADER);
  }

  // 既存の動画IDを収集（2列目 = 動画ID）
  const existingIds = {};
  if (!isNew) {
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
      for (let i = 0; i < ids.length; i++) {
        if (ids[i][0]) existingIds[String(ids[i][0])] = true;
      }
    }
  }

  // YouTubeから最新50件取得
  const channelResponse = YouTube.Channels.list('contentDetails', { id: HARERUYA_CHANNEL_ID });
  const uploadsPlaylistId = channelResponse.items[0].contentDetails.relatedPlaylists.uploads;

  const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
    playlistId: uploadsPlaylistId,
    maxResults: 50,
    pageToken: '',
  });

  if (!playlistResponse.items) {
    Logger.log('動画が取得できませんでした');
    return;
  }

  const videoIds = playlistResponse.items.map(function(item) { return item.contentDetails.videoId; }).join(',');
  const videoDetails = YouTube.Videos.list('contentDetails,snippet,liveStreamingDetails', { id: videoIds });

  const newRows = [];
  videoDetails.items.forEach(function(item) {
    if (!isOtakuCardVideo_(item)) return;
    if (existingIds[String(item.id)]) return;
    newRows.push(extractVideoRow_(item));
  });

  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, VIDEO_HEADER.length).setValues(newRows);
  }

  Logger.log('最新動画取得: 新規 ' + newRows.length + '件を「' + SHEET_NAME + '」に追加しました');
}

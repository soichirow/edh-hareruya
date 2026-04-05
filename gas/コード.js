function searchOtakuCardVideosFromHareluya() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '動画自動取得';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  const header = ['タイトル', '動画ID', '公開日(UTC)', '公開日(JST)', 'URL', 'サムネイルURL', '再生時間'];
  sheet.appendRow(header);

  const channelId = 'UC1l7GtlvAmCOXRlxjImbWvw'; // 晴れる屋MTG
  let videoData = [];
  let total = 0;

  // チャンネル情報から「アップロード動画リスト」のIDを取得
  const channelResponse = YouTube.Channels.list('contentDetails', { id: channelId });
  const uploadsPlaylistId = channelResponse.items[0].contentDetails.relatedPlaylists.uploads;

  let nextPageToken = '';

  // アップロードリストをページングで全部取得
  do {
    const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
      playlistId: uploadsPlaylistId,
      maxResults: 50,
      pageToken: nextPageToken
    });

    if (!playlistResponse.items) break;

    // videoIdまとめ
    const videoIds = playlistResponse.items.map(item => item.contentDetails.videoId).join(',');
    const videoDetails = YouTube.Videos.list('contentDetails,snippet,liveStreamingDetails', {
      id: videoIds
    });

    videoDetails.items.forEach(item => {
      const title = item.snippet.title;

      // 生放送（ライブ配信・プレミア公開）は除外
      if (item.snippet.liveBroadcastContent && item.snippet.liveBroadcastContent !== 'none') {
        return;
      }

      // 「オタクカード」を含む動画だけ残す
      if (title.includes('オタクカード')) {
        const videoId = item.id;
        const publishedAtUTC = item.snippet.publishedAt;
        const publishedDateJST = Utilities.formatDate(new Date(publishedAtUTC), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
        const url = `https://www.youtube.com/watch?v=${videoId}`;
        const thumbnailUrl = item.snippet.thumbnails?.high?.url || item.snippet.thumbnails?.default?.url;
        const durationISO = item.contentDetails?.duration || '';
        const duration = convertISO8601ToTime(durationISO);

        videoData.push([title, videoId, publishedAtUTC, publishedDateJST, url, thumbnailUrl, duration]);
        total++;
      }
    });

    nextPageToken = playlistResponse.nextPageToken;
  } while (nextPageToken);

  if (videoData.length > 0) {
    sheet.getRange(2, 1, videoData.length, 7).setValues(videoData);
  }

  Logger.log(`「オタクカード」を含む動画 ${total}件を「${sheetName}」に出力しました。`);
}

/**
 * ISO8601の再生時間を "HH:mm:ss" に変換
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

  return `${hh}:${mm}:${ss}`;
}

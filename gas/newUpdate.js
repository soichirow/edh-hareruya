function newUpdate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '最新動画';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }

  const header = ['タイトル', '動画ID', '公開日(UTC)', '公開日(JST)', 'URL', 'サムネイルURL', '再生時間'];
  sheet.appendRow(header);

  const channelId = 'UC1l7GtlvAmCOXRlxjImbWvw'; // 晴れる屋MTG
  const videoData = [];
  let total = 0;

  // チャンネル情報から「アップロード動画リスト」のIDを取得
  const channelResponse = YouTube.Channels.list('contentDetails', { id: channelId });
  const uploadsPlaylistId = channelResponse.items[0].contentDetails.relatedPlaylists.uploads;

  const nextPageToken = '';

  const playlistResponse = YouTube.PlaylistItems.list('snippet,contentDetails', {
    playlistId: uploadsPlaylistId,
    maxResults: 50,
    pageToken: nextPageToken
  });

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

  if (videoData.length > 0) {
    sheet.getRange(2, 1, videoData.length, 7).setValues(videoData);
  }

  Logger.log(`「オタクカード」を含む動画 ${total}件を「${sheetName}」に出力しました。`);
}


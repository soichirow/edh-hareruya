import { describe, it, expect } from 'vitest';
import {
  createMockSheet,
  createMockSpreadsheetApp,
  createMockUrlFetchApp,
  createMockUtilities,
  createMockContentService,
} from './gas-mocks.js';
import {
  convertISO8601ToTime,
  getJoined,
  getImageNormal,
  isOtakuCardVideo,
  extractVideoData,
  createFetchOtakuVideos,
  createFetchJsonWithRetry,
  createDoGet,
  createSheetClass,
} from './gas-adapter.js';

// ========================================
// convertISO8601ToTime（純粋関数）
// ========================================
describe('convertISO8601ToTime', () => {
  it('時間・分・秒すべてある場合', () => {
    expect(convertISO8601ToTime('PT1H30M45S')).toBe('01:30:45');
  });

  it('分・秒のみ', () => {
    expect(convertISO8601ToTime('PT5M30S')).toBe('00:05:30');
  });

  it('秒のみ', () => {
    expect(convertISO8601ToTime('PT45S')).toBe('00:00:45');
  });

  it('時間のみ', () => {
    expect(convertISO8601ToTime('PT2H')).toBe('02:00:00');
  });

  it('不正な文字列は空文字を返す', () => {
    expect(convertISO8601ToTime('')).toBe('');
    expect(convertISO8601ToTime('invalid')).toBe('');
  });

  it('ゼロパディングされる', () => {
    expect(convertISO8601ToTime('PT1M5S')).toBe('00:01:05');
  });
});

// ========================================
// getJoined（純粋関数）
// ========================================
describe('getJoined', () => {
  it('単面カードの日本語テキストを返す', () => {
    const card = { mana_cost: '{2}{W}', printed_name: '太陽の指輪' };
    expect(getJoined(card, 'mana_cost', 'printed_name')).toBe('太陽の指輪');
  });

  it('printed_keyがなければbase_keyを返す', () => {
    const card = { name: 'Sol Ring' };
    expect(getJoined(card, 'name', 'printed_name')).toBe('Sol Ring');
  });

  it('多面カードは // で結合する', () => {
    const card = {
      card_faces: [
        { mana_cost: '{1}', printed_name: '表面' },
        { mana_cost: '{2}', printed_name: '裏面' },
      ],
    };
    expect(getJoined(card, 'mana_cost', 'printed_name')).toBe('表面 // 裏面');
  });

  it('多面カードでprinted_keyがなければbase_keyを使う', () => {
    const card = {
      card_faces: [
        { name: 'Front' },
        { name: 'Back' },
      ],
    };
    expect(getJoined(card, 'name', 'printed_name')).toBe('Front // Back');
  });

  it('空のカードは空文字を返す', () => {
    expect(getJoined({}, 'name', 'printed_name')).toBe('');
  });
});

// ========================================
// getImageNormal（純粋関数）
// ========================================
describe('getImageNormal', () => {
  it('image_uris.normalを返す', () => {
    const card = { image_uris: { normal: 'https://example.com/img.jpg' } };
    expect(getImageNormal(card)).toBe('https://example.com/img.jpg');
  });

  it('多面カードは1面目の画像を返す', () => {
    const card = {
      card_faces: [
        { image_uris: { normal: 'https://example.com/front.jpg' } },
        { image_uris: { normal: 'https://example.com/back.jpg' } },
      ],
    };
    expect(getImageNormal(card)).toBe('https://example.com/front.jpg');
  });

  it('画像がない場合は空文字を返す', () => {
    expect(getImageNormal({})).toBe('');
  });

  it('card_facesがあるがimage_urisがない場合は空文字を返す', () => {
    const card = { card_faces: [{ name: 'Test' }] };
    expect(getImageNormal(card)).toBe('');
  });
});

// ========================================
// fetchJsonWithRetry_（モック付き）
// ========================================
describe('fetchJsonWithRetry_', () => {
  it('200で正常レスポンスを返す', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(200, { data: 'ok' });
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(fn('https://api.example.com', {})).toEqual({ data: 'ok' });
  });

  it('429でリトライして最終的に成功する', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(429, '');
    mockFetch._addResponse(429, '');
    mockFetch._addResponse(200, { data: 'retry-ok' });
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(fn('https://api.example.com', {})).toEqual({ data: 'retry-ok' });
  });

  it('429が最大リトライ回数を超えたらエラー', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    // maxRetries=1 → attempt 0 (429) + attempt 1 (429) → fail
    mockFetch._addResponse(429, '');
    mockFetch._addResponse(429, '{"details":"rate limited"}');
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(() => fn('https://api.example.com', {}, 1)).toThrow('rate limited');
  });

  it('500エラーはリトライせずにスロー', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(500, '{"details":"server error"}');
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(() => fn('https://api.example.com', {})).toThrow('server error');
  });

  it('404エラーのdetailsメッセージを返す', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(404, '{"details":"not found"}');
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(() => fn('https://api.example.com', {})).toThrow('not found');
  });
});

// ========================================
// Sheet クラス
// ========================================
describe('Sheet class', () => {
  const SpreadsheetApp = createMockSpreadsheetApp();
  const SheetClass = createSheetClass(SpreadsheetApp);

  it('ヘッダーを取得できる', () => {
    const mockSheet = createMockSheet(['名前', '色', 'CMC'], [['Sol Ring', '', '1']]);
    const s = new SheetClass(mockSheet);
    expect(s.getHeaders()).toEqual(['名前', '色', 'CMC']);
  });

  it('データ行を取得できる', () => {
    const mockSheet = createMockSheet(['名前', 'CMC'], [['Sol Ring', 1], ['Lightning Bolt', 1]]);
    const s = new SheetClass(mockSheet);
    expect(s.getDataValues()).toHaveLength(2);
    expect(s.getDataValues()[0]).toEqual(['Sol Ring', 1]);
  });

  it('getAsDictsでMap配列を返す', () => {
    const mockSheet = createMockSheet(['名前', 'CMC'], [['Sol Ring', 1]]);
    const s = new SheetClass(mockSheet);
    const dicts = s.getAsDicts();
    expect(dicts).toHaveLength(1);
    expect(dicts[0].get('名前')).toBe('Sol Ring');
    expect(dicts[0].get('CMC')).toBe(1);
  });

  it('getColumnByHeaderNameで列番号を返す', () => {
    const mockSheet = createMockSheet(['A', 'B', 'C'], []);
    const s = new SheetClass(mockSheet);
    expect(s.getColumnByHeaderName('B')).toBe(2);
  });

  it('存在しないヘッダー名でエラー', () => {
    const mockSheet = createMockSheet(['A', 'B'], []);
    const s = new SheetClass(mockSheet);
    expect(() => s.getColumnByHeaderName('Z')).toThrow('does not exist');
  });

  it('filterDictsで条件一致する行を返す', () => {
    const mockSheet = createMockSheet(
      ['名前', '色'],
      [['Sol Ring', ''], ['Lightning Bolt', 'R'], ['Counterspell', 'U']]
    );
    const s = new SheetClass(mockSheet);
    const result = s.filterDicts('色', 'R');
    expect(result).toHaveLength(1);
    expect(result[0].get('名前')).toBe('Lightning Bolt');
  });

  it('findDictで見つからない場合はエラー', () => {
    const mockSheet = createMockSheet(['名前'], [['Sol Ring']]);
    const s = new SheetClass(mockSheet);
    expect(() => s.findDict('名前', 'Missing Card')).toThrow('does not exist');
  });

  it('findDictで見つかった場合はMapを返す', () => {
    const mockSheet = createMockSheet(['名前', 'CMC'], [['Sol Ring', 1]]);
    const s = new SheetClass(mockSheet);
    const dict = s.findDict('名前', 'Sol Ring');
    expect(dict.get('CMC')).toBe(1);
  });

  it('hasValueInFieldで存在チェック', () => {
    const mockSheet = createMockSheet(['名前'], [['Sol Ring'], ['島']]);
    const s = new SheetClass(mockSheet);
    expect(s.hasValueInField('名前', 'Sol Ring')).toBe(true);
    expect(s.hasValueInField('名前', 'Missing')).toBe(false);
  });

  it('selectで指定カラムだけ取得', () => {
    const mockSheet = createMockSheet(['A', 'B', 'C'], [['a1', 'b1', 'c1'], ['a2', 'b2', 'c2']]);
    const s = new SheetClass(mockSheet);
    expect(s.select(['A', 'C'])).toEqual([['a1', 'c1'], ['a2', 'c2']]);
  });

  it('getFieldValuesで単一カラム取得', () => {
    const mockSheet = createMockSheet(['名前', 'CMC'], [['Sol Ring', 1], ['島', 0]]);
    const s = new SheetClass(mockSheet);
    expect(s.getFieldValues('名前')).toEqual(['Sol Ring', '島']);
  });
});

// ========================================
// doGet（モック付き）
// ========================================
describe('doGet (JSON API)', () => {
  function setupDoGet() {
    const dbSheet = createMockSheet(
      ['名前', '色', 'CMC'],
      [['Sol Ring', '', 1], ['Lightning Bolt', 'R', 1]]
    );
    const SpreadsheetApp = createMockSpreadsheetApp({ 'データベース': dbSheet });
    const ContentService = createMockContentService();
    const SheetClass = createSheetClass(SpreadsheetApp);

    const { doGet, fetchDatabaseJson } = createDoGet({
      SpreadsheetApp,
      ContentService,
      Sheet: SheetClass,
      DB_SHEET_NAME: 'データベース',
    });

    return { doGet, fetchDatabaseJson };
  }

  it('JSONレスポンスを返す', () => {
    const { doGet } = setupDoGet();
    const result = doGet({ parameter: {} });
    expect(result._text).toBeDefined();
    const parsed = JSON.parse(result._text);
    expect(parsed).toHaveLength(2);
    expect(parsed[0]['名前']).toBe('Sol Ring');
  });

  it('limitパラメータが効く', () => {
    const { doGet } = setupDoGet();
    const result = doGet({ parameter: { limit: '1' } });
    const parsed = JSON.parse(result._text);
    expect(parsed).toHaveLength(1);
  });

  it('パラメータなしでもJSONを返す', () => {
    const { doGet } = setupDoGet();
    const result = doGet(null);
    const parsed = JSON.parse(result._text);
    expect(parsed).toHaveLength(2);
  });

  it('fetchDatabaseJsonのlimitは1-3000にクランプ', () => {
    const { fetchDatabaseJson } = setupDoGet();
    const r1 = JSON.parse(fetchDatabaseJson(0));
    expect(r1.length).toBeGreaterThanOrEqual(1);
    const r2 = JSON.parse(fetchDatabaseJson(99999));
    expect(r2).toHaveLength(2);
  });
});

// logSearch テストは削除（GAS版WebApp廃止、ログ機能不要）

// ========================================
// isOtakuCardVideo（純粋関数）
// ========================================
describe('isOtakuCardVideo', () => {
  it('タイトルに「オタクカード」を含む通常動画はtrue', () => {
    const item = { snippet: { title: '【MTG】EDHオタクカード184', liveBroadcastContent: 'none' } };
    expect(isOtakuCardVideo(item)).toBe(true);
  });

  it('タイトルに「オタクカード」がなければfalse', () => {
    const item = { snippet: { title: 'MTG開封動画', liveBroadcastContent: 'none' } };
    expect(isOtakuCardVideo(item)).toBe(false);
  });

  it('ライブ配信はfalse', () => {
    const item = { snippet: { title: 'EDHオタクカードLIVE', liveBroadcastContent: 'live' } };
    expect(isOtakuCardVideo(item)).toBe(false);
  });

  it('liveBroadcastContentがnoneなら通す', () => {
    const item = { snippet: { title: 'EDHオタクカード100', liveBroadcastContent: 'none' } };
    expect(isOtakuCardVideo(item)).toBe(true);
  });

  it('liveBroadcastContentが未設定なら通す', () => {
    const item = { snippet: { title: 'EDHオタクカード50' } };
    expect(isOtakuCardVideo(item)).toBe(true);
  });
});

// ========================================
// extractVideoData（純粋関数）
// ========================================
describe('extractVideoData', () => {
  const mockUtilities = createMockUtilities();

  it('動画データを正しい配列形式で返す', () => {
    const item = {
      id: 'abc123',
      snippet: {
        title: 'EDHオタクカード184',
        publishedAt: '2026-03-30T10:00:17Z',
        thumbnails: { high: { url: 'https://i.ytimg.com/vi/abc123/hqdefault.jpg' } },
      },
      contentDetails: { duration: 'PT15M30S' },
    };
    const result = extractVideoData(item, mockUtilities);
    expect(result).toHaveLength(7);
    expect(result[0]).toBe('EDHオタクカード184');
    expect(result[1]).toBe('abc123');
    expect(result[4]).toBe('https://www.youtube.com/watch?v=abc123');
    expect(result[6]).toBe('00:15:30');
  });

  it('サムネイルがない場合は空文字', () => {
    const item = {
      id: 'xyz',
      snippet: { title: 'Test', publishedAt: '2026-01-01T00:00:00Z', thumbnails: {} },
      contentDetails: { duration: 'PT5M' },
    };
    const result = extractVideoData(item, mockUtilities);
    expect(result[5]).toBe('');
  });
});

// ========================================
// createFetchOtakuVideos（モック付き統合テスト）
// ========================================
describe('fetchOtakuVideos', () => {
  function createMockYouTube(videoItems, hasNextPage) {
    return {
      Channels: {
        list: () => ({ items: [{ contentDetails: { relatedPlaylists: { uploads: 'UU123' } } }] }),
      },
      PlaylistItems: {
        list: () => ({
          items: videoItems.map(v => ({ contentDetails: { videoId: v.id }, snippet: {} })),
          nextPageToken: hasNextPage ? 'page2' : undefined,
        }),
      },
      Videos: {
        list: () => ({ items: videoItems }),
      },
    };
  }

  it('オタクカード動画をシートに書き込む', () => {
    const sheets = {};
    const SpreadsheetApp = createMockSpreadsheetApp(sheets);
    const mockItems = [
      { id: 'v1', snippet: { title: 'EDHオタクカード1', publishedAt: '2026-01-01T00:00:00Z', liveBroadcastContent: 'none', thumbnails: {} }, contentDetails: { duration: 'PT10M' } },
      { id: 'v2', snippet: { title: '普通の動画', publishedAt: '2026-01-02T00:00:00Z', liveBroadcastContent: 'none', thumbnails: {} }, contentDetails: { duration: 'PT5M' } },
    ];
    const YouTube = createMockYouTube(mockItems, false);
    const fn = createFetchOtakuVideos({ SpreadsheetApp, YouTube, Utilities: createMockUtilities(), Logger: { log: () => {} } });

    const count = fn('テスト動画', false);
    expect(count).toBe(1); // v1のみ（v2はタイトル不一致）
    expect(sheets['テスト動画']).toBeDefined();
  });

  it('ライブ配信を除外する', () => {
    const sheets = {};
    const SpreadsheetApp = createMockSpreadsheetApp(sheets);
    const mockItems = [
      { id: 'v1', snippet: { title: 'EDHオタクカードLIVE', publishedAt: '2026-01-01T00:00:00Z', liveBroadcastContent: 'live', thumbnails: {} }, contentDetails: { duration: 'PT60M' } },
    ];
    const YouTube = createMockYouTube(mockItems, false);
    const fn = createFetchOtakuVideos({ SpreadsheetApp, YouTube, Utilities: createMockUtilities(), Logger: { log: () => {} } });

    const count = fn('テスト', false);
    expect(count).toBe(0);
  });
});

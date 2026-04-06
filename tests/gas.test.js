import { describe, it, expect, vi } from 'vitest';
import {
  createMockSheet,
  createMockSpreadsheetApp,
  createMockUrlFetchApp,
  createMockUtilities,
  createMockContentService,
  createMockHtmlService,
} from './gas-mocks.js';
import {
  convertISO8601ToTime,
  getJoined,
  getImageNormal,
  createFetchJsonWithRetry,
  createDoGet,
  createLogSearch,
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
describe('doGet', () => {
  function setupDoGet() {
    const dbSheet = createMockSheet(
      ['名前', '色', 'CMC'],
      [['Sol Ring', '', 1], ['Lightning Bolt', 'R', 1]]
    );
    const SpreadsheetApp = createMockSpreadsheetApp({ 'データベース': dbSheet });
    const ContentService = createMockContentService();
    const HtmlService = createMockHtmlService();
    const SheetClass = createSheetClass(SpreadsheetApp);
    const logAccess_ = vi.fn();

    const { doGet, fetchDatabaseJson } = createDoGet({
      SpreadsheetApp,
      ContentService,
      HtmlService,
      Sheet: SheetClass,
      logAccess_,
      DB_SHEET_NAME: 'データベース',
    });

    return { doGet, fetchDatabaseJson, logAccess_, ContentService };
  }

  it('format=jsonでJSON文字列を返す', () => {
    const { doGet } = setupDoGet();
    const result = doGet({ parameter: { format: 'json' } });
    expect(result._text).toBeDefined();
    const parsed = JSON.parse(result._text);
    expect(parsed).toHaveLength(2);
    expect(parsed[0]['名前']).toBe('Sol Ring');
  });

  it('format=jsonでlimitが効く', () => {
    const { doGet } = setupDoGet();
    const result = doGet({ parameter: { format: 'json', limit: '1' } });
    const parsed = JSON.parse(result._text);
    expect(parsed).toHaveLength(1);
  });

  it('format指定なしでHTMLを返す', () => {
    const { doGet, logAccess_ } = setupDoGet();
    const result = doGet({ parameter: {} });
    expect(result._type).toBe('html');
    expect(logAccess_).toHaveBeenCalledOnce();
  });

  it('パラメータなしでもHTMLを返す', () => {
    const { doGet, logAccess_ } = setupDoGet();
    const result = doGet(null);
    expect(result._type).toBe('html');
    expect(logAccess_).toHaveBeenCalled();
  });

  it('fetchDatabaseJsonのlimitは1-3000にクランプ', () => {
    const { fetchDatabaseJson } = setupDoGet();
    const r1 = JSON.parse(fetchDatabaseJson(0));
    expect(r1.length).toBeGreaterThanOrEqual(1);
    const r2 = JSON.parse(fetchDatabaseJson(99999));
    expect(r2).toHaveLength(2); // データ2件しかないので2
  });
});

// ========================================
// logSearch（モック付き）
// ========================================
describe('logSearch', () => {
  it('2文字以上の検索語をログに記録する', () => {
    const sheets = {};
    const SpreadsheetApp = createMockSpreadsheetApp(sheets);
    const logSearch = createLogSearch(SpreadsheetApp, '検索ログ');

    logSearch('テスト', { presenter: 'トロピ大塚', colors: ['W', 'U'] });

    expect(sheets['検索ログ']).toBeDefined();
    // ヘッダー + 1行 = 2行
    expect(sheets['検索ログ']._data.length).toBeGreaterThanOrEqual(2);
  });

  it('1文字の検索語は無視する', () => {
    const sheets = {};
    const SpreadsheetApp = createMockSpreadsheetApp(sheets);
    const logSearch = createLogSearch(SpreadsheetApp, '検索ログ');

    logSearch('あ');
    expect(sheets['検索ログ']).toBeUndefined();
  });

  it('空文字は無視する', () => {
    const sheets = {};
    const SpreadsheetApp = createMockSpreadsheetApp(sheets);
    const logSearch = createLogSearch(SpreadsheetApp, '検索ログ');

    logSearch('');
    expect(sheets['検索ログ']).toBeUndefined();
  });

  it('100文字超は切り詰める', () => {
    const sheets = {};
    const SpreadsheetApp = createMockSpreadsheetApp(sheets);
    const logSearch = createLogSearch(SpreadsheetApp, '検索ログ');
    const longTerm = 'あ'.repeat(200);

    logSearch(longTerm);

    const data = sheets['検索ログ']._data;
    const lastRow = data[data.length - 1];
    expect(lastRow[1].length).toBe(100);
  });
});

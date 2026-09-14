/**
 * GAS本体のエラー耐性と中断ガードのテスト
 * - fetchJsonWithRetry_ の 5xx / ネットワーク例外リトライ・Retry-After 対応
 * - fetchMtgCardDataJa のサマリ・中断仕様
 */
import { describe, it, expect } from 'vitest';
import {
  createMockSheet,
  createMockSpreadsheetApp,
  createMockUrlFetchApp,
  createMockUtilities,
} from './gas-mocks.js';
import * as gasAdapter from './gas-adapter.js';

const { createFetchJsonWithRetry } = gasAdapter;

// ========================================
// A. fetchJsonWithRetry_ の新リトライ仕様
// ========================================
describe('fetchJsonWithRetry_ (error resilience)', () => {
  it('5xxはリトライして最終的に成功する (500 → 503 → 200)', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(500, '{"details":"server error"}');
    mockFetch._addResponse(503, '{"details":"unavailable"}');
    mockFetch._addResponse(200, { data: '5xx-retry-ok' });
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(fn('https://api.example.com', {})).toEqual({ data: '5xx-retry-ok' });
  });

  it('5xxがリトライ上限を超えたらスロー (maxRetries=1, 500×2)', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(500, '{"details":"server error"}');
    mockFetch._addResponse(500, '{"details":"server error"}');
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(() => fn('https://api.example.com', {}, 1)).toThrow();
    // リトライが実際に行われたこと（2回フェッチ）を確認
    expect(mockFetch._urls.length).toBe(2);
  });

  it('fetch例外（ネットワークエラー）はリトライして成功する', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addError('DNS error: api.scryfall.com');
    mockFetch._addResponse(200, { data: 'network-retry-ok' });
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(fn('https://api.example.com', {})).toEqual({ data: 'network-retry-ok' });
  });

  it('fetch例外がリトライ上限を超えたらスロー (maxRetries=1, エラー×2)', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addError('DNS error: api.scryfall.com');
    mockFetch._addError('DNS error: api.scryfall.com');
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(() => fn('https://api.example.com', {}, 1)).toThrow();
    // リトライが実際に行われたこと（2回フェッチ）を確認
    expect(mockFetch._urls.length).toBe(2);
  });

  it('429のRetry-Afterヘッダーを尊重して待機する (Retry-After: 7 → 7000ms)', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(429, '', { 'Retry-After': '7' });
    mockFetch._addResponse(200, { data: 'retry-after-ok' });
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(fn('https://api.example.com', {})).toEqual({ data: 'retry-after-ok' });
    // ヘッダー指定が尊重される（デフォルトバックオフではない）
    expect(mockUtil._sleeps).toContain(7000);
    expect(mockUtil._sleeps[0]).toBe(7000);
  });

  it('429でヘッダーなしは指数バックオフ (1000 → 2000)', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(429, '');
    mockFetch._addResponse(429, '');
    mockFetch._addResponse(200, { data: 'backoff-ok' });
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(fn('https://api.example.com', {})).toEqual({ data: 'backoff-ok' });
    expect(mockUtil._sleeps[0]).toBe(1000);
    expect(mockUtil._sleeps[1]).toBe(2000);
  });

  it('404はリトライせず即座にスローする', () => {
    const mockFetch = createMockUrlFetchApp();
    const mockUtil = createMockUtilities();
    mockFetch._addResponse(404, '{"details":"not found"}');
    const fn = createFetchJsonWithRetry(mockFetch, mockUtil);
    expect(() => fn('https://api.example.com', {})).toThrow('not found');
    expect(mockFetch._urls.length).toBe(1);
  });
});

// ========================================
// B. fetchMtgCardDataJa
// ========================================

// データベースシートの29列ヘッダー（L列=12列目: 紹介カード、R列=18列目〜: Scryfallデータ）
const DB_HEADERS = ['動画タイトル','ID','公開日(UTC)','公開日(JST)','動画URL','サムネイルURL','再生時間','話数','紹介順','今のうちだぞ','プレゼンター','紹介カード','関連する統率者','関連する統率者2','関連カード','テーマ','上位互換？','カード名 (日本語)','カード名 (英語)','マナコスト','カードタイプ','オラクルテキスト','Power','Toughness','Colors','Color Identity','画像URL','Scryfall','CMC'];

/** L列(12列目)にカード名を入れた29列の行を作る */
function makeDbRow(cardName, filledR = '', filledS = '') {
  const row = new Array(29).fill('');
  row[11] = cardName; // L列 = 紹介カード
  row[17] = filledR;  // R列 = カード名 (日本語)
  row[18] = filledS;  // S列 = カード名 (英語)
  return row;
}

/** Scryfall検索（lang:ja）レスポンス */
function jaSearchResponse(printedName = '太陽の指輪', overrides = {}) {
  return {
    object: 'list',
    data: [{
      object: 'card',
      lang: 'ja',
      name: 'Sol Ring',
      printed_name: printedName,
      image_uris: { normal: 'https://x/ja.jpg' },
      scryfall_uri: 'https://scryfall.com/x',
      cmc: 1,
      colors: [],
      color_identity: [],
      ...overrides,
    }],
  };
}

/** cards/named?exact（英語画像）レスポンス */
function enExactResponse() {
  return {
    object: 'card',
    lang: 'en',
    image_uris: { normal: 'https://x/en.jpg' },
  };
}

function setup(rows, { now } = {}) {
  const dbSheet = createMockSheet(DB_HEADERS, rows);
  const SpreadsheetApp = createMockSpreadsheetApp({ 'データベース': dbSheet });
  const UrlFetchApp = createMockUrlFetchApp();
  const Utilities = createMockUtilities();
  const deps = { SpreadsheetApp, UrlFetchApp, Utilities };
  if (now) deps.now = now;
  return { dbSheet, SpreadsheetApp, UrlFetchApp, Utilities, deps };
}

describe('createFetchMtgCardDataJa', () => {
  it('ハッピーパス: 2行のカードを取得してR..AC列を埋め、サマリを返す', () => {
    const { dbSheet, UrlFetchApp, deps } = setup([
      makeDbRow('Sol Ring'),
      makeDbRow('Mana Crypt'),
    ]);
    // カードごとに (1) cards/search lang:ja (2) cards/named?exact の2フェッチ
    UrlFetchApp._addResponse(200, jaSearchResponse());
    UrlFetchApp._addResponse(200, enExactResponse());
    UrlFetchApp._addResponse(200, jaSearchResponse());
    UrlFetchApp._addResponse(200, enExactResponse());

    const run = gasAdapter.createFetchMtgCardDataJa(deps);
    const summary = run();

    expect(summary).toMatchObject({ processed: 2, succeeded: 2, failed: 0, aborted: null });

    // 両行のR..AC列(18..29列, 0-indexed 17..28)が埋まる
    for (const rowIdx of [1, 2]) {
      const row = dbSheet._data[rowIdx];
      expect(row[17]).toBe('太陽の指輪');        // R: カード名(日本語)
      expect(row[18]).toBe('Sol Ring');          // S: カード名(英語)
      expect(row[26]).toBe('https://x/en.jpg');  // AA: 画像URL（英語版優先）
      expect(row[27]).toBe('https://scryfall.com/x'); // AB: Scryfall
      expect(row[28]).toBe(1);                   // AC: CMC
    }
  });

  it('レジューム: R列が既に埋まっている行はスキップしfetchしない', () => {
    const { UrlFetchApp, deps } = setup([
      makeDbRow('Sol Ring', '太陽の指輪', 'Sol Ring'),
    ]);

    const run = gasAdapter.createFetchMtgCardDataJa(deps);
    const summary = run();

    expect(summary.skipped).toBe(1);
    expect(UrlFetchApp._urls.length).toBe(0);
  });

  it('紹介カードと取得済みカード名が違う行は再取得して修復する', () => {
    const { dbSheet, UrlFetchApp, deps } = setup([
      makeDbRow('むさぼり喰うストロサス', 'Wall of Kelp', 'Wall of Kelp'),
    ]);
    UrlFetchApp._addResponse(200, jaSearchResponse('むさぼり喰うストロサス', {
      name: 'Devouring Strossus',
      mana_cost: '{5}{B}{B}{B}',
      colors: ['B'],
      color_identity: ['B'],
      cmc: 8,
    }));
    UrlFetchApp._addResponse(200, enExactResponse());

    const summary = gasAdapter.createFetchMtgCardDataJa(deps)();

    expect(summary).toMatchObject({ processed: 1, succeeded: 1, skipped: 0 });
    expect(dbSheet._data[1][17]).toBe('むさぼり喰うストロサス');
    expect(dbSheet._data[1][18]).toBe('Devouring Strossus');
    expect(dbSheet._data[1][19]).toBe('{5}{B}{B}{B}');
    expect(dbSheet._data[1][25]).toBe('B');
    expect(dbSheet._data[1][28]).toBe(8);
  });

  it('取得中に紹介カード行が移動したら誤った行へ書かず中断する', () => {
    const { dbSheet, UrlFetchApp, deps } = setup([makeDbRow('Sol Ring')]);
    const fetch = UrlFetchApp.fetch;
    UrlFetchApp.fetch = (...args) => {
      const response = fetch(...args);
      dbSheet._data[1][11] = 'Mana Crypt';
      return response;
    };
    UrlFetchApp._addResponse(200, jaSearchResponse());
    UrlFetchApp._addResponse(200, enExactResponse());

    const summary = gasAdapter.createFetchMtgCardDataJa(deps)();

    expect(summary.aborted).toBe('source-rows-changed');
    expect(summary.succeeded).toBe(0);
    expect(dbSheet._data[1][17]).toBe('');
  });

  it('連続失敗ブレーカー: 3カード連続失敗で中断し、4枚目は試行しない', () => {
    const { dbSheet, UrlFetchApp, deps } = setup([
      makeDbRow('CardA'),
      makeDbRow('CardB'),
      makeDbRow('CardC'),
      makeDbRow('CardD'),
    ]);
    // すべてのfetchが例外で失敗し続ける（リトライ分も含めて十分な数を積む）
    for (let i = 0; i < 40; i++) {
      UrlFetchApp._addError('DNS error: api.scryfall.com');
    }

    const run = gasAdapter.createFetchMtgCardDataJa(deps);
    const summary = run();

    expect(summary.aborted).toBe('consecutive-failures');
    expect(summary.failed).toBe(3);
    // 4枚目(CardD)へのfetchは発生していない
    expect(UrlFetchApp._urls.every(u => !u.includes('CardD'))).toBe(true);
    // 4枚目の行は未記入のまま（再実行可能）
    expect(dbSheet._data[4][17]).toBe('');
  });

  it('時間予算ガード: 5分超過で中断し、残り行は未記入のまま', () => {
    let callCount = 0;
    const now = () => (callCount++ === 0 ? 0 : 10 * 60 * 1000); // 初回0、以降は5分超
    const { dbSheet, UrlFetchApp, deps } = setup(
      [makeDbRow('Sol Ring'), makeDbRow('Mana Crypt')],
      { now }
    );
    UrlFetchApp._addResponse(200, jaSearchResponse());
    UrlFetchApp._addResponse(200, enExactResponse());

    const run = gasAdapter.createFetchMtgCardDataJa(deps);
    const summary = run();

    expect(summary.aborted).toBe('time-budget');
    // 残り行（2行目）は未記入のまま再実行可能
    expect(dbSheet._data[2][17]).toBe('');
  });
});

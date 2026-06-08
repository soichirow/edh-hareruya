import { describe, it, expect } from 'vitest';
import {
  getTitle,
  getPresenterBucket,
  ciSet,
  matchColorsCommander,
  colorSortKey,
  getCmc,
  buildSearchText,
  filterCards,
  sortCards,
  MAIN_PRESENTERS,
} from '../js/card-utils.js';

// --- テスト用ヘルパー ---
function makeCard(overrides = {}) {
  return {
    'カード名 (日本語)': '',
    'カード名 (英語)': '',
    'プレゼンター': '',
    'Color Identity': '',
    '話数': '',
    'CMC': '',
    '画像URL': 'https://example.com/img.jpg',
    '動画URL': '',
    'Scryfall': '',
    'テーマ': '',
    '関連する統率者': '',
    ...overrides,
  };
}

// ========================================
// getTitle
// ========================================
describe('getTitle', () => {
  it('日本語名を優先して返す', () => {
    const card = makeCard({ 'カード名 (日本語)': '太陽の指輪', 'カード名 (英語)': 'Sol Ring' });
    expect(getTitle(card)).toBe('太陽の指輪');
  });

  it('日本語名が空なら英語名を返す', () => {
    const card = makeCard({ 'カード名 (英語)': 'Sol Ring' });
    expect(getTitle(card)).toBe('Sol Ring');
  });

  it('どちらも空なら「カード」を返す', () => {
    const card = makeCard();
    expect(getTitle(card)).toBe('カード');
  });

  it('前後の空白をトリムする', () => {
    const card = makeCard({ 'カード名 (日本語)': '  太陽の指輪  ' });
    expect(getTitle(card)).toBe('太陽の指輪');
  });
});

// ========================================
// getPresenterBucket
// ========================================
describe('getPresenterBucket', () => {
  it('主要プレゼンターはそのまま返す', () => {
    const card = makeCard({ 'プレゼンター': 'トロピ大塚' });
    expect(getPresenterBucket(card)).toBe('トロピ大塚');
  });

  it('カンマ区切りで最初に一致した主要プレゼンターを返す', () => {
    const card = makeCard({ 'プレゼンター': 'ゲスト,いってつ' });
    expect(getPresenterBucket(card)).toBe('いってつ');
  });

  it('主要プレゼンターに該当しなければ「その他」', () => {
    const card = makeCard({ 'プレゼンター': '謎の人物' });
    expect(getPresenterBucket(card)).toBe('その他');
  });

  it('空文字なら空文字を返す', () => {
    const card = makeCard({ 'プレゼンター': '' });
    expect(getPresenterBucket(card)).toBe('');
  });

  it('全主要プレゼンターが正しく定義されている', () => {
    const expected = ['トロピ大塚', 'いってつ', 'スギちゃん', 'タイシン', 'タカノシゲキ', 'ソラノ', 'ブチャラティ', '卍幻日輪廻卍のタイシン'];
    for (const name of expected) {
      expect(MAIN_PRESENTERS.has(name)).toBe(true);
    }
  });
});

// ========================================
// ciSet
// ========================================
describe('ciSet', () => {
  it('カンマ区切りの色文字列をSetに変換', () => {
    const card = makeCard({ 'Color Identity': 'W,U' });
    expect(ciSet(card)).toEqual(new Set(['W', 'U']));
  });

  it('配列をそのままSetに変換', () => {
    const card = makeCard({ 'Color Identity': ['R', 'G'] });
    expect(ciSet(card)).toEqual(new Set(['R', 'G']));
  });

  it('空文字列は空Setを返す（無色）', () => {
    const card = makeCard({ 'Color Identity': '' });
    expect(ciSet(card)).toEqual(new Set());
  });

  it('小文字でも大文字に正規化する', () => {
    const card = makeCard({ 'Color Identity': 'w,b' });
    expect(ciSet(card)).toEqual(new Set(['W', 'B']));
  });

  it('WUBRG以外の文字は無視する', () => {
    const card = makeCard({ 'Color Identity': 'W,X,Z' });
    expect(ciSet(card)).toEqual(new Set(['W']));
  });

  it('full color words are parsed as tokens, not individual letters', () => {
    const card = makeCard({ 'Color Identity': 'Green' });
    expect(ciSet(card)).toEqual(new Set(['G']));
  });

  it('colorless words and C are treated as no color identity', () => {
    expect(ciSet(makeCard({ 'Color Identity': 'Colorless' }))).toEqual(new Set());
    expect(ciSet(makeCard({ 'Color Identity': 'C' }))).toEqual(new Set());
  });
});

// ========================================
// matchColorsCommander
// ========================================
describe('matchColorsCommander', () => {
  it('色未選択なら全カードを通す', () => {
    const card = makeCard({ 'Color Identity': 'W,U,B' });
    expect(matchColorsCommander(card, [])).toBe(true);
  });

  it('無色(C)のみ選択 → 無色カードだけ通す', () => {
    const colorless = makeCard({ 'Color Identity': '' });
    const hasColor = makeCard({ 'Color Identity': 'R' });
    expect(matchColorsCommander(colorless, ['C'])).toBe(true);
    expect(matchColorsCommander(hasColor, ['C'])).toBe(false);
  });

  it('選択色の部分集合ならOK (R,U選択 → Rだけのカードも通す)', () => {
    const cardR = makeCard({ 'Color Identity': 'R' });
    const cardRU = makeCard({ 'Color Identity': 'R,U' });
    const cardRUG = makeCard({ 'Color Identity': 'R,U,G' });
    expect(matchColorsCommander(cardR, ['R', 'U'])).toBe(true);
    expect(matchColorsCommander(cardRU, ['R', 'U'])).toBe(true);
    expect(matchColorsCommander(cardRUG, ['R', 'U'])).toBe(false);
  });

  it('colorless cards require C to be selected', () => {
    const colorless = makeCard({ 'Color Identity': '' });
    expect(matchColorsCommander(colorless, ['R'])).toBe(false);
    expect(matchColorsCommander(colorless, ['R', 'C'])).toBe(true);
  });
});

// ========================================
// colorSortKey
// ========================================
describe('colorSortKey', () => {
  it('無色は [0, 0, name] を返す', () => {
    const card = makeCard({ 'Color Identity': '', 'カード名 (日本語)': 'Sol Ring' });
    const key = colorSortKey(card);
    expect(key[0]).toBe(0);
    expect(key[1]).toBe(0);
  });

  it('白単色は [1, 1, name]', () => {
    const card = makeCard({ 'Color Identity': 'W', 'カード名 (日本語)': 'テスト' });
    const key = colorSortKey(card);
    expect(key[0]).toBe(1);
    expect(key[1]).toBe(1);
  });

  it('多色はminRankが小さい色の順番、色数で比較', () => {
    const wu = makeCard({ 'Color Identity': 'W,U', 'カード名 (日本語)': 'A' });
    const rb = makeCard({ 'Color Identity': 'R,B', 'カード名 (日本語)': 'B' });
    expect(colorSortKey(wu)[0]).toBeLessThan(colorSortKey(rb)[0]);
  });
});

// ========================================
// getCmc
// ========================================
describe('getCmc', () => {
  it('数値を返す', () => {
    const card = makeCard({ 'CMC': 3 });
    expect(getCmc(card)).toBe(3);
  });

  it('文字列の数字もパースする', () => {
    const card = makeCard({ 'CMC': '5' });
    expect(getCmc(card)).toBe(5);
  });

  it('空・不正値はPositive Infinityを返す', () => {
    const card = makeCard({ 'CMC': '' });
    expect(getCmc(card)).toBe(Number.POSITIVE_INFINITY);
  });

  it('0は0を返す', () => {
    const card = makeCard({ 'CMC': 0 });
    expect(getCmc(card)).toBe(0);
  });
});

// ========================================
// buildSearchText
// ========================================
describe('buildSearchText', () => {
  it('日本語名・英語名・統率者・テーマを結合する', () => {
    const card = makeCard({
      'カード名 (日本語)': '太陽の指輪',
      'カード名 (英語)': 'Sol Ring',
      '関連する統率者': 'ウルザ',
      'テーマ': 'マナ加速',
    });
    const text = buildSearchText(card);
    expect(text).toContain('太陽の指輪');
    expect(text).toContain('sol ring');
    expect(text).toContain('ウルザ');
    expect(text).toContain('マナ加速');
  });

  it('空フィールドは無視される', () => {
    const card = makeCard({ 'カード名 (日本語)': 'テスト' });
    const text = buildSearchText(card);
    expect(text).toContain('テスト');
    expect(text).not.toContain('undefined');
  });
});

// ========================================
// filterCards
// ========================================
describe('filterCards', () => {
  const cards = [
    makeCard({ 'カード名 (日本語)': '太陽の指輪', 'プレゼンター': 'トロピ大塚', 'Color Identity': '', 'CMC': 1 }),
    makeCard({ 'カード名 (日本語)': '稲妻', 'プレゼンター': 'いってつ', 'Color Identity': 'R', 'CMC': 1 }),
    makeCard({ 'カード名 (日本語)': '対抗呪文', 'プレゼンター': 'スギちゃん', 'Color Identity': 'U', 'CMC': 2 }),
    makeCard({ 'カード名 (英語)': 'Demonic Tutor', 'プレゼンター': '謎の人', 'Color Identity': 'B', 'CMC': 2 }),
  ];
  // buildSearchText を事前に付与
  const prepared = cards.map(c => ({ ...c, _searchText: buildSearchText(c) }));

  it('検索語でフィルタ', () => {
    const result = filterCards(prepared, { query: '稲妻' });
    expect(result).toHaveLength(1);
    expect(result[0]['カード名 (日本語)']).toBe('稲妻');
  });

  it('英語名でも検索できる', () => {
    const result = filterCards(prepared, { query: 'demonic' });
    expect(result).toHaveLength(1);
  });

  it('プレゼンターでフィルタ', () => {
    const result = filterCards(prepared, { presenter: 'いってつ' });
    expect(result).toHaveLength(1);
  });

  it('その他プレゼンターでフィルタ', () => {
    const result = filterCards(prepared, { presenter: 'その他' });
    expect(result).toHaveLength(1);
  });

  it('色でフィルタ（R選択 → R以下のカードのみ）', () => {
    const result = filterCards(prepared, { colors: ['R'] });
    expect(result).toHaveLength(1);
    expect(result[0]['Color Identity']).toBe('R');
  });

  it('複合フィルタ', () => {
    const result = filterCards(prepared, { query: '', presenter: '', colors: ['U'] });
    expect(result).toHaveLength(1);
    expect(result[0]['Color Identity']).toBe('U');
  });

  it('C can be combined with a color to include colorless cards explicitly', () => {
    const result = filterCards(prepared, { colors: ['R', 'C'] });
    expect(result).toHaveLength(2);
    expect(result.map(card => card['Color Identity']).sort()).toEqual(['', 'R']);
  });

  it('条件なしなら全件返す', () => {
    const result = filterCards(prepared, {});
    expect(result).toHaveLength(4);
  });
});

// ========================================
// sortCards
// ========================================
describe('sortCards', () => {
  const cards = [
    makeCard({ 'カード名 (日本語)': 'B', 'Color Identity': 'R', 'CMC': 3, '話数': 5 }),
    makeCard({ 'カード名 (日本語)': 'A', 'Color Identity': 'W', 'CMC': 1, '話数': 10 }),
    makeCard({ 'カード名 (日本語)': 'C', 'Color Identity': '', 'CMC': 0, '話数': 1 }),
  ];

  it('色ソート昇順: 無色→白→赤', () => {
    const sorted = sortCards([...cards], 'color', 'asc');
    expect(sorted[0]['カード名 (日本語)']).toBe('C'); // 無色
    expect(sorted[1]['カード名 (日本語)']).toBe('A'); // 白
    expect(sorted[2]['カード名 (日本語)']).toBe('B'); // 赤
  });

  it('色ソート降順', () => {
    const sorted = sortCards([...cards], 'color', 'desc');
    expect(sorted[0]['カード名 (日本語)']).toBe('B'); // 赤
    expect(sorted[2]['カード名 (日本語)']).toBe('C'); // 無色
  });

  it('話数ソート昇順', () => {
    const sorted = sortCards([...cards], 'episode', 'asc');
    expect(sorted[0]['話数']).toBe(1);
    expect(sorted[2]['話数']).toBe(10);
  });

  it('CMCソート昇順', () => {
    const sorted = sortCards([...cards], 'cmc', 'asc');
    expect(getCmc(sorted[0])).toBe(0);
    expect(getCmc(sorted[2])).toBe(3);
  });

  it('ソートキーnoneなら順序変更なし', () => {
    const original = [...cards];
    const sorted = sortCards([...cards], 'none', 'asc');
    expect(sorted.map(c => c['カード名 (日本語)'])).toEqual(original.map(c => c['カード名 (日本語)']));
  });
});

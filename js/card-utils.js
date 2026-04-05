// カード列名定数
export const KEY_IMG = '画像URL';
export const KEY_JA = 'カード名 (日本語)';
export const KEY_EN = 'カード名 (英語)';
export const KEY_PRES = 'プレゼンター';
export const KEY_COLOR_ID = 'Color Identity';
export const KEY_EPISODE = '話数';
export const KEY_CMC = 'CMC';
export const KEY_YT = '動画URL';
export const KEY_SCRY = 'Scryfall';
export const KEY_THEME = 'テーマ';
export const KEY_COMM = '関連する統率者';

export const MAIN_PRESENTERS = new Set([
  'トロピ大塚', 'いってつ', 'スギちゃん', 'タイシン',
  'タカノシゲキ', 'ソラノ', 'ブチャラティ', '卍幻日輪廻卍のタイシン',
]);

// --- ユーティリティ ---

const s = (v) => String(v ?? '').trim();

export function getTitle(row) {
  return s(row[KEY_JA]) || s(row[KEY_EN]) || 'カード';
}

function getPresenterRaw(row) {
  return s(row[KEY_PRES]);
}

export function getPresenterBucket(row) {
  const raw = getPresenterRaw(row);
  if (!raw) return '';
  const parts = raw.split(/[,\u3001/／・\s]+/).map(x => x.trim()).filter(Boolean);
  for (const p of parts) {
    if (MAIN_PRESENTERS.has(p)) return p;
  }
  return 'その他';
}

export function ciSet(row) {
  const v = row[KEY_COLOR_ID];
  if (Array.isArray(v)) {
    return new Set(v.map(x => String(x).toUpperCase()).filter(c => 'WUBRG'.includes(c)));
  }
  const letters = String(v ?? '').toUpperCase().match(/[WUBRG]/g) || [];
  return new Set(letters);
}

export function matchColorsCommander(row, checked) {
  const set = new Set(checked);
  const have = ciSet(row);

  if (set.size === 0) return true;
  if (set.size === 1 && set.has('C')) return have.size === 0;
  if (set.has('C')) set.delete('C');

  for (const c of have) {
    if (!set.has(c)) return false;
  }
  return true;
}

const COLOR_ORDER = { 'W': 1, 'U': 2, 'B': 3, 'R': 4, 'G': 5 };

export function colorSortKey(row) {
  const have = ciSet(row);
  if (have.size === 0) return [0, 0, getTitle(row).toLowerCase()];
  const seq = ['W', 'U', 'B', 'R', 'G'].filter(c => have.has(c));
  const minRank = COLOR_ORDER[seq[0]];
  return [minRank, seq.length, getTitle(row).toLowerCase()];
}

function toNum(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[^\d.-]/g, ''));
  return Number.isFinite(n) ? n : Number.NEGATIVE_INFINITY;
}

export function getCmc(row) {
  const raw = row?.[KEY_CMC];
  const str = String(raw ?? '').trim();
  if (str === '') return Number.POSITIVE_INFINITY;
  const n = Number(str);
  return Number.isFinite(n) ? n : Number.POSITIVE_INFINITY;
}

export function buildSearchText(row) {
  const fields = [
    row[KEY_JA],
    row[KEY_EN],
    row[KEY_COMM],
    row[KEY_THEME],
  ];
  return fields
    .flatMap(v => {
      const str = String(v ?? '');
      return str.split(/[\s,、，/／・]+/).filter(Boolean);
    })
    .map(x => x.toLowerCase())
    .join(' ');
}

// --- フィルタ・ソート ---

export function filterCards(cards, { query = '', presenter = '', colors = [] } = {}) {
  const q = s(query).toLowerCase();

  return cards.filter(row => {
    if (q && !String(row._searchText || '').includes(q)) return false;
    if (presenter) {
      const bucket = getPresenterBucket(row);
      if (presenter !== bucket) return false;
    }
    if (!matchColorsCommander(row, colors)) return false;
    return true;
  });
}

export function sortCards(cards, sortKey, sortDir) {
  const mul = sortDir === 'desc' ? -1 : 1;

  if (sortKey === 'color') {
    cards.sort((a, b) => {
      const aa = colorSortKey(a);
      const bb = colorSortKey(b);
      for (let i = 0; i < Math.max(aa.length, bb.length); i++) {
        if (aa[i] === bb[i]) continue;
        return (aa[i] < bb[i] ? -1 : 1) * mul;
      }
      return 0;
    });
  } else if (sortKey === 'episode') {
    cards.sort((a, b) => (toNum(a[KEY_EPISODE]) - toNum(b[KEY_EPISODE])) * mul);
  } else if (sortKey === 'cmc') {
    cards.sort((a, b) => {
      const d = getCmc(a) - getCmc(b);
      if (d !== 0) return d * mul;
      const aa = colorSortKey(a);
      const bb = colorSortKey(b);
      for (let i = 0; i < Math.max(aa.length, bb.length); i++) {
        if (aa[i] === bb[i]) continue;
        return (aa[i] < bb[i] ? -1 : 1) * mul;
      }
      return 0;
    });
  }

  return cards;
}

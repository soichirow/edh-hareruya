function fetchMtgCardDataJa() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("データベース");
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const cardNames = sheet.getRange(2, 12, lastRow - 1, 1).getValues();

  const fetchOptions = {
    muteHttpExceptions: true,
    headers: {
      // ここはあなたのスクリプト名/バージョンに置き換え推奨
      // 連絡先があると運用上安全
      "User-Agent": "MtgSheetFetcher/1.0 (contact: sowatanabe@bushiroad-gp.com)",
      "Accept": "application/json",
    },
  };

  for (let i = 0; i < cardNames.length; i++) {
    const row = i + 2;
    const cardNameRaw = (cardNames[i][0] || "").toString().trim();
    if (!cardNameRaw) continue;

    // すでにR列が埋まっていたらスキップ
    const already = sheet.getRange(row, 18).getValue();
    if (already) continue;

    try {
      const card = findCardPreferJa(cardNameRaw, fetchOptions);

      const nameJa = card.printed_name || "";
      const nameEn = card.name || "";

      const manaCost = getJoined(card, "mana_cost", "mana_cost");
      const typeLine = getJoined(card, "type_line", "printed_type_line");
      const oracleText = getJoined(card, "oracle_text", "printed_text");

      const power = getJoined(card, "power", "power");
      const toughness = getJoined(card, "toughness", "toughness");

      const colors = (card.colors || []).join(",");
      const colorIdentity = (card.color_identity || []).join(",");
      const imageUri = getImageNormal(card);
      const scryfallUri = card.scryfall_uri || "";
      const cmc = (card.cmc !== null && card.cmc !== undefined) ? card.cmc : "";

      // R列以降に書き込み（12列）
      sheet.getRange(row, 18, 1, 12).setValues([[
        nameJa, nameEn, manaCost, typeLine, oracleText,
        power, toughness, colors, colorIdentity, imageUri, scryfallUri, cmc
      ]]);

      // 1リクエスト/カードでも安全側で少し待つ（10req/sec未満目安）
      Utilities.sleep(150);

    } catch (e) {
      console.error(`カード取得失敗: ${cardNameRaw} - ${e && e.message ? e.message : e}`);
      // 次へ進む前に少し長めに待つ（ブロック回避）
      Utilities.sleep(1000);
    }
  }
}

// 日本語印刷を優先して1枚取得。なければ /cards/named fuzzy で英語等を拾う。
function findCardPreferJa(cardName, fetchOptions) {
  // /cards/search で完全一致 + 日本語印刷を優先
  // ! は完全一致（語句やスペースがあるので常にクォート）
  const escaped = cardName.replace(/"/g, '\\"');
  const qJa = `!"${escaped}" lang:ja`;
  const urlJa = `https://api.scryfall.com/cards/search?q=${encodeURIComponent(qJa)}&unique=prints&order=released&dir=desc`;
  const ja = fetchJsonWithRetry_(urlJa, fetchOptions);

  if (ja && ja.object === "list" && Array.isArray(ja.data) && ja.data.length > 0) {
    return ja.data[0];
  }

  // 日本語印刷がない/名前が一致しない場合のフォールバック（曖昧検索）
  const urlFuzzy = `https://api.scryfall.com/cards/named?fuzzy=${encodeURIComponent(cardName)}`;
  const fuzzy = fetchJsonWithRetry_(urlFuzzy, fetchOptions);

  if (fuzzy && fuzzy.object === "card") return fuzzy;

  const detail = (ja && ja.details) ? ja.details : (fuzzy && fuzzy.details) ? fuzzy.details : "not found";
  throw new Error(detail);
}

// 429 を中心に簡易リトライ（指数バックオフ）
function fetchJsonWithRetry_(url, fetchOptions, maxRetries) {
  const retries = (maxRetries === null || maxRetries === undefined) ? 5 : maxRetries;

  for (let attempt = 0; attempt <= retries; attempt++) {
    const res = UrlFetchApp.fetch(url, fetchOptions);
    const code = res.getResponseCode();
    const text = res.getContentText();

    // 成功
    if (code >= 200 && code < 300) return JSON.parse(text);

    // 429 は待ってリトライ
    if (code === 429 && attempt < retries) {
      const waitMs = Math.min(8000, 500 * Math.pow(2, attempt)); // 500,1000,2000,4000,8000...
      Utilities.sleep(waitMs);
      continue;
    }

    // その他エラー
    let payload;
    try { payload = JSON.parse(text); } catch (_) { payload = { details: text }; }
    const msg = payload && payload.details ? payload.details : `HTTP ${code}`;
    throw new Error(msg);
  }

  throw new Error("retry exceeded");
}

// 多面カードは face ごとに結合して返す（日本語優先、なければ英語）
function getJoined(card, baseKey, printedKey) {
  if (Array.isArray(card.card_faces) && card.card_faces.length > 0) {
    const parts = card.card_faces.map(f => (f[printedKey] || f[baseKey] || "").toString());
    return parts.join(" // ");
  }
  return (card[printedKey] || card[baseKey] || "").toString();
}

function getImageNormal(card) {
  if (card.image_uris && card.image_uris.normal) return card.image_uris.normal;
  if (Array.isArray(card.card_faces) && card.card_faces[0] && card.card_faces[0].image_uris && card.card_faces[0].image_uris.normal) {
    return card.card_faces[0].image_uris.normal;
  }
  return "";
}
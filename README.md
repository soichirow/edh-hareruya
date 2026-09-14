# 晴れる屋コマンダーオタクカード紹介カード一覧

晴れる屋MTGの YouTube チャンネルで紹介された「コマンダーオタクカード」を検索・閲覧する非公式ファンサイトです。

- 公開サイト: <https://soichirow.github.io/edh-hareruya/>
- 運用・データ契約・監査結果: [docs/site-governance.md](docs/site-governance.md)

## 機能

- カード名検索（日本語・英語）
- プレゼンター、色、マナ総量（CMC）による絞り込み
- 色、話数、マナ総量による並び替え
- お気に入り、ダークモード、日本語 / English 切替
- カード画像の拡大表示
- 同意した場合だけ読み込む Google Analytics 4

色フィルターは `Color Identity`、マナ総量は `CMC` を使用します。欠損 CMC は昇順・降順のどちらでも末尾です。

## 構成

| レイヤー | 実装 |
|---|---|
| フロントエンド | `docs/` の静的 HTML / JavaScript、GitHub Pages |
| データ API | Google Apps Script の JSON Web API |
| データソース | Google Spreadsheet の `データベース` シート |
| カード情報 | [Scryfall API](https://scryfall.com/) |
| 品質確認 | Vitest、ESLint、公開サイト E2E |

GitHub Pages、GAS、Spreadsheet は別々に反映されます。`main` への push だけでは GAS やシートは更新されません。変更前に必ず[運用手順](docs/site-governance.md)を確認してください。

## ローカル確認

```bash
npm install
npm test
npm run lint
```

`npm run gas:pull` / `npm run gas:push` には、認証済みの clasp と、Git 管理しない `gas/.clasp.json` が必要です。クラウド側の Head とローカル差分を確認せずに `gas:push` しないでください。GAS の公開反映には push 後のデプロイ版更新も必要です。

## クレジット

- カードデータ・画像: [Scryfall API](https://scryfall.com/)
- 動画: [晴れる屋MTG YouTube チャンネル](https://www.youtube.com/@haaboreMTG)
- このサイトは[ファンコンテンツ・ポリシー](https://company.wizards.com/ja/legal/fancontentpolicy)に沿った非公式のファンコンテンツです。ウィザーズ社の認可・許諾を受けたものではありません。題材の一部にウィザーズ・オブ・ザ・コースト社の財産を含みます。© Wizards of the Coast LLC.

# EDH オタクカード画像一覧

晴れる屋MTGの YouTube チャンネルで紹介された「EDH オタクカード」の一覧サイトです。

**非公式ファンサイト** - ウィザーズ社の認可/許諾は得ていません。

## サイト

https://soichirow.github.io/edh-hareruya/

## 機能

- カード名検索（日本語・英語）
- プレゼンター別フィルター
- MTG 色フィルター（W/U/B/R/G/C）
- ソート（色・話数・CMC）
- お気に入り（localStorage 保存）
- カード画像クリックで拡大表示（モーダル）
- ダークモード
- 日本語 / English 切替

## 技術構成

| レイヤー | 技術 |
|---------|------|
| フロントエンド | GitHub Pages（静的 HTML + Bootstrap 5） |
| データ API | Google Apps Script（JSON エンドポイント） |
| データソース | Google Spreadsheet |
| カード情報 | [Scryfall API](https://scryfall.com/) |
| テスト | Vitest（86 テスト） |
| Lint | ESLint |

## 開発

```bash
npm install
npm test        # テスト実行
npm run lint    # Lint
npm run gas:push  # GAS コードをプッシュ
npm run gas:pull  # GAS コードをプル
```

## クレジット

- カードデータ・画像: [Scryfall API](https://scryfall.com/)
- 動画: [晴れる屋MTG YouTube チャンネル](https://www.youtube.com/@haaboreMTG)
- このサイトは[ファンコンテンツ・ポリシー](https://company.wizards.com/ja/legal/fancontentpolicy)に沿った非公式のファンコンテンツです。題材の一部に、ウィザーズ・オブ・ザ・コースト社の財産を含んでいます。&copy; Wizards of the Coast LLC.

# Everyone's Answer Board

協働回答ボードシステム - Google Apps Script (GAS) ベースのWebアプリケーション

[![License: ISC](https://img.shields.io/badge/License-ISC-blue.svg)](https://opensource.org/licenses/ISC)
[![GAS](https://img.shields.io/badge/Platform-Google%20Apps%20Script-4285F4?logo=google)](https://developers.google.com/apps-script)

## 概要

Everyone's Answer Boardは、Googleフォームの回答をリアルタイムで共有・表示するためのWebアプリケーションです。教育機関やチームでの協働作業に最適化されています。

### 主な機能

- Googleフォームとの自動連携
- リアルタイム回答表示
- 表示モードの切替: 掲示板 / 数直線 / マトリクス（2軸散布図） / ワードランキング / 円グラフ
  （`config.displaySettings.boardMode`。定義は `src/validators.js` の `BOARD_MODES`）
- 授業モード: フェーズ（考える→出会う→議論する→もう一度考える→ふりかえる）で
  児童画面の権能を切り替える。Google フォームを介さない native 入力にも対応
- リアクション機能（いいね、なるほど等）
- ハイライト機能
- 管理パネル（ボードの設定・授業の作成と進行・見取りグリッド）
- 複数ボード（プロファイル）の切り替え

### 技術スタック

- **バックエンド**: Google Apps Script (V8 runtime)
- **フロントエンド**: HTML + Tailwind CSS + Vanilla JavaScript
- **データストア**: 先生のスプレッドシート 1 つ（Google Sheets）
- **実行**: 先生ごとの配置を「先生として実行」（児童は自分の権限を使わない）
- **認証**: Session.getActiveUser（組織のドメイン内のみ）
- **デプロイ**: clasp CLI（`scripts/deploy-single.js`）

---

## 先生ごとの配置

先生ごとに、先生の Drive のスプレッドシートに紐づいたアプリを「先生として実行」で配置する。
データ（ボードの設定・授業・回答）はそのスプレッドシート 1 つに入る。サービスアカウントや共有の DB は使わない。

### 前提条件

- Node.js 20 以上
- 学校の Google Workspace アカウントで `clasp login` 済み（`~/.clasprc.json` の名前を `--user` に渡す）

### 先生を足す

```bash
git clone https://github.com/atariryuma/Everyone-s-Answer-Board.git
cd Everyone-s-Answer-Board
npm install
npm run deploy:single -- create --name <名前> --adminEmail <先生のメール> --user organization
```

1. 表示された URL を、その先生のアカウントで 1 度開いて「許可」を押す
2. 管理パネルが出れば準備完了（初回にプロパティと DB のシートが自動でそろう。ログイン・初期設定の画面は無い）
3. 児童には管理パネルの「ボードを開く」の URL（`?mode=view&userId=...`）を配る

### 更新する

```bash
npm test
npm run deploy             # 全員の配置を更新（URL はそのまま）
npm run api -- systemDiagnosis
```

---

## 開発者向け情報

### プロジェクト構成

```
Everyone-s-Answer-Board/
├── src/                      # GAS ソースコード
│   ├── main.js              # エントリーポイント、認証
│   ├── Deployment.js        # 配置ごとの先生と DB（リポジトリでは空。配置時に焼き込む）
│   ├── SystemController.js  # システム管理
│   ├── DatabaseCore.js      # users シートとアクセス判定
│   ├── *Service.js          # ビジネスロジック
│   ├── *Apis.js             # API エンドポイント
│   └── *.html               # フロントエンド
├── tests/                    # ユニットテスト（node:test）
├── scripts/                  # CLIツール
│   ├── deploy-single.js     # 先生ごとの配置（create / deploy）
│   ├── admin-api.js         # 配置への管理用コマンド
│   ├── migrate-to-single.js # 共有の本体からの引っ越し
│   ├── verify/              # 検証用の配置でだけ使う入口と、端から端までの検証
│   └── lib/gas-auth.js      # 共通認証モジュール
├── CLAUDE.md                # AI開発ガイド
└── README.md
```

### 開発コマンド

```bash
npm test                                   # テスト実行（node:test）
npm run deploy:single -- deploy --name <検証用の名前>   # 検証用の配置へ
npm run verify:e2e -- <検証用の名前>      # 画面の機能を端から端まで（54 項目）
npm run deploy                             # 全員の配置を更新

npm run api -- systemDiagnosis             # システム診断
npm run api -- getAppStatus                # アプリ状態
npm run api -- getLogs --limit 20          # セキュリティログ
npm run api -- listProperties              # Script Properties一覧
```

### 開発ワークフロー

```text
コード編集 → npm test → 検証用の配置 + verify:e2e → npm run deploy → git commit → git push
                                                                          ↓
                                                                GitHub CI（構文チェック+テスト）
```

- **CI**: 構文チェック + lint（error）+ テスト（品質ゲート）
- **テスト**: pre-commit フックで自動実行（`npm install` 時に自動有効化。テスト失敗時はコミットをブロック）

### GAS保守ルール

- `src/page.js.html` / `src/AdminPanel.js.html` は単一ファイル運用を維持
- トップレベル副作用を禁止（`google.script.run`・DOM操作は `init()` 内のみ）
- `doPost` の `action` は allowlist 管理し、追加時は入力検証を同時に追加
- `publishApp` は allowlist フィールドのみ受け付け、`etag` 競合検知を維持
- include順は固定し、変更する場合は単独コミットで扱う

---

## トラブルシューティング

**Q: 開くと「このアプリはまだ準備できていません」**
- 先生のアカウントで 1 度開いて許可したか確認（初回にプロパティがそろう）
- `npm run api -- systemDiagnosis` で診断を実行

**Q: 児童が「アクセスが制限されています」になる**
- 配った URL に `?mode=view&userId=...` が付いているか確認（付いていないトップの URL は先生用）
- ボードが「公開中」か確認

**Q: データが表示されない**
- `npm run api -- getUserConfig --userId <uuid>` でボードの設定（spreadsheetId / sheetName）を確認

### ログの確認

先生ごとの配置の実行ログは、スプレッドシート → 拡張機能 → Apps Script の「実行数」で見る。

---

## ライセンス

[ISC License](LICENSE)

## Author

Ryuma Atari

## リンク

- [GitHub Repository](https://github.com/atariryuma/Everyone-s-Answer-Board)
- [Issues](https://github.com/atariryuma/Everyone-s-Answer-Board/issues)
- [CLAUDE.md - AI開発ガイド](CLAUDE.md)

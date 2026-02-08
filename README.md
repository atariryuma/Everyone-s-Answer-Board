# Everyone's Answer Board

協働回答ボードシステム - Google Apps Script (GAS) ベースのWebアプリケーション

[![License: ISC](https://img.shields.io/badge/License-ISC-blue.svg)](https://opensource.org/licenses/ISC)
[![GAS](https://img.shields.io/badge/Platform-Google%20Apps%20Script-4285F4?logo=google)](https://developers.google.com/apps-script)

## 概要

Everyone's Answer Boardは、Googleフォームの回答をリアルタイムで共有・表示するためのWebアプリケーションです。教育機関やチームでの協働作業に最適化されています。

### 主な機能

- Googleフォームとの自動連携
- リアルタイム回答表示
- リアクション機能（いいね、なるほど等）
- ハイライト機能
- 管理者ダッシュボード
- 複数ユーザー・複数ボード対応

---

## システム管理者向け: デプロイ手順

### 前提条件

- Google Workspace アカウント（組織管理者権限推奨）
- Node.js 18以上
- npm または yarn

### Step 1: リポジトリのクローンと依存関係のインストール

```bash
git clone https://github.com/atariryuma/Everyone-s-Answer-Board.git
cd Everyone-s-Answer-Board
npm install
```

### Step 2: Google Cloud Console でプロジェクト設定

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成（または既存プロジェクトを選択）
3. **APIs & Services** > **Enable APIs and Services** で以下を有効化:
   - Google Sheets API
   - Google Drive API
   - Google Forms API

### Step 3: サービスアカウントの作成

1. **APIs & Services** > **Credentials** > **Create Credentials** > **Service Account**
2. サービスアカウント名を入力（例: `answer-board-service`）
3. 作成後、**Keys** タブで **Add Key** > **Create new key** > **JSON** を選択
4. ダウンロードしたJSONファイルを安全に保管（後で使用）

### Step 4: clasp のセットアップ

```bash
# clasp にログイン（Googleアカウント認証）
npx clasp login

# 新しいGASプロジェクトを作成
npx clasp create --type webapp --title "Everyone's Answer Board"

# または既存プロジェクトに接続
# .clasp.json.template を .clasp.json にコピーし、scriptId を設定
cp .clasp.json.template .clasp.json
# .clasp.json を編集して scriptId を設定
```

### Step 5: データベーススプレッドシートの準備

1. [Google Sheets](https://sheets.google.com/) で新しいスプレッドシートを作成
2. シート名を `users` に変更
3. 1行目にヘッダーを追加:
   ```
   userId | userEmail | userName | isActive | createdAt | lastModified | config
   ```
4. スプレッドシートIDをメモ（URLの `/d/` と `/edit` の間の文字列）
5. サービスアカウントのメールアドレス（`xxx@xxx.iam.gserviceaccount.com`）を編集者として共有

### Step 6: コードのデプロイ

```bash
# ソースコードをGASにプッシュ
npm run push

# GASエディタを開いて確認
npm run open
```

### Step 7: スクリプトプロパティの設定

GASエディタで **プロジェクトの設定** > **スクリプトプロパティ** を開き、以下を設定:

| プロパティ名 | 値 |
|-------------|-----|
| `ADMIN_EMAIL` | 管理者のメールアドレス |
| `DATABASE_SPREADSHEET_ID` | Step 5で作成したスプレッドシートのID |
| `SERVICE_ACCOUNT_CREDS` | Step 3でダウンロードしたJSONの内容（全文貼り付け） |

### Step 8: Webアプリとしてデプロイ

1. GASエディタで **デプロイ** > **新しいデプロイ**
2. **種類の選択** で **ウェブアプリ** を選択
3. 設定:
   - **説明**: バージョン説明（例: `v1.0.0 Initial release`）
   - **次のユーザーとして実行**: **ウェブアプリにアクセスしているユーザー**
   - **アクセスできるユーザー**: **全員**（または組織内のみ）
4. **デプロイ** をクリック
5. 表示されるURLが本番環境のアクセスURLです

### Step 9: 初期設定の確認

1. デプロイURLに `?mode=setup` を追加してアクセス
2. システム設定が正しく読み込まれていることを確認
3. `?mode=login` でログインページにアクセスし、管理者としてログイン

---

## 開発者向け情報

### プロジェクト構成

```
Everyone-s-Answer-Board/
├── src/                      # GAS ソースコード
│   ├── main.js              # エントリーポイント、認証
│   ├── SystemController.js  # システム管理
│   ├── DatabaseCore.js      # データベース操作
│   ├── *Service.js          # ビジネスロジック
│   ├── *Apis.js             # API エンドポイント
│   └── *.html               # フロントエンド
├── .clasp.json.template     # clasp 設定テンプレート
├── .claspignore             # clasp push 除外設定
├── package.json             # npm 設定
├── CLAUDE.md                # 開発ガイド（AI向け）
└── README.md                # このファイル
```

### 開発コマンド

```bash
npm run pull    # GAS からコードを取得
npm run push    # GAS にコードをプッシュ
npm run open    # GAS エディタを開く
npm run logs    # 実行ログを確認
npm run deploy  # 新バージョンをデプロイ
```

### 開発ワークフロー

1. `npm run pull` で最新コードを取得
2. ローカルでコードを編集
3. `npm run push` でGASにプッシュ
4. ブラウザでテスト（`/exec?mode=...`）
5. 動作確認後、Gitにコミット

### GAS保守ルール

- `src/page.js.html` / `src/AdminPanel.js.html` は当面「単一ファイル運用」を維持（無理な物理分割はしない）
- トップレベル副作用を禁止（`google.script.run`・DOM操作・自動実行タイマーは `init()` 内のみ）
- `doPost` の `action` は allowlist 管理し、追加時は入力検証を同時に追加
- `publishApp` は allowlist フィールドのみ受け付け、`etag` 競合検知を維持
- include順は固定し、変更する場合は単独コミットで扱う

### 技術スタック

- **バックエンド**: Google Apps Script (V8 runtime)
- **フロントエンド**: HTML + Tailwind CSS + Vanilla JavaScript
- **データストア**: Google Sheets
- **認証**: Google OAuth 2.0
- **デプロイ**: clasp CLI

### アーキテクチャ特徴

- **Zero-Dependency**: 外部ライブラリ不使用
- **Direct GAS API**: SpreadsheetApp, DriveApp 等を直接使用
- **Batch Operations**: 70倍のパフォーマンス改善
- **TTL Caching**: API呼び出しを80-90%削減

---

## トラブルシューティング

### よくある問題

**Q: デプロイ後にアクセスできない**
- スクリプトプロパティが正しく設定されているか確認
- サービスアカウントがスプレッドシートに共有されているか確認

**Q: ログインできない**
- `ADMIN_EMAIL` が正しく設定されているか確認
- ブラウザのCookieをクリアして再試行

**Q: データが表示されない**
- `DATABASE_SPREADSHEET_ID` が正しいか確認
- `users` シートが存在するか確認

### ログの確認

```bash
npm run logs
```

または、GASエディタで **実行数** > **実行ログ** を確認

---

## ライセンス

[ISC License](LICENSE)

## Author

Ryuma Atari

## リンク

- [GitHub Repository](https://github.com/atariryuma/Everyone-s-Answer-Board)
- [Issues](https://github.com/atariryuma/Everyone-s-Answer-Board/issues)
- [CLAUDE.md - 開発ガイド](CLAUDE.md)

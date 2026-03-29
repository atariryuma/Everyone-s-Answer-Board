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

### 技術スタック

- **バックエンド**: Google Apps Script (V8 runtime)
- **フロントエンド**: HTML + Tailwind CSS + Vanilla JavaScript
- **データストア**: Google Sheets
- **認証**: Google OAuth 2.0（Session.getActiveUser）
- **デプロイ**: clasp CLI

---

## システム管理者向け: デプロイ手順

### 前提条件

- Google Workspace アカウント（組織管理者権限推奨）
- Node.js 18以上
- npm

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
   - Apps Script API

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
npm run push        # GASにコードをプッシュ
npm run deploy:prod # 本番デプロイ（URL維持）
```

### Step 7: スクリプトプロパティの設定

GASエディタで **プロジェクトの設定** > **スクリプトプロパティ** を開き、以下を設定:

| プロパティ名 | 値 |
| ----------- | --- |
| `ADMIN_EMAIL` | 管理者のメールアドレス |
| `DATABASE_SPREADSHEET_ID` | Step 5で作成したスプレッドシートのID |
| `SERVICE_ACCOUNT_CREDS` | Step 3でダウンロードしたJSONの内容（全文貼り付け） |
| `ADMIN_API_KEY` | 任意の秘密文字列（16文字以上、CLIツール認証用） |

または、SetupPage（`デプロイURL?mode=setup`）から設定可能です。

### Step 8: Webアプリとしてデプロイ

1. GASエディタで **デプロイ** > **新しいデプロイ**
2. **種類の選択** で **ウェブアプリ** を選択
3. 設定:
   - **次のユーザーとして実行**: **ウェブアプリにアクセスしているユーザー**
   - **アクセスできるユーザー**: 組織のポリシーに合わせて選択
4. **デプロイ** をクリック

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
├── tests/                    # ユニットテスト（74件）
├── scripts/                  # CLIツール
│   ├── admin-api.js         # 本番API操作
│   ├── deploy-prod.js       # URL維持デプロイ
│   ├── logs.js              # Cloud Logging取得
│   └── lib/gas-auth.js      # 共通認証モジュール
├── .clasp.json.template     # clasp 設定テンプレート
├── CLAUDE.md                # AI開発ガイド
└── README.md
```

### 開発コマンド

```bash
# 開発
npm test                  # テスト実行（74件）
npm run push              # GASにコードをプッシュ
npm run deploy:prod       # 本番デプロイ（URL維持、pushも含む）
npm run deploy            # 新しいURLでデプロイ（テスト用）
npm run pull              # GASからコードを取得
npm run logs              # GAS実行ログを確認

# 本番運用
npm run api -- systemDiagnosis     # システム診断
npm run api -- getUsers            # ユーザー一覧
npm run api -- getAppStatus        # アプリ状態
npm run api -- getLogs --limit 20  # セキュリティログ
npm run api -- perfMetrics         # パフォーマンス
npm run api -- listProperties      # Script Properties一覧
npm run logs:cloud                 # Cloud Logging（直近の警告/エラー）
npm run logs:cloud -- --severity ERROR --hours 24  # エラーのみ
```

### 開発ワークフロー

```text
コード編集 → npm test → npm run deploy:prod → git commit → git push
                                                              ↓
                                                    GitHub CI（構文チェック+テスト）
```

- **CI**: 構文チェック + テストのみ（品質ゲート）
- **デプロイ**: ローカルの `deploy:prod` で実行（CLIから本番URLを維持して更新）
- **テスト**: pre-pushフックで自動実行（テスト失敗時はpushブロック）

### GAS保守ルール

- `src/page.js.html` / `src/AdminPanel.js.html` は単一ファイル運用を維持
- トップレベル副作用を禁止（`google.script.run`・DOM操作は `init()` 内のみ）
- `doPost` の `action` は allowlist 管理し、追加時は入力検証を同時に追加
- `publishApp` は allowlist フィールドのみ受け付け、`etag` 競合検知を維持
- include順は固定し、変更する場合は単独コミットで扱う

---

## トラブルシューティング

### よくある問題

**Q: デプロイ後にアクセスできない**
- スクリプトプロパティが正しく設定されているか確認
- サービスアカウントがスプレッドシートに共有されているか確認
- `npm run api -- systemDiagnosis` で診断を実行

**Q: ログインできない**
- `ADMIN_EMAIL` が正しく設定されているか確認
- ドメイン制限が有効な場合、同一ドメインのアカウントでアクセス

**Q: データが表示されない**
- `DATABASE_SPREADSHEET_ID` が正しいか確認
- `users` シートが存在するか確認
- `npm run api -- getUsers` でユーザー状態を確認

**Q: リアクションがエラーになる**

- `npm run logs:cloud -- --severity ERROR --hours 1` でエラーログを確認
- スプレッドシートの共有設定（同一ドメイン内の編集権限）を確認

### ログの確認

```bash
npm run logs:cloud                              # 直近の警告/エラー
npm run logs:cloud -- --severity INFO --hours 24 # 過去24時間の全ログ
npm run logs:cloud -- --function doPost          # 特定関数のログ
```

---

## ライセンス

[ISC License](LICENSE)

## Author

Ryuma Atari

## リンク

- [GitHub Repository](https://github.com/atariryuma/Everyone-s-Answer-Board)
- [Issues](https://github.com/atariryuma/Everyone-s-Answer-Board/issues)
- [CLAUDE.md - AI開発ガイド](CLAUDE.md)

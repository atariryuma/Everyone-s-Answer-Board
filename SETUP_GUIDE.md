# StudyQuest ログイン設定ガイド

## 必要な設定項目

ログイン機能を正常に動作させるために、以下の設定項目をGoogle Apps Scriptのスクリプトプロパティに設定する必要があります。

### 1. スクリプトプロパティの設定方法

1. Google Apps Script エディターを開く
2. 左メニューの「設定」(⚙️)をクリック
3. 「スクリプトプロパティ」タブを選択
4. 以下の項目を「プロパティ」と「値」に設定

### 2. 必須設定項目

#### `GOOGLE_CLIENT_ID`
- **説明**: Google OAuth 2.0 クライアントID
- **取得方法**: 
  1. [Google Cloud Console](https://console.cloud.google.com)にアクセス
  2. プロジェクトを選択または作成
  3. 「APIとサービス」→「認証情報」
  4. 「認証情報を作成」→「OAuth 2.0 クライアントID」
  5. アプリケーションの種類で「ウェブアプリケーション」を選択
  6. 承認済みのリダイレクトURIにGoogle Apps ScriptのURLを追加
- **形式**: `123456789-abcdefghijklmnopqrstuvwxyz.apps.googleusercontent.com`

#### `DATABASE_SPREADSHEET_ID`
- **説明**: ユーザーデータベースとして使用するGoogle SpreadsheetのID
- **取得方法**: 
  1. Google Driveで新しいスプレッドシートを作成
  2. URLの`/d/`と`/edit`の間の文字列をコピー
- **形式**: `1abcDefGhIjKlMnOpQrStUvWxYz1234567890AbCdEfG`

#### `ADMIN_EMAIL`
- **説明**: システム管理者のメールアドレス
- **形式**: `admin@example.com`

#### `SERVICE_ACCOUNT_CREDS`
- **説明**: サービスアカウントの認証情報（JSON形式）
- **取得方法**: 
  1. Google Cloud Consoleで「APIとサービス」→「認証情報」
  2. 「認証情報を作成」→「サービスアカウント」
  3. サービスアカウントを作成後、「キー」タブで「新しいキーを作成」
  4. JSON形式でダウンロード
  5. ダウンロードしたJSONファイルの内容全体をコピー
- **形式**: `{"type": "service_account", "project_id": "...", ...}`

### 3. 設定確認方法

1. ログイン画面にアクセス
2. コンソールに表示される初期化ログを確認
3. 設定が不完全な場合、「システム設定を確認」ボタンが表示される
4. ボタンをクリックして詳細な設定状況を確認

### 4. トラブルシューティング

#### 「GOOGLE_CLIENT_ID not found」エラー
- スクリプトプロパティに`GOOGLE_CLIENT_ID`が設定されていない
- 設定後、ページを再読み込みしてください

#### 「System configuration incomplete」エラー
- 必須設定項目のうち、いくつかが未設定です
- 「システム設定を確認」ボタンで不足している項目を確認してください

#### Google OAuth認証エラー
- `GOOGLE_CLIENT_ID`が正しく設定されているか確認
- Google Cloud Consoleで認証情報の設定を確認
- リダイレクトURIが正しく設定されているか確認

### 5. セキュリティ上の注意

- `SERVICE_ACCOUNT_CREDS`には機密情報が含まれます
- スクリプトプロパティは安全に管理されますが、コードには直接記述しないでください
- 定期的に認証情報を更新することを推奨します

## 設定完了後の確認

すべての設定が完了すると、ログイン画面で以下の動作が確認できます：

1. ✅ システム設定チェック完了
2. ✅ Google APIs読み込み完了
3. ✅ Google Client ID取得完了
4. ✅ Google認証初期化完了
5. ✅ ログインフロー開始完了

これらのステップがすべて完了すると、ユーザーのメールアドレスが表示され、ログイン機能が正常に動作します。
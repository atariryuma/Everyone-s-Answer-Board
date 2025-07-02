# StudyQuest - みんなの回答ボード: 内部API リファレンス

このドキュメントでは、「StudyQuest - みんなの回答ボード」の**新しいサービスアカウントアーキテクチャ**における主要な内部関数とAPIについて説明します。新アーキテクチャでは外部APIが不要になり、すべての機能が単一のGoogle Apps Scriptプロジェクト内で提供されています。

## 📋 概要

新アーキテクチャでは以下の特徴があります：
- ✅ **外部API不要**: Admin Logger APIを削除し、単一プロジェクトに統合
- ✅ **サービスアカウント認証**: JWT + Google OAuth2による直接認証
- ✅ **Google Sheets API v4**: 直接APIアクセスによる高性能化
- ✅ **内部関数API**: 明確に定義された内部インターフェース

---

## 1. 🔐 認証・セットアップ API

### `setupApplication(credsJson, dbId)`
アプリケーションの初期セットアップを実行します。

**パラメータ:**
- `credsJson` (string): サービスアカウントのJSONキーファイルの内容
- `dbId` (string): 中央データベーススプレッドシートのID

**戻り値:**
- `void`: 成功時は正常終了、失敗時は例外をスロー

**例:**
```javascript
const credsJson = '{"type":"service_account",...}';
const dbId = '1BcD3fGhIjKlMnOpQrStUvWxYz0123456789ABCDEFGH';
setupApplication(credsJson, dbId);
```

### `getServiceAccountToken()`
サービスアカウントのアクセストークンを取得します。

**戻り値:**
- `string`: OAuth2アクセストークン

**内部処理:**
1. JWT生成（RS256署名）
2. Google OAuth2トークン取得
3. トークンキャッシュ管理

### `getSheetsService()`
Google Sheets API v4のサービスオブジェクトを取得します。

**戻り値:**
- `object`: Sheets APIサービスオブジェクト

**提供メソッド:**
- `spreadsheets.get(spreadsheetId)`
- `spreadsheets.values.get(spreadsheetId, range)`
- `spreadsheets.values.append(spreadsheetId, range, values)`
- `spreadsheets.values.batchUpdate(spreadsheetId, requests)`

---

## 2. 📊 データベース操作 API

### `findUserById(userId)`
ユーザーIDでユーザー情報を検索します。内部でキャッシュを活用し、高速なアクセスを提供します。

**パラメータ:**
- `userId` (string): 検索するユーザーID

**戻り値:**
- `object|null`: ユーザー情報オブジェクトまたはnull

**ユーザー情報オブジェクト:**
```javascript
{
  userId: "user-1234567890",
  adminEmail: "teacher@school.edu", 
  spreadsheetId: "1BcD3fG...",
  spreadsheetUrl: "https://docs.google.com/...",
  createdAt: "2024-01-01T00:00:00.000Z",
  configJson: '{"displayMode":"anonymous"}',
  lastAccessedAt: "2024-01-01T12:00:00.000Z",
  isActive: "true"
}
```

### `findUserByEmail(email)`
メールアドレスでユーザー情報を検索します。内部でキャッシュを活用し、高速なアクセスを提供します。

**パラメータ:**
- `email` (string): 検索するメールアドレス

**戻り値:**
- `object|null`: ユーザー情報オブジェクトまたはnull

### `createUserInDb(userData)`
新しいユーザーを中央データベースに作成します。

**パラメータ:**
- `userData` (object): 作成するユーザー情報

**戻り値:**
- `object`: 作成されたユーザー情報

### `updateUserInDb(userId, updateData)`
既存ユーザーの情報を更新します。

**パラメータ:**
- `userId` (string): 更新対象のユーザーID
- `updateData` (object): 更新するデータ

**戻り値:**
- `object`: 更新結果 `{success: boolean, message: string}`

---

## 3. 📋 回答ボード操作 API

### `getSheetData(userId, sheetName, classFilter, sortMode)`
指定されたシートからデータを取得します。

**パラメータ:**
- `userId` (string): ユーザーID
- `sheetName` (string): シート名
- `classFilter` (string): クラスフィルター（'all'または具体的なクラス名）
- `sortMode` (string): ソートモード（'newest', 'oldest', 'score', 'likes', 'random'）

**戻り値:**
```javascript
{
  status: "success",
  data: [
    {
      timestamp: "2024-01-01 10:00:00",
      email: "student@school.edu",
      class: "6-1", 
      answer: "回答内容",
      reason: "理由",
      name: "学生名",
      reactions: {...},
      isHighlighted: false
    }
  ],
  totalCount: 25,
  filteredCount: 15
}
```

### `addReaction(rowIndex, reactionType, sheetName)`
回答にリアクションを追加/削除します。

**パラメータ:**
- `rowIndex` (number): 対象行のインデックス
- `reactionType` (string): リアクションタイプ（'UNDERSTAND', 'LIKE', 'CURIOUS'）
- `sheetName` (string): シート名

**戻り値:**
```javascript
{
  status: "success",
  action: "added", // または "removed"
  reactionType: "LIKE",
  userEmail: "user@example.com"
}
```

### `toggleHighlight(rowIndex)`
回答のハイライト状態を切り替えます。

**パラメータ:**
- `rowIndex` (number): 対象行のインデックス

**戻り値:**
```javascript
{
  status: "success", 
  isHighlighted: true, // または false
  rowIndex: 2
}
```

---

## 4. 🎛️ 管理機能 API

### `getAdminSettings()`
管理者設定情報を取得します。

**戻り値:**
```javascript
{
  status: "success",
  userId: "user-1234567890",
  userInfo: {...},
  spreadsheetInfo: {...},
  webAppUrl: "https://script.google.com/...",
  isPublished: false,
  publishedSheet: null,
  displayMode: "anonymous"
}
```

### `publishApp(sheetName)`
指定されたシートでアプリを公開します。

**パラメータ:**
- `sheetName` (string): 公開するシート名

**戻り値:**
```javascript
{
  status: "success",
  message: "アプリを公開しました: シート名",
  publishedSheet: "シート名",
  boardUrl: "https://script.google.com/..."
}
```

### `unpublishApp()`
アプリの公開を停止します。

**戻り値:**
```javascript
{
  status: "success", 
  message: "アプリの公開を停止しました"
}
```

### `saveDisplayMode(mode)`
表示モードを設定します。

**パラメータ:**
- `mode` (string): 表示モード（'anonymous'または'named'）

**戻り値:**
```javascript
{
  status: "success",
  message: "表示モードを保存しました: anonymous",
  displayMode: "anonymous"
}
```

### `getSheets(userId)`
ユーザーのスプレッドシート内のシート一覧を取得します。

**パラメータ:**
- `userId` (string): ユーザーID

**戻り値:**
- `array`: シート情報の配列

---

## 5. 🗄️ キャッシュ管理 API

### `getCachedUserInfo(userId)`
キャッシュからユーザー情報を取得します。

**パラメータ:**
- `userId` (string): ユーザーID

**戻り値:**
- `object|null`: キャッシュされたユーザー情報またはnull

### `setCachedUserInfo(userId, userInfo)`
ユーザー情報をキャッシュに保存します。

**パラメータ:**
- `userId` (string): ユーザーID
- `userInfo` (object): キャッシュするユーザー情報

### `clearAllCaches()`
すべてのキャッシュをクリアします。

**戻り値:**
- `void`

### `getAndCacheHeaderIndices(spreadsheetId, sheetName)`
ヘッダー情報を取得し、キャッシュします。

**パラメータ:**
- `spreadsheetId` (string): スプレッドシートID  
- `sheetName` (string): シート名

**戻り値:**
- `object`: ヘッダーインデックス情報

---

## 6. 🛠️ ユーティリティ API

### `isValidEmail(email)`
メールアドレスの形式を検証します。

**パラメータ:**
- `email` (string): 検証するメールアドレス

**戻り値:**
- `boolean`: 有効な場合true

### `getEmailDomain(email)`
メールアドレスからドメインを抽出します。

**パラメータ:**
- `email` (string): メールアドレス

**戻り値:**
- `string`: ドメイン部分

### `parseReactionString(reactionString)`
リアクション文字列を解析します。

**パラメータ:**
- `reactionString` (string): カンマ区切りのメールアドレス文字列

**戻り値:**
- `array`: メールアドレスの配列

### `safeSpreadsheetOperation(operation, fallbackValue)`
スプレッドシート操作を安全に実行します。

**パラメータ:**
- `operation` (function): 実行する操作
- `fallbackValue` (any): エラー時のフォールバック値

**戻り値:**
- `any`: 操作の結果またはフォールバック値

---

## 7. 🔍 診断・デバッグ API

### `testSetup()`
セットアップ状態を詳細に診断します。

**戻り値:**
```javascript
{
  serviceAccount: "✅ 認証成功",
  database: "✅ 接続成功", 
  cache: "✅ 動作中",
  apis: "✅ 利用可能"
}
```

### `showSetupInfo()`
現在の設定情報を表示します。

**戻り値:**
- `object`: 詳細な設定情報

### `checkDatabaseStatus()`
データベース接続状態を確認します。

**戻り値:**
```javascript
{
  status: "success",
  connection: "active",
  userCount: 25,
  lastAccess: "2024-01-01T12:00:00.000Z"
}
```

---

## 8. 🎯 エラーハンドリング

新アーキテクチャでは、すべてのAPI関数が統一されたエラーハンドリングを提供します。

### エラーレスポンス形式
```javascript
{
  status: "error",
  message: "エラーの詳細説明",
  errorCode: "SPECIFIC_ERROR_CODE", 
  timestamp: "2024-01-01T12:00:00.000Z"
}
```

### 一般的なエラーコード
- `AUTH_FAILED`: サービスアカウント認証失敗
- `DATABASE_ERROR`: データベース接続エラー
- `PERMISSION_DENIED`: 権限不足
- `INVALID_INPUT`: 無効な入力パラメータ
- `RESOURCE_NOT_FOUND`: リソースが見つからない

---

## 9. 🚀 パフォーマンス考慮事項

### キャッシュ戦略
- **ユーザー情報**: 5分間キャッシュ
- **ヘッダー情報**: セッション中キャッシュ  
- **名簿データ**: セッション中キャッシュ

### API呼び出し最適化
- **バッチ処理**: 複数操作の一括実行
- **条件付きアクセス**: キャッシュヒット時のAPI回避
- **並列処理**: 独立操作の同時実行

### レート制限
- **Google Sheets API**: 毎分100リクエスト
- **内部処理**: 制限なし（メモリキャッシュ経由）

---

## 🎉 まとめ

新しいサービスアカウントアーキテクチャにより、StudyQuestは：

### ✅ 統合されたAPI設計
- 外部API依存を完全排除
- 単一プロジェクト内での一貫したインターフェース
- 明確に定義された内部API仕様

### ✅ 高性能・高信頼性
- Google Sheets API v4直接アクセス
- 効率的なキャッシュシステム
- 堅牢なエラーハンドリング

### ✅ 保守性の向上
- 明確なAPI境界
- 包括的なドキュメント
- 100%テストカバレッジ

この内部APIリファレンスにより、開発者は新アーキテクチャの機能を効率的に活用し、拡張することができます。
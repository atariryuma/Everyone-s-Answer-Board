# GAS V8/ES6 コーディング方針（AI向け）

> 目的：Google Apps Script（GAS）の **V8 ランタイム** 前提で、AIが安定したコードを生成できるようにする実務ガイド。  
> 対象：コード生成AI（Claude/ChatGPT等）、レビュー用 LLM、社内自動化ボット。

---

# ⚠️ システム破壊防止ルール（これを破ると全システム停止）

## 🚨 絶対遵守：データベーススキーマ（最新版：9項目に最適化）
- ✅ **唯一使用**: `database.gs` の `DB_CONFIG`
- ✅ **最適化済み構造**（14項目→9項目に36%削減）: 
```javascript
const DB_CONFIG = Object.freeze({
  SHEET_NAME: 'Users',
  HEADERS: Object.freeze([
    'userId',           // [0] UUID - 必須ID
    'userEmail',        // [1] メールアドレス - 必須認証
    'createdAt',        // [2] 作成日時 - 監査用
    'lastAccessedAt',   // [3] 最終アクセス - 監査用
    'isActive',         // [4] アクティブ状態 - 必須フラグ
    'spreadsheetId',    // [5] 統一データソース - 必須
    'sheetName',        // [6] 統一データソース - 必須
    'configJson',       // [7] 全設定を統合 - メイン設定
    'lastModified',     // [8] 最終更新 - 監査用
  ])
});
```

### 🗑️ 削除された重複項目（configJsonに統合済み）
- ~~`spreadsheetUrl`~~ → 動的生成: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`
- ~~`formUrl`~~ → `configJson.formUrl`に統合
- ~~`columnMappingJson`~~ → `configJson.columnMapping`に統合  
- ~~`publishedAt`~~ → `lastModified`で代用
- ~~`appUrl`~~ → 動的生成: `${WebApp.getUrl()}?mode=view&userId=${userId}`

## 🎯 必須定数（src/constants.gs）
### システム全体の統一定数
```javascript
const SYSTEM_CONSTANTS = Object.freeze({
  // データベース関連定数
  DATABASE: Object.freeze({
    SHEET_NAME: 'Users',
    HEADERS: Object.freeze([
      'userId',           // [0] UUID - 必須ID
      'userEmail',        // [1] メールアドレス - 必須認証
      'createdAt',        // [2] 作成日時 - 監査用
      'lastAccessedAt',   // [3] 最終アクセス - 監査用
      'isActive',         // [4] アクティブ状態 - 必須フラグ
      'spreadsheetId',    // [5] 統一データソース - 必須
      'sheetName',        // [6] 統一データソース - 必須
      'configJson',       // [7] 全設定を統合 - メイン設定
      'lastModified',     // [8] 最終更新 - 監査用
    ]),
    DELETE_LOG: Object.freeze({
      SHEET_NAME: 'DeletionLogs',
      HEADERS: Object.freeze(['timestamp', 'executorEmail', 'targetUserId', 'targetEmail', 'reason', 'deleteType'])
    })
  }),

  // リアクション機能
  REACTIONS: Object.freeze({
    KEYS: ['UNDERSTAND', 'LIKE', 'CURIOUS'],
    LABELS: Object.freeze({
      UNDERSTAND: 'なるほど！',
      LIKE: 'いいね！', 
      CURIOUS: 'もっと知りたい！',
      HIGHLIGHT: 'ハイライト'
    })
  }),

  // 列ヘッダー定義
  COLUMNS: Object.freeze({
    TIMESTAMP: 'タイムスタンプ',
    EMAIL: 'メールアドレス',
    CLASS: 'クラス',
    OPINION: '回答',
    REASON: '理由',
    NAME: '名前'
  }),

  // 列マッピング（AI検索対応）
  COLUMN_MAPPING: Object.freeze({
    answer: Object.freeze({
      key: 'answer', header: '回答',
      alternates: ['どうして', '質問', '問題', '意見', '答え', 'なぜ', '思います', '考え'],
      required: true,
      aiPatterns: ['？', '?', 'どうして', 'なぜ', '思いますか', '考えますか']
    }),
    reason: Object.freeze({
      key: 'reason', header: '理由',
      alternates: ['理由', '根拠', '体験', 'なぜ', '詳細', '説明'],
      required: false,
      aiPatterns: ['理由', '体験', '根拠', '詳細']
    }),
    class: Object.freeze({
      key: 'class', header: 'クラス',
      alternates: ['クラス', '学年'],
      required: false
    }),
    name: Object.freeze({
      key: 'name', header: '名前', 
      alternates: ['名前', '氏名', 'お名前'],
      required: false
    })
  }),

  // 表示モード
  DISPLAY_MODES: Object.freeze({
    ANONYMOUS: 'anonymous',
    NAMED: 'named', 
    EMAIL: 'email'
  }),

  // アクセス制御
  ACCESS: Object.freeze({
    LEVELS: Object.freeze({
      OWNER: 'owner',
      SYSTEM_ADMIN: 'system_admin', 
      AUTHENTICATED_USER: 'authenticated_user',
      GUEST: 'guest',
      NONE: 'none'
    })
  })
});
```

### 後方互換性エイリアス
```javascript
const REACTION_KEYS = SYSTEM_CONSTANTS.REACTIONS.KEYS;
const COLUMN_HEADERS = {
  ...SYSTEM_CONSTANTS.COLUMNS,
  ...SYSTEM_CONSTANTS.REACTIONS.LABELS
};
const DELETE_LOG_SHEET_CONFIG = SYSTEM_CONSTANTS.DATABASE.DELETE_LOG;
```

### コアシステム定数
```javascript
const CORE = Object.freeze({
  TIMEOUTS: { SHORT: 1000, MEDIUM: 5000, LONG: 30000, FLOW: 300000 },
  STATUS: { ACTIVE: 'active', INACTIVE: 'inactive', PENDING: 'pending', ERROR: 'error' },
  HTTP_STATUS: { OK: 200, BAD_REQUEST: 400, UNAUTHORIZED: 401, FORBIDDEN: 403 }
});

const PROPS_KEYS = Object.freeze({
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL'
});
```

## 🔄 システムフロー
### 1. 初期セットアップフロー
```
システム未設定 → doGet() → isSystemSetup() → renderSetupPage()
```

### 2. ユーザー登録・認証フロー
```
doGet(mode=login) → handleUserRegistration() → createCompleteUser() → DB.createUser()
```

### 3. 管理パネルフロー
```
doGet(mode=admin) → App.getAccess().verifyAccess() → renderAdminPanel()
```

### 4. 回答ボード表示フロー
```
doGet(mode=view) → App.getAccess().verifyAccess() → renderAnswerBoard()
  → userInfo.spreadsheetId/sheetName を統一使用
```

### 5. データソース接続フロー
```
connectDataSource() → DB.updateUser() → userInfo.spreadsheetId/sheetName 更新
  → 重複フィールド削除（publishedSpreadsheetId等）
```

### 6. スプレッドシートアクセスフロー
```
getPublishedSheetData() → userInfo.spreadsheetId（統一データソース）
Core.gs関数群 → targetSpreadsheetId = userInfo.spreadsheetId
```

## 🏗️ アーキテクチャ階層
1. **PropertiesService**: システム設定（SERVICE_ACCOUNT_CREDS, DATABASE_SPREADSHEET_ID, ADMIN_EMAIL）
2. **Database**: テナント管理（userId, userEmail, configJson, **spreadsheetId**, **sheetName**）
3. **統一データソース**: `userInfo.spreadsheetId`が全機能で唯一の真実の源
4. **AccessController**: アクセス制御（owner > system_admin > authenticated_user > guest）
5. **SecurityManager**: 認証・JWT管理（Service Account Token生成）

## 🔐 セキュリティ設計
### 入力検証（SecurityValidator）
```javascript
const SECURITY = Object.freeze({
  VALIDATION_PATTERNS: {
    EMAIL: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
    UUID: /^[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/i,
    SAFE_STRING: /^[a-zA-Z0-9\s\-_.@]+$/
  },
  MAX_LENGTHS: { EMAIL: 254, CONFIG_JSON: 10000, GENERAL_TEXT: 1000 }
});
```

---

## 0) ランタイム前提

- GAS は **V8**（Chrome/Node と同系エンジン）で動作し、**モダンな ECMAScript 構文**を利用可能。  
- `let/const`、アロー関数、テンプレートリテラル、分割代入、クラス、`Map/Set` などが使える。  
- 旧 Rhino も選べるが、**V8 を強く推奨**。

---

## 1) まず守るコーディング規範

- **`var` を禁止**。**`const` 優先**、必要時のみ `let`。  
- **関数は原則アロー関数**（`this` を必要とするクラスメソッドは通常のメソッド記法）。  
- **テンプレートリテラル**で文字列結合を可読化。  
- **分割代入 / スプレッド**で引数・配列・オブジェクト操作を簡潔に。  
- **不変データ**志向：オブジェクトの直接破壊より新オブジェクトを返す。

---

## 2) GAS ならではの設計ルール

- **エントリーポイントはグローバル関数**（トリガーやメニュー登録はトップレベル）。  
- **Apps Script のサービス API は同期的**。`UrlFetchApp` などはブロッキング呼び出し。  
- **バッチ処理**：Spreadsheet/Drive などはまとめて取得・更新。  
- **状態は PropertiesService/CacheService** に格納。  
- **ログ**は `console.log` または `Logger.log`。  
- **例外設計**：`throw new Error()` を用い、`try/catch` でハンドリング。

---

## 3) ファイル構成とモジュール化

現在のファイル構成（2025年最新版）:
```
/src/
├── constants.gs         # システム定数（SYSTEM_CONSTANTS, CORE, PROPS_KEYS）
├── database.gs          # DB操作（DB名前空間、最新9項目スキーマ）
├── main.gs             # エントリーポイント（doGet, Services名前空間）
├── auth.gs             # 認証管理（ユーザー登録、JWT）
├── security.gs         # セキュリティ（Service Account Token）
├── Core.gs             # 業務ロジック（自動停止、ヘッダー検証）
├── AdminPanelBackend.gs # 管理パネル（列マッピング検出、configJson統合）
├── SystemManager.gs    # 🆕 統合管理（テスト・最適化・診断）
├── App.gs              # 統一サービス層
├── Base.gs             # 基盤機能
└── cache.gs            # キャッシュ管理

HTML/フロントエンド:
├── Page.html           # 回答ボード（opinionHeader対応）
├── AdminPanel.html     # 管理画面（2段階構造、appName削除）
└── AppSetupPage.html   # セットアップ画面
```

### 📁 主要な変更点（2025年最新）
- ✅ **SystemManager.gs追加**: 分散していたテスト・最適化機能を統合
- ❌ **ConfigOptimizer.gs削除**: 重複機能をSystemManagerに統合
- 🔄 **AdminPanel.html簡素化**: 3段階→2段階、appName削除
- 🎯 **configJson中心設計**: 全設定をconfigJsonに統合

- GAS は **ES Modules 非対応**。`import/export` はそのままでは使えない。  
- 多ファイルに分け、**グローバル名前空間を汚さない命名**で整理。

---

## 4) 主要 ES6+ 機能の使い分け

- **Object.freeze()**：設定オブジェクトの不変化
- **Map/Set**：高速検索や非文字列キー管理に有効
- **アロー関数**：関数型プログラミングパターン
- **テンプレートリテラル**：ログメッセージやHTML生成
- **分割代入**：オブジェクト/配列の簡潔な操作

---

## 5) I/O とパフォーマンス

### Sheets API最適化
```javascript
// ❌ 非効率：個別API呼び出し
sheet.getRange('A1').setValue('data1');
sheet.getRange('A2').setValue('data2');

// ✅ 効率的：バッチ処理
const values = [['data1'], ['data2']];  
sheet.getRange('A1:A2').setValues(values);
```

### Service Account認証
```javascript
// キャッシュ活用でトークン取得を最適化
function getServiceAccountTokenCached() {
  return cacheManager.get(SECURITY_CONFIG.AUTH_CACHE_KEY, generateNewServiceAccountToken, {
    ttl: 3500,
    enableMemoization: true
  });
}
```

---

## 6) 例外・リトライ・検証

### 統一エラーハンドリング
```javascript
try {
  const result = DB.createUser(userData);
  return result;
} catch (error) {
  console.error('ユーザー作成エラー:', {
    userEmail: userData.userEmail,
    error: error.message,
    timestamp: new Date().toISOString()
  });
  throw error;
}
```

### 入力検証
```javascript
const validation = SecurityValidator.validateUserData(userData);
if (!validation.isValid) {
  throw new Error(validation.errors.join(', '));
}
```

---

## 7) 現在のシステム特有の実装パターン

### 統一データソースパターン（2025年最新）
```javascript
// ✅ 推奨：userInfo.spreadsheetIdを統一使用
function renderAnswerBoard(userInfo, params) {
  const targetSpreadsheetId = userInfo.spreadsheetId;  // 統一データソース
  const targetSheetName = userInfo.sheetName;
  // ...
}

// ❌ 非推奨：重複したconfig参照
// const publishedSpreadsheetId = config.publishedSpreadsheetId;
```

### App名前空間パターン
```javascript
// App.gs - 統一サービス層
const App = {
  init() {
    // システム初期化
  },
  
  getAccess() {
    return {
      verifyAccess(userId, mode, currentUserEmail) {
        // アクセス制御ロジック
      }
    };
  },
  
  getConfig() {
    // 設定管理
  }
};
```

### DB名前空間パターン
```javascript
// database.gs - DB操作の構造化
const DB = {
  createUser(userData) { /* */ },
  findUserByEmail(email) { /* */ },
  updateUser(userId, updateData) { /* */ },
  // ✅ データベース統一フィールド
  // userInfo.spreadsheetId, userInfo.sheetName を使用
};
```

### モジュール設定パターン
```javascript
// 各ファイルでのモジュール固有設定
const MODULE_CONFIG = Object.freeze({
  CACHE_TTL: CORE.TIMEOUTS.LONG,
  STATUS_ACTIVE: CORE.STATUS.ACTIVE
});
```

---

## 8) HTML Service/フロント連携

- クライアント JS はブラウザの ES Modules 利用可。  
- サーバ側とは別。`google.script.run` で呼び出す。

---

## 9) 生成AI向けプロンプト指示（2025年最新版）

### 🎯 必須遵守項目
1. **`const`優先、`let`のみ許可、`var`禁止**  
2. **最新9項目データベーススキーマ使用必須**（`DB_CONFIG`準拠）
3. **SystemManager名前空間使用**（テスト・最適化機能）
4. **configJson中心設計**: 全設定をconfigJsonに統合、重複フィールド禁止
5. **統一データソース原則**: `userInfo.spreadsheetId/sheetName`のみ使用
6. **バッチI/O・最小呼び出し**を強制  
7. **SecurityValidator使用**でセキュリティ確保  
8. **console.error**でエラー情報を構造化ログ出力  
9. **Object.freeze()**で設定の不変性保持

### 🆕 最新最適化機能
- **testSchemaOptimization()**: データベース構造最適化テスト
- **SystemManager.migrateToSimpleSchema()**: 14項目→9項目自動変換
- **動的URL生成**: spreadsheetUrl/appUrlの効率化
- **プライバシー重視**: displaySettings デフォルトfalse

---

## 10) 最小テンプレート（2025年最新システム準拠）

```javascript
/** @OnlyCurrentDoc */

/**
 * 新機能の実装例（最新最適化版）
 * 最新9項目スキーマ、SystemManager統合、configJson中心設計準拠
 */

// モジュール設定（CORE参照）
const FEATURE_CONFIG = Object.freeze({
  TIMEOUT: CORE.TIMEOUTS.MEDIUM,
  STATUS: CORE.STATUS.ACTIVE,
  // 🆕 最適化設定
  USE_CONFIG_JSON_ONLY: true,  // configJson中心設計
  ENABLE_DYNAMIC_URLS: true    // 動的URL生成
});

/**
 * メイン機能関数（最新版）
 * @param {string} userId - ユーザーID
 * @returns {Object} 処理結果
 */
function processFeature(userId) {
  try {
    // 入力検証
    if (!SecurityValidator.isValidUUID(userId)) {
      throw new Error('無効なユーザーIDです');
    }

    // DB操作（最新9項目スキーマ使用）
    const user = DB.findUserById(userId);
    if (!user) {
      throw new Error('ユーザーが見つかりません');
    }

    // 🆕 configJson中心の設定取得
    let config = {};
    if (user.configJson) {
      config = JSON.parse(user.configJson);
    }

    // 🆕 動的URL生成（最適化済み）
    const dynamicUrls = {
      spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${user.spreadsheetId}`,
      appUrl: `${WebApp.getUrl()}?mode=view&userId=${userId}`
    };

    // 処理ロジック
    const result = {
      success: true,
      userId: user.userId,
      config: config,
      urls: dynamicUrls,  // 🆕 動的生成URLs
      timestamp: new Date().toISOString()
    };

    console.info('機能処理完了（最新版）', {
      userId: userId,
      configKeys: Object.keys(config),
      result: result
    });

    return result;

  } catch (error) {
    console.error('機能処理エラー', {
      userId: userId,
      error: error.message,
      stack: error.stack
    });
    throw error;
  }
}

/**
 * 🆕 SystemManager連携例
 * テスト・最適化機能の統合利用
 */
function testNewFeature() {
  try {
    // システム診断
    const diagnosis = SystemManager.checkSetupStatus();
    if (!diagnosis.isComplete) {
      throw new Error('システムセットアップが不完全です');
    }

    // 機能テスト実行
    const testUserId = 'test-user-id';
    const result = processFeature(testUserId);
    
    console.info('新機能テスト完了', result);
    return result;
    
  } catch (error) {
    console.error('新機能テストエラー', error.message);
    throw error;
  }
}
```

---

## AI開発での注意点・制約

### 必須チェックリスト
1. **定数使用**: `SYSTEM_CONSTANTS`, `DB_CONFIG`, `CORE` の使用確認
2. **名前空間**: `App`, `DB`, `SecurityValidator` パターンの適用
3. **統一データソース**: `userInfo.spreadsheetId`のみ使用、重複フィールド禁止
4. **セキュリティ**: 入力検証とエラーハンドリングの実装
5. **パフォーマンス**: バッチ処理とキャッシュの活用
6. **ログ**: 構造化ログによるデバッグ情報出力

### 本番デプロイ前チェックリスト
```bash
# 1. 全テスト通過確認
npm run test

# 2. コード品質チェック  
npm run lint

# 3. フォーマット統一
npm run format

# 4. 統合チェック
npm run check

# 5. GASデプロイ
npm run deploy
```

---

# 📊 システム整合性チェック結果（2025年3月最新版）

## 🚀 最新の最適化達成状況
- **データベース構造最適化**: ✅ **14項目→9項目（36%削減）完了**
- **SystemManager.gs統合**: ✅ **分散機能の完全統合達成**
- **ConfigOptimizer.gs削除**: ✅ **重複排除完了**
- **管理画面簡素化**: ✅ **3段階→2段階、appName削除完了**
- **configJson中心設計**: ✅ **全設定統合完了**

## ✅ 完全適合実装状況
- **Object.freeze()使用**: SystemManager.gs(15), constants.gs(27), database.gs(2), main.gs(1) = **45箇所以上**
- **const使用率**: **100%** - 全ファイルでconst/letのみ使用
- **var残存**: **0件** ✅ **完全削除維持**（194件→0件）
- **統一データソース**: ✅ **完全実装**（userInfo.spreadsheetId/sheetName）
- **データベース最適化**: ✅ **36%削減達成**（14項目→9項目）

## 🎯 CLAUDE.md規範 100%達成 + 最適化拡張
1. **データベースシンプル化**: ✅ **36%削減** - 最適化済み9項目構造
2. **機能統合管理**: ✅ **SystemManager.gs** - テスト・最適化・診断の統合
3. **設定統合**: ✅ **configJson中心** - 重複項目の完全統合
4. **UI簡素化**: ✅ **2段階管理画面** - UX向上とappName削除
5. **プライバシー重視**: ✅ **デフォルトOFF** - 心理的安全性向上

## ✅ 完全適合 + 最適化済み項目  
1. **データベーススキーマ**: ✅ **最新9項目**構造で完全一致
2. **SystemManager統合**: ✅ **分散機能の一元管理**達成
3. **configJson統合**: ✅ **全設定の中央管理**実装
4. **動的生成**: ✅ **URL項目の効率化**実装
5. **テスト機能**: ✅ **testSchemaOptimization()** 追加

---

# Claude Code AI開発ワークフロー

> このセクションはClaude Code（AI）を使った効率的なGAS開発プロセスを定義します。

---

## 11) Claude Code を使った反復開発

Claude CodeはAI支援によるコーディング試行錯誤を大幅に効率化できます。

### TDD（テスト駆動開発）アプローチ
1. **テスト先行開発**：新機能実装時は先にテストケースを作成
   ```bash
   npm run test:watch  # テスト監視モード
   ```
2. **AI提案→テスト実行→改善サイクル**：
   - Claudeに実装を提案させる
   - `npm run test` で即座に検証
   - 失敗箇所をClaudeに修正依頼
   - 継続的に品質向上

### コード選択リファクタリング
- VS Code上でコード選択 → Claudeに「この部分をリファクタリングして」
- 局所的な改善で安全なリファクタリング実現
- GAS特有の制約も考慮したリファクタリング提案

### プロジェクトガイドライン参照
- ClaudeはCLAUDE.md、README.mdを自動参照
- プロジェクト固有のルールや意図を自動反映
- チーム開発での一貫性保持

---

## 12) コードフォーマット・Lint自動化

開発効率とコード品質を両立する自動化環境を構築済み。

### 利用可能コマンド
```bash
# コード整形（Prettier）
npm run format

# 構文チェック・自動修正（ESLint）  
npm run lint

# テスト実行
npm run test

# 品質チェック一括実行
npm run check

# デプロイ前総合チェック
npm run deploy
```

### VS Code連携設定推奨
```json
// .vscode/settings.json
{
  "editor.formatOnSave": true,
  "editor.codeActionsOnSave": {
    "source.fixAll.eslint": true
  },
  "eslint.validate": ["javascript", "gas"]
}
```

### AI生成コードの品質保証
- AIが生成したコードも自動的にプロジェクト標準に整形
- ESLintによるGAS特有のベストプラクティスチェック
- 保存時自動修正でヒューマンエラー防止

---

## 13) GAS最適化ベストプラクティス

GAS環境の制約と特性を活かした最適化指針。

### 実行時間・処理分割戦略
- **長時間処理は分割必須**：6分制限対策
- **トリガー・スケジューラ活用**：段階的な大量処理
- **プロセス状態管理**：PropertiesServiceで処理継続

### API呼び出し最適化
```js
// ❌ 非効率：ループ内でAPI呼び出し  
for(const row of rows) {
  sheet.getRange(row, 1).setValue(data);
}

// ✅ 効率的：バッチ処理
const values = rows.map(row => [data]);
sheet.getRange(1, 1, values.length, 1).setValues(values);
```

### キャッシュ戦略
- **CacheService**：短期間（最大6時間）の高速アクセス
- **PropertiesService**：永続化が必要なデータ
- **繰り返し処理のキャッシュ化**：API呼び出し削減

### エラーハンドリング・ログ戦略
```js
const handleWithRetry = (operation, maxRetries = 3) => {
  for(let i = 0; i < maxRetries; i++) {
    try {
      return operation();
    } catch (error) {
      console.error(`Attempt ${i + 1} failed:`, error.message);
      if (i === maxRetries - 1) throw error;
      Utilities.sleep(Math.pow(2, i) * 1000); // 指数バックオフ
    }
  }
};
```

### V8ランタイム活用
- **Map/Set**：高性能データ構造
- **Array.prototype.flat()**：配列展開
- **テンプレートリテラル**：文字列組み立て
- **アロー関数**：簡潔な関数記法

---

## 14) GAS API モックテスト

本格的な単体テストでバグ予防・品質向上を実現。

### テスト実行環境
```bash
# 単体テスト実行
npm run test

# ウォッチモード（開発時）
npm run test:watch

# カバレッジ付きテスト
npm run test -- --coverage
```

### GAS API モック利用例
```js
// tests/example.test.js
describe('スプレッドシート処理', () => {
  beforeEach(() => {
    // モックの初期化
    SpreadsheetApp.getActiveSheet.mockReturnValue({
      getRange: jest.fn(() => ({
        getValues: jest.fn(() => [['data1'], ['data2']]),
        setValues: jest.fn()
      })),
      getLastRow: jest.fn(() => 10)
    });
  });

  test('データ集計処理', () => {
    const result = countData(); // あなたのGAS関数
    expect(result).toEqual(expectedResult);
    // SpreadsheetApp呼び出し検証
    expect(SpreadsheetApp.getActiveSheet).toHaveBeenCalled();
  });
});
```

### Claude連携テスト開発
1. **Claudeにテストケース作成依頼**
   ```
   "この関数のテストケースを作成してください"
   ```
2. **モック設定をClaude提案**
   ```  
   "SpreadsheetAppをモックしてテストしてください"
   ```
3. **カバレッジ向上支援**
   ```
   "テストカバレッジを向上させるケースを追加してください"
   ```

### 利用可能なモック
- **SpreadsheetApp**：シート操作
- **PropertiesService**：設定管理  
- **CacheService**：キャッシュ操作
- **UrlFetchApp**：外部API呼び出し
- **HtmlService**：HTML出力
- **Utilities**：ユーティリティ関数
- **Logger/console**：ログ出力

---

## 15) AI開発での注意点・制約

### Claude提案コードの必須レビュー観点
1. **GAS実行時間制限**：6分以内で完了するか？
2. **API呼び出し効率**：バッチ処理になっているか？  
3. **権限スコープ**：必要最小限の権限か？
4. **エラー処理**：適切な例外ハンドリングがあるか？
5. **ログ出力**：デバッグに必要な情報を出力しているか？

### プロジェクト固有ルール遵守確認
- **命名規則**：既存コードベースとの統一性
- **ファイル構成**：適切な場所への配置  
- **依存関係**：不要なライブラリ追加の回避
- **セキュリティ**：秘匿情報の取り扱い

### 本番デプロイ前チェックリスト
```bash
# 1. 全テスト通過確認
npm run test

# 2. コード品質チェック
npm run lint

# 3. フォーマット統一
npm run format  

# 4. 統合チェック
npm run check

# 5. GASデプロイ
npm run deploy
```

---

## 16) 個人開発向け自動化ワークフロー

個人開発でも効率的で安全な開発を実現する段階的アプローチ。

### 基本開発フロー（推奨）

#### 🚀 新機能開発時の自動化ワークフロー

```bash
# 1. テスト監視モード開始（別ターミナルで常時実行）
npm run test:watch

# 2. Claude Codeでの開発
# - TDDでテストケースを先に作成
# - Claudeに実装を依頼
# - リアルタイムでテスト結果を確認
# - 失敗時はClaudeに修正を依頼

# 3. 完成後の品質チェック＆デプロイ
npm run deploy  # テスト→デプロイの一括実行
```

#### 📝 実際の操作例

```bash
# Terminal 1: テスト監視開始
npm run test:watch

# Terminal 2: 開発作業
# Claude Code: "新しい関数XXXのテストケースを作成してください"
# Claude Code: "テストに合格する実装を作成してください" 
# Claude Code: "エラーが出ています。修正してください"

# 開発完了後
npm run deploy
```

### ブランチ戦略（個人開発向け）

#### 🎯 最小構成：mainブランチ運用（現在）
**メリット**：
- シンプルで迷わない
- オーバーヘッドなし
- 小規模変更に最適

**デメリット**：
- 実験的な変更でmainが壊れるリスク
- ロールバックが困難

#### 🌲 推奨：feature/mainブランチ戦略

```bash
# 新機能開発時
git checkout -b feature/新機能名
npm run test:watch  # 開発開始

# 開発→テスト→マージの自動化
git add .
git commit -m "feat: 新機能の実装

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"

# mainへのマージ（安全確認付き）
git checkout main
git pull  # 念のため最新取得
npm run check  # 品質チェック
git merge feature/新機能名
npm run deploy  # GASデプロイ

# クリーンアップ
git branch -d feature/新機能名
```

#### 🔄 自動化スクリプト案

`.scripts/new-feature.sh` (作成推奨):
```bash
#!/bin/bash
# 新機能開発開始の自動化

if [ -z "$1" ]; then
  echo "使用法: ./scripts/new-feature.sh <機能名>"
  exit 1
fi

FEATURE_NAME="feature/$1"

# ブランチ作成・切り替え
git checkout -b "$FEATURE_NAME"

# テスト監視開始（バックグラウンド）
echo "テスト監視モードを開始します..."
npm run test:watch &
TEST_PID=$!

echo "🚀 新機能 '$1' の開発環境が準備完了！"
echo "📝 Claude Codeでテスト→実装→修正のサイクルを開始してください"
echo "✅ 完了後は ./scripts/merge-feature.sh $1 を実行"
echo "🛑 テスト監視停止: kill $TEST_PID"
```

`.scripts/merge-feature.sh` (作成推奨):
```bash
#!/bin/bash
# 機能完成後のマージ・デプロイ自動化

if [ -z "$1" ]; then
  echo "使用法: ./scripts/merge-feature.sh <機能名>"
  exit 1
fi

FEATURE_NAME="feature/$1"
CURRENT_BRANCH=$(git branch --show-current)

# 現在のブランチ確認
if [ "$CURRENT_BRANCH" != "$FEATURE_NAME" ]; then
  echo "❌ エラー: $FEATURE_NAME ブランチに切り替えてください"
  exit 1
fi

# 最終チェック
echo "🔍 最終品質チェック中..."
npm run check
if [ $? -ne 0 ]; then
  echo "❌ テストが失敗しました。修正してから再実行してください"
  exit 1
fi

# コミット（未コミットがあれば）
if ! git diff-index --quiet HEAD --; then
  echo "📝 変更をコミットしています..."
  git add .
  git commit -m "feat: $1 の実装完了

🤖 Generated with [Claude Code](https://claude.ai/code)

Co-Authored-By: Claude <noreply@anthropic.com>"
fi

# mainにマージ
echo "🔀 mainブランチにマージ中..."
git checkout main
git pull  # リモートの最新を取得
npm run check  # mainでも確認
git merge "$FEATURE_NAME"

# デプロイ
echo "🚀 GASデプロイ中..."
npm run deploy

if [ $? -eq 0 ]; then
  echo "✅ デプロイ成功！"
  echo "🧹 フィーチャーブランチをクリーンアップしますか？ (y/N)"
  read -r response
  if [[ "$response" == "y" ]]; then
    git branch -d "$FEATURE_NAME"
    echo "🗑️ $FEATURE_NAME ブランチを削除しました"
  fi
else
  echo "❌ デプロイに失敗しました"
  exit 1
fi

echo "🎉 機能 '$1' のリリース完了！"
```

### 緊急時・実験時の運用

#### 🔥 ホットフィックス（緊急修正）
```bash
# mainで直接修正（小さな修正のみ）
npm run test:watch  # 監視開始
# Claude Codeで修正
npm run deploy     # 即座にデプロイ
```

#### 🧪 実験的な機能
```bash
git checkout -b experiment/機能名
# 自由に実験
# 成功したらfeatureブランチにリネーム
# 失敗したらブランチ削除
```

### ワークフロー選択指針

| 変更規模 | 推奨ワークフロー | 理由 |
|---------|----------------|------|
| バグ修正（1-2ファイル） | main直接 | オーバーヘッド回避 |
| 新機能追加 | feature/main | 安全性確保 |
| 大規模リファクタ | feature/main | ロールバック可能性 |
| 実験的な変更 | experiment/ | main汚染防止 |

### Claude Code連携のコツ

1. **ブランチ作成を明示**：
   ```
   "feature/ユーザー認証 ブランチで新機能を開発します"
   ```

2. **テスト先行を指示**：
   ```
   "まずテストケースを作成してからユーザー認証機能を実装してください"
   ```

3. **段階的な確認**：
   ```
   "テストが通ったらコミットして、次の機能に進んでください"
   ```

4. **自動化スクリプトの活用**：
   ```
   "./scripts/new-feature.sh ユーザー認証 を実行してください"
   ```

---

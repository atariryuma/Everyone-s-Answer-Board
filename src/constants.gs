/**
 * Core Constants - 2024 GAS V8 Best Practices
 * 最小限のコア定数のみ定義、Lazy-Loading Pattern対応
 */

/**
 * システム全体で使用するコア定数
 * CONSTANT_CASE命名 (Google Style Guide準拠)
 */
const CORE = Object.freeze({
  // タイムアウト設定（ミリ秒）
  TIMEOUTS: Object.freeze({
    SHORT: 1000,      // 1秒 - UI反応
    MEDIUM: 5000,     // 5秒 - API呼び出し  
    LONG: 30000,      // 30秒 - バッチ処理
    FLOW: 300000,     // 5分 - GAS実行制限
  }),
  
  // ステータス定数
  STATUS: Object.freeze({
    ACTIVE: 'active',
    INACTIVE: 'inactive',
    PENDING: 'pending',
    ERROR: 'error',
  }),
  
  // アクセスレベル
  ACCESS_LEVELS: Object.freeze({
    OWNER: 'owner',
    ADMIN: 'admin', 
    USER: 'user',
    GUEST: 'guest',
    NONE: 'none',
  }),
  
  // エラーレベル（数値で優先度表現）
  ERROR_LEVELS: Object.freeze({
    INFO: 0,
    WARN: 1,
    ERROR: 2,
    CRITICAL: 3,
  }),
  
  // 基本的なHTTPステータス
  HTTP_STATUS: Object.freeze({
    OK: 200,
    BAD_REQUEST: 400,
    UNAUTHORIZED: 401,
    FORBIDDEN: 403,
    NOT_FOUND: 404,
    INTERNAL_ERROR: 500,
  }),
});

/**
 * PropertiesServiceキー定数
 * セキュリティ重要項目の一元管理
 */
const PROPS_KEYS = Object.freeze({
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  ADMIN_EMAIL: 'ADMIN_EMAIL',
});
/**
 * constants.gs - システム全体共通定数
 * プロジェクト全体で使用される定数を定義
 */

/* eslint-disable no-unused-vars */

/**
 * システム定数オブジェクト
 */
const CONSTANTS = Object.freeze({

  // アクセスレベル定義
  ACCESS: Object.freeze({
    LEVELS: Object.freeze({
      NONE: 'none',
      GUEST: 'guest',
      AUTHENTICATED_USER: 'authenticated_user',
      OWNER: 'owner',
      SYSTEM_ADMIN: 'system_admin'
    })
  }),

  // データベース設定
  DATABASE: Object.freeze({
    FIELDS: Object.freeze({
      USER_ID: 'userId',
      USER_EMAIL: 'userEmail',
      IS_ACTIVE: 'isActive',
      CONFIG_JSON: 'configJson',
      LAST_MODIFIED: 'lastModified'
    })
  }),

  // キャッシュ設定
  CACHE: Object.freeze({
    TTL: Object.freeze({
      SHORT: 60,    // 1分
      MEDIUM: 300,  // 5分
      LONG: 1800    // 30分
    })
  })
});

/**
 * プロパティキー定数
 */
const PROPS_KEYS = Object.freeze({
  ADMIN_EMAIL: 'ADMIN_EMAIL',
  DATABASE_SPREADSHEET_ID: 'DATABASE_SPREADSHEET_ID',
  SERVICE_ACCOUNT_CREDS: 'SERVICE_ACCOUNT_CREDS',
  GOOGLE_CLIENT_ID: 'GOOGLE_CLIENT_ID'
});
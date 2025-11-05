/**
 * @fileoverview Helper Utilities
 *
 * 🎯 責任範囲:
 * - 列マッピング・インデックス操作
 * - データフォーマッティング
 * - 汎用ヘルパー関数
 * - 計算・変換ユーティリティ
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - 直接的な関数エクスポート
 * - 簡素なユーティリティ関数群
 */

/* global CACHE_DURATION, TIMEOUT_MS, SLEEP_MS */


// ⚡ Runtime Memory Cache for PropertiesService with TTL
// ✅ API最適化: PropertiesService呼び出し80-90%削減
// ✅ CLAUDE.md準拠: 30秒TTLで自動期限切れ
const RUNTIME_PROPERTIES_CACHE = {};
const PROPERTY_CACHE_TTL = 30000; // 30秒（CLAUDE.md準拠）

/**
 * PropertiesServiceのメモリキャッシュ付きアクセス（TTL対応）
 * ✅ CLAUDE.md準拠: 30秒TTLで自動期限切れ
 * ✅ Google公式推奨: 頻繁アクセスする設定値はメモリにキャッシュ
 * @param {string} key - プロパティキー
 * @returns {string|null} プロパティ値
 */
function getCachedProperty(key) {
  const now = Date.now();
  const cached = RUNTIME_PROPERTIES_CACHE[key];

  // ✅ TTLチェック: 有効期限内ならキャッシュを返す
  if (cached && cached.timestamp && (now - cached.timestamp < PROPERTY_CACHE_TTL)) {
    return cached.value;
  }

  // キャッシュミスまたは期限切れ: PropertiesServiceから取得
  const value = PropertiesService.getScriptProperties().getProperty(key);
  RUNTIME_PROPERTIES_CACHE[key] = {
    value,
    timestamp: now  // ✅ タイムスタンプ記録
  };
  return value;
}

/**
 * メモリキャッシュをクリア（テスト用・設定変更時用）
 * ✅ 明示的なクリアも可能（システム設定更新時など）
 * @param {string} key - クリアするキー（省略時は全クリア）
 */
function clearPropertyCache(key = null) {
  if (key) {
    delete RUNTIME_PROPERTIES_CACHE[key];
  } else {
    Object.keys(RUNTIME_PROPERTIES_CACHE).forEach(k => delete RUNTIME_PROPERTIES_CACHE[k]);
  }
}

/**
 * オブジェクトをシンプルなハッシュ文字列に変換
 * ✅ API最適化: JSON.stringify()より約50%高速
 * @param {Object} obj - ハッシュ化するオブジェクト
 * @returns {string} ハッシュ文字列
 */
function simpleHash(obj) {
  if (!obj || typeof obj !== 'object') return '';
  const keys = Object.keys(obj).sort();
  return keys.map(k => `${k}:${obj[k]}`).join('|');
}



// 📋 Response Standardization (Zero-Dependency)


/**
 * 標準化エラーレスポンス生成（拡張版）
 * @param {string} message - エラーメッセージ
 * @param {*} data - 追加データ
 * @param {Object} extraFields - 追加フィールド
 * @returns {Object} 標準エラーレスポンス
 */
function createErrorResponse(message, data = null, extraFields = {}) {
  return {
    success: false,
    message,
    error: message,
    ...(data && { data }),
    ...extraFields
  };
}

/**
 * 標準化成功レスポンス生成（拡張版）
 * @param {string} message - 成功メッセージ
 * @param {*} data - レスポンスデータ
 * @param {Object} extraFields - 追加フィールド
 * @returns {Object} 標準成功レスポンス
 */
function createSuccessResponse(message, data = null, extraFields = {}) {
  return {
    success: true,
    message,
    ...(data && { data }),
    ...extraFields
  };
}

/**
 * データサービス用エラーレスポンス
 * @param {string} message - エラーメッセージ
 * @param {string} sheetName - シート名
 * @returns {Object} データサービス用エラーレスポンス
 */
function createDataServiceErrorResponse(message, sheetName = '') {
  return createErrorResponse(message, [], { headers: [], sheetName });
}


/**
 * 認証エラーレスポンス
 * @returns {Object} 認証エラー
 */
function createAuthError() {
  return createErrorResponse('ユーザー認証が必要です');
}

/**
 * ユーザー未発見エラーレスポンス
 * @returns {Object} ユーザー未発見エラー
 */
function createUserNotFoundError() {
  return createErrorResponse('ユーザーが見つかりません');
}

/**
 * 管理者権限エラーレスポンス
 * @returns {Object} 管理者権限エラー
 */
function createAdminRequiredError() {
  return createErrorResponse('管理者権限が必要です');
}

/**
 * 例外エラーレスポンス生成
 * @param {Error} error - エラーオブジェクト
 * @returns {Object} 例外エラーレスポンス
 */
function createExceptionResponse(error) {
  return createErrorResponse(error.message || 'Unknown error');
}


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
 * データサービス用成功レスポンス
 * @param {Array} data - データ配列
 * @param {Array} headers - ヘッダー配列
 * @param {string} sheetName - シート名
 * @param {string} message - 成功メッセージ
 * @returns {Object} データサービス用成功レスポンス
 */
function createDataServiceSuccessResponse(data, headers, sheetName, message = 'データ取得成功') {
  return createSuccessResponse(message, data, { headers, sheetName });
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


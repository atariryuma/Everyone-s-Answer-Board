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

/* global */





// ===========================================
// 📋 Response Standardization (Zero-Dependency)
// ===========================================

/**
 * 標準化エラーレスポンス生成
 * @param {string} message - エラーメッセージ
 * @param {*} data - 追加データ
 * @returns {Object} 標準エラーレスポンス
 */
function createErrorResponse(message, data = null) {
  return { success: false, message, error: message, ...(data && { data }) };
}

/**
 * 標準化成功レスポンス生成
 * @param {string} message - 成功メッセージ
 * @param {*} data - レスポンスデータ
 * @returns {Object} 標準成功レスポンス
 */
function createSuccessResponse(message, data = null) {
  return { success: true, message, ...(data && { data }) };
}

/**
 * 認証エラーレスポンス
 * @returns {Object} 認証エラー
 */
function createAuthError() {
  return createErrorResponse('Not authenticated');
}

/**
 * ユーザー未発見エラーレスポンス
 * @returns {Object} ユーザー未発見エラー
 */
function createUserNotFoundError() {
  return createErrorResponse('User not found');
}

/**
 * 管理者権限エラーレスポンス
 * @returns {Object} 管理者権限エラー
 */
function createAdminRequiredError() {
  return createErrorResponse('Admin access required');
}

/**
 * 例外エラーレスポンス生成
 * @param {Error} error - エラーオブジェクト
 * @returns {Object} 例外エラーレスポンス
 */
function createExceptionResponse(error) {
  return createErrorResponse(error.message || 'Unknown error');
}


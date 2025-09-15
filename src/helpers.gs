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
// 📋 列操作ヘルパー関数群
// ===========================================

/**
 * 統一列インデックス取得
 * @param {Object} config - ユーザー設定
 * @param {string} columnType - 列タイプ (answer/reason/class/name)
 * @returns {number} 列インデックス（見つからない場合は-1）
 */
function getHelperColumnIndex(config, columnType) {
    const index = config?.columnMapping?.mapping?.[columnType];
    return typeof index === 'number' ? index : -1;
}




// ===========================================
// 📏 フォーマッティングヘルパー関数群
// ===========================================

// formatTimestamp - formatters.jsに統一 (重複削除完了)





// ===========================================
// 🔍 検証ヘルパー関数群
// ===========================================




// ===========================================
// 🧮 計算ヘルパー関数群
// ===========================================





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
  return { success: false, message, ...(data && { data }) };
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

// ===========================================
// 🔄 レガシー互換関数 (GAS Best Practices準拠)
// ===========================================


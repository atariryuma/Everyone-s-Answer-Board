/**
 * @fileoverview Data Formatting & Transformation
 *
 * 🎯 責任範囲:
 * - データの表示用フォーマット
 * - 型変換・データ変換
 * - レスポンス形式の統一
 * - 出力データの正規化
 *
 * 🔄 GAS Best Practices準拠:
 * - フラット関数構造 (Object.freeze削除)
 * - 直接的な関数エクスポート
 * - 簡素なユーティリティ関数群
 */

/* global */

// ===========================================
// 📅 データ表示用フォーマット関数群
// ===========================================

/**
 * タイムスタンプをフォーマット
 * @param {string|Date} timestamp - タイムスタンプ
 * @returns {string} フォーマット済みタイムスタンプ
 */
function formatTimestamp(timestamp) {
  try {
    if (!timestamp) return '-';

    const date = new Date(timestamp);
    if (isNaN(date.getTime())) return '-';

    // YYYY/MM/DD HH:MM形式
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  } catch (error) {
    console.warn('formatTimestamp error:', error.message);
    return '-';
  }
}







// ===========================================
// 🌐 HTML出力用フォーマット関数群
// ===========================================





// ===========================================
// ⚙️ 設定表示用フォーマット関数群
// ===========================================





// ===========================================
// 🔄 レガシー互換関数 (GAS Best Practices準拠)
// ===========================================





/**
 * フォーマッター診断
 * @returns {Object} 診断結果
 */

// ===========================================
// 🎯 GAS Best Practice: Simple Functions
// ===========================================

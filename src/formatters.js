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

    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');
  } catch (error) {
    console.warn('formatTimestamp error:', error.message);
    return '-';
  }
}


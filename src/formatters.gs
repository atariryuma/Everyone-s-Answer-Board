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


// 📅 データ表示用フォーマット関数群


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


// 🌐 HTML出力用フォーマット関数群


/**
 * HTMLエスケープ処理
 * @param {string} text - エスケープするテキスト
 * @returns {string} HTMLエスケープ済みテキスト
 */
function htmlEncode(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
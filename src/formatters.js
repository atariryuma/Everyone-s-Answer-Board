/**
 * @fileoverview Data formatting helpers (timestamp formatting only — currently).
 */

/* global */

/**
 * タイムスタンプを `yyyy/MM/dd HH:mm` で整形（不正値は '-'）
 * @param {string|Date} timestamp
 * @returns {string}
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


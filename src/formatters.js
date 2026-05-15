/**
 * @fileoverview Data formatting helpers (timestamp formatting only — currently).
 */

// タイムスタンプを `yyyy/MM/dd HH:mm` で整形（不正値は '-'）。
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


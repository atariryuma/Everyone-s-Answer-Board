/**
 * @fileoverview 共通ユーティリティ関数群
 * 使用されている必要最低限のユーティリティを提供
 */

/**
 * 現在のアクティブユーザーのメールアドレスを安全に取得
 * @returns {string} ユーザーメールアドレス（取得失敗時は空文字）
 */
function getCurrentUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    debugLog('getCurrentUserEmail: Retrieved email:', email); // 追加
    return email || '';
  } catch (error) {
    logError(error, 'getCurrentUserEmail', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.AUTHENTICATION);
    debugLog('getCurrentUserEmail: Error retrieving email, returning empty string. Error:', error.message); // 追加
    return '';
  }
}

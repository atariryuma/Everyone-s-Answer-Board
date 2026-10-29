/**
 * @fileoverview UserService - 認証・現在ユーザー情報取得（isAdmin /
 *   IDトークン由来の Google 不変 ID 取得）。
 */

/* global getCurrentEmail, isAdministrator */

/**
 * ユーザー情報キャッシュキーを生成（UserApis.js から参照される）
 */
function getUserInfoCacheKey(email) {
  return `current_user_info_${email}`;
}

/**
 * IDトークンからGoogleアカウントの不変ID（sub）を取得
 * メールアドレスが変更されても sub は不変のため、アカウント同一性の判定に使用
 * @returns {Object|null} { googleId: string, email: string } or null
 */
function getGoogleIdFromToken() {
  try {
    const token = ScriptApp.getIdentityToken();
    if (!token) {
      return null;
    }

    const parts = token.split('.');
    if (parts.length !== 3) {
      console.warn('getGoogleIdFromToken: Invalid JWT format');
      return null;
    }

    const payload = JSON.parse(
      Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString()
    );

    if (!payload.sub) {
      return null;
    }

    // プレフィックス付きで返す（スプレッドシートの数値自動変換による精度損失を防止）
    return { googleId: 'gid_' + payload.sub, email: payload.email || null };
  } catch (error) {
    logError_('getGoogleIdFromToken', error);
    return null;
  }
}

/**
 * 管理者権限確認（フロントエンド互換性）
 * 統一認証システムのAdministrator権限をチェック
 * @returns {boolean} 管理者権限があるか
 */
function isAdmin() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return false;
    }
    return isAdministrator(email);
  } catch (error) {
    logError_('isAdmin', error);
    return false;
  }
}

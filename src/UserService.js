/**
 * @fileoverview UserService - 統一ユーザー管理サービス (遅延初期化対応)
 *
 * 責任範囲:
 * - ユーザー認証・セッション管理
 * - ユーザー情報の取得・更新
 * - 権限・アクセス制御
 * - ユーザーキャッシュ管理
 *
 * GAS Best Practices準拠:
 * - 遅延初期化パターン (各公開関数先頭でinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global validateUrl, validateEmail, getCurrentEmail, findUserByEmail, findUserById, findUserByGoogleId, openSpreadsheet, updateUser, getUserConfig, getConfigOrDefault, isAdministrator, CACHE_DURATION, clearConfigCache, SYSTEM_LIMITS, createExceptionResponse, ScriptApp, Utilities */

/**
 * UserService - ゼロ依存アーキテクチャ
 * GAS-Nativeパターンによる直接APIアクセス
 * DB, CONSTANTS, PROPS_KEYS依存を排除
 */

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
    console.error('getGoogleIdFromToken error:', error.message);
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
    console.error('isAdmin error:', error.message);
    return false;
  }
}

/**
 * Get user information with specified type
 * @param {string} infoType - Type of info: 'email', 'full'
 * @returns {Object} User information
 */
function getUser(infoType = 'email') {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return {
        success: false,
        message: 'Not authenticated'
      };
    }

    if (infoType === 'email') {
      return {
        success: true,
        email
      };
    }

    if (infoType === 'full') {
      const user = findUserByEmail(email, { requestingUser: email });
      return {
        success: true,
        email,
        user: user || null
      };
    }

    return {
      success: false,
      message: `Unsupported infoType: ${infoType}`
    };
  } catch (error) {
    console.error('getUser error:', error.message);
    return {
      success: false,
      message: error.message || 'User retrieval failed'
    };
  }
}

/**
 * メールアドレス検証（SecurityServiceに委譲）
 * @param {string} email - メールアドレス
 * @returns {Object} 検証結果
 */

/**
 * フォームURL検証
 * @param {string} formUrl - フォームURL
 * @returns {boolean} 有効かどうか
 */

/**
 * サービス状態診断
 * @returns {Object} 診断結果
 */

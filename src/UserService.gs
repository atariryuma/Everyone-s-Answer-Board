/**
 * @fileoverview UserService - 統一ユーザー管理サービス (遅延初期化対応)
 *
 * 🎯 責任範囲:
 * - ユーザー認証・セッション管理
 * - ユーザー情報の取得・更新
 * - 権限・アクセス制御
 * - ユーザーキャッシュ管理
 *
 * 🔄 GAS Best Practices準拠:
 * - 遅延初期化パターン (各公開関数先頭でinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global validateUrl, validateEmail, getCurrentEmail, findUserByEmail, findUserById, openSpreadsheet, updateUser, getUserSpreadsheetData, getUserConfig, isAdministrator, CACHE_DURATION */

// ===========================================
// 🔧 GAS-Native UserService (直接API版)
// ===========================================

/**
 * UserService - ゼロ依存アーキテクチャ
 * GAS-Nativeパターンによる直接APIアクセス
 * DB, CONSTANTS, PROPS_KEYS依存を排除
 */



/**
 * 現在のユーザー情報取得（キャッシュ対応）
 * @returns {Object|null} ユーザー情報オブジェクト
 */
function getCurrentUserInfo() {
  const cacheKey = 'current_user_info';

  try {
    // キャッシュ確認
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    // セッション取得
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      console.warn('getCurrentUserInfo: 有効なセッションなし');
      return null;
    }

    // データベース検索
    const userInfo = findUserByEmail(email);
    if (!userInfo) {
      console.info('UserService.getCurrentUserInfo: 新規ユーザーの可能性', { email });
      return null;
    }

    // 設定情報を統合
    const completeUserInfo = enrichUserInfo(userInfo);

    // キャッシュ保存（5分間）
    cache.put(cacheKey, JSON.stringify(completeUserInfo), CACHE_DURATION.LONG);

    return completeUserInfo;
  } catch (error) {
    console.error('UserService.getCurrentUserInfo: エラー', {
      error: error.message,
      stack: error.stack
    });
    return null;
  }
}

/**
 * ユーザー情報を設定で拡張
 * @param {Object} userInfo - 基本ユーザー情報
 * @returns {Object} 拡張されたユーザー情報
 */
function enrichUserInfo(userInfo) {
    try {
      if (!userInfo || !userInfo.userId) {
        throw new Error('無効なユーザー情報');
      }

      // 統一API使用: 構造化パース
      const configResult = getUserConfig(userInfo.userId);
      const config = configResult.success ? configResult.config : {};

      // 動的URLを生成・キャッシュ
      const enrichedConfig = generateDynamicUserUrls(config);

      return {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive,
        lastModified: userInfo.lastModified,
        config: enrichedConfig,
        userInfo: {
          userId: userInfo.userId,
          userEmail: userInfo.userEmail,
          isActive: userInfo.isActive
        }
      };
    } catch (error) {
      console.error('UserService.enrichUserInfo: エラー', error.message);
      return userInfo; // フォールバック
    }
}

/**
 * 動的URL生成（spreadsheetUrl, appUrl等）
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} URL付き設定オブジェクト
 */
function generateDynamicUserUrls(config) {
    try {
      const enhanced = { ...config };

      // スプレッドシートURL生成
      if (config.spreadsheetId && !config.spreadsheetUrl) {
        if (config.spreadsheetId && typeof config.spreadsheetId === 'string') {
          enhanced.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}/edit`;
        }
      }

      // アプリURL生成（公開済みの場合）
      if (config.isPublished && !config.appUrl) {
        enhanced.appUrl = ScriptApp.getService().getUrl();
      }

      // フォームURL存在確認
      if (config.formUrl) {
        enhanced.hasValidForm = validateUrl(config.formUrl)?.isValid || false;
      }

      return enhanced;
    } catch (error) {
      console.error('UserService.generateDynamicUrls: エラー', error.message);
      return config; // フォールバック
    }
}

// ===========================================
// 🛡️ 権限・アクセス制御
// ===========================================


/**
 * 編集者権限確認（旧: 所有者権限確認）
 * @param {string} userId - 確認対象ユーザーID
 * @returns {boolean} 編集者かどうか
 */
function checkUserEditorAccess(userId) {
  // userIdからemailを取得してgetUnifiedAccessLevelを使用
  const user = findUserById(userId);
  if (!user) return false;

  const accessLevel = getUnifiedAccessLevel(user.userEmail, userId);
  // 🔧 用語統一: owner → editor
  return accessLevel === 'editor';
}


// ===========================================
// 🔐 統一認証システム（Administrator/Editor/Viewer）
// ===========================================


/**
 * 編集者権限判定（Editor）
 * @param {string} email - メールアドレス
 * @param {string} targetUserId - 対象ユーザーID
 * @returns {boolean} 編集権限があるか
 */
function isEditor(email, targetUserId) {
  if (!email || !targetUserId) {
    return false;
  }

  try {
    const user = findUserByEmail(email);
    return user && user.userId === targetUserId;
  } catch (error) {
    console.error('UserService.isEditor: エラー', error.message);
    return false;
  }
}

/**
 * 統一アクセスレベル取得（Email-based）
 * @param {string} email - メールアドレス
 * @param {string} targetUserId - 対象ユーザーID（オプショナル）
 * @returns {string} アクセスレベル
 */
function getUnifiedAccessLevel(email, targetUserId) {
  if (isAdministrator(email)) return 'administrator';
  if (targetUserId && isEditor(email, targetUserId)) return 'editor';
  return email ? 'authenticated_user' : 'guest';
}

// ===========================================
// 🧹 キャッシュ・セッション管理
// ===========================================

/**
 * ユーザーキャッシュクリア
 * @param {string} userId - ユーザーID（オプション、未指定時は全体）
 */


// ===========================================
// 🔧 ユーティリティ
// ===========================================


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

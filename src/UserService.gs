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

/* global validateUrl, validateEmail, getCurrentEmail, findUserByEmail, findUserById, openSpreadsheet, updateUser, getUserConfig, isAdministrator, CACHE_DURATION, clearConfigCache, SYSTEM_LIMITS, createExceptionResponse */


// GAS-Native UserService (直接API版)


/**
 * UserService - ゼロ依存アーキテクチャ
 * GAS-Nativeパターンによる直接APIアクセス
 * DB, CONSTANTS, PROPS_KEYS依存を排除
 */



/**
 * 現在のユーザー情報取得（キャッシュ対応）
 * ✅ SECURITY: ユーザー固有キャッシュキーで個人情報隔離
 * @returns {Object|null} ユーザー情報オブジェクト
 */
function getCurrentUserInfo() {
  try {
    // セッション取得
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      console.warn('getCurrentUserInfo: 有効なセッションなし');
      return null;
    }

    // ✅ SECURITY FIX: ユーザー固有のキャッシュキー（共有キャッシュの個人情報流出防止）
    const cacheKey = `current_user_info_${email}`;

    // キャッシュ確認
    const cache = CacheService.getScriptCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    // データベース検索
    const userInfo = findUserByEmail(email, { requestingUser: email });
    if (!userInfo) {
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


// 🛡️ 権限・アクセス制御



/**
 * 編集者権限確認（旧: 所有者権限確認）
 * @param {string} userId - 確認対象ユーザーID
 * @returns {boolean} 編集者かどうか
 */
function checkUserEditorAccess(userId) {
  // userIdからemailを取得してgetUnifiedAccessLevelを使用
  const user = findUserById(userId, { requestingUser: getCurrentEmail() });
  if (!user) return false;

  const accessLevel = getUnifiedAccessLevel(user.userEmail, userId);
  // 🔧 用語統一: owner → editor
  return accessLevel === 'editor';
}



// 🔐 統一認証システム（Administrator/Editor/Viewer）



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
    const user = findUserByEmail(email, { requestingUser: email });
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


// 🧹 キャッシュ・セッション管理


/**
 * ユーザーキャッシュクリア
 * @param {string} userId - ユーザーID（オプション、未指定時は全体）
 */



// 🔧 ユーティリティ



// 🔧 User Management Functions (from main.gs)


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
      // Get user from database if available
      // 🔧 GAS-Native統一: 直接findUserByEmail使用
      const user = findUserByEmail(email, { requestingUser: email });
      return {
        success: true,
        email,
        user: user || null
      };
    }

    // Default return for unsupported infoType
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
 * Reset authentication and clear all user session data
 * ✅ CLAUDE.md準拠: 包括的キャッシュクリア with 論理的破綻修正
 */
function resetAuth() {
  try {
    const cache = CacheService.getScriptCache();
    let clearedKeysCount = 0;
    let clearConfigResult = null;

    // 🔧 修正1: 現在ユーザー情報を事前取得（キャッシュクリア前）
    const currentEmail = getCurrentEmail();
    const currentUser = currentEmail ? findUserByEmail(currentEmail, { requestingUser: currentEmail }) : null;
    const userId = currentUser?.userId;

    // 🔧 修正2: ConfigService専用クリア関数の活用
    if (userId) {
      try {
        clearConfigCache(userId);
        clearConfigResult = 'ConfigService cache cleared successfully';
      } catch (configError) {
        console.warn('resetAuth: ConfigService cache clear failed:', configError.message);
        clearConfigResult = `ConfigService cache clear failed: ${configError.message}`;
      }
    }

    // 🔧 修正3: 包括的キャッシュキーリスト（実際の使用パターンに合わせて更新）
    // ✅ SECURITY NOTE: current_user_info はユーザー固有キー（current_user_info_${email}）に変更済み
    // グローバルキャッシュキーのみをリストアップ
    const globalCacheKeysToRemove = [
      'admin_auth_cache',
      'session_data',
      'system_diagnostic_cache',
      'bulk_admin_data_cache'
    ];

    // グローバルキャッシュクリア
    globalCacheKeysToRemove.forEach(key => {
      try {
        cache.remove(key);
        clearedKeysCount++;
      } catch (e) {
        console.warn(`resetAuth: Failed to remove global cache key ${key}:`, e.message);
      }
    });

    // 🔧 修正4: User固有キャッシュの完全クリア（email + userId ベース）
    const userSpecificKeysCleared = [];
    if (currentEmail) {
      const emailBasedKeys = [
        `current_user_info_${currentEmail}`,  // ✅ SECURITY FIX: ユーザー固有キー追加
        `board_data_${currentEmail}`,
        `user_data_${currentEmail}`,
        `admin_panel_${currentEmail}`
      ];

      emailBasedKeys.forEach(key => {
        try {
          cache.remove(key);
          userSpecificKeysCleared.push(key);
          clearedKeysCount++;
        } catch (e) {
          console.warn(`resetAuth: Failed to remove email-based cache key ${key}:`, e.message);
        }
      });
    }

    if (userId) {
      const userIdBasedKeys = [
        `user_config_${userId}`,
        `config_${userId}`,
        `user_${userId}`,
        `board_data_${userId}`,
        `question_text_${userId}`
      ];

      userIdBasedKeys.forEach(key => {
        try {
          cache.remove(key);
          userSpecificKeysCleared.push(key);
          clearedKeysCount++;
        } catch (e) {
          console.warn(`resetAuth: Failed to remove userId-based cache key ${key}:`, e.message);
        }
      });
    }

    // 🔧 修正5: リアクション・ハイライトロック完全クリア
    let reactionLocksCleared = 0;
    if (userId) {
      try {
        // リアクション・ハイライトロックの動的検索・削除は複雑なので、
        // 基本的なロックキーのパターンをクリア
        const lockPatterns = [
          `reaction_${userId}_`,
          `highlight_${userId}_`
        ];

        // Note: GAS CacheService doesn't support pattern matching,
        // so we clear known common patterns and rely on TTL expiration
        for (let i = 0; i < SYSTEM_LIMITS.MAX_LOCK_ROWS; i++) { // 最大100行のロックをクリア
          lockPatterns.forEach(pattern => {
            try {
              cache.remove(`${pattern}${i}`);
              reactionLocksCleared++;
            } catch (e) {
              // Lock key might not exist - this is expected
            }
          });
        }
      } catch (lockError) {
        console.warn('resetAuth: Reaction lock clearing failed:', lockError.message);
      }
    }

    // 🔧 修正6: 包括的なログ出力
    const logDetails = {
      currentUser: currentEmail ? `${currentEmail.substring(0, 8)}***@${currentEmail.split('@')[1]}` : 'N/A',
      userId: userId ? `${userId.substring(0, 8)}***` : 'N/A',
      globalKeysCleared: globalCacheKeysToRemove.length,
      userSpecificKeysCleared: userSpecificKeysCleared.length,
      reactionLocksCleared,
      configServiceResult: clearConfigResult,
      totalKeysCleared: clearedKeysCount
    };


    return {
      success: true,
      message: 'Authentication and session data cleared successfully',
      details: {
        clearedKeys: clearedKeysCount,
        userSpecificKeys: userSpecificKeysCleared.length,
        reactionLocks: reactionLocksCleared,
        configService: clearConfigResult ? 'success' : 'skipped'
      }
    };
  } catch (error) {
    console.error('resetAuth error:', error.message, error.stack);
    return createExceptionResponse(error);
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

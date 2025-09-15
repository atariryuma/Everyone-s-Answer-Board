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

/* global ServiceFactory, DatabaseOperations, validateUrl, validateEmail, getCurrentEmail */

// ===========================================
// 🔧 Zero-Dependency UserService (ServiceFactory版)
// ===========================================

/**
 * UserService - ゼロ依存アーキテクチャ
 * ServiceFactoryパターンによる依存関係完全除去
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
    const cache = ServiceFactory.getCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return JSON.parse(cached);
    }

    // セッション取得
    const session = ServiceFactory.getSession();
    const {email} = session;
    if (!email) {
      console.warn('getCurrentUserInfo: 有効なセッションなし');
      return null;
    }

    // データベース検索
    const userInfo = DatabaseOperations.findUserByEmail(email);
    if (!userInfo) {
      console.info('UserService.getCurrentUserInfo: 新規ユーザーの可能性', { email });
      return null;
    }

    // 設定情報を統合
    const completeUserInfo = enrichUserInfo(userInfo);

    // キャッシュ保存（5分間）
    cache.put(cacheKey, JSON.stringify(completeUserInfo), 300);

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

      // configJsonを解析
      let config = {};
      try {
        config = JSON.parse(userInfo.configJson || '{}');
      } catch (parseError) {
        console.warn('UserService.enrichUserInfo: configJson解析エラー', parseError.message);
        config = {};
      }

      // 動的URLを生成・キャッシュ
      const enrichedConfig = generateDynamicUserUrls(config);

      return {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive,
        lastModified: userInfo.lastModified,
        config: enrichedConfig,
        // レガシー互換性のためのフィールド
        currentUserEmail: userInfo.userEmail,
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
        enhanced.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}/edit`;
      }

      // アプリURL生成（公開済みの場合）
      if (config.appPublished && !config.appUrl) {
        enhanced.appUrl = ServiceFactory.getUtils().getWebAppUrl();
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
 * ユーザーアクセスレベル取得
 * @param {string} userId - ユーザーID
 * @returns {string} アクセスレベル (owner/system_admin/authenticated_user/guest/none)
 */
function getUserAccessLevel(userId) {
  try {
    const ACCESS_LEVELS = {
      NONE: 'none',
      GUEST: 'guest',
      AUTHENTICATED_USER: 'authenticated',
      OWNER: 'owner',
      SYSTEM_ADMIN: 'system_admin'
    };

    if (!userId) {
      return ACCESS_LEVELS.GUEST;
    }

    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      return ACCESS_LEVELS.NONE;
    }

    const userInfo = DatabaseOperations.findUserById(userId);
    if (!userInfo) {
      return ACCESS_LEVELS.NONE;
    }

    // 所有者チェック
    if (userInfo.userEmail === currentEmail) {
      return ACCESS_LEVELS.OWNER;
    }

    // システム管理者チェック
    if (isSystemAdmin(currentEmail)) {
      return ACCESS_LEVELS.SYSTEM_ADMIN;
    }

    // 認証済みユーザー
    return ACCESS_LEVELS.AUTHENTICATED_USER;
  } catch (error) {
    console.error('UserService.getAccessLevel: エラー', error.message);
    return 'none';
  }
}

/**
 * 所有者権限確認
 * @param {string} userId - 確認対象ユーザーID
 * @returns {boolean} 所有者かどうか
 */
function verifyUserOwnership(userId) {
  const accessLevel = getUserAccessLevel(userId);
  // 🔧 CONSTANTS依存除去: 直接値比較
  return accessLevel === 'owner';
}

/**
 * システム管理者確認
 * @param {string} email - メールアドレス
 * @returns {boolean} システム管理者かどうか
 */
function isSystemAdmin(email) {
  try {
    if (!email) {
      return false;
    }

    const adminEmail = ServiceFactory.getProperties().getProperty('ADMIN_EMAIL');

    if (!adminEmail) {
      console.warn('UserService.isSystemAdmin: ADMIN_EMAILが設定されていません');
      return false;
    }

    // 管理者メールと一致チェック
    const isAdmin = email.toLowerCase() === adminEmail.toLowerCase();

    if (isAdmin) {
      console.info('UserService.isSystemAdmin: システム管理者を認証', { email });
    }

    return isAdmin;
  } catch (error) {
    console.error('UserService.isSystemAdmin: エラー', {
      email,
      error: error.message
    });
    return false;
  }
}

// ===========================================
// 🔄 ユーザー操作
// ===========================================

/**
 * 新規ユーザー作成
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {Object} initialConfig - 初期設定（オプション）
 * @returns {Object} 作成されたユーザー情報
 */
function createUser(userEmail, initialConfig = {}) {
  try {
    if (!userEmail || !validateEmail(userEmail).isValid) {
      throw new Error('無効なメールアドレス');
    }

    // 既存ユーザーチェック
    const existingUser = DatabaseOperations.findUserByEmail(userEmail);
    if (existingUser) {
      console.info('UserService.createUser: 既存ユーザーを返却', { userEmail });
      return existingUser;
    }

    // 新規ユーザーデータ作成
    const userData = buildNewUserData(userEmail, initialConfig);

    // データベースに保存
    const success = DatabaseOperations.createUser(userData);
    if (!success) {
      throw new Error('ユーザー作成に失敗');
    }

    console.info('UserService.createUser: 新規ユーザー作成完了', {
      userEmail,
      userId: userData.userId
    });

    return userData;
  } catch (error) {
    console.error('UserService.createUser: エラー', {
      userEmail,
      error: error.message
    });
    throw error;
  }
}

/**
 * 新規ユーザーデータ構築
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {Object} initialConfig - 初期設定
 * @returns {Object} 新規ユーザーデータ
 */
function buildNewUserData(userEmail, initialConfig) {
    // 統一ID生成関数を使用（main.gsのgenerateUserIdと同一）
    const userId = Utilities.getUuid();
    const timestamp = new Date().toISOString();

    // CLAUDE.md準拠: 最小限configJSON
    const minimalConfig = {
      setupStatus: 'pending',
      appPublished: false,
      displaySettings: {
        showNames: false, // 心理的安全性重視
        showReactions: false
      },
      createdAt: timestamp,
      lastModified: timestamp,
      ...initialConfig // カスタム初期設定をマージ
    };

    return {
      userId,
      userEmail,
      isActive: true,
      configJson: JSON.stringify(minimalConfig),
      lastModified: timestamp
    };
}

// ===========================================
// 🧹 キャッシュ・セッション管理
// ===========================================

/**
 * ユーザーキャッシュクリア
 * @param {string} userId - ユーザーID（オプション、未指定時は全体）
 */
function clearUserCache(userId = null) {
    // CacheServiceに統一委譲
    const cache = ServiceFactory.getCache();
    const cacheKey = userId ? `user_info_${userId}` : 'current_user_info';
    return cache.remove(cacheKey);
}

/**
 * セッション状態確認
 * @returns {Object} セッション状態情報
 */
function getUserSessionStatus() {
    try {
      const email = getCurrentEmail();
      const userInfo = email ? getCurrentUserInfo() : null;

      return {
        isAuthenticated: !!email,
        userEmail: email,
        hasUserInfo: !!userInfo,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('UserService.getSessionStatus: エラー', error.message);
      return {
        isAuthenticated: false,
        userEmail: null,
        hasUserInfo: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
}

// ===========================================
// 🔧 ユーティリティ
// ===========================================

// findUserByEmail is provided by DatabaseOperations (duplicate removed)

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
// validateUserFormUrl function removed - use validateUrl from validators.gs instead

/**
 * サービス状態診断
 * @returns {Object} 診断結果
 */

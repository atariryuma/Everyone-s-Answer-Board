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

/* global ServiceFactory */

// ===========================================
// 🔧 Zero-Dependency UserService (ServiceFactory版)
// ===========================================

/**
 * UserService - ゼロ依存アーキテクチャ
 * ServiceFactoryパターンによる依存関係完全除去
 * DB, CONSTANTS, PROPS_KEYS依存を排除
 */

/**
 * ServiceFactory統合初期化（UserService版）
 * 依存関係チェックなしの即座初期化
 * @returns {boolean} 初期化成功可否
 */
function initUserServiceZero() {
  try {
    // ServiceFactory利用可能性確認
    if (typeof ServiceFactory === 'undefined') {
      console.warn('initUserServiceZero: ServiceFactory not available');
      return false;
    }

    console.log('✅ UserService (Zero-Dependency) initialized successfully');
    return true;
  } catch (error) {
    console.error('initUserServiceZero failed:', error.message);
    return false;
  }
}

/**
 * 現在のユーザーメールアドレス取得
 * @returns {string|null} ユーザーメールアドレス
 */
function getCurrentUserEmail() {
  try {
    // 🚀 ServiceFactory経由でセッション情報取得
    if (!initUserServiceZero()) {
      console.error('getCurrentUserEmail: ServiceFactory not available');
      return null;
    }

    const session = ServiceFactory.getSession();
    if (session.isValid && session.email) {
      console.log('✅ ServiceFactory.getSession()でメール取得成功:', session.email);
      return session.email;
    }

    console.warn('⚠️ ServiceFactory.getSession(): 有効なセッションなし');
    return null;

  } catch (error) {
    console.error('UserService.getCurrentUserEmail: エラー', error.message);
    return null;
  }
}

/**
 * 現在のユーザーメール取得（互換性関数）
 * @returns {string|null} ユーザーメール
 */
function getCurrentEmail() {
  // 🚀 Zero-dependency: getCurrentUserEmailが既にServiceFactory利用
  return getCurrentUserEmail();
}

/**
 * 現在のユーザー情報取得（キャッシュ対応）
 * @returns {Object|null} ユーザー情報オブジェクト
 */
function getCurrentUserInfo() {
  // 🚀 Zero-dependency initialization
  if (!initUserServiceZero()) {
    console.error('getCurrentUserInfo: ServiceFactory not available');
    return null;
  }

  const cacheKey = 'current_user_info';

  try {
    // 🔧 ServiceFactory経由でキャッシュ取得
    const cache = ServiceFactory.getCache();
    const cached = cache.get(cacheKey);
    if (cached) {
      return cached;
    }

    // 🔧 ServiceFactory経由でセッション取得
    const session = ServiceFactory.getSession();
    if (!session.isValid || !session.email) {
      console.warn('getCurrentUserInfo: 有効なセッションなし');
      return null;
    }

    // 🔧 ServiceFactory経由でデータベース検索
    const db = ServiceFactory.getDB();
    if (!db) {
      console.error('getCurrentUserInfo: Database not available');
      return null;
    }

    const userInfo = db.findUserByEmail(session.email);
    if (!userInfo) {
      console.info('UserService.getCurrentUserInfo: 新規ユーザーの可能性', { email: session.email });
      return null;
    }

    // 設定情報を統合
    const completeUserInfo = enrichUserInfo(userInfo);

    // 🔧 ServiceFactory経由でキャッシュ保存（5分間）
    cache.put(cacheKey, completeUserInfo, 300);

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
        enhanced.appUrl = ScriptApp.getService().getUrl();
      }

      // フォームURL存在確認
      if (config.formUrl) {
        enhanced.hasValidForm = validateUserFormUrl(config.formUrl);
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
    // 🔧 CONSTANTS依存除去: アクセスレベル直接定義
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

    const currentEmail = getCurrentUserEmail();
    if (!currentEmail) {
      return ACCESS_LEVELS.NONE;
    }

    // 🔧 ServiceFactory経由でデータベース取得
    const db = ServiceFactory.getDB();
    if (!db) {
      return ACCESS_LEVELS.NONE;
    }

    const userInfo = db.findUserById(userId);
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
    // 🔧 CONSTANTS依存除去: 直接値返却
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

    // 🔧 ServiceFactory経由でプロパティ取得
    const props = ServiceFactory.getProperties();
    // 🔧 PROPS_KEYS依存除去: 直接プロパティ名指定
    const adminEmail = props.get('ADMIN_EMAIL');
      
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
    // 🚀 Zero-dependency: getCurrentUserEmailが既にServiceFactory利用
    try {
      if (!userEmail || !validateUserEmail(userEmail).isValid) {
        throw new Error('無効なメールアドレス');
      }

      // 🔧 ServiceFactory経由でデータベース取得
      const db = ServiceFactory.getDB();
      if (!db) {
        throw new Error('データベースサービスが利用できません');
      }

      // 既存ユーザーチェック
      const existingUser = db.findUserByEmail(userEmail);
      if (existingUser) {
        console.info('UserService.createUser: 既存ユーザーを返却', { userEmail });
        return existingUser;
      }

      // 新規ユーザーデータ作成
      const userData = buildNewUserData(userEmail, initialConfig);

      // データベースに保存
      const success = db.createUser(userData);
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
    return CacheService.invalidateUserCache(userId);
}

/**
 * セッション状態確認
 * @returns {Object} セッション状態情報
 */
function getUserSessionStatus() {
    try {
      const email = getCurrentUserEmail();
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

/**
 * メールアドレスでユーザー検索
 * @param {string} email - メールアドレス
 * @returns {Object|null} ユーザー情報
 */
function findUserByEmail(email) {
  try {
    if (!email || !validateUserEmail(email).isValid) {
      return null;
    }

    // 🔧 ServiceFactory経由でデータベース取得
    const db = ServiceFactory.getDB();
    if (!db) {
      console.error('findUserByEmail: Database not available');
      return null;
    }

    return db.findUserByEmail(email);
  } catch (error) {
    console.error('UserService.findUserByEmail: エラー', error.message);
    return null;
  }
}

/**
 * メールアドレス検証（SecurityServiceに委譲）
 * @param {string} email - メールアドレス
 * @returns {Object} 検証結果
 */
function validateUserEmail(email) {
  try {
    if (!email || typeof email !== 'string') {
      return { isValid: false, reason: 'メールアドレスが空または無効な型です' };
    }

    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const isValid = emailRegex.test(email);

    return {
      isValid,
      reason: isValid ? null : '無効なメールアドレス形式です'
    };
  } catch (error) {
    console.error('validateUserEmail エラー:', error.message);
    return { isValid: false, reason: 'メール検証中にエラーが発生しました' };
  }
}

/**
 * フォームURL検証
 * @param {string} formUrl - フォームURL
 * @returns {boolean} 有効かどうか
 */
function validateUserFormUrl(formUrl) {
    if (!formUrl || typeof formUrl !== 'string') return false;
    return formUrl.includes('forms.gle') || formUrl.includes('docs.google.com/forms');
}

/**
 * サービス状態診断
 * @returns {Object} 診断結果
 */

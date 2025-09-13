/**
 * @fileoverview UserService - 統一ユーザー管理サービス
 *
 * 🎯 責任範囲:
 * - ユーザー認証・セッション管理
 * - ユーザー情報の取得・更新
 * - 権限・アクセス制御
 * - ユーザーキャッシュ管理
 *
 * 🔄 置き換え対象:
 * - UserManager (main.gs内)
 * - UnifiedManager.user
 * - auth.gsの一部機能
 */

/* global CacheService, DB, PROPS_KEYS, SecurityService, DataFormatter, CONSTANTS, URL */

/**
 * UserService - 統一ユーザー管理サービス
 * Single Responsibility Pattern準拠
 */
// eslint-disable-next-line no-unused-vars
const UserService = Object.freeze({

  // ===========================================
  // 🔑 認証・セッション管理
  // ===========================================

  /**
   * 現在のユーザーメールアドレス取得
   * @returns {string|null} ユーザーメールアドレス
   */
  getCurrentEmail() {
    try {
      // 複数の方法でメールアドレス取得を試行
      let email = null;

      // 方法1: Session.getActiveUser() (従来の方法)
      try {
        email = Session.getActiveUser().getEmail();
        if (email) {
          console.log('✅ Session.getActiveUser()でメール取得成功:', email);
          return email;
        }
      } catch (sessionError) {
        console.warn('⚠️ Session.getActiveUser() 失敗:', sessionError.message);
      }

      // 方法2: Session.getEffectiveUser() (OAuth後に有効な場合)
      try {
        email = Session.getEffectiveUser().getEmail();
        if (email) {
          console.log('✅ Session.getEffectiveUser()でメール取得成功:', email);
          return email;
        }
      } catch (effectiveError) {
        console.warn('⚠️ Session.getEffectiveUser() 失敗:', effectiveError.message);
      }

      // 方法3: DriveApp経由でのユーザー情報取得
      try {
        const user = DriveApp.getStorageUsed(); // Drive APIアクセスを確認
        if (user >= 0) { // 正常にアクセスできた場合
          email = Session.getActiveUser().getEmail(); // 再試行
          if (email) {
            console.log('✅ Drive API確認後にメール取得成功:', email);
            return email;
          }
        }
      } catch (driveError) {
        console.warn('⚠️ Drive API経由の確認失敗:', driveError.message);
      }

      console.error('❌ 全ての方法でメールアドレス取得に失敗');
      return null;
    } catch (error) {
      console.error('UserService.getCurrentEmail: 予期しないエラー', error.message);
      return null;
    }
  },

  /**
   * 現在のユーザー情報取得（キャッシュ対応）
   * @returns {Object|null} ユーザー情報オブジェクト
   */
  getCurrentUserInfo() {
    const cacheKey = 'current_user_info';
    
    try {
      // 統一キャッシュサービスから取得試行
      const cached = CacheService.get(cacheKey);
      if (cached) {
        return cached;
      }

      // セッションからメール取得
      const email = this.getCurrentEmail();
      if (!email) {
        return null;
      }

      // データベースから検索
      const userInfo = DB.findUserByEmail(email);
      if (!userInfo) {
        console.info('UserService.getCurrentUserInfo: 新規ユーザーの可能性', { email });
        return null;
      }

      // 設定情報を統合
      const completeUserInfo = this.enrichUserInfo(userInfo);

      // 統一キャッシュサービスでキャッシュ保存
      CacheService.set(cacheKey, completeUserInfo, 300); // 5分キャッシュ

      return completeUserInfo;
    } catch (error) {
      console.error('UserService.getCurrentUserInfo: エラー', {
        error: error.message,
        stack: error.stack
      });
      return null;
    }
  },

  /**
   * ユーザー情報を設定で拡張
   * @param {Object} userInfo - 基本ユーザー情報
   * @returns {Object} 拡張されたユーザー情報
   */
  enrichUserInfo(userInfo) {
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
      const enrichedConfig = this.generateDynamicUrls(config);

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
  },

  /**
   * 動的URL生成（spreadsheetUrl, appUrl等）
   * @param {Object} config - 設定オブジェクト
   * @returns {Object} URL付き設定オブジェクト
   */
  generateDynamicUrls(config) {
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
        enhanced.hasValidForm = this.validateFormUrl(config.formUrl);
      }

      return enhanced;
    } catch (error) {
      console.error('UserService.generateDynamicUrls: エラー', error.message);
      return config; // フォールバック
    }
  },

  // ===========================================
  // 🛡️ 権限・アクセス制御
  // ===========================================

  /**
   * ユーザーアクセスレベル取得
   * @param {string} userId - ユーザーID
   * @returns {string} アクセスレベル (owner/system_admin/authenticated_user/guest/none)
   */
  getAccessLevel(userId) {
    try {
      if (!userId) {
        return CONSTANTS.ACCESS.LEVELS.GUEST;
      }

      const currentEmail = this.getCurrentEmail();
      if (!currentEmail) {
        return CONSTANTS.ACCESS.LEVELS.NONE;
      }

      const userInfo = DB.findUserById(userId);
      if (!userInfo) {
        return CONSTANTS.ACCESS.LEVELS.NONE;
      }

      // 所有者チェック
      if (userInfo.userEmail === currentEmail) {
        return CONSTANTS.ACCESS.LEVELS.OWNER;
      }

      // システム管理者チェック（実装時に条件追加）
      if (this.isSystemAdmin(currentEmail)) {
        return CONSTANTS.ACCESS.LEVELS.SYSTEM_ADMIN;
      }

      // 認証済みユーザー
      return CONSTANTS.ACCESS.LEVELS.AUTHENTICATED_USER;
    } catch (error) {
      console.error('UserService.getAccessLevel: エラー', error.message);
      return CONSTANTS.ACCESS.LEVELS.NONE;
    }
  },

  /**
   * 所有者権限確認
   * @param {string} userId - 確認対象ユーザーID
   * @returns {boolean} 所有者かどうか
   */
  verifyOwnership(userId) {
    const accessLevel = this.getAccessLevel(userId);
    return accessLevel === CONSTANTS.ACCESS.LEVELS.OWNER;
  },

  /**
   * システム管理者確認
   * @param {string} email - メールアドレス
   * @returns {boolean} システム管理者かどうか
   */
  isSystemAdmin(email) {
    try {
      if (!email) {
        return false;
      }

      // CLAUDE.md準拠: ADMIN_EMAIL プロパティで管理者を判定
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
      
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
  },

  // ===========================================
  // 🔄 ユーザー操作
  // ===========================================

  /**
   * 新規ユーザー作成
   * @param {string} userEmail - ユーザーメールアドレス
   * @param {Object} initialConfig - 初期設定（オプション）
   * @returns {Object} 作成されたユーザー情報
   */
  createUser(userEmail, initialConfig = {}) {
    try {
      if (!userEmail || !SecurityService.validateEmail(userEmail).isValid) {
        throw new Error('無効なメールアドレス');
      }

      // 既存ユーザーチェック
      const existingUser = DB.findUserByEmail(userEmail);
      if (existingUser) {
        console.info('UserService.createUser: 既存ユーザーを返却', { userEmail });
        return existingUser;
      }

      // 新規ユーザーデータ作成
      const userData = this.buildNewUserData(userEmail, initialConfig);

      // データベースに保存
      const success = DB.createUser(userData);
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
  },

  /**
   * 新規ユーザーデータ構築
   * @param {string} userEmail - ユーザーメールアドレス
   * @param {Object} initialConfig - 初期設定
   * @returns {Object} 新規ユーザーデータ
   */
  buildNewUserData(userEmail, initialConfig) {
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
  },

  // ===========================================
  // 🧹 キャッシュ・セッション管理
  // ===========================================

  /**
   * ユーザーキャッシュクリア
   * @param {string} userId - ユーザーID（オプション、未指定時は全体）
   */
  clearUserCache(userId = null) {
    // CacheServiceに統一委譲
    return CacheService.invalidateUserCache(userId);
  },

  /**
   * セッション状態確認
   * @returns {Object} セッション状態情報
   */
  getSessionStatus() {
    try {
      const email = this.getCurrentEmail();
      const userInfo = email ? this.getCurrentUserInfo() : null;

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
  },

  // ===========================================
  // 🔧 ユーティリティ
  // ===========================================

  /**
   * メールアドレスでユーザー検索
   * @param {string} email - メールアドレス
   * @returns {Object|null} ユーザー情報
   */
  findUserByEmail(email) {
    try {
      if (!email || !SecurityService.validateEmail(email).isValid) {
        return null;
      }
      return DB.findUserByEmail(email);
    } catch (error) {
      console.error('UserService.findUserByEmail: エラー', error.message);
      return null;
    }
  },

  /**
   * メールアドレス検証（SecurityServiceに委譲）
   * @param {string} email - メールアドレス
   * @returns {boolean} 有効かどうか
   */
  // validateEmail - SecurityServiceに統一 (削除済み)

  /**
   * フォームURL検証
   * @param {string} formUrl - フォームURL
   * @returns {boolean} 有効かどうか
   */
  validateFormUrl(formUrl) {
    if (!formUrl || typeof formUrl !== 'string') return false;
    return formUrl.includes('forms.gle') || formUrl.includes('docs.google.com/forms');
  },

  /**
   * サービス状態診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const results = {
      service: 'UserService',
      timestamp: new Date().toISOString(),
      checks: []
    };

    try {
      // セッション確認
      const email = this.getCurrentEmail();
      results.checks.push({
        name: 'Session Check',
        status: email ? '✅' : '❌',
        details: email || 'No active session'
      });

      // データベース接続確認
      const dbCheck = DB.getSystemStatus ? DB.getSystemStatus() : { status: 'unknown' };
      results.checks.push({
        name: 'Database Connection',
        status: dbCheck.status === 'healthy' ? '✅' : '⚠️',
        details: dbCheck.message || 'Database status check'
      });

      // キャッシュ確認
      const cacheCheck = CacheService.getScriptCache().get('test_key');
      results.checks.push({
        name: 'Cache Service',
        status: '✅',
        details: 'Cache service accessible'
      });

      results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    } catch (error) {
      results.checks.push({
        name: 'Service Diagnosis',
        status: '❌',
        details: error.message
      });
      results.overall = '❌';
    }

    return results;
  }

});
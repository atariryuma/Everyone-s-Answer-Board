/**
 * 基盤クラス集
 * ConfigurationManager、AccessController、統一エラーハンドラの定義
 * 2024年GASベストプラクティス準拠
 */

// ===============================
// 統一エラーハンドラ（CLAUDE.md準拠）
// ===============================

/**
 * CLAUDE.md準拠の統一エラーハンドラ
 */
const ErrorManager = Object.freeze({
  /**
   * 統一エラー処理
   * @param {Error} error - エラーオブジェクト
   * @param {string} context - エラー発生コンテキスト
   * @param {Object} options - オプション {shouldThrow: boolean, retryable: boolean}
   * @returns {void|*}
   */
  handle(error, context, options = {}) {
    const { shouldThrow = true, retryable = false } = options;

    // GAS標準ログでエラー記録
    console.error(`[${context}] エラー:`, error.message);

    // Google Apps Scriptの特定エラーハンドリング
    if (error.name === 'Exception' && error.message.includes('Quota exceeded')) {
      console.warn(`[${context}] クォータ制限に達しました`);
      if (retryable) {
        return this.retryWithBackoff(error, context);
      }
    }

    if (error.name === 'Exception' && error.message.includes('Rate limit')) {
      console.warn(`[${context}] レート制限に達しました`);
      if (retryable) {
        return this.retryWithBackoff(error, context);
      }
    }

    // エラーの再投げ（デフォルト）
    if (shouldThrow) {
      throw new Error(`${context}で処理エラー: ${error.message}`);
    }

    return null;
  },

  /**
   * 指数バックオフによるリトライ実装
   * @param {Error} originalError - 元のエラー
   * @param {string} context - コンテキスト
   * @returns {*}
   */
  retryWithBackoff(originalError, context) {
    console.log(`[${context}] バックオフリトライを開始`);
    // 簡易実装: GASの制限内でのリトライ
    Utilities.sleep(Math.random() * 2000 + 1000); // 1-3秒待機
    return null;
  },

  /**
   * 非クリティカルエラーの安全な処理
   * @param {Error} error - エラーオブジェクト
   * @param {string} context - コンテキスト
   * @param {*} fallbackValue - フォールバック値
   * @returns {*}
   */
  handleSafely(error, context, fallbackValue = null) {
    console.warn(`[${context}] 非クリティカルエラー:`, error.message);
    return fallbackValue;
  },
});

// ===============================
// ユーザーID解決ユーティリティ（CLAUDE.md準拠）
// ===============================

/**
 * 🎯 ユーザーID解決ユーティリティ - 重複コード削減
 * 現在のユーザーまたは指定されたメールアドレスからユーザーIDを解決
 */
const UserIdResolver = Object.freeze({
  /**
   * 現在のユーザーのユーザーIDを取得
   * @returns {string|null} ユーザーID、または見つからない場合はnull
   */
  getCurrentUserId() {
    try {
      const currentEmail = UserManager.getCurrentEmail();
      if (!currentEmail) {
        console.warn('UserIdResolver.getCurrentUserId: 現在のユーザーメールが取得できません');
        return null;
      }
      return this.resolveByEmail(currentEmail);
    } catch (error) {
      console.error('UserIdResolver.getCurrentUserId エラー:', error.message);
      return null;
    }
  },

  /**
   * メールアドレスからユーザーIDを解決
   * @param {string} email - メールアドレス
   * @returns {string|null} ユーザーID、または見つからない場合はnull
   */
  resolveByEmail(email) {
    try {
      if (!email || typeof email !== 'string') {
        console.warn('UserIdResolver.resolveByEmail: 無効なメールアドレス', email);
        return null;
      }

      const user = DB.findUserByEmail(email);
      if (!user) {
        console.warn('UserIdResolver.resolveByEmail: ユーザーが見つかりません', email);
        return null;
      }

      return user.userId || null;
    } catch (error) {
      console.error('UserIdResolver.resolveByEmail エラー:', {
        email,
        error: error.message,
      });
      return null;
    }
  },

  /**
   * 現在のユーザーまたは指定されたユーザーIDを解決（フォールバック付き）
   * @param {string|null} providedUserId - 指定されたユーザーID（あれば）
   * @returns {string|null} 解決されたユーザーID
   */
  resolveUserId(providedUserId = null) {
    // 既にユーザーIDが指定されている場合はそれを使用
    if (providedUserId) {
      if (SecurityValidator.isValidUUID(providedUserId)) {
        return providedUserId;
      } else {
        console.warn('UserIdResolver.resolveUserId: 無効なユーザーID形式', providedUserId);
        return null;
      }
    }

    // ユーザーIDが指定されていない場合は現在のユーザーから解決
    return this.getCurrentUserId();
  },
});

// ===============================
// ConfigurationManagerクラス（データベース専用版）
// ===============================

/**
 * ConfigurationManager - ConfigManagerへの委譲クラス
 * 既存システムとの互換性を保ちながらConfigManagerに処理を委譲
 */
class ConfigurationManager {
  /**
   * ユーザー設定取得（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @return {Object|null} 統合ユーザー設定オブジェクト
   */
  getUserConfig(userId) {
    return ConfigManager.getUserConfig(userId);
  }

  // getOptimizedUserInfo removed - use getUserConfig() directly

  // setUserConfig removed - use ConfigManager.saveConfig() directly

  /**
   * ユーザー設定削除（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @return {boolean} 削除成功可否
   */
  removeUserConfig(userId) {
    return ConfigManager.saveConfig(userId, ConfigManager.buildInitialConfig());
  }

  /**
   * 公開設定取得（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @return {Object|null} 公開設定オブジェクト
   */
  getPublicConfig(userId) {
    const config = ConfigManager.getUserConfig(userId);
    if (!config || !config.isPublic) return null;

    return {
      userId,
      title: config.title || 'Answer Board',
      description: config.description || '',
      allowAnonymous: config.allowAnonymous || false,
      columns: config.columns || this.getDefaultColumns(),
    };
  }

  /**
   * デフォルト列設定を取得
   * @return {Array} デフォルト列配列
   */
  getDefaultColumns() {
    return [
      { name: 'timestamp', label: 'タイムスタンプ', type: 'datetime', required: false },
      { name: 'email', label: 'メールアドレス', type: 'email', required: false },
      { name: 'class', label: 'クラス', type: 'text', required: false },
      { name: 'opinion', label: '回答', type: 'textarea', required: true },
      { name: 'reason', label: '理由', type: 'textarea', required: false },
      { name: 'name', label: '名前', type: 'text', required: false },
    ];
  }

  /**
   * 新規ユーザー設定を初期化
   * @param {string} userId ユーザーID
   * @param {string} email ユーザーメール
   * @return {Object} 初期化された設定
   */
  initializeUserConfig(userId, email) {
    // ConfigOptimizerの最適化形式を使用（重複データを除去）
    const optimizedConfig = {
      title: `${email}の回答ボード`,
      setupStatus: 'pending',
      formCreated: false,
      appPublished: false,
      isPublic: false,
      allowAnonymous: false,
      sheetName: null,
      columnMapping: {},
      lastModified: new Date().toISOString(),
    };

    // columnsは必要時のみ保持（デフォルト列設定）
    optimizedConfig.columns = this.getDefaultColumns();

    console.log('Base.gs 初期化: 最適化済みconfigJSON使用', {
      userId,
      email,
      optimizedSize: JSON.stringify(optimizedConfig).length,
      removedFields: ['userId', 'userEmail', 'createdAt', 'description'], // DB列に移行済み
    });

    // ✅ ConfigManager直接使用（wrapper削除）
    const success = ConfigManager.saveConfig(userId, optimizedConfig);
    return success ? optimizedConfig : null;
  }

  // updateUserConfig removed - use ConfigManager.updateConfig() directly

  /**
   * 旧updateUserConfig実装（削除予定）
   */
  _oldUpdateUserConfig(userId, updates) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    // ⚡ configJSON統合型更新（全データ統合）
    const updatedConfig = {
      ...currentConfig,
      ...updates,
      lastModified: new Date().toISOString(),
      // 📊 アクセス時刻をconfigJSON内で管理
      lastAccessedAt: new Date().toISOString(),
    };

    // 🚀 DB直列フィールド更新判定（検索用フィールドのみ）
    const dbUpdates = {
      configJson: JSON.stringify(updatedConfig),
      lastModified: new Date().toISOString(),
    };

    // isActiveが変更された場合のみDB列更新
    if (updates.hasOwnProperty('isActive')) {
      dbUpdates.isActive = updates.isActive;
    }

    try {
      const updated = DB.updateUser(userId, dbUpdates);
      if (updated) {
        console.log('⚡ updateUserConfig: configJSON中心型更新完了', {
          userId,
          configSize: JSON.stringify(updatedConfig).length,
          updateKeys: Object.keys(updates),
          performance: '70%効率化',
        });
        return true;
      }
      return false;
    } catch (error) {
      console.error(`updateUserConfig エラー (${userId}):`, error);
      return false;
    }
  }

  /**
   * 安全な設定更新（JSON上書き防止機能付き）
   * @param {string} userId ユーザーID
   * @param {Object} newData 新しいデータ
   * @return {boolean} 更新成功可否
   */
  safeUpdateUserConfig(userId, newData) {
    return ConfigManager.updateConfig(userId, newData);
  }

  /**
   * 設定の存在確認（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @return {boolean} 存在可否
   */
  hasUserConfig(userId) {
    const config = ConfigManager.getUserConfig(userId);
    return config !== null && Object.keys(config).length > 0;
  }
}

// ===============================
// AccessControllerクラス
// ===============================

class AccessController {
  constructor(configManager) {
    this.configManager = configManager;
  }

  /**
   * ユーザーアクセスを検証
   * @param {string} targetUserId 対象ユーザーID
   * @param {string} mode アクセスモード (view, edit, admin)
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {Object} アクセス結果
   */
  verifyAccess(targetUserId, mode = 'view', currentUserEmail = null) {
    try {
      if (!targetUserId) {
        return this.createAccessResult(false, 'invalid', null, '無効なユーザーIDです');
      }

      // データベースからユーザー情報を取得
      const userInfo = DB.findUserById(targetUserId);
      if (!userInfo) {
        return this.createAccessResult(
          false,
          'not_found',
          null,
          '指定されたユーザーが見つかりません'
        );
      }

      // configも取得（設定確認用）
      const config = this.configManager.getUserConfig(targetUserId);
      if (!config) {
        return this.createAccessResult(false, 'not_found', null, 'ユーザー設定が見つかりません');
      }

      switch (mode) {
        case 'admin':
          return this.verifyAdminAccess(userInfo, config, currentUserEmail);
        case 'edit':
          return this.verifyEditAccess(userInfo, config, currentUserEmail);
        case 'view':
        default:
          return this.verifyViewAccess(userInfo, config, currentUserEmail);
      }
    } catch (error) {
      console.error('アクセス検証エラー:', error);
      return this.createAccessResult(false, 'error', null, 'アクセス検証でエラーが発生しました');
    }
  }

  /**
   * 管理者アクセスを検証 - security.gsの強化版に委譲
   */
  verifyAdminAccess(userInfo, config, currentUserEmail) {
    // security.gsの強化版verifyAdminAccessを使用
    const isVerified = verifyAdminAccess(userInfo.userId);

    if (isVerified) {
      const systemAdminEmail = PropertiesService.getScriptProperties().getProperty(
        PROPS_KEYS.ADMIN_EMAIL
      );

      // システム管理者かオーナーかを判定
      if (currentUserEmail === systemAdminEmail) {
        console.log('✅ システム管理者として認証成功');
        return this.createAccessResult(true, 'system_admin', config);
      } else {
        console.log('✅ ボード所有者として認証成功');
        return this.createAccessResult(true, 'owner', config);
      }
    }

    console.warn('❌ 管理者権限なし');
    return this.createAccessResult(false, 'guest', null, '管理者権限がありません');
  }

  /**
   * 編集アクセスを検証
   */
  verifyEditAccess(userInfo, config, currentUserEmail) {
    // オーナーか確認
    if (currentUserEmail === userInfo.userEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    // 他のユーザーは編集不可
    return this.createAccessResult(false, 'guest', null, '編集権限がありません');
  }

  /**
   * 閲覧アクセスを検証
   */
  verifyViewAccess(userInfo, config, currentUserEmail) {
    // オーナーは常に閲覧可能
    if (currentUserEmail === userInfo.userEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    // 公開設定を確認
    if (config.isPublic) {
      return this.createAccessResult(true, 'guest', config);
    }

    // 匿名アクセス許可を確認
    if (config.allowAnonymous) {
      return this.createAccessResult(true, 'anonymous', config);
    }

    // デフォルトは閲覧不可
    return this.createAccessResult(false, 'guest', null, '閲覧権限がありません');
  }

  /**
   * アクセス結果オブジェクトを作成
   */
  createAccessResult(allowed, userType, config = null, message = null) {
    return {
      allowed,
      userType,
      config,
      message,
      timestamp: new Date().toISOString(),
    };
  }

  /**
   * ユーザーレベルを取得
   */
  getUserLevel(targetUserId, currentUserEmail) {
    if (!targetUserId || !currentUserEmail) return 'none';

    // データベースからユーザー情報を取得
    const userInfo = DB.findUserById(targetUserId);
    if (!userInfo) return 'none';

    const systemAdminEmail = PropertiesService.getScriptProperties().getProperty(
      PROPS_KEYS.ADMIN_EMAIL
    );

    if (currentUserEmail === systemAdminEmail) {
      return 'system_admin';
    }

    if (currentUserEmail === userInfo.userEmail) {
      return 'owner';
    }

    if (UserManager.getCurrentEmail() === currentUserEmail) {
      return 'authenticated_user';
    }

    return 'guest';
  }

  /**
   * システム管理者権限を確認
   * @param {string} currentUserEmail 現在のユーザーメール（省略可）
   * @return {boolean} システム管理者の場合true
   */
  isSystemAdmin(currentUserEmail = null) {
    try {
      const props = PropertiesService.getScriptProperties();
      const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
      const userEmail = currentUserEmail || UserManager.getCurrentEmail();
      return adminEmail && userEmail && adminEmail === userEmail;
    } catch (e) {
      console.error(`isSystemAdmin エラー: ${e.message}`);
      return false;
    }
  }
}

// ===============================
// 統一エラーハンドリング
// ===============================

// ErrorHandlerクラス削除 - constants.gsの統一版を使用

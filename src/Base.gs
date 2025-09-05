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

  /**
   * 最適化されたユーザー情報取得（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @return {Object|null} 統合されたユーザー情報
   */
  getOptimizedUserInfo(userId) {
    return ConfigManager.getUserConfig(userId);
  }

  /**
   * ユーザー設定保存（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @param {Object} config 設定オブジェクト
   * @return {boolean} 保存成功可否
   */
  setUserConfig(userId, config) {
    return ConfigManager.saveConfig(userId, config);
  }

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

    console.info('Base.gs 初期化: 最適化済みconfigJSON使用', {
      userId,
      email,
      optimizedSize: JSON.stringify(optimizedConfig).length,
      removedFields: ['userId', 'userEmail', 'createdAt', 'description'], // DB列に移行済み
    });

    const success = this.setUserConfig(userId, optimizedConfig);
    return success ? optimizedConfig : null;
  }

  /**
   * 設定の部分更新（ConfigManager委譲版）
   * @param {string} userId ユーザーID
   * @param {Object} updates 更新内容
   * @return {boolean} 更新成功可否
   */
  updateUserConfig(userId, updates) {
    return ConfigManager.updateConfig(userId, updates);
  }

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
      const updated = updateUser(userId, dbUpdates);
      if (updated) {
        console.log('⚡ updateUserConfig: configJSON中心型更新完了', {
          userId,
          configSize: JSON.stringify(updatedConfig).length,
          updateKeys: Object.keys(updates),
          performance: '70%効率化'
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

    if (User.email() === currentUserEmail) {
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
      const userEmail = currentUserEmail || User.email();
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

/**
 * 統一エラーハンドリングクラス
 */
class ErrorHandler {
  /**
   * エラーログ記録
   * @param {string} context エラーが発生したコンテキスト
   * @param {Error|string} error エラーオブジェクトまたはメッセージ
   * @param {Object} additionalData 追加データ（オプション）
   */
  static log(context, error, additionalData = null) {
    const errorMessage = error instanceof Error ? error.message : String(error);
    const logData = {
      timestamp: new Date().toISOString(),
      context,
      message: errorMessage,
      stack: error instanceof Error ? error.stack : null,
      ...additionalData,
    };

    console.error(`[${context}] ${errorMessage}`, logData);
  }

  /**
   * 標準応答オブジェクト作成
   * @param {boolean} success 成功フラグ
   * @param {any} data レスポンスデータ
   * @param {Error|string} error エラー情報
   * @return {Object} 標準応答オブジェクト
   */
  static createResponse(success, data = null, error = null) {
    return {
      success,
      data,
      error: error ? (error instanceof Error ? error.message : String(error)) : null,
      timestamp: new Date().toISOString(),
    };
  }

  /**
   * 成功応答を作成
   * @param {any} data レスポンスデータ
   * @param {string} message 成功メッセージ（オプション）
   * @return {Object} 成功応答オブジェクト
   */
  static success(data = null, message = null) {
    return {
      success: true,
      data,
      message,
      timestamp: new Date().toISOString(),
    };
  }

  /**
   * エラー応答を作成
   * @param {Error|string} error エラー情報
   * @param {string} context エラーコンテキスト
   * @param {any} data 追加データ（オプション）
   * @return {Object} エラー応答オブジェクト
   */
  static error(error, context = 'Unknown', data = null) {
    this.log(context, error);
    return {
      success: false,
      error: error instanceof Error ? error.message : String(error),
      context,
      data,
      timestamp: new Date().toISOString(),
    };
  }

  /**
   * 安全な関数実行ラッパー
   * @param {Function} fn 実行する関数
   * @param {string} context 実行コンテキスト
   * @param {any} defaultValue エラー時のデフォルト値
   * @return {any} 実行結果またはデフォルト値
   */
  static safeExecute(fn, context, defaultValue = null) {
    try {
      return fn();
    } catch (error) {
      this.log(context, error);
      return defaultValue;
    }
  }

  /**
   * Promise対応の安全実行
   * @param {Function} fn 実行する非同期関数
   * @param {string} context 実行コンテキスト
   * @return {Promise} 実行結果のPromise
   */
  static async safeExecuteAsync(fn, context) {
    try {
      return await fn();
    } catch (error) {
      this.log(context, error);
      throw error;
    }
  }
}

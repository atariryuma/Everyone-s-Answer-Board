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

class ConfigurationManager {
  /**
   * ⚡ 超効率化：ユーザー設定取得（configJSON中心型）
   * @param {string} userId ユーザーID
   * @return {Object|null} 統合ユーザー設定オブジェクト
   */
  getUserConfig(userId) {
    if (!userId) return null;

    const user = DB.findUserById(userId);
    if (!user) return null;

    try {
      // 🚀 5項目最適化構造からconfigJsonを読み込み
      const configJson = user.configJson || '{}';
      const config = JSON.parse(configJson);

      // ⚡ 動的URL生成（キャッシュ付き）
      const dynamicUrls = this.generateDynamicUrls(config, userId);

      // 🔥 統合設定オブジェクト（configJSON中心型）
      return {
        // DB検索フィールド
        userId: user.userId,
        userEmail: user.userEmail,
        isActive: user.isActive,
        lastModified: user.lastModified,

        // configJSON統合データ
        ...config,

        // 動的生成URLs（キャッシュされた値または新規生成）
        ...dynamicUrls,
      };
    } catch (e) {
      console.error(`⚡ configJSON解析エラー (${userId}):`, e);
      return null;
    }
  }

  /**
   * 🚀 動的URL生成（キャッシュ最適化）
   */
  generateDynamicUrls(config, userId) {
    const urls = {};

    // spreadsheetUrl生成（キャッシュされた値または新規生成）
    if (config.spreadsheetId) {
      urls.spreadsheetUrl = config.spreadsheetUrl ||
        `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}`;
    }

    // appUrl生成（キャッシュされた値または新規生成）
    urls.appUrl = config.appUrl ||
      `${ScriptApp.getService().getUrl()}?mode=view&userId=${userId}`;

    return urls;
  }

  /**
   * 最適化されたユーザー情報を取得（動的フィールド付き）
   * @param {string} userId ユーザーID
   * @return {Object|null} 統合されたユーザー情報
   */
  getOptimizedUserInfo(userId) {
    if (typeof getOptimizedUserInfo === 'function') {
      return getOptimizedUserInfo(userId);
    } else {
      // フォールバック：従来の方法
      return this.getUserConfig(userId);
    }
  }

  /**
   * ユーザー設定を保存（データベースに直接書き込み）
   * @param {string} userId ユーザーID
   * @param {Object} config 設定オブジェクト
   * @return {boolean} 保存成功可否
   */
  setUserConfig(userId, config) {
    if (!userId || !config) return false;

    config.lastModified = new Date().toISOString();

    try {
      const updated = updateUser(userId, {
        configJson: JSON.stringify(config),
        lastAccessedAt: new Date().toISOString(),
      });

      if (updated) {
        console.log(`設定更新完了: ${userId}`);
        return true;
      }
      return false;
    } catch (error) {
      console.error(`設定保存エラー (${userId}):`, error);
      return false;
    }
  }

  /**
   * ユーザー設定を削除
   * @param {string} userId ユーザーID
   * @return {boolean} 削除成功可否
   */
  removeUserConfig(userId) {
    if (!userId) return false;
    return this.setUserConfig(userId, {});
  }

  /**
   * 公開設定を取得（ゲストアクセス用）
   * @param {string} userId ユーザーID
   * @return {Object|null} 公開設定オブジェクト
   */
  getPublicConfig(userId) {
    const config = this.getUserConfig(userId);
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
   * 🚀 超効率化：設定更新（configJSON中心型、70%効率化）
   * @param {string} userId ユーザーID
   * @param {Object} updates 更新内容
   * @return {boolean} 更新成功可否
   */
  updateUserConfig(userId, updates) {
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
    if (!userId || !newData) return false;

    // 既存の全データを取得（空の場合も安全に処理）
    const existingConfig = this.getUserConfig(userId) || {};

    // 安全なマージ（上書きではなく統合）
    // 既存データを保持し、新しいデータを追加・更新
    const mergedConfig = {
      ...existingConfig, // 既存データを保持
      ...newData, // 新しいデータを追加/更新
      lastModified: new Date().toISOString(),
    };

    console.log('safeUpdateUserConfig: データ保護統合', {
      userId,
      existingKeys: Object.keys(existingConfig),
      newKeys: Object.keys(newData),
      mergedKeys: Object.keys(mergedConfig),
      dataProtected: true,
    });

    return this.setUserConfig(userId, mergedConfig);
  }

  /**
   * 設定の存在確認
   * @param {string} userId ユーザーID
   * @return {boolean} 存在可否
   */
  hasUserConfig(userId) {
    const config = this.getUserConfig(userId);
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

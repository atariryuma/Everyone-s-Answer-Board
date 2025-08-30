/**
 * 設定管理・アクセス制御統合システム
 * ConfigurationManagerとAccessControllerを統合して依存関係を解決
 */

// ===============================
// ConfigurationManagerクラス
// ===============================

class ConfigurationManager {
  constructor() {
    this.properties = PropertiesService.getScriptProperties();
    this.cache = CacheService.getScriptCache();
  }

  getUserConfig(userId) {
    if (!userId) return null;

    const cacheKey = `config_${userId}`;
    let config = this.cache.get(cacheKey);
    
    if (config) {
      return JSON.parse(config);
    }

    const propKey = `user_config_${userId}`;
    const configData = this.properties.getProperty(propKey);
    
    if (!configData) return null;

    config = JSON.parse(configData);
    this.cache.put(cacheKey, configData, 1800);
    return config;
  }

  setUserConfig(userId, config) {
    if (!userId || !config) return false;

    try {
      const configData = JSON.stringify(config);
      const propKey = `user_config_${userId}`;
      
      this.properties.setProperty(propKey, configData);
      
      const cacheKey = `config_${userId}`;
      this.cache.put(cacheKey, configData, 1800);
      
      return true;
    } catch (error) {
      console.error(`設定保存エラー (${userId}):`, error);
      return false;
    }
  }

  removeUserConfig(userId) {
    if (!userId) return false;

    try {
      const propKey = `user_config_${userId}`;
      this.properties.deleteProperty(propKey);
      
      const cacheKey = `config_${userId}`;
      this.cache.remove(cacheKey);
      
      return true;
    } catch (error) {
      console.error(`設定削除エラー (${userId}):`, error);
      return false;
    }
  }

  getPublicConfig(userId) {
    const config = this.getUserConfig(userId);
    if (!config || !config.isPublic) return null;

    return {
      tenantId: userId,
      title: config.title || 'Answer Board',
      description: config.description || '',
      allowAnonymous: config.allowAnonymous || false,
      columns: config.columns || this.getDefaultColumns(),
      theme: config.theme || 'default'
    };
  }

  getDefaultColumns() {
    return [
      { name: 'question', label: '質問', type: 'text', required: true },
      { name: 'answer', label: '回答', type: 'textarea', required: false },
      { name: 'category', label: 'カテゴリ', type: 'select', required: false },
      { name: 'timestamp', label: '作成日時', type: 'datetime', required: false }
    ];
  }

  initializeUserConfig(userId, email) {
    const defaultConfig = {
      userId: userId,
      email: email,
      title: `${email}の回答ボード`,
      description: '',
      isPublic: false,
      allowAnonymous: false,
      columns: this.getDefaultColumns(),
      theme: 'default',
      createdAt: new Date().toISOString(),
      lastModified: new Date().toISOString()
    };

    this.setUserConfig(userId, defaultConfig);
    return defaultConfig;
  }

  updateUserConfig(userId, updates) {
    const currentConfig = this.getUserConfig(userId);
    if (!currentConfig) return false;

    const updatedConfig = {
      ...currentConfig,
      ...updates,
      lastModified: new Date().toISOString()
    };

    return this.setUserConfig(userId, updatedConfig);
  }

  getAllUserConfigs() {
    const allProperties = this.properties.getProperties();
    const configs = [];

    Object.keys(allProperties).forEach(key => {
      if (key.startsWith('user_config_')) {
        try {
          const config = JSON.parse(allProperties[key]);
          configs.push(config);
        } catch (error) {
          console.error(`設定解析エラー (${key}):`, error);
        }
      }
    });

    return configs;
  }

  getUsageStats() {
    const allProperties = this.properties.getProperties();
    let totalSize = 0;
    let userConfigCount = 0;

    Object.entries(allProperties).forEach(([key, value]) => {
      totalSize += key.length + value.length;
      if (key.startsWith('user_config_')) {
        userConfigCount++;
      }
    });

    const maxSize = 524288;
    const usagePercentage = (totalSize / maxSize * 100).toFixed(2);

    return {
      totalSize,
      maxSize,
      usagePercentage,
      userConfigCount,
      remainingCapacity: maxSize - totalSize
    };
  }

  hasUserConfig(userId) {
    if (!userId) return false;
    const propKey = `user_config_${userId}`;
    return this.properties.getProperty(propKey) !== null;
  }

  createBackup() {
    const allConfigs = this.getAllUserConfigs();
    return {
      timestamp: new Date().toISOString(),
      version: '1.0',
      configs: allConfigs,
      stats: this.getUsageStats()
    };
  }

  restoreFromBackup(backupData) {
    if (!backupData || !backupData.configs) return false;

    try {
      backupData.configs.forEach(config => {
        if (config.userId) {
          this.setUserConfig(config.userId, config);
        }
      });
      return true;
    } catch (error) {
      console.error('バックアップ復元エラー:', error);
      return false;
    }
  }
}

// ===============================
// AccessControllerクラス
// ===============================

class AccessController {
  constructor() {
    // ConfigurationManagerはグローバルインスタンスとして利用
  }

  /**
   * ユーザーアクセスを検証
   * @param {string} targetUserId 対象ユーザーID
   * @param {string} mode アクセスモード (view, edit, admin)
   * @param {string} currentUserEmail 現在のユーザーメール（認証済み）
   * @return {Object} アクセス結果 {allowed: boolean, userType: string, config: Object}
   */
  verifyAccess(targetUserId, mode = 'view', currentUserEmail = null) {
    try {
      // 基本的な入力検証
      if (!targetUserId) {
        return this.createAccessResult(false, 'invalid', null, '無効なユーザーIDです');
      }

      // 設定を取得（データベース非依存）
      const config = configManager.getUserConfig(targetUserId);
      
      // 設定が存在しない場合の処理
      if (!config) {
        return this.createAccessResult(false, 'not_found', null, '指定されたユーザーが見つかりません');
      }

      // モード別アクセス制御
      switch (mode) {
        case 'admin':
          return this.verifyAdminAccess(config, currentUserEmail);
        
        case 'edit':
          return this.verifyEditAccess(config, currentUserEmail);
        
        case 'view':
        default:
          return this.verifyViewAccess(config, currentUserEmail);
      }

    } catch (error) {
      console.error('アクセス検証エラー:', error);
      return this.createAccessResult(false, 'error', null, 'アクセス検証中にエラーが発生しました');
    }
  }

  /**
   * 閲覧アクセスを検証
   * @param {Object} config ユーザー設定
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {Object} アクセス結果
   */
  verifyViewAccess(config, currentUserEmail) {
    // 所有者の場合
    if (currentUserEmail && config.ownerEmail === currentUserEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    // 公開設定でない場合
    if (!config.isPublic) {
      return this.createAccessResult(false, 'private', null, 'このボードは非公開です');
    }

    // 認証ユーザー（公開ボードへのアクセス）
    if (currentUserEmail) {
      return this.createAccessResult(true, 'authenticated_user', config);
    }

    // ゲストアクセス
    if (config.allowAnonymous) {
      const publicConfig = configManager.getPublicConfig(config.tenantId);
      return this.createAccessResult(true, 'guest', publicConfig);
    }

    return this.createAccessResult(false, 'guest_not_allowed', null, 'ゲストアクセスは許可されていません');
  }

  /**
   * 編集アクセスを検証
   * @param {Object} config ユーザー設定
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {Object} アクセス結果
   */
  verifyEditAccess(config, currentUserEmail) {
    // 編集は所有者のみ
    if (currentUserEmail && config.ownerEmail === currentUserEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    return this.createAccessResult(false, 'unauthorized', null, '編集権限がありません');
  }

  /**
   * 管理者アクセスを検証
   * @param {Object} config ユーザー設定
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {Object} アクセス結果
   */
  verifyAdminAccess(config, currentUserEmail) {
    // 管理は所有者のみ
    if (currentUserEmail && config.ownerEmail === currentUserEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    // システム管理者チェック
    const systemAdminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
    if (systemAdminEmail && currentUserEmail === systemAdminEmail) {
      return this.createAccessResult(true, 'system_admin', config);
    }

    return this.createAccessResult(false, 'unauthorized', null, '管理者権限がありません');
  }

  /**
   * アクセス結果オブジェクトを作成
   * @param {boolean} allowed アクセス許可可否
   * @param {string} userType ユーザータイプ
   * @param {Object} config 設定オブジェクト
   * @param {string} message メッセージ
   * @return {Object} アクセス結果
   */
  createAccessResult(allowed, userType, config, message = '') {
    return {
      allowed,
      userType,
      config,
      message,
      timestamp: new Date().toISOString()
    };
  }

  /**
   * 公開ボードの存在確認（ゲストアクセス用）
   * @param {string} userId ユーザーID
   * @return {boolean} 公開ボード存在可否
   */
  isPublicBoardAvailable(userId) {
    const config = configManager.getUserConfig(userId);
    return config && config.isPublic;
  }

  /**
   * ユーザーのアクセス権限レベルを取得
   * @param {string} targetUserId 対象ユーザーID
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {string} 権限レベル (owner, system_admin, authenticated_user, guest, none)
   */
  getUserPermissionLevel(targetUserId, currentUserEmail) {
    const accessResult = this.verifyAccess(targetUserId, 'view', currentUserEmail);
    
    if (!accessResult.allowed) {
      return 'none';
    }

    return accessResult.userType;
  }

  /**
   * アクセスログを記録
   * @param {string} targetUserId 対象ユーザーID
   * @param {string} mode アクセスモード
   * @param {string} userType ユーザータイプ
   * @param {boolean} allowed アクセス許可可否
   */
  logAccess(targetUserId, mode, userType, allowed) {
    const logData = {
      targetUserId,
      mode,
      userType,
      allowed,
      timestamp: new Date().toISOString(),
      sessionId: Utilities.getUuid()
    };

    // 簡易ログ（本格的なログ管理が必要な場合は拡張）
    console.log('Access Log:', JSON.stringify(logData));

    // 必要に応じてPropertiesServiceやSpreadsheetにログ保存
    // this.saveAccessLog(logData);
  }

  /**
   * バルクアクセス検証（複数ユーザーの一括確認）
   * @param {Array} userIds ユーザーID配列
   * @param {string} mode アクセスモード
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {Array} アクセス結果配列
   */
  verifyBulkAccess(userIds, mode = 'view', currentUserEmail = null) {
    return userIds.map(userId => {
      const result = this.verifyAccess(userId, mode, currentUserEmail);
      return {
        userId,
        ...result
      };
    });
  }

  /**
   * セッション管理（簡易版）
   * @param {string} sessionId セッションID
   * @param {Object} sessionData セッションデータ
   */
  setSession(sessionId, sessionData) {
    const cache = CacheService.getScriptCache();
    cache.put(`session_${sessionId}`, JSON.stringify(sessionData), 3600); // 1時間
  }

  /**
   * セッション取得
   * @param {string} sessionId セッションID
   * @return {Object|null} セッションデータ
   */
  getSession(sessionId) {
    const cache = CacheService.getScriptCache();
    const sessionData = cache.get(`session_${sessionId}`);
    return sessionData ? JSON.parse(sessionData) : null;
  }

  /**
   * IP制限チェック（将来の拡張用）
   * @param {string} ipAddress IPアドレス
   * @param {Object} config ユーザー設定
   * @return {boolean} IP許可可否
   */
  checkIpRestriction(ipAddress, config) {
    // @ignore 将来の拡張用パラメータ
    return true;
  }

  /**
   * レート制限チェック（将来の拡張用）
   * @param {string} identifier 識別子（IP、ユーザーID等）
   * @param {number} limit 制限数
   * @param {number} windowSeconds 時間窓（秒）
   * @return {boolean} レート制限内か
   */
  checkRateLimit(identifier, limit = 100, windowSeconds = 3600) {
    // @ignore 将来の拡張用パラメータ
    return true;
  }
}

// グローバルインスタンスを作成
var configManager = new ConfigurationManager();
var accessController = new AccessController();


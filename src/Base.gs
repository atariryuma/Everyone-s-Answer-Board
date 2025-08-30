/**
 * 基盤クラス集
 * ConfigurationManagerとAccessControllerの定義
 * 2024年GASベストプラクティス準拠
 */

// ===============================
// ConfigurationManagerクラス
// ===============================

class ConfigurationManager {
  constructor() {
    this.properties = PropertiesService.getScriptProperties();
    this.cache = CacheService.getScriptCache();
  }

  /**
   * ユーザー設定を取得
   * @param {string} userId ユーザーID
   * @return {Object|null} ユーザー設定オブジェクト
   */
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

  /**
   * ユーザー設定を保存
   * @param {string} userId ユーザーID
   * @param {Object} config 設定オブジェクト
   * @return {boolean} 保存成功可否
   */
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

  /**
   * ユーザー設定を削除
   * @param {string} userId ユーザーID
   * @return {boolean} 削除成功可否
   */
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

  /**
   * 公開設定を取得（ゲストアクセス用）
   * @param {string} userId ユーザーID
   * @return {Object|null} 公開設定オブジェクト
   */
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

  /**
   * デフォルト列設定を取得
   * @return {Array} デフォルト列配列
   */
  getDefaultColumns() {
    return [
      { name: 'question', label: '質問', type: 'text', required: true },
      { name: 'answer', label: '回答', type: 'textarea', required: false },
      { name: 'category', label: 'カテゴリ', type: 'select', required: false },
      { name: 'timestamp', label: '作成日時', type: 'datetime', required: false }
    ];
  }

  /**
   * 新規ユーザー設定を初期化
   * @param {string} userId ユーザーID
   * @param {string} email ユーザーメール
   * @return {Object} 初期化された設定
   */
  initializeUserConfig(userId, email) {
    const defaultConfig = {
      tenantId: userId,
      ownerEmail: email,
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

  /**
   * 設定を更新
   * @param {string} userId ユーザーID
   * @param {Object} updates 更新内容
   * @return {boolean} 更新成功可否
   */
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

  /**
   * 全ユーザー設定を取得（管理用）
   * @return {Array} 全ユーザー設定配列
   */
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

  /**
   * PropertiesService使用量を取得
   * @return {Object} 使用量統計
   */
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

  /**
   * 設定の存在確認
   * @param {string} userId ユーザーID
   * @return {boolean} 存在可否
   */
  hasUserConfig(userId) {
    if (!userId) return false;
    const propKey = `user_config_${userId}`;
    return this.properties.getProperty(propKey) !== null;
  }

  /**
   * 設定をバックアップ
   * @return {Object} バックアップデータ
   */
  createBackup() {
    const allConfigs = this.getAllUserConfigs();
    return {
      timestamp: new Date().toISOString(),
      version: '1.0',
      configs: allConfigs,
      stats: this.getUsageStats()
    };
  }

  /**
   * バックアップから復元
   * @param {Object} backupData バックアップデータ
   * @return {boolean} 復元成功可否
   */
  restoreFromBackup(backupData) {
    if (!backupData || !backupData.configs) return false;

    try {
      backupData.configs.forEach(config => {
        if (config.tenantId) {
          this.setUserConfig(config.tenantId, config);
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

      const config = this.configManager.getUserConfig(targetUserId);
      
      if (!config) {
        return this.createAccessResult(false, 'not_found', null, '指定されたユーザーが見つかりません');
      }

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
    if (currentUserEmail && config.ownerEmail === currentUserEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

    if (!config.isPublic) {
      return this.createAccessResult(false, 'private', null, 'このボードは非公開です');
    }

    if (currentUserEmail) {
      return this.createAccessResult(true, 'authenticated_user', config);
    }

    if (config.allowAnonymous) {
      const publicConfig = this.configManager.getPublicConfig(config.tenantId);
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
    if (currentUserEmail && config.ownerEmail === currentUserEmail) {
      return this.createAccessResult(true, 'owner', config);
    }

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
   * 公開ボードの存在確認
   * @param {string} userId ユーザーID
   * @return {boolean} 公開ボード存在可否
   */
  isPublicBoardAvailable(userId) {
    const config = this.configManager.getUserConfig(userId);
    return config && config.isPublic;
  }

  /**
   * ユーザーのアクセス権限レベルを取得
   * @param {string} targetUserId 対象ユーザーID
   * @param {string} currentUserEmail 現在のユーザーメール
   * @return {string} 権限レベル
   */
  getUserPermissionLevel(targetUserId, currentUserEmail) {
    const accessResult = this.verifyAccess(targetUserId, 'view', currentUserEmail);
    
    if (!accessResult.allowed) {
      return 'none';
    }

    return accessResult.userType;
  }
}
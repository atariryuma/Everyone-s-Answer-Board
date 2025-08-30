/**
 * 汎用設定管理システム
 * PropertiesServiceを使用した効率的な設定データ管理
 * データベース非依存でゲストアクセスにも対応
 */

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
    this.cache.put(cacheKey, configData, 1800); // 30分キャッシュ
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

    const maxSize = 524288; // 512KB
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

// グローバルインスタンス
if (typeof configManager === 'undefined') {
  var configManager = new ConfigurationManager();
}


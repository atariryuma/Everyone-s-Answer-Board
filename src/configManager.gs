/**
 * @fileoverview 統一設定マネージャー - PropertiesServiceとconfigJsonの重複解消
 */

/**
 * 設定の優先順位定義
 */
const CONFIG_PRIORITY = {
  SYSTEM: 1,      // PropertiesService（システム全体）
  USER: 2,        // configJson（ユーザー固有）
  RUNTIME: 3,     // 実行時設定（一時的）
  DEFAULT: 4      // デフォルト値
};

/**
 * 設定項目の定義
 */
const CONFIG_SCHEMA = {
  // システム設定（PropertiesService）
  system: {
    DATABASE_SPREADSHEET_ID: { type: 'string', required: true, scope: 'system' },
    SERVICE_ACCOUNT_CREDS: { type: 'object', required: true, scope: 'system' },
    ADMIN_EMAIL: { type: 'string', required: true, scope: 'system' },
    BACKUP_CONFIG: { type: 'object', required: false, scope: 'system' }
  },
  
  // ユーザー設定（configJson）
  user: {
    spreadsheetId: { type: 'string', required: false, scope: 'user' },
    spreadsheetUrl: { type: 'string', required: false, scope: 'user' },
    setupStatus: { type: 'string', required: false, scope: 'user', default: 'initial' },
    formUrl: { type: 'string', required: false, scope: 'user' },
    publishedSheetName: { type: 'string', required: false, scope: 'user' },
    opinionHeader: { type: 'string', required: false, scope: 'user' },
    reasonHeader: { type: 'string', required: false, scope: 'user' },
    nameHeader: { type: 'string', required: false, scope: 'user' },
    timestampHeader: { type: 'string', required: false, scope: 'user' },
    showNames: { type: 'boolean', required: false, scope: 'user', default: false },
    showCounts: { type: 'boolean', required: false, scope: 'user', default: true },
    appPublished: { type: 'boolean', required: false, scope: 'user', default: false }
  },
  
  // 実行時設定（セッション固有）
  runtime: {
    currentUserId: { type: 'string', required: false, scope: 'runtime' },
    debugMode: { type: 'boolean', required: false, scope: 'runtime', default: false },
    cacheEnabled: { type: 'boolean', required: false, scope: 'runtime', default: true }
  }
};

/**
 * 統一設定マネージャークラス
 */
class ConfigManager {
  constructor() {
    this.systemProps = PropertiesService.getScriptProperties();
    this.userProps = PropertiesService.getUserProperties();
    this.configCache = new Map();
    this.lastCacheUpdate = 0;
    this.cacheValidityMs = 300000; // 5分間
  }

  /**
   * 設定値を取得（優先順位に基づく統合取得）
   * @param {string} key - 設定キー
   * @param {string} userId - ユーザーID（user scopeの場合必須）
   * @param {object} options - オプション { useCache: boolean, priority: array }
   * @returns {any} 設定値
   */
  get(key, userId = null, options = {}) {
    const { useCache = true, priority = null } = options;
    
    // キャッシュチェック
    if (useCache) {
      const cachedValue = this._getCachedValue(key, userId);
      if (cachedValue !== undefined) {
        return cachedValue;
      }
    }
    
    // 設定スキーマを確認
    const schema = this._findConfigSchema(key);
    if (!schema) {
      console.warn(`[ConfigManager] Unknown config key: ${key}`);
      return null;
    }
    
    // 優先順位に基づいて値を取得
    const priorityOrder = priority || this._getDefaultPriority(schema.scope);
    let value = null;
    
    for (const source of priorityOrder) {
      value = this._getFromSource(source, key, userId, schema);
      if (value !== null && value !== undefined) {
        break;
      }
    }
    
    // デフォルト値を適用
    if ((value === null || value === undefined) && schema.default !== undefined) {
      value = schema.default;
    }
    
    // 型変換
    value = this._convertType(value, schema.type);
    
    // キャッシュに保存
    if (useCache && value !== null) {
      this._setCachedValue(key, userId, value);
    }
    
    return value;
  }

  /**
   * 設定値を保存
   * @param {string} key - 設定キー
   * @param {any} value - 設定値
   * @param {string} userId - ユーザーID
   * @param {object} options - オプション { scope: string, merge: boolean }
   * @returns {boolean} 保存成功フラグ
   */
  set(key, value, userId = null, options = {}) {
    const { scope = null, merge = false } = options;
    
    const schema = this._findConfigSchema(key);
    if (!schema) {
      console.warn(`[ConfigManager] Unknown config key: ${key}`);
      return false;
    }
    
    // 型検証
    if (!this._validateType(value, schema.type)) {
      console.error(`[ConfigManager] Type validation failed for ${key}: expected ${schema.type}, got ${typeof value}`);
      return false;
    }
    
    // スコープ決定
    const targetScope = scope || schema.scope;
    
    try {
      const success = this._setToSource(targetScope, key, value, userId, merge);
      
      if (success) {
        // キャッシュを無効化
        this._invalidateCache(key, userId);
        console.log(`[ConfigManager] Set ${key} = ${JSON.stringify(value)} (scope: ${targetScope})`);
      }
      
      return success;
      
    } catch (error) {
      console.error(`[ConfigManager] Failed to set ${key}:`, error.message);
      return false;
    }
  }

  /**
   * ユーザー設定を一括取得
   * @param {string} userId - ユーザーID
   * @param {object} options - オプション { includeSystem: boolean }
   * @returns {object} 設定オブジェクト
   */
  getUserConfig(userId, options = {}) {
    const { includeSystem = false } = options;
    
    if (!userId) {
      throw new Error('ユーザーIDが必要です');
    }
    
    try {
      const user = findUserById(userId);
      if (!user) {
        throw new Error('ユーザーが見つかりません');
      }
      
      // configJsonをパース
      let userConfig = {};
      try {
        userConfig = JSON.parse(user.configJson || '{}');
      } catch (parseError) {
        console.warn(`[ConfigManager] Failed to parse configJson for user ${userId}:`, parseError.message);
        userConfig = {};
      }
      
      // スキーマに基づいて統合設定を構築
      const config = {};
      
      // ユーザー設定
      for (const [key, schema] of Object.entries(CONFIG_SCHEMA.user)) {
        config[key] = this.get(key, userId);
      }
      
      // システム設定（要求された場合）
      if (includeSystem) {
        for (const [key, schema] of Object.entries(CONFIG_SCHEMA.system)) {
          config[key] = this.get(key, userId);
        }
      }
      
      return config;
      
    } catch (error) {
      const processedError = processError(error, {
        function: 'getUserConfig',
        userId: userId,
        operation: 'config_retrieval'
      });
      throw new Error(processedError.userMessage);
    }
  }

  /**
   * ユーザー設定を一括保存
   * @param {string} userId - ユーザーID
   * @param {object} config - 設定オブジェクト
   * @param {object} options - オプション { merge: boolean, backup: boolean }
   * @returns {boolean} 保存成功フラグ
   */
  setUserConfig(userId, config, options = {}) {
    const { merge = true, backup = true } = options;
    
    if (!userId) {
      throw new Error('ユーザーIDが必要です');
    }
    
    try {
      // バックアップ作成
      if (backup) {
        createPreOperationBackup(userId, 'config_update');
      }
      
      // 現在の設定を取得（マージの場合）
      let currentConfig = {};
      if (merge) {
        currentConfig = this.getUserConfig(userId, { includeSystem: false });
      }
      
      // 設定を統合
      const updatedConfig = merge ? Object.assign(currentConfig, config) : config;
      
      // 個別設定を保存
      let successCount = 0;
      const errors = [];
      
      for (const [key, value] of Object.entries(updatedConfig)) {
        try {
          const success = this.set(key, value, userId);
          if (success) {
            successCount++;
          } else {
            errors.push(`Failed to set ${key}`);
          }
        } catch (error) {
          errors.push(`Error setting ${key}: ${error.message}`);
        }
      }
      
      // configJsonを更新
      try {
        const configJson = JSON.stringify(updatedConfig);
        updateUser(userId, { configJson });
        console.log(`[ConfigManager] Updated configJson for user ${userId}`);
      } catch (updateError) {
        console.error(`[ConfigManager] Failed to update configJson:`, updateError.message);
        errors.push('Failed to update configJson');
      }
      
      if (errors.length > 0) {
        console.warn(`[ConfigManager] Some config updates failed:`, errors);
        return successCount > 0; // 部分的成功
      }
      
      return true;
      
    } catch (error) {
      const processedError = processError(error, {
        function: 'setUserConfig',
        userId: userId,
        operation: 'config_update'
      });
      throw new Error(processedError.userMessage);
    }
  }

  /**
   * 設定の検証
   * @param {string} userId - ユーザーID
   * @returns {object} 検証結果
   */
  validateConfig(userId) {
    const validation = {
      isValid: true,
      errors: [],
      warnings: [],
      missing: []
    };
    
    try {
      const config = this.getUserConfig(userId, { includeSystem: true });
      
      // 必須項目チェック
      for (const [category, items] of Object.entries(CONFIG_SCHEMA)) {
        for (const [key, schema] of Object.entries(items)) {
          if (schema.required) {
            const value = config[key];
            if (value === null || value === undefined || value === '') {
              validation.missing.push(key);
              validation.isValid = false;
            }
          }
        }
      }
      
      // 型チェック
      for (const [key, value] of Object.entries(config)) {
        const schema = this._findConfigSchema(key);
        if (schema && value !== null && value !== undefined) {
          if (!this._validateType(value, schema.type)) {
            validation.errors.push(`${key}: 型が正しくありません (期待: ${schema.type}, 実際: ${typeof value})`);
            validation.isValid = false;
          }
        }
      }
      
      // 依存関係チェック
      if (config.appPublished && !config.publishedSheetName) {
        validation.warnings.push('アプリが公開されていますが、公開シート名が設定されていません');
      }
      
      if (config.spreadsheetId && !config.spreadsheetUrl) {
        validation.warnings.push('スプレッドシートIDが設定されていますが、URLが設定されていません');
      }
      
    } catch (error) {
      validation.errors.push(`設定検証中にエラーが発生しました: ${error.message}`);
      validation.isValid = false;
    }
    
    return validation;
  }

  /**
   * 設定スキーマを検索
   * @private
   */
  _findConfigSchema(key) {
    for (const category of Object.values(CONFIG_SCHEMA)) {
      if (category[key]) {
        return category[key];
      }
    }
    return null;
  }

  /**
   * デフォルト優先順位を取得
   * @private
   */
  _getDefaultPriority(scope) {
    switch (scope) {
      case 'system':
        return ['system', 'default'];
      case 'user':
        return ['user', 'system', 'default'];
      case 'runtime':
        return ['runtime', 'user', 'system', 'default'];
      default:
        return ['user', 'system', 'default'];
    }
  }

  /**
   * ソースから値を取得
   * @private
   */
  _getFromSource(source, key, userId, schema) {
    try {
      switch (source) {
        case 'system':
          return this.systemProps.getProperty(key);
          
        case 'user':
          if (!userId) return null;
          const user = findUserById(userId);
          if (!user || !user.configJson) return null;
          const config = JSON.parse(user.configJson);
          return config[key];
          
        case 'runtime':
          return this.userProps.getProperty(`runtime_${key}`);
          
        case 'default':
          return schema.default;
          
        default:
          return null;
      }
    } catch (error) {
      console.warn(`[ConfigManager] Failed to get from source ${source}:`, error.message);
      return null;
    }
  }

  /**
   * ソースに値を保存
   * @private
   */
  _setToSource(scope, key, value, userId, merge) {
    try {
      switch (scope) {
        case 'system':
          if (typeof value === 'object') {
            this.systemProps.setProperty(key, JSON.stringify(value));
          } else {
            this.systemProps.setProperty(key, String(value));
          }
          return true;
          
        case 'user':
          if (!userId) {
            throw new Error('ユーザーIDが必要です');
          }
          // configJsonに保存（updateUser経由で更新される）
          return true;
          
        case 'runtime':
          this.userProps.setProperty(`runtime_${key}`, String(value));
          return true;
          
        default:
          throw new Error(`未知のスコープ: ${scope}`);
      }
    } catch (error) {
      console.error(`[ConfigManager] Failed to set to source ${scope}:`, error.message);
      return false;
    }
  }

  /**
   * 型変換
   * @private
   */
  _convertType(value, expectedType) {
    if (value === null || value === undefined) {
      return value;
    }
    
    switch (expectedType) {
      case 'boolean':
        if (typeof value === 'string') {
          return value.toLowerCase() === 'true';
        }
        return Boolean(value);
        
      case 'number':
        return Number(value);
        
      case 'object':
        if (typeof value === 'string') {
          try {
            return JSON.parse(value);
          } catch (e) {
            console.warn(`[ConfigManager] Failed to parse object: ${value}`);
            return null;
          }
        }
        return value;
        
      case 'string':
      default:
        return String(value);
    }
  }

  /**
   * 型検証
   * @private
   */
  _validateType(value, expectedType) {
    if (value === null || value === undefined) {
      return true; // null/undefinedは常に有効
    }
    
    switch (expectedType) {
      case 'boolean':
        return typeof value === 'boolean' || (typeof value === 'string' && ['true', 'false'].includes(value.toLowerCase()));
      case 'number':
        return typeof value === 'number' || !isNaN(Number(value));
      case 'object':
        return typeof value === 'object' || typeof value === 'string';
      case 'string':
        return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
      default:
        return true;
    }
  }

  /**
   * キャッシュから値を取得
   * @private
   */
  _getCachedValue(key, userId) {
    if (Date.now() - this.lastCacheUpdate > this.cacheValidityMs) {
      this.configCache.clear();
      this.lastCacheUpdate = Date.now();
      return undefined;
    }
    
    const cacheKey = userId ? `${userId}:${key}` : key;
    return this.configCache.get(cacheKey);
  }

  /**
   * キャッシュに値を保存
   * @private
   */
  _setCachedValue(key, userId, value) {
    const cacheKey = userId ? `${userId}:${key}` : key;
    this.configCache.set(cacheKey, value);
    
    // キャッシュサイズ制限
    if (this.configCache.size > 200) {
      const keysToDelete = Array.from(this.configCache.keys()).slice(0, 50);
      keysToDelete.forEach(k => this.configCache.delete(k));
    }
  }

  /**
   * キャッシュを無効化
   * @private
   */
  _invalidateCache(key, userId) {
    if (userId) {
      const cacheKey = `${userId}:${key}`;
      this.configCache.delete(cacheKey);
    } else {
      // 部分マッチで削除
      for (const cacheKey of this.configCache.keys()) {
        if (cacheKey.includes(key)) {
          this.configCache.delete(cacheKey);
        }
      }
    }
  }

  /**
   * 設定の詳細レポートを生成
   * @param {string} userId - ユーザーID
   * @returns {string} レポート
   */
  generateConfigReport(userId) {
    try {
      const config = this.getUserConfig(userId, { includeSystem: false });
      const validation = this.validateConfig(userId);
      
      let report = '=== 設定レポート ===\n';
      report += `ユーザーID: ${userId}\n`;
      report += `検証結果: ${validation.isValid ? '✅ 正常' : '❌ 問題あり'}\n\n`;
      
      if (validation.missing.length > 0) {
        report += '⚠️ 未設定項目:\n';
        validation.missing.forEach(key => {
          report += `  - ${key}\n`;
        });
        report += '\n';
      }
      
      if (validation.errors.length > 0) {
        report += '❌ エラー:\n';
        validation.errors.forEach(error => {
          report += `  - ${error}\n`;
        });
        report += '\n';
      }
      
      if (validation.warnings.length > 0) {
        report += '⚠️ 警告:\n';
        validation.warnings.forEach(warning => {
          report += `  - ${warning}\n`;
        });
        report += '\n';
      }
      
      report += '📋 現在の設定:\n';
      for (const [key, value] of Object.entries(config)) {
        if (value !== null && value !== undefined) {
          const displayValue = typeof value === 'object' ? JSON.stringify(value) : String(value);
          report += `  ${key}: ${displayValue}\n`;
        }
      }
      
      return report;
      
    } catch (error) {
      return `設定レポートの生成に失敗しました: ${error.message}`;
    }
  }
}

// シングルトンインスタンス
const configManager = new ConfigManager();

/**
 * 後方互換性のためのラッパー関数
 */

/**
 * 統一設定取得関数
 * @param {string} key - 設定キー
 * @param {string} userId - ユーザーID
 * @returns {any} 設定値
 */
function getUnifiedConfig(key, userId = null) {
  return configManager.get(key, userId);
}

/**
 * 統一設定保存関数
 * @param {string} key - 設定キー
 * @param {any} value - 設定値
 * @param {string} userId - ユーザーID
 * @returns {boolean} 保存成功フラグ
 */
function setUnifiedConfig(key, value, userId = null) {
  return configManager.set(key, value, userId);
}
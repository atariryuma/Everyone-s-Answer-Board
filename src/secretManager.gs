/**
 * @fileoverview 統一秘密情報管理システム
 * Google Secret Manager統合とセキュリティ強化
 */

/**
 * 統一秘密情報管理クラス
 * Google Secret ManagerとPropertiesServiceのハイブリッド管理
 */
class UnifiedSecretManager {
  constructor(options = {}) {
    this.config = {
      useSecretManager: options.useSecretManager !== false, // デフォルトtrue
      projectId: options.projectId || this.getProjectIdFromEnvironment(),
      fallbackToProperties: options.fallbackToProperties !== false, // デフォルトtrue
      encryptionEnabled: options.encryptionEnabled !== false, // デフォルトtrue
      auditLogging: options.auditLogging !== false, // デフォルトtrue
      rotationEnabled: options.rotationEnabled === true, // デフォルトfalse
      cacheSecretsLocally: options.cacheSecretsLocally !== false, // デフォルトtrue
      cacheTTL: options.cacheTTL || 300 // 5分
    };

    // セキュリティ監査ログ
    this.auditLog = [];
    this.secretCache = new Map();
    
    // 秘密情報のタイプ分類
    this.secretTypes = {
      SERVICE_ACCOUNT: 'service_account',
      API_KEY: 'api_key', 
      DATABASE_CREDS: 'database_creds',
      WEBHOOK_SECRET: 'webhook_secret',
      ENCRYPTION_KEY: 'encryption_key'
    };

    // 重要な秘密情報のマッピング
    this.criticalSecrets = {
      'SERVICE_ACCOUNT_CREDS': this.secretTypes.SERVICE_ACCOUNT,
      'DATABASE_SPREADSHEET_ID': this.secretTypes.DATABASE_CREDS,
      'WEBHOOK_SECRET': this.secretTypes.WEBHOOK_SECRET
    };
  }

  /**
   * 秘密情報の安全な取得
   * @param {string} secretName - 秘密情報名
   * @param {object} options - オプション
   * @returns {Promise<string>} 秘密情報の値
   */
  async getSecret(secretName, options = {}) {
    const {
      useCache = this.config.cacheSecretsLocally,
      version = 'latest',
      fallback = this.config.fallbackToProperties,
      auditLog = this.config.auditLogging
    } = options;

    // 入力検証
    if (!secretName || typeof secretName !== 'string') {
      throw new Error('SECURITY_ERROR: 無効な秘密情報名');
    }

    // 監査ログ記録
    if (auditLog) {
      this.logSecretAccess('GET', secretName, {
        useCache,
        version,
        timestamp: new Date().toISOString(),
        userEmail: Session.getActiveUser().getEmail()
      });
    }

    // キャッシュ確認
    if (useCache && this.isSecretCached(secretName)) {
      const cached = this.getSecretFromCache(secretName);
      if (cached && !this.isCacheExpired(secretName)) {
        debugLog(`🔐 秘密情報キャッシュヒット: ${secretName}`);
        return cached;
      }
    }

    let secretValue = null;

    // Google Secret Manager から取得を試行
    if (this.config.useSecretManager && this.config.projectId) {
      try {
        secretValue = await this.getSecretFromManager(secretName, version);
        if (secretValue) {
          debugLog(`🔐 Secret Manager取得成功: ${secretName}`);
          if (useCache) {
            this.cacheSecret(secretName, secretValue);
          }
          return secretValue;
        }
      } catch (error) {
        warnLog(`Secret Manager取得エラー (${secretName}):`, error.message);
        if (!fallback) {
          throw error;
        }
      }
    }

    // フォールバック: PropertiesService から取得
    if (fallback) {
      try {
        secretValue = await this.getSecretFromProperties(secretName);
        if (secretValue) {
          debugLog(`🔐 Properties Service取得成功: ${secretName}`);
          if (useCache) {
            this.cacheSecret(secretName, secretValue);
          }
          return secretValue;
        }
      } catch (error) {
        errorLog(`Properties Service取得エラー (${secretName}):`, error.message);
        throw error;
      }
    }

    throw new Error(`SECURITY_ERROR: 秘密情報が見つかりません: ${secretName}`);
  }

  /**
   * 秘密情報の安全な設定
   * @param {string} secretName - 秘密情報名
   * @param {string} secretValue - 秘密情報の値
   * @param {object} options - オプション
   * @returns {Promise<boolean>} 設定成功フラグ
   */
  async setSecret(secretName, secretValue, options = {}) {
    const {
      useSecretManager = this.config.useSecretManager,
      updateProperties = true,
      encrypt = this.config.encryptionEnabled,
      auditLog = this.config.auditLogging
    } = options;

    // 入力検証とセキュリティチェック
    if (!secretName || !secretValue) {
      throw new Error('SECURITY_ERROR: 秘密情報名と値は必須です');
    }

    if (typeof secretValue !== 'string') {
      throw new Error('SECURITY_ERROR: 秘密情報の値は文字列である必要があります');
    }

    // 秘密情報の機密性チェック
    if (this.isCriticalSecret(secretName)) {
      if (!this.validateCriticalSecret(secretName, secretValue)) {
        throw new Error(`SECURITY_ERROR: 重要な秘密情報の検証に失敗: ${secretName}`);
      }
    }

    // 監査ログ記録
    if (auditLog) {
      this.logSecretAccess('SET', secretName, {
        useSecretManager,
        updateProperties,
        encrypt,
        valueLength: secretValue.length,
        timestamp: new Date().toISOString(),
        userEmail: Session.getActiveUser().getEmail()
      });
    }

    let success = false;

    // Google Secret Manager に保存を試行
    if (useSecretManager && this.config.projectId) {
      try {
        await this.setSecretInManager(secretName, secretValue);
        success = true;
        debugLog(`🔐 Secret Manager保存成功: ${secretName}`);
      } catch (error) {
        warnLog(`Secret Manager保存エラー (${secretName}):`, error.message);
        if (!updateProperties) {
          throw error;
        }
      }
    }

    // PropertiesService にも保存（バックアップまたはフォールバック）
    if (updateProperties) {
      try {
        const valueToStore = encrypt ? this.encryptValue(secretValue) : secretValue;
        await this.setSecretInProperties(secretName, valueToStore, { encrypted: encrypt });
        success = true;
        debugLog(`🔐 Properties Service保存成功: ${secretName}`);
      } catch (error) {
        errorLog(`Properties Service保存エラー (${secretName}):`, error.message);
        if (!success) {
          throw error;
        }
      }
    }

    // キャッシュをクリア
    this.clearSecretCache(secretName);

    return success;
  }

  /**
   * Google Secret Manager から秘密情報を取得
   * @private
   */
  async getSecretFromManager(secretName, version = 'latest') {
    const secretPath = `projects/${this.config.projectId}/secrets/${secretName}/versions/${version}`;
    
    try {
      // Google Cloud Secret Manager API 呼び出し
      const response = await resilientUrlFetch(
        `https://secretmanager.googleapis.com/v1/${secretPath}:access`,
        {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${await getServiceAccountTokenCached()}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.getResponseCode() !== 200) {
        throw new Error(`Secret Manager API error: ${response.getResponseCode()}`);
      }

      const data = JSON.parse(response.getContentText());
      return data.payload ? Utilities.newBlob(Utilities.base64Decode(data.payload.data)).getDataAsString() : null;

    } catch (error) {
      if (error.message.includes('404')) {
        debugLog(`Secret not found in Secret Manager: ${secretName}`);
        return null;
      }
      throw error;
    }
  }

  /**
   * Google Secret Manager に秘密情報を保存
   * @private
   */
  async setSecretInManager(secretName, secretValue) {
    // まずシークレットが存在するか確認
    const secretsListUrl = `https://secretmanager.googleapis.com/v1/projects/${this.config.projectId}/secrets`;
    
    try {
      // シークレットの存在確認
      const existsResponse = await resilientUrlFetch(
        `${secretsListUrl}/${secretName}`,
        {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${await getServiceAccountTokenCached()}`,
            'Content-Type': 'application/json'
          }
        }
      );

      let secretExists = existsResponse.getResponseCode() === 200;

      // シークレットが存在しない場合は作成
      if (!secretExists) {
        const createResponse = await resilientUrlFetch(secretsListUrl, {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${await getServiceAccountTokenCached()}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify({
            secretId: secretName,
            secret: {
              replication: { automatic: {} }
            }
          })
        });

        if (createResponse.getResponseCode() !== 200) {
          throw new Error(`Failed to create secret: ${createResponse.getContentText()}`);
        }
      }

      // バージョンを追加
      const addVersionResponse = await resilientUrlFetch(
        `${secretsListUrl}/${secretName}:addVersion`,
        {
          method: 'POST',
          headers: {
            'Authorization': `Bearer ${await getServiceAccountTokenCached()}`,
            'Content-Type': 'application/json'
          },
          payload: JSON.stringify({
            payload: {
              data: Utilities.base64Encode(secretValue)
            }
          })
        }
      );

      if (addVersionResponse.getResponseCode() !== 200) {
        throw new Error(`Failed to add secret version: ${addVersionResponse.getContentText()}`);
      }

      return true;

    } catch (error) {
      errorLog(`Secret Manager保存エラー:`, error.message);
      throw error;
    }
  }

  /**
   * Properties Service から秘密情報を取得
   * @private
   */
  async getSecretFromProperties(secretName) {
    try {
      const props = await getResilientScriptProperties();
      let value = props.getProperty(secretName);
      
      if (!value) {
        return null;
      }

      // 暗号化されているかチェック
      if (this.isEncryptedValue(value)) {
        value = this.decryptValue(value);
      }

      return value;
    } catch (error) {
      errorLog(`Properties Service取得エラー:`, error.message);
      throw error;
    }
  }

  /**
   * Properties Service に秘密情報を保存
   * @private
   */
  async setSecretInProperties(secretName, secretValue, options = {}) {
    try {
      const props = await getResilientScriptProperties();
      props.setProperty(secretName, secretValue);
      
      // 暗号化メタデータの保存
      if (options.encrypted) {
        props.setProperty(`${secretName}_ENCRYPTED`, 'true');
      }
      
      return true;
    } catch (error) {
      errorLog(`Properties Service保存エラー:`, error.message);
      throw error;
    }
  }

  /**
   * 値の暗号化（簡易実装）
   * @private
   */
  encryptValue(value) {
    try {
      // 簡易暗号化（Base64エンコード + 固定キーでのXOR）
      const key = 'ENCRYPTION_KEY_2024';
      let encrypted = '';
      
      for (let i = 0; i < value.length; i++) {
        const charCode = value.charCodeAt(i) ^ key.charCodeAt(i % key.length);
        encrypted += String.fromCharCode(charCode);
      }
      
      return 'ENC:' + Utilities.base64Encode(encrypted);
    } catch (error) {
      warnLog('暗号化エラー:', error.message);
      return value; // 暗号化失敗時は平文を返す
    }
  }

  /**
   * 値の復号化
   * @private
   */
  decryptValue(encryptedValue) {
    try {
      if (!encryptedValue.startsWith('ENC:')) {
        return encryptedValue;
      }
      
      const encrypted = Utilities.base64Decode(encryptedValue.substring(4));
      const encryptedString = Utilities.newBlob(encrypted).getDataAsString();
      const key = 'ENCRYPTION_KEY_2024';
      let decrypted = '';
      
      for (let i = 0; i < encryptedString.length; i++) {
        const charCode = encryptedString.charCodeAt(i) ^ key.charCodeAt(i % key.length);
        decrypted += String.fromCharCode(charCode);
      }
      
      return decrypted;
    } catch (error) {
      warnLog('復号化エラー:', error.message);
      throw new Error('SECURITY_ERROR: 秘密情報の復号化に失敗');
    }
  }

  /**
   * 暗号化された値かチェック
   * @private
   */
  isEncryptedValue(value) {
    return typeof value === 'string' && value.startsWith('ENC:');
  }

  /**
   * 重要な秘密情報かチェック
   * @private
   */
  isCriticalSecret(secretName) {
    return this.criticalSecrets.hasOwnProperty(secretName);
  }

  /**
   * 重要な秘密情報の検証
   * @private
   */
  validateCriticalSecret(secretName, secretValue) {
    const secretType = this.criticalSecrets[secretName];
    
    switch (secretType) {
      case this.secretTypes.SERVICE_ACCOUNT:
        try {
          const creds = JSON.parse(secretValue);
          return creds.client_email && creds.private_key && creds.type === 'service_account';
        } catch {
          return false;
        }
        
      case this.secretTypes.DATABASE_CREDS:
        // スプレッドシートIDの形式チェック
        return /^[a-zA-Z0-9_-]{44}$/.test(secretValue);
        
      case this.secretTypes.WEBHOOK_SECRET:
        // Webhook秘密情報の最小長チェック
        return secretValue.length >= 32;
        
      default:
        return true;
    }
  }

  /**
   * プロジェクトIDを環境から取得
   * @private
   */
  getProjectIdFromEnvironment() {
    try {
      // GAS環境での取得を試行
      return PropertiesService.getScriptProperties().getProperty('GCP_PROJECT_ID');
    } catch {
      return null;
    }
  }

  /**
   * キャッシュ関連のメソッド群
   */
  isSecretCached(secretName) {
    return this.secretCache.has(secretName);
  }

  getSecretFromCache(secretName) {
    const cached = this.secretCache.get(secretName);
    return cached ? cached.value : null;
  }

  isCacheExpired(secretName) {
    const cached = this.secretCache.get(secretName);
    if (!cached) return true;
    
    return (Date.now() - cached.timestamp) > (this.config.cacheTTL * 1000);
  }

  cacheSecret(secretName, secretValue) {
    this.secretCache.set(secretName, {
      value: secretValue,
      timestamp: Date.now()
    });
  }

  clearSecretCache(secretName) {
    if (secretName) {
      this.secretCache.delete(secretName);
    } else {
      this.secretCache.clear();
    }
  }

  /**
   * 監査ログ記録
   * @private
   */
  logSecretAccess(action, secretName, metadata) {
    if (!this.config.auditLogging) return;

    const logEntry = {
      timestamp: metadata.timestamp || new Date().toISOString(),
      action: action,
      secretName: secretName,
      userEmail: metadata.userEmail,
      metadata: metadata
    };

    this.auditLog.push(logEntry);
    
    // ログサイズ制限（最大1000件）
    if (this.auditLog.length > 1000) {
      this.auditLog.shift();
    }

    // 重要な操作の場合は詳細ログ出力
    if (this.isCriticalSecret(secretName)) {
      infoLog(`🔐 重要秘密情報アクセス: ${action} ${secretName}`, logEntry);
    } else {
      debugLog(`🔐 秘密情報アクセス: ${action} ${secretName}`);
    }
  }

  /**
   * 監査ログの取得
   * @returns {Array} 監査ログ
   */
  getAuditLog() {
    return [...this.auditLog];
  }

  /**
   * 秘密情報管理の健全性チェック
   * @returns {object} チェック結果
   */
  async performHealthCheck() {
    const results = {
      timestamp: new Date().toISOString(),
      secretManagerStatus: 'UNKNOWN',
      propertiesServiceStatus: 'UNKNOWN',
      encryptionStatus: 'UNKNOWN',
      criticalSecretsStatus: 'UNKNOWN',
      issues: []
    };

    try {
      // Secret Manager接続テスト
      if (this.config.useSecretManager && this.config.projectId) {
        try {
          await resilientUrlFetch(
            `https://secretmanager.googleapis.com/v1/projects/${this.config.projectId}/secrets`,
            {
              method: 'GET',
              headers: {
                'Authorization': `Bearer ${await getServiceAccountTokenCached()}`
              }
            }
          );
          results.secretManagerStatus = 'OK';
        } catch (error) {
          results.secretManagerStatus = 'ERROR';
          results.issues.push(`Secret Manager接続エラー: ${error.message}`);
        }
      } else {
        results.secretManagerStatus = 'DISABLED';
      }

      // Properties Service テスト
      try {
        const props = await getResilientScriptProperties();
        props.getProperty('TEST_KEY'); // 存在しないキーでのテスト
        results.propertiesServiceStatus = 'OK';
      } catch (error) {
        results.propertiesServiceStatus = 'ERROR';
        results.issues.push(`Properties Service エラー: ${error.message}`);
      }

      // 暗号化機能テスト
      if (this.config.encryptionEnabled) {
        try {
          const testValue = 'test_encryption_value';
          const encrypted = this.encryptValue(testValue);
          const decrypted = this.decryptValue(encrypted);
          
          if (decrypted === testValue) {
            results.encryptionStatus = 'OK';
          } else {
            results.encryptionStatus = 'ERROR';
            results.issues.push('暗号化/復号化の整合性エラー');
          }
        } catch (error) {
          results.encryptionStatus = 'ERROR';
          results.issues.push(`暗号化機能エラー: ${error.message}`);
        }
      } else {
        results.encryptionStatus = 'DISABLED';
      }

      // 重要な秘密情報の存在確認
      try {
        let criticalCount = 0;
        let foundCount = 0;

        for (const secretName of Object.keys(this.criticalSecrets)) {
          criticalCount++;
          try {
            const value = await this.getSecret(secretName, { auditLog: false });
            if (value) {
              foundCount++;
            }
          } catch (error) {
            // 存在しない場合はカウントしない
          }
        }

        if (foundCount === criticalCount) {
          results.criticalSecretsStatus = 'OK';
        } else {
          results.criticalSecretsStatus = 'WARNING';
          results.issues.push(`重要秘密情報: ${foundCount}/${criticalCount} が利用可能`);
        }
      } catch (error) {
        results.criticalSecretsStatus = 'ERROR';
        results.issues.push(`重要秘密情報チェックエラー: ${error.message}`);
      }

    } catch (error) {
      results.issues.push(`ヘルスチェック中にエラー: ${error.message}`);
    }

    return results;
  }
}

/**
 * グローバルな統一秘密情報管理インスタンス
 */
const unifiedSecretManager = new UnifiedSecretManager({
  useSecretManager: true,
  fallbackToProperties: true,
  encryptionEnabled: true,
  auditLogging: true,
  cacheSecretsLocally: true,
  cacheTTL: 300 // 5分
});

/**
 * 便利なヘルパー関数群（既存コードとの互換性）
 */

/**
 * 安全なサービスアカウント認証情報取得
 * @returns {Promise<object>} サービスアカウント認証情報
 */
async function getSecureServiceAccountCreds() {
  const credsString = await unifiedSecretManager.getSecret('SERVICE_ACCOUNT_CREDS');
  if (!credsString) {
    throw new Error('SECURITY_ERROR: サービスアカウント認証情報が設定されていません');
  }

  try {
    const creds = JSON.parse(credsString);
    if (!creds.client_email || !creds.private_key) {
      throw new Error('SECURITY_ERROR: 不完全なサービスアカウント認証情報');
    }
    return creds;
  } catch (error) {
    throw new Error(`SECURITY_ERROR: サービスアカウント認証情報の解析に失敗: ${error.message}`);
  }
}

/**
 * 安全なデータベースID取得
 * @returns {Promise<string>} データベーススプレッドシートID
 */
async function getSecureDatabaseId() {
  return await unifiedSecretManager.getSecret('DATABASE_SPREADSHEET_ID');
}

/**
 * 秘密情報の安全な取得（既存コードとの互換性）
 * @param {string} key - 秘密情報キー
 * @returns {Promise<string>} 秘密情報の値
 */
async function getSecureProperty(key) {
  return await unifiedSecretManager.getSecret(key);
}

/**
 * 秘密情報の安全な設定（既存コードとの互換性）
 * @param {string} key - 秘密情報キー  
 * @param {string} value - 秘密情報の値
 * @param {object} options - オプション
 * @returns {Promise<boolean>} 設定成功フラグ
 */
async function setSecureProperty(key, value, options = {}) {
  return await unifiedSecretManager.setSecret(key, value, options);
}
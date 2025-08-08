/**
 * @fileoverview マルチテナントデータ分離と検証強化
 * サイロ型マルチテナンシー実現とセキュリティ強化
 */

/**
 * マルチテナントセキュリティ管理クラス
 * ユーザー間のデータ分離を強制し、テナント境界を厳密に管理
 */
class MultiTenantSecurityManager {
  constructor() {
    this.tenantBoundaryEnforcement = true;
    this.strictDataIsolation = true;
    this.auditTrailEnabled = true;
  }

  /**
   * ユーザーIDベースのテナント境界検証
   * @param {string} requestUserId - リクエスト元ユーザーID
   * @param {string} targetUserId - 対象データのユーザーID
   * @param {string} operation - 実行しようとする操作
   * @returns {boolean} アクセス許可の可否
   */
  async validateTenantBoundary(requestUserId, targetUserId, operation) {
    if (!requestUserId || !targetUserId) {
      this.logSecurityViolation('MISSING_USER_ID', { requestUserId, targetUserId, operation });
      return false;
    }

    // 同一テナント（ユーザー）内のアクセスは許可
    if (requestUserId === targetUserId) {
      this.logDataAccess('TENANT_BOUNDARY_VALID', { requestUserId, operation });
      return true;
    }

    // 管理者権限チェック（必要に応じて実装）
    const hasAdminAccess = await this.checkAdminAccess(requestUserId, targetUserId, operation);
    if (hasAdminAccess) {
      this.logDataAccess('ADMIN_ACCESS_GRANTED', { requestUserId, targetUserId, operation });
      return true;
    }

    // テナント境界違反
    this.logSecurityViolation('TENANT_BOUNDARY_VIOLATION', {
      requestUserId,
      targetUserId,
      operation
    });
    
    return false;
  }

  /**
   * データアクセスパターンの検証
   * @param {string} userId - ユーザーID
   * @param {string} dataType - データタイプ
   * @param {string} operation - 操作タイプ
   * @returns {boolean} アクセス許可の可否
   */
  async validateDataAccessPattern(userId, dataType, operation) {
    const allowedPatterns = {
      'user_data': ['read', 'write', 'delete'],
      'user_config': ['read', 'write'],
      'user_cache': ['read', 'write', 'delete'],
      'shared_config': ['read'], // 共有設定は読み取りのみ
    };

    if (!allowedPatterns[dataType] || !allowedPatterns[dataType].includes(operation)) {
      this.logSecurityViolation('INVALID_ACCESS_PATTERN', { userId, dataType, operation });
      return false;
    }

    return true;
  }

  /**
   * キャッシュキーのテナント分離強制
   * @param {string} baseKey - ベースキー
   * @param {string} userId - ユーザーID
   * @param {string} context - コンテキスト情報
   * @returns {string} テナント分離されたキー
   */
  generateSecureKey(baseKey, userId, context = '') {
    if (!userId) {
      throw new Error('SECURITY_ERROR: ユーザーIDが必要です');
    }

    // ユーザーIDの暗号化ハッシュを使用してテナント分離を強化
    const userHash = this.generateUserHash(userId);
    const timestamp = Math.floor(Date.now() / 3600000); // 1時間単位
    
    let secureKey = `MT_${userHash}_${baseKey}`;
    if (context) {
      secureKey += `_${context}`;
    }
    
    // タイムベースの検証情報を追加
    secureKey += `_T${timestamp}`;
    
    return secureKey;
  }

  /**
   * セキュアなユーザーハッシュ生成
   * @param {string} userId - ユーザーID
   * @returns {string} ハッシュ値
   */
  generateUserHash(userId) {
    if (!userId) return null;
    
    // MD5ハッシュを使用してユーザーIDを匿名化
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      userId + '_TENANT_SALT_2024',
      Utilities.Charset.UTF_8
    );
    
    return digest.map(function(byte) {
      return (byte + 256).toString(16).slice(-2);
    }).join('').substring(0, 16); // 16文字に短縮
  }

  /**
   * 管理者権限チェック
   * @param {string} requestUserId - リクエスト元ユーザーID
   * @param {string} targetUserId - 対象ユーザーID
   * @param {string} operation - 操作
   * @returns {boolean} 管理者権限の有無
   */
  async checkAdminAccess(requestUserId, targetUserId, operation) {
    // 管理者権限は現在は実装しない（必要に応じて実装）
    // 実装する場合は、管理者ロールの検証ロジックをここに追加
    return false;
  }

  /**
   * セキュリティ違反の記録
   * @param {string} violationType - 違反タイプ
   * @param {object} details - 詳細情報
   */
  logSecurityViolation(violationType, details) {
    if (!this.auditTrailEnabled) return;

    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'SECURITY_VIOLATION',
      violation: violationType,
      details: details,
      userAgent: Session.getActiveUser().getEmail()
    };

    errorLog(`🚨 セキュリティ違反: ${violationType}`, JSON.stringify(logEntry));
    
    // 重大な違反の場合は追加の通知を送信する可能性
    if (violationType === 'TENANT_BOUNDARY_VIOLATION') {
      this.handleCriticalSecurityViolation(logEntry);
    }
  }

  /**
   * データアクセスの記録
   * @param {string} accessType - アクセスタイプ
   * @param {object} details - 詳細情報
   */
  logDataAccess(accessType, details) {
    if (!this.auditTrailEnabled) return;

    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'DATA_ACCESS',
      access: accessType,
      details: details
    };

    debugLog(`🔒 データアクセス: ${accessType}`, JSON.stringify(logEntry));
  }

  /**
   * 重大なセキュリティ違反の処理
   * @param {object} logEntry - ログエントリ
   */
  async handleCriticalSecurityViolation(logEntry) {
    // 重大な違反の場合の追加処理
    // 例：管理者への通知、一時的なアクセス制限など
    warnLog('🚨 重大なセキュリティ違反が発生しました', logEntry);
    
    // フューチャー実装: 通知システムとの連携
    // await this.sendSecurityAlert(logEntry);
  }

  /**
   * テナント分離検証の統計情報取得
   * @returns {object} 統計情報
   */
  getSecurityStats() {
    return {
      tenantBoundaryEnforcement: this.tenantBoundaryEnforcement,
      strictDataIsolation: this.strictDataIsolation,
      auditTrailEnabled: this.auditTrailEnabled,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * グローバルなマルチテナントセキュリティマネージャーインスタンス
 */
const multiTenantSecurity = new MultiTenantSecurityManager();

/**
 * テナント分離を強制するキャッシュ操作ラッパー
 * @param {string} operation - 操作タイプ（'get', 'set', 'delete'）
 * @param {string} baseKey - ベースキー
 * @param {string} userId - ユーザーID
 * @param {any} value - 設定値（set操作時）
 * @param {object} options - オプション
 * @returns {any} 結果
 */
async function secureMultiTenantCacheOperation(operation, baseKey, userId, value = null, options = {}) {
  // テナント境界検証
  const currentUserId = Session.getActiveUser().getEmail(); // 簡易実装
  if (!await multiTenantSecurity.validateTenantBoundary(currentUserId, userId, `cache_${operation}`)) {
    throw new Error(`SECURITY_ERROR: テナント境界違反 - ${operation} operation denied`);
  }

  // データアクセスパターン検証
  if (!await multiTenantSecurity.validateDataAccessPattern(userId, 'user_cache', operation)) {
    throw new Error(`SECURITY_ERROR: 不正なデータアクセスパターン - ${operation} denied`);
  }

  // セキュアキー生成
  const secureKey = multiTenantSecurity.generateSecureKey(baseKey, userId, options.context);
  
  // 回復力のある実行で操作実行
  try {
    switch (operation) {
      case 'get':
        return await resilientCacheOperation(
          () => CacheService.getUserCache().get(secureKey),
          `SecureCache_Get_${baseKey}`,
          () => null // フォールバック
        );

      case 'set':
        const ttl = options.ttl || 300; // 5分デフォルト
        return await resilientCacheOperation(
          () => CacheService.getUserCache().put(secureKey, JSON.stringify(value), ttl),
          `SecureCache_Set_${baseKey}`
        );

      case 'delete':
        return await resilientCacheOperation(
          () => CacheService.getUserCache().remove(secureKey),
          `SecureCache_Delete_${baseKey}`
        );

      default:
        throw new Error(`SECURITY_ERROR: 不正な操作タイプ - ${operation}`);
    }
  } catch (error) {
    multiTenantSecurity.logSecurityViolation('CACHE_OPERATION_FAILED', {
      operation,
      baseKey,
      userId,
      error: error.message
    });
    throw error;
  }
}

/**
 * セキュアなユーザー情報取得
 * @param {string} userId - ユーザーID
 * @param {object} options - オプション
 * @returns {object|null} ユーザー情報
 */
async function getSecureUserInfo(userId, options = {}) {
  const currentUserId = Session.getActiveUser().getEmail();
  
  // テナント境界検証
  if (!await multiTenantSecurity.validateTenantBoundary(currentUserId, userId, 'user_info_access')) {
    throw new Error('SECURITY_ERROR: ユーザー情報アクセス拒否 - テナント境界違反');
  }

  // セキュアなキャッシュキーでユーザー情報を取得
  return await secureMultiTenantCacheOperation('get', 'user_info', userId, null, options);
}

/**
 * セキュアなユーザー設定情報取得
 * @param {string} userId - ユーザーID
 * @param {string} configKey - 設定キー
 * @param {object} options - オプション
 * @returns {any} 設定値
 */
async function getSecureUserConfig(userId, configKey, options = {}) {
  const currentUserId = Session.getActiveUser().getEmail();
  
  // テナント境界検証
  if (!await multiTenantSecurity.validateTenantBoundary(currentUserId, userId, 'user_config_access')) {
    throw new Error('SECURITY_ERROR: ユーザー設定アクセス拒否 - テナント境界違反');
  }

  return await secureMultiTenantCacheOperation('get', `config_${configKey}`, userId, null, options);
}

/**
 * マルチテナントセキュリティの健全性チェック
 * @returns {object} 健全性チェック結果
 */
async function performMultiTenantHealthCheck() {
  const results = {
    timestamp: new Date().toISOString(),
    securityManagerStatus: 'OK',
    tenantIsolationTest: 'OK',
    cacheSecurityTest: 'OK',
    issues: []
  };

  try {
    // セキュリティマネージャーの状態確認
    const stats = multiTenantSecurity.getSecurityStats();
    if (!stats.tenantBoundaryEnforcement) {
      results.issues.push('テナント境界強制が無効化されています');
      results.securityManagerStatus = 'WARNING';
    }

    // テナント分離テスト
    const testUserId1 = 'test_user_1@example.com';
    const testUserId2 = 'test_user_2@example.com';
    
    const isolation1 = await multiTenantSecurity.validateTenantBoundary(testUserId1, testUserId2, 'test');
    if (isolation1) {
      results.issues.push('テナント分離テストが失敗 - 異なるテナント間のアクセスが許可されました');
      results.tenantIsolationTest = 'FAILED';
    }

    const isolation2 = await multiTenantSecurity.validateTenantBoundary(testUserId1, testUserId1, 'test');
    if (!isolation2) {
      results.issues.push('テナント分離テストが失敗 - 同一テナント内のアクセスが拒否されました');
      results.tenantIsolationTest = 'FAILED';
    }

    // キャッシュセキュリティテスト
    const secureKey1 = multiTenantSecurity.generateSecureKey('test', testUserId1);
    const secureKey2 = multiTenantSecurity.generateSecureKey('test', testUserId2);
    
    if (secureKey1 === secureKey2) {
      results.issues.push('セキュアキー生成テストが失敗 - 異なるユーザーで同じキーが生成されました');
      results.cacheSecurityTest = 'FAILED';
    }

  } catch (error) {
    results.issues.push(`健全性チェック中にエラー: ${error.message}`);
    results.securityManagerStatus = 'ERROR';
  }

  return results;
}